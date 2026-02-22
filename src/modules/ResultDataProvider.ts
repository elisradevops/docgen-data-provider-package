import DataProviderUtils from '../utils/DataProviderUtils';
import { TFSServices } from '../helpers/tfs';
import { OpenPcrRequest, PlainTestResult, TestSteps } from '../models/tfs-data';
import { AdoWorkItemComment, AdoWorkItemCommentsResponse } from '../models/ado-comments';
import type {
  MewpBugLink,
  MewpCoverageBugCell,
  MewpCoverageFlatPayload,
  MewpExternalFilesValidationResponse,
  MewpCoverageL3L4Cell,
  MewpCoverageRequestOptions,
  MewpExternalFileRef,
  MewpExternalTableValidationResult,
  MewpCoverageRow,
  MewpInternalValidationFlatPayload,
  MewpInternalValidationRequestOptions,
  MewpInternalValidationRow,
  MewpL2RequirementFamily,
  MewpL2RequirementWorkItem,
  MewpL3L4Link,
  MewpLinkedRequirementsByTestCase,
  MewpRequirementIndex,
  MewpRunStatus,
} from '../models/mewp-reporting';
import logger from '../utils/logger';
import MewpExternalIngestionUtils from '../utils/mewpExternalIngestionUtils';
import MewpExternalTableUtils, {
  MewpExternalFileValidationError,
} from '../utils/mewpExternalTableUtils';
import Utils from '../utils/testStepParserHelper';
import TicketsDataProvider from './TicketsDataProvider';
const pLimit = require('p-limit');
/**
 * Provides methods to fetch, process, and summarize test data from Azure DevOps.
 *
 * This class includes functionalities for:
 * - Fetching test suites, test points, and test cases.
 * - Aligning test steps with iterations and generating detailed results.
 * - Fetching result data based on work items and test runs.
 * - Summarizing test group results, test results, and detailed results.
 * - Fetching linked work items and open PCRs (Problem Change Requests).
 * - Generating test logs and mapping attachments for download.
 * - Supporting test reporter functionalities with customizable field selection and filtering.
 *
 * Key Features:
 * - Hierarchical and flat test suite processing.
 * - Step-level and test-level result alignment.
 * - Customizable result summaries and detailed reports.
 * - Integration with Azure DevOps APIs for fetching and processing test data.
 * - Support for additional configurations like open PCRs, test logs, and step execution analysis.
 *
 * Usage:
 * Instantiate the class with the organization URL and token, and use the provided methods to fetch and process test data.
 */
export default class ResultDataProvider {
  private static readonly MEWP_L2_COVERAGE_COLUMNS = [
    'L2 REQ ID',
    'L2 REQ Title',
    'L2 SubSystem',
    'L2 Run Status',
    'Bug ID',
    'Bug Title',
    'Bug Responsibility',
    'L3 REQ ID',
    'L3 REQ Title',
    'L4 REQ ID',
    'L4 REQ Title',
  ];
  private static readonly INTERNAL_VALIDATION_COLUMNS = [
    'Test Case ID',
    'Test Case Title',
    'Mentioned but Not Linked',
    'Linked but Not Mentioned',
    'Validation Status',
  ];

  orgUrl: string = '';
  token: string = '';
  private limit = pLimit(10);
  private testStepParserHelper: Utils;
  private testToAssociatedItemMap: Map<number, Set<any>>;
  private querySelectedColumns: any[];
  private workItemDiscussionCache: Map<number, any[]>;
  private mewpExternalTableUtils: MewpExternalTableUtils;
  private mewpExternalIngestionUtils: MewpExternalIngestionUtils;
  constructor(orgUrl: string, token: string) {
    this.orgUrl = orgUrl;
    this.token = token;
    this.testStepParserHelper = new Utils(orgUrl, token);
    this.testToAssociatedItemMap = new Map<number, Set<any>>();
    this.querySelectedColumns = [];
    this.workItemDiscussionCache = new Map<number, any[]>();
    this.mewpExternalTableUtils = new MewpExternalTableUtils();
    this.mewpExternalIngestionUtils = new MewpExternalIngestionUtils(this.mewpExternalTableUtils);
  }

  /**
   * Combines the results of test group result summary, test results summary, and detailed results summary into a single key-value pair array.
   */
  public async getCombinedResultsSummary(
    testPlanId: string,
    projectName: string,
    selectedSuiteIds?: number[],
    addConfiguration: boolean = false,
    isHierarchyGroupName: boolean = false,
    openPcrRequest: OpenPcrRequest | null = null,
    includeTestLog: boolean = false,
    stepExecution?: any,
    stepAnalysis?: any,
    includeHardCopyRun: boolean = false
  ): Promise<{
    combinedResults: any[];
    openPcrToTestCaseTraceMap: Map<string, string[]>;
    testCaseToOpenPcrTraceMap: Map<string, string[]>;
  }> {
    const combinedResults: any[] = [];
    try {
      const openPcrToTestCaseTraceMap = new Map<string, string[]>();
      const testCaseToOpenPcrTraceMap = new Map<string, string[]>();
      // Fetch test suites
      const suites = await this.fetchTestSuites(
        testPlanId,
        projectName,
        selectedSuiteIds,
        isHierarchyGroupName
      );

      // Prepare test data for summaries
      const testPointsPromises = suites.map((suite) =>
        this.limit(() =>
          this.fetchTestPoints(projectName, testPlanId, suite.testSuiteId)
            .then((testPointsItems) => ({ ...suite, testPointsItems }))
            .catch((error: any) => {
              logger.error(`Error occurred for suite ${suite.testSuiteId}: ${error.message}`);
              return { ...suite, testPointsItems: [] };
            })
        )
      );
      const testPoints = await Promise.all(testPointsPromises);

      // 1. Calculate Test Group Result Summary
      const summarizedResults = testPoints
        .filter((testPoint: any) => testPoint.testPointsItems && testPoint.testPointsItems.length > 0)
        .map((testPoint: any) => {
          const groupResultSummary = this.calculateGroupResultSummary(
            testPoint.testPointsItems || [],
            includeHardCopyRun
          );
          return { ...testPoint, groupResultSummary };
        });

      const totalSummary = this.calculateTotalSummary(summarizedResults, includeHardCopyRun);
      const testGroupArray = summarizedResults.map((item: any) => ({
        testGroupName: item.testGroupName,
        ...item.groupResultSummary,
      }));
      testGroupArray.push({ testGroupName: 'Total', ...totalSummary });

      // Add test group result summary to combined results
      combinedResults.push({
        contentControl: 'test-group-summary-content-control',
        data: testGroupArray,
        skin: 'test-result-test-group-summary-table',
      });

      // 2. Calculate Test Results Summary
      const flattenedTestPoints = this.flattenTestPoints(testPoints);
      const testResultsSummary = flattenedTestPoints.map((testPoint) =>
        this.formatTestResult(testPoint, addConfiguration, includeHardCopyRun)
      );

      // Add test results summary to combined results
      combinedResults.push({
        contentControl: 'test-result-summary-content-control',
        data: testResultsSummary,
        skin: 'test-result-table',
      });

      // 3. Calculate Detailed Results Summary
      const testData = await this.fetchTestData(suites, projectName, testPlanId, false);
      const runResults = await this.fetchAllResultData(testData, projectName);
      const detailedStepResultsSummary = this.alignStepsWithIterations(testData, runResults)?.filter(
        (results) => results !== null
      );
      //Filter out all the results with no comment
      const filteredDetailedResults = detailedStepResultsSummary.filter(
        (result) => result && (result.stepComments !== '' || result.stepStatus === 'Failed')
      );

      // Add detailed results summary to combined results
      combinedResults.push({
        contentControl: 'detailed-test-result-content-control',
        data: !includeHardCopyRun ? filteredDetailedResults : [],
        skin: 'detailed-test-result-table',
      });

      if (openPcrRequest?.openPcrMode === 'linked') {
        //5. Open PCRs data (only if enabled)
        await this.fetchOpenPcrData(
          testResultsSummary,
          projectName,
          openPcrToTestCaseTraceMap,
          testCaseToOpenPcrTraceMap
        );
      }

      //6. Test Log (only if enabled)
      if (includeTestLog) {
        this.fetchTestLogData(flattenedTestPoints, combinedResults);
      }

      if (stepAnalysis && stepAnalysis.isEnabled) {
        const mappedAnalysisData = runResults.filter(
          (result) =>
            result.comment ||
            result.iteration?.attachments?.length > 0 ||
            result.analysisAttachments?.length > 0
        );

        const mappedAnalysisResultData = stepAnalysis.generateRunAttachments.isEnabled
          ? this.mapAttachmentsUrl(mappedAnalysisData, projectName)
          : mappedAnalysisData;
        if (mappedAnalysisResultData?.length > 0) {
          combinedResults.push({
            contentControl: 'appendix-a-content-control',
            data: mappedAnalysisResultData,
            skin: 'step-analysis-appendix-skin',
          });
        }
      }

      if (stepExecution && stepExecution.isEnabled) {
        const mappedAnalysisData =
          stepExecution.generateAttachments.isEnabled &&
          stepExecution.generateAttachments.runAttachmentMode !== 'planOnly'
            ? runResults.filter((result) => result.iteration?.attachments?.length > 0)
            : [];
        const mappedAnalysisResultData =
          mappedAnalysisData.length > 0 ? this.mapAttachmentsUrl(mappedAnalysisData, projectName) : [];

        const mappedDetailedResults = this.mapStepResultsForExecutionAppendix(
          detailedStepResultsSummary,
          mappedAnalysisResultData
        );
        combinedResults.push({
          contentControl: 'appendix-b-content-control',
          data: mappedDetailedResults,
          skin: 'step-execution-appendix-skin',
        });
      }

      return { combinedResults, openPcrToTestCaseTraceMap, testCaseToOpenPcrTraceMap };
    } catch (error: any) {
      logger.error(`Error during getCombinedResultsSummary: ${error.message}`);
      if (error.response) {
        logger.error(`Response Data: ${JSON.stringify(error.response.data)}`);
      }
      // Ensure the error is rethrown to propagate it correctly
      throw new Error(error.message || 'Unknown error occurred during getCombinedResultsSummary');
    }
  }

  /**
   * Fetches and processes test reporter results for a given test plan and project.
   *
   * @param testPlanId - The ID of the test plan to fetch results for.
   * @param projectName - The name of the project associated with the test plan.
   * @param selectedSuiteIds - An array of suite IDs to filter the test results.
   * @param selectedFields - An array of field names to include in the results.
   * @param enableRunStepStatusFilter - A flag to enable filtering out test steps with a "Not Run" status.
   * @returns A promise that resolves to an array of test reporter results, formatted for use in a test-reporter-table.
   *
   * @throws Will log an error if any step in the process fails.
   */
  public async getTestReporterResults(
    testPlanId: string,
    projectName: string,
    selectedSuiteIds: number[],
    selectedFields: string[],
    allowCrossTestPlan: boolean,
    enableRunTestCaseFilter: boolean,
    enableRunStepStatusFilter: boolean,
    linkedQueryRequest: any,
    errorFilterMode: string = 'none',
    includeAllHistory: boolean = false
  ) {
    const fetchedTestResults: any[] = [];
    logger.debug(
      `Fetching test reporter results for test plan ID: ${testPlanId}, project name: ${projectName}`
    );
    logger.debug(`Selected suite IDs: ${selectedSuiteIds}`);
    try {
      const ticketsDataProvider = new TicketsDataProvider(this.orgUrl, this.token);
      logger.debug(`Fetching Plan info for test plan ID: ${testPlanId}, project name: ${projectName}`);
      const plan = await this.fetchTestPlanName(testPlanId, projectName);
      logger.debug(`Fetching Test suites for test plan ID: ${testPlanId}, project name: ${projectName}`);
      const suites = await this.fetchTestSuites(testPlanId, projectName, selectedSuiteIds, true);
      logger.debug(`Fetching test data for test plan ID: ${testPlanId}, project name: ${projectName}`);
      const testData = await this.fetchTestData(suites, projectName, testPlanId, allowCrossTestPlan);
      logger.debug(`Fetching Run results for test data, project name: ${projectName}`);
      const isQueryMode = linkedQueryRequest.linkedQueryMode === 'query';
      if (isQueryMode) {
        // Fetch associated items
        await ticketsDataProvider.GetQueryResultsFromWiql(
          linkedQueryRequest.testAssociatedQuery.wiql.href,
          true,
          this.testToAssociatedItemMap
        );
        this.querySelectedColumns = linkedQueryRequest.testAssociatedQuery.columns;
      }

      const runResults = await this.fetchAllResultDataTestReporter(
        testData,
        projectName,
        selectedFields,
        isQueryMode,
        includeAllHistory
      );
      logger.debug(`Aligning steps with iterations for test reporter results`);
      const testReporterData = this.alignStepsWithIterationsTestReporter(
        testData,
        runResults,
        selectedFields,
        !enableRunTestCaseFilter
      );
      // Apply filters sequentially based on enabled flags
      let filteredResults = testReporterData;
      if (errorFilterMode !== 'none') {
        switch (errorFilterMode) {
          case 'onlyTestCaseResult':
            logger.debug(`Filtering test reporter results for only test case result`);
            filteredResults = filteredResults.filter(
              (result: any) =>
                result &&
                ((result.testCase?.comment && result.testCase.comment !== '') ||
                  result.testCase?.result?.resultMessage?.includes('Failed'))
            );
            break;
          case 'onlyTestStepsResult':
            logger.debug(`Filtering test reporter results for only test steps result`);
            filteredResults = filteredResults.filter(
              (result: any) =>
                result &&
                ((result.stepComments && result.stepComments !== '') || result.stepStatus === 'Failed')
            );
            break;
          case 'both':
            logger.debug(`Filtering test reporter results for both test case and test steps result`);
            filteredResults = filteredResults
              ?.filter(
                (result: any) =>
                  result &&
                  ((result.testCase?.comment && result.testCase.comment !== '') ||
                    result.testCase?.result?.resultMessage?.includes('Failed'))
              )
              ?.filter(
                (result: any) =>
                  result &&
                  ((result.stepComments && result.stepComments !== '') || result.stepStatus === 'Failed')
              );
            break;
          default:
            break;
        }
      }

      // filter: Test step run status
      if (enableRunStepStatusFilter) {
        filteredResults = filteredResults.filter((result) => !this.isNotRunStep(result));
      }

      // Use the final filtered results
      fetchedTestResults.push({
        contentControl: 'test-reporter-table',
        customName: `Test Results for ${plan}`,
        data: filteredResults || [],
        skin: 'test-reporter-table',
      });

      return fetchedTestResults;
    } catch (error: any) {
      logger.error(`Error during getTestReporterResults: ${error.message}`);
    }
  }

  /**
   * Fetches and processes a flat list of test reporter rows for a given test plan.
   * Returns raw row data suitable for post-processing/formatting by the caller.
   */
  public async getTestReporterFlatResults(
    testPlanId: string,
    projectName: string,
    selectedSuiteIds: number[] | undefined,
    selectedFields: string[] = [],
    includeAllHistory: boolean = false
  ) {
    logger.debug(
      `Fetching flat test reporter results for test plan ID: ${testPlanId}, project name: ${projectName}`
    );
    try {
      const planName = await this.fetchTestPlanName(testPlanId, projectName);
      const suites = await this.fetchTestSuites(testPlanId, projectName, selectedSuiteIds, true);
      const testData = await this.fetchTestData(suites, projectName, testPlanId, false);
      const runResults = await this.fetchAllResultDataTestReporter(
        testData,
        projectName,
        selectedFields,
        false,
        includeAllHistory
      );

      const rows = this.alignStepsWithIterationsFlatReport(
        testData,
        runResults,
        true,
        testPlanId,
        planName
      );

      return { planId: testPlanId, planName, rows: rows || [] };
    } catch (error: any) {
      logger.error(`Error during getTestReporterFlatResults: ${error.message}`);
      return { planId: testPlanId, planName: '', rows: [] };
    }
  }

  /**
   * Builds MEWP L2 requirement coverage rows for audit reporting.
   * Rows are one Requirement-TestCase pair; uncovered requirements are emitted with empty test-case columns.
   */
  public async getMewpL2CoverageFlatResults(
    testPlanId: string,
    projectName: string,
    selectedSuiteIds: number[] | undefined,
    linkedQueryRequest?: any,
    options?: MewpCoverageRequestOptions
  ): Promise<MewpCoverageFlatPayload> {
    const defaultPayload: MewpCoverageFlatPayload = {
      sheetName: `MEWP L2 Coverage - Plan ${testPlanId}`,
      columnOrder: [...ResultDataProvider.MEWP_L2_COVERAGE_COLUMNS],
      rows: [],
    };

    try {
      const planName = await this.fetchTestPlanName(testPlanId, projectName);
      const testData = await this.fetchMewpScopedTestData(
        testPlanId,
        projectName,
        selectedSuiteIds,
        !!options?.useRelFallback
      );

      const allRequirements = await this.fetchMewpL2Requirements(projectName);
      if (allRequirements.length === 0) {
        return {
          ...defaultPayload,
          sheetName: this.buildMewpCoverageSheetName(planName, testPlanId),
        };
      }

      const linkedRequirementsByTestCase = await this.buildLinkedRequirementsByTestCase(
        allRequirements,
        testData,
        projectName
      );
      const scopedRequirementKeys = await this.resolveMewpRequirementScopeKeysFromQuery(
        linkedQueryRequest,
        allRequirements,
        linkedRequirementsByTestCase
      );
      const requirements = this.collapseMewpRequirementFamilies(
        allRequirements,
        scopedRequirementKeys?.size ? scopedRequirementKeys : undefined
      );
      const requirementSapWbsByBaseKey = this.buildRequirementSapWbsByBaseKey(allRequirements);
      const externalBugsByTestCase = await this.loadExternalBugsByTestCase(options?.externalBugsFile);
      const externalL3L4ByBaseKey = await this.loadExternalL3L4ByBaseKey(
        options?.externalL3L4File,
        requirementSapWbsByBaseKey
      );
      const hasExternalBugsFile = !!String(
        options?.externalBugsFile?.name ||
          options?.externalBugsFile?.objectName ||
          options?.externalBugsFile?.text ||
          options?.externalBugsFile?.url ||
          ''
      ).trim();
      const hasExternalL3L4File = !!String(
        options?.externalL3L4File?.name ||
          options?.externalL3L4File?.objectName ||
          options?.externalL3L4File?.text ||
          options?.externalL3L4File?.url ||
          ''
      ).trim();
      const externalBugLinksCount = [...externalBugsByTestCase.values()].reduce(
        (sum, items) => sum + (Array.isArray(items) ? items.length : 0),
        0
      );
      const externalL3L4LinksCount = [...externalL3L4ByBaseKey.values()].reduce(
        (sum, items) => sum + (Array.isArray(items) ? items.length : 0),
        0
      );
      logger.info(
        `MEWP coverage external ingestion summary: ` +
          `bugsFileProvided=${hasExternalBugsFile} bugsTestCases=${externalBugsByTestCase.size} bugsLinks=${externalBugLinksCount}; ` +
          `l3l4FileProvided=${hasExternalL3L4File} l3l4BaseKeys=${externalL3L4ByBaseKey.size} l3l4Links=${externalL3L4LinksCount}`
      );
      if (hasExternalBugsFile && externalBugLinksCount === 0) {
        logger.warn(
          `MEWP coverage: external bugs file was provided but produced 0 links. ` +
            `Check SR/test-case/state values in ingestion logs.`
        );
      }
      if (hasExternalL3L4File && externalL3L4LinksCount === 0) {
        logger.warn(
          `MEWP coverage: external L3/L4 file was provided but produced 0 links. ` +
            `Check SR/AREA34/state/SAPWBS filters in ingestion logs.`
        );
      }
      if (requirements.length === 0) {
        return {
          ...defaultPayload,
          sheetName: this.buildMewpCoverageSheetName(planName, testPlanId),
        };
      }

      const requirementIndex: MewpRequirementIndex = new Map();
      const observedTestCaseIdsByRequirement = new Map<string, Set<number>>();
      const requirementKeys = new Set<string>();
      requirements.forEach((requirement) => {
        const key = String(requirement?.baseKey || '').trim();
        if (!key) return;
        requirementKeys.add(key);
      });

      const parsedDefinitionStepsByTestCase = new Map<number, TestSteps[]>();
      const testCaseStepsXmlMap = this.buildTestCaseStepsXmlMap(testData);
      const runResults = await this.fetchAllResultDataTestReporter(testData, projectName, [], false, false);
      for (const runResult of runResults) {
        const testCaseId = this.extractMewpTestCaseId(runResult);
        const rawActionResults = Array.isArray(runResult?.iteration?.actionResults)
          ? runResult.iteration.actionResults.filter((item: any) => !item?.isSharedStepTitle)
          : [];
        const actionResults = rawActionResults.sort((a: any, b: any) =>
          this.compareActionResults(
            String(a?.stepPosition || a?.stepIdentifier || ''),
            String(b?.stepPosition || b?.stepIdentifier || '')
          )
        );
        const hasExecutedRun =
          Number(runResult?.lastRunId || 0) > 0 && Number(runResult?.lastResultId || 0) > 0;

        if (actionResults.length > 0) {
          this.accumulateRequirementCountsFromActionResults(
            actionResults,
            testCaseId,
            requirementKeys,
            requirementIndex,
            observedTestCaseIdsByRequirement
          );
          continue;
        }

        // Do not force "not run" from definition steps when a run exists:
        // some runs may have missing/unmapped actionResults.
        if (hasExecutedRun) {
          continue;
        }

        if (!Number.isFinite(testCaseId)) continue;
        if (!parsedDefinitionStepsByTestCase.has(testCaseId)) {
          const stepsXml = testCaseStepsXmlMap.get(testCaseId) || '';
          const parsed =
            stepsXml && String(stepsXml).trim() !== ''
              ? await this.testStepParserHelper.parseTestSteps(stepsXml, new Map<number, number>())
              : [];
          parsedDefinitionStepsByTestCase.set(testCaseId, parsed);
        }

        const definitionSteps = parsedDefinitionStepsByTestCase.get(testCaseId) || [];
        const fallbackActionResults = definitionSteps
          .filter((step) => !step?.isSharedStepTitle)
          .sort((a, b) =>
            this.compareActionResults(String(a?.stepPosition || ''), String(b?.stepPosition || ''))
          )
          .map((step) => ({
            stepPosition: step?.stepPosition,
            expected: step?.expected,
            outcome: 'Unspecified',
          }));

        this.accumulateRequirementCountsFromActionResults(
          fallbackActionResults,
          testCaseId,
          requirementKeys,
          requirementIndex,
          observedTestCaseIdsByRequirement
        );
      }

      const requirementBaseKeys = new Set<string>(
        requirements.map((item) => String(item?.baseKey || '').trim()).filter((item) => !!item)
      );
      const externalL3L4BaseKeys = new Set<string>([...externalL3L4ByBaseKey.keys()]);
      const externalL3L4OverlapKeys = [...externalL3L4BaseKeys].filter((key) => requirementBaseKeys.has(key));
      const failedRequirementBaseKeys = new Set<string>();
      const failedTestCaseIds = new Set<number>();
      for (const [requirementBaseKey, byTestCase] of requirementIndex.entries()) {
        for (const [testCaseId, counts] of byTestCase.entries()) {
          if (Number(counts?.failed || 0) > 0) {
            failedRequirementBaseKeys.add(requirementBaseKey);
            failedTestCaseIds.add(testCaseId);
          }
        }
      }
      const externalBugTestCaseIds = new Set<number>([...externalBugsByTestCase.keys()]);
      const externalBugFailedTestCaseOverlap = [...externalBugTestCaseIds].filter((id) =>
        failedTestCaseIds.has(id)
      );
      const externalBugBaseKeys = new Set<string>();
      for (const bugs of externalBugsByTestCase.values()) {
        for (const bug of bugs || []) {
          const key = String(bug?.requirementBaseKey || '').trim();
          if (key) externalBugBaseKeys.add(key);
        }
      }
      const externalBugRequirementOverlap = [...externalBugBaseKeys].filter((key) =>
        requirementBaseKeys.has(key)
      );
      const externalBugFailedRequirementOverlap = [...externalBugBaseKeys].filter((key) =>
        failedRequirementBaseKeys.has(key)
      );
      logger.info(
        `MEWP coverage join diagnostics: requirementBaseKeys=${requirementBaseKeys.size} ` +
          `failedRequirementBaseKeys=${failedRequirementBaseKeys.size} failedTestCases=${failedTestCaseIds.size}; ` +
          `externalL3L4BaseKeys=${externalL3L4BaseKeys.size} externalL3L4Overlap=${externalL3L4OverlapKeys.length}; ` +
          `externalBugTestCases=${externalBugTestCaseIds.size} externalBugFailedTestCaseOverlap=${externalBugFailedTestCaseOverlap.length}; ` +
          `externalBugBaseKeys=${externalBugBaseKeys.size} externalBugRequirementOverlap=${externalBugRequirementOverlap.length} ` +
          `externalBugFailedRequirementOverlap=${externalBugFailedRequirementOverlap.length}`
      );
      if (externalL3L4BaseKeys.size > 0 && externalL3L4OverlapKeys.length === 0) {
        const sampleReq = [...requirementBaseKeys].slice(0, 5);
        const sampleExt = [...externalL3L4BaseKeys].slice(0, 5);
        logger.warn(
          `MEWP coverage join diagnostics: no L3/L4 key overlap found. ` +
            `sampleRequirementKeys=${sampleReq.join(', ')} sampleExternalL3L4Keys=${sampleExt.join(', ')}`
        );
      }
      if (externalBugBaseKeys.size > 0 && externalBugRequirementOverlap.length === 0) {
        const sampleReq = [...requirementBaseKeys].slice(0, 5);
        const sampleExt = [...externalBugBaseKeys].slice(0, 5);
        logger.warn(
          `MEWP coverage join diagnostics: no bug requirement-key overlap found. ` +
            `sampleRequirementKeys=${sampleReq.join(', ')} sampleExternalBugKeys=${sampleExt.join(', ')}`
        );
      }
      if (externalBugTestCaseIds.size > 0 && externalBugFailedTestCaseOverlap.length === 0) {
        logger.warn(
          `MEWP coverage join diagnostics: no overlap between external bug test cases and failed test cases. ` +
            `Bug rows remain empty because bugs are shown only for failed L2s.`
        );
      }

      const rows = this.buildMewpCoverageRows(
        requirements,
        requirementIndex,
        observedTestCaseIdsByRequirement,
        linkedRequirementsByTestCase,
        externalL3L4ByBaseKey,
        externalBugsByTestCase
      );
      const coverageRowStats = rows.reduce(
        (acc, row) => {
          const hasBug = Number(row?.['Bug ID'] || 0) > 0;
          const hasL3 = String(row?.['L3 REQ ID'] || '').trim() !== '';
          const hasL4 = String(row?.['L4 REQ ID'] || '').trim() !== '';
          if (hasBug) acc.bugRows += 1;
          if (hasL3) acc.l3Rows += 1;
          if (hasL4) acc.l4Rows += 1;
          if (!hasBug && !hasL3 && !hasL4) acc.baseOnlyRows += 1;
          return acc;
        },
        { bugRows: 0, l3Rows: 0, l4Rows: 0, baseOnlyRows: 0 }
      );
      logger.info(
        `MEWP coverage output summary: requirements=${requirements.length} rows=${rows.length} ` +
          `bugRows=${coverageRowStats.bugRows} l3Rows=${coverageRowStats.l3Rows} ` +
          `l4Rows=${coverageRowStats.l4Rows} baseOnlyRows=${coverageRowStats.baseOnlyRows}`
      );

      return {
        sheetName: this.buildMewpCoverageSheetName(planName, testPlanId),
        columnOrder: [...ResultDataProvider.MEWP_L2_COVERAGE_COLUMNS],
        rows,
      };
    } catch (error: any) {
      logger.error(`Error during getMewpL2CoverageFlatResults: ${error.message}`);
      if (error instanceof MewpExternalFileValidationError) {
        throw error;
      }
      return defaultPayload;
    }
  }

  public async getMewpInternalValidationFlatResults(
    testPlanId: string,
    projectName: string,
    selectedSuiteIds: number[] | undefined,
    linkedQueryRequest?: any,
    options?: MewpInternalValidationRequestOptions
  ): Promise<MewpInternalValidationFlatPayload> {
    const defaultPayload: MewpInternalValidationFlatPayload = {
      sheetName: `MEWP Internal Validation - Plan ${testPlanId}`,
      columnOrder: [...ResultDataProvider.INTERNAL_VALIDATION_COLUMNS],
      rows: [],
    };

    try {
      const planName = await this.fetchTestPlanName(testPlanId, projectName);
      const testData = await this.fetchMewpScopedTestData(
        testPlanId,
        projectName,
        selectedSuiteIds,
        !!options?.useRelFallback
      );
      const allRequirements = await this.fetchMewpL2Requirements(projectName);
      const linkedRequirementsByTestCase = await this.buildLinkedRequirementsByTestCase(
        allRequirements,
        testData,
        projectName
      );
      const scopedRequirementKeys = await this.resolveMewpRequirementScopeKeysFromQuery(
        linkedQueryRequest,
        allRequirements,
        linkedRequirementsByTestCase
      );
      const requirementFamilies = this.buildRequirementFamilyMap(
        allRequirements,
        scopedRequirementKeys?.size ? scopedRequirementKeys : undefined
      );

      const rows: MewpInternalValidationRow[] = [];
      const stepsXmlByTestCase = this.buildTestCaseStepsXmlMap(testData);
      const testCaseTitleMap = this.buildMewpTestCaseTitleMap(testData);
      const allTestCaseIds = new Set<number>();
      for (const suite of testData || []) {
        const testCasesItems = Array.isArray(suite?.testCasesItems) ? suite.testCasesItems : [];
        for (const testCase of testCasesItems) {
          const id = Number(testCase?.workItem?.id || testCase?.testCaseId || testCase?.id || 0);
          if (Number.isFinite(id) && id > 0) allTestCaseIds.add(id);
        }
        const testPointsItems = Array.isArray(suite?.testPointsItems) ? suite.testPointsItems : [];
        for (const testPoint of testPointsItems) {
          const id = Number(testPoint?.testCaseId || testPoint?.testCase?.id || 0);
          if (Number.isFinite(id) && id > 0) allTestCaseIds.add(id);
        }
      }

      const validL2BaseKeys = new Set<string>([...requirementFamilies.keys()]);

      for (const testCaseId of [...allTestCaseIds].sort((a, b) => a - b)) {
        const stepsXml = stepsXmlByTestCase.get(testCaseId) || '';
        const parsedSteps =
          stepsXml && String(stepsXml).trim() !== ''
            ? await this.testStepParserHelper.parseTestSteps(stepsXml, new Map<number, number>())
            : [];
        const mentionEntries = this.extractRequirementMentionsFromExpectedSteps(parsedSteps, true);
        const mentionedL2Only = new Set<string>();
        const mentionedCodeFirstStep = new Map<string, string>();
        const mentionedBaseFirstStep = new Map<string, string>();
        for (const mentionEntry of mentionEntries) {
          const scopeFilteredCodes =
            scopedRequirementKeys?.size && mentionEntry.codes.size > 0
              ? [...mentionEntry.codes].filter((code) => scopedRequirementKeys.has(this.toRequirementKey(code)))
              : [...mentionEntry.codes];
          for (const code of scopeFilteredCodes) {
            const baseKey = this.toRequirementKey(code);
            if (!baseKey) continue;
            if (validL2BaseKeys.has(baseKey)) {
              mentionedL2Only.add(code);
              if (!mentionedCodeFirstStep.has(code)) {
                mentionedCodeFirstStep.set(code, mentionEntry.stepRef);
              }
              if (!mentionedBaseFirstStep.has(baseKey)) {
                mentionedBaseFirstStep.set(baseKey, mentionEntry.stepRef);
              }
            }
          }
        }

        const mentionedBaseKeys = new Set<string>(
          [...mentionedL2Only].map((code) => this.toRequirementKey(code)).filter((code) => !!code)
        );

        const expectedFamilyCodes = new Set<string>();
        for (const baseKey of mentionedBaseKeys) {
          const familyCodes = requirementFamilies.get(baseKey);
          if (familyCodes?.size) {
            familyCodes.forEach((code) => expectedFamilyCodes.add(code));
          } else {
            for (const code of mentionedL2Only) {
              if (this.toRequirementKey(code) === baseKey) expectedFamilyCodes.add(code);
            }
          }
        }

        const linkedFullCodesRaw = linkedRequirementsByTestCase.get(testCaseId)?.fullCodes || new Set<string>();
        const linkedFullCodes =
          scopedRequirementKeys?.size && linkedFullCodesRaw.size > 0
            ? new Set<string>(
                [...linkedFullCodesRaw].filter((code) =>
                  scopedRequirementKeys.has(this.toRequirementKey(code))
                )
              )
            : linkedFullCodesRaw;
        const linkedBaseKeys = new Set<string>(
          [...linkedFullCodes].map((code) => this.toRequirementKey(code)).filter((code) => !!code)
        );

        const missingMentioned = [...mentionedL2Only].filter((code) => {
          const baseKey = this.toRequirementKey(code);
          if (!baseKey) return false;
          const hasSpecificSuffix = /-\d+$/.test(code);
          if (hasSpecificSuffix) return !linkedFullCodes.has(code);
          return !linkedBaseKeys.has(baseKey);
        });
        const missingFamily = [...expectedFamilyCodes].filter((code) => !linkedFullCodes.has(code));
        const extraLinked = [...linkedFullCodes].filter((code) => !expectedFamilyCodes.has(code));
        const mentionedButNotLinkedByStep = new Map<string, Set<string>>();
        const appendMentionedButNotLinked = (requirementId: string, stepRef: string) => {
          const normalizedRequirementId = this.normalizeMewpRequirementCodeWithSuffix(requirementId);
          if (!normalizedRequirementId) return;
          const normalizedStepRef = String(stepRef || 'Step ?').trim() || 'Step ?';
          if (!mentionedButNotLinkedByStep.has(normalizedStepRef)) {
            mentionedButNotLinkedByStep.set(normalizedStepRef, new Set<string>());
          }
          mentionedButNotLinkedByStep.get(normalizedStepRef)!.add(normalizedRequirementId);
        };

        const sortedMissingMentioned = [...new Set(missingMentioned)].sort((a, b) => a.localeCompare(b));
        const sortedMissingFamily = [...new Set(missingFamily)].sort((a, b) => a.localeCompare(b));
        for (const code of sortedMissingMentioned) {
          const stepRef = mentionedCodeFirstStep.get(code) || 'Step ?';
          appendMentionedButNotLinked(code, stepRef);
        }
        for (const code of sortedMissingFamily) {
          const baseKey = this.toRequirementKey(code);
          const stepRef = mentionedBaseFirstStep.get(baseKey) || 'Step ?';
          appendMentionedButNotLinked(code, stepRef);
        }

        const sortedExtraLinked = [...new Set(extraLinked)]
          .map((code) => this.normalizeMewpRequirementCodeWithSuffix(code))
          .filter((code) => !!code)
          .sort((a, b) => a.localeCompare(b));

        const parseStepOrder = (stepRef: string): number => {
          const match = /step\s+(\d+)/i.exec(String(stepRef || ''));
          const parsed = Number(match?.[1] || Number.POSITIVE_INFINITY);
          return Number.isFinite(parsed) ? parsed : Number.POSITIVE_INFINITY;
        };
        const mentionedButNotLinked = [...mentionedButNotLinkedByStep.entries()]
          .sort((a, b) => {
            const stepOrderA = parseStepOrder(a[0]);
            const stepOrderB = parseStepOrder(b[0]);
            if (stepOrderA !== stepOrderB) return stepOrderA - stepOrderB;
            return String(a[0]).localeCompare(String(b[0]));
          })
          .map(([stepRef, requirementIds]) => {
            const requirementList = [...requirementIds].sort((a, b) => a.localeCompare(b));
            return `${stepRef}: ${requirementList.join(', ')}`;
          })
          .join('; ');
        const linkedButNotMentioned = sortedExtraLinked.join('; ');
        const validationStatus: 'Pass' | 'Fail' =
          mentionedButNotLinked || linkedButNotMentioned ? 'Fail' : 'Pass';

        rows.push({
          'Test Case ID': testCaseId,
          'Test Case Title': String(testCaseTitleMap.get(testCaseId) || '').trim(),
          'Mentioned but Not Linked': mentionedButNotLinked,
          'Linked but Not Mentioned': linkedButNotMentioned,
          'Validation Status': validationStatus,
        });
      }

      return {
        sheetName: this.buildInternalValidationSheetName(planName, testPlanId),
        columnOrder: [...ResultDataProvider.INTERNAL_VALIDATION_COLUMNS],
        rows,
      };
    } catch (error: any) {
      logger.error(`Error during getMewpInternalValidationFlatResults: ${error.message}`);
      return defaultPayload;
    }
  }

  public async validateMewpExternalFiles(options: {
    externalBugsFile?: MewpExternalFileRef | null;
    externalL3L4File?: MewpExternalFileRef | null;
  }): Promise<MewpExternalFilesValidationResponse> {
    const response: MewpExternalFilesValidationResponse = { valid: true };
    const validateOne = async (
      file: MewpExternalFileRef | null | undefined,
      tableType: 'bugs' | 'l3l4'
    ): Promise<MewpExternalTableValidationResult | undefined> => {
      const sourceName = String(file?.name || file?.objectName || file?.text || file?.url || '').trim();
      if (!sourceName) return undefined;

      try {
        const { rows, meta } = await this.mewpExternalTableUtils.loadExternalTableRowsWithMeta(
          file,
          tableType
        );
        return {
          tableType,
          sourceName: meta.sourceName,
          valid: true,
          headerRow: meta.headerRow,
          matchedRequiredColumns: meta.matchedRequiredColumns,
          totalRequiredColumns: meta.totalRequiredColumns,
          missingRequiredColumns: [],
          rowCount: rows.length,
          message: 'File schema is valid',
        };
      } catch (error: any) {
        if (error instanceof MewpExternalFileValidationError) {
          return {
            ...error.details,
            valid: false,
          };
        }
        return {
          tableType,
          sourceName: sourceName || tableType,
          valid: false,
          headerRow: '',
          matchedRequiredColumns: 0,
          totalRequiredColumns: this.mewpExternalTableUtils.getRequiredColumnCount(tableType),
          missingRequiredColumns: [],
          rowCount: 0,
          message: String(error?.message || error || 'Unknown validation error'),
        };
      }
    };

    response.bugs = await validateOne(options?.externalBugsFile, 'bugs');
    response.l3l4 = await validateOne(options?.externalL3L4File, 'l3l4');
    response.valid = [response.bugs, response.l3l4].filter(Boolean).every((item) => !!item?.valid);
    return response;
  }

  /**
   * Mapping each attachment to a proper URL for downloading it
   * @param runResults Array of run results
   */
  public mapAttachmentsUrl(runResults: any[], project: string) {
    return runResults.map((result) => {
      if (!result.iteration) {
        return result;
      }
      const { iteration, analysisAttachments, ...restResult } = result;
      //add downloadUri field for each attachment
      const baseDownloadUrl = `${this.orgUrl}${project}/_apis/test/runs/${result.lastRunId}/results/${result.lastResultId}/attachments`;
      if (iteration && iteration.attachments?.length > 0) {
        const { attachments, actionResults, ...restOfIteration } = iteration;
        const attachmentPathToIndexMap: Map<string, number> =
          this.CreateAttachmentPathIndexMap(actionResults);

        const mappedAttachments = attachments.map((attachment: any) => ({
          ...attachment,
          stepNo: attachmentPathToIndexMap.has(attachment.actionPath)
            ? attachmentPathToIndexMap.get(attachment.actionPath)
            : undefined,
          downloadUrl: `${baseDownloadUrl}/${attachment.id}/${attachment.name}`,
        }));

        restResult.iteration = { ...restOfIteration, attachments: mappedAttachments };
      }
      if (analysisAttachments && analysisAttachments.length > 0) {
        restResult.analysisAttachments = analysisAttachments.map((attachment: any) => ({
          ...attachment,
          downloadUrl: `${baseDownloadUrl}/${attachment.id}/${attachment.fileName}`,
        }));
      }

      return { ...restResult };
    });
  }

  private buildMewpCoverageSheetName(planName: string, testPlanId: string): string {
    const suffix = String(planName || '').trim() || `Plan ${testPlanId}`;
    return `MEWP L2 Coverage - ${suffix}`;
  }

  private buildInternalValidationSheetName(planName: string, testPlanId: string): string {
    const suffix = String(planName || '').trim() || `Plan ${testPlanId}`;
    return `MEWP Internal Validation - ${suffix}`;
  }

  private createMewpCoverageRow(
    requirement: Pick<MewpL2RequirementFamily, 'requirementId' | 'title' | 'subSystem' | 'responsibility'>,
    runStatus: MewpRunStatus,
    bug: MewpCoverageBugCell,
    linkedL3L4: MewpCoverageL3L4Cell
  ): MewpCoverageRow {
    const l2ReqId = this.formatMewpCustomerId(requirement.requirementId);
    const l2ReqTitle = this.toMewpComparableText(requirement.title);
    const l2SubSystem = this.toMewpComparableText(requirement.subSystem);

    return {
      'L2 REQ ID': l2ReqId,
      'L2 REQ Title': l2ReqTitle,
      'L2 SubSystem': l2SubSystem,
      'L2 Run Status': runStatus,
      'Bug ID': Number.isFinite(Number(bug?.id)) && Number(bug?.id) > 0 ? Number(bug?.id) : '',
      'Bug Title': String(bug?.title || '').trim(),
      'Bug Responsibility': String(bug?.responsibility || '').trim(),
      'L3 REQ ID': String(linkedL3L4?.l3Id || '').trim(),
      'L3 REQ Title': String(linkedL3L4?.l3Title || '').trim(),
      'L4 REQ ID': String(linkedL3L4?.l4Id || '').trim(),
      'L4 REQ Title': String(linkedL3L4?.l4Title || '').trim(),
    };
  }

  private createEmptyMewpCoverageBugCell(): MewpCoverageBugCell {
    return { id: '' as '', title: '', responsibility: '' };
  }

  private createEmptyMewpCoverageL3L4Cell(): MewpCoverageL3L4Cell {
    return { l3Id: '', l3Title: '', l4Id: '', l4Title: '' };
  }

  private buildMewpCoverageL3L4Rows(links: MewpL3L4Link[]): MewpCoverageL3L4Cell[] {
    const sorted = [...(links || [])].sort((a, b) => {
      if (a.level !== b.level) return a.level === 'L3' ? -1 : 1;
      return String(a.id || '').localeCompare(String(b.id || ''));
    });

    const rows: MewpCoverageL3L4Cell[] = [];
    for (const item of sorted) {
      const isL3 = item.level === 'L3';
      rows.push({
        l3Id: isL3 ? String(item?.id || '').trim() : '',
        l3Title: isL3 ? String(item?.title || '').trim() : '',
        l4Id: isL3 ? '' : String(item?.id || '').trim(),
        l4Title: isL3 ? '' : String(item?.title || '').trim(),
      });
    }
    return rows;
  }

  private formatMewpCustomerId(rawValue: string): string {
    const normalized = this.normalizeMewpRequirementCode(this.toMewpComparableText(rawValue));
    if (normalized) return normalized;

    const onlyDigits = String(rawValue || '').replace(/\D/g, '');
    if (onlyDigits) return `SR${onlyDigits}`;
    return '';
  }

  private buildMewpCoverageRows(
    requirements: MewpL2RequirementFamily[],
    requirementIndex: MewpRequirementIndex,
    observedTestCaseIdsByRequirement: Map<string, Set<number>>,
    linkedRequirementsByTestCase: MewpLinkedRequirementsByTestCase,
    l3l4ByBaseKey: Map<string, MewpL3L4Link[]>,
    externalBugsByTestCase: Map<number, MewpBugLink[]>
  ): MewpCoverageRow[] {
    const rows: MewpCoverageRow[] = [];
    const linkedByRequirement = this.invertBaseRequirementLinks(linkedRequirementsByTestCase);
    for (const requirement of requirements) {
      const key = String(requirement?.baseKey || this.toRequirementKey(requirement.requirementId) || '').trim();
      const linkedTestCaseIds = (requirement?.linkedTestCaseIds || []).filter(
        (id) => Number.isFinite(id) && Number(id) > 0
      );
      const linkedByTestCase = key ? Array.from(linkedByRequirement.get(key) || []) : [];
      const observedTestCaseIds = key
        ? Array.from(observedTestCaseIdsByRequirement.get(key) || [])
        : [];

      const testCaseIds = Array.from(
        new Set<number>([...linkedTestCaseIds, ...linkedByTestCase, ...observedTestCaseIds])
      ).sort((a, b) => a - b);

      let totalPassed = 0;
      let totalFailed = 0;
      let totalNotRun = 0;
      const aggregatedBugs = new Map<number, MewpBugLink>();

      for (const testCaseId of testCaseIds) {
        const summary = key
          ? requirementIndex.get(key)?.get(testCaseId) || { passed: 0, failed: 0, notRun: 0 }
          : { passed: 0, failed: 0, notRun: 0 };
        totalPassed += summary.passed;
        totalFailed += summary.failed;
        totalNotRun += summary.notRun;

        if (summary.failed > 0) {
          const externalBugs = externalBugsByTestCase.get(testCaseId) || [];
          for (const bug of externalBugs) {
            const bugBaseKey = String(bug?.requirementBaseKey || '').trim();
            if (bugBaseKey && bugBaseKey !== key) continue;
            const bugId = Number(bug?.id || 0);
            if (!Number.isFinite(bugId) || bugId <= 0) continue;
            aggregatedBugs.set(bugId, {
              ...bug,
              responsibility: this.resolveCoverageBugResponsibility(
                String(bug?.responsibility || ''),
                requirement
              ),
            });
          }
        }
      }

      const runStatus = this.resolveMewpL2RunStatus({
        passed: totalPassed,
        failed: totalFailed,
        notRun: totalNotRun,
        hasAnyTestCase: testCaseIds.length > 0,
      });

      const bugsForRows =
        runStatus === 'Fail'
          ? Array.from(aggregatedBugs.values()).sort((a, b) => a.id - b.id)
          : [];
      const l3l4ForRows = [...(l3l4ByBaseKey.get(key) || [])];

      const bugRows: MewpCoverageBugCell[] =
        bugsForRows.length > 0
          ? bugsForRows
          : [];
      const l3l4Rows: MewpCoverageL3L4Cell[] = this.buildMewpCoverageL3L4Rows(l3l4ForRows);

      if (bugRows.length === 0 && l3l4Rows.length === 0) {
        rows.push(
          this.createMewpCoverageRow(
            requirement,
            runStatus,
            this.createEmptyMewpCoverageBugCell(),
            this.createEmptyMewpCoverageL3L4Cell()
          )
        );
        continue;
      }

      for (const bug of bugRows) {
        rows.push(
          this.createMewpCoverageRow(
            requirement,
            runStatus,
            bug,
            this.createEmptyMewpCoverageL3L4Cell()
          )
        );
      }

      for (const linkedL3L4 of l3l4Rows) {
        rows.push(
          this.createMewpCoverageRow(
            requirement,
            runStatus,
            this.createEmptyMewpCoverageBugCell(),
            linkedL3L4
          )
        );
      }
    }

    return rows;
  }

  private resolveCoverageBugResponsibility(
    rawResponsibility: string,
    requirement: Pick<MewpL2RequirementFamily, 'responsibility'>
  ): string {
    const direct = String(rawResponsibility || '').trim();
    if (direct && direct.toLowerCase() !== 'unknown') return direct;

    const requirementResponsibility = String(requirement?.responsibility || '')
      .trim()
      .toUpperCase();
    if (requirementResponsibility === 'ESUK') return 'ESUK';
    if (requirementResponsibility === 'IL' || requirementResponsibility === 'ELISRA') return 'Elisra';

    return direct || 'Unknown';
  }

  private resolveMewpL2RunStatus(input: {
    passed: number;
    failed: number;
    notRun: number;
    hasAnyTestCase: boolean;
  }): MewpRunStatus {
    if ((input?.failed || 0) > 0) return 'Fail';
    if ((input?.notRun || 0) > 0) return 'Not Run';
    if ((input?.passed || 0) > 0) return 'Pass';
    return input?.hasAnyTestCase ? 'Not Run' : 'Not Run';
  }

  private async fetchMewpScopedTestData(
    testPlanId: string,
    projectName: string,
    selectedSuiteIds: number[] | undefined,
    useRelFallback: boolean
  ): Promise<any[]> {
    if (!useRelFallback) {
      const suites = await this.fetchTestSuites(testPlanId, projectName, selectedSuiteIds, true);
      return this.fetchTestData(suites, projectName, testPlanId, false);
    }

    const selectedSuites = await this.fetchTestSuites(testPlanId, projectName, selectedSuiteIds, true);
    const selectedRel = this.resolveMaxRelNumberFromSuites(selectedSuites);
    if (selectedRel <= 0) {
      return this.fetchTestData(selectedSuites, projectName, testPlanId, false);
    }

    const allSuites = await this.fetchTestSuites(testPlanId, projectName, undefined, true);
    const relScopedSuites = allSuites.filter((suite) => {
      const rel = this.extractRelNumberFromSuite(suite);
      return rel > 0 && rel <= selectedRel;
    });
    const suitesForFetch = relScopedSuites.length > 0 ? relScopedSuites : selectedSuites;
    const rawTestData = await this.fetchTestData(suitesForFetch, projectName, testPlanId, false);
    return this.reduceToLatestRelRunPerTestCase(rawTestData);
  }

  private extractRelNumberFromSuite(suite: any): number {
    const candidates = [
      suite?.suiteName,
      suite?.parentSuiteName,
      suite?.suitePath,
      suite?.testGroupName,
    ];
    const pattern = /(?:^|[^a-z0-9])rel\s*([0-9]+)/i;
    for (const item of candidates) {
      const match = pattern.exec(String(item || ''));
      if (!match) continue;
      const parsed = Number(match[1]);
      if (Number.isFinite(parsed) && parsed > 0) {
        return parsed;
      }
    }
    return 0;
  }

  private resolveMaxRelNumberFromSuites(suites: any[]): number {
    let maxRel = 0;
    for (const suite of suites || []) {
      const rel = this.extractRelNumberFromSuite(suite);
      if (rel > maxRel) maxRel = rel;
    }
    return maxRel;
  }

  private reduceToLatestRelRunPerTestCase(testData: any[]): any[] {
    type Candidate = {
      point: any;
      rel: number;
      runId: number;
      resultId: number;
      hasRun: boolean;
    };

    const candidatesByTestCase = new Map<number, Candidate[]>();
    const testCaseDefinitionById = new Map<number, any>();

    for (const suite of testData || []) {
      const rel = this.extractRelNumberFromSuite(suite);
      const testPointsItems = Array.isArray(suite?.testPointsItems) ? suite.testPointsItems : [];
      const testCasesItems = Array.isArray(suite?.testCasesItems) ? suite.testCasesItems : [];

      for (const testCase of testCasesItems) {
        const testCaseId = Number(testCase?.workItem?.id || testCase?.testCaseId || testCase?.id || 0);
        if (!Number.isFinite(testCaseId) || testCaseId <= 0) continue;
        if (!testCaseDefinitionById.has(testCaseId)) {
          testCaseDefinitionById.set(testCaseId, testCase);
        }
      }

      for (const point of testPointsItems) {
        const testCaseId = Number(point?.testCaseId || point?.testCase?.id || 0);
        if (!Number.isFinite(testCaseId) || testCaseId <= 0) continue;

        const runId = Number(point?.lastRunId || 0);
        const resultId = Number(point?.lastResultId || 0);
        const hasRun = runId > 0 && resultId > 0;
        if (!candidatesByTestCase.has(testCaseId)) {
          candidatesByTestCase.set(testCaseId, []);
        }
        candidatesByTestCase.get(testCaseId)!.push({
          point,
          rel,
          runId,
          resultId,
          hasRun,
        });
      }
    }

    const selectedPoints: any[] = [];
    const selectedTestCaseIds = new Set<number>();
    for (const [testCaseId, candidates] of candidatesByTestCase.entries()) {
      const sorted = [...candidates].sort((a, b) => {
        if (a.hasRun !== b.hasRun) return a.hasRun ? -1 : 1;
        if (a.rel !== b.rel) return b.rel - a.rel;
        if (a.runId !== b.runId) return b.runId - a.runId;
        return b.resultId - a.resultId;
      });
      const chosen = sorted[0];
      if (!chosen?.point) continue;
      selectedPoints.push(chosen.point);
      selectedTestCaseIds.add(testCaseId);
    }

    const selectedTestCases: any[] = [];
    for (const testCaseId of selectedTestCaseIds) {
      const definition = testCaseDefinitionById.get(testCaseId);
      if (definition) {
        selectedTestCases.push(definition);
      }
    }

    return [
      {
        testSuiteId: 'MEWP_REL_SCOPED',
        suiteId: 'MEWP_REL_SCOPED',
        suiteName: 'MEWP Rel Scoped',
        parentSuiteId: '',
        parentSuiteName: '',
        suitePath: 'MEWP Rel Scoped',
        testGroupName: 'MEWP Rel Scoped',
        testPointsItems: selectedPoints,
        testCasesItems: selectedTestCases,
      },
    ];
  }

  private async loadExternalBugsByTestCase(
    externalBugsFile: MewpExternalFileRef | null | undefined
  ): Promise<Map<number, MewpBugLink[]>> {
    return this.mewpExternalIngestionUtils.loadExternalBugsByTestCase(externalBugsFile, {
      toComparableText: (value) => this.toMewpComparableText(value),
      toRequirementKey: (value) => this.toRequirementKey(value),
      resolveBugResponsibility: (fields) => this.resolveBugResponsibility(fields),
      isExternalStateInScope: (value, itemType) => this.isExternalStateInScope(value, itemType),
      isExcludedL3L4BySapWbs: (value) => this.isExcludedL3L4BySapWbs(value),
      resolveRequirementSapWbsByBaseKey: () => '',
    });
  }

  private async loadExternalL3L4ByBaseKey(
    externalL3L4File: MewpExternalFileRef | null | undefined,
    requirementSapWbsByBaseKey: Map<string, string> = new Map<string, string>()
  ): Promise<Map<string, MewpL3L4Link[]>> {
    return this.mewpExternalIngestionUtils.loadExternalL3L4ByBaseKey(externalL3L4File, {
      toComparableText: (value) => this.toMewpComparableText(value),
      toRequirementKey: (value) => this.toRequirementKey(value),
      resolveBugResponsibility: (fields) => this.resolveBugResponsibility(fields),
      isExternalStateInScope: (value, itemType) => this.isExternalStateInScope(value, itemType),
      isExcludedL3L4BySapWbs: (value) => this.isExcludedL3L4BySapWbs(value),
      resolveRequirementSapWbsByBaseKey: (baseKey) => String(requirementSapWbsByBaseKey.get(baseKey) || ''),
    });
  }

  private buildRequirementSapWbsByBaseKey(
    requirements: Array<Pick<MewpL2RequirementWorkItem, 'baseKey' | 'responsibility'>>
  ): Map<string, string> {
    const out = new Map<string, string>();
    for (const requirement of requirements || []) {
      const baseKey = String(requirement?.baseKey || '').trim();
      if (!baseKey) continue;

      const normalized = this.resolveMewpResponsibility(this.toMewpComparableText(requirement?.responsibility));
      if (!normalized) continue;

      const existing = out.get(baseKey) || '';
      // Keep ESUK as dominant if conflicting values are ever present across family items.
      if (existing === 'ESUK') continue;
      if (normalized === 'ESUK' || !existing) {
        out.set(baseKey, normalized);
      }
    }
    return out;
  }

  private isExternalStateInScope(value: string, itemType: 'bug' | 'requirement'): boolean {
    const normalized = String(value || '').trim().toLowerCase();
    if (!normalized) return true;

    // TFS/ADO processes usually don't expose a literal "Open" state.
    // Keep non-terminal states, exclude terminal states.
    const terminalStates = new Set<string>([
      'resolved',
      'closed',
      'done',
      'completed',
      'complete',
      'removed',
      'rejected',
      'cancelled',
      'canceled',
      'obsolete',
    ]);

    if (terminalStates.has(normalized)) return false;

    // Bug-specific terminal variants often used in custom processes.
    if (itemType === 'bug') {
      if (normalized === 'fixed') return false;
    }

    return true;
  }

  private invertBaseRequirementLinks(
    linkedRequirementsByTestCase: MewpLinkedRequirementsByTestCase
  ): Map<string, Set<number>> {
    const out = new Map<string, Set<number>>();
    for (const [testCaseId, links] of linkedRequirementsByTestCase.entries()) {
      for (const baseKey of links?.baseKeys || []) {
        if (!out.has(baseKey)) out.set(baseKey, new Set<number>());
        out.get(baseKey)!.add(testCaseId);
      }
    }
    return out;
  }

  private buildMewpTestCaseTitleMap(testData: any[]): Map<number, string> {
    const map = new Map<number, string>();

    const readTitleFromWorkItemFields = (workItemFields: any): string => {
      if (!Array.isArray(workItemFields)) return '';
      for (const field of workItemFields) {
        const keyCandidates = [field?.key, field?.name, field?.referenceName, field?.id]
          .map((item) => String(item || '').toLowerCase().trim());
        const isTitleField =
          keyCandidates.includes('system.title') || keyCandidates.includes('title');
        if (!isTitleField) continue;
        const value = this.toMewpComparableText(field?.value);
        if (value) return value;
      }
      return '';
    };

    for (const suite of testData || []) {
      const testPointsItems = Array.isArray(suite?.testPointsItems) ? suite.testPointsItems : [];
      for (const point of testPointsItems) {
        const pointTestCaseId = Number(point?.testCaseId || point?.testCase?.id);
        if (!Number.isFinite(pointTestCaseId) || pointTestCaseId <= 0 || map.has(pointTestCaseId)) continue;
        const pointTitle = this.toMewpComparableText(point?.testCaseName || point?.testCase?.name);
        if (pointTitle) map.set(pointTestCaseId, pointTitle);
      }

      const testCasesItems = Array.isArray(suite?.testCasesItems) ? suite.testCasesItems : [];
      for (const testCase of testCasesItems) {
        const id = Number(testCase?.workItem?.id || testCase?.testCaseId || testCase?.id);
        if (!Number.isFinite(id) || id <= 0 || map.has(id)) continue;
        const fromDirectFields = this.toMewpComparableText(
          testCase?.testCaseName || testCase?.name || testCase?.workItem?.name
        );
        if (fromDirectFields) {
          map.set(id, fromDirectFields);
          continue;
        }
        const fromWorkItemField = readTitleFromWorkItemFields(testCase?.workItem?.workItemFields);
        if (fromWorkItemField) {
          map.set(id, fromWorkItemField);
        }
      }
    }

    return map;
  }

  private extractMewpTestCaseId(runResult: any): number {
    const testCaseId = Number(runResult?.testCaseId || runResult?.testCase?.id || 0);
    return Number.isFinite(testCaseId) ? testCaseId : 0;
  }

  private buildTestCaseStepsXmlMap(testData: any[]): Map<number, string> {
    const map = new Map<number, string>();
    for (const suite of testData || []) {
      const testCasesItems = Array.isArray(suite?.testCasesItems) ? suite.testCasesItems : [];
      for (const testCase of testCasesItems) {
        const id = Number(testCase?.workItem?.id);
        if (!Number.isFinite(id)) continue;
        if (map.has(id)) continue;
        const fields = testCase?.workItem?.workItemFields;
        const stepsXml = this.extractStepsXmlFromWorkItemFields(fields);
        if (stepsXml) {
          map.set(id, stepsXml);
        }
      }
    }
    return map;
  }

  private extractStepsXmlFromWorkItemFields(workItemFields: any): string {
    if (!Array.isArray(workItemFields)) return '';
    const isStepsKey = (name: string) => {
      const normalized = String(name || '').toLowerCase().trim();
      return normalized === 'steps' || normalized === 'microsoft.vsts.tcm.steps';
    };

    for (const field of workItemFields) {
      const keyCandidates = [field?.key, field?.name, field?.referenceName, field?.id];
      const hasStepsKey = keyCandidates.some((candidate) => isStepsKey(String(candidate || '')));
      if (!hasStepsKey) continue;
      const value = String(field?.value || '').trim();
      if (value) return value;
    }

    return '';
  }

  private classifyRequirementStepOutcome(outcome: any): 'passed' | 'failed' | 'notRun' {
    const normalized = String(outcome || '')
      .trim()
      .toLowerCase();
    if (normalized === 'passed') return 'passed';
    if (normalized === 'failed') return 'failed';
    return 'notRun';
  }

  private accumulateRequirementCountsFromActionResults(
    actionResults: any[],
    testCaseId: number,
    requirementKeys: Set<string>,
    counters: Map<string, Map<number, { passed: number; failed: number; notRun: number }>>,
    observedTestCaseIdsByRequirement: Map<string, Set<number>>
  ) {
    if (!Number.isFinite(testCaseId) || testCaseId <= 0) return;
    const sortedResults = Array.isArray(actionResults) ? actionResults : [];
    let previousRequirementStepIndex = -1;

    for (let i = 0; i < sortedResults.length; i++) {
      const actionResult = sortedResults[i];
      if (actionResult?.isSharedStepTitle) continue;
      const requirementCodes = this.extractRequirementCodesFromText(actionResult?.expected || '');
      if (requirementCodes.size === 0) continue;

      const startIndex = previousRequirementStepIndex + 1;
      const status = this.resolveRequirementStatusForWindow(sortedResults, startIndex, i);
      previousRequirementStepIndex = i;

      for (const code of requirementCodes) {
        if (requirementKeys.size > 0 && !requirementKeys.has(code)) continue;
        if (!counters.has(code)) {
          counters.set(code, new Map<number, { passed: number; failed: number; notRun: number }>());
        }
        const perTestCaseCounters = counters.get(code)!;
        if (!perTestCaseCounters.has(testCaseId)) {
          perTestCaseCounters.set(testCaseId, { passed: 0, failed: 0, notRun: 0 });
        }

        if (!observedTestCaseIdsByRequirement.has(code)) {
          observedTestCaseIdsByRequirement.set(code, new Set<number>());
        }
        observedTestCaseIdsByRequirement.get(code)!.add(testCaseId);

        const counter = perTestCaseCounters.get(testCaseId)!;
        if (status === 'passed') counter.passed += 1;
        else if (status === 'failed') counter.failed += 1;
        else counter.notRun += 1;
      }
    }
  }

  private resolveRequirementStatusForWindow(
    actionResults: any[],
    startIndex: number,
    endIndex: number
  ): 'passed' | 'failed' | 'notRun' {
    let hasNotRun = false;
    for (let index = startIndex; index <= endIndex; index++) {
      const status = this.classifyRequirementStepOutcome(actionResults[index]?.outcome);
      if (status === 'failed') return 'failed';
      if (status === 'notRun') hasNotRun = true;
    }
    return hasNotRun ? 'notRun' : 'passed';
  }

  private extractRequirementCodesFromText(text: string): Set<string> {
    return this.extractRequirementCodesFromExpectedText(text, false);
  }

  private extractRequirementMentionsFromExpectedSteps(
    steps: TestSteps[],
    includeSuffix: boolean
  ): Array<{ stepRef: string; codes: Set<string> }> {
    const out: Array<{ stepRef: string; codes: Set<string> }> = [];
    const allSteps = Array.isArray(steps) ? steps : [];
    for (let index = 0; index < allSteps.length; index += 1) {
      const step = allSteps[index];
      if (step?.isSharedStepTitle) continue;
      const codes = this.extractRequirementCodesFromExpectedText(step?.expected || '', includeSuffix);
      if (codes.size === 0) continue;
      out.push({
        stepRef: this.resolveValidationStepReference(step, index),
        codes,
      });
    }
    return out;
  }

  private extractRequirementCodesFromExpectedSteps(steps: TestSteps[], includeSuffix: boolean): Set<string> {
    const out = new Set<string>();
    for (const step of Array.isArray(steps) ? steps : []) {
      if (step?.isSharedStepTitle) continue;
      const codes = this.extractRequirementCodesFromExpectedText(step?.expected || '', includeSuffix);
      codes.forEach((code) => out.add(code));
    }
    return out;
  }

  private extractRequirementCodesFromExpectedText(text: string, includeSuffix: boolean): Set<string> {
    const out = new Set<string>();
    const source = this.normalizeRequirementStepText(text);
    if (!source) return out;

    const tokens = source
      .split(';')
      .map((token) => String(token || '').trim())
      .filter((token) => token !== '');

    for (const token of tokens) {
      const candidates = this.extractRequirementCandidatesFromToken(token);
      for (const candidate of candidates) {
        const expandedTokens = this.expandRequirementTokenByComma(candidate);
        for (const expandedToken of expandedTokens) {
          if (!expandedToken || /vvrm/i.test(expandedToken)) continue;
          const normalized = this.normalizeRequirementCodeToken(expandedToken, includeSuffix);
          if (normalized) {
            out.add(normalized);
          }
        }
      }
    }

    return out;
  }

  private extractRequirementCandidatesFromToken(token: string): string[] {
    const source = String(token || '');
    if (!source) return [];
    const out = new Set<string>();
    const collectCandidates = (input: string, rejectTailPattern: RegExp) => {
      for (const match of input.matchAll(/SR\d{4,}(?:-\d+(?:,\d+)*)?/gi)) {
        const matchedValue = String(match?.[0] || '')
          .trim()
          .toUpperCase();
        if (!matchedValue) continue;
        const endIndex = Number(match?.index || 0) + matchedValue.length;
        const tail = String(input.slice(endIndex) || '');
        if (rejectTailPattern.test(tail)) continue;
        out.add(matchedValue);
      }
    };

    // Normal scan keeps punctuation context (" SR0817-V3.2 " -> reject via tail).
    collectCandidates(source, /^\s*(?:V\d|VVRM|-V\d)/i);

    // Compact scan preserves legacy support for spaced SR letters/digits
    // such as "S R 0 0 0 1" and HTML-fragmented tokens.
    const compactSource = source.replace(/\s+/g, '');
    if (compactSource && compactSource !== source) {
      collectCandidates(compactSource, /^(?:V\d|VVRM|-V\d)/i);
    }

    return [...out];
  }

  private expandRequirementTokenByComma(token: string): string[] {
    const compact = String(token || '').trim().toUpperCase();
    if (!compact) return [];

    const suffixBatchMatch = /^SR(\d{4,})-(\d+(?:,\d+)+)$/.exec(compact);
    if (suffixBatchMatch) {
      const base = String(suffixBatchMatch[1] || '').trim();
      const suffixes = String(suffixBatchMatch[2] || '')
        .split(',')
        .map((item) => String(item || '').trim())
        .filter((item) => /^\d+$/.test(item));
      return suffixes.map((suffix) => `SR${base}-${suffix}`);
    }

    return compact
      .split(',')
      .map((part) => String(part || '').trim())
      .filter((part) => !!part);
  }

  private normalizeRequirementCodeToken(token: string, includeSuffix: boolean): string {
    const compact = String(token || '')
      .trim()
      .replace(/[\u200B-\u200D\uFEFF]/g, '')
      .toUpperCase();
    if (!compact) return '';

    const pattern = includeSuffix ? /^SR(\d{4,})(?:-(\d+))?$/ : /^SR(\d{4,})(?:-\d+)?$/;
    const match = pattern.exec(compact);
    if (!match) return '';

    const baseDigits = String(match[1] || '').trim();
    if (!baseDigits) return '';

    if (includeSuffix && match[2]) {
      const suffixDigits = String(match[2] || '').trim();
      if (!suffixDigits) return '';
      return `SR${baseDigits}-${suffixDigits}`;
    }

    return `SR${baseDigits}`;
  }

  private normalizeRequirementStepText(text: string): string {
    const raw = String(text || '');
    if (!raw) return '';

    return raw
      .replace(/&nbsp;|&#160;|&#xA0;/gi, ' ')
      .replace(/&lt;/gi, '<')
      .replace(/&gt;/gi, '>')
      .replace(/&amp;/gi, '&')
      .replace(/&quot;/gi, '"')
      .replace(/&#39;|&apos;/gi, "'")
      .replace(/<[^>]*>/g, ' ')
      .replace(/[\u200B-\u200D\uFEFF]/g, '')
      .replace(/\s+/g, ' ');
  }

  private resolveValidationStepReference(step: TestSteps, index: number): string {
    const fromPosition = String(step?.stepPosition || '').trim();
    if (fromPosition) return `Step ${fromPosition}`;
    const fromId = String(step?.stepId || '').trim();
    if (fromId) return `Step ${fromId}`;
    return `Step ${index + 1}`;
  }

  private toRequirementKey(requirementId: string): string {
    return this.normalizeMewpRequirementCode(requirementId);
  }

  private async fetchMewpL2Requirements(projectName: string): Promise<MewpL2RequirementWorkItem[]> {
    const workItemTypeNames = await this.fetchMewpRequirementTypeNames(projectName);
    if (workItemTypeNames.length === 0) {
      return [];
    }

    const quotedTypeNames = workItemTypeNames
      .map((name) => `'${String(name).replace(/'/g, "''")}'`)
      .join(', ');
    const queryRequirementIds = async (l2AreaPath: string | null): Promise<number[]> => {
      const escapedAreaPath = l2AreaPath ? String(l2AreaPath).replace(/'/g, "''") : '';
      const areaFilter = escapedAreaPath ? `\n  AND [System.AreaPath] UNDER '${escapedAreaPath}'` : '';
      const wiql = `SELECT [System.Id]
	FROM WorkItems
	WHERE [System.TeamProject] = @project
	  AND [System.WorkItemType] IN (${quotedTypeNames})${areaFilter}
	ORDER BY [System.Id]`;
      const wiqlUrl = `${this.orgUrl}${projectName}/_apis/wit/wiql?api-version=7.1-preview.2`;
      const wiqlResponse = await TFSServices.postRequest(wiqlUrl, this.token, 'Post', { query: wiql }, null);
      const workItemRefs = Array.isArray(wiqlResponse?.data?.workItems) ? wiqlResponse.data.workItems : [];
      return workItemRefs
        .map((item: any) => Number(item?.id))
        .filter((id: number) => Number.isFinite(id));
    };

    const defaultL2AreaPath = `${String(projectName || '').trim()}\\Customer Requirements\\Level 2`;
    let requirementIds: number[] = [];
    try {
      requirementIds = await queryRequirementIds(defaultL2AreaPath);
    } catch (error: any) {
      logger.warn(
        `Could not apply MEWP L2 WIQL area-path optimization. Falling back to full requirement scope: ${
          error?.message || error
        }`
      );
    }
    if (requirementIds.length === 0) {
      requirementIds = await queryRequirementIds(null);
    }

    if (requirementIds.length === 0) {
      return [];
    }

    const workItems = await this.fetchWorkItemsByIds(projectName, requirementIds, true);
    const requirements = workItems.map((wi: any) => {
      const fields = wi?.fields || {};
      const requirementId = this.extractMewpRequirementIdentifier(fields, Number(wi?.id || 0));
      const areaPath = this.toMewpComparableText(fields?.['System.AreaPath']);
      return {
        workItemId: Number(wi?.id || 0),
        requirementId,
        baseKey: this.toRequirementKey(requirementId),
        title: this.toMewpComparableText(fields?.['System.Title'] || wi?.title),
        subSystem: this.deriveMewpSubSystem(fields),
        responsibility: this.deriveMewpResponsibility(fields),
        linkedTestCaseIds: this.extractLinkedTestCaseIdsFromRequirement(wi?.relations || []),
        relatedWorkItemIds: this.extractLinkedWorkItemIdsFromRelations(wi?.relations || []),
        areaPath,
      };
    });

    return requirements
      .filter((item) => {
        if (!item.baseKey) return false;
        if (!item.areaPath) return true;
        return this.isMewpL2AreaPath(item.areaPath);
      })
      .sort((a, b) => String(a.requirementId).localeCompare(String(b.requirementId)));
  }

  private isMewpL2AreaPath(areaPath: string): boolean {
    const normalized = String(areaPath || '')
      .trim()
      .toLowerCase()
      .replace(/\//g, '\\');
    if (!normalized) return false;
    return normalized.includes('\\customer requirements\\level 2');
  }

  private collapseMewpRequirementFamilies(
    requirements: MewpL2RequirementWorkItem[],
    scopedRequirementKeys?: Set<string>
  ): MewpL2RequirementFamily[] {
    const families = new Map<
      string,
      {
        representative: MewpL2RequirementWorkItem;
        score: number;
        linkedTestCaseIds: Set<number>;
      }
    >();

    const calcScore = (item: MewpL2RequirementWorkItem) => {
      const requirementId = String(item?.requirementId || '').trim();
      const areaPath = String(item?.areaPath || '')
        .trim()
        .toLowerCase();
      let score = 0;
      if (/^SR\d+$/i.test(requirementId)) score += 6;
      if (areaPath.includes('\\customer requirements\\level 2')) score += 3;
      if (!areaPath.includes('\\mop')) score += 2;
      if (String(item?.title || '').trim()) score += 1;
      if (String(item?.subSystem || '').trim()) score += 1;
      if (String(item?.responsibility || '').trim()) score += 1;
      return score;
    };

    for (const requirement of requirements || []) {
      const baseKey = String(requirement?.baseKey || '').trim();
      if (!baseKey) continue;
      if (scopedRequirementKeys?.size && !scopedRequirementKeys.has(baseKey)) continue;

      if (!families.has(baseKey)) {
        families.set(baseKey, {
          representative: requirement,
          score: calcScore(requirement),
          linkedTestCaseIds: new Set<number>(),
        });
      }
      const family = families.get(baseKey)!;
      const score = calcScore(requirement);
      if (score > family.score) {
        family.representative = requirement;
        family.score = score;
      }
      for (const testCaseId of requirement?.linkedTestCaseIds || []) {
        if (Number.isFinite(testCaseId) && Number(testCaseId) > 0) {
          family.linkedTestCaseIds.add(Number(testCaseId));
        }
      }
    }

    return [...families.entries()]
      .map(([baseKey, family]) => ({
        requirementId: String(family?.representative?.requirementId || baseKey),
        baseKey,
        title: String(family?.representative?.title || ''),
        subSystem: String(family?.representative?.subSystem || ''),
        responsibility: String(family?.representative?.responsibility || ''),
        linkedTestCaseIds: [...family.linkedTestCaseIds].sort((a, b) => a - b),
      }))
      .sort((a, b) => String(a.requirementId).localeCompare(String(b.requirementId)));
  }

  private buildRequirementFamilyMap(
    requirements: Array<Pick<MewpL2RequirementWorkItem, 'requirementId' | 'baseKey'>>,
    scopedRequirementKeys?: Set<string>
  ): Map<string, Set<string>> {
    const familyMap = new Map<string, Set<string>>();
    for (const requirement of requirements || []) {
      const baseKey = String(requirement?.baseKey || '').trim();
      if (!baseKey) continue;
      if (scopedRequirementKeys?.size && !scopedRequirementKeys.has(baseKey)) continue;
      const fullCode = this.normalizeMewpRequirementCodeWithSuffix(requirement?.requirementId || '');
      if (!fullCode) continue;
      if (!familyMap.has(baseKey)) familyMap.set(baseKey, new Set<string>());
      familyMap.get(baseKey)!.add(fullCode);
    }
    return familyMap;
  }

  private async buildLinkedRequirementsByTestCase(
    requirements: Array<
      Pick<MewpL2RequirementWorkItem, 'workItemId' | 'requirementId' | 'baseKey' | 'linkedTestCaseIds'>
    >,
    testData: any[],
    projectName: string
  ): Promise<MewpLinkedRequirementsByTestCase> {
    const map: MewpLinkedRequirementsByTestCase = new Map();
    const ensure = (testCaseId: number) => {
      if (!map.has(testCaseId)) {
        map.set(testCaseId, {
          baseKeys: new Set<string>(),
          fullCodes: new Set<string>(),
          bugIds: new Set<number>(),
        });
      }
      return map.get(testCaseId)!;
    };

    const requirementById = new Map<number, { baseKey: string; fullCode: string }>();
    for (const requirement of requirements || []) {
      const workItemId = Number(requirement?.workItemId || 0);
      const baseKey = String(requirement?.baseKey || '').trim();
      const fullCode = this.normalizeMewpRequirementCodeWithSuffix(requirement?.requirementId || '');
      if (workItemId > 0 && baseKey && fullCode) {
        requirementById.set(workItemId, { baseKey, fullCode });
      }

      for (const testCaseIdRaw of requirement?.linkedTestCaseIds || []) {
        const testCaseId = Number(testCaseIdRaw);
        if (!Number.isFinite(testCaseId) || testCaseId <= 0 || !baseKey || !fullCode) continue;
        const entry = ensure(testCaseId);
        entry.baseKeys.add(baseKey);
        entry.fullCodes.add(fullCode);
      }
    }

    const testCaseIds = new Set<number>();
    for (const suite of testData || []) {
      const testCasesItems = Array.isArray(suite?.testCasesItems) ? suite.testCasesItems : [];
      for (const testCase of testCasesItems) {
        const id = Number(testCase?.workItem?.id || testCase?.testCaseId || testCase?.id || 0);
        if (Number.isFinite(id) && id > 0) testCaseIds.add(id);
      }
      const testPointsItems = Array.isArray(suite?.testPointsItems) ? suite.testPointsItems : [];
      for (const testPoint of testPointsItems) {
        const id = Number(testPoint?.testCaseId || testPoint?.testCase?.id || 0);
        if (Number.isFinite(id) && id > 0) testCaseIds.add(id);
      }
    }

    const relatedIdsByTestCase = new Map<number, Set<number>>();
    const allRelatedIds = new Set<number>();
    if (testCaseIds.size > 0) {
      const testCaseWorkItems = await this.fetchWorkItemsByIds(projectName, [...testCaseIds], true);
      for (const workItem of testCaseWorkItems || []) {
        const testCaseId = Number(workItem?.id || 0);
        if (!Number.isFinite(testCaseId) || testCaseId <= 0) continue;
        const relations = Array.isArray(workItem?.relations) ? workItem.relations : [];
        if (!relatedIdsByTestCase.has(testCaseId)) relatedIdsByTestCase.set(testCaseId, new Set<number>());
        for (const relation of relations) {
          const linkedWorkItemId = this.extractLinkedWorkItemIdFromRelation(relation);
          if (!linkedWorkItemId) continue;
          relatedIdsByTestCase.get(testCaseId)!.add(linkedWorkItemId);
          allRelatedIds.add(linkedWorkItemId);

          if (this.isTestCaseToRequirementRelation(relation) && requirementById.has(linkedWorkItemId)) {
            const linkedRequirement = requirementById.get(linkedWorkItemId)!;
            const entry = ensure(testCaseId);
            entry.baseKeys.add(linkedRequirement.baseKey);
            entry.fullCodes.add(linkedRequirement.fullCode);
          }
        }
      }
    }

    if (allRelatedIds.size > 0) {
      const relatedWorkItems = await this.fetchWorkItemsByIds(projectName, [...allRelatedIds], false);
      const typeById = new Map<number, string>();
      for (const workItem of relatedWorkItems || []) {
        const id = Number(workItem?.id || 0);
        if (!Number.isFinite(id) || id <= 0) continue;
        const type = String(workItem?.fields?.['System.WorkItemType'] || '')
          .trim()
          .toLowerCase();
        typeById.set(id, type);
      }

      for (const [testCaseId, ids] of relatedIdsByTestCase.entries()) {
        const entry = ensure(testCaseId);
        for (const linkedId of ids) {
          const linkedType = typeById.get(linkedId) || '';
          if (linkedType === 'bug') {
            entry.bugIds.add(linkedId);
          }
        }
      }
    }

    return map;
  }

  private async resolveMewpRequirementScopeKeysFromQuery(
    linkedQueryRequest: any,
    requirements: Array<Pick<MewpL2RequirementWorkItem, 'workItemId' | 'baseKey'>>,
    linkedRequirementsByTestCase: MewpLinkedRequirementsByTestCase
  ): Promise<Set<string> | undefined> {
    const mode = String(linkedQueryRequest?.linkedQueryMode || '')
      .trim()
      .toLowerCase();
    const wiqlHref = String(linkedQueryRequest?.testAssociatedQuery?.wiql?.href || '').trim();
    if (mode !== 'query' || !wiqlHref) return undefined;

    try {
      const queryResult = await TFSServices.getItemContent(wiqlHref, this.token);
      const queryIds = new Set<number>();
      if (Array.isArray(queryResult?.workItems)) {
        for (const workItem of queryResult.workItems) {
          const id = Number(workItem?.id || 0);
          if (Number.isFinite(id) && id > 0) queryIds.add(id);
        }
      }
      if (Array.isArray(queryResult?.workItemRelations)) {
        for (const relation of queryResult.workItemRelations) {
          const sourceId = Number(relation?.source?.id || 0);
          const targetId = Number(relation?.target?.id || 0);
          if (Number.isFinite(sourceId) && sourceId > 0) queryIds.add(sourceId);
          if (Number.isFinite(targetId) && targetId > 0) queryIds.add(targetId);
        }
      }

      if (queryIds.size === 0) return undefined;

      const reqIdToBaseKey = new Map<number, string>();
      for (const requirement of requirements || []) {
        const id = Number(requirement?.workItemId || 0);
        const baseKey = String(requirement?.baseKey || '').trim();
        if (id > 0 && baseKey) reqIdToBaseKey.set(id, baseKey);
      }

      const scopedKeys = new Set<string>();
      for (const queryId of queryIds) {
        if (reqIdToBaseKey.has(queryId)) {
          scopedKeys.add(reqIdToBaseKey.get(queryId)!);
          continue;
        }

        const linked = linkedRequirementsByTestCase.get(queryId);
        if (!linked?.baseKeys?.size) continue;
        linked.baseKeys.forEach((baseKey) => scopedKeys.add(baseKey));
      }

      return scopedKeys.size > 0 ? scopedKeys : undefined;
    } catch (error: any) {
      logger.warn(`Could not resolve MEWP query scope: ${error?.message || error}`);
      return undefined;
    }
  }

  private isTestCaseToRequirementRelation(relation: any): boolean {
    const rel = String(relation?.rel || '')
      .trim()
      .toLowerCase();
    if (!rel) return false;
    return rel.includes('testedby-reverse') || (rel.includes('tests') && rel.includes('reverse'));
  }

  private extractLinkedWorkItemIdFromRelation(relation: any): number {
    const url = String(relation?.url || '');
    const match = /\/workItems\/(\d+)/i.exec(url);
    if (!match) return 0;
    const parsed = Number(match[1]);
    return Number.isFinite(parsed) ? parsed : 0;
  }

  private async fetchMewpRequirementTypeNames(projectName: string): Promise<string[]> {
    try {
      const url = `${this.orgUrl}${projectName}/_apis/wit/workitemtypes?api-version=7.1-preview.2`;
      const result = await TFSServices.getItemContent(url, this.token);
      const values = Array.isArray(result?.value) ? result.value : [];
      const matched = values
        .map((item: any) => String(item?.name || ''))
        .filter((name: string) => /requirement/i.test(name) || /^epic$/i.test(name));
      const unique = Array.from(new Set<string>(matched));
      if (unique.length > 0) {
        return unique;
      }
    } catch (error: any) {
      logger.debug(`Could not fetch MEWP work item types, using defaults: ${error?.message || error}`);
    }

    return ['Requirement', 'Epic'];
  }

  private async fetchWorkItemsByIds(
    projectName: string,
    workItemIds: number[],
    includeRelations: boolean
  ): Promise<any[]> {
    const ids = [...new Set(workItemIds.filter((id) => Number.isFinite(id)))];
    if (ids.length === 0) return [];

    const CHUNK_SIZE = 200;
    const allItems: any[] = [];

    for (let i = 0; i < ids.length; i += CHUNK_SIZE) {
      const chunk = ids.slice(i, i + CHUNK_SIZE);
      const idsQuery = chunk.join(',');
      const expandParam = includeRelations ? '&$expand=relations' : '';
      const url = `${this.orgUrl}${projectName}/_apis/wit/workitems?ids=${idsQuery}${expandParam}&api-version=7.1-preview.3`;
      const response = await TFSServices.getItemContent(url, this.token);
      const values = Array.isArray(response?.value) ? response.value : [];
      allItems.push(...values);
    }

    return allItems;
  }

  private extractLinkedTestCaseIdsFromRequirement(relations: any[]): number[] {
    const out = new Set<number>();
    for (const relation of Array.isArray(relations) ? relations : []) {
      const rel = String(relation?.rel || '')
        .trim()
        .toLowerCase();
      const isRequirementToTestLink = rel.includes('testedby') || rel.includes('.tests');
      if (!isRequirementToTestLink) continue;

      const url = String(relation?.url || '');
      const match = /\/workItems\/(\d+)/i.exec(url);
      if (!match) continue;
      const id = Number(match[1]);
      if (Number.isFinite(id)) out.add(id);
    }
    return [...out].sort((a, b) => a - b);
  }

  private extractLinkedWorkItemIdsFromRelations(relations: any[]): number[] {
    const out = new Set<number>();
    for (const relation of Array.isArray(relations) ? relations : []) {
      const url = String(relation?.url || '');
      const match = /\/workItems\/(\d+)/i.exec(url);
      if (!match) continue;
      const id = Number(match[1]);
      if (Number.isFinite(id) && id > 0) out.add(id);
    }
    return [...out].sort((a, b) => a - b);
  }

  private extractMewpRequirementIdentifier(fields: Record<string, any>, fallbackWorkItemId: number): string {
    const entries = Object.entries(fields || {});

    // First pass: only trusted identifier-like fields.
    const strictHints = [
      'customerid',
      'customer id',
      'customerrequirementid',
      'requirementid',
      'externalid',
      'srid',
      'sapwbsid',
    ];
    for (const [key, value] of entries) {
      const normalizedKey = String(key || '').toLowerCase();
      if (!strictHints.some((hint) => normalizedKey.includes(hint))) continue;

      const valueAsString = this.toMewpComparableText(value);
      if (!valueAsString) continue;
      const normalized = this.normalizeMewpRequirementCodeWithSuffix(valueAsString);
      if (normalized) return normalized;
    }

    // Second pass: weaker hints, but still key-based only.
    const looseHints = ['customer', 'requirement', 'external', 'sapwbs', 'sr'];
    for (const [key, value] of entries) {
      const normalizedKey = String(key || '').toLowerCase();
      if (!looseHints.some((hint) => normalizedKey.includes(hint))) continue;

      const valueAsString = this.toMewpComparableText(value);
      if (!valueAsString) continue;
      const normalized = this.normalizeMewpRequirementCodeWithSuffix(valueAsString);
      if (normalized) return normalized;
    }

    // Optional fallback from title only (avoid scanning all fields and accidental SR matches).
    const title = this.toMewpComparableText(fields?.['System.Title']);
    const titleCode = this.normalizeMewpRequirementCodeWithSuffix(title);
    if (titleCode) return titleCode;

    return fallbackWorkItemId ? `SR${fallbackWorkItemId}` : '';
  }

  private deriveMewpResponsibility(fields: Record<string, any>): string {
    const explicitSapWbs = this.toMewpComparableText(fields?.['Custom.SAPWBS']);
    const fromExplicitSapWbs = this.resolveMewpResponsibility(explicitSapWbs);
    if (fromExplicitSapWbs) return fromExplicitSapWbs;
    if (explicitSapWbs) return explicitSapWbs;

    const explicitSapWbsByLabel = this.toMewpComparableText(fields?.['SAPWBS']);
    const fromExplicitLabel = this.resolveMewpResponsibility(explicitSapWbsByLabel);
    if (fromExplicitLabel) return fromExplicitLabel;
    if (explicitSapWbsByLabel) return explicitSapWbsByLabel;

    const areaPath = this.toMewpComparableText(fields?.['System.AreaPath']);
    const fromAreaPath = this.resolveMewpResponsibility(areaPath);
    if (fromAreaPath) return fromAreaPath;

    const keyHints = ['sapwbs', 'responsibility', 'owner'];
    for (const [key, value] of Object.entries(fields || {})) {
      const normalizedKey = String(key || '').toLowerCase();
      if (!keyHints.some((hint) => normalizedKey.includes(hint))) continue;
      const resolved = this.resolveMewpResponsibility(this.toMewpComparableText(value));
      if (resolved) return resolved;
    }

    return '';
  }

  private deriveMewpSubSystem(fields: Record<string, any>): string {
    const directCandidates = [
      fields?.['Custom.SubSystem'],
      fields?.['Custom.Subsystem'],
      fields?.['SubSystem'],
      fields?.['Subsystem'],
      fields?.['subSystem'],
    ];
    for (const candidate of directCandidates) {
      const value = this.toMewpComparableText(candidate);
      if (value) return value;
    }

    const keyHints = ['subsystem', 'sub system', 'sub_system'];
    for (const [key, value] of Object.entries(fields || {})) {
      const normalizedKey = String(key || '').toLowerCase();
      if (!keyHints.some((hint) => normalizedKey.includes(hint))) continue;
      const resolved = this.toMewpComparableText(value);
      if (resolved) return resolved;
    }

    return '';
  }

  private resolveBugResponsibility(fields: Record<string, any>): string {
    const sapWbsRaw = this.toMewpComparableText(fields?.['Custom.SAPWBS'] || fields?.['SAPWBS']);
    const fromSapWbs = this.resolveMewpResponsibility(sapWbsRaw);
    if (fromSapWbs === 'ESUK') return 'ESUK';
    if (fromSapWbs === 'IL') return 'Elisra';

    const areaPathRaw = this.toMewpComparableText(fields?.['System.AreaPath']);
    const fromAreaPath = this.resolveMewpResponsibility(areaPathRaw);
    if (fromAreaPath === 'ESUK') return 'ESUK';
    if (fromAreaPath === 'IL') return 'Elisra';

    return 'Unknown';
  }

  private resolveMewpResponsibility(value: string): string {
    const raw = String(value || '').trim();
    if (!raw) return '';

    const rawUpper = raw.toUpperCase();
    if (rawUpper === 'ESUK') return 'ESUK';
    if (rawUpper === 'IL') return 'IL';

    const normalizedPath = raw
      .toLowerCase()
      .replace(/\//g, '\\')
      .replace(/\\+/g, '\\')
      .trim();

    if (normalizedPath.endsWith('\\atp\\esuk') || normalizedPath === 'atp\\esuk') return 'ESUK';
    if (normalizedPath.endsWith('\\atp') || normalizedPath === 'atp') return 'IL';

    return '';
  }

  private isExcludedL3L4BySapWbs(value: string): boolean {
    const responsibility = this.resolveMewpResponsibility(this.toMewpComparableText(value));
    return responsibility === 'ESUK';
  }

  private normalizeMewpRequirementCode(value: string): string {
    const text = String(value || '').trim();
    if (!text) return '';
    const match = /\bSR[\s\-_]*([0-9]+)\b/i.exec(text);
    if (!match) return '';
    return `SR${match[1]}`;
  }

  private normalizeMewpRequirementCodeWithSuffix(value: string): string {
    const text = String(value || '').trim();
    if (!text) return '';
    const compact = text.replace(/\s+/g, '');
    const match = /^SR(\d+)(?:-(\d+))?$/i.exec(compact);
    if (!match) return '';
    if (match[2]) return `SR${match[1]}-${match[2]}`;
    return `SR${match[1]}`;
  }

  private toMewpComparableText(value: any): string {
    if (value === null || value === undefined) return '';
    if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') {
      return String(value).trim();
    }
    if (typeof value === 'object') {
      const displayName = (value as any).displayName;
      if (displayName) return String(displayName).trim();
      const name = (value as any).name;
      if (name) return String(name).trim();
      const uniqueName = (value as any).uniqueName;
      if (uniqueName) return String(uniqueName).trim();
      const objectValue = (value as any).value;
      if (objectValue !== undefined && objectValue !== null) return String(objectValue).trim();
    }
    return String(value).trim();
  }

  private async fetchTestPlanName(testPlanId: string, teamProject: string): Promise<string> {
    try {
      const url = `${this.orgUrl}${teamProject}/_apis/testplan/Plans/${testPlanId}?api-version=5.1`;
      const testPlan = await TFSServices.getItemContent(url, this.token);
      return testPlan.name;
    } catch (error: any) {
      logger.error(`Error during fetching Test Plan Name: ${error.message}`);
      return '';
    }
  }

  /**
   * Fetches test suites for a given test plan and project, optionally filtering by selected suite IDs.
   *
   * @param testPlanId - The ID of the test plan to fetch suites for.
   * @param projectName - The name of the project containing the test plan.
   * @param selectedSuiteIds - An optional array of suite IDs to filter the results by.
   * @param isHierarchyGroupName - A flag indicating whether to build the test group name hierarchically. Defaults to `true`.
   * @returns A promise that resolves to an array of objects containing `testSuiteId` and `testGroupName`.
   *          Returns an empty array if no test suites are found or an error occurs.
   * @throws Will log an error message if fetching test suites fails.
   */
  private async fetchTestSuites(
    testPlanId: string,
    projectName: string,
    selectedSuiteIds?: number[],
    isHierarchyGroupName: boolean = true
  ): Promise<any[]> {
    try {
      const treeUrl = `${this.orgUrl}${projectName}/_apis/testplan/Plans/${testPlanId}/Suites?asTreeView=true`;
      const { value: treeTestSuites, count: treeCount } = await TFSServices.getItemContent(
        treeUrl,
        this.token
      );

      if (treeCount === 0) throw new Error('No test suites found');

      const flatTestSuites = this.flattenSuites(treeTestSuites);
      const filteredSuites = this.filterSuites(flatTestSuites, selectedSuiteIds);
      const suiteMap = this.createSuiteMap(treeTestSuites);

      return filteredSuites.map((testSuite: any) => {
        const parentSuite = testSuite.parentSuite;
        return {
          testSuiteId: testSuite.id,
          suiteId: testSuite.id,
          suiteName: testSuite.name,
          parentSuiteId: parentSuite?.id,
          parentSuiteName: parentSuite?.name,
          suitePath: this.buildSuitePath(testSuite.id, suiteMap),
          testGroupName: this.buildTestGroupName(testSuite.id, suiteMap, isHierarchyGroupName),
        };
      });
    } catch (error: any) {
      logger.error(`Error during fetching Test Suites: ${error.message}`);
      return [];
    }
  }

  /**
   * Flattens a hierarchical suite structure into a single-level array.
   */
  private flattenSuites(suites: any[]): any[] {
    const flatSuites: any[] = [];
    const flatten = (suites: any[]) => {
      suites.forEach((suite: any) => {
        flatSuites.push(suite);
        if (suite.children) flatten(suite.children);
      });
    };
    flatten(suites);
    return flatSuites;
  }

  /**
   * Filters test suites based on the selected suite IDs.
   */
  private filterSuites(testSuites: any[], selectedSuiteIds?: number[]): any[] {
    return selectedSuiteIds
      ? testSuites.filter((suite) => selectedSuiteIds.includes(suite.id))
      : testSuites.filter((suite) => suite.parentSuite);
  }

  /**
   * Creates a quick-lookup map of suites by their IDs.
   */
  private createSuiteMap(suites: any[]): Map<number, any> {
    const suiteMap = new Map<number, any>();
    const addToMap = (suites: any[]) => {
      suites.forEach((suite: any) => {
        suiteMap.set(suite.id, suite);
        if (suite.children) addToMap(suite.children);
      });
    };
    addToMap(suites);
    return suiteMap;
  }

  /**
   * Constructs the test group name using a hierarchical format.
   */
  private buildTestGroupName(
    suiteId: number,
    suiteMap: Map<number, any>,
    isHierarchyGroupName: boolean
  ): string {
    if (!isHierarchyGroupName) return suiteMap.get(suiteId)?.name || '';

    let currentSuite = suiteMap.get(suiteId);
    let path = currentSuite?.name || '';

    while (currentSuite?.parentSuite) {
      const parentSuite = suiteMap.get(currentSuite.parentSuite.id);
      if (!parentSuite) break;
      path = `${parentSuite.name}/${path}`;
      currentSuite = parentSuite;
    }

    const parts = path.split('/');
    if (parts.length - 1 === 1) return parts[1];
    return parts.length > 3
      ? `${parts[1]}/.../${parts[parts.length - 1]}`
      : `${parts[1]}/${parts[parts.length - 1]}`;
  }

  /**
   * Builds the full suite path from the root to the suite.
   */
  private buildSuitePath(suiteId: number, suiteMap: Map<number, any>): string {
    const parts: string[] = [];
    let currentSuite = suiteMap.get(suiteId);

    while (currentSuite) {
      const name = String(currentSuite?.name || '').trim();
      if (name) parts.unshift(name);
      const parentId = currentSuite?.parentSuite?.id;
      if (!parentId) break;
      currentSuite = suiteMap.get(parentId);
    }

    return parts.join('/');
  }

  private async fetchCrossTestPoints(projectName: string, testCaseIds: any[]): Promise<any[]> {
    try {
      const url = `${this.orgUrl}${projectName}/_apis/test/points?api-version=6.0`;
      if (testCaseIds.length === 0) {
        return [];
      }

      const requestBody: any = {
        PointsFilter: {
          TestcaseIds: testCaseIds,
        },
      };

      const { data: value } = await TFSServices.postRequest(url, this.token, 'Post', requestBody, null);
      if (!value || !value.points || !Array.isArray(value.points)) {
        logger.warn('No test points found or invalid response format');
        return [];
      }

      // Group test points by test case ID
      const pointsByTestCase = new Map();
      value.points.forEach((point: any) => {
        const testCaseId = point.testCase.id;

        if (!pointsByTestCase.has(testCaseId)) {
          pointsByTestCase.set(testCaseId, []);
        }

        pointsByTestCase.get(testCaseId).push(point);
      });

      // For each test case, find the point with the most recent run
      const latestPoints: any[] = [];

      for (const [testCaseId, points] of pointsByTestCase.entries()) {
        // Sort by lastTestRun.id (descending), then by lastResult.id (descending)
        const sortedPoints = points.sort((a: any, b: any) => {
          // Parse IDs as numbers for proper comparison
          const aRunId = parseInt(a.lastTestRun?.id || '0');
          const bRunId = parseInt(b.lastTestRun?.id || '0');

          if (aRunId !== bRunId) {
            return bRunId - aRunId; // Sort by run ID first (descending)
          }

          const aResultId = parseInt(a.lastResult?.id || '0');
          const bResultId = parseInt(b.lastResult?.id || '0');
          return bResultId - aResultId; // Then by result ID (descending)
        });

        // Take the first item (most recent)
        latestPoints.push(sortedPoints[0]);
      }

      // Fetch detailed information for each test point and map to required format
      const detailedPoints = await Promise.all(
        latestPoints.map(async (point: any) => {
          const url = `${point.url}?witFields=Microsoft.VSTS.TCM.Steps&includePointDetails=true`;
          const detailedPoint = await TFSServices.getItemContent(url, this.token);
          return this.mapTestPointForCrossPlans(detailedPoint, projectName);
          // return this.mapTestPointForCrossPlans(detailedPoint, projectName);
        })
      );
      return detailedPoints;
    } catch (err: any) {
      logger.error(`Error during fetching Cross Test Points: ${err.message}`);
      logger.error(`Error stack: ${err.stack}`);
      return [];
    }
  }

  /**
   * Fetches test points by suite ID.
   */
  private async fetchTestPoints(
    projectName: string,
    testPlanId: string,
    testSuiteId: string
  ): Promise<any[]> {
    try {
      const url = `${this.orgUrl}${projectName}/_apis/testplan/Plans/${testPlanId}/Suites/${testSuiteId}/TestPoint?includePointDetails=true`;
      const { value: testPoints, count } = await TFSServices.getItemContent(url, this.token);

      return count !== 0 ? testPoints.map((testPoint: any) => this.mapTestPoint(testPoint, projectName)) : [];
    } catch (error: any) {
      logger.error(`Error during fetching Test Points: ${error.message}`);
      return [];
    }
  }

  /**
   * Maps raw test point data to a simplified object.
   */
  private mapTestPoint(testPoint: any, projectName: string): any {
    return {
      testPointId: testPoint.id,
      testCaseId: testPoint.testCaseReference.id,
      testCaseName: testPoint.testCaseReference.name,
      testCaseUrl: `${this.orgUrl}${projectName}/_workitems/edit/${testPoint.testCaseReference.id}`,
      configurationName: testPoint.configuration?.name,
      outcome: testPoint.results?.outcome || 'Not Run',
      testSuite: testPoint.testSuite,
      lastRunId: testPoint.results?.lastTestRunId,
      lastResultId: testPoint.results?.lastResultId,
      lastResultDetails: testPoint.results?.lastResultDetails,
    };
  }

  /**
   * Maps raw test point data to a simplified object.
   */
  private mapTestPointForCrossPlans(testPoint: any, projectName: string): any {
    return {
      testPointId: testPoint.id,
      testCaseId: testPoint.testCase.id,
      testCaseName: testPoint.testCase.name,
      testCaseUrl: `${this.orgUrl}${projectName}/_workitems/edit/${testPoint.testCase.id}`,
      testSuite: testPoint.testSuite,
      configurationName: testPoint.configuration?.name,
      outcome: testPoint.outcome || 'Not Run',
      lastRunId: testPoint.lastTestRun?.id,
      lastResultId: testPoint.lastResult?.id,
      lastResultDetails: testPoint.lastResultDetails || {
        duration: 0,
        dateCompleted: '0000-00-00T00:00:00.000Z',
        runBy: { displayName: 'No tester', id: '00000000-0000-0000-0000-000000000000' },
      },
    };
  }

  // Helper method to get all test points for a test case
  async getTestPointsForTestCases(projectName: string, testCaseId: string[]): Promise<any> {
    const url = `${this.orgUrl}${projectName}/_apis/test/points`;
    const requestBody = {
      PointsFilter: {
        TestcaseIds: testCaseId,
      },
    };

    return await TFSServices.postRequest(url, this.token, 'Post', requestBody, null);
  }

  /**
   * Fetches test cases by suite ID.
   */
  private async fetchTestCasesBySuiteId(
    projectName: string,
    testPlanId: string,
    suiteId: string
  ): Promise<any[]> {
    const url = `${this.orgUrl}${projectName}/_apis/testplan/Plans/${testPlanId}/Suites/${suiteId}/TestCase?witFields=Microsoft.VSTS.TCM.Steps`;

    const { value: testCases } = await TFSServices.getItemContent(url, this.token);

    return testCases;
  }

  /**
   * Fetches result data based on the Work Item Test Reporter.
   *
   * This method retrieves detailed result data for a specific test run and result ID,
   * including related work items, selected fields, and additional processing options.
   *
   * @param projectName - The name of the project containing the test run.
   * @param runId - The unique identifier of the test run.
   * @param resultId - The unique identifier of the test result.
   * @param isTestReporter - (Optional) A flag indicating whether the result data is being fetched for the Test Reporter.
   * @param selectedFields - (Optional) An array of field names to include in the result data.
   * @param isQueryMode - (Optional) A flag indicating whether the result data is being fetched in query mode.
   * @returns A promise that resolves to the fetched result data.
   */
  private async fetchResultDataBasedOnWiBase(
    projectName: string,
    runId: string,
    resultId: string,
    isTestReporter: boolean = false,
    selectedFields?: string[],
    isQueryMode?: boolean,
    point?: any,
    includeAllHistory: boolean = false
  ): Promise<any> {
    try {
      let filteredFields: any = {};
      let relatedRequirements: any[] = [];
      let relatedBugs: any[] = [];
      let relatedCRs: any[] = [];
      if (runId === '0' || resultId === '0') {
        if (!point) {
          logger.warn(`Invalid run result ${runId} or result ${resultId}`);
          return null;
        }
        logger.warn(`Current Test point for Test case ${point.testCaseId} is in Active state`);
        const url = `${this.orgUrl}${projectName}/_apis/wit/workItems/${point.testCaseId}?$expand=all`;
        const testCaseData = await TFSServices.getItemContent(url, this.token);
        const newResultData: PlainTestResult = {
          id: 0,
          outcome: point.outcome,
          revision: testCaseData?.rev || 1,
          testCase: { id: point.testCaseId, name: point.testCaseName },
          state: testCaseData?.fields?.['System.State'] || 'Active',
          priority: testCaseData?.fields?.['Microsoft.VSTS.TCM.Priority'] || 0,
          createdDate: testCaseData?.fields?.['System.CreatedDate'] || '0001-01-01T00:00:00',
          testSuite: point.testSuite,
          failureType: 'None',
        };
        if (isQueryMode) {
          this.appendQueryRelations(point.testCaseId, relatedRequirements, relatedBugs, relatedCRs);
        } else {
          const filteredLinkedFields = selectedFields
            ?.filter((field: string) => field.includes('@linked'))
            ?.map((field: string) => field.split('@')[0]);
          const selectedLinkedFieldSet = new Set(filteredLinkedFields);
          const { relations } = testCaseData;
          if (relations) {
            await this.appendLinkedRelations(
              relations,
              relatedRequirements,
              relatedBugs,
              relatedCRs,
              testCaseData,
              selectedLinkedFieldSet
            );
          }
          selectedLinkedFieldSet.clear();
        }
        const filteredTestCaseFields = selectedFields
          ?.filter((field: string) => field.includes('@testCaseWorkItemField'))
          ?.map((field: string) => field.split('@')[0]);
        const selectedFieldSet = new Set(filteredTestCaseFields);
        // Filter fields based on selected field set
        if (selectedFieldSet.size !== 0) {
          filteredFields = [...selectedFieldSet].reduce((obj: any, key) => {
            obj[key] = testCaseData.fields?.[key]?.displayName ?? testCaseData.fields?.[key] ?? '';
            return obj;
          }, {});
          if (selectedFieldSet.has('System.History')) {
            filteredFields['System.History'] = await this.getWorkItemDiscussionHistoryEntries(
              projectName,
              Number(point.testCaseId),
              includeAllHistory
            );
          }
        }
        selectedFieldSet.clear();
        return {
          ...newResultData,
          stepsResultXml: testCaseData.fields['Microsoft.VSTS.TCM.Steps'] || undefined,
          testCaseRevision: testCaseData.rev,
          filteredFields,
          relatedRequirements,
          relatedBugs,
          relatedCRs,
        };
      }
      const url = `${this.orgUrl}${projectName}/_apis/test/runs/${runId}/results/${resultId}?detailsToInclude=Iterations`;
      const resultData = await TFSServices.getItemContent(url, this.token);

      const attachmentsUrl = `${this.orgUrl}${projectName}/_apis/test/runs/${runId}/results/${resultId}/attachments`;
      const { value: analysisAttachments } = await TFSServices.getItemContent(attachmentsUrl, this.token);

      // Build workItem URL with optional expand parameter
      const expandParam = isTestReporter ? '?$expand=all' : '';
      const wiUrl = `${this.orgUrl}${projectName}/_apis/wit/workItems/${resultData.testCase.id}/revisions/${resultData.testCaseRevision}${expandParam}`;
      const wiByRevision = await TFSServices.getItemContent(wiUrl, this.token);

      // Process selected fields if provided
      if (isTestReporter) {
        // Process related requirements if needed
        if (isQueryMode) {
          this.appendQueryRelations(resultData.testCase.id, relatedRequirements, relatedBugs, relatedCRs);
        } else {
          const filteredLinkedFields = selectedFields
            ?.filter((field: string) => field.includes('@linked'))
            ?.map((field: string) => field.split('@')[0]);
          const selectedLinkedFieldSet = new Set(filteredLinkedFields);
          const { relations } = wiByRevision;
          if (relations) {
            await this.appendLinkedRelations(
              relations,
              relatedRequirements,
              relatedBugs,
              relatedCRs,
              wiByRevision,
              selectedLinkedFieldSet
            );
          }
          selectedLinkedFieldSet.clear();
        }
        const filteredTestCaseFields = selectedFields
          ?.filter((field: string) => field.includes('@testCaseWorkItemField'))
          ?.map((field: string) => field.split('@')[0]);
        const selectedFieldSet = new Set(filteredTestCaseFields);
        // Filter fields based on selected field set
        if (selectedFieldSet.size !== 0) {
          filteredFields = [...selectedFieldSet].reduce((obj: any, key) => {
            obj[key] = wiByRevision.fields?.[key]?.displayName ?? wiByRevision.fields?.[key] ?? '';
            return obj;
          }, {});
          if (selectedFieldSet.has('System.History')) {
            filteredFields['System.History'] = await this.getWorkItemDiscussionHistoryEntries(
              projectName,
              Number(resultData.testCase.id),
              includeAllHistory
            );
          }
        }
        selectedFieldSet.clear();
      }
      return {
        ...resultData,

        stepsResultXml: wiByRevision.fields['Microsoft.VSTS.TCM.Steps'] || undefined,
        analysisAttachments,
        testCaseRevision: resultData.testCaseRevision,
        filteredFields,
        relatedRequirements,
        relatedBugs,
        relatedCRs,
      };
    } catch (error: any) {
      logger.error(`Error while fetching run result: ${error.message}`);
      if (isTestReporter) {
        logger.error(`Error stack: ${error.stack}`);
      }
      return null;
    }
  }

  private async getWorkItemDiscussionHistoryEntries(
    projectName: string,
    workItemId: number,
    includeAllHistory: boolean = false
  ): Promise<any[]> {
    const id = Number(workItemId);
    if (!Number.isFinite(id)) return [];

    const cached = this.workItemDiscussionCache.get(id);
    if (cached !== undefined) {
      if (cached.length === 0 && this.isVerboseHistoryDebugEnabled()) {
        logger.debug(
          `[History] Cache hit but empty for work item ${id} (project=${projectName}, includeAll=${includeAllHistory})`
        );
      }
      return includeAllHistory ? cached : cached.slice(0, 1);
    }

    const fromComments = await this.tryFetchDiscussionFromComments(projectName, id);
    if (fromComments !== null) {
      if (fromComments.length === 0 && this.isVerboseHistoryDebugEnabled()) {
        logger.debug(
          `[History] Comments API returned 0 entries for work item ${id} (project=${projectName}, includeAll=${includeAllHistory})`
        );
      }
      const sorted = this.sortDiscussionEntries(fromComments);
      const normalized = this.normalizeDiscussionEntries(sorted);
      if (normalized.length === 0 && fromComments.length > 0) {
        logger.warn(
          `[History] Comments API returned ${fromComments.length} items but normalized to 0 for work item ${id} ` +
            `(project=${projectName}, includeAll=${includeAllHistory})`
        );
      }
      this.workItemDiscussionCache.set(id, normalized);
      return includeAllHistory ? normalized : normalized.slice(0, 1);
    }

    logger.warn(
      `[History] Comments API fetch failed for work item ${id} (project=${projectName}). Returning empty history.`
    );

    // Comments endpoint is the source-of-truth for discussion history.
    // If it's unavailable / returns nothing, don't fall back to System.History updates
    // (those often contain system-authored noise and empty rows).
    return [];
  }

  private normalizeDiscussionEntries(entries: any[]): any[] {
    const list = Array.isArray(entries) ? entries : [];
    const seen = new Set<string>();
    const out: any[] = [];

    for (const e of list) {
      const createdDate = String(e?.createdDate ?? '').trim();
      const createdBy = String(e?.createdBy ?? '').trim();
      const textRaw = e?.text;
      const text = typeof textRaw === 'string' ? textRaw.trim() : '';

      if (!text) {
        continue;
      }
      if (this.isSystemIdentity(createdBy)) {
        continue;
      }

      const textForKey = this.stripHtmlForEmptiness(text);
      if (!textForKey) {
        continue;
      }

      const key = `${createdDate}|${createdBy}|${textForKey}`;
      if (seen.has(key)) {
        continue;
      }
      seen.add(key);

      out.push({ createdDate, createdBy, text });
    }

    return out;
  }

  private isVerboseHistoryDebugEnabled(): boolean {
    return (
      String(process?.env?.DOCGEN_VERBOSE_HISTORY_DEBUG ?? '').toLowerCase() === 'true' ||
      String(process?.env?.DOCGEN_VERBOSE_HISTORY_DEBUG ?? '') === '1'
    );
  }

  private isRunResultDebugEnabled(): boolean {
    return (
      String(process?.env?.DOCGEN_DEBUG_RUNRESULT ?? '').toLowerCase() === 'true' ||
      String(process?.env?.DOCGEN_DEBUG_RUNRESULT ?? '') === '1'
    );
  }

  private extractCommentText(comment: AdoWorkItemComment): string {
    const rendered = comment?.renderedText;
    // In Azure DevOps the `renderedText` field can be present but empty ("") even when `text` is populated.
    // Prefer `renderedText` only when it is a non-empty string.
    if (typeof rendered === 'string' && rendered.trim() !== '') return rendered;

    const text = comment?.text;
    if (typeof text === 'string') return text;

    // Defensive fallbacks for server variants that may wrap text differently
    const anyComment: any = comment as any;
    const candidates = [
      anyComment?.text?.value,
      anyComment?.text?.content,
      anyComment?.renderedText?.value,
      anyComment?.renderedText?.content,
    ];
    for (const c of candidates) {
      if (typeof c === 'string') return c;
    }
    return '';
  }

  private stringifyForDebug(value: any, maxChars: number): string {
    try {
      const s = JSON.stringify(value);
      if (typeof maxChars === 'number' && maxChars > 0 && s.length > maxChars) {
        return `${s.slice(0, maxChars)}...[truncated ${s.length - maxChars} chars]`;
      }
      return s;
    } catch (e) {
      return String(value);
    }
  }

  private stripHtmlForEmptiness(html: string): string {
    return String(html ?? '')
      .replace(/<[^>]*>/g, ' ')
      .replace(/&nbsp;/gi, ' ')
      .replace(/&amp;/gi, '&')
      .replace(/\s+/g, ' ')
      .trim();
  }

  private isSystemIdentity(displayName: string): boolean {
    const s = String(displayName ?? '')
      .toLowerCase()
      .trim();
    if (!s) return false;
    return s === 'microsoft.teamfoundation.system' || s.includes('microsoft.teamfoundation.system');
  }

  private getCommentAuthor(comment: AdoWorkItemComment): string {
    // Azure DevOps / TFS may record comments as created by the System identity,
    // but include the real author under createdOnBehalfOf.
    const onBehalf = comment?.createdOnBehalfOf?.displayName ?? comment?.createdOnBehalfOf?.uniqueName ?? '';
    if (onBehalf && !this.isSystemIdentity(onBehalf)) return String(onBehalf);

    return String(comment?.createdBy?.displayName ?? comment?.createdBy?.uniqueName ?? '');
  }

  private sortDiscussionEntries(entries: any[]): any[] {
    const list = Array.isArray(entries) ? entries.slice() : [];
    list.sort((a, b) => {
      const ta = new Date(a?.createdDate ?? 0).getTime();
      const tb = new Date(b?.createdDate ?? 0).getTime();
      if (tb !== ta) return tb - ta;
      const ia = Number.isFinite(a?._idx) ? a._idx : 0;
      const ib = Number.isFinite(b?._idx) ? b._idx : 0;
      return ia - ib;
    });
    return list.map(({ _idx, ...rest }) => rest);
  }

  private getContinuationToken(headers: any, response?: AdoWorkItemCommentsResponse): string | undefined {
    const candidates = ['x-ms-continuationtoken', 'x-ms-continuation-token'];

    if (headers) {
      // If headers is a fetch/Headers-like object
      if (typeof headers.get === 'function') {
        for (const key of candidates) {
          const val = headers.get(key) ?? headers.get(key.toLowerCase()) ?? headers.get(key.toUpperCase());
          if (typeof val === 'string' && val.trim() !== '') return val;
        }
      }

      // If headers is a plain object (axios-style)
      if (typeof headers === 'object') {
        for (const key of candidates) {
          const direct =
            (headers as any)[key] ??
            (headers as any)[key.toLowerCase()] ??
            (headers as any)[key.toUpperCase()];
          if (typeof direct === 'string' && direct.trim() !== '') return direct;
        }

        for (const [k, v] of Object.entries(headers)) {
          const lk = String(k).toLowerCase();
          if (!candidates.includes(lk)) continue;
          if (typeof v === 'string' && v.trim() !== '') return v;
          // Some libs return arrays for repeated headers; take the first string.
          if (Array.isArray(v) && typeof v[0] === 'string' && v[0].trim() !== '') return v[0];
        }
      }
    }

    const fromBody = response?.continuationToken;
    if (typeof fromBody === 'string' && fromBody.trim() !== '') return fromBody;
    return undefined;
  }

  private async tryFetchDiscussionFromComments(
    projectName: string,
    workItemId: number
  ): Promise<any[] | null> {
    try {
      // NOTE: Some Azure DevOps Server / collection configs omit comment text unless includeText=true.
      const baseUrl = `${this.orgUrl}${projectName}/_apis/wit/workItems/${workItemId}/comments?order=desc&$top=200&includeText=true&api-version=7.1-preview.3`;
      const all: AdoWorkItemComment[] = [];
      let continuationToken: string | undefined = undefined;
      let page = 0;
      const MAX_PAGES = 50;
      const seenTokens = new Set<string>();
      const verbose = this.isVerboseHistoryDebugEnabled();
      const pageResponsesForDebug: { page: number; data: any; headers: any }[] = [];
      let firstPageDebug: { url: string; data: any; headers: any } | null = null;

      do {
        if (page >= MAX_PAGES) {
          logger.warn(
            `[History][comments] Reached max pages (${MAX_PAGES}) for work item ${workItemId}. ` +
              `Stopping pagination to avoid infinite loop (lastToken=${String(continuationToken ?? '')}).`
          );
          break;
        }

        let url = baseUrl;
        if (continuationToken) {
          url += `&continuationToken=${encodeURIComponent(continuationToken)}`;
        }

        const { data, headers } = await this.limit(() =>
          TFSServices.getItemContentWithHeaders(url, this.token, 'get', {}, {}, false)
        );

        if (verbose && !firstPageDebug) {
          firstPageDebug = { url, data, headers };
        }
        const response = (data ?? {}) as AdoWorkItemCommentsResponse;
        const comments = Array.isArray(response.comments) ? response.comments : [];
        all.push(...comments);

        if (verbose && pageResponsesForDebug.length < 2) {
          pageResponsesForDebug.push({ page: page + 1, data, headers });
        }

        const nextToken = this.getContinuationToken(headers, response);
        if (nextToken) {
          if (seenTokens.has(nextToken)) {
            logger.warn(
              `[History][comments] Continuation token repeated for work item ${workItemId}. ` +
                `Stopping pagination to avoid infinite loop (token=${nextToken}).`
            );
            continuationToken = undefined;
          } else {
            seenTokens.add(nextToken);
            continuationToken = nextToken;
          }
        } else {
          continuationToken = undefined;
        }

        page++;
      } while (continuationToken);

      if (all.length === 0) {
        if (verbose && firstPageDebug) {
          logger.debug(
            `[History][comments] 0 comments returned for work item ${workItemId}. url=${firstPageDebug.url}`
          );
          logger.debug(
            `[History][comments] Raw comments API response (truncated) for work item ${workItemId} (page=1): ` +
              this.stringifyForDebug(firstPageDebug.data, 5000)
          );
        }
        return [];
      }

      let deletedCount = 0;
      let emptyRawCount = 0;
      let emptyAfterStripCount = 0;
      let systemIdentityCount = 0;
      const emptyRawIds: number[] = [];

      const entries = all
        .filter((c) => {
          const deleted = !!c?.isDeleted;
          if (deleted) deletedCount++;
          return !deleted;
        })
        .map((c, idx) => {
          const raw = this.extractCommentText(c);
          const text = typeof raw === 'string' ? raw.trim() : '';
          const createdBy = this.getCommentAuthor(c);
          const createdDate = c?.createdDate ?? '';

          if (!text) {
            emptyRawCount++;
            if (emptyRawIds.length < 10 && Number.isFinite((c as any)?.id))
              emptyRawIds.push(Number((c as any)?.id));
            return null;
          }

          if (!this.stripHtmlForEmptiness(text)) {
            emptyAfterStripCount++;
            return null;
          }

          if (this.isSystemIdentity(createdBy)) {
            systemIdentityCount++;
            return null;
          }

          return { _idx: idx, createdDate, createdBy, text };
        })
        .filter(Boolean) as any[];

      if (entries.length === 0 && all.length > 0) {
        if (verbose) {
          logger.debug(
            `[History][comments] Work item ${workItemId} returned ${all.length} comments but 0 usable entries ` +
              `(deleted=${deletedCount}, emptyRaw=${emptyRawCount}, emptyAfterStrip=${emptyAfterStripCount}, system=${systemIdentityCount}, pages=${page}, emptyRawIds=${emptyRawIds.join(
                ','
              )})`
          );
          for (const p of pageResponsesForDebug) {
            logger.debug(
              `[History][comments] Raw comments API response (truncated) for work item ${workItemId} (page=${p.page}): ` +
                this.stringifyForDebug(p.data, 5000)
            );
          }
        }
      }

      return entries;
    } catch (e) {
      logger.debug(
        `[History][comments] Failed fetching comments for work item ${workItemId}: ${
          (e as any)?.message ?? e
        }`
      );
      return null;
    }
  }

  private async debugProbeHistoryFromUpdates(projectName: string, workItemId: number): Promise<void> {
    if (!this.isVerboseHistoryDebugEnabled()) return;
    try {
      const url = `${this.orgUrl}${projectName}/_apis/wit/workItems/${workItemId}/updates?$top=200&api-version=7.1-preview.3`;
      const { data, headers } = await this.limit(() =>
        TFSServices.getItemContentWithHeaders(url, this.token, 'get', {}, {}, false)
      );

      const updates = Array.isArray((data as any)?.value) ? (data as any).value : [];
      let historyUpdates = 0;
      const samples: string[] = [];
      const fieldKeySet = new Set<string>();
      const updateSamplesForDebug: any[] = [];

      for (const u of updates) {
        if (updateSamplesForDebug.length < 2) updateSamplesForDebug.push(u);
        const fields = (u as any)?.fields;
        if (fields && typeof fields === 'object') {
          for (const k of Object.keys(fields)) fieldKeySet.add(String(k));
        }
        const h = fields?.['System.History'];
        const candidate =
          (typeof h?.newValue === 'string' && h.newValue) ||
          (typeof h?.oldValue === 'string' && h.oldValue) ||
          (typeof h === 'string' && h) ||
          '';
        const normalized = this.stripHtmlForEmptiness(String(candidate));
        if (!normalized) continue;
        historyUpdates++;
        if (samples.length < 3) {
          const oneLine = normalized.replace(/\s+/g, ' ').trim();
          samples.push(oneLine.length > 300 ? `${oneLine.slice(0, 300)}...[truncated]` : oneLine);
        }
      }

      logger.debug(
        `[History][updates] Probe for work item ${workItemId}: updates=${updates.length}, ` +
          `historyUpdates=${historyUpdates}, url=${url}`
      );
      if (samples.length > 0) {
        logger.debug(
          `[History][updates] Sample System.History entries for work item ${workItemId}: ${this.stringifyForDebug(
            samples,
            2000
          )}`
        );
      }

      if (historyUpdates === 0 && updates.length > 0) {
        const fieldKeys = [...fieldKeySet].slice(0, 80);
        logger.debug(
          `[History][updates] No System.History found for work item ${workItemId}. ` +
            `Field keys seen (first ${fieldKeys.length}): ${this.stringifyForDebug(fieldKeys, 4000)}`
        );
      }

      const continuation = this.getContinuationToken(headers, undefined);
      if (continuation) {
        logger.debug(
          `[History][updates] Updates endpoint returned continuation token for work item ${workItemId}: ${continuation}`
        );
      }
    } catch (e) {
      logger.debug(
        `[History][updates] Probe failed for work item ${workItemId}: ${(e as any)?.message ?? e}`
      );
    }
  }

  private async debugProbeHistoryFromRevisions(projectName: string, workItemId: number): Promise<void> {
    if (!this.isVerboseHistoryDebugEnabled()) return;
    try {
      const url = `${this.orgUrl}${projectName}/_apis/wit/workItems/${workItemId}/revisions?$top=200&api-version=7.1-preview.3`;
      const { data } = await this.limit(() =>
        TFSServices.getItemContentWithHeaders(url, this.token, 'get', {}, {}, false)
      );

      const revisions = Array.isArray((data as any)?.value) ? (data as any).value : [];
      let historyRevs = 0;
      const samples: string[] = [];

      for (const r of revisions) {
        const fields = (r as any)?.fields;
        const h = fields?.['System.History'];
        const candidate = typeof h === 'string' ? h : '';
        const normalized = this.stripHtmlForEmptiness(String(candidate));
        if (!normalized) continue;
        historyRevs++;
        if (samples.length < 3) {
          const oneLine = normalized.replace(/\s+/g, ' ').trim();
          samples.push(oneLine.length > 300 ? `${oneLine.slice(0, 300)}...[truncated]` : oneLine);
        }
      }

      logger.debug(
        `[History][revisions] Probe for work item ${workItemId}: revisions=${revisions.length}, historyRevisions=${historyRevs}, url=${url}`
      );
      if (samples.length > 0) {
        logger.debug(
          `[History][revisions] Sample System.History entries for work item ${workItemId}: ${this.stringifyForDebug(
            samples,
            2000
          )}`
        );
      }
    } catch (e) {
      logger.debug(
        `[History][revisions] Probe failed for work item ${workItemId}: ${(e as any)?.message ?? e}`
      );
    }
  }

  private isNotRunStep = (result: any): boolean => {
    return result && result.stepStatus === 'Not Run';
  };

  private appendQueryRelations(
    testCaseId: any,
    relatedRequirements: any[],
    relatedBugs: any[],
    relatedCRs: any[]
  ) {
    if (this.testToAssociatedItemMap.size !== 0) {
      const relatedItemSet = this.testToAssociatedItemMap.get(Number(testCaseId));
      if (relatedItemSet) {
        for (const relatedItem of relatedItemSet) {
          const { id, fields, _links } = relatedItem;
          const itemTitle = fields['System.Title'];
          const itemUrl = _links.html.href;
          const workItemType = fields['System.WorkItemType'];
          delete fields['System.Title'];
          delete fields['System.WorkItemType'];

          const customFields = this.standardCustomField(fields, this.querySelectedColumns);
          let objectToSave = {
            id,
            title: itemTitle,
            url: itemUrl,
            workItemType,
            ...customFields,
          };

          switch (workItemType) {
            case 'Requirement':
              relatedRequirements.push(objectToSave);
              break;
            case 'Bug':
              relatedBugs.push(objectToSave);
              break;
            case 'Change Request':
              relatedCRs.push(objectToSave);
              break;
          }
        }
      }
    }
  }

  private async appendLinkedRelations(
    relations: any,
    relatedRequirements: any[],
    relatedBugs: any[],
    relatedCRs: any[],
    wiByRevision: any,
    selectedLinkedFieldSet: Set<string>
  ) {
    for (const relation of relations) {
      if (
        relation.rel?.includes('System.LinkTypes') ||
        relation.rel?.includes('Microsoft.VSTS.Common.TestedBy')
      ) {
        const relatedUrl = relation.url;
        try {
          const wi = await TFSServices.getItemContent(relatedUrl, this.token);
          if (wi.fields['System.State'] === 'Closed') {
            continue;
          }
          if (
            selectedLinkedFieldSet.has('associatedRequirement') &&
            wi.fields['System.WorkItemType'] === 'Requirement'
          ) {
            const { id, fields, _links } = wi;
            const title = fields['System.Title'];
            const customerFieldKey = Object.keys(fields).find((key) =>
              key.toLowerCase().includes('customer')
            );
            const customerId = customerFieldKey ? fields[customerFieldKey] : undefined;
            const url = _links.html.href;
            relatedRequirements.push({
              id,
              title,
              workItemType: 'Requirement',
              customerId,
              url,
            });
          } else if (
            selectedLinkedFieldSet.has('associatedBug') &&
            wi.fields['System.WorkItemType'] === 'Bug'
          ) {
            const { id, fields, _links } = wi;
            const title = fields['System.Title'];
            const url = _links.html.href;
            relatedBugs.push({ id, title, workItemType: 'Bug', url });
          } else if (
            selectedLinkedFieldSet.has('associatedCR') &&
            wi.fields['System.WorkItemType'] === 'Change Request'
          ) {
            const { id, fields, _links } = wi;
            const title = fields['System.Title'];
            const url = _links.html.href;
            relatedCRs.push({ id, title, workItemType: 'Change Request', url });
          }
        } catch (err: any) {
          logger.error(`Could not append related work item to test case ${wiByRevision.id}: ${err.message}`);
        }
      }
    }
  }

  /**
   * Fetches result data based on the specified work item (WI) details.
   *
   * @param projectName - The name of the project associated with the work item.
   * @param runId - The unique identifier of the test run.
   * @param resultId - The unique identifier of the test result.
   * @returns A promise that resolves to the result data.
   */
  private async fetchResultDataBasedOnWi(projectName: string, runId: string, resultId: string): Promise<any> {
    return this.fetchResultDataBasedOnWiBase(projectName, runId, resultId);
  }

  /**
   * Converts a run status string into a human-readable format.
   *
   * @param status - The status string to convert. Expected values are:
   *   - `'passed'`: Indicates the run was successful.
   *   - `'failed'`: Indicates the run was unsuccessful.
   *   - `'notApplicable'`: Indicates the run is not applicable.
   *   - Any other value will default to `'Not Run'`.
   * @returns A human-readable string representing the run status.
   */
  private convertRunStatus(status: string): string {
    switch (status.toLowerCase()) {
      case 'passed':
        return 'Passed';
      case 'failed':
        return 'Failed';
      case 'notapplicable':
        return 'Not Applicable';
      default:
        return 'Not Run';
    }
  }

  /**
   * Converts the outcome of an action result based on specific conditions.
   *
   * - If the outcome is 'Unspecified' and the action result is a shared step title,
   *   it returns an empty string.
   * - If the outcome is 'Unspecified' but not a shared step title, it returns 'Not Run'.
   * - If the outcome is not 'Not Run', it returns the original outcome.
   * - Otherwise, it returns an empty string.
   *
   * @param actionResult - The action result object containing the outcome and other properties.
   * @returns A string representing the converted outcome.
   */
  private convertUnspecifiedRunStatus(actionResult: any) {
    if (!actionResult || (actionResult.outcome === 'Unspecified' && actionResult.isSharedStepTitle)) {
      return '';
    }

    return actionResult.outcome === 'Unspecified'
      ? 'Not Run'
      : actionResult.outcome !== 'Not Run'
      ? actionResult.outcome
      : '';
  }

  /**
   * Aligns test steps with iterations and generates detailed results based on the provided options.
   *
   * @param testData - An array of test data objects containing test points and test cases.
   * @param iterations - An array of iteration data used to map test cases to their respective iterations.
   * @param includeNotRunTestCases - including not run test cases in the results.
   * @param options - Configuration options for processing the test data.
   * @param options.selectedFields - An optional array of selected fields to filter step-level properties.
   * @param options.createResultObject - A callback function to create a result object for each test or step.
   * @param options.shouldProcessStepLevel - A callback function to determine whether to process at the step level.
   * @returns An array of detailed result objects, either at the test level or step level, depending on the options.
   */
  private alignStepsWithIterationsBase(
    testData: any[],
    iterations: any[],
    includeNotRunTestCases: boolean,
    includeItemsWithNoIterations: boolean,
    isTestReporter: boolean,
    options: {
      selectedFields?: any[];
      createResultObject: (params: {
        testItem: any;
        point: any;
        fetchedTestCase: any;
        actionResult?: any;
        filteredFields?: Set<string>;
      }) => any;
      shouldProcessStepLevel: (fetchedTestCase: any, filteredFields: Set<string>) => boolean;
    }
  ): any[] {
    const detailedResults: any[] = [];

    if (!iterations || iterations?.length === 0) {
      return detailedResults;
    }

    // Process filtered fields if available
    const filteredFields = new Set(
      options.selectedFields
        ?.filter((field: string) => field.includes('@stepsRunProperties'))
        ?.map((field: string) => field.split('@')[0]) || []
    );
    const iterationsMap = this.createIterationsMap(iterations, isTestReporter, includeNotRunTestCases);

    for (const testItem of testData) {
      const testCasesItems = Array.isArray(testItem?.testCasesItems) ? testItem.testCasesItems : [];
      const testCaseById = new Map<number, any>();
      for (const tc of testCasesItems) {
        const id = Number(tc?.workItem?.id);
        if (Number.isFinite(id)) {
          testCaseById.set(id, tc);
        }
      }
      for (const point of testItem.testPointsItems) {
        const testCase = testCaseById.get(Number(point.testCaseId));
        if (!testCase) continue;

        if (testCase.workItem.workItemFields.length === 0) {
          logger.warn(`Could not fetch the steps from WI ${JSON.stringify(testCase.workItem.id)}`);
          if (!isTestReporter) {
            continue;
          }
        }
        const iterationKey =
          !point.lastRunId || !point.lastResultId
            ? `${testCase.workItem.id}`
            : `${point.lastRunId}-${point.lastResultId}-${testCase.workItem.id}`;
        const fetchedTestCase =
          iterationsMap[iterationKey] || (includeNotRunTestCases ? testCase : undefined);
        // First check if fetchedTestCase exists
        if (!fetchedTestCase) continue;

        // Then separately check for iteration only if configured not to include items without iterations
        if (!includeItemsWithNoIterations && !fetchedTestCase.iteration) continue;

        // Determine if we should process at step level
        this.AppendResults(options, fetchedTestCase, filteredFields, testItem, point, detailedResults);
      }
    }
    return detailedResults;
  }

  /**
   * Appends results to the detailed results array based on the provided options and test data.
   *
   * @param options - Configuration options for processing the test data.
   * @param fetchedTestCase - The fetched test case object containing iteration and action results.
   * @param filteredFields - A set of filtered fields to include in the result object.
   * @param testItem - The test item object containing test point information.
   * @param point - The test point object containing details about the test case.
   * @param detailedResults - The array to which the result objects will be appended.
   */
  private AppendResults(
    options: {
      selectedFields?: any[];
      createResultObject: (params: {
        testItem: any;
        point: any;
        fetchedTestCase: any;
        actionResult?: any;
        filteredFields?: Set<string>;
      }) => any;
      shouldProcessStepLevel: (fetchedTestCase: any, filteredFields: Set<string>) => boolean;
    },
    fetchedTestCase: any,
    filteredFields: Set<string>,
    testItem: any,
    point: any,
    detailedResults: any[]
  ) {
    const shouldProcessSteps = options.shouldProcessStepLevel(fetchedTestCase, filteredFields);

    if (!shouldProcessSteps) {
      // Create a test-level result object
      const resultObj = options.createResultObject({
        testItem,
        point,
        fetchedTestCase,
        filteredFields,
      });
      detailedResults.push(resultObj);
    } else {
      // Create step-level result objects
      const { actionResults } = fetchedTestCase?.iteration;
      const resultObjectsToAdd: any[] = [];
      for (let i = 0; i < actionResults?.length; i++) {
        const resultObj = options.createResultObject({
          testItem,
          point,
          fetchedTestCase,
          actionResult: actionResults[i],
          filteredFields,
        });
        resultObjectsToAdd.push(resultObj);
      }
      resultObjectsToAdd.length > 0
        ? detailedResults.push(...resultObjectsToAdd)
        : detailedResults.push(
            options.createResultObject({
              testItem,
              point,
              fetchedTestCase,
              filteredFields,
            })
          );
    }
  }

  /**
   * Aligns test steps with their corresponding iterations by processing the provided test data and iterations.
   * This method utilizes a base alignment function with custom logic for processing step-level data and creating result objects.
   *
   * @param testData - An array of test data objects containing information about test cases and their steps.
   * @param iterations - An array of iteration objects containing action results for each test case.
   * @returns An array of aligned step data objects, each containing detailed information about a test step and its status.
   *
   * The alignment process includes:
   * - Filtering and processing step-level data based on the presence of iteration and action results.
   * - Creating result objects for each step, including details such as test ID, step identifier, action path, step status, and comments.
   */
  private alignStepsWithIterations(testData: any[], iterations: any[]): any[] {
    return this.alignStepsWithIterationsBase(testData, iterations, false, false, false, {
      shouldProcessStepLevel: (fetchedTestCase) =>
        fetchedTestCase != null &&
        fetchedTestCase.iteration != null &&
        fetchedTestCase.iteration.actionResults != null,
      createResultObject: ({ point, fetchedTestCase, actionResult }) => {
        if (!actionResult) return null;

        const stepIdentifier = parseInt(actionResult.stepIdentifier, 10);
        return {
          testId: point.testCaseId,
          testCaseRevision: fetchedTestCase.testCaseRevision,
          testName: point.testCaseName,
          stepIdentifier: stepIdentifier,
          actionPath: actionResult.actionPath,
          stepNo: actionResult.stepPosition,
          stepAction: actionResult.action,
          stepExpected: actionResult.expected,
          isSharedStepTitle: actionResult.isSharedStepTitle,
          stepStatus: this.convertUnspecifiedRunStatus(actionResult),
          stepComments: actionResult.errorMessage || '',
        };
      },
    });
  }
  /**
   * Creates a mapping of iterations by their unique keys.
   */
  private createIterationsMap(
    iterations: any[],
    isTestReporter: boolean,
    includeNotRunTestCases: boolean
  ): Record<string, any> {
    return iterations.reduce((map, iterationItem) => {
      if (
        (isTestReporter && iterationItem.lastRunId && iterationItem.lastResultId) ||
        iterationItem.iteration
      ) {
        const key = `${iterationItem.lastRunId}-${iterationItem.lastResultId}-${iterationItem.testCaseId}`;
        map[key] = iterationItem;
      } else if (includeNotRunTestCases) {
        const key = `${iterationItem.testCaseId}`;
        map[key] = iterationItem;
      }
      return map;
    }, {} as Record<string, any>);
  }

  /**
   * Fetches test data for all suites, including test points and test cases.
   */
  private async fetchTestData(
    suites: any[],
    projectName: string,
    testPlanId: string,
    fetchCrossPlans: boolean = false
  ): Promise<any[]> {
    return await Promise.all(
      suites.map((suite) =>
        this.limit(async () => {
          try {
            const testCasesItems = await this.fetchTestCasesBySuiteId(
              projectName,
              testPlanId,
              suite.testSuiteId
            );
            const testCaseIds = testCasesItems.map((testCase: any) => testCase.workItem.id);
            const testPointsItems = !fetchCrossPlans
              ? await this.fetchTestPoints(projectName, testPlanId, suite.testSuiteId)
              : await this.fetchCrossTestPoints(projectName, testCaseIds);

            return { ...suite, testPointsItems, testCasesItems };
          } catch (error: any) {
            logger.error(`Error occurred for suite ${suite.testSuiteId}: ${error.message}`);
            return suite;
          }
        })
      )
    );
  }

  /**
   * Fetches all result data based on the provided test data, project name, and fetch strategy.
   * This method processes the test data, filters valid test points, and sequentially fetches
   * result data for each valid point using the provided fetch strategy.
   *
   * @param testData - An array of test data objects containing test suite and test points information.
   * @param projectName - The name of the project for which the result data is being fetched.
   * @param fetchStrategy - A function that defines the strategy for fetching result data. It takes
   *                        the project name, test suite ID, a test point, and additional arguments,
   *                        and returns a Promise resolving to the fetched result data.
   * @param additionalArgs - An optional array of additional arguments to be passed to the fetch strategy.
   * @returns A Promise that resolves to an array of fetched result data.
   */
  private async fetchAllResultDataBase(
    testData: any[],
    projectName: string,
    isTestReporter: boolean,
    fetchStrategy: (projectName: string, testSuiteId: string, point: any, ...args: any[]) => Promise<any>,
    additionalArgs: any[] = []
  ): Promise<any[]> {
    const results = [];

    for (const item of testData) {
      if (item.testPointsItems && item.testPointsItems.length > 0) {
        const { testSuiteId, testPointsItems } = item;

        // Filter and sort valid points
        const validPoints = isTestReporter
          ? testPointsItems
          : testPointsItems.filter((point: any) => point && point.lastRunId && point.lastResultId);

        // Fetch results for each point.
        // For non-test-reporter paths we keep sequential fetching to avoid context-related errors.
        if (!isTestReporter) {
          for (const point of validPoints) {
            const resultData = await fetchStrategy(projectName, testSuiteId, point, ...additionalArgs);
            if (resultData !== null) {
              results.push(resultData);
            }
          }
        } else {
          // Test Reporter: bounded concurrency + batching to avoid creating a massive promise graph.
          const POINT_BATCH_SIZE = 50;
          const pointLimit = pLimit(6);

          for (let i = 0; i < validPoints.length; i += POINT_BATCH_SIZE) {
            const batch = validPoints.slice(i, i + POINT_BATCH_SIZE);
            const batchResults = await Promise.all(
              batch.map((point: any) =>
                pointLimit(async () => {
                  try {
                    return await fetchStrategy(projectName, testSuiteId, point, ...additionalArgs);
                  } catch (e) {
                    logger.error(
                      `Error occurred for point ${point?.testCaseId ?? 'unknown'}: ${
                        (e as any)?.message ?? e
                      }`
                    );
                    return null;
                  }
                })
              )
            );

            for (const resultData of batchResults) {
              if (resultData !== null) {
                results.push(resultData);
              }
            }
          }
        }
      }
    }

    return results;
  }

  /**
   * Fetches result data for all test points within the given test data, sequentially to avoid context-related errors.
   */
  private async fetchAllResultData(testData: any[], projectName: string): Promise<any[]> {
    return this.fetchAllResultDataBase(testData, projectName, false, (projectName, testSuiteId, point) =>
      this.fetchResultData(projectName, testSuiteId, point)
    );
  }

  //Sorting step positions
  private compareActionResults = (a: string, b: string) => {
    const aParts = a.split('.').map(Number);
    const bParts = b.split('.').map(Number);
    const maxLength = Math.max(aParts.length, bParts.length);

    for (let i = 0; i < maxLength; i++) {
      const aNum = aParts[i] || 0; // Default to 0 if undefined
      const bNum = bParts[i] || 0;

      if (aNum > bNum) return 1;
      if (aNum < bNum) return -1;
      // If equal, continue to next segment
    }

    return 0; // Versions are equal
  };

  /**
   * Fetches result Data for a specific test point
   */
  private async fetchResultDataBase(
    projectName: string,
    testSuiteId: string,
    point: any,
    fetchResultMethod: (project: string, runId: string, resultId: string, ...args: any[]) => Promise<any>,
    createResponseObject: (resultData: any, testSuiteId: string, point: any, ...args: any[]) => any,
    additionalArgs: any[] = []
  ): Promise<any> {
    try {
      const { lastRunId, lastResultId } = point;

      const resultData = await fetchResultMethod(
        projectName,
        lastRunId?.toString() || '0',
        lastResultId?.toString() || '0',
        ...additionalArgs
      );

      let iteration =
        resultData.iterationDetails?.length > 0
          ? resultData.iterationDetails[resultData.iterationDetails.length - 1]
          : undefined;

      if (resultData.stepsResultXml && !iteration) {
        iteration = { actionResults: [] };
        if (!Array.isArray(resultData.iterationDetails)) {
          resultData.iterationDetails = [];
        }
        resultData.iterationDetails.push(iteration);
      }

      if (resultData.stepsResultXml && iteration) {
        const actionResults = Array.isArray(iteration.actionResults) ? iteration.actionResults : [];
        const actionResultsWithSharedModels = actionResults.filter(
          (result: any) => result.sharedStepModel
        );

        const sharedStepIdToRevisionLookupMap: Map<number, number> = new Map();

        if (actionResultsWithSharedModels?.length > 0) {
          actionResultsWithSharedModels.forEach((actionResult: any) => {
            const { sharedStepModel } = actionResult;
            sharedStepIdToRevisionLookupMap.set(Number(sharedStepModel.id), Number(sharedStepModel.revision));
          });
        }
        const stepsList = await this.testStepParserHelper.parseTestSteps(
          resultData.stepsResultXml,
          sharedStepIdToRevisionLookupMap
        );

        sharedStepIdToRevisionLookupMap.clear();

        const stepMap = new Map<string, TestSteps>();
        for (const step of stepsList) {
          stepMap.set(step.stepId.toString(), step);
        }

        for (const actionResult of actionResults) {
          const step = stepMap.get(actionResult.stepIdentifier);
          if (step) {
            actionResult.stepPosition = step.stepPosition;
            actionResult.action = step.action;
            actionResult.expected = step.expected;
            actionResult.isSharedStepTitle = step.isSharedStepTitle;
          }
        }

        if (actionResults.length > 0) {
          // Sort mapped action results by logical step position.
          iteration.actionResults = actionResults
            .filter((result: any) => result.stepPosition)
            .sort((a: any, b: any) => this.compareActionResults(a.stepPosition, b.stepPosition));
        } else {
          // Fallback for runs that have no action results: emit test definition steps as Not Run.
          iteration.actionResults = stepsList
            .filter((step: any) => step?.stepPosition)
            .map((step: any) => ({
              stepIdentifier: String(step?.stepId ?? step?.stepPosition ?? ''),
              stepPosition: step.stepPosition,
              action: step.action,
              expected: step.expected,
              isSharedStepTitle: step.isSharedStepTitle,
              outcome: 'Unspecified',
              errorMessage: '',
              actionPath: String(step?.stepPosition ?? ''),
            }))
            .sort((a: any, b: any) => this.compareActionResults(a.stepPosition, b.stepPosition));
        }
      }

      return resultData?.testCase
        ? createResponseObject(resultData, testSuiteId, point, ...additionalArgs)
        : null;
    } catch (error: any) {
      logger.error(`Error occurred for point ${point.testCaseId}: ${error.message}`);
      logger.error(`Stack trace: ${error.stack}`);
      return null;
    }
  }

  /**
   * Fetches result Data for a specific test point
   */
  private async fetchResultData(projectName: string, testSuiteId: string, point: any) {
    return this.fetchResultDataBase(
      projectName,
      testSuiteId,
      point,
      (project, runId, resultId) => this.fetchResultDataBasedOnWi(project, runId, resultId),
      (resultData, testSuiteId, point) => ({
        testCaseName: `${resultData?.testCase?.name ?? ''} - ${resultData?.testCase?.id ?? ''}`,
        testCaseId: resultData?.testCase?.id,
        testSuiteName: `${resultData?.testSuite?.name ?? ''}`,
        testSuiteId,
        lastRunId: point.lastRunId,
        lastResultId: point.lastResultId,
        iteration:
          resultData.iterationDetails?.length > 0
            ? resultData.iterationDetails[resultData.iterationDetails.length - 1]
            : undefined,
        testCaseRevision: resultData.testCaseRevision,
        failureType: resultData.failureType,
        resolution: resultData.resolutionState,
        comment: resultData.comment,
        analysisAttachments: resultData.analysisAttachments,
      })
    );
  }

  /**
   * Fetches all the linked work items (WI) for the given test case.
   * @param project Project name
   * @param testItems Test cases
   * @returns Array of linked Work Items
   */
  private async fetchLinkedWi(project: string, testItems: any[]): Promise<any[]> {
    logger.info('Fetching linked work items for test cases');
    const CHUNK_SIZE = 100;
    const summarizedItemMap: Map<number, any> = new Map();

    // Prepare map and ID array
    const testIds = testItems.map((item) => {
      summarizedItemMap.set(item.testId, { ...item, linkItems: [] });
      return item.testId.toString();
    });

    logger.info(`Fetching linked work items for ${testIds.length} test cases`);
    // Split test IDs into chunks
    const chunks: string[][] = [];
    for (let i = 0; i < testIds.length; i += CHUNK_SIZE) {
      chunks.push(testIds.slice(i, i + CHUNK_SIZE));
    }
    logger.info(`Fetching linked work items in ${chunks.length} chunks`);

    // Process each chunk
    for (const chunk of chunks) {
      try {
        const url = `${this.orgUrl}${project}/_apis/wit/workItems?ids=${chunk.join(',')}&$expand=relations`;
        const { value: workItems } = await TFSServices.getItemContent(url, this.token);

        for (const wi of workItems || []) {
          const mappedItem = summarizedItemMap.get(wi.id);
          if (!mappedItem || !wi.relations) continue;
          const relatedIds = wi.relations
            .filter((relation: any) => relation?.url?.includes('workItems'))
            .map((rel: any) => rel.url.split('/').pop());

          // Fetch related items in batches
          let allRelatedWi: any[] = [];
          for (let i = 0; i < relatedIds.length; i += CHUNK_SIZE) {
            const relChunk = relatedIds.slice(i, i + CHUNK_SIZE);
            const relatedUrl = `${this.orgUrl}${project}/_apis/wit/workItems?ids=${relChunk.join(
              ','
            )}&$expand=1`;
            const { value: rwi } = await TFSServices.getItemContent(relatedUrl, this.token);
            allRelatedWi = [...allRelatedWi, ...rwi];
          }

          // Filter
          const filtered = allRelatedWi.filter(({ fields }) => {
            const t = fields?.['System.WorkItemType'];
            const s = fields?.['System.State'];
            return (t === 'Change Request' || t === 'Bug') && s !== 'Closed' && s !== 'Resolved';
          });

          mappedItem.linkItems = filtered.length > 0 ? this.MapLinkedWorkItem(filtered, project) : [];
          summarizedItemMap.set(wi.id, mappedItem);
        }
      } catch (error: any) {
        logger.error(`Error occurred while fetching linked work items: ${error.message}`);
        logger.error(`Error Stack: ${error.stack}`);
      }
    }

    return [...summarizedItemMap.values()];
  }

  /**
   * Mapping the linked work item of the OpenPcr table
   * @param wis Work item list
   * @param project
   * @returns array of mapped workitems
   */

  private MapLinkedWorkItem(wis: any[], project: string): any[] {
    return wis.map((item) => {
      const { id, fields } = item;
      return {
        pcrId: id,
        workItemType: fields['System.WorkItemType'],
        title: fields['System.Title'],
        severity: fields['Microsoft.VSTS.Common.Severity'] || '',
        pcrUrl: `${this.orgUrl}${project}/_workitems/edit/${id}`,
      };
    });
  }

  /**
   * Fetching Open PCRs data
   */

  private async fetchOpenPcrData(
    testItems: any[],
    projectName: string,
    openPcrToTestCaseTraceMap: Map<string, string[]>,
    testCaseToOpenPcrTraceMap: Map<string, string[]>
  ) {
    const linkedWorkItems = await this.fetchLinkedWi(projectName, testItems);
    for (const wi of linkedWorkItems) {
      const { linkItems, ...restItem } = wi;
      const stringifiedTestCase = JSON.stringify({
        id: restItem.testId,
        title: restItem.testName,
        testCaseUrl: restItem.testCaseUrl,
        runStatus: restItem.runStatus,
      });
      for (const linkedItem of linkItems) {
        const stringifiedPcr = JSON.stringify({
          pcrId: linkedItem.pcrId,
          workItemType: linkedItem.workItemType,
          title: linkedItem.title,
          severity: linkedItem.severity,
          pcrUrl: linkedItem.pcrUrl,
        });
        DataProviderUtils.addToTraceMap(openPcrToTestCaseTraceMap, stringifiedPcr, stringifiedTestCase);
        DataProviderUtils.addToTraceMap(testCaseToOpenPcrTraceMap, stringifiedTestCase, stringifiedPcr);
      }
    }
  }

  /**
   * Fetch Test log data
   */
  private formatUtcToLocalDateTimeString(utcDateString: string, ianaTimeZone: string): string {
    try {
      const date = new Date(utcDateString);
      return new Intl.DateTimeFormat('en-US', {
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit',
        hour12: true,
        timeZone: ianaTimeZone,
      }).format(date);
    } catch (error) {
      logger.error(`Error formatting date string "${utcDateString}" to timezone "${ianaTimeZone}":`, error);
      // Fallback to original UTC string if formatting fails, or handle as per requirements
      return utcDateString;
    }
  }

  private fetchTestLogData(testItems: any[], combinedResults: any[]) {
    const processedItems = testItems
      .filter((item) => item.lastResultDetails && item.lastResultDetails.runBy.displayName !== null)
      .map((item) => {
        const { dateCompleted, runBy } = item.lastResultDetails;
        return {
          testId: item.testCaseId,
          testName: item.testCaseName,
          originalUtcDate: new Date(dateCompleted), // For sorting
          utcDateString: dateCompleted, // For formatting
          performedBy: runBy.displayName,
        };
      })
      .sort((a, b) => b.originalUtcDate.getTime() - a.originalUtcDate.getTime());

    const testLogData = processedItems.map((item) => {
      return {
        testId: item.testId,
        testName: item.testName,
        executedDate: this.formatUtcToLocalDateTimeString(item.utcDateString, 'Asia/Jerusalem'),
        performedBy: item.performedBy,
      };
    });

    if (testLogData?.length > 0) {
      // Add openPCR to combined results
      combinedResults.push({
        contentControl: 'test-execution-content-control',
        data: testLogData,
        skin: 'test-log-table',
        insertPageBreak: true,
      });
    }
  }

  private CreateAttachmentPathIndexMap(actionResults: any) {
    const attachmentPathToIndexMap: Map<string, number> = new Map();

    for (let i = 0; i < actionResults.length; i++) {
      const actionPath = actionResults[i].actionPath;
      attachmentPathToIndexMap.set(actionPath, actionResults[i].stepPosition);
    }
    return attachmentPathToIndexMap;
  }

  private mapStepResultsForExecutionAppendix(detailedResults: any[], runResultData: any[]): Map<string, any> {
    // Create maps first to avoid repeated lookups
    const testCaseIdToStepsMap = new Map<string, any>();
    const actionPathToPositionMap = this.mapActionPathToPosition(detailedResults);

    // Pre-process detailedResults to set up the basic structure for testCaseIdToStepsMap
    detailedResults.forEach((result) => {
      const testCaseId = result.testId.toString();
      if (!testCaseIdToStepsMap.has(testCaseId)) {
        testCaseIdToStepsMap.set(testCaseId, {
          ...result.testCaseRevision,
          stepList: [],
          caseEvidenceAttachments: [],
        });
      }

      testCaseIdToStepsMap.get(testCaseId).stepList.push({
        stepPosition: result.stepNo,
        stepId: result.stepIdentifier,
        action: result.stepAction,
        expected: result.stepExpected,
        stepStatus: result.stepStatus,
        stepComments: result.stepComments,
        isSharedStepTitle: result.isSharedStepTitle,
      });
    });

    // Process run attachments in a single pass
    for (const result of runResultData) {
      const testCaseId = result.testCaseId.toString();
      const testCase = testCaseIdToStepsMap.get(testCaseId);

      if (!testCase) continue;

      // Process all attachments for this iteration at once
      result.iteration.attachments.forEach((attachment: any) => {
        const position = `${result.testCaseId}-${attachment.actionPath}`;
        const stepNo = actionPathToPositionMap.get(position) || '';

        testCase.caseEvidenceAttachments.push({
          name: attachment.name,
          testCaseId: result.testCaseId,
          stepNo,
          downloadUrl: attachment.downloadUrl,
        });
      });
    }

    return testCaseIdToStepsMap;
  }

  private mapActionPathToPosition(actionResults: any[]): Map<string, number> {
    return new Map(actionResults.map((result) => [`${result.testId}-${result.actionPath}`, result.stepNo]));
  }

  /**
   * Calculates a summary of test group results.
   */
  private calculateGroupResultSummary(testPointsItems: any[], includeHardCopyRun: boolean): any {
    if (includeHardCopyRun) {
      return {
        passed: '',
        failed: '',
        notApplicable: '',
        blocked: '',
        notRun: '',
        total: '',
        successPercentage: '',
      };
    }

    const summary = {
      passed: 0,
      failed: 0,
      notApplicable: 0,
      blocked: 0,
      notRun: 0,
      total: testPointsItems.length,
      successPercentage: '0.00%',
    };

    testPointsItems.forEach((item) => {
      switch (item.outcome) {
        case 'passed':
          summary.passed++;
          break;
        case 'failed':
          summary.failed++;
          break;
        case 'notApplicable':
          summary.notApplicable++;
          break;
        case 'blocked':
          summary.blocked++;
          break;
        default:
          summary.notRun++;
      }
    });

    summary.successPercentage =
      summary.total > 0 ? `${((summary.passed / summary.total) * 100).toFixed(2)}%` : '0.00%';

    return summary;
  }

  /**
   * Calculates the total summary of all test group results.
   */
  private calculateTotalSummary(results: any[], includeHardCopyRun: boolean): any {
    if (includeHardCopyRun) {
      return {
        passed: '',
        failed: '',
        notApplicable: '',
        blocked: '',
        notRun: '',
        total: '',
        successPercentage: '',
      };
    }
    const totalSummary = results.reduce(
      (acc, { groupResultSummary }) => {
        acc.passed += groupResultSummary.passed;
        acc.failed += groupResultSummary.failed;
        acc.notApplicable += groupResultSummary.notApplicable;
        acc.blocked += groupResultSummary.blocked;
        acc.notRun += groupResultSummary.notRun;
        acc.total += groupResultSummary.total;
        return acc;
      },
      {
        passed: 0,
        failed: 0,
        notApplicable: 0,
        blocked: 0,
        notRun: 0,
        total: 0,
        successPercentage: '0.00%',
      }
    );

    totalSummary.successPercentage =
      totalSummary.total > 0 ? `${((totalSummary.passed / totalSummary.total) * 100).toFixed(2)}%` : '0.00%';

    return totalSummary;
  }

  /**
   * Flattens the test points for easier access.
   */
  private flattenTestPoints(testPoints: any[]): any[] {
    return testPoints
      .filter((point) => point.testPointsItems && point.testPointsItems.length > 0)
      .flatMap((point) => {
        const { testPointsItems, ...restOfPoint } = point;
        return testPointsItems.map((item: any) => ({ ...restOfPoint, ...item }));
      });
  }

  /**
   * Formats a test result for display.
   */
  private formatTestResult(testPoint: any, addConfiguration: boolean, includeHardCopyRun: boolean): any {
    const formattedResult: any = {
      testGroupName: testPoint.testGroupName,
      testId: testPoint.testCaseId,
      testName: testPoint.testCaseName,
      testCaseUrl: testPoint.testCaseUrl,
      runStatus: !includeHardCopyRun ? this.convertRunStatus(testPoint.outcome) : '',
    };

    if (addConfiguration) {
      formattedResult.configuration = testPoint.configurationName;
    }

    return formattedResult;
  }

  /**
   * Fetches result data based on the Work Item Test Reporter.
   *
   * This method retrieves detailed result data for a specific test run and result ID,
   * including related work items, selected fields, and additional processing options.
   *
   * @param projectName - The name of the project containing the test run.
   * @param runId - The unique identifier of the test run.
   * @param resultId - The unique identifier of the test result.
   * @param selectedFields - (Optional) An array of field names to include in the result data.
   * @returns A promise that resolves to the fetched result data.
   */
  private async fetchResultDataBasedOnWiTestReporter(
    projectName: string,
    runId: string,
    resultId: string,
    selectedFields?: string[],
    isQueryMode?: boolean,
    point?: any,
    includeAllHistory: boolean = false
  ): Promise<any> {
    return this.fetchResultDataBasedOnWiBase(
      projectName,
      runId,
      resultId,
      true,
      selectedFields,
      isQueryMode,
      point,
      includeAllHistory
    );
  }

  /**
   * Fetches all result data for the test reporter by processing the provided test data.
   *
   * This method utilizes the `fetchAllResultDataBase` function to retrieve and process
   * result data for a specific project and test reporter. It applies a callback to fetch
   * result data for individual test points.
   *
   * @param testData - An array of test data objects to process.
   * @param projectName - The name of the project for which result data is being fetched.
   * @param selectedFields - An optional array of field names to include in the result data.
   * @returns A promise that resolves to an array of processed result data.
   */
  private async fetchAllResultDataTestReporter(
    testData: any[],
    projectName: string,
    selectedFields?: string[],
    isQueryMode?: boolean,
    includeAllHistory: boolean = false
  ): Promise<any[]> {
    return this.fetchAllResultDataBase(
      testData,
      projectName,
      true,
      (projectName, testSuiteId, point, selectedFields, isQueryMode, includeAllHistory) =>
        this.fetchResultDataForTestReporter(
          projectName,
          testSuiteId,
          point,
          selectedFields,
          isQueryMode,
          includeAllHistory
        ),
      [selectedFields, isQueryMode, includeAllHistory]
    );
  }

  /**
   * Aligns test steps with iterations for the test reporter by processing test data and iterations
   * and generating a structured result object based on the provided selected fields.
   *
   * @param testData - An array of test data objects to be processed.
   * @param iterations - An array of iteration objects to align with the test data.
   * @param selectedFields - An array of selected fields to determine which properties to include in the response.
   * @returns An array of structured result objects containing aligned test steps and iterations.
   *
   * The method uses a base alignment function and provides custom logic for:
   * - Determining whether step-level processing should occur based on the fetched test case and filtered fields.
   * - Creating a result object with properties such as suite name, test case details, priority, run information,
   *   failure type, automation status, execution date, configuration name, state, error message, and related requirements.
   * - Including step-specific properties (e.g., step number, action, expected result, status, and comments) if action results are available
   *   and the corresponding fields are selected.
   */
  private alignStepsWithIterationsTestReporter(
    testData: any[],
    iterations: any[],
    selectedFields: any[],
    includeNotRunTestCases: boolean
  ): any[] {
    return this.alignStepsWithIterationsBase(testData, iterations, includeNotRunTestCases, true, true, {
      selectedFields,
      shouldProcessStepLevel: (fetchedTestCase, filteredFields) =>
        fetchedTestCase != null &&
        fetchedTestCase.iteration != null &&
        fetchedTestCase.iteration.actionResults != null &&
        filteredFields.size > 0,

      createResultObject: ({
        testItem,
        point,
        fetchedTestCase,
        actionResult,
        filteredFields = new Set(),
      }) => {
        const baseObj: any = {
          suiteName: testItem.testGroupName,
          testCase: {
            id: point.testCaseId,
            title: point.testCaseName,
            url: point.testCaseUrl,
            result: fetchedTestCase.testCaseResult,
            comment: fetchedTestCase.comment,
          },
          runBy: fetchedTestCase.runBy,
          failureType: fetchedTestCase.failureType,
          executionDate: fetchedTestCase.executionDate,
          configurationName: fetchedTestCase.configurationName,
          relatedRequirements: fetchedTestCase.relatedRequirements,
          relatedBugs: fetchedTestCase.relatedBugs,
          relatedCRs: fetchedTestCase.relatedCRs,
          ...fetchedTestCase.customFields,
        };

        // If we have action results, add step-specific properties
        if (actionResult) {
          return {
            ...baseObj,
            stepNo:
              filteredFields.has('includeSteps') ||
              filteredFields.has('stepRunStatus') ||
              filteredFields.has('testStepComment')
                ? actionResult.stepPosition
                : undefined,
            stepAction: filteredFields.has('includeSteps') ? actionResult.action : undefined,
            stepExpected: filteredFields.has('includeSteps') ? actionResult.expected : undefined,
            stepStatus: filteredFields.has('stepRunStatus')
              ? this.convertUnspecifiedRunStatus(actionResult)
              : undefined,
            stepComments: filteredFields.has('testStepComment') ? actionResult.errorMessage : undefined,
          };
        }

        return baseObj;
      },
    });
  }

  /**
   * Builds flat test reporter rows that include suite/testcase/run/step data.
   */
  private alignStepsWithIterationsFlatReport(
    testData: any[],
    iterations: any[],
    includeNotRunTestCases: boolean,
    planId: string,
    planName: string
  ): any[] {
    return this.alignStepsWithIterationsBase(testData, iterations, includeNotRunTestCases, true, true, {
      selectedFields: [],
      shouldProcessStepLevel: (fetchedTestCase) =>
        fetchedTestCase != null &&
        fetchedTestCase.iteration != null &&
        fetchedTestCase.iteration.actionResults != null,
      createResultObject: ({ testItem, point, fetchedTestCase, actionResult }) => {
        const suiteId = testItem?.testSuiteId ?? testItem?.suiteId;
        const suiteName = testItem?.suiteName ?? testItem?.testGroupName ?? '';
        const parentSuiteId = testItem?.parentSuiteId;
        const parentSuiteName = testItem?.parentSuiteName;
        const suitePath = testItem?.suitePath ?? '';
        const customFields = fetchedTestCase?.customFields ?? {};
        const toNumber = (value: any) => {
          if (value === null || value === undefined) return undefined;
          const n = Number.parseInt(String(value), 10);
          return Number.isFinite(n) ? n : undefined;
        };
        const parsedStepIdentifier = toNumber(actionResult?.stepIdentifier);

        return {
          planId,
          planName,
          suiteId,
          suiteName,
          parentSuiteId,
          parentSuiteName,
          suitePath,
          testCaseId: point?.testCaseId,
          testCaseName: point?.testCaseName,
          customFields,
          pointOutcome: point?.outcome,
          testCaseResultMessage: fetchedTestCase?.testCaseResult?.resultMessage ?? '',
          executionDate: fetchedTestCase?.executionDate ?? '',
          runDateCompleted: point?.lastResultDetails?.dateCompleted ?? fetchedTestCase?.executionDate ?? '',
          runStatsOutcome: point?.lastResultDetails?.outcome ?? point?.outcome,
          testRunId: point?.lastRunId,
          testPointId: point?.testPointId,
          tester: fetchedTestCase?.runBy ?? point?.lastResultDetails?.runBy?.displayName ?? '',
          stepOutcome: actionResult?.outcome,
          stepStepIdentifier: actionResult?.stepPosition ?? '',
        };
      },
    });
  }

  /**
   * Fetches result data for a test reporter based on the provided project, test suite, and point information.
   * This method processes the result data and formats it according to the selected fields.
   *
   * @param projectName - The name of the project.
   * @param testSuiteId - The ID of the test suite.
   * @param point - The test point containing details such as last run ID, result ID, and configuration name.
   * @param selectedFields - An optional array of field names to filter and include in the response.
   *
   * @returns A promise that resolves to the formatted result data object containing details about the test case,
   *          test suite, last run, iteration, and other selected fields.
   */
  private async fetchResultDataForTestReporter(
    projectName: string,
    testSuiteId: string,
    point: any,
    selectedFields?: string[],
    isQueryMode?: boolean,
    includeAllHistory: boolean = false
  ) {
    return this.fetchResultDataBase(
      projectName,
      testSuiteId,
      point,
      (project, runId, resultId, fields, isQueryMode, point, includeAllHistory) =>
        this.fetchResultDataBasedOnWiTestReporter(
          project,
          runId,
          resultId,
          fields,
          isQueryMode,
          point,
          includeAllHistory
        ),
      (resultData, testSuiteId, point, selectedFields) => {
        const { lastRunId, lastResultId, configurationName, lastResultDetails } = point;
        try {
          const iteration =
            resultData.iterationDetails?.length > 0
              ? resultData.iterationDetails[resultData.iterationDetails?.length - 1]
              : undefined;

          if (!resultData?.testCase || !resultData?.testSuite) {
            logger.debug(
              `[RunResult] Missing testCase/testSuite for point testCaseId=${String(
                point?.testCaseId ?? 'unknown'
              )} (lastRunId=${String(lastRunId ?? '')}, lastResultId=${String(
                lastResultId ?? ''
              )}). hasTestCase=${Boolean(resultData?.testCase)} hasTestSuite=${Boolean(
                resultData?.testSuite
              )}`
            );
          }
          const resultDataResponse: any = {
            testCaseName: `${resultData?.testCase?.name ?? ''} - ${resultData?.testCase?.id ?? ''}`,
            testCaseId: resultData?.testCase?.id,
            testSuiteName: `${resultData?.testSuite?.name ?? ''}`,
            testSuiteId,
            lastRunId,
            lastResultId,
            iteration,
            testCaseRevision: resultData.testCaseRevision,
            resolution: resultData.resolutionState,
            failureType: undefined as string | undefined,
            runBy: undefined as string | undefined,
            executionDate: undefined as string | undefined,
            testCaseResult: undefined as any,
            comment: undefined as string | undefined,
            configurationName: undefined as string | undefined,
            relatedRequirements: resultData.relatedRequirements || undefined,
            relatedBugs: resultData.relatedBugs || undefined,
            relatedCRs: resultData.relatedCRs || undefined,
            lastRunResult: undefined as any,
            customFields: {}, // Create an object to store custom fields
          };

          // Process all custom fields from resultData.filteredFields
          if (resultData.filteredFields) {
            const customFields = this.standardCustomField(resultData.filteredFields);
            resultDataResponse.customFields = customFields;
          }

          const filteredFields = selectedFields
            ?.filter((field: string) => field.includes('@runResultField'))
            ?.map((field: string) => field.split('@')[0]);

          if (filteredFields && filteredFields.length > 0) {
            for (const field of filteredFields) {
              switch (field) {
                case 'priority':
                  resultDataResponse.priority = resultData.priority;
                  break;
                case 'testCaseResult':
                  const outcome = this.getTestOutcome(resultData);
                  if (lastRunId === undefined || lastResultId === undefined) {
                    resultDataResponse.testCaseResult = {
                      resultMessage: `${this.convertRunStatus(outcome)}`,
                      url: '',
                    };
                  } else {
                    resultDataResponse.testCaseResult = {
                      resultMessage: `${this.convertRunStatus(outcome)} in Run ${lastRunId}`,
                      url: `${this.orgUrl}${projectName}/_testManagement/runs?runId=${lastRunId}&_a=resultSummary&resultId=${lastResultId}`,
                    };
                  }
                  break;
                case 'testCaseComment':
                  resultDataResponse.comment = iteration?.comment;
                  break;
                case 'failureType':
                  resultDataResponse.failureType = resultData.failureType;
                  break;
                case 'runBy':
                  if (!lastResultDetails?.runBy?.displayName) {
                    logger.debug(
                      `[RunResult] Missing runBy for testCaseId=${String(
                        resultData?.testCase?.id ?? point?.testCaseId ?? 'unknown'
                      )} (lastRunId=${String(lastRunId ?? '')}, lastResultId=${String(
                        lastResultId ?? ''
                      )}). lastResultDetails=${this.stringifyForDebug(lastResultDetails, 2000)}`
                    );
                  }
                  resultDataResponse.runBy = lastResultDetails?.runBy?.displayName ?? '';
                  break;
                case 'executionDate':
                  if (!lastResultDetails?.dateCompleted) {
                    logger.debug(
                      `[RunResult] Missing dateCompleted for testCaseId=${String(
                        resultData?.testCase?.id ?? point?.testCaseId ?? 'unknown'
                      )} (lastRunId=${String(lastRunId ?? '')}, lastResultId=${String(
                        lastResultId ?? ''
                      )}). lastResultDetails=${this.stringifyForDebug(lastResultDetails, 2000)}`
                    );
                  }
                  resultDataResponse.executionDate = lastResultDetails?.dateCompleted ?? '';
                  break;
                case 'configurationName':
                  resultDataResponse.configurationName = configurationName;
                  break;
                default:
                  logger.debug(`Field ${field} not handled`);
                  break;
              }
            }
          }
          return resultDataResponse;
        } catch (err: any) {
          logger.error(`Error occurred while fetching result data: ${err.message}`);
          logger.error(`Error stack: ${err.stack}`);
          return null;
        }
      },
      [selectedFields, isQueryMode, point, includeAllHistory]
    );
  }

  private getTestOutcome(resultData: any) {
    // Default outcome if nothing else is available
    const defaultOutcome = 'NotApplicable';

    // Check if we have iteration details
    const hasIterationDetails = resultData?.iterationDetails && resultData.iterationDetails.length > 0;

    if (hasIterationDetails) {
      // Get the last iteration's outcome if available
      const lastIteration = resultData.iterationDetails[resultData.iterationDetails.length - 1];
      return lastIteration?.outcome || resultData.outcome || defaultOutcome;
    }

    // No iteration details, use result outcome or default
    return resultData.outcome || defaultOutcome;
  }

  private standardCustomField(fieldsToProcess: any, selectedColumns?: any[]): any {
    const customFields: any = {};
    if (selectedColumns) {
      const standardFields = ['id', 'title', 'workItemType'];

      for (const column of selectedColumns) {
        const fieldName = column.referenceName;
        let propertyName = column.name.replace(/\s+/g, '');
        if (propertyName === propertyName.toUpperCase()) {
          propertyName = propertyName.toLowerCase();
        } else {
          propertyName = propertyName.charAt(0).toLowerCase() + propertyName.slice(1);
        }

        if (standardFields.includes(propertyName)) {
          continue;
        }
        const fieldValue = fieldsToProcess[fieldName];
        if (fieldValue === undefined || fieldValue === null) {
          customFields[propertyName] = null;
        } else {
          customFields[propertyName] = (fieldValue as any)?.displayName ?? fieldValue;
        }
      }
    } else {
      for (const [fieldName, fieldValue] of Object.entries(fieldsToProcess)) {
        const nameParts = fieldName.split('.');
        let propertyName = nameParts[nameParts.length - 1];
        propertyName = propertyName.charAt(0).toLowerCase() + propertyName.slice(1);
        customFields[propertyName] = (fieldValue as any)?.displayName ?? fieldValue ?? '';
      }
    }

    return customFields;
  }
}
