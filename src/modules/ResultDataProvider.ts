import DataProviderUtils from '../utils/DataProviderUtils';
import { TFSServices } from '../helpers/tfs';
import { OpenPcrRequest, PlainTestResult, TestSteps } from '../models/tfs-data';
import { AdoWorkItemComment, AdoWorkItemCommentsResponse } from '../models/ado-comments';
import logger from '../utils/logger';
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
  orgUrl: string = '';
  token: string = '';
  private limit = pLimit(10);
  private testStepParserHelper: Utils;
  private testToAssociatedItemMap: Map<number, Set<any>>;
  private querySelectedColumns: any[];
  private workItemDiscussionCache: Map<number, any[]>;
  constructor(orgUrl: string, token: string) {
    this.orgUrl = orgUrl;
    this.token = token;
    this.testStepParserHelper = new Utils(orgUrl, token);
    this.testToAssociatedItemMap = new Map<number, Set<any>>();
    this.querySelectedColumns = [];
    this.workItemDiscussionCache = new Map<number, any[]>();
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

      const iteration =
        resultData.iterationDetails?.length > 0
          ? resultData.iterationDetails[resultData.iterationDetails.length - 1]
          : undefined;

      if (resultData.stepsResultXml && iteration) {
        const actionResultsWithSharedModels = iteration.actionResults.filter(
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

        for (const actionResult of iteration.actionResults) {
          const step = stepMap.get(actionResult.stepIdentifier);
          if (step) {
            actionResult.stepPosition = step.stepPosition;
            actionResult.action = step.action;
            actionResult.expected = step.expected;
            actionResult.isSharedStepTitle = step.isSharedStepTitle;
          }
        }
        //Sort by step position
        iteration.actionResults = iteration.actionResults
          .filter((result: any) => result.stepPosition)
          .sort((a: any, b: any) => this.compareActionResults(a.stepPosition, b.stepPosition));
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
        const customFields = fetchedTestCase?.customFields ?? {};
        const toNumber = (value: any) => {
          if (value === null || value === undefined) return undefined;
          const n = Number.parseInt(String(value), 10);
          return Number.isFinite(n) ? n : undefined;
        };
        const stepPosition = actionResult?.stepPosition;
        const parsedStepIdentifier = toNumber(actionResult?.stepIdentifier);

        return {
          planId,
          planName,
          suiteId,
          suiteName,
          parentSuiteId,
          parentSuiteName,
          testCaseId: point?.testCaseId,
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
          stepStepIdentifier: stepPosition ?? parsedStepIdentifier,
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
