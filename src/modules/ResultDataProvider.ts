import { log } from 'winston';
import { TFSServices } from '../helpers/tfs';
import { TestSteps } from '../models/tfs-data';
import logger from '../utils/logger';
import TestStepParserHelper from '../utils/testStepParserHelper';
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
  private testStepParserHelper: TestStepParserHelper;
  constructor(orgUrl: string, token: string) {
    this.orgUrl = orgUrl;
    this.token = token;
    this.testStepParserHelper = new TestStepParserHelper(orgUrl, token);
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
    includeOpenPCRs: boolean = false,
    includeTestLog: boolean = false,
    stepExecution?: any,
    stepAnalysis?: any,
    includeHardCopyRun: boolean = false
  ): Promise<any[]> {
    const combinedResults: any[] = [];
    try {
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
      const testData = await this.fetchTestData(suites, projectName, testPlanId);
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

      if (includeOpenPCRs) {
        //5. Open PCRs data (only if enabled)
        await this.fetchOpenPcrData(testResultsSummary, projectName, combinedResults);
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

      return combinedResults;
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
    enableRunTestCaseFilter: boolean,
    enableRunStepStatusFilter: boolean
  ) {
    const fetchedTestResults: any[] = [];
    logger.debug(
      `Fetching test reporter results for test plan ID: ${testPlanId}, project name: ${projectName}`
    );
    logger.debug(`Selected suite IDs: ${selectedSuiteIds}`);
    try {
      const plan = await this.fetchTestPlanName(testPlanId, projectName);
      const suites = await this.fetchTestSuites(testPlanId, projectName, selectedSuiteIds, true);
      const testData = await this.fetchTestData(suites, projectName, testPlanId);
      const runResults = await this.fetchAllResultDataTestReporter(testData, projectName, selectedFields);
      const testReporterData = this.alignStepsWithIterationsTestReporter(
        testData,
        runResults,
        selectedFields,
        !enableRunTestCaseFilter
      );

      // Apply filters sequentially based on enabled flags
      let filteredResults = testReporterData;

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

      return filteredSuites.map((testSuite: any) => ({
        testSuiteId: testSuite.id,
        testGroupName: this.buildTestGroupName(testSuite.id, suiteMap, isHierarchyGroupName),
      }));
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
      testCaseId: testPoint.testCaseReference.id,
      testCaseName: testPoint.testCaseReference.name,
      testCaseUrl: `${this.orgUrl}${projectName}/_workitems/edit/${testPoint.testCaseReference.id}`,
      configurationName: testPoint.configuration?.name,
      outcome: testPoint.results?.outcome || 'Not Run',
      lastRunId: testPoint.results?.lastTestRunId,
      lastResultId: testPoint.results?.lastResultId,
      lastResultDetails: testPoint.results?.lastResultDetails,
    };
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
   * Fetches result data based on a work item base (WiBase) for a specific test run and result.
   *
   * @param projectName - The name of the project in Azure DevOps.
   * @param runId - The ID of the test run.
   * @param resultId - The ID of the test result.
   * @param options - Optional parameters for customizing the data retrieval.
   * @param options.expandWorkItem - If true, expands all fields of the work item.
   * @param options.selectedFields - An array of field names to filter the work item fields.
   * @param options.processRelatedRequirements - If true, processes related requirements linked to the work item.
   * @param options.includeFullErrorStack - If true, includes the full error stack in the logs when an error occurs.
   * @returns A promise that resolves to an object containing the fetched result data, including:
   * - `stepsResultXml`: The test steps result in XML format.
   * - `analysisAttachments`: Attachments related to the test result analysis.
   * - `testCaseRevision`: The revision number of the test case.
   * - `filteredFields`: The filtered fields from the work item based on the selected fields.
   * - `relatedRequirements`: An array of related requirements with details such as ID, title, customer ID, and URL.
   * - `relatedBugs`: An array of related bugs with details such as ID, title, and URL.
   * If an error occurs, logs the error and returns `null`.
   *
   * @throws Logs an error message if the data retrieval fails.
   */
  private async fetchResultDataBasedOnWiBase(
    projectName: string,
    runId: string,
    resultId: string,
    options: {
      expandWorkItem?: boolean;
      selectedFields?: string[];
      processRelatedRequirements?: boolean;
      processRelatedBugs?: boolean;
      includeFullErrorStack?: boolean;
    } = {}
  ): Promise<any> {
    try {
      const url = `${this.orgUrl}${projectName}/_apis/test/runs/${runId}/results/${resultId}?detailsToInclude=Iterations`;
      const resultData = await TFSServices.getItemContent(url, this.token);

      const attachmentsUrl = `${this.orgUrl}${projectName}/_apis/test/runs/${runId}/results/${resultId}/attachments`;
      const { value: analysisAttachments } = await TFSServices.getItemContent(attachmentsUrl, this.token);

      // Build workItem URL with optional expand parameter
      const expandParam = options.expandWorkItem ? '?$expand=all' : '';
      const wiUrl = `${this.orgUrl}${projectName}/_apis/wit/workItems/${resultData.testCase.id}/revisions/${resultData.testCaseRevision}${expandParam}`;
      const wiByRevision = await TFSServices.getItemContent(wiUrl, this.token);
      let filteredFields: any = {};
      let relatedRequirements: any[] = [];
      let relatedBugs: any[] = [];
      //TODO: Add CR support as well, and also add the logic to fetch the CR details
      // TODO: Add logic for grabbing the relations from cross projects

      // Process selected fields if provided
      if (
        options.selectedFields?.length &&
        (options.processRelatedRequirements || options.processRelatedBugs)
      ) {
        const filtered = options.selectedFields
          ?.filter((field: string) => field.includes('@testCaseWorkItemField'))
          ?.map((field: string) => field.split('@')[0]);
        const selectedFieldSet = new Set(filtered);

        if (selectedFieldSet.size !== 0) {
          // Process related requirements if needed
          const { relations } = wiByRevision;
          if (relations) {
            for (const relation of relations) {
              if (
                relation.rel?.includes('System.LinkTypes.Hierarchy') ||
                relation.rel?.includes('Microsoft.VSTS.Common.TestedBy')
              ) {
                const relatedUrl = relation.url;
                const wi = await TFSServices.getItemContent(relatedUrl, this.token);
                if (wi.fields['System.WorkItemType'] === 'Requirement') {
                  const { id, fields, _links } = wi;
                  const requirementTitle = fields['System.Title'];
                  const customerFieldKey = Object.keys(fields).find((key) =>
                    key.toLowerCase().includes('customer')
                  );
                  const customerId = customerFieldKey ? fields[customerFieldKey] : undefined;
                  const url = _links.html.href;
                  relatedRequirements.push({ id, requirementTitle, customerId, url });
                } else if (wi.fields['System.WorkItemType'] === 'Bug') {
                  const { id, fields, _links } = wi;
                  const bugTitle = fields['System.Title'];
                  const url = _links.html.href;
                  relatedBugs.push({ id, bugTitle, url });
                }
              }
            }
          }

          // Filter fields based on selected field set
          filteredFields = Object.keys(wiByRevision.fields)
            .filter((key) => selectedFieldSet.has(key))
            .reduce((obj: any, key) => {
              obj[key] = wiByRevision.fields[key]?.displayName ?? wiByRevision.fields[key];
              return obj;
            }, {});
        }
      }
      return {
        ...resultData,
        stepsResultXml: wiByRevision.fields['Microsoft.VSTS.TCM.Steps'] || undefined,
        analysisAttachments,
        testCaseRevision: resultData.testCaseRevision,
        filteredFields,
        relatedRequirements,
        relatedBugs,
      };
    } catch (error: any) {
      logger.error(`Error while fetching run result: ${error.message}`);
      if (options.includeFullErrorStack) {
        logger.error(`Error stack: ${error.stack}`);
      }
      return null;
    }
  }

  private isNotRunStep = (result: any): boolean => {
    return result && result.stepStatus === 'Not Run';
  };

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

    for (const testItem of testData) {
      for (const point of testItem.testPointsItems) {
        const testCase = testItem.testCasesItems.find((tc: any) => tc.workItem.id === point.testCaseId);
        if (!testCase) continue;

        if (testCase.workItem.workItemFields.length === 0) {
          logger.warn(`Could not fetch the steps from WI ${JSON.stringify(testCase.workItem.id)}`);
          continue;
        }

        if (includeNotRunTestCases && !point.lastRunId && !point.lastResultId) {
          this.AppendResults(options, testCase, filteredFields, testItem, point, detailedResults);
        } else {
          const iterationsMap = this.createIterationsMap(iterations, testCase.workItem.id, isTestReporter);

          const iterationKey = `${point.lastRunId}-${point.lastResultId}-${testCase.workItem.id}`;
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
    testCaseId: number,
    isTestReporter: boolean
  ): Record<string, any> {
    return iterations.reduce((map, iterationItem) => {
      if (
        (isTestReporter && iterationItem.lastRunId && iterationItem.lastResultId) ||
        iterationItem.iteration
      ) {
        const key = `${iterationItem.lastRunId}-${iterationItem.lastResultId}-${testCaseId}`;
        map[key] = iterationItem;
      }
      return map;
    }, {} as Record<string, any>);
  }

  /**
   * Fetches test data for all suites, including test points and test cases.
   */
  private async fetchTestData(suites: any[], projectName: string, testPlanId: string): Promise<any[]> {
    return await Promise.all(
      suites.map((suite) =>
        this.limit(async () => {
          try {
            const testPointsItems = await this.fetchTestPoints(projectName, testPlanId, suite.testSuiteId);
            const testCasesItems = await this.fetchTestCasesBySuiteId(
              projectName,
              testPlanId,
              suite.testSuiteId
            );
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
    fetchStrategy: (projectName: string, testSuiteId: string, point: any, ...args: any[]) => Promise<any>,
    additionalArgs: any[] = []
  ): Promise<any[]> {
    const results = [];

    for (const item of testData) {
      if (item.testPointsItems && item.testPointsItems.length > 0) {
        const { testSuiteId, testPointsItems } = item;

        // Filter and sort valid points
        const validPoints = testPointsItems.filter(
          (point: any) => point && point.lastRunId && point.lastResultId
        );

        // Fetch results for each point sequentially
        for (const point of validPoints) {
          const resultData = await fetchStrategy(projectName, testSuiteId, point, ...additionalArgs);
          if (resultData !== null) {
            results.push(resultData);
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
    return this.fetchAllResultDataBase(testData, projectName, (projectName, testSuiteId, point) =>
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
   * Base method for fetching result data for test points
   */
  private async fetchResultDataBase(
    projectName: string,
    testSuiteId: string,
    point: any,
    fetchResultMethod: (project: string, runId: string, resultId: string, ...args: any[]) => Promise<any>,
    createResponseObject: (resultData: any, testSuiteId: string, point: any, ...args: any[]) => any,
    additionalArgs: any[] = []
  ): Promise<any> {
    const { lastRunId, lastResultId } = point;
    const resultData = await fetchResultMethod(
      projectName,
      lastRunId.toString(),
      lastResultId.toString(),
      ...additionalArgs
    );

    const iteration =
      resultData.iterationDetails.length > 0
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
        testCaseName: `${resultData.testCase.name} - ${resultData.testCase.id}`,
        testCaseId: resultData.testCase.id,
        testSuiteName: `${resultData.testSuite.name}`,
        testSuiteId,
        lastRunId: point.lastRunId,
        lastResultId: point.lastResultId,
        iteration:
          resultData.iterationDetails.length > 0
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

  private async fetchOpenPcrData(testItems: any[], projectName: string, combinedResults: any[]) {
    const linkedWorkItems = await this.fetchLinkedWi(projectName, testItems);
    const flatOpenPcrsItems = linkedWorkItems
      .filter((item) => item.linkItems.length > 0)
      .flatMap((item) => {
        const { linkItems, ...restItem } = item;
        return linkItems.map((linkedItem: any) => ({ ...restItem, ...linkedItem }));
      });
    if (flatOpenPcrsItems?.length > 0) {
      // Add openPCR to combined results
      combinedResults.push({
        contentControl: 'open-pcr-content-control',
        data: flatOpenPcrsItems,
        skin: 'open-pcr-table',
      });
    }
  }

  /**
   * Fetch Test log data
   */
  private fetchTestLogData(testItems: any[], combinedResults: any[]) {
    const testLogData = testItems
      .filter((item) => item.lastResultDetails && item.lastResultDetails.runBy.displayName !== null)
      .map((item) => {
        const { dateCompleted, runBy } = item.lastResultDetails;
        return {
          testId: item.testCaseId,
          testName: item.testCaseName,
          executedDate: dateCompleted,
          performedBy: runBy.displayName,
        };
      })
      .sort((a, b) => new Date(b.executedDate).getTime() - new Date(a.executedDate).getTime());

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
    selectedFields?: string[]
  ): Promise<any> {
    return this.fetchResultDataBasedOnWiBase(projectName, runId, resultId, {
      expandWorkItem: true,
      selectedFields,
      processRelatedRequirements: true,
      processRelatedBugs: true,
      includeFullErrorStack: true,
    });
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
    selectedFields?: string[]
  ): Promise<any[]> {
    return this.fetchAllResultDataBase(
      testData,
      projectName,
      (projectName, testSuiteId, point, selectedFields) =>
        this.fetchResultDataForTestReporter(projectName, testSuiteId, point, selectedFields),
      [selectedFields]
    );
  }

  /**
   * Aligns test steps with iterations for the test reporter by processing test data and iterations
   * and generating a structured result object based on the provided selected fields.
   *
   * @param testData - An array of test data objects to be processed.
   * @param iterations - An array of iteration objects to align with the test data.
   * @param selectedFields - An array of selected fields to determine which properties to include in the result.
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
        const baseObj = {
          suiteName: testItem.testGroupName,
          testCase: {
            id: point.testCaseId,
            title: point.testCaseName,
            url: point.testCaseUrl,
            result: fetchedTestCase.testCaseResult,
            comment: fetchedTestCase.iteration?.comment,
          },
          priority: fetchedTestCase.priority,
          runBy: fetchedTestCase.runBy,
          activatedBy: fetchedTestCase.activatedBy,
          assignedTo: fetchedTestCase.assignedTo,
          failureType: fetchedTestCase.failureType,
          automationStatus: fetchedTestCase.automationStatus,
          executionDate: fetchedTestCase.executionDate,
          configurationName: fetchedTestCase.configurationName,
          errorMessage: fetchedTestCase.errorMessage,
          relatedRequirements: fetchedTestCase.relatedRequirements,
          relatedBugs: fetchedTestCase.relatedBugs,
        };

        // If we have action results, add step-specific properties
        if (actionResult) {
          return {
            ...baseObj,
            stepNo: actionResult.stepPosition,
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
    selectedFields?: string[]
  ) {
    return this.fetchResultDataBase(
      projectName,
      testSuiteId,
      point,
      (project, runId, resultId, fields) =>
        this.fetchResultDataBasedOnWiTestReporter(project, runId, resultId, fields),
      (resultData, testSuiteId, point, selectedFields) => {
        const { lastRunId, lastResultId, configurationName, lastResultDetails } = point;
        try {
          const iteration =
            resultData.iterationDetails.length > 0
              ? resultData.iterationDetails[resultData.iterationDetails.length - 1]
              : undefined;

          const resultDataResponse = {
            testCaseName: `${resultData.testCase.name} - ${resultData.testCase.id}`,
            testCaseId: resultData.testCase.id,
            testSuiteName: `${resultData.testSuite.name}`,
            testSuiteId,
            lastRunId,
            lastResultId,
            iteration,
            testCaseRevision: resultData.testCaseRevision,
            resolution: resultData.resolutionState,
            automationStatus: resultData.filteredFields['Microsoft.VSTS.TCM.AutomationStatus'] || undefined,
            analysisAttachments: resultData.analysisAttachments,
            failureType: undefined as string | undefined,
            comment: undefined as string | undefined,
            priority: undefined,
            runBy: undefined as string | undefined,
            executionDate: undefined as string | undefined,
            testCaseResult: undefined as any | undefined,
            errorMessage: undefined as string | undefined,
            configurationName: undefined as string | undefined,
            relatedRequirements: undefined,
            relatedBugs: undefined,
            lastRunResult: undefined as any,
          };

          const filteredFields = selectedFields
            ?.filter((field: string) => field.includes('@runResultField') || field.includes('@linked'))
            ?.map((field: string) => field.split('@')[0]);

          if (filteredFields && filteredFields.length > 0) {
            for (const field of filteredFields) {
              switch (field) {
                case 'priority':
                  resultDataResponse.priority = resultData.priority;
                  break;
                case 'testCaseResult':
                  const outcome =
                    resultData.iterationDetails[resultData.iterationDetails.length - 1]?.outcome ||
                    resultData.outcome ||
                    'NotApplicable';
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
                case 'failureType':
                  resultDataResponse.failureType = resultData.failureType;
                  break;
                case 'testCaseComment':
                  resultDataResponse.comment = resultData.comment || undefined;
                  break;
                case 'runBy':
                  const runBy = lastResultDetails.runBy.displayName;
                  resultDataResponse.runBy = runBy;
                  break;
                case 'executionDate':
                  resultDataResponse.executionDate = lastResultDetails.dateCompleted;
                  break;
                case 'configurationName':
                  resultDataResponse.configurationName = configurationName;
                  break;
                case 'associatedRequirement':
                  resultDataResponse.relatedRequirements = resultData.relatedRequirements;
                  break;
                case 'associatedBug':
                  resultDataResponse.relatedBugs = resultData.relatedBugs;
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
      [selectedFields]
    );
  }
}
