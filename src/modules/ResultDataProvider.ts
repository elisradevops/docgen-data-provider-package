import { TFSServices } from '../helpers/tfs';
import { TestSteps, Workitem } from '../models/tfs-data';
import * as xml2js from 'xml2js';
import logger from '../utils/logger';
import TestStepParserHelper from '../utils/testStepParserHelper';
const pLimit = require('p-limit');
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
   * Retrieves test suites by test plan ID and processes them into a flat structure with hierarchy names.
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

      return count !== 0 ? testPoints.map(this.mapTestPoint) : [];
    } catch (error: any) {
      logger.error(`Error during fetching Test Points: ${error.message}`);
      return [];
    }
  }

  /**
   * Maps raw test point data to a simplified object.
   */
  private mapTestPoint(testPoint: any): any {
    return {
      testCaseId: testPoint.testCaseReference.id,
      testCaseName: testPoint.testCaseReference.name,
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
   * Fetches iterations data by run and result IDs.
   */
  private async fetchResult(projectName: string, runId: string, resultId: string): Promise<any> {
    try {
      const url = `${this.orgUrl}${projectName}/_apis/test/runs/${runId}/results/${resultId}?detailsToInclude=Iterations`;
      const resultData = await TFSServices.getItemContent(url, this.token);
      const attachmentsUrl = `${this.orgUrl}${projectName}/_apis/test/runs/${runId}/results/${resultId}/attachments`;
      const { value: analysisAttachments } = await TFSServices.getItemContent(attachmentsUrl, this.token);
      const wiUrl = `${this.orgUrl}${projectName}/_apis/wit/workItems/${resultData.testCase.id}/revisions/${resultData.testCaseRevision}`;
      const wiByRevision = await TFSServices.getItemContent(wiUrl, this.token);

      return {
        ...resultData,
        stepsResultXml: wiByRevision.fields['Microsoft.VSTS.TCM.Steps'] || undefined,
        analysisAttachments,
        testCaseRevision: resultData.testCaseRevision,
      };
    } catch (error: any) {
      logger.error(`Error while fetching run result: ${error.message}`);
      return null;
    }
  }

  /**
   * Converts run status from API format to a more readable format.
   */
  private convertRunStatus(status: string): string {
    switch (status) {
      case 'passed':
        return 'Passed';
      case 'failed':
        return 'Failed';
      case 'notApplicable':
        return 'Not Applicable';
      default:
        return 'Not Run';
    }
  }

  private setRunStatus(actionResult: any) {
    if (actionResult.outcome === 'Unspecified' && actionResult.isSharedStepTitle) {
      return '';
    }

    return actionResult.outcome === 'Unspecified'
      ? 'Not Run'
      : actionResult.outcome !== 'Not Run'
      ? actionResult.outcome
      : '';
  }

  /**
   * Aligns test steps with their corresponding iterations.
   */
  private alignStepsWithIterations(testData: any[], iterations: any[]): any[] {
    const detailedResults: any[] = [];
    if (!iterations || iterations?.length === 0) {
      return detailedResults;
    }

    for (const testItem of testData) {
      for (const point of testItem.testPointsItems) {
        const testCase = testItem.testCasesItems.find((tc: any) => tc.workItem.id === point.testCaseId);
        if (!testCase) continue;
        const iterationsMap = this.createIterationsMap(iterations, testCase.workItem.id);
        if (testCase.workItem.workItemFields.length === 0) {
          logger.warn(`Could not fetch the steps from WI ${JSON.stringify(testCase.workItem.id)}`);
          continue;
        }

        if (point.lastRunId && point.lastResultId) {
          const iterationKey = `${point.lastRunId}-${point.lastResultId}-${testCase.workItem.id}`;

          const testCastObj = iterationsMap[iterationKey];
          if (!testCastObj || !testCastObj.iteration || !testCastObj.iteration.actionResults) continue;

          const { actionResults } = testCastObj?.iteration;

          for (let i = 0; i < actionResults?.length; i++) {
            const stepIdentifier = parseInt(actionResults[i].stepIdentifier, 10);
            const resultObj = {
              testId: point.testCaseId,
              testCaseRevision: testCastObj.testCaseRevision,
              testName: point.testCaseName,
              stepIdentifier: stepIdentifier,
              actionPath: actionResults[i].actionPath,
              stepNo: actionResults[i].stepPosition,
              stepAction: actionResults[i].action,
              stepExpected: actionResults[i].expected,
              isSharedStepTitle: actionResults[i].isSharedStepTitle,
              stepStatus: this.setRunStatus(actionResults[i]),
              stepComments: actionResults[i].errorMessage || '',
            };

            detailedResults.push(resultObj);
          }
        }
      }
    }
    return detailedResults;
  }

  /**
   * Creates a mapping of iterations by their unique keys.
   */
  private createIterationsMap(iterations: any[], testCaseId: number): Record<string, any> {
    return iterations.reduce((map, iterationItem) => {
      if (iterationItem.iteration) {
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
   * Fetches result data for all test points within the given test data.
   */
  /**
   * Fetches result data for all test points within the given test data, sequentially to avoid context-related errors.
   */
  private async fetchAllResultData(testData: any[], projectName: string): Promise<any[]> {
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
          const resultData = await this.fetchResultData(projectName, testSuiteId, point);
          if (resultData !== null) {
            results.push(resultData);
          }
        }
      }
    }

    return results;
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
   * Fetches result Data data for a specific test point.
   */
  private async fetchResultData(projectName: string, testSuiteId: string, point: any) {
    const { lastRunId, lastResultId } = point;
    const resultData = await this.fetchResult(projectName, lastRunId.toString(), lastResultId.toString());

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
      ? {
          testCaseName: `${resultData.testCase.name} - ${resultData.testCase.id}`,
          testCaseId: resultData.testCase.id,
          testSuiteName: `${resultData.testSuite.name}`,
          testSuiteId,
          lastRunId,
          lastResultId,
          iteration,
          testCaseRevision: resultData.testCaseRevision,
          failureType: resultData.failureType,
          resolution: resultData.resolutionState,
          comment: resultData.comment,
          analysisAttachments: resultData.analysisAttachments,
        }
      : null;
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
            .filter((rel: any) => rel?.attributes?.name === 'Tests')
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

  private CreateAttachmentPathIndexMap(actionResults: any) {
    const attachmentPathToIndexMap: Map<string, number> = new Map();

    for (let i = 0; i < actionResults.length; i++) {
      const actionPath = actionResults[i].actionPath;
      attachmentPathToIndexMap.set(actionPath, actionResults[i].stepPosition);
    }
    return attachmentPathToIndexMap;
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
      const detailedStepResultsSummary = this.alignStepsWithIterations(testData, runResults);
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
      throw error;
    }
  }

  // private mapStepResultsForExecutionAppendix(detailedResults: any[]): any {
  //   return detailedResults?.length > 0
  //     ? detailedResults.map((result) => {
  //         return {
  //           testId: result.testId,
  //           testCaseRevision: result.testCaseRevision || undefined,
  //           stepNo: result.stepNo,
  //           stepIdentifier: result.stepIdentifier,
  //           stepStatus: result.stepStatus,
  //           stepComments: result.stepComments,
  //         };
  //       })
  //     : [];
  // }

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
}
