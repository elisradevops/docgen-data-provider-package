import { TFSServices } from '../helpers/tfs';
import { TestSteps } from '../models/tfs-data';
import * as xml2js from 'xml2js';
import logger from '../utils/logger';

export default class ResultDataProvider {
  orgUrl: string = '';
  token: string = '';

  constructor(orgUrl: string, token: string) {
    this.orgUrl = orgUrl;
    this.token = token;
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
      outcome: testPoint.results?.outcome,
      lastRunId: testPoint.results?.lastTestRunId,
      lastResultId: testPoint.results?.lastResultId,
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
  private async fetchIterations(projectName: string, runId: string, resultId: string): Promise<any[]> {
    const url = `${this.orgUrl}${projectName}/_apis/test/runs/${runId}/Results/${resultId}/iterations`;
    const { value: iterations } = await TFSServices.getItemContent(url, this.token);
    return iterations;
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
        return 'Unspecified';
    }
  }

  /**
   * Parses test steps from XML format into a structured array.
   */
  private parseTestSteps(xmlSteps: string): TestSteps[] {
    const stepsList: TestSteps[] = [];
    xml2js.parseString(xmlSteps, { explicitArray: false }, (err, result) => {
      if (err) {
        logger.warn('Failed to parse XML test steps.');
        return;
      }

      const stepsArray = Array.isArray(result.steps?.step) ? result.steps.step : [result.steps?.step];

      stepsArray.forEach((stepObj: any) => {
        const step = new TestSteps();
        step.action = stepObj.parameterizedString?.[0]?._ || '';
        step.expected = stepObj.parameterizedString?.[1]?._ || '';
        stepsList.push(step);
      });
    });
    return stepsList;
  }

  /**
   * Aligns test steps with their corresponding iterations.
   */
  private alignStepsWithIterations(testData: any[], iterations: any[]): any[] {
    const detailedResults: any[] = [];
    const iterationsMap = this.createIterationsMap(iterations);

    for (const testItem of testData) {
      for (const point of testItem.testPointsItems) {
        const testCase = testItem.testCasesItems.find((tc: any) => tc.workItem.id === point.testCaseId);
        if (!testCase) continue;

        const steps = this.parseTestSteps(testCase.workItem.workItemFields[0]['Microsoft.VSTS.TCM.Steps']);
        const iterationKey = `${point.lastRunId}-${point.lastResultId}`;
        const iteration = iterationsMap[iterationKey]?.iteration;

        if (!iteration) continue;

        for (const actionResult of iteration.actionResults) {
          const stepIndex = parseInt(actionResult.stepIdentifier, 10) - 2;
          if (!steps[stepIndex]) continue;

          // const actionPath = actionResult.actionPath;

          // const testStepAttachments = iteration.attachments.find(
          //   (attachment: any) => attachment.actionPath === actionPath
          // );

          detailedResults.push({
            testId: point.testCaseId,
            testName: point.testCaseName,
            stepNo: stepIndex + 1,
            stepAction: steps[stepIndex].action,
            stepExpected: steps[stepIndex].expected,
            stepStatus: actionResult.outcome !== 'Unspecified' ? actionResult.outcome : '',
            stepComments: actionResult.errorMessage || '',
          });
        }
      }
    }
    return detailedResults;
  }

  /**
   * Creates a mapping of iterations by their unique keys.
   */
  private createIterationsMap(iterations: any[]): Record<string, any> {
    return iterations.reduce(
      (map, iterationItem) => {
        if (iterationItem.iteration) {
          const key = `${iterationItem.lastRunId}-${iterationItem.lastResultId}`;
          map[key] = iterationItem;
        }
        return map;
      },
      {} as Record<string, any>
    );
  }

  /**
   * Fetches test data for all suites, including test points and test cases.
   */
  private async fetchTestData(suites: any[], projectName: string, testPlanId: string): Promise<any[]> {
    return Promise.all(
      suites.map((suite) =>
        Promise.all([
          this.fetchTestPoints(projectName, testPlanId, suite.testSuiteId),
          this.fetchTestCasesBySuiteId(projectName, testPlanId, suite.testSuiteId),
        ])
          .then(([testPointsItems, testCasesItems]) => ({ ...suite, testPointsItems, testCasesItems }))
          .catch((error: any) => {
            logger.error(`Error occurred for suite ${suite.testSuiteId}: ${error.message}`);
            return suite;
          })
      )
    );
  }

  /**
   * Fetches iterations for all test points within the given test data.
   */
  private async fetchAllIterations(testData: any[], projectName: string): Promise<any[]> {
    const pointsToFetch = testData
      .filter((item) => item.testPointsItems && item.testPointsItems.length > 0)
      .flatMap((item) => {
        const { testSuiteId, testPointsItems } = item;
        const validPoints = testPointsItems.filter((point: any) => point.lastRunId && point.lastResultId);
        return validPoints.map((point: any) => this.fetchIterationData(projectName, testSuiteId, point));
      });

    return Promise.all(pointsToFetch);
  }

  /**
   * Fetches iteration data for a specific test point.
   */
  private async fetchIterationData(projectName: string, testSuiteId: string, point: any): Promise<any> {
    const { lastRunId, lastResultId } = point;
    const iterations = await this.fetchIterations(projectName, lastRunId.toString(), lastResultId.toString());
    return {
      testSuiteId,
      lastRunId,
      lastResultId,
      iteration: iterations.length > 0 ? iterations[iterations.length - 1] : undefined,
    };
  }

  /**
   * Combines the results of test group result summary, test results summary, and detailed results summary into a single key-value pair array.
   */
  public async getCombinedResultsSummary(
    testPlanId: string,
    projectName: string,
    selectedSuiteIds?: number[],
    addConfiguration: boolean = false,
    isHierarchyGroupName: boolean = false
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
        this.fetchTestPoints(projectName, testPlanId, suite.testSuiteId)
          .then((testPointsItems) => ({ ...suite, testPointsItems }))
          .catch((error: any) => {
            logger.error(`Error occurred for suite ${suite.testSuiteId}: ${error.message}`);
            return { ...suite, testPointsItems: [] };
          })
      );
      const testPoints = await Promise.all(testPointsPromises);

      // 1. Calculate Test Group Result Summary
      const summarizedResults = testPoints.map((testPoint) => {
        const groupResultSummary = this.calculateGroupResultSummary(testPoint.testPointsItems || []);
        return { ...testPoint, groupResultSummary };
      });

      const totalSummary = this.calculateTotalSummary(summarizedResults);
      const testGroupArray = summarizedResults.map((item) => ({
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
        this.formatTestResult(testPoint, addConfiguration)
      );

      // Add test results summary to combined results
      combinedResults.push({
        contentControl: 'test-result-summary-content-control',
        data: testResultsSummary,
        skin: 'test-result-table',
      });

      // 3. Calculate Detailed Results Summary
      const testData = await this.fetchTestData(suites, projectName, testPlanId);
      const iterations = await this.fetchAllIterations(testData, projectName);
      const detailedResultsSummary = this.alignStepsWithIterations(testData, iterations);

      // Add detailed results summary to combined results
      combinedResults.push({
        contentControl: 'detailed-test-result-content-control',
        data: detailedResultsSummary,
        skin: 'detailed-test-result-table',
      });

      return combinedResults;
    } catch (error: any) {
      logger.error(`Error during getCombinedResultsSummary: ${error.message}`);
      return combinedResults; // Return whatever is computed even in case of error
    }
  }

  /**
   * Calculates a summary of test group results.
   */
  private calculateGroupResultSummary(testPointsItems: any[]): any {
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
  private calculateTotalSummary(results: any[]): any {
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
  private formatTestResult(testPoint: any, addConfiguration: boolean): any {
    const formattedResult: any = {
      testGroupName: testPoint.testGroupName,
      testId: testPoint.testCaseId,
      testName: testPoint.testCaseName,
      runStatus: this.convertRunStatus(testPoint.outcome),
    };

    if (addConfiguration) {
      formattedResult.configuration = testPoint.configurationName;
    }

    return formattedResult;
  }
}
