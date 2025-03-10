import { TFSServices } from '../helpers/tfs';
import { Helper, suiteData } from '../helpers/helper';
import { TestSteps, createRequirementRelation } from '../models/tfs-data';
import { TestCase } from '../models/tfs-data';
import * as xml2js from 'xml2js';
import logger from '../utils/logger';
import TestStepParserHelper from '../utils/testStepParserHelper';
const pLimit = require('p-limit');

export default class TestDataProvider {
  orgUrl: string = '';
  token: string = '';
  private testStepParserHelper: TestStepParserHelper;
  private cache = new Map<string, any>(); // Cache for API responses
  private limit = pLimit(10);

  constructor(orgUrl: string, token: string) {
    this.orgUrl = orgUrl;
    this.token = token;
    this.testStepParserHelper = new TestStepParserHelper(orgUrl, token);
  }

  private async fetchWithCache(url: string, ttlMs = 60000): Promise<any> {
    if (this.cache.has(url)) {
      const cached = this.cache.get(url);
      if (cached.timestamp + ttlMs > Date.now()) {
        return cached.data;
      }
    }

    try {
      const result = await TFSServices.getItemContent(url, this.token);

      this.cache.set(url, {
        data: result,
        timestamp: Date.now(),
      });

      return result;
    } catch (error: any) {
      logger.error(`Error fetching ${url}: ${error.message}`);
      throw error;
    }
  }

  async GetTestSuiteByTestCase(testCaseId: string): Promise<any> {
    let url = `${this.orgUrl}/_apis/testplan/suites?testCaseId=${testCaseId}`;
    return this.fetchWithCache(url);
  }

  async GetTestPlans(project: string): Promise<string> {
    let testPlanUrl: string = `${this.orgUrl}${project}/_apis/test/plans`;
    return this.fetchWithCache(testPlanUrl);
  }

  async GetTestSuites(project: string, planId: string): Promise<any> {
    let testsuitesUrl: string = this.orgUrl + project + '/_apis/test/Plans/' + planId + '/suites';
    try {
      return this.fetchWithCache(testsuitesUrl);
    } catch (e) {
      logger.error(`Failed to get test suites: ${e}`);
      return null;
    }
  }

  async GetTestSuitesForPlan(project: string, planid: string): Promise<any> {
    if (!project) {
      throw new Error('Project not selected');
    }
    if (!planid) {
      throw new Error('Plan not selected');
    }
    let url =
      this.orgUrl + '/' + project + '/_api/_testManagement/GetTestSuitesForPlan?__v=5&planId=' + planid;
    return this.fetchWithCache(url);
  }

  async GetTestSuitesByPlan(project: string, planId: string, recursive: boolean): Promise<any> {
    let suiteId = Number(planId) + 1;
    return this.GetTestSuiteById(project, planId, suiteId.toString(), recursive);
  }

  async GetTestSuiteById(project: string, planId: string, suiteId: string, recursive: boolean): Promise<any> {
    let testSuites = await this.GetTestSuitesForPlan(project, planId);
    Helper.suitList = [];
    let dataSuites: any = Helper.findSuitesRecursive(
      planId,
      this.orgUrl,
      project,
      testSuites.testSuites,
      suiteId,
      recursive
    );
    Helper.first = true;
    return dataSuites;
  }

  async GetTestCasesBySuites(
    project: string,
    planId: string,
    suiteId: string,
    recursive: boolean,
    includeRequirements: boolean,
    CustomerRequirementId: boolean,
    stepResultDetailsMap?: Map<string, any>
  ): Promise<any> {
    let testCasesList: Array<any> = new Array<any>();
    const requirementToTestCaseTraceMap: Map<string, string[]> = new Map();
    const testCaseToRequirementsTraceMap: Map<string, string[]> = new Map();
    // const startTime = performance.now();

    let suitesTestCasesList: Array<suiteData> = await this.GetTestSuiteById(
      project,
      planId,
      suiteId,
      recursive
    );

    // Create array of promises that each return their test cases
    const testCaseListPromises = suitesTestCasesList.map((suite) =>
      this.limit(async () => {
        try {
          const testCases = await this.GetTestCases(project, planId, suite.id);
          // const structureStartTime = performance.now();
          const testCasesWithSteps = await this.StructureTestCase(
            project,
            testCases,
            suite,
            includeRequirements,
            CustomerRequirementId,
            requirementToTestCaseTraceMap,
            testCaseToRequirementsTraceMap,
            stepResultDetailsMap
          );
          // logger.debug(
          //   `Performance: structured suite ${suite.id} in ${performance.now() - structureStartTime}ms`
          // );

          // Return the results instead of modifying shared array
          return testCasesWithSteps || [];
        } catch (error) {
          logger.error(`Error processing suite ${suite.id}: ${error}`);
          return []; // Return empty array on error
        }
      })
    );

    // Wait for all promises and only then combine the results
    const results = await Promise.all(testCaseListPromises);
    testCasesList = results.flat(); // Combine all results into a single array
    // logger.debug(`Performance: GetTestCasesBySuites completed in ${performance.now() - startTime}ms`);
    return { testCasesList, requirementToTestCaseTraceMap, testCaseToRequirementsTraceMap };
  }

  async StructureTestCase(
    project: string,
    testCases: any,
    suite: suiteData,
    includeRequirements: boolean,
    CustomerRequirementId: boolean,
    requirementToTestCaseTraceMap: Map<string, string[]>,
    testCaseToRequirementsTraceMap: Map<string, string[]>,
    stepResultDetailsMap?: Map<string, any>
  ): Promise<any[]> {
    let url = this.orgUrl + project + '/_workitems/edit/';
    let testCasesUrlList: any[] = [];
    logger.debug(`Trying to structure Test case for ${project} suite: ${suite.id}:${suite.name}`);
    try {
      if (!testCases || !testCases.value || testCases.count === 0) {
        logger.warn(`No test cases found for suite: ${suite.id}`);
        return [];
      }

      for (let i = 0; i < testCases.count; i++) {
        try {
          let stepDetailObject =
            stepResultDetailsMap?.get(testCases.value[i].testCase.id.toString()) || undefined;

          let newurl = !stepDetailObject?.testCaseRevision
            ? testCases.value[i].testCase.url + '?$expand=All'
            : `${testCases.value[i].testCase.url}/revisions/${stepDetailObject.testCaseRevision}?$expand=All`;
          let test: any = await this.fetchWithCache(newurl);
          let testCase: TestCase = new TestCase();

          testCase.title = test.fields['System.Title'];
          testCase.area = test.fields['System.AreaPath'];
          testCase.description = test.fields['System.Description'];
          testCase.url = url + test.id;
          //testCase.steps = test.fields["Microsoft.VSTS.TCM.Steps"];
          testCase.id = test.id;
          testCase.suit = suite.id;

          if (!stepDetailObject && test.fields['Microsoft.VSTS.TCM.Steps'] != null) {
            let steps = await this.testStepParserHelper.parseTestSteps(
              test.fields['Microsoft.VSTS.TCM.Steps'],
              new Map<number, number>()
            );
            testCase.steps = steps;
            //In case its already parsed during the STR
          } else if (stepDetailObject) {
            testCase.steps = stepDetailObject.stepList;
            testCase.caseEvidenceAttachments = stepDetailObject.caseEvidenceAttachments;
          }
          if (test.relations) {
            for (const relation of test.relations) {
              // Only proceed if the URL contains 'workItems'
              if (relation.url.includes('/workItems/')) {
                try {
                  let relatedItemContent: any = await this.fetchWithCache(relation.url);
                  // Check if the WorkItemType is "Requirement" before adding to relations
                  if (relatedItemContent.fields['System.WorkItemType'] === 'Requirement') {
                    const newRequirementRelation = this.createNewRequirement(
                      CustomerRequirementId,
                      relatedItemContent
                    );

                    const stringifiedTestCase = JSON.stringify({
                      id: testCase.id,
                      title: testCase.title,
                    });
                    const stringifiedRequirement = JSON.stringify(newRequirementRelation);

                    // Add the test case to the requirement-to-test-case trace map
                    this.addToMap(requirementToTestCaseTraceMap, stringifiedRequirement, stringifiedTestCase);

                    // Add the requirement to the test-case-to-requirements trace map
                    this.addToMap(
                      testCaseToRequirementsTraceMap,
                      stringifiedTestCase,
                      stringifiedRequirement
                    );

                    if (includeRequirements) {
                      testCase.relations.push(newRequirementRelation);
                    }
                  }
                } catch (fetchError) {
                  // Log error silently or handle as needed
                  console.error('Failed to fetch relation content', fetchError);
                  logger.error(`Failed to fetch relation content for URL ${relation.url}: ${fetchError}`);
                }
              }
            }
          }
          testCasesUrlList.push(testCase);
        } catch {
          const errorMsg = `ran into an issue while retrieving testCase ${testCases.value[i].testCase.id}`;
          logger.error(`Error: ${errorMsg}`);
          throw new Error(errorMsg);
        }
      }
    } catch (err: any) {
      logger.error(`Error: ${err.message} while trying to structure testCases for test suite ${suite.id}`);
    }

    return testCasesUrlList;
  }

  private createNewRequirement(CustomerRequirementId: boolean, relatedItemContent: any) {
    let customerId = undefined;
    // Check if CustomerRequirementId is true and set customerId
    if (CustomerRequirementId) {
      // Here we check for either of the two potential fields for customer ID
      customerId =
        relatedItemContent.fields['Custom.CustomerRequirementId'] ||
        relatedItemContent.fields['Custom.CustomerID'] ||
        relatedItemContent.fields['Elisra.CustomerRequirementId'] ||
        ' ';
    }
    const newRequirementRelation = createRequirementRelation(
      relatedItemContent.id,
      relatedItemContent.fields['System.Title'],
      customerId
    );
    return newRequirementRelation;
  }

  ParseSteps(steps: string) {
    let stepsLsist: Array<TestSteps> = new Array<TestSteps>();
    const start: string = ';P&gt;';
    const end: string = '&lt;/P';
    let totalString: String = steps;
    xml2js.parseString(steps, function (err, result) {
      if (err) console.log(err);
      if (result.steps.step != null)
        for (let i = 0; i < result.steps.step.length; i++) {
          let step: TestSteps = new TestSteps();
          try {
            if (result.steps.step[i].parameterizedString[0]._ != null)
              step.action = result.steps.step[i].parameterizedString[0]._;
          } catch (e) {
            logger.warn(`No test step action data to parse for testcase `);
          }
          try {
            if (result.steps.step[i].parameterizedString[1]._ != null)
              step.expected = result.steps.step[i].parameterizedString[1]._;
          } catch (e) {
            logger.warn(`No test step expected data to parse for testcase `);
          }
          stepsLsist.push(step);
        }
    });

    return stepsLsist;
  }

  async GetTestCases(project: string, planId: string, suiteId: string): Promise<any> {
    let testCaseUrl: string =
      this.orgUrl + project + '/_apis/test/Plans/' + planId + '/suites/' + suiteId + '/testcases/';
    let testCases: any = await this.fetchWithCache(testCaseUrl);
    logger.debug(`test cases for plan ${planId} and ${suiteId} were ${testCases ? 'found' : 'not found'}`);
    return testCases;
  }

  async GetTestPoint(project: string, planId: string, suiteId: string, testCaseId: string): Promise<any> {
    let testPointUrl: string = `${this.orgUrl}${project}/_apis/test/Plans/${planId}/Suites/${suiteId}/points?testCaseId=${testCaseId}`;
    return this.fetchWithCache(testPointUrl);
  }

  async CreateTestRun(
    projectName: string,
    testRunName: string,
    testPlanId: string,
    testPointId: string
  ): Promise<any> {
    try {
      logger.info(`Create test run op test point  ${testPointId} ,test planId : ${testPlanId}`);
      let Url = `${this.orgUrl}${projectName}/_apis/test/runs`;
      let data = {
        name: testRunName,
        plan: {
          id: testPlanId,
        },
        pointIds: [testPointId],
      };
      let res = await TFSServices.postRequest(Url, this.token, 'Post', data, null);
      return res;
    } catch (err) {
      logger.error(`Error : ${err}`);
      throw new Error(String(err));
    }
  }

  async UpdateTestRun(projectName: string, runId: string, state: string): Promise<any> {
    logger.info(`Update runId : ${runId} to state : ${state}`);
    let Url = `${this.orgUrl}${projectName}/_apis/test/Runs/${runId}?api-version=5.0`;
    let data = {
      state: state,
    };
    let res = await TFSServices.postRequest(Url, this.token, 'PATCH', data, null);
    return res;
  }

  async UpdateTestCase(projectName: string, runId: string, state: number): Promise<any> {
    let data: any;
    logger.info(`Update test case, runId : ${runId} to state : ${state}`);
    let Url = `${this.orgUrl}${projectName}/_apis/test/Runs/${runId}/results?api-version=5.0`;
    switch (state) {
      case 0:
        logger.info(`Reset test case to Active state `);
        data = [
          {
            id: 100000,
            outcome: '0',
          },
        ];
        break;
      case 1:
        logger.info(`Update test case to complite state `);
        data = [
          {
            id: 100000,
            state: 'Completed',
            outcome: '1',
          },
        ];
        break;
      case 2:
        logger.info(`Update test case to passed state `);
        data = [
          {
            id: 100000,
            state: 'Completed',
            outcome: '2',
          },
        ];
        break;
      case 3:
        logger.info(`Update test case to failed state `);
        data = [
          {
            id: 100000,
            state: 'Completed',
            outcome: '3',
          },
        ];
        break;
    }
    let res = await TFSServices.postRequest(Url, this.token, 'PATCH', data, null);
    return res;
  }

  async UploadTestAttachment(
    runID: string,
    projectName: string,
    stream: any,
    fileName: string,
    comment: string,
    attachmentType: string
  ): Promise<any> {
    logger.info(`Upload attachment to test run : ${runID}`);
    let Url = `${this.orgUrl}${projectName}/_apis/test/Runs/${runID}/attachments?api-version=5.0-preview.1`;
    let data = {
      stream: stream,
      fileName: fileName,
      comment: comment,
      attachmentType: attachmentType,
    };
    let res = await TFSServices.postRequest(Url, this.token, 'Post', data, null);
    return res;
  }

  async GetTestRunById(projectName: string, runId: string): Promise<any> {
    logger.info(`getting test run id: ${runId}`);
    let url = `${this.orgUrl}${projectName}/_apis/test/Runs/${runId}`;
    let res = await TFSServices.getItemContent(url, this.token, 'get', null, null);
    return res;
  }

  async GetTestPointByTestCaseId(projectName: string, testCaseId: string): Promise<any> {
    logger.info(`get test points at project ${projectName} , of testCaseId : ${testCaseId}`);
    let url = `${this.orgUrl}${projectName}/_apis/test/points`;
    let data = {
      PointsFilter: {
        TestcaseIds: [testCaseId],
      },
    };
    let res = await TFSServices.postRequest(url, this.token, 'Post', data, null);
    return res;
  }

  private addToMap = (map: Map<string, string[]>, key: string, value: string) => {
    if (!map.has(key)) {
      map.set(key, []);
    }
    map.get(key)?.push(value);
  };

  /**
   * Clears the cache to free memory
   */
  public clearCache(): void {
    this.cache.clear();
    logger.debug('Cache cleared');
  }
}
