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
  ): Promise<Array<any>> {
    const baseUrl = this.orgUrl + project + '/_workitems/edit/';
    const testCasesUrlList: Array<any> = [];

    logger.debug(`Structuring test cases for ${project} suite: ${suite.id}:${suite.name}`);

    try {
      if (!testCases || !testCases.value || testCases.count === 0) {
        logger.warn(`No test cases found for suite: ${suite.id}`);
        return [];
      }

      // Step 1: Prepare all test case fetch requests
      const testCaseRequests = testCases.value.map((item: any) => {
        const testCaseId = item.testCase.id.toString();
        const stepDetailObject = stepResultDetailsMap?.get(testCaseId);

        const url = !stepDetailObject?.testCaseRevision
          ? item.testCase.url + '?$expand=All'
          : `${item.testCase.url}/revisions/${stepDetailObject.testCaseRevision}?$expand=All`;

        return {
          url,
          testCaseId,
          stepDetailObject,
        };
      });

      // Step 2: Fetch all test cases in parallel using limit to control concurrency
      logger.debug(`Fetching ${testCaseRequests.length} test cases concurrently`);
      const testCaseResults = await Promise.all(
        testCaseRequests.map((request: any) =>
          this.limit(async () => {
            try {
              const test = await this.fetchWithCache(request.url);
              return {
                test,
                testCaseId: request.testCaseId,
                stepDetailObject: request.stepDetailObject,
              };
            } catch (error) {
              logger.error(`Error fetching test case ${request.testCaseId}: ${error}`);
              return null;
            }
          })
        )
      );

      // Step 3: Process test cases and collect relation URLs
      const validResults = testCaseResults.filter((result) => result !== null);
      const relationRequests: Array<{ url: string; testCaseIndex: number }> = [];

      const testCaseObjects = await Promise.all(
        validResults.map(async (result, index) => {
          try {
            const test = result!.test;
            const testCase = new TestCase();

            // Build test case object
            testCase.title = test.fields['System.Title'];
            testCase.area = test.fields['System.AreaPath'];
            testCase.description = test.fields['System.Description'];
            testCase.url = baseUrl + test.id;
            testCase.id = test.id;
            testCase.suit = suite.id;

            // Handle steps
            if (!result!.stepDetailObject && test.fields['Microsoft.VSTS.TCM.Steps'] != null) {
              testCase.steps = await this.testStepParserHelper.parseTestSteps(
                test.fields['Microsoft.VSTS.TCM.Steps'],
                new Map<number, number>()
              );
            } else if (result!.stepDetailObject) {
              testCase.steps = result!.stepDetailObject.stepList;
              testCase.caseEvidenceAttachments = result!.stepDetailObject.caseEvidenceAttachments;
            }

            // Collect relation URLs for batch processing
            if (test.relations) {
              test.relations.forEach((relation: any) => {
                if (relation.url.includes('/workItems/')) {
                  relationRequests.push({
                    url: relation.url,
                    testCaseIndex: index,
                  });
                }
              });
            }

            return testCase;
          } catch (error) {
            logger.error(`Error processing test case ${result!.testCaseId}: ${error}`);
            return null;
          }
        })
      );

      // Filter out any errors during test case processing
      const validTestCases = testCaseObjects.filter((tc) => tc !== null) as TestCase[];

      // Step 4: Fetch all relations concurrently
      if (relationRequests.length > 0) {
        logger.debug(`Fetching ${relationRequests.length} relations concurrently`);
        const relationResults = await Promise.all(
          relationRequests.map((request) =>
            this.limit(async () => {
              try {
                const relatedItemContent = await this.fetchWithCache(request.url);
                return {
                  content: relatedItemContent,
                  testCaseIndex: request.testCaseIndex,
                };
              } catch (error) {
                logger.error(`Failed to fetch relation: ${error}`);
                return null;
              }
            })
          )
        );

        // Step 5: Process relations and update test cases
        relationResults
          .filter((result) => result !== null)
          .forEach((result) => {
            const relatedItemContent = result!.content;
            const testCase = validTestCases[result!.testCaseIndex];

            // Only process requirement relations
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

              // Update trace maps
              this.addToMap(requirementToTestCaseTraceMap, stringifiedRequirement, stringifiedTestCase);
              this.addToMap(testCaseToRequirementsTraceMap, stringifiedTestCase, stringifiedRequirement);

              // Add to test case relations if needed
              if (includeRequirements) {
                testCase.relations.push(newRequirementRelation);
              }
            }
          });
      }

      // Add all valid test cases to the result list
      testCasesUrlList.push(...validTestCases);
    } catch (err: any) {
      logger.error(`Error: ${err.message} while structuring test cases for suite ${suite.id}`);
    }

    logger.info(
      `StructureTestCase for suite ${suite.id} completed with ${testCasesUrlList.length} test cases`
    );
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
