import { TFSServices } from "../helpers/tfs";
import { Workitem } from "../models/tfs-data";
import { Helper, suiteData, Links, Trace, Relations } from "../helpers/helper";
import { Query, TestSteps } from "../models/tfs-data";
import { QueryType } from "../models/tfs-data";
import { QueryAllTypes } from "../models/tfs-data";
import { Column } from "../models/tfs-data";
import { value } from "../models/tfs-data";
import { TestCase } from "../models/tfs-data";
import * as xml2js from "xml2js";

import logger from "../utils/logger";

export default class TestDataProvider {
  orgUrl: string = "";
  token: string = "";
  
  constructor(orgUrl: string, token: string) {
    this.orgUrl = orgUrl;
    this.token = token;
  }

  async GetTestSuiteByTestCase(
    testCaseId: string
  ): Promise<any> {
    let url = `${this.orgUrl}/_apis/testplan/suites?testCaseId=${testCaseId}`;
    let testCaseData = await TFSServices.getItemContent(url, this.token);
    return testCaseData;
  }

  //get all test plans in the project
  async GetTestPlans(
    project: string
  ): Promise<string> {
    let testPlanUrl: string = `${this.orgUrl}${project}/_apis/test/plans`;
    return TFSServices.getItemContent(testPlanUrl,this.token);
  }

  // get all test suits in projct test plan
  async GetTestSuites(
    project: string,
    planId: string
  ): Promise<any> {
    let testsuitesUrl: string =
      this.orgUrl + project + "/_apis/test/Plans/" + planId + "/suites";
    try {
      let testSuites = await TFSServices.getItemContent(
        testsuitesUrl,
        this.token
      );
      return testSuites;
    } catch (e) {}
  }

  async GetTestSuitesForPlan(
    project: string,
    planid: string
  ): Promise<any> {
    let url =
      this.orgUrl +
      "/" +
      project +
      "/_api/_testManagement/GetTestSuitesForPlan?__v=5&planId=" +
      planid;
    let suites = await TFSServices.getItemContent(url, this.token);
    return suites;
  }

  async GetTestSuitesByPlan(
    project: string,
    planId: string,
    recursive: boolean
  ): Promise<any> {
    let suiteId = Number(planId) + 1;
    let suites = await this.GetTestSuiteById(
      project,
      planId,
      suiteId.toString(),
      recursive
    );
    return suites;
  }
  //gets all testsuits recorsivly under test suite

  async GetTestSuiteById(
    project: string,
    planId: string,
    suiteId: string,
    recursive: boolean
  ): Promise<any> {
    let testSuites = await this.GetTestSuitesForPlan(project, planId);
    // GetTestSuites(project, planId);
    Helper.suitList = [];
    Helper.level = 1;
    let dataSuites: any = Helper.findSuitesRecursive(
      planId,
      this.orgUrl,
      project,
      testSuites.testSuites,
      suiteId,
      recursive
    );
    Helper.first = true;
    // let levledSuites: any = Helper.buildSuiteslevel(dataSuites);

    return dataSuites;
  }
  //gets all testcase under test suite acording to recursive flag

  async GetTestCasesBySuites(
    project: string,
    planId: string,
    suiteId: string,
    recursiv: boolean
  ): Promise<Array<any>> {
    let testCasesList: Array<any> = new Array<any>();
    let suitesTestCasesList: Array<suiteData> = await this.GetTestSuiteById(
      project,
      planId,
      suiteId,
      recursiv
    );
    for (let i = 0; i < suitesTestCasesList.length; i++) {
      let testCases: any = await this.GetTestCases(
        project,
        planId,
        suitesTestCasesList[i].id
      );
      let testCseseWithSteps: any = await this.StructureTestCase(
        project,
        testCases,
        suitesTestCasesList[i]
      );
      if (testCseseWithSteps.length > 0)
        testCasesList = [...testCasesList, ...testCseseWithSteps];
    }
    return testCasesList;
  }

  async StructureTestCase(
    project: string,
    testCases: any,
    suite: suiteData
  ): Promise<Array<any>> {
    let url = this.orgUrl + project + "/_workitems/edit/";
    let testCasesUrlList: Array<any> = new Array<any>();
    //let tesrCase:TestCase;
    for (let i = 0; i < testCases.count; i++) {
      try{
      let test: any = await TFSServices.getItemContent(
        testCases.value[i].testCase.url,
        this.token
      );
      let testCase: TestCase = new TestCase();
      testCase.title = test.fields["System.Title"];
      testCase.area = test.fields["System.AreaPath"];
      testCase.description = test.fields["System.Description"];
      testCase.url = url + test.id;
      //testCase.steps = test.fields["Microsoft.VSTS.TCM.Steps"];
      testCase.id = test.id;
      testCase.suit = suite.id;
      if (test.fields["Microsoft.VSTS.TCM.Steps"] != null) {
        let steps: Array<TestSteps> = this.ParseSteps(
          test.fields["Microsoft.VSTS.TCM.Steps"]
        );
        testCase.steps = steps;
      }
      testCasesUrlList.push(testCase);
    }
    catch{
      logger.error(`ran into an issue while retriving testCase ${testCases.value[i].testCase.id}`);
    }
  }

    return testCasesUrlList;
  }

  ParseSteps(
    steps: string
  ) {
    let stepsLsist: Array<TestSteps> = new Array<TestSteps>();
    const start: string = ";P&gt;";
    const end: string = "&lt;/P";
    let totalString: String = steps;
    xml2js.parseString(steps, function (err, result) {
      if (err) console.log(err);
      if (result.steps.step != null)
        for (let i = 0; i < result.steps.step.length; i++) {
          let step: TestSteps = new TestSteps();
          try {
            if (result.steps.step[i].parameterizedString[0]._ != null)
              step.action = result.steps.step[
                i
              ].parameterizedString[0]._
          } catch (e) {
            logger.warn(`No test step action data to parse for testcase `);
          }
          try {
            if (result.steps.step[i].parameterizedString[1]._ != null)
              step.expected = result.steps.step[
                i
              ].parameterizedString[1]._
          } catch (e) {
            logger.warn(`No test step expected data to parse for testcase `);
          }
          stepsLsist.push(step);
        }
    });

    return stepsLsist;
  }

  async GetTestCases(
    project: string,
    planId: string,
    suiteId: string
  ): Promise<any> {
    let testCaseUrl: string =
      this.orgUrl +
      project +
      "/_apis/test/Plans/" +
      planId +
      "/suites/" +
      suiteId +
      "/testcases/";
    let testCases: any = await TFSServices.getItemContent(
      testCaseUrl,
      this.token
    );
    return testCases;
  }

  //gets all test point in a test case
  async GetTestPoint(
    project: string,
    planId: string,
    suiteId: string,
    testCaseId: string
  ): Promise<any> {
    let testPointUrl: string =
    `${this.orgUrl}${project}/_apis/test/Plans/${planId}/Suites/${suiteId}/points?testCaseId=${testCaseId}`;
    let testPoints: any = await TFSServices.getItemContent(
      testPointUrl,
      this.token
    );
    return testPoints;
  }

  async CreateTestRun(
    projectName: string,
    testRunName: string,
    testPlanId: string,
    testPointId: string
  ): Promise<any> {
    try {
      logger.info(
        `Create test run op test point  ${testPointId} ,test planId : ${testPlanId}`
      );
      let Url = `${this.orgUrl}${projectName}/_apis/test/runs`;
      let data = {
        name: testRunName,
        plan: {
          id: testPlanId,
        },
        pointIds: [testPointId],
      };
      let res = await TFSServices.postRequest(
        Url,
        this.token,
        "Post",
        data,
        null
      );
      return res;
    } catch (err) {
      logger.error(`Error : ${err}`);
      throw new Error(String(err));
    }
  }

  async UpdateTestRun(
    projectName: string,
    runId: string,
    state: string
  ): Promise<any> {
    logger.info(`Update runId : ${runId} to state : ${state}`);
    let Url = `${this.orgUrl}${projectName}/_apis/test/Runs/${runId}?api-version=5.0`;
    let data = {
      state: state,
    };
    let res = await TFSServices.postRequest(
      Url,
      this.token,
      "PATCH",
      data,
      null
    );
    return res;
  }

  async UpdateTestCase(
    projectName: string,
    runId: string,
    state: number
  ): Promise<any> {
    let data: any;
    logger.info(`Update test case, runId : ${runId} to state : ${state}`);
    let Url = `${this.orgUrl}${projectName}/_apis/test/Runs/${runId}/results?api-version=5.0`;
    switch (state) {
      case 0:
        logger.info(`Reset test case to Active state `);
        data = [
          {
            id: 100000,
            outcome: "0",
          },
        ];
        break;
      case 1:
        logger.info(`Update test case to complite state `);
        data = [
          {
            id: 100000,
            state: "Completed",
            outcome: "1",
          },
        ];
        break;
      case 2:
        logger.info(`Update test case to passed state `);
        data = [
          {
            id: 100000,
            state: "Completed",
            outcome: "2",
          },
        ];
        break;
      case 3:
        logger.info(`Update test case to failed state `);
        data = [
          {
            id: 100000,
            state: "Completed",
            outcome: "3",
          },
        ];
        break;
    }
    let res = await TFSServices.postRequest(
      Url,
      this.token,
      "PATCH",
      data,
      null
    );
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
    let res = await TFSServices.postRequest(
      Url,
      this.token,
      "Post",
      data,
      null
    );
    return res;
  }

  async GetTestRunById(
    projectName: string, runId: string
  ): Promise<any> {
    logger.info(`getting test run id: ${runId}`);
    let url = `${this.orgUrl}${projectName}/_apis/test/Runs/${runId}`;
    let res = await TFSServices.getItemContent(
      url,
      this.token,
      "get",
      null,
      null
    );
    return res;
  }

  async GetTestPointByTestCaseId(
    projectName: string,
    testCaseId: string
  ): Promise<any> {
    logger.info(
      `get test points at project ${projectName} , of testCaseId : ${testCaseId}`
    );
    let url = `${this.orgUrl}${projectName}/_apis/test/points`;
    let data = {
      PointsFilter: {
        TestcaseIds: [testCaseId],
      },
    };
    let res = await TFSServices.postRequest(
      url,
      this.token,
      "Post",
      data,
      null
    );
    return res;
  }
}
