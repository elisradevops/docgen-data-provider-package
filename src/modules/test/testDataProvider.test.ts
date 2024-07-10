import DgDataProviderAzureDevOps from '../..';

require('dotenv').config();
jest.setTimeout(600000);

const orgUrl = process.env.ORG_URL;
const token = process.env.PAT;
const dgDataProviderAzureDevOps = new DgDataProviderAzureDevOps(orgUrl, token);

describe('Test module - tests', () => {
  test('should return test plans', async () => {
    let TestDataProvider = await dgDataProviderAzureDevOps.getTestDataProvider();
    let json: any = await TestDataProvider.GetTestPlans('tests');
    expect(json.count).toBeGreaterThanOrEqual(1);
  });
  test('should return test suites by plan', async () => {
    //not working yet
    let TestDataProvider = await dgDataProviderAzureDevOps.getTestDataProvider();
    let testSuites = await TestDataProvider.GetTestSuitesByPlan('tests', '540', true);
    expect(testSuites[0].name).toBe('TestSuite');
  });
  test('should return list of test cases', async () => {
    let TestDataProvider = await dgDataProviderAzureDevOps.getTestDataProvider();
    let attachList: any = await TestDataProvider.GetTestCasesBySuites(
      'tests',
      '545',
      '546',
      true,
      true,
      true,
      true,
      true
    );
    expect(attachList.length > 0).toBeDefined();
  });
  test('should use Helper.findSuitesRecursive twice after restarting static value of Helper.first=True ', async () => {
    let TestDataProvider = await dgDataProviderAzureDevOps.getTestDataProvider();
    let suitesByPlan = await TestDataProvider.GetTestSuitesByPlan('tests', '545', true);
    expect(suitesByPlan.length > 0).toBeDefined();
  });
  test.skip('should return list of test cases - stress test - big testplan 1400 cases', async () => {
    jest.setTimeout(1000000);
    let TestDataProvider = await dgDataProviderAzureDevOps.getTestDataProvider();
    let attachList: any = await TestDataProvider.GetTestCasesBySuites(
      'tests',
      '540',
      '549',
      true,
      true,
      true,
      true,
      true
    );
    expect(attachList.length > 1000).toBeDefined(); //not enough test cases for stress test
  });
  test('should return test cases by suite', async () => {
    let TestDataProvider = await dgDataProviderAzureDevOps.getTestDataProvider();
    let json = await TestDataProvider.GetTestCases('tests', '540', '541');
    expect(json.count).toBeGreaterThan(0);
  });
  test('should return test points by testcase', async () => {
    let TestDataProvider = await dgDataProviderAzureDevOps.getTestDataProvider();
    let json = await TestDataProvider.GetTestPoint('tests', '540', '541', '542');
    expect(json.count).toBeGreaterThan(0);
  });
  test('should return test runs by testcaseid', async () => {
    let TestDataProvider = await dgDataProviderAzureDevOps.getTestDataProvider();
    let json = await TestDataProvider.GetTestRunById('tests', '1000120');
    expect(json.id).toBe(1000120);
  });
  test('should create run test according test pointId and return OK(200) as response ', async () => {
    let TestDataProvider = await dgDataProviderAzureDevOps.getTestDataProvider();
    let result: any = await TestDataProvider.CreateTestRun('tests', 'testrun', '540', '3');
    expect(result.status).toBe(200);
  });
  test('should Update runId state and return OK(200) as response ', async () => {
    let TestDataProvider = await dgDataProviderAzureDevOps.getTestDataProvider();
    let result: any = await TestDataProvider.UpdateTestRun(
      'tests',
      '1000124', //runId
      'NeedsInvestigation' //Unspecified ,NotStarted, InProgress, Completed, Waiting, Aborted, NeedsInvestigation (State)
    );
    expect(result.status).toBe(200);
  });
  test('should Update test case state and return OK(200) as response ', async () => {
    let TestDataProvider = await dgDataProviderAzureDevOps.getTestDataProvider();
    let result: any = await TestDataProvider.UpdateTestCase(
      'tests',
      '1000120',
      2 //0-reset , 1-complite , 2-passed , 3-failed (State)
    );
    expect(result.status).toBe(200);
  });
  test('should Upload attachment for test run and return OK(200) as response ', async () => {
    let TestDataProvider = await dgDataProviderAzureDevOps.getTestDataProvider();
    let data = 'This is test line of data';
    let buff = new Buffer(data);
    let base64data = buff.toString('base64');
    let result: any = await TestDataProvider.UploadTestAttachment(
      '1000120', //runID
      'tests',
      base64data, //stream
      'testAttachment2.json', //fileName
      'Test attachment upload', //comment
      'GeneralAttachment' //attachmentType
    );
    expect(result.status).toBe(200);
  });
  test('should Get all test case data', async () => {
    let TestDataProvider = await dgDataProviderAzureDevOps.getTestDataProvider();
    let result: any = await TestDataProvider.GetTestSuiteByTestCase('544');
    expect(result).toBeDefined;
  });
  test('should Get test points by test case id', async () => {
    let TestDataProvider = await dgDataProviderAzureDevOps.getTestDataProvider();
    let result: any = await TestDataProvider.GetTestPointByTestCaseId(
      'tests',
      '544' //testCaseId
    );
    expect(result).toBeDefined;
  });
});
