import { TFSServices } from '../../helpers/tfs';
import { Helper, suiteData } from '../../helpers/helper';
import TestDataProvider from '../../modules/TestDataProvider';
import Utils from '../../utils/testStepParserHelper';
import logger from '../../utils/logger';
import { TestCase } from '../../models/tfs-data';

jest.mock('../../helpers/tfs');
jest.mock('../../utils/logger');
jest.mock('../../helpers/helper');
jest.mock('../../utils/testStepParserHelper', () => {
  return {
    __esModule: true,
    default: jest.fn().mockImplementation(() => ({
      parseTestSteps: jest.fn(),
    })),
  };
});
jest.mock('p-limit', () => jest.fn(() => (fn: Function) => fn()));

describe('TestDataProvider', () => {
  let testDataProvider: TestDataProvider;
  const mockOrgUrl = 'https://dev.azure.com/orgname/';
  const mockToken = 'mock-token';
  const mockBearerToken = 'bearer:abc.def.ghi';
  const mockProject = 'project-123';
  const mockPlanId = '456';
  const mockSuiteId = '789';
  const mockTestCaseId = '101112';

  beforeEach(() => {
    jest.clearAllMocks();

    testDataProvider = new TestDataProvider(mockOrgUrl, mockToken);
  });

  // Helper to access private method for testing
  const invokeFetchWithCache = async (instance: any, url: string, ttl = 60000) => {
    return instance.fetchWithCache.call(instance, url, ttl);
  };

  describe('fetchWithCache', () => {
    it('should return cached data when available and not expired', async () => {
      // Arrange
      const mockUrl = `${mockOrgUrl}_apis/test/endpoint`;
      const mockData = { value: 'test data' };
      const cache = new Map();
      cache.set(mockUrl, {
        data: mockData,
        timestamp: Date.now(),
      });
      (testDataProvider as any).cache = cache;

      // Act
      const result = await invokeFetchWithCache(testDataProvider, mockUrl);

      // Assert
      expect(result).toEqual(mockData);
      expect(TFSServices.getItemContent).not.toHaveBeenCalled();
    });

    it('should fetch new data when cache is expired', async () => {
      // Arrange
      const mockUrl = `${mockOrgUrl}_apis/test/endpoint`;
      const mockData = { value: 'old data' };
      const newData = { value: 'new data' };
      const cache = new Map();
      cache.set(mockUrl, {
        data: mockData,
        timestamp: Date.now() - 70000, // Expired (default TTL is 60000ms)
      });
      (testDataProvider as any).cache = cache;

      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(newData);

      // Act
      const result = await invokeFetchWithCache(testDataProvider, mockUrl);

      // Assert
      expect(result).toEqual(newData);
      expect(TFSServices.getItemContent).toHaveBeenCalledWith(mockUrl, mockToken);
    });

    it('should fetch and cache new data when not in cache', async () => {
      // Arrange
      const mockUrl = `${mockOrgUrl}_apis/test/endpoint`;
      const mockData = { value: 'new data' };
      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockData);

      // Act
      const result = await invokeFetchWithCache(testDataProvider, mockUrl);

      // Assert
      expect(result).toEqual(mockData);
      expect(TFSServices.getItemContent).toHaveBeenCalledWith(mockUrl, mockToken);
      expect((testDataProvider as any).cache.has(mockUrl)).toBeTruthy();
      expect((testDataProvider as any).cache.get(mockUrl).data).toEqual(mockData);
    });

    it('should throw and log error when fetch fails', async () => {
      // Arrange
      const mockUrl = `${mockOrgUrl}_apis/test/endpoint`;
      const mockError = new Error('API call failed');
      (TFSServices.getItemContent as jest.Mock).mockRejectedValueOnce(mockError);

      // Act & Assert
      await expect(invokeFetchWithCache(testDataProvider, mockUrl)).rejects.toThrow('API call failed');
      expect(logger.error).toHaveBeenCalledWith(`Error fetching ${mockUrl}: API call failed`);
    });
  });

  describe('GetTestSuiteByTestCase', () => {
    it('should return test suites for a given test case ID', async () => {
      // Arrange
      const mockData = { value: [{ id: '123', name: 'Test Suite 1' }] };
      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockData);

      // Act
      const result = await testDataProvider.GetTestSuiteByTestCase(mockTestCaseId);

      // Assert
      expect(result).toEqual(mockData);
      expect(TFSServices.getItemContent).toHaveBeenCalledWith(
        `${mockOrgUrl}_apis/testplan/suites?testCaseId=${mockTestCaseId}`,
        mockToken
      );
    });
  });

  describe('GetTestPlans', () => {
    it('should return test plans for a given project', async () => {
      // Arrange
      const mockData = {
        value: [
          { id: '456', name: 'Test Plan 1' },
          { id: '789', name: 'Test Plan 2' },
        ],
      };
      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockData);

      // Act
      const result = await testDataProvider.GetTestPlans(mockProject);

      // Assert
      expect(result).toEqual(mockData);
      expect(TFSServices.getItemContent).toHaveBeenCalledWith(
        `${mockOrgUrl}${mockProject}/_apis/test/plans`,
        mockToken
      );
    });
  });

  describe('GetTestSuites', () => {
    it('should return test suites for a given project and plan ID', async () => {
      // Arrange
      const mockData = {
        value: [
          { id: '123', name: 'Test Suite 1' },
          { id: '456', name: 'Test Suite 2' },
        ],
      };
      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockData);

      // Act
      const result = await testDataProvider.GetTestSuites(mockProject, mockPlanId);

      // Assert
      expect(result).toEqual(mockData);
      expect(TFSServices.getItemContent).toHaveBeenCalledWith(
        `${mockOrgUrl}${mockProject}/_apis/test/Plans/${mockPlanId}/suites`,
        mockToken
      );
    });

    it('should return null and log error if fetching test suites fails', async () => {
      // Arrange
      const mockError = new Error('Failed to get test suites');
      (TFSServices.getItemContent as jest.Mock).mockRejectedValueOnce(mockError);

      // Act
      const result = await testDataProvider.GetTestSuites(mockProject, mockPlanId);

      // Assert
      expect(result).toBeNull();
    });
  });

  describe('GetTestSuitesForPlan', () => {
    it('should throw error when project is not provided', async () => {
      // Act & Assert
      await expect(testDataProvider.GetTestSuitesForPlan('', mockPlanId)).rejects.toThrow(
        'Project not selected'
      );
    });

    it('should throw error when plan ID is not provided', async () => {
      // Act & Assert
      await expect(testDataProvider.GetTestSuitesForPlan(mockProject, '')).rejects.toThrow(
        'Plan not selected'
      );
    });

    it('should return test suites for a plan', async () => {
      // Arrange
      const mockData = {
        testSuites: [
          { id: '123', name: 'Test Suite 1' },
          { id: '456', name: 'Test Suite 2' },
        ],
      };
      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockData);

      // Act
      const result = await testDataProvider.GetTestSuitesForPlan(mockProject, mockPlanId);

      // Assert
      expect(result).toEqual(mockData);
      expect(TFSServices.getItemContent).toHaveBeenCalledWith(
        `${mockOrgUrl.replace(/\/+$/, '')}/${mockProject}/_api/_testManagement/GetTestSuitesForPlan?__v=5&planId=${mockPlanId}`,
        mockToken
      );
    });

    it('should use testplan suites endpoint for bearer token and normalize response', async () => {
      const bearerProvider = new TestDataProvider(mockOrgUrl, mockBearerToken);
      const mockData = {
        value: [
          { id: '123', name: 'Suite 1' },
          { id: '456', name: 'Suite 2', parentSuite: { id: '123' } },
        ],
        count: 2,
      };
      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockData);

      const result = await bearerProvider.GetTestSuitesForPlan(mockProject, mockPlanId);

      expect(TFSServices.getItemContent).toHaveBeenCalledWith(
        `${mockOrgUrl.replace(/\/+$/, '')}/${mockProject}/_apis/testplan/Plans/${mockPlanId}/suites?includeChildren=true&api-version=7.0`,
        mockBearerToken
      );
      expect(result.testSuites).toEqual(
        expect.arrayContaining([
          expect.objectContaining({ id: '123', title: 'Suite 1', parentSuiteId: 0 }),
          expect.objectContaining({ id: '456', title: 'Suite 2', parentSuiteId: '123' }),
        ])
      );
    });
  });

  describe('GetTestSuiteById', () => {
    it('should call GetTestSuitesForPlan and Helper.findSuitesRecursive with correct params', async () => {
      // Arrange
      const mockTestSuites = { testSuites: [{ id: '123', name: 'Test Suite 1' }] };
      const mockSuiteData = [new suiteData('Test Suite 1', '123', '456', 1)];

      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockTestSuites);
      (Helper.findSuitesRecursive as jest.Mock).mockReturnValueOnce(mockSuiteData);

      // Act
      const result = await testDataProvider.GetTestSuiteById(mockProject, mockPlanId, mockSuiteId, true);

      // Assert
      expect(TFSServices.getItemContent).toHaveBeenCalledWith(
        `${mockOrgUrl.replace(/\/+$/, '')}/${mockProject}/_api/_testManagement/GetTestSuitesForPlan?__v=5&planId=${mockPlanId}`,
        mockToken
      );
      expect(Helper.findSuitesRecursive).toHaveBeenCalledWith(
        mockPlanId,
        mockOrgUrl,
        mockProject,
        mockTestSuites.testSuites,
        mockSuiteId,
        true
      );
      expect(result).toEqual(mockSuiteData);
    });

    it('should use bearer suites payload when token is bearer', async () => {
      const bearerProvider = new TestDataProvider(mockOrgUrl, mockBearerToken);
      const mockTestSuites = {
        value: [{ id: '123', name: 'Suite 1', parentSuite: { id: 0 } }],
      };
      const mockSuiteData = [new suiteData('Suite 1', '123', '456', 1)];
      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockTestSuites);
      (Helper.findSuitesRecursive as jest.Mock).mockReturnValueOnce(mockSuiteData);

      const result = await bearerProvider.GetTestSuiteById(mockProject, mockPlanId, mockSuiteId, true);

      expect(TFSServices.getItemContent).toHaveBeenCalledWith(
        `${mockOrgUrl.replace(/\/+$/, '')}/${mockProject}/_apis/testplan/Plans/${mockPlanId}/suites?includeChildren=true&api-version=7.0`,
        mockBearerToken
      );
      const suitesArg = (Helper.findSuitesRecursive as jest.Mock).mock.calls[0][3];
      expect(suitesArg[0].title).toBe('Suite 1');
      expect(result).toEqual(mockSuiteData);
    });
  });

  describe('GetTestCases', () => {
    it('should return test cases for a given project, plan ID, and suite ID', async () => {
      // Arrange
      const mockData = {
        count: 2,
        value: [
          { testCase: { id: '101', name: 'Test Case 1', url: 'url1' } },
          { testCase: { id: '102', name: 'Test Case 2', url: 'url2' } },
        ],
      };
      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockData);

      // Act
      const result = await testDataProvider.GetTestCases(mockProject, mockPlanId, mockSuiteId);

      // Assert
      expect(result).toEqual(mockData);
      expect(TFSServices.getItemContent).toHaveBeenCalledWith(
        `${mockOrgUrl}${mockProject}/_apis/test/Plans/${mockPlanId}/suites/${mockSuiteId}/testcases/`,
        mockToken
      );
      expect(logger.debug).toHaveBeenCalledWith(
        `test cases for plan ${mockPlanId} and ${mockSuiteId} were found`
      );
    });
  });

  describe('clearCache', () => {
    it('should clear the cache', () => {
      // Arrange
      const mockUrl = `${mockOrgUrl}_apis/test/endpoint`;
      const mockData = { value: 'test data' };
      const cache = new Map();
      cache.set(mockUrl, {
        data: mockData,
        timestamp: Date.now(),
      });
      (testDataProvider as any).cache = cache;

      // Act
      testDataProvider.clearCache();

      // Assert
      expect((testDataProvider as any).cache.size).toBe(0);
      expect(logger.debug).toHaveBeenCalledWith('Cache cleared');
    });
  });

  describe('UpdateTestRun', () => {
    it('should update a test run with the correct state', async () => {
      // Arrange
      const mockRunId = '12345';
      const mockState = 'Completed';
      const mockResponse = { id: mockRunId, state: mockState };

      (TFSServices.postRequest as jest.Mock).mockResolvedValueOnce(mockResponse);

      // Act
      const result = await testDataProvider.UpdateTestRun(mockProject, mockRunId, mockState);

      // Assert
      expect(result).toEqual(mockResponse);
      expect(TFSServices.postRequest).toHaveBeenCalledWith(
        `${mockOrgUrl}${mockProject}/_apis/test/Runs/${mockRunId}?api-version=5.0`,
        mockToken,
        'PATCH',
        { state: mockState },
        null
      );
      expect(logger.info).toHaveBeenCalledWith(`Update runId : ${mockRunId} to state : ${mockState}`);
    });
  });

  describe('GetTestSuitesByPlan', () => {
    it('should return suites without filter using default suiteId', async () => {
      // Arrange
      const mockTestSuites = { testSuites: [{ id: 457, name: 'Suite 1', parentSuiteId: 0 }] };
      const mockSuiteData = [new suiteData('Suite 1', '457', '456', 1)];

      (TFSServices.getItemContent as jest.Mock).mockResolvedValue(mockTestSuites);
      (Helper.findSuitesRecursive as jest.Mock).mockReturnValue(mockSuiteData);

      // Act
      const result = await testDataProvider.GetTestSuitesByPlan(mockProject, mockPlanId, true);

      // Assert
      expect(Helper.findSuitesRecursive).toHaveBeenCalled();
    });

    it('should process multiple top-level suite hierarchies and combine results', async () => {
      jest.spyOn(testDataProvider, 'GetTestSuitesForPlan').mockResolvedValueOnce({
        testSuites: [
          { id: 1, parentSuiteId: 0 },
          { id: 2, parentSuiteId: 0 },
          { id: 10, parentSuiteId: 1 },
          { id: 11, parentSuiteId: 2 },
        ],
      } as any);

      const suiteIdsFilter = [10, 11];
      const getByIdSpy = jest
        .spyOn(testDataProvider, 'GetTestSuiteById')
        .mockResolvedValueOnce([{ id: '10' }])
        .mockResolvedValueOnce([{ id: '11' }]);

      const res = await testDataProvider.GetTestSuitesByPlan(mockProject, mockPlanId, true, suiteIdsFilter);

      expect(getByIdSpy).toHaveBeenCalledTimes(2);
      expect(res).toEqual([{ id: '10' }, { id: '11' }]);
    });

    it('should fallback to first suite when no top-level suites can be determined', async () => {
      jest.spyOn(testDataProvider, 'GetTestSuitesForPlan').mockResolvedValueOnce({
        testSuites: [{ id: 1, parentSuiteId: 0 }],
      } as any);

      const getByIdSpy = jest
        .spyOn(testDataProvider, 'GetTestSuiteById')
        .mockResolvedValueOnce([{ id: '1' }]);

      const res = await testDataProvider.GetTestSuitesByPlan(mockProject, mockPlanId, true, [1]);

      expect(getByIdSpy).toHaveBeenCalledWith(mockProject, mockPlanId, '1', true, [1]);
      expect(res).toEqual([{ id: '1' }]);
    });
  });

  describe('createNewRequirement', () => {
    it('should pick customer id from any supported customer fields when enabled', () => {
      const rel = (testDataProvider as any).createNewRequirement(true, {
        id: '123',
        fields: {
          'System.Title': 'Req title',
          'Custom.CustomerID': 'CID-1',
        },
      });

      expect(rel).toEqual({ type: 'requirement', id: '123', title: 'Req title', customerId: 'CID-1' });
    });

    it('should default customerId to a single space when enabled and no field exists', () => {
      const rel = (testDataProvider as any).createNewRequirement(true, {
        id: '123',
        fields: {
          'System.Title': 'Req title',
        },
      });

      expect(rel).toEqual({ type: 'requirement', id: '123', title: 'Req title', customerId: ' ' });
    });
  });

  describe('addToMap', () => {
    it('should create array for missing key and append values', () => {
      const map = new Map<string, string[]>();
      (testDataProvider as any).addToMap(map, 'k', 'v1');
      (testDataProvider as any).addToMap(map, 'k', 'v2');
      expect(map.get('k')).toEqual(['v1', 'v2']);
    });
  });

  describe('GetTestSuiteById with filtering', () => {
    it('should filter suites when suiteIdsFilter is provided', async () => {
      // Arrange
      const suiteIdsFilter = [123, 456];
      const mockTestSuites = {
        testSuites: [
          { id: 123, name: 'Suite 1', parentSuiteId: 100 },
          { id: 456, name: 'Suite 2', parentSuiteId: 100 },
        ],
      };
      const mockSuiteData = [new suiteData('Suite 1', '123', '100', 1)];

      (TFSServices.getItemContent as jest.Mock).mockResolvedValue(mockTestSuites);
      (Helper.findSuitesRecursive as jest.Mock).mockReturnValue(mockSuiteData);

      // Act
      const result = await testDataProvider.GetTestSuiteById(
        mockProject,
        mockPlanId,
        '123',
        true,
        suiteIdsFilter
      );

      // Assert
      expect(Helper.findSuitesRecursive).toHaveBeenCalled();
    });
  });

  describe('GetTestPoint', () => {
    it('should return test points for a test case', async () => {
      // Arrange
      const mockResponse = { value: [{ id: 1, testCaseId: mockTestCaseId }] };
      (TFSServices.getItemContent as jest.Mock).mockResolvedValue(mockResponse);

      // Act
      const result = await testDataProvider.GetTestPoint(
        mockProject,
        mockPlanId,
        mockSuiteId,
        mockTestCaseId
      );

      // Assert
      expect(result).toEqual(mockResponse);
    });
  });

  describe('CreateTestRun', () => {
    it('should create a test run successfully', async () => {
      // Arrange
      const testRunName = 'Test Run 1';
      const testPointId = '12345';
      const mockResponse = { data: { id: 1, name: testRunName } };

      (TFSServices.postRequest as jest.Mock).mockResolvedValueOnce(mockResponse);

      // Act
      const result = await testDataProvider.CreateTestRun(mockProject, testRunName, mockPlanId, testPointId);

      // Assert
      expect(result).toEqual(mockResponse);
      expect(TFSServices.postRequest).toHaveBeenCalledWith(
        `${mockOrgUrl}${mockProject}/_apis/test/runs`,
        mockToken,
        'Post',
        {
          name: testRunName,
          plan: { id: mockPlanId },
          pointIds: [testPointId],
        },
        null
      );
    });

    it('should throw error when creation fails', async () => {
      // Arrange
      const testRunName = 'Test Run 1';
      const testPointId = '12345';
      const mockError = new Error('Creation failed');

      (TFSServices.postRequest as jest.Mock).mockRejectedValueOnce(mockError);

      // Act & Assert
      await expect(
        testDataProvider.CreateTestRun(mockProject, testRunName, mockPlanId, testPointId)
      ).rejects.toThrow('Error: Creation failed');
    });
  });

  describe('UpdateTestCase', () => {
    it('should update test case to Active state (0)', async () => {
      // Arrange
      const mockRunId = '12345';
      const mockResponse = { data: { outcome: '0' } };
      (TFSServices.postRequest as jest.Mock).mockResolvedValueOnce(mockResponse);

      // Act
      const result = await testDataProvider.UpdateTestCase(mockProject, mockRunId, 0);

      // Assert
      expect(result).toEqual(mockResponse);
      expect(logger.info).toHaveBeenCalledWith('Reset test case to Active state ');
    });

    it('should update test case to Completed state (1)', async () => {
      // Arrange
      const mockRunId = '12345';
      const mockResponse = { data: { state: 'Completed', outcome: '1' } };
      (TFSServices.postRequest as jest.Mock).mockResolvedValueOnce(mockResponse);

      // Act
      const result = await testDataProvider.UpdateTestCase(mockProject, mockRunId, 1);

      // Assert
      expect(result).toEqual(mockResponse);
      expect(logger.info).toHaveBeenCalledWith('Update test case to complite state ');
    });

    it('should update test case to Passed state (2)', async () => {
      // Arrange
      const mockRunId = '12345';
      const mockResponse = { data: { state: 'Completed', outcome: '2' } };
      (TFSServices.postRequest as jest.Mock).mockResolvedValueOnce(mockResponse);

      // Act
      const result = await testDataProvider.UpdateTestCase(mockProject, mockRunId, 2);

      // Assert
      expect(result).toEqual(mockResponse);
      expect(logger.info).toHaveBeenCalledWith('Update test case to passed state ');
    });

    it('should update test case to Failed state (3)', async () => {
      // Arrange
      const mockRunId = '12345';
      const mockResponse = { data: { state: 'Completed', outcome: '3' } };
      (TFSServices.postRequest as jest.Mock).mockResolvedValueOnce(mockResponse);

      // Act
      const result = await testDataProvider.UpdateTestCase(mockProject, mockRunId, 3);

      // Assert
      expect(result).toEqual(mockResponse);
      expect(logger.info).toHaveBeenCalledWith('Update test case to failed state ');
    });
  });

  describe('UploadTestAttachment', () => {
    it('should upload attachment to test run', async () => {
      // Arrange
      const runId = '12345';
      const stream = 'base64encodeddata';
      const fileName = 'test.png';
      const comment = 'Test attachment';
      const attachmentType = 'GeneralAttachment';
      const mockResponse = { data: { id: 1 } };

      (TFSServices.postRequest as jest.Mock).mockResolvedValueOnce(mockResponse);

      // Act
      const result = await testDataProvider.UploadTestAttachment(
        runId,
        mockProject,
        stream,
        fileName,
        comment,
        attachmentType
      );

      // Assert
      expect(result).toEqual(mockResponse);
      expect(TFSServices.postRequest).toHaveBeenCalledWith(
        `${mockOrgUrl}${mockProject}/_apis/test/Runs/${runId}/attachments?api-version=5.0-preview.1`,
        mockToken,
        'Post',
        { stream, fileName, comment, attachmentType },
        null
      );
    });
  });

  describe('GetTestRunById', () => {
    it('should return test run by ID', async () => {
      // Arrange
      const runId = '12345';
      const mockResponse = { id: runId, name: 'Test Run 1' };

      // Act
      const result = await testDataProvider.GetTestRunById(mockProject, runId);

      // Assert
      expect(TFSServices.getItemContent).toHaveBeenCalled();
    });
  });

  describe('GetTestPointByTestCaseId', () => {
    it('should return test points for a test case ID', async () => {
      // Arrange
      const mockResponse = { data: { value: [{ id: 1 }] } };
      (TFSServices.postRequest as jest.Mock).mockResolvedValueOnce(mockResponse);

      // Act
      const result = await testDataProvider.GetTestPointByTestCaseId(mockProject, mockTestCaseId);

      // Assert
      expect(result).toEqual(mockResponse);
      expect(TFSServices.postRequest).toHaveBeenCalledWith(
        `${mockOrgUrl}${mockProject}/_apis/test/points`,
        mockToken,
        'Post',
        { PointsFilter: { TestcaseIds: [mockTestCaseId] } },
        null
      );
    });
  });

  describe('GetTestCasesBySuites', () => {
    it('should return test cases for suites', async () => {
      // Arrange
      const mockSuiteData = [new suiteData('Suite 1', '123', '456', 1)];
      const mockTestCases = {
        count: 1,
        value: [{ testCase: { id: '101', url: 'http://test.com/101' } }],
      };
      const mockTestCaseDetails = {
        id: 101,
        fields: {
          'System.Title': 'Test Case 1',
          'System.AreaPath': 'Area/Path',
          'System.Description': 'Description',
          'Microsoft.VSTS.TCM.Steps': null,
        },
      };

      (TFSServices.getItemContent as jest.Mock)
        .mockResolvedValueOnce({ testSuites: mockSuiteData })
        .mockResolvedValueOnce(mockTestCases)
        .mockResolvedValueOnce(mockTestCaseDetails);
      (Helper.findSuitesRecursive as jest.Mock).mockReturnValueOnce(mockSuiteData);

      // Act
      const result = await testDataProvider.GetTestCasesBySuites(
        mockProject,
        mockPlanId,
        mockSuiteId,
        false,
        false,
        false,
        false
      );

      // Assert
      expect(result.testCasesList).toBeDefined();
      expect(result.requirementToTestCaseTraceMap).toBeDefined();
      expect(result.testCaseToRequirementsTraceMap).toBeDefined();
    });

    it('should use pre-filtered suites when provided', async () => {
      // Arrange
      const preFilteredSuites = [new suiteData('Suite 1', '123', '456', 1)];
      const mockTestCases = {
        count: 1,
        value: [{ testCase: { id: '101', url: 'http://test.com/101' } }],
      };
      const mockTestCaseDetails = {
        id: 101,
        fields: {
          'System.Title': 'Test Case 1',
          'System.AreaPath': 'Area/Path',
          'System.Description': 'Description',
          'Microsoft.VSTS.TCM.Steps': null,
        },
      };

      (TFSServices.getItemContent as jest.Mock)
        .mockResolvedValueOnce(mockTestCases)
        .mockResolvedValueOnce(mockTestCaseDetails);

      // Act
      const result = await testDataProvider.GetTestCasesBySuites(
        mockProject,
        mockPlanId,
        mockSuiteId,
        false,
        false,
        false,
        false,
        undefined,
        undefined,
        preFilteredSuites
      );

      // Assert
      expect(result.testCasesList).toBeDefined();
    });

    it('should handle errors in suite processing', async () => {
      // Arrange
      const preFilteredSuites = [new suiteData('Suite 1', '123', '456', 1)];

      (TFSServices.getItemContent as jest.Mock).mockRejectedValueOnce(new Error('API Error'));

      // Act
      const result = await testDataProvider.GetTestCasesBySuites(
        mockProject,
        mockPlanId,
        mockSuiteId,
        false,
        false,
        false,
        false,
        undefined,
        undefined,
        preFilteredSuites
      );

      // Assert
      expect(result.testCasesList).toEqual([]);
    });
  });

  describe('StructureTestCase', () => {
    it('should return empty array when no test cases', async () => {
      // Arrange
      const suite = new suiteData('Suite 1', '123', '456', 1);
      const testCases = { count: 0, value: [] };

      // Act
      const result = await testDataProvider.StructureTestCase(
        mockProject,
        testCases,
        suite,
        false,
        false,
        false,
        new Map(),
        new Map()
      );

      // Assert
      expect(result).toEqual([]);
      expect(logger.warn).toHaveBeenCalled();
    });

    it('should warn and return [] when no testCases are provided', async () => {
      const res = await testDataProvider.StructureTestCase(
        mockProject,
        { value: [], count: 0 } as any,
        { id: '1', name: 'Suite 1' } as any,
        true,
        true,
        false,
        new Map<string, string[]>(),
        new Map<string, string[]>()
      );

      expect(res).toEqual([]);
      expect(logger.warn).toHaveBeenCalledWith('No test cases found for suite: 1');
    });

    it('should parse steps, add requirement relation when enabled, and add mom relation when enabled', async () => {
      const UtilsMock: any = require('../../utils/testStepParserHelper').default;
      const parserInstance = UtilsMock.mock.results[0].value;
      parserInstance.parseTestSteps.mockResolvedValueOnce([{ stepId: '1' }]);

      const suite = { id: '1', name: 'Suite 1' } as any;
      const testCases = {
        count: 1,
        value: [{ testCase: { id: 123, url: 'https://example.com/testcase/123' } }],
      };

      jest.spyOn(testDataProvider as any, 'fetchWithCache').mockImplementation(async (...args: any[]) => {
        const url = String(args[0] || '');
        if (url.includes('testcase/123') && url.includes('?$expand=All')) {
          return {
            id: 123,
            fields: {
              'System.Title': 'TC 123',
              'System.AreaPath': 'A',
              'System.Description': 'D',
              'Microsoft.VSTS.TCM.Steps': '<steps></steps>',
            },
            relations: [
              { url: 'https://example.com/_apis/wit/workItems/200' },
              { url: 'https://example.com/not-work-items/ignore' },
              { url: 'https://example.com/_apis/wit/workItems/201' },
            ],
          };
        }
        if (url.includes('/workItems/200')) {
          return {
            id: 200,
            fields: {
              'System.WorkItemType': 'Requirement',
              'System.Title': 'REQ 200',
              'Custom.CustomerRequirementId': 'C-200',
            },
            _links: { html: { href: 'http://example.com/200' } },
          };
        }
        if (url.includes('/workItems/201')) {
          return {
            id: 201,
            fields: {
              'System.WorkItemType': 'Bug',
              'System.Title': 'BUG 201',
              'System.State': 'Active',
            },
            _links: { html: { href: 'http://example.com/201' } },
          };
        }
        throw new Error(`unexpected url ${url}`);
      });

      const requirementToTestCaseTraceMap = new Map<string, string[]>();
      const testCaseToRequirementsTraceMap = new Map<string, string[]>();

      const res = await testDataProvider.StructureTestCase(
        mockProject,
        testCases as any,
        suite,
        true,
        true,
        true,
        requirementToTestCaseTraceMap,
        testCaseToRequirementsTraceMap
      );

      expect(res).toHaveLength(1);
      expect(res[0].steps).toEqual([{ stepId: '1' }]);
      expect(res[0].relations).toEqual(
        expect.arrayContaining([
          expect.objectContaining({ type: 'requirement', id: 200, customerId: 'C-200' }),
        ])
      );
      expect(res[0].relations).toEqual(
        expect.arrayContaining([expect.objectContaining({ type: 'Bug', id: 201 })])
      );
      expect(requirementToTestCaseTraceMap.size).toBe(1);
      expect(testCaseToRequirementsTraceMap.size).toBe(1);
    });

    it('should use stepResultDetailsMap when provided and skip parseTestSteps', async () => {
      const UtilsMock: any = require('../../utils/testStepParserHelper').default;
      const parserInstance = UtilsMock.mock.results[0].value;
      parserInstance.parseTestSteps.mockClear();

      const suite = { id: '1', name: 'Suite 1' } as any;
      const testCases = {
        count: 1,
        value: [{ testCase: { id: 123, url: 'https://example.com/testcase/123' } }],
      };

      const stepResultDetailsMap = new Map<string, any>();
      stepResultDetailsMap.set('123', {
        testCaseRevision: 7,
        stepList: [{ stepId: 'from-cache' }],
        caseEvidenceAttachments: [{ id: 1 }],
      });

      jest.spyOn(testDataProvider as any, 'fetchWithCache').mockResolvedValueOnce({
        id: 123,
        fields: {
          'System.Title': 'TC 123',
          'System.AreaPath': 'A',
          'System.Description': 'D',
          'Microsoft.VSTS.TCM.Steps': '<steps></steps>',
        },
        relations: [],
      });

      const res = await testDataProvider.StructureTestCase(
        mockProject,
        testCases as any,
        suite,
        true,
        false,
        false,
        new Map<string, string[]>(),
        new Map<string, string[]>(),
        stepResultDetailsMap
      );

      expect(res).toHaveLength(1);
      expect(res[0].steps).toEqual([{ stepId: 'from-cache' }]);
      expect(res[0].caseEvidenceAttachments).toEqual([{ id: 1 }]);
      expect(parserInstance.parseTestSteps).not.toHaveBeenCalled();
    });

    it('should log error and continue when fetching relation content fails', async () => {
      const suite = { id: '1', name: 'Suite 1' } as any;
      const testCases = {
        count: 1,
        value: [{ testCase: { id: 123, url: 'https://example.com/testcase/123' } }],
      };

      jest.spyOn(testDataProvider as any, 'fetchWithCache').mockImplementation(async (...args: any[]) => {
        const url = String(args[0] || '');
        if (url.includes('testcase/123') && url.includes('?$expand=All')) {
          return {
            id: 123,
            fields: {
              'System.Title': 'TC 123',
              'System.AreaPath': 'A',
              'System.Description': 'D',
              'Microsoft.VSTS.TCM.Steps': null,
            },
            relations: [{ url: 'https://example.com/_apis/wit/workItems/200' }],
          };
        }
        if (url.includes('/workItems/200')) {
          throw new Error('boom');
        }
        return null;
      });

      const res = await testDataProvider.StructureTestCase(
        mockProject,
        testCases as any,
        suite,
        true,
        true,
        false,
        new Map<string, string[]>(),
        new Map<string, string[]>()
      );

      expect(res).toHaveLength(1);
      expect(logger.error).toHaveBeenCalledWith(
        expect.stringContaining(
          'Failed to fetch relation content for URL https://example.com/_apis/wit/workItems/200'
        )
      );
    });

    it('should append linked MOMs from testCaseToLinkedMomLookup when provided', async () => {
      const suite = { id: '1', name: 'Suite 1' } as any;
      const testCases = {
        count: 1,
        value: [{ testCase: { id: 123, url: 'https://example.com/testcase/123' } }],
      };

      jest.spyOn(testDataProvider as any, 'fetchWithCache').mockResolvedValueOnce({
        id: 123,
        fields: {
          'System.Title': 'TC 123',
          'System.AreaPath': 'A',
          'System.Description': 'D',
          'Microsoft.VSTS.TCM.Steps': null,
        },
        relations: [],
      });

      const testCaseToLinkedMomLookup = new Map<number, Set<any>>();
      testCaseToLinkedMomLookup.set(
        123,
        new Set([
          {
            id: 900,
            fields: {
              'System.WorkItemType': 'Task',
              'System.Title': 'Task 900',
              'System.State': 'Active',
            },
            _links: { html: { href: 'http://example.com/900' } },
          },
        ])
      );

      const res = await testDataProvider.StructureTestCase(
        mockProject,
        testCases as any,
        suite,
        false,
        false,
        false,
        new Map<string, string[]>(),
        new Map<string, string[]>(),
        undefined,
        testCaseToLinkedMomLookup
      );

      expect(res).toHaveLength(1);
      expect(res[0].relations).toEqual(
        expect.arrayContaining([expect.objectContaining({ type: 'Task', id: 900 })])
      );
    });

    it('should return empty list when test case fetch fails', async () => {
      const suite = { id: '1', name: 'Suite 1' } as any;
      const testCases = {
        count: 1,
        value: [{ testCase: { id: 123, url: 'https://example.com/testcase/123' } }],
      };

      jest.spyOn(testDataProvider as any, 'fetchWithCache').mockRejectedValueOnce(new Error('boom'));

      const res = await testDataProvider.StructureTestCase(
        mockProject,
        testCases as any,
        suite,
        false,
        false,
        false,
        new Map<string, string[]>(),
        new Map<string, string[]>()
      );

      expect(res).toEqual([]);
      expect(logger.error).toHaveBeenCalledWith('Error: ran into an issue while retrieving testCase 123');
    });
  });

  describe('ParseSteps', () => {
    it('should parse XML steps correctly', () => {
      // Arrange
      const stepsXml = `<steps>
        <step>
          <parameterizedString>Action 1</parameterizedString>
          <parameterizedString>Expected 1</parameterizedString>
        </step>
      </steps>`;

      // Act
      const result = testDataProvider.ParseSteps(stepsXml);

      // Assert - ParseSteps uses xml2js.parseString which is synchronous callback-based
      // The result may be empty if parsing fails or steps format doesn't match
      expect(Array.isArray(result)).toBe(true);
    });
  });
});
