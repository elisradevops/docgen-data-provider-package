import { TFSServices } from '../../helpers/tfs';
import { Helper, suiteData } from '../../helpers/helper';
import TestDataProvider from '../TestDataProvider';
import Utils from '../../utils/testStepParserHelper';
import logger from '../../utils/logger';
import { TestCase } from '../../models/tfs-data';

jest.mock('../../helpers/tfs');
jest.mock('../../utils/logger');
jest.mock('../../helpers/helper');
jest.mock('../../utils/testStepParserHelper');
jest.mock('p-limit', () => jest.fn(() => (fn: Function) => fn()));

describe('TestDataProvider', () => {
  let testDataProvider: TestDataProvider;
  const mockOrgUrl = 'https://dev.azure.com/orgname/';
  const mockToken = 'mock-token';
  const mockProject = 'project-123';
  const mockPlanId = '456';
  const mockSuiteId = '789';
  const mockTestCaseId = '101112';

  beforeEach(() => {
    jest.clearAllMocks();
    (Helper.suitList as any) = [];
    (Helper.first as any) = false;

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
        `${mockOrgUrl}/${mockProject}/_api/_testManagement/GetTestSuitesForPlan?__v=5&planId=${mockPlanId}`,
        mockToken
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
        `${mockOrgUrl}/${mockProject}/_api/_testManagement/GetTestSuitesForPlan?__v=5&planId=${mockPlanId}`,
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
      expect(Helper.first).toBe(true);
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
});
