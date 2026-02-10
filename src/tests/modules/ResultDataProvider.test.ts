import { TFSServices } from '../../helpers/tfs';
import ResultDataProvider from '../../modules/ResultDataProvider';
import logger from '../../utils/logger';
import Utils from '../../utils/testStepParserHelper';

// Mock dependencies
jest.mock('../../helpers/tfs');
jest.mock('../../utils/logger');
jest.mock('../../utils/testStepParserHelper');
jest.mock('../../modules/TicketsDataProvider', () => {
  return {
    __esModule: true,
    default: jest.fn().mockImplementation(() => ({
      GetQueryResultsFromWiql: jest.fn(),
    })),
  };
});
jest.mock('p-limit', () => () => (fn: Function) => fn());

describe('ResultDataProvider', () => {
  let resultDataProvider: ResultDataProvider;
  const mockOrgUrl = 'https://dev.azure.com/organization/';
  const mockToken = 'mock-token';
  const mockProjectName = 'test-project';
  const mockTestPlanId = '12345';

  beforeEach(() => {
    jest.clearAllMocks();
    resultDataProvider = new ResultDataProvider(mockOrgUrl, mockToken);
  });

  describe('Utility methods', () => {
    describe('flattenSuites', () => {
      it('should flatten a hierarchical suite structure into a single-level array', () => {
        // Arrange
        const suites = [
          {
            id: 1,
            name: 'Parent 1',
            children: [
              { id: 2, name: 'Child 1' },
              {
                id: 3,
                name: 'Child 2',
                children: [{ id: 4, name: 'Grandchild 1' }],
              },
            ],
          },
          { id: 5, name: 'Parent 2' },
        ];

        // Act
        const result: any[] = (resultDataProvider as any).flattenSuites(suites);

        // Assert
        expect(result).toHaveLength(5);
        expect(result.map((s) => s.id)).toEqual([1, 2, 3, 4, 5]);
      });
    });

    describe('filterSuites', () => {
      it('should filter suites based on selected suite IDs', () => {
        // Arrange
        const testSuites = [
          { id: 1, name: 'Suite 1', parentSuite: { id: 0 } },
          { id: 2, name: 'Suite 2', parentSuite: { id: 1 } },
          { id: 3, name: 'Suite 3', parentSuite: { id: 1 } },
        ];
        const selectedSuiteIds = [1, 3];

        // Act
        const result: any[] = (resultDataProvider as any).filterSuites(testSuites, selectedSuiteIds);

        // Assert
        expect(result).toHaveLength(2);
        expect(result.map((s) => s.id)).toEqual([1, 3]);
      });

      it('should return all suites with parent when no suite IDs are selected', () => {
        // Arrange
        const testSuites = [
          { id: 1, name: 'Suite 1', parentSuite: { id: 0 } },
          { id: 2, name: 'Suite 2', parentSuite: { id: 1 } },
          { id: 3, name: 'Suite 3', parentSuite: null },
        ];

        // Act
        const result: any[] = (resultDataProvider as any).filterSuites(testSuites);

        // Assert
        expect(result).toHaveLength(2);
        expect(result.map((s) => s.id)).toEqual([1, 2]);
      });
    });

    describe('buildTestGroupName', () => {
      it('should return simple suite name when hierarchy is disabled', () => {
        // Arrange
        const suiteMap = new Map([[1, { id: 1, name: 'Suite 1', parentSuite: { id: 0 } }]]);

        // Act
        const result = (resultDataProvider as any).buildTestGroupName(1, suiteMap, false);

        // Assert
        expect(result).toBe('Suite 1');
      });

      it('should build hierarchical name with parent info', () => {
        // Arrange
        const suiteMap = new Map([
          [1, { id: 1, name: 'Parent', parentSuite: null }],
          [2, { id: 2, name: 'Child', parentSuite: { id: 1 } }],
        ]);

        // Act
        const result = (resultDataProvider as any).buildTestGroupName(2, suiteMap, true);

        // Assert
        expect(result).toBe('Child');
      });

      it('should abbreviate deep hierarchies', () => {
        // Arrange
        const suiteMap = new Map([
          [1, { id: 1, name: 'Root', parentSuite: null }],
          [2, { id: 2, name: 'Level1', parentSuite: { id: 1 } }],
          [3, { id: 3, name: 'Level2', parentSuite: { id: 2 } }],
          [4, { id: 4, name: 'Level3', parentSuite: { id: 3 } }],
        ]);

        // Act
        const result = (resultDataProvider as any).buildTestGroupName(4, suiteMap, true);

        // Assert
        expect(result).toBe('Level1/.../Level3');
      });
    });

    describe('convertRunStatus', () => {
      it('should convert API status to readable format', () => {
        // Arrange & Act & Assert
        expect((resultDataProvider as any).convertRunStatus('passed')).toBe('Passed');
        expect((resultDataProvider as any).convertRunStatus('failed')).toBe('Failed');
        expect((resultDataProvider as any).convertRunStatus('notApplicable')).toBe('Not Applicable');
        expect((resultDataProvider as any).convertRunStatus('unknown')).toBe('Not Run');
      });
    });

    describe('compareActionResults', () => {
      it('should compare version-like step positions correctly', () => {
        // Act & Assert
        const compare = (resultDataProvider as any).compareActionResults;
        expect(compare('1', '2')).toBe(-1);
        expect(compare('2', '1')).toBe(1);
        expect(compare('1.1', '1.2')).toBe(-1);
        expect(compare('1.2', '1.1')).toBe(1);
        expect(compare('1.1', '1.1')).toBe(0);
        expect(compare('1.1.1', '1.1')).toBe(1);
        expect(compare('1.1', '1.1.1')).toBe(-1);
      });
    });
  });

  describe('Data fetching methods', () => {
    describe('fetchTestSuites', () => {
      it('should fetch and process test suites correctly', async () => {
        // Arrange
        const mockTestSuites = {
          value: [
            {
              id: 1,
              name: 'Root Suite',
              children: [{ id: 2, name: 'Child Suite 1', parentSuite: { id: 1 } }],
            },
          ],
          count: 1,
        };

        (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockTestSuites);

        // Act
        const result = await (resultDataProvider as any).fetchTestSuites(mockTestPlanId, mockProjectName);

        // Assert
        expect(TFSServices.getItemContent).toHaveBeenCalledWith(
          `${mockOrgUrl}${mockProjectName}/_apis/testplan/Plans/${mockTestPlanId}/Suites?asTreeView=true`,
          mockToken
        );
        expect(result).toHaveLength(1);
        expect(result[0]).toHaveProperty('testSuiteId', 2);
        expect(result[0]).toHaveProperty('testGroupName');
      });

      it('should handle errors and return empty array', async () => {
        // Arrange
        const mockError = new Error('API error');
        (TFSServices.getItemContent as jest.Mock).mockRejectedValueOnce(mockError);

        // Act
        const result = await (resultDataProvider as any).fetchTestSuites(mockTestPlanId, mockProjectName);

        // Assert
        expect(logger.error).toHaveBeenCalled();
        expect(result).toEqual([]);
      });
    });

    describe('fetchTestPoints', () => {
      it('should fetch and map test points correctly', async () => {
        // Arrange
        const mockSuiteId = '123';
        const mockTestPoints = {
          value: [
            {
              testCaseReference: { id: 1, name: 'Test Case 1' },
              configuration: { name: 'Config 1' },
              results: {
                outcome: 'passed',
                lastTestRunId: 100,
                lastResultId: 200,
                lastResultDetails: { dateCompleted: '2023-01-01', runBy: { displayName: 'Test User' } },
              },
            },
          ],
          count: 1,
        };

        (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockTestPoints);

        // Act
        const result = await (resultDataProvider as any).fetchTestPoints(
          mockProjectName,
          mockTestPlanId,
          mockSuiteId
        );

        // Assert
        expect(TFSServices.getItemContent).toHaveBeenCalledWith(
          `${mockOrgUrl}${mockProjectName}/_apis/testplan/Plans/${mockTestPlanId}/Suites/${mockSuiteId}/TestPoint?includePointDetails=true`,
          mockToken
        );
        expect(result).toHaveLength(1);
        expect(result[0]).toEqual({
          testCaseId: 1,
          testCaseName: 'Test Case 1',
          configurationName: 'Config 1',
          outcome: 'passed',
          lastRunId: 100,
          lastResultId: 200,
          lastResultDetails: { dateCompleted: '2023-01-01', runBy: { displayName: 'Test User' } },
          testCaseUrl: 'https://dev.azure.com/organization/test-project/_workitems/edit/1',
        });
      });

      it('should handle errors and return empty array', async () => {
        // Arrange
        const mockSuiteId = '123';
        const mockError = new Error('API error');
        (TFSServices.getItemContent as jest.Mock).mockRejectedValueOnce(mockError);

        // Act
        const result = await (resultDataProvider as any).fetchTestPoints(
          mockProjectName,
          mockTestPlanId,
          mockSuiteId
        );

        // Assert
        expect(logger.error).toHaveBeenCalled();
        expect(result).toEqual([]);
      });
    });
  });

  describe('Data transformation methods', () => {
    describe('mapTestPoint', () => {
      it('should transform test point data correctly', () => {
        // Arrange
        const testPoint = {
          testCaseReference: { id: 1, name: 'Test Case 1' },
          configuration: { name: 'Config 1' },
          results: {
            outcome: 'passed',
            lastTestRunId: 100,
            lastResultId: 200,
            lastResultDetails: { dateCompleted: '2023-01-01', runBy: { displayName: 'Test User' } },
          },
        };

        // Act
        const result = (resultDataProvider as any).mapTestPoint(testPoint, mockProjectName);

        // Assert
        expect(result).toEqual({
          testCaseId: 1,
          testCaseName: 'Test Case 1',
          configurationName: 'Config 1',
          outcome: 'passed',
          lastRunId: 100,
          lastResultId: 200,
          lastResultDetails: { dateCompleted: '2023-01-01', runBy: { displayName: 'Test User' } },
          testCaseUrl: 'https://dev.azure.com/organization/test-project/_workitems/edit/1',
        });
      });

      it('should handle missing fields', () => {
        // Arrange
        const testPoint = {
          testCaseReference: { id: 1, name: 'Test Case 1' },
          // No configuration or results
        };

        // Act
        const result = (resultDataProvider as any).mapTestPoint(testPoint, mockProjectName);

        // Assert
        expect(result).toEqual({
          testCaseId: 1,
          testCaseName: 'Test Case 1',
          configurationName: undefined,
          outcome: 'Not Run',
          lastRunId: undefined,
          lastResultId: undefined,
          lastResultDetails: undefined,
          testCaseUrl: 'https://dev.azure.com/organization/test-project/_workitems/edit/1',
        });
      });
    });

    describe('calculateGroupResultSummary', () => {
      it('should return empty strings when includeHardCopyRun is true', () => {
        // Arrange
        const testPoints = [{ outcome: 'passed' }, { outcome: 'failed' }];

        // Act
        const result = (resultDataProvider as any).calculateGroupResultSummary(testPoints, true);

        // Assert
        expect(result).toEqual({
          passed: '',
          failed: '',
          notApplicable: '',
          blocked: '',
          notRun: '',
          total: '',
          successPercentage: '',
        });
      });

      it('should calculate summary statistics correctly', () => {
        // Arrange
        const testPoints = [
          { outcome: 'passed' },
          { outcome: 'passed' },
          { outcome: 'failed' },
          { outcome: 'notApplicable' },
          { outcome: 'blocked' },
          { outcome: 'something else' },
        ];

        // Act
        const result = (resultDataProvider as any).calculateGroupResultSummary(testPoints, false);

        // Assert
        expect(result).toEqual({
          passed: 2,
          failed: 1,
          notApplicable: 1,
          blocked: 1,
          notRun: 1,
          total: 6,
          successPercentage: '33.33%',
        });
      });

      it('should handle empty array', () => {
        // Arrange
        const testPoints: any[] = [];

        // Act
        const result = (resultDataProvider as any).calculateGroupResultSummary(testPoints, false);

        // Assert
        expect(result).toEqual({
          passed: 0,
          failed: 0,
          notApplicable: 0,
          blocked: 0,
          notRun: 0,
          total: 0,
          successPercentage: '0.00%',
        });
      });
    });

    describe('mapAttachmentsUrl', () => {
      it('should map attachment URLs correctly', () => {
        // Arrange
        const mockRunResults = [
          {
            testCaseId: 1,
            lastRunId: 100,
            lastResultId: 200,
            iteration: {
              attachments: [{ id: 1, name: 'attachment1.png', actionPath: 'path1' }],
              actionResults: [{ actionPath: 'path1', stepPosition: '1.1' }],
            },
            analysisAttachments: [{ id: 2, fileName: 'analysis1.txt' }],
          },
        ];

        // Act
        const result = resultDataProvider.mapAttachmentsUrl(mockRunResults, mockProjectName);

        // Assert
        expect(result[0].iteration.attachments[0].downloadUrl).toBe(
          `${mockOrgUrl}${mockProjectName}/_apis/test/runs/100/results/200/attachments/1/attachment1.png`
        );
        expect(result[0].iteration.attachments[0].stepNo).toBe('1.1');
        expect(result[0].analysisAttachments[0].downloadUrl).toBe(
          `${mockOrgUrl}${mockProjectName}/_apis/test/runs/100/results/200/attachments/2/analysis1.txt`
        );
      });

      it('should handle missing iteration', () => {
        // Arrange
        const mockRunResults = [
          {
            testCaseId: 1,
            lastRunId: 100,
            lastResultId: 200,
            // No iteration
            analysisAttachments: [{ id: 2, fileName: 'analysis1.txt' }],
          },
        ];

        // Act
        const result = resultDataProvider.mapAttachmentsUrl(mockRunResults, mockProjectName);

        // Assert
        expect(result[0]).toEqual(mockRunResults[0]);
      });
    });
  });

  describe('formatTestResult', () => {
    it('should format test result correctly', () => {
      // Arrange
      const testPoint = {
        testCaseId: 1,
        testCaseName: 'Test Case 1',
        testGroupName: 'Suite 1',
        testCaseUrl: 'https://example.com/workitems/1',
        configurationName: 'Config 1',
        outcome: 'passed',
      };

      // Act
      const result = (resultDataProvider as any).formatTestResult(testPoint, true, false);

      // Assert
      expect(result.testId).toBe(1);
      expect(result.testName).toBe('Test Case 1');
      expect(result.configuration).toBe('Config 1');
      expect(result.runStatus).toBe('Passed');
    });

    it('should return empty strings when includeHardCopyRun is true', () => {
      // Arrange
      const testPoint = {
        testCaseId: 1,
        testCaseName: 'Test Case 1',
        testGroupName: 'Suite 1',
        outcome: 'passed',
      };

      // Act
      const result = (resultDataProvider as any).formatTestResult(testPoint, false, true);

      // Assert
      expect(result.runStatus).toBe('');
    });
  });

  describe('calculateTotalSummary', () => {
    it('should calculate total summary from summarized results', () => {
      // Arrange
      const summarizedResults = [
        { groupResultSummary: { passed: 2, failed: 1, notApplicable: 0, blocked: 0, notRun: 1, total: 4 } },
        { groupResultSummary: { passed: 3, failed: 0, notApplicable: 1, blocked: 1, notRun: 0, total: 5 } },
      ];

      // Act
      const result = (resultDataProvider as any).calculateTotalSummary(summarizedResults, false);

      // Assert
      expect(result.passed).toBe(5);
      expect(result.failed).toBe(1);
      expect(result.notApplicable).toBe(1);
      expect(result.blocked).toBe(1);
      expect(result.notRun).toBe(1);
      expect(result.total).toBe(9);
    });

    it('should return empty strings when includeHardCopyRun is true', () => {
      // Arrange
      const summarizedResults = [{ groupResultSummary: { passed: 2, failed: 1, total: 3 } }];

      // Act
      const result = (resultDataProvider as any).calculateTotalSummary(summarizedResults, true);

      // Assert
      expect(result.passed).toBe('');
      expect(result.failed).toBe('');
      expect(result.total).toBe('');
    });
  });

  describe('flattenTestPoints', () => {
    it('should flatten test points from suites', () => {
      // Arrange
      const testPoints = [
        { testSuiteId: 1, testGroupName: 'Suite 1', testPointsItems: [{ id: 1 }, { id: 2 }] },
        { testSuiteId: 2, testGroupName: 'Suite 2', testPointsItems: [{ id: 3 }] },
      ];

      // Act
      const result = (resultDataProvider as any).flattenTestPoints(testPoints);

      // Assert
      expect(result).toHaveLength(3);
      expect(result[0].testGroupName).toBe('Suite 1');
      expect(result[2].testGroupName).toBe('Suite 2');
    });
  });

  describe('createSuiteMap', () => {
    it('should create a map of suites by ID', () => {
      // Arrange
      const suites = [
        { id: 1, name: 'Suite 1', children: [{ id: 2, name: 'Suite 2' }] },
        { id: 3, name: 'Suite 3' },
      ];

      // Act
      const result = (resultDataProvider as any).createSuiteMap(suites);

      // Assert
      expect(result).toBeInstanceOf(Map);
      expect(result.size).toBe(3);
      expect(result.get(1).name).toBe('Suite 1');
      expect(result.get(2).name).toBe('Suite 2');
    });
  });

  describe('isNotRunStep', () => {
    it('should return true for not run step', () => {
      // Arrange
      const step = { stepStatus: 'Not Run' };

      // Act
      const result = (resultDataProvider as any).isNotRunStep(step);

      // Assert
      expect(result).toBe(true);
    });

    it('should return false for run step', () => {
      // Arrange
      const step = { stepStatus: 'Passed' };

      // Act
      const result = (resultDataProvider as any).isNotRunStep(step);

      // Assert
      expect(result).toBe(false);
    });
  });

  describe('CreateAttachmentPathIndexMap', () => {
    it('should create map of action paths to step positions', () => {
      // Arrange
      const actionResults = [
        { actionPath: 'path1', stepPosition: '1.1' },
        { actionPath: 'path2', stepPosition: '1.2' },
      ];

      // Act
      const result = (resultDataProvider as any).CreateAttachmentPathIndexMap(actionResults);

      // Assert
      expect(result).toBeInstanceOf(Map);
      expect(result.get('path1')).toBe('1.1');
      expect(result.get('path2')).toBe('1.2');
    });
  });

  describe('getTestPointsForTestCases', () => {
    it('should fetch test points for test cases', async () => {
      // Arrange
      const mockTestCaseIds = ['1', '2'];
      const mockResponse = { data: { points: [{ id: 1 }] } };
      (TFSServices.postRequest as jest.Mock).mockResolvedValueOnce(mockResponse);

      // Act
      const result = await resultDataProvider.getTestPointsForTestCases(mockProjectName, mockTestCaseIds);

      // Assert
      expect(TFSServices.postRequest).toHaveBeenCalled();
      expect(result).toEqual(mockResponse);
    });
  });

  describe('mapActionPathToPosition', () => {
    it('should map action paths to positions', () => {
      // Arrange
      const actionResults = [
        { testId: 1, actionPath: 'path1', stepNo: '1.1' },
        { testId: 1, actionPath: 'path2', stepNo: '1.2' },
      ];

      // Act
      const result = (resultDataProvider as any).mapActionPathToPosition(actionResults);

      // Assert
      expect(result).toBeInstanceOf(Map);
      expect(result.get('1-path1')).toBe('1.1');
      expect(result.get('1-path2')).toBe('1.2');
    });
  });

  describe('fetchTestPlanName', () => {
    it('should fetch test plan name', async () => {
      // Arrange
      const mockPlan = { name: 'Test Plan 1' };
      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockPlan);

      // Act
      const result = await (resultDataProvider as any).fetchTestPlanName(mockTestPlanId, mockProjectName);

      // Assert
      expect(result).toBe('Test Plan 1');
    });

    it('should return empty string on error', async () => {
      // Arrange
      (TFSServices.getItemContent as jest.Mock).mockRejectedValueOnce(new Error('API Error'));

      // Act
      const result = await (resultDataProvider as any).fetchTestPlanName(mockTestPlanId, mockProjectName);

      // Assert
      expect(result).toBe('');
      expect(logger.error).toHaveBeenCalled();
    });
  });

  describe('fetchTestCasesBySuiteId', () => {
    it('should fetch test cases by suite ID', async () => {
      // Arrange
      const mockSuiteId = '123';
      const mockTestCases = { value: [{ workItem: { id: 1 } }] };
      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockTestCases);

      // Act
      const result = await (resultDataProvider as any).fetchTestCasesBySuiteId(
        mockProjectName,
        mockTestPlanId,
        mockSuiteId
      );

      // Assert
      expect(result).toEqual(mockTestCases.value);
    });
  });

  describe('mapTestPointForCrossPlans', () => {
    it('should map test point for cross plans', () => {
      // Arrange
      const testPoint = {
        testCase: { id: 1, name: 'Test Case 1' },
        testSuite: { id: 2, name: 'Suite 1' },
        configuration: { name: 'Config 1' },
        outcome: 'passed',
        lastTestRun: { id: '100' },
        lastResult: { id: '200' },
      };

      // Act
      const result = (resultDataProvider as any).mapTestPointForCrossPlans(testPoint, mockProjectName);

      // Assert
      expect(result.testCaseId).toBe(1);
      expect(result.testCaseName).toBe('Test Case 1');
      expect(result.outcome).toBe('passed');
      expect(result.lastRunId).toBe('100');
    });

    it('should provide default values for missing fields', () => {
      // Arrange
      const testPoint = {
        testCase: { id: 1, name: 'Test Case 1' },
        testSuite: { id: 2, name: 'Suite 1' },
      };

      // Act
      const result = (resultDataProvider as any).mapTestPointForCrossPlans(testPoint, mockProjectName);

      // Assert
      expect(result.outcome).toBe('Not Run');
      expect(result.lastResultDetails).toBeDefined();
      expect(result.lastResultDetails.duration).toBe(0);
    });
  });

  describe('mapStepResultsForExecutionAppendix', () => {
    it('should map step results for execution appendix', () => {
      // Arrange
      const detailedResults = [
        {
          testId: 1,
          testCaseRevision: { rev: 1 },
          stepNo: '1.1',
          stepIdentifier: 'step1',
          stepAction: 'Do something',
          stepExpected: 'Something happens',
          stepStatus: 'Passed',
          stepComments: '',
          isSharedStepTitle: false,
          actionPath: 'path1',
        },
      ];
      const runResultData = [
        {
          testCaseId: 1,
          iteration: {
            attachments: [{ name: 'attachment.png', actionPath: 'path1', downloadUrl: 'http://example.com' }],
          },
        },
      ];

      // Act
      const result = (resultDataProvider as any).mapStepResultsForExecutionAppendix(
        detailedResults,
        runResultData
      );

      // Assert
      expect(result).toBeInstanceOf(Map);
      expect(result.has('1')).toBe(true);
    });
  });

  describe('fetchCrossTestPoints', () => {
    it('should return empty array when no test case IDs provided', async () => {
      // Act
      const result = await (resultDataProvider as any).fetchCrossTestPoints(mockProjectName, []);

      // Assert
      expect(result).toEqual([]);
    });

    it('should handle error and return empty array', async () => {
      // Arrange
      (TFSServices.postRequest as jest.Mock).mockRejectedValueOnce(new Error('API Error'));

      // Act
      const result = await (resultDataProvider as any).fetchCrossTestPoints(mockProjectName, [1, 2]);

      // Assert
      expect(result).toEqual([]);
      expect(logger.error).toHaveBeenCalled();
    });

    it('should return empty array when API returns invalid response format', async () => {
      (TFSServices.postRequest as jest.Mock).mockResolvedValueOnce({ data: {} });

      const result = await (resultDataProvider as any).fetchCrossTestPoints(mockProjectName, [1, 2]);

      expect(result).toEqual([]);
      expect(logger.warn).toHaveBeenCalledWith('No test points found or invalid response format');
    });

    it('should pick the latest point per test case and map details with defaults', async () => {
      (TFSServices.postRequest as jest.Mock).mockResolvedValueOnce({
        data: {
          points: [
            {
              testCase: { id: 1 },
              lastTestRun: { id: '1' },
              lastResult: { id: '5' },
              url: 'https://example.com/points/1',
            },
            {
              testCase: { id: 1 },
              lastTestRun: { id: '2' },
              lastResult: { id: '1' },
              url: 'https://example.com/points/2',
            },
            {
              testCase: { id: 2 },
              lastTestRun: { id: '1' },
              lastResult: { id: '1' },
              url: 'https://example.com/points/3',
            },
          ],
        },
      });

      (TFSServices.getItemContent as jest.Mock)
        .mockResolvedValueOnce({
          testCase: { id: 1, name: 'TC 1' },
          testSuite: { id: 10, name: 'Suite' },
          configuration: { name: 'Config' },
          outcome: 'passed',
          lastTestRun: { id: '2' },
          lastResult: { id: '1' },
        })
        .mockResolvedValueOnce({
          testCase: { id: 2, name: 'TC 2' },
          testSuite: { id: 11, name: 'Suite' },
          configuration: { name: 'Config' },
          outcome: 'failed',
          lastTestRun: { id: '1' },
          lastResult: { id: '1' },
          lastResultDetails: { duration: 5, dateCompleted: '2023-01-01', runBy: { displayName: 'User' } },
        });

      const result = await (resultDataProvider as any).fetchCrossTestPoints(mockProjectName, [1, 2]);

      expect(TFSServices.getItemContent).toHaveBeenCalledWith(
        'https://example.com/points/2?witFields=Microsoft.VSTS.TCM.Steps&includePointDetails=true',
        mockToken
      );
      expect(result).toHaveLength(2);
      const tc1 = result.find((r: any) => r.testCaseId === 1);
      expect(tc1.lastResultDetails).toEqual(
        expect.objectContaining({
          duration: 0,
          runBy: expect.objectContaining({ displayName: 'No tester' }),
        })
      );
    });
  });

  describe('fetchLinkedWi', () => {
    it('should fetch linked work items and filter only open Bugs/Change Requests', async () => {
      const testItems = [
        {
          testId: 1,
          testName: 'TC 1',
          testCaseUrl: 'http://example.com/tc/1',
          runStatus: 'Passed',
        },
      ];

      (TFSServices.getItemContent as jest.Mock)
        .mockResolvedValueOnce({
          value: [
            {
              id: 1,
              relations: [
                { url: `${mockOrgUrl}_apis/wit/workItems/100` },
                { url: `${mockOrgUrl}_apis/wit/workItems/101` },
              ],
            },
          ],
        })
        .mockResolvedValueOnce({
          value: [
            {
              id: 100,
              fields: {
                'System.WorkItemType': 'Bug',
                'System.State': 'Active',
                'System.Title': 'Open bug',
                'Microsoft.VSTS.Common.Severity': '1 - Critical',
              },
            },
            {
              id: 101,
              fields: {
                'System.WorkItemType': 'Bug',
                'System.State': 'Closed',
                'System.Title': 'Closed bug',
              },
            },
          ],
        });

      const result = await (resultDataProvider as any).fetchLinkedWi(mockProjectName, testItems);

      expect(result).toHaveLength(1);
      expect(result[0].linkItems).toHaveLength(1);
      expect(result[0].linkItems[0]).toEqual(
        expect.objectContaining({
          pcrId: 100,
          workItemType: 'Bug',
          title: 'Open bug',
          pcrUrl: `${mockOrgUrl}${mockProjectName}/_workitems/edit/100`,
        })
      );
    });
  });

  describe('getTestReporterResults filtering', () => {
    it('should apply errorFilterMode=both and run-step filter and return the filtered rows', async () => {
      const testReporterRows = [
        {
          testCase: {
            comment: 'has comment',
            result: { resultMessage: 'Failed in Run 1' },
          },
          stepComments: 'step comment',
          stepStatus: 'Failed',
        },
        {
          testCase: {
            comment: '',
            result: { resultMessage: 'Passed' },
          },
          stepComments: '',
          stepStatus: 'Not Run',
        },
      ];

      jest.spyOn(resultDataProvider as any, 'fetchTestPlanName').mockResolvedValueOnce('Plan A');
      jest.spyOn(resultDataProvider as any, 'fetchTestSuites').mockResolvedValueOnce([]);
      jest.spyOn(resultDataProvider as any, 'fetchTestData').mockResolvedValueOnce([]);
      jest.spyOn(resultDataProvider as any, 'fetchAllResultDataTestReporter').mockResolvedValueOnce([]);
      jest
        .spyOn(resultDataProvider as any, 'alignStepsWithIterationsTestReporter')
        .mockReturnValueOnce(testReporterRows);

      const linkedQueryRequest = {
        linkedQueryMode: 'none',
        testAssociatedQuery: { wiql: { href: 'https://example.com/wiql' }, columns: [] },
      };

      const result = await resultDataProvider.getTestReporterResults(
        'planId',
        mockProjectName,
        [],
        [],
        false,
        true,
        true,
        linkedQueryRequest,
        'both'
      );

      expect(result).toBeDefined();
      const first = (result as any[])[0];
      expect(first).toBeDefined();
      expect(first.data).toHaveLength(1);
      expect(first.data[0].stepStatus).toBe('Failed');
    });

    it('should filter only test case results when errorFilterMode=onlyTestCaseResult', async () => {
      const rows = [
        {
          testCase: { comment: 'c', result: { resultMessage: 'Passed' } },
          stepComments: '',
          stepStatus: 'Not Run',
        },
        {
          testCase: { comment: '', result: { resultMessage: 'Failed in Run 1' } },
          stepComments: '',
          stepStatus: 'Not Run',
        },
        {
          testCase: { comment: '', result: { resultMessage: 'Passed' } },
          stepComments: '',
          stepStatus: 'Not Run',
        },
      ];

      jest.spyOn(resultDataProvider as any, 'fetchTestPlanName').mockResolvedValueOnce('Plan A');
      jest.spyOn(resultDataProvider as any, 'fetchTestSuites').mockResolvedValueOnce([]);
      jest.spyOn(resultDataProvider as any, 'fetchTestData').mockResolvedValueOnce([]);
      jest.spyOn(resultDataProvider as any, 'fetchAllResultDataTestReporter').mockResolvedValueOnce([]);
      jest.spyOn(resultDataProvider as any, 'alignStepsWithIterationsTestReporter').mockReturnValueOnce(rows);

      const res = await resultDataProvider.getTestReporterResults(
        'planId',
        mockProjectName,
        [],
        [],
        false,
        true,
        false,
        { linkedQueryMode: 'none', testAssociatedQuery: { wiql: { href: 'x' }, columns: [] } },
        'onlyTestCaseResult'
      );

      expect((res as any[])[0].data).toHaveLength(2);
    });

    it('should filter only step results when errorFilterMode=onlyTestStepsResult', async () => {
      const rows = [
        {
          testCase: { comment: '', result: { resultMessage: 'Passed' } },
          stepComments: 'x',
          stepStatus: 'Not Run',
        },
        {
          testCase: { comment: '', result: { resultMessage: 'Passed' } },
          stepComments: '',
          stepStatus: 'Failed',
        },
        {
          testCase: { comment: '', result: { resultMessage: 'Passed' } },
          stepComments: '',
          stepStatus: 'Not Run',
        },
      ];

      jest.spyOn(resultDataProvider as any, 'fetchTestPlanName').mockResolvedValueOnce('Plan A');
      jest.spyOn(resultDataProvider as any, 'fetchTestSuites').mockResolvedValueOnce([]);
      jest.spyOn(resultDataProvider as any, 'fetchTestData').mockResolvedValueOnce([]);
      jest.spyOn(resultDataProvider as any, 'fetchAllResultDataTestReporter').mockResolvedValueOnce([]);
      jest.spyOn(resultDataProvider as any, 'alignStepsWithIterationsTestReporter').mockReturnValueOnce(rows);

      const res = await resultDataProvider.getTestReporterResults(
        'planId',
        mockProjectName,
        [],
        [],
        false,
        true,
        false,
        { linkedQueryMode: 'none', testAssociatedQuery: { wiql: { href: 'x' }, columns: [] } },
        'onlyTestStepsResult'
      );

      expect((res as any[])[0].data).toHaveLength(2);
    });

    it('should execute query mode and call TicketsDataProvider when linkedQueryMode=query', async () => {
      jest.spyOn(resultDataProvider as any, 'fetchTestPlanName').mockResolvedValueOnce('Plan A');
      jest.spyOn(resultDataProvider as any, 'fetchTestSuites').mockResolvedValueOnce([]);
      jest.spyOn(resultDataProvider as any, 'fetchTestData').mockResolvedValueOnce([]);
      jest.spyOn(resultDataProvider as any, 'fetchAllResultDataTestReporter').mockResolvedValueOnce([]);
      jest.spyOn(resultDataProvider as any, 'alignStepsWithIterationsTestReporter').mockReturnValueOnce([]);

      const linkedQueryRequest = {
        linkedQueryMode: 'query',
        testAssociatedQuery: {
          wiql: { href: 'https://example.com/wiql' },
          columns: [{ referenceName: 'X', name: 'X' }],
        },
      };

      await resultDataProvider.getTestReporterResults(
        'planId',
        mockProjectName,
        [],
        [],
        false,
        true,
        false,
        linkedQueryRequest,
        'none'
      );

      const TicketsProviderMock: any = require('../../modules/TicketsDataProvider').default;
      expect(TicketsProviderMock).toHaveBeenCalled();
      const instance = TicketsProviderMock.mock.results[0].value;
      expect(instance.GetQueryResultsFromWiql).toHaveBeenCalledWith(
        'https://example.com/wiql',
        true,
        expect.any(Map)
      );
    });
  });

  describe('getMewpL2CoverageFlatResults', () => {
    it('should support query-mode requirement scope for MEWP coverage', async () => {
      jest.spyOn(resultDataProvider as any, 'fetchTestPlanName').mockResolvedValueOnce('Plan A');
      jest.spyOn(resultDataProvider as any, 'fetchTestSuites').mockResolvedValueOnce([{ testSuiteId: 1 }]);
      jest.spyOn(resultDataProvider as any, 'fetchTestData').mockResolvedValueOnce([
        {
          testPointsItems: [{ testCaseId: 101, lastRunId: 10, lastResultId: 20, testCaseName: 'TC 101' }],
          testCasesItems: [
            {
              workItem: {
                id: 101,
                workItemFields: [{ key: 'System.Title', value: 'TC 101' }],
              },
            },
          ],
        },
      ]);
      jest.spyOn(resultDataProvider as any, 'fetchMewpRequirementTypeNames').mockResolvedValueOnce([
        'Requirement',
      ]);
      jest.spyOn(resultDataProvider as any, 'fetchWorkItemsByIds').mockResolvedValueOnce([
        {
          id: 9001,
          fields: {
            'System.WorkItemType': 'Requirement',
            'System.Title': 'Requirement from query',
            'Custom.CustomerId': 'SR3001',
            'System.AreaPath': 'MEWP\\IL',
          },
          relations: [
            {
              rel: 'Microsoft.VSTS.Common.TestedBy-Forward',
              url: 'https://dev.azure.com/org/_apis/wit/workItems/101',
            },
          ],
        },
      ]);
      jest.spyOn(resultDataProvider as any, 'fetchAllResultDataTestReporter').mockResolvedValueOnce([
        {
          testCaseId: 101,
          testCase: { id: 101, name: 'TC 101' },
          iteration: {
            actionResults: [{ action: 'Validate SR3001', expected: '', outcome: 'Passed' }],
          },
        },
      ]);

      const TicketsProviderMock: any = require('../../modules/TicketsDataProvider').default;
      TicketsProviderMock.mockImplementationOnce(() => ({
        GetQueryResultsFromWiql: jest.fn().mockResolvedValue({
          fetchedWorkItems: [
            {
              id: 9001,
              fields: {
                'System.WorkItemType': 'Requirement',
                'System.Title': 'Requirement from query',
                'Custom.CustomerId': 'SR3001',
                'System.AreaPath': 'MEWP\\IL',
              },
            },
          ],
        }),
      }));

      const result = await (resultDataProvider as any).getMewpL2CoverageFlatResults(
        '123',
        mockProjectName,
        [1],
        {
          linkedQueryMode: 'query',
          testAssociatedQuery: { wiql: { href: 'https://example.com/wiql' } },
        }
      );

      const row = result.rows.find((item: any) => item['Customer ID'] === 'SR3001');
      expect(row).toEqual(
        expect.objectContaining({
          'Title (Customer name)': 'Requirement from query',
          'Responsibility - SAPWBS (ESUK/IL)': 'IL',
          'Test case id': 101,
          'Test case title': 'TC 101',
          'Number of passed steps': 1,
          'Number of failed steps': 0,
          'Number of not run tests': 0,
        })
      );

      expect(TicketsProviderMock).toHaveBeenCalled();
      const instance = TicketsProviderMock.mock.results[0].value;
      expect(instance.GetQueryResultsFromWiql).toHaveBeenCalledWith(
        'https://example.com/wiql',
        true,
        expect.any(Map)
      );
    });

    it('should map SR ids from steps and output requirement-test-case coverage rows', async () => {
      jest.spyOn(resultDataProvider as any, 'fetchTestPlanName').mockResolvedValueOnce('Plan A');
      jest.spyOn(resultDataProvider as any, 'fetchTestSuites').mockResolvedValueOnce([{ testSuiteId: 1 }]);
      jest.spyOn(resultDataProvider as any, 'fetchTestData').mockResolvedValueOnce([
        {
          testPointsItems: [
            { testCaseId: 101, lastRunId: 11, lastResultId: 22, testCaseName: 'TC 101' },
            { testCaseId: 102, lastRunId: 0, lastResultId: 0, testCaseName: 'TC 102' },
          ],
          testCasesItems: [
            {
              workItem: {
                id: 102,
                workItemFields: [{ key: 'Steps', value: '<steps></steps>' }],
              },
            },
          ],
        },
      ]);
      jest.spyOn(resultDataProvider as any, 'fetchMewpL2Requirements').mockResolvedValueOnce([
        {
          workItemId: 5001,
          requirementId: 'SR1001',
          title: 'Covered requirement',
          responsibility: 'ESUK',
          linkedTestCaseIds: [101],
        },
        {
          workItemId: 5002,
          requirementId: 'SR1002',
          title: 'Referenced from non-linked step text',
          responsibility: 'IL',
          linkedTestCaseIds: [],
        },
        {
          workItemId: 5003,
          requirementId: 'SR1003',
          title: 'Not covered by any test case',
          responsibility: 'IL',
          linkedTestCaseIds: [],
        },
      ]);
      jest.spyOn(resultDataProvider as any, 'fetchAllResultDataTestReporter').mockResolvedValueOnce([
        {
          testCaseId: 101,
          testCase: { id: 101, name: 'TC 101' },
          iteration: {
            actionResults: [
              {
                action: 'Validate <b>S</b><b>R</b> 1 0 0 1 happy path',
                expected: '',
                outcome: 'Passed',
              },
              { action: 'Validate SR1001 failed flow', expected: '&nbsp;', outcome: 'Failed' },
              { action: '', expected: 'Pending S R 1 0 0 1 scenario', outcome: 'Unspecified' },
            ],
          },
        },
        {
          testCaseId: 102,
          testCase: { id: 102, name: 'TC 102' },
          iteration: undefined,
        },
      ]);
      jest
        .spyOn((resultDataProvider as any).testStepParserHelper, 'parseTestSteps')
        .mockResolvedValueOnce([
          {
            stepId: '1',
            stepPosition: '1',
            action: 'Definition contains SR1002',
            expected: '',
            isSharedStepTitle: false,
          },
        ]);

      const result = await (resultDataProvider as any).getMewpL2CoverageFlatResults(
        '123',
        mockProjectName,
        [1]
      );

      expect(result).toEqual(
        expect.objectContaining({
          sheetName: expect.stringContaining('MEWP L2 Coverage'),
          columnOrder: expect.arrayContaining(['Customer ID', 'Test case id', 'Number of not run tests']),
        })
      );

      const covered = result.rows.find(
        (row: any) => row['Customer ID'] === 'SR1001' && row['Test case id'] === 101
      );
      const inferredByStepText = result.rows.find(
        (row: any) => row['Customer ID'] === 'SR1002' && row['Test case id'] === 102
      );
      const uncovered = result.rows.find(
        (row: any) =>
          row['Customer ID'] === 'SR1003' &&
          (row['Test case id'] === '' || row['Test case id'] === undefined || row['Test case id'] === null)
      );

      expect(covered).toEqual(
        expect.objectContaining({
          'Title (Customer name)': 'Covered requirement',
          'Responsibility - SAPWBS (ESUK/IL)': 'ESUK',
          'Test case title': 'TC 101',
          'Number of passed steps': 1,
          'Number of failed steps': 1,
          'Number of not run tests': 1,
        })
      );
      expect(inferredByStepText).toEqual(
        expect.objectContaining({
          'Title (Customer name)': 'Referenced from non-linked step text',
          'Responsibility - SAPWBS (ESUK/IL)': 'IL',
          'Test case title': 'TC 102',
          'Number of passed steps': 0,
          'Number of failed steps': 0,
          'Number of not run tests': 1,
        })
      );
      expect(uncovered).toEqual(
        expect.objectContaining({
          'Title (Customer name)': 'Not covered by any test case',
          'Responsibility - SAPWBS (ESUK/IL)': 'IL',
          'Test case title': '',
          'Number of passed steps': 0,
          'Number of failed steps': 0,
          'Number of not run tests': 0,
        })
      );
    });

    it('should extract SR ids from HTML/spacing and return unique ids per step text', () => {
      const text =
        'A: <b>S</b><b>R</b> 0 0 0 1; B: SR0002; C: S R 0 0 0 3; D: SR0002; E: &lt;b&gt;SR&lt;/b&gt;0004';
      const codes = (resultDataProvider as any).extractRequirementCodesFromText(text);
      expect([...codes].sort()).toEqual(['SR1', 'SR2', 'SR3', 'SR4']);
    });

    it('should not backfill definition steps as not-run when a real run exists but has no action results', async () => {
      jest.spyOn(resultDataProvider as any, 'fetchTestPlanName').mockResolvedValueOnce('Plan A');
      jest.spyOn(resultDataProvider as any, 'fetchTestSuites').mockResolvedValueOnce([{ testSuiteId: 1 }]);
      jest.spyOn(resultDataProvider as any, 'fetchTestData').mockResolvedValueOnce([
        {
          testPointsItems: [{ testCaseId: 101, lastRunId: 88, lastResultId: 99, testCaseName: 'TC 101' }],
          testCasesItems: [
            {
              workItem: {
                id: 101,
                workItemFields: [{ key: 'Steps', value: '<steps></steps>' }],
              },
            },
          ],
        },
      ]);
      jest.spyOn(resultDataProvider as any, 'fetchMewpL2Requirements').mockResolvedValueOnce([
        {
          workItemId: 7001,
          requirementId: 'SR2001',
          title: 'Has run but no actions',
          responsibility: 'ESUK',
          linkedTestCaseIds: [101],
        },
      ]);
      jest.spyOn(resultDataProvider as any, 'fetchAllResultDataTestReporter').mockResolvedValueOnce([
        {
          testCaseId: 101,
          lastRunId: 88,
          lastResultId: 99,
          iteration: {
            actionResults: [],
          },
        },
      ]);

      const parseSpy = jest
        .spyOn((resultDataProvider as any).testStepParserHelper, 'parseTestSteps')
        .mockResolvedValueOnce([
          {
            stepId: '1',
            stepPosition: '1',
            action: 'SR2001 from definition',
            expected: '',
            isSharedStepTitle: false,
          },
        ]);

      const result = await (resultDataProvider as any).getMewpL2CoverageFlatResults(
        '123',
        mockProjectName,
        [1]
      );

      const row = result.rows.find(
        (item: any) => item['Customer ID'] === 'SR2001' && item['Test case id'] === 101
      );
      expect(parseSpy).not.toHaveBeenCalled();
      expect(row).toEqual(
        expect.objectContaining({
          'Number of passed steps': 0,
          'Number of failed steps': 0,
          'Number of not run tests': 0,
        })
      );
    });

    it('should not infer requirement id from unrelated SR text in non-identifier fields', () => {
      const requirementId = (resultDataProvider as any).extractMewpRequirementIdentifier(
        {
          'System.Description': 'random text with SR9999 that is unrelated',
          'Custom.CustomerId': 'customer id unknown',
          'System.Title': 'Requirement without explicit SR code',
        },
        4321
      );

      expect(requirementId).toBe('4321');
    });
  });

  describe('fetchResultDataForTestReporter (runResultField switch)', () => {
    it('should populate requested runResultField values including testCaseResult URL branches', async () => {
      jest
        .spyOn(resultDataProvider as any, 'fetchResultDataBase')
        .mockImplementation(async (...args: any[]) => {
          const testSuiteId = args[1];
          const formatter = args[4];
          const extra = args[5] as any[];
          const selectedFields = extra[0];
          const isQueryMode = extra[1];
          const pt = extra[2];
          return formatter(
            {
              testCase: { id: 1, name: 'TC 1' },
              testSuite: { name: 'Suite' },
              testCaseRevision: 7,
              resolutionState: 'x',
              failureType: 'FT',
              priority: 2,
              outcome: 'passed',
              iterationDetails: [],
              filteredFields: { 'Custom.Field1': 'v1' },
              relatedRequirements: [],
              relatedBugs: [],
              relatedCRs: [],
            },
            testSuiteId,
            pt,
            selectedFields,
            isQueryMode
          );
        });

      const selectedFields = [
        'priority@runResultField',
        'testCaseResult@runResultField',
        'testCaseComment@runResultField',
        'failureType@runResultField',
        'runBy@runResultField',
        'executionDate@runResultField',
        'configurationName@runResultField',
        'unknownField@runResultField',
      ];

      const point = {
        lastRunId: 10,
        lastResultId: 20,
        configurationName: 'Cfg',
        lastResultDetails: { runBy: { displayName: 'User' }, dateCompleted: '2023-01-01' },
      };

      const res = await (resultDataProvider as any).fetchResultDataForTestReporter(
        mockProjectName,
        'suite1',
        point,
        selectedFields,
        false
      );

      expect(res.priority).toBe(2);
      expect(res.testCaseResult).toEqual(
        expect.objectContaining({
          resultMessage: expect.stringContaining('Run 10'),
          url: expect.stringContaining('runId=10'),
        })
      );
      expect(res.runBy).toBe('User');
      expect(res.executionDate).toBe('2023-01-01');
      expect(res.configurationName).toBe('Cfg');
      expect(res.customFields).toEqual(expect.objectContaining({ field1: 'v1' }));
    });

    it('should set testCaseResult url empty when lastRunId/lastResultId are undefined', async () => {
      jest
        .spyOn(resultDataProvider as any, 'fetchResultDataBase')
        .mockImplementation(async (...args: any[]) => {
          const testSuiteId = args[1];
          const point = args[2];
          const formatter = args[4];
          const extra = args[5] as any[];
          return formatter(
            {
              testCase: { id: 1, name: 'TC 1' },
              testSuite: { name: 'Suite' },
              testCaseRevision: 1,
              resolutionState: 'x',
              failureType: 'FT',
              priority: 1,
              outcome: 'passed',
              iterationDetails: [],
            },
            testSuiteId,
            point,
            extra[0]
          );
        });

      const res = await (resultDataProvider as any).fetchResultDataForTestReporter(
        mockProjectName,
        'suite1',
        {
          lastRunId: undefined,
          lastResultId: undefined,
          configurationName: 'Cfg',
          lastResultDetails: { runBy: { displayName: 'U' }, dateCompleted: 'd' },
        },
        ['testCaseResult@runResultField'],
        false
      );

      expect(res.testCaseResult).toEqual(expect.objectContaining({ url: '' }));
    });
  });

  describe('fetchResultDataBasedOnWiBase', () => {
    it('should return null and warn when runId/resultId are 0 and no point is provided', async () => {
      const res = await (resultDataProvider as any).fetchResultDataBasedOnWiBase(mockProjectName, '0', '0');
      expect(res).toBeNull();
      expect(logger.warn).toHaveBeenCalled();
    });

    it('should build synthetic result for Active state when runId/resultId are 0 and point is provided', async () => {
      const point = {
        testCaseId: '123',
        testCaseName: 'TC 123',
        outcome: 'passed',
        testSuite: { id: '1', name: 'Suite' },
      };

      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce({
        id: 123,
        rev: 7,
        fields: {
          'System.State': 'Active',
          'System.CreatedDate': '2023-01-01T00:00:00',
          'Microsoft.VSTS.TCM.Priority': 2,
          'System.Title': 'Title 123',
          'Microsoft.VSTS.TCM.Steps': '<steps></steps>',
        },
        relations: null,
      });

      const selectedFields = ['System.Title@testCaseWorkItemField'];
      const res = await (resultDataProvider as any).fetchResultDataBasedOnWiBase(
        mockProjectName,
        '0',
        '0',
        true,
        selectedFields,
        false,
        point
      );

      expect(res).toEqual(
        expect.objectContaining({
          id: 0,
          failureType: 'None',
          testCaseRevision: 7,
          stepsResultXml: '<steps></steps>',
          filteredFields: { 'System.Title': 'Title 123' },
        })
      );
    });

    it('should append linked relations and filter testCaseWorkItemField when isTestReporter=true and isQueryMode=false', async () => {
      (TFSServices.getItemContent as jest.Mock).mockReset();

      const selectedFields = ['associatedBug@linked', 'System.Title@testCaseWorkItemField'];

      // 1) run result
      (TFSServices.getItemContent as jest.Mock)
        .mockResolvedValueOnce({
          testCase: { id: 123 },
          testCaseRevision: 7,
          testSuite: { name: 'S' },
        })
        // 2) attachments
        .mockResolvedValueOnce({ value: [] })
        // 3) wiByRevision (with relations)
        .mockResolvedValueOnce({
          id: 123,
          fields: {
            'Microsoft.VSTS.TCM.Steps': '<steps></steps>',
            'System.Title': { displayName: 'My Title' },
          },
          relations: [{ rel: 'System.LinkTypes.Related', url: 'https://example.com/wi/200' }],
        })
        // 4) linked bug
        .mockResolvedValueOnce({
          id: 200,
          fields: { 'System.WorkItemType': 'Bug', 'System.State': 'Active', 'System.Title': 'B200' },
          _links: { html: { href: 'http://example.com/200' } },
        });

      const res = await (resultDataProvider as any).fetchResultDataBasedOnWiBase(
        mockProjectName,
        '10',
        '20',
        true,
        selectedFields,
        false,
        undefined,
        false
      );

      expect(res).toEqual(expect.objectContaining({ testCaseRevision: 7 }));
      expect(res.filteredFields).toEqual({ 'System.Title': 'My Title' });
      expect(res.relatedBugs).toEqual(
        expect.arrayContaining([expect.objectContaining({ id: 200, title: 'B200', workItemType: 'Bug' })])
      );
    });

    it('should return only the most recent System.History entry by default', async () => {
      (TFSServices.getItemContentWithHeaders as jest.Mock).mockReset();

      const selectedFields = ['System.History@testCaseWorkItemField'];

      (TFSServices.getItemContent as jest.Mock).mockReset();
      (TFSServices.getItemContent as jest.Mock)
        // 1) run result
        .mockResolvedValueOnce({
          testCase: { id: 123 },
          testCaseRevision: 7,
          testSuite: { name: 'S' },
        })
        // 2) attachments
        .mockResolvedValueOnce({ value: [] })
        // 3) wiByRevision (no System.History field on this revision)
        .mockResolvedValueOnce({
          id: 123,
          fields: { 'Microsoft.VSTS.TCM.Steps': '<steps></steps>' },
          relations: [],
        });

      // 4) comments (new shape: { totalCount, count, comments })
      (TFSServices.getItemContentWithHeaders as jest.Mock).mockResolvedValueOnce({
        data: {
          totalCount: 2,
          count: 2,
          comments: [
            {
              text: 'Second comment',
              createdDate: '2024-01-02T00:00:00Z',
              createdBy: { displayName: 'Bob' },
              isDeleted: false,
            },
            {
              text: 'First comment',
              createdDate: '2024-01-01T00:00:00Z',
              createdBy: { displayName: 'Alice' },
              isDeleted: false,
            },
          ],
        },
        headers: {},
      });

      const res = await (resultDataProvider as any).fetchResultDataBasedOnWiBase(
        mockProjectName,
        '10',
        '20',
        true,
        selectedFields,
        false
      );

      expect(res.filteredFields['System.History']).toEqual([
        { createdDate: '2024-01-02T00:00:00Z', createdBy: 'Bob', text: 'Second comment' },
      ]);
    });

    it('should backfill System.History from work item comments when requested (includeAllHistory=true)', async () => {
      (TFSServices.getItemContent as jest.Mock).mockReset();
      (TFSServices.getItemContentWithHeaders as jest.Mock).mockReset();

      const selectedFields = ['System.History@testCaseWorkItemField'];

      (TFSServices.getItemContent as jest.Mock)
        // 1) run result
        .mockResolvedValueOnce({
          testCase: { id: 123 },
          testCaseRevision: 7,
          testSuite: { name: 'S' },
        })
        // 2) attachments
        .mockResolvedValueOnce({ value: [] })
        // 3) wiByRevision (no System.History field on this revision)
        .mockResolvedValueOnce({
          id: 123,
          fields: { 'Microsoft.VSTS.TCM.Steps': '<steps></steps>' },
          relations: [],
        });

      (TFSServices.getItemContentWithHeaders as jest.Mock).mockResolvedValueOnce({
        data: {
          totalCount: 2,
          count: 2,
          comments: [
            {
              text: 'Second comment',
              createdDate: '2024-01-02T00:00:00Z',
              createdBy: { displayName: 'Bob' },
              isDeleted: false,
            },
            {
              text: 'First comment',
              createdDate: '2024-01-01T00:00:00Z',
              createdBy: { displayName: 'Alice' },
              isDeleted: false,
            },
          ],
        },
        headers: {},
      });

      const res = await (resultDataProvider as any).fetchResultDataBasedOnWiBase(
        mockProjectName,
        '10',
        '20',
        true,
        selectedFields,
        false,
        undefined,
        true
      );

      expect(res.filteredFields['System.History']).toEqual([
        { createdDate: '2024-01-02T00:00:00Z', createdBy: 'Bob', text: 'Second comment' },
        { createdDate: '2024-01-01T00:00:00Z', createdBy: 'Alice', text: 'First comment' },
      ]);
    });

    it('should paginate work item comments when continuationToken is returned in the response body', async () => {
      (TFSServices.getItemContent as jest.Mock).mockReset();
      (TFSServices.getItemContentWithHeaders as jest.Mock).mockReset();

      const selectedFields = ['System.History@testCaseWorkItemField'];

      (TFSServices.getItemContent as jest.Mock)
        // 1) run result
        .mockResolvedValueOnce({
          testCase: { id: 123 },
          testCaseRevision: 7,
          testSuite: { name: 'S' },
        })
        // 2) attachments
        .mockResolvedValueOnce({ value: [] })
        // 3) wiByRevision (no System.History field on this revision)
        .mockResolvedValueOnce({
          id: 123,
          fields: { 'Microsoft.VSTS.TCM.Steps': '<steps></steps>' },
          relations: [],
        });

      // 4) comments page 1 (continuationToken in body, not headers)
      (TFSServices.getItemContentWithHeaders as jest.Mock)
        .mockResolvedValueOnce({
          data: {
            totalCount: 3,
            count: 2,
            continuationToken: 'abc',
            comments: [
              {
                text: 'Second comment',
                createdDate: '2024-01-02T00:00:00Z',
                createdBy: { displayName: 'Bob' },
                isDeleted: false,
              },
              {
                text: 'First comment',
                createdDate: '2024-01-01T00:00:00Z',
                createdBy: { displayName: 'Alice' },
                isDeleted: false,
              },
            ],
          },
          headers: {},
        })
        // 5) comments page 2 (no continuation token)
        .mockResolvedValueOnce({
          data: {
            totalCount: 3,
            count: 1,
            comments: [
              {
                text: 'Third comment',
                createdDate: '2024-01-03T00:00:00Z',
                createdBy: { displayName: 'Cara' },
                isDeleted: false,
              },
            ],
          },
          headers: {},
        });

      const res = await (resultDataProvider as any).fetchResultDataBasedOnWiBase(
        mockProjectName,
        '10',
        '20',
        true,
        selectedFields,
        false,
        undefined,
        true
      );

      expect((TFSServices.getItemContentWithHeaders as jest.Mock).mock.calls).toHaveLength(2);
      expect((TFSServices.getItemContentWithHeaders as jest.Mock).mock.calls[1][0]).toContain(
        'continuationToken=abc'
      );

      expect(res.filteredFields['System.History']).toEqual([
        { createdDate: '2024-01-03T00:00:00Z', createdBy: 'Cara', text: 'Third comment' },
        { createdDate: '2024-01-02T00:00:00Z', createdBy: 'Bob', text: 'Second comment' },
        { createdDate: '2024-01-01T00:00:00Z', createdBy: 'Alice', text: 'First comment' },
      ]);
    });

    it('should return empty System.History when comments endpoint fails', async () => {
      (TFSServices.getItemContent as jest.Mock).mockReset();
      (TFSServices.getItemContentWithHeaders as jest.Mock).mockReset();

      const selectedFields = ['System.History@testCaseWorkItemField'];

      (TFSServices.getItemContent as jest.Mock)
        // 1) run result
        .mockResolvedValueOnce({
          testCase: { id: 123 },
          testCaseRevision: 7,
          testSuite: { name: 'S' },
        })
        // 2) attachments
        .mockResolvedValueOnce({ value: [] })
        // 3) wiByRevision
        .mockResolvedValueOnce({
          id: 123,
          fields: { 'Microsoft.VSTS.TCM.Steps': '<steps></steps>' },
          relations: [],
        })
        ;

      // 4) comments (fails)
      (TFSServices.getItemContentWithHeaders as jest.Mock).mockRejectedValueOnce(new Error('comments failed'));

      const res = await (resultDataProvider as any).fetchResultDataBasedOnWiBase(
        mockProjectName,
        '10',
        '20',
        true,
        selectedFields,
        false
      );

      expect(res.filteredFields['System.History']).toEqual([]);
    });

  });

  describe('alignStepsWithIterationsBase - additional branches', () => {
    it('should include not-run test cases when enabled and includeItemsWithNoIterations=true (creates test-level result)', () => {
      const testData = [
        {
          testPointsItems: [{ testCaseId: 123, lastRunId: undefined, lastResultId: undefined }],
          testCasesItems: [
            { workItem: { id: 123, workItemFields: [{ key: 'Steps', value: '<steps></steps>' }] } },
          ],
        },
      ];
      const iterations = [
        { testCaseId: 123, lastRunId: undefined, lastResultId: undefined, iteration: null },
      ];

      const createResultObject = jest.fn().mockReturnValue({ ok: true });
      const shouldProcessStepLevel = jest.fn().mockReturnValue(false);

      const res = (resultDataProvider as any).alignStepsWithIterationsBase(
        testData,
        iterations,
        true,
        true,
        true,
        {
          selectedFields: [],
          createResultObject,
          shouldProcessStepLevel,
        }
      );

      expect(res).toEqual([{ ok: true }]);
      expect(createResultObject).toHaveBeenCalled();
    });

    it('should skip items without iterations when includeItemsWithNoIterations=false', () => {
      const testData = [
        {
          testPointsItems: [{ testCaseId: 123, lastRunId: undefined, lastResultId: undefined }],
          testCasesItems: [
            { workItem: { id: 123, workItemFields: [{ key: 'Steps', value: '<steps></steps>' }] } },
          ],
        },
      ];
      const iterations = [
        { testCaseId: 123, lastRunId: undefined, lastResultId: undefined, iteration: null },
      ];

      const res = (resultDataProvider as any).alignStepsWithIterationsBase(
        testData,
        iterations,
        true,
        false,
        true,
        {
          selectedFields: [],
          createResultObject: jest.fn().mockReturnValue({ ok: true }),
          shouldProcessStepLevel: jest.fn().mockReturnValue(false),
        }
      );

      expect(res).toEqual([]);
    });

    it('should fallback to test-level result when step-level processing enabled but actionResults is empty', () => {
      const testData = [
        {
          testPointsItems: [{ testCaseId: 123, lastRunId: '10', lastResultId: '20' }],
          testCasesItems: [
            { workItem: { id: 123, workItemFields: [{ key: 'Steps', value: '<steps></steps>' }] } },
          ],
        },
      ];
      const iterations = [
        {
          testCaseId: 123,
          lastRunId: '10',
          lastResultId: '20',
          iteration: { actionResults: [] },
        },
      ];

      const createResultObject = jest.fn().mockReturnValue({ mode: 'test-level' });
      const shouldProcessStepLevel = jest.fn().mockReturnValue(true);

      const res = (resultDataProvider as any).alignStepsWithIterationsBase(
        testData,
        iterations,
        false,
        true,
        false,
        {
          selectedFields: ['includeSteps@stepsRunProperties'],
          createResultObject,
          shouldProcessStepLevel,
        }
      );

      expect(res).toEqual([{ mode: 'test-level' }]);
      expect(createResultObject).toHaveBeenCalledTimes(1);
    });
  });

  describe('getCombinedResultsSummary - optional outputs', () => {
    it('should include open PCRs, test log, appendix-a and appendix-b when enabled', async () => {
      jest
        .spyOn(resultDataProvider as any, 'fetchTestSuites')
        .mockResolvedValueOnce([{ testSuiteId: '1', testGroupName: 'Group 1' }]);

      jest.spyOn(resultDataProvider as any, 'fetchTestPoints').mockResolvedValueOnce([
        {
          testCaseId: 1,
          testCaseName: 'TC 1',
          testCaseUrl: 'http://example.com/1',
          configurationName: 'Cfg',
          outcome: 'passed',
          lastRunId: 10,
          lastResultId: 20,
          lastResultDetails: { dateCompleted: '2023-01-01T00:00:00.000Z', runBy: { displayName: 'User 1' } },
        },
      ]);

      jest.spyOn(resultDataProvider as any, 'fetchTestData').mockResolvedValueOnce([]);

      // runResults for appendix-a filtering
      jest.spyOn(resultDataProvider as any, 'fetchAllResultData').mockResolvedValueOnce([
        {
          comment: 'has comment',
          iteration: { attachments: [] },
          analysisAttachments: [],
          lastRunId: 1,
          lastResultId: 2,
        },
      ]);

      jest.spyOn(resultDataProvider as any, 'alignStepsWithIterations').mockReturnValueOnce([]);

      const openSpy = jest
        .spyOn(resultDataProvider as any, 'fetchOpenPcrData')
        .mockResolvedValueOnce(undefined);

      // Ensure appendix-b mapping runs but doesn't require attachments
      jest
        .spyOn(resultDataProvider as any, 'mapStepResultsForExecutionAppendix')
        .mockReturnValueOnce(new Map());

      const res = await resultDataProvider.getCombinedResultsSummary(
        mockTestPlanId,
        mockProjectName,
        undefined,
        false,
        false,
        { openPcrMode: 'linked', openPcrLinkedQuery: { wiql: { href: 'x' }, columns: [] } } as any,
        true,
        { isEnabled: true, generateAttachments: { isEnabled: true, runAttachmentMode: 'planOnly' } },
        { isEnabled: true, generateRunAttachments: { isEnabled: false } },
        false
      );

      expect(openSpy).toHaveBeenCalled();

      const hasTestLog = res.combinedResults.some(
        (x: any) => x.contentControl === 'test-execution-content-control'
      );
      expect(hasTestLog).toBe(true);

      const hasAppendixA = res.combinedResults.some(
        (x: any) => x.contentControl === 'appendix-a-content-control'
      );
      expect(hasAppendixA).toBe(true);

      const hasAppendixB = res.combinedResults.some(
        (x: any) => x.contentControl === 'appendix-b-content-control'
      );
      expect(hasAppendixB).toBe(true);
    });
  });

  describe('appendLinkedRelations', () => {
    it('should append requirement/bug/cr when enabled and skip closed items', async () => {
      const relations = [
        { rel: 'System.LinkTypes.Related', url: 'https://example.com/wi/1' },
        { rel: 'System.LinkTypes.Related', url: 'https://example.com/wi/2' },
        { rel: 'System.LinkTypes.Related', url: 'https://example.com/wi/3' },
        { rel: 'System.LinkTypes.Related', url: 'https://example.com/wi/4' },
      ];
      const relatedRequirements: any[] = [];
      const relatedBugs: any[] = [];
      const relatedCRs: any[] = [];
      const selected = new Set(['associatedRequirement', 'associatedBug', 'associatedCR']);

      (TFSServices.getItemContent as jest.Mock)
        .mockResolvedValueOnce({
          id: 1,
          fields: {
            'System.WorkItemType': 'Requirement',
            'System.State': 'Active',
            'System.Title': 'Req 1',
            'Custom.CustomerRequirementId': 'CUST-1',
          },
          _links: { html: { href: 'http://example.com/1' } },
        })
        .mockResolvedValueOnce({
          id: 2,
          fields: {
            'System.WorkItemType': 'Bug',
            'System.State': 'Active',
            'System.Title': 'Bug 2',
          },
          _links: { html: { href: 'http://example.com/2' } },
        })
        .mockResolvedValueOnce({
          id: 3,
          fields: {
            'System.WorkItemType': 'Change Request',
            'System.State': 'Active',
            'System.Title': 'CR 3',
          },
          _links: { html: { href: 'http://example.com/3' } },
        })
        .mockResolvedValueOnce({
          id: 4,
          fields: {
            'System.WorkItemType': 'Bug',
            'System.State': 'Closed',
            'System.Title': 'Closed bug',
          },
          _links: { html: { href: 'http://example.com/4' } },
        });

      await (resultDataProvider as any).appendLinkedRelations(
        relations,
        relatedRequirements,
        relatedBugs,
        relatedCRs,
        { id: 123 },
        selected
      );

      expect(relatedRequirements).toHaveLength(1);
      expect(relatedRequirements[0]).toEqual(
        expect.objectContaining({ id: 1, customerId: 'CUST-1', workItemType: 'Requirement' })
      );
      expect(relatedBugs).toHaveLength(1);
      expect(relatedCRs).toHaveLength(1);
    });

    it('should log an error when fetching a related item fails', async () => {
      const relations = [{ rel: 'System.LinkTypes.Related', url: 'https://example.com/wi/1' }];
      (TFSServices.getItemContent as jest.Mock).mockRejectedValueOnce(new Error('network'));

      await (resultDataProvider as any).appendLinkedRelations(
        relations,
        [],
        [],
        [],
        { id: 999 },
        new Set(['associatedBug'])
      );

      expect(logger.error).toHaveBeenCalledWith(
        expect.stringContaining('Could not append related work item to test case 999')
      );
    });
  });

  describe('appendQueryRelations', () => {
    it('should append query relations when map has items', () => {
      // Arrange
      const testCaseId = 1;
      const relatedRequirements: any[] = [];
      const relatedBugs: any[] = [];
      const relatedCRs: any[] = [];

      // Set up the map
      (resultDataProvider as any).testToAssociatedItemMap = new Map([
        [
          1,
          [
            {
              id: 100,
              fields: { 'System.Title': 'Req 1', 'System.WorkItemType': 'Requirement' },
              _links: { html: { href: 'http://example.com/100' } },
            },
          ],
        ],
      ]);
      (resultDataProvider as any).querySelectedColumns = [];

      // Act
      (resultDataProvider as any).appendQueryRelations(
        testCaseId,
        relatedRequirements,
        relatedBugs,
        relatedCRs
      );

      // Assert
      expect(relatedRequirements).toHaveLength(1);
      expect(relatedRequirements[0].id).toBe(100);
    });

    it('should map Requirement/Bug/Change Request into correct buckets', () => {
      const testCaseId = 1;
      const relatedRequirements: any[] = [];
      const relatedBugs: any[] = [];
      const relatedCRs: any[] = [];

      jest.spyOn(resultDataProvider as any, 'standardCustomField').mockReturnValue({});

      (resultDataProvider as any).testToAssociatedItemMap = new Map([
        [
          1,
          [
            {
              id: 10,
              fields: { 'System.Title': 'R', 'System.WorkItemType': 'Requirement', X: '1' },
              _links: { html: { href: 'http://example.com/10' } },
            },
            {
              id: 11,
              fields: { 'System.Title': 'B', 'System.WorkItemType': 'Bug', Y: '2' },
              _links: { html: { href: 'http://example.com/11' } },
            },
            {
              id: 12,
              fields: { 'System.Title': 'C', 'System.WorkItemType': 'Change Request', Z: '3' },
              _links: { html: { href: 'http://example.com/12' } },
            },
          ],
        ],
      ]);
      (resultDataProvider as any).querySelectedColumns = [];

      (resultDataProvider as any).appendQueryRelations(
        testCaseId,
        relatedRequirements,
        relatedBugs,
        relatedCRs
      );

      expect(relatedRequirements).toHaveLength(1);
      expect(relatedBugs).toHaveLength(1);
      expect(relatedCRs).toHaveLength(1);
    });

    it('should handle empty map', () => {
      // Arrange
      const testCaseId = 1;
      const relatedRequirements: any[] = [];
      const relatedBugs: any[] = [];
      const relatedCRs: any[] = [];
      (resultDataProvider as any).testToAssociatedItemMap = new Map();

      // Act
      (resultDataProvider as any).appendQueryRelations(
        testCaseId,
        relatedRequirements,
        relatedBugs,
        relatedCRs
      );

      // Assert
      expect(relatedRequirements).toHaveLength(0);
    });
  });

  describe('convertUnspecifiedRunStatus', () => {
    it('should return empty string for null actionResult', () => {
      // Act
      const result = (resultDataProvider as any).convertUnspecifiedRunStatus(null);

      // Assert
      expect(result).toBe('');
    });

    it('should return empty string for Unspecified shared step title', () => {
      // Arrange
      const actionResult = { outcome: 'Unspecified', isSharedStepTitle: true };

      // Act
      const result = (resultDataProvider as any).convertUnspecifiedRunStatus(actionResult);

      // Assert
      expect(result).toBe('');
    });

    it('should return Not Run for Unspecified non-shared step', () => {
      // Arrange
      const actionResult = { outcome: 'Unspecified', isSharedStepTitle: false };

      // Act
      const result = (resultDataProvider as any).convertUnspecifiedRunStatus(actionResult);

      // Assert
      expect(result).toBe('Not Run');
    });

    it('should return original outcome for non-Unspecified status', () => {
      // Arrange
      const actionResult = { outcome: 'Passed', isSharedStepTitle: false };

      // Act
      const result = (resultDataProvider as any).convertUnspecifiedRunStatus(actionResult);

      // Assert
      expect(result).toBe('Passed');
    });
  });

  describe('fetchResultDataBasedOnWi', () => {
    it('should call fetchResultDataBasedOnWiBase', async () => {
      // Arrange
      const mockResult = { id: 1, outcome: 'passed' };
      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockResult);

      // Act - just verify it doesn't throw
      const spy = jest.spyOn(resultDataProvider as any, 'fetchResultDataBasedOnWiBase');
      try {
        await (resultDataProvider as any).fetchResultDataBasedOnWi(mockProjectName, '100', '200');
      } catch {
        // Expected to fail due to missing mocks
      }

      // Assert
      expect(spy).toHaveBeenCalledWith(mockProjectName, '100', '200');
    });
  });

  describe('alignStepsWithIterationsBase', () => {
    it('should return empty array when no iterations', () => {
      // Arrange
      const testData: any[] = [];
      const iterations: any[] = [];
      const options = {
        createResultObject: jest.fn(),
        shouldProcessStepLevel: jest.fn(),
      };

      // Act
      const result = (resultDataProvider as any).alignStepsWithIterationsBase(
        testData,
        iterations,
        false,
        false,
        false,
        options
      );

      // Assert
      expect(result).toEqual([]);
    });

    it('should return [null] when fetchedTestCase.iteration.actionResults is null (shouldProcessStepLevel=false)', () => {
      const testData = [
        {
          testGroupName: 'G',
          testPointsItems: [{ testCaseId: 1, testCaseName: 'TC', lastRunId: 10, lastResultId: 20 }],
          testCasesItems: [{ workItem: { id: 1, workItemFields: [{ key: 'Steps', value: '<steps />' }] } }],
        },
      ];
      const iterations = [
        {
          testCaseId: 1,
          lastRunId: 10,
          lastResultId: 20,
          iteration: { actionResults: null },
          testCaseRevision: 1,
        },
      ];

      const res = (resultDataProvider as any).alignStepsWithIterations(testData, iterations);
      expect(res).toEqual([null]);
    });

    it('should include a null row when actionResults contains an undefined element', () => {
      const testData = [
        {
          testGroupName: 'G',
          testPointsItems: [{ testCaseId: 1, testCaseName: 'TC', lastRunId: 10, lastResultId: 20 }],
          testCasesItems: [{ workItem: { id: 1, workItemFields: [{ key: 'Steps', value: '<steps />' }] } }],
        },
      ];
      const iterations = [
        {
          testCaseId: 1,
          lastRunId: 10,
          lastResultId: 20,
          iteration: { actionResults: [undefined] },
          testCaseRevision: 1,
        },
      ];

      const res = (resultDataProvider as any).alignStepsWithIterations(testData, iterations);
      expect(res).toEqual([null]);
    });
  });

  describe('standardCustomField', () => {
    it('should standardize custom fields with columns', () => {
      // Arrange
      const fields = { 'Custom.Field1': 'value1', 'Custom.Field2': 'value2' };
      const columns = [
        { referenceName: 'Custom.Field1', name: 'Field 1' },
        { referenceName: 'Custom.Field2', name: 'Field 2' },
      ];

      // Act
      const result = (resultDataProvider as any).standardCustomField(fields, columns);

      // Assert
      expect(result).toBeDefined();
      expect(result.field1).toBe('value1');
    });

    it('should handle uppercase field names', () => {
      // Arrange
      const fields = { 'Custom.ABC': 'value1' };
      const columns = [{ referenceName: 'Custom.ABC', name: 'ABC' }];

      // Act
      const result = (resultDataProvider as any).standardCustomField(fields, columns);

      // Assert
      expect(result.abc).toBe('value1');
    });

    it('should skip standard fields', () => {
      // Arrange
      const fields = { 'System.Id': 1, 'Custom.Field1': 'value1' };
      const columns = [
        { referenceName: 'System.Id', name: 'id' },
        { referenceName: 'Custom.Field1', name: 'Field 1' },
      ];

      // Act
      const result = (resultDataProvider as any).standardCustomField(fields, columns);

      // Assert
      expect(result.id).toBeUndefined();
      expect(result.field1).toBe('value1');
    });

    it('should handle null/undefined field values', () => {
      // Arrange
      const fields = { 'Custom.Field1': null };
      const columns = [{ referenceName: 'Custom.Field1', name: 'Field 1' }];

      // Act
      const result = (resultDataProvider as any).standardCustomField(fields, columns);

      // Assert
      expect(result.field1).toBeNull();
    });

    it('should handle fields without columns', () => {
      // Arrange
      const fields = { 'Custom.Field1': 'value1', 'System.Title': 'Title' };

      // Act
      const result = (resultDataProvider as any).standardCustomField(fields);

      // Assert
      expect(result).toBeDefined();
    });

    it('should handle displayName property', () => {
      // Arrange
      const fields = { 'Custom.Field1': { displayName: 'Display Value' } };
      const columns = [{ referenceName: 'Custom.Field1', name: 'Field 1' }];

      // Act
      const result = (resultDataProvider as any).standardCustomField(fields, columns);

      // Assert
      expect(result.field1).toBe('Display Value');
    });
  });

  describe('getTestOutcome', () => {
    it('should return outcome from last iteration', () => {
      // Arrange
      const resultData = {
        iterationDetails: [{ outcome: 'Failed' }, { outcome: 'Passed' }],
        outcome: 'Failed',
      };

      // Act
      const result = (resultDataProvider as any).getTestOutcome(resultData);

      // Assert
      expect(result).toBe('Passed');
    });

    it('should return result outcome when no iteration details', () => {
      // Arrange
      const resultData = { outcome: 'Passed' };

      // Act
      const result = (resultDataProvider as any).getTestOutcome(resultData);

      // Assert
      expect(result).toBe('Passed');
    });

    it('should return default outcome when no data', () => {
      // Arrange
      const resultData = {};

      // Act
      const result = (resultDataProvider as any).getTestOutcome(resultData);

      // Assert
      expect(result).toBe('NotApplicable');
    });
  });

  describe('createIterationsMap', () => {
    it('should create iterations map from results', () => {
      // Arrange
      const iterations = [{ lastRunId: 100, lastResultId: 200, testCase: { id: 1 } }];

      // Act
      const result = (resultDataProvider as any).createIterationsMap(iterations, false, false);

      // Assert
      expect(result).toBeDefined();
    });

    it('should create iterations map from results with iteration', () => {
      // Arrange
      const iterations = [{ lastRunId: 100, lastResultId: 200, testCaseId: 1, iteration: { id: 1 } }];

      // Act
      const result = (resultDataProvider as any).createIterationsMap(iterations, false, false);

      // Assert
      expect(result).toBeDefined();
      expect(result['100-200-1']).toBeDefined();
    });

    it('should create iterations map for test reporter mode', () => {
      // Arrange
      const iterations = [{ lastRunId: 100, lastResultId: 200, testCaseId: 1 }];

      // Act
      const result = (resultDataProvider as any).createIterationsMap(iterations, true, false);

      // Assert
      expect(result['100-200-1']).toBeDefined();
    });

    it('should include not run test cases when flag is set', () => {
      // Arrange
      const iterations = [{ testCaseId: 1 }];

      // Act
      const result = (resultDataProvider as any).createIterationsMap(iterations, false, true);

      // Assert
      expect(result['1']).toBeDefined();
    });
  });

  describe('alignStepsWithIterationsBase', () => {
    it('should return empty array when no iterations', () => {
      // Arrange
      const testData: any[] = [];
      const iterations: any[] = [];
      const options = {
        createResultObject: jest.fn(),
        shouldProcessStepLevel: jest.fn(),
      };

      // Act
      const result = (resultDataProvider as any).alignStepsWithIterationsBase(
        testData,
        iterations,
        false,
        false,
        false,
        options
      );

      // Assert
      expect(result).toEqual([]);
    });
  });

  describe('alignStepsWithIterations', () => {
    it('should return empty array when no iterations', () => {
      // Arrange
      const testData: any[] = [];
      const iterations: any[] = [];

      // Act
      const result = (resultDataProvider as any).alignStepsWithIterations(testData, iterations);

      // Assert
      expect(result).toEqual([]);
    });
  });

  describe('fetchTestData', () => {
    it('should fetch test data for suites', async () => {
      // Arrange
      const suites = [{ testSuiteId: 1, testGroupName: 'Suite 1' }];
      const mockTestCases = { value: [{ workItem: { id: 1 } }] };
      const mockTestPoints = { value: [{ testCaseReference: { id: 1 } }], count: 1 };

      (TFSServices.getItemContent as jest.Mock)
        .mockResolvedValueOnce(mockTestCases)
        .mockResolvedValueOnce(mockTestPoints);

      // Act
      const result = await (resultDataProvider as any).fetchTestData(
        suites,
        mockProjectName,
        mockTestPlanId,
        false
      );

      // Assert
      expect(result).toHaveLength(1);
      expect(result[0].testCasesItems).toBeDefined();
    });

    it('should handle errors gracefully', async () => {
      // Arrange
      const suites = [{ testSuiteId: 1, testGroupName: 'Suite 1' }];
      (TFSServices.getItemContent as jest.Mock).mockRejectedValueOnce(new Error('API Error'));

      // Act
      const result = await (resultDataProvider as any).fetchTestData(
        suites,
        mockProjectName,
        mockTestPlanId,
        false
      );

      // Assert
      expect(result).toHaveLength(1);
      expect(logger.error).toHaveBeenCalled();
    });
  });

  describe('fetchAllResultData', () => {
    it('should return empty array when no test data', async () => {
      // Arrange
      const testData: any[] = [];

      // Act
      const result = await (resultDataProvider as any).fetchAllResultData(testData, mockProjectName);

      // Assert
      expect(result).toEqual([]);
    });

    it('should fetch result data for test points', async () => {
      // Arrange
      const testData = [
        {
          testPointsItems: [{ testCaseId: 1, lastRunId: 100, lastResultId: 200 }],
        },
      ];
      const mockResult = {
        testCase: { id: 1 },
        iteration: { actionResults: [] },
      };
      const mockAttachments = { value: [] };
      const mockWi = { fields: {} };

      (TFSServices.getItemContent as jest.Mock)
        .mockResolvedValueOnce(mockResult)
        .mockResolvedValueOnce(mockAttachments)
        .mockResolvedValueOnce(mockWi);

      // Act
      const result = await (resultDataProvider as any).fetchAllResultData(testData, mockProjectName);

      // Assert
      expect(result).toBeDefined();
    });

    it('should log response data and rethrow when an error with response is thrown', async () => {
      const err: any = new Error('boom');
      err.response = { data: { detail: 'bad' } };

      jest.spyOn(resultDataProvider as any, 'fetchTestSuites').mockRejectedValueOnce(err);

      await expect(
        resultDataProvider.getCombinedResultsSummary(mockTestPlanId, mockProjectName)
      ).rejects.toThrow('boom');

      expect(logger.error).toHaveBeenCalledWith('Error during getCombinedResultsSummary: boom');
      expect(logger.error).toHaveBeenCalledWith('Response Data: {"detail":"bad"}');
    });
  });

  describe('fetchAllResultDataTestReporter', () => {
    it('should return empty array when no test data', async () => {
      // Arrange
      const testData: any[] = [];

      // Act
      const result = await (resultDataProvider as any).fetchAllResultDataTestReporter(
        testData,
        mockProjectName,
        [],
        false
      );

      // Assert
      expect(result).toEqual([]);
    });
  });

  describe('alignStepsWithIterationsTestReporter', () => {
    it('should return empty array when no iterations', () => {
      // Arrange
      const testData: any[] = [];
      const iterations: any[] = [];

      // Act
      const result = (resultDataProvider as any).alignStepsWithIterationsTestReporter(
        testData,
        iterations,
        [],
        false
      );

      // Assert
      expect(result).toEqual([]);
    });
  });

  describe('fetchAllResultDataBase', () => {
    it('should return empty array when no test data', async () => {
      // Arrange
      const testData: any[] = [];
      const fetchStrategy = jest.fn();

      // Act
      const result = await (resultDataProvider as any).fetchAllResultDataBase(
        testData,
        mockProjectName,
        false,
        fetchStrategy
      );

      // Assert
      expect(result).toEqual([]);
    });

    it('should filter out points without run/result IDs when not test reporter', async () => {
      // Arrange
      const testData = [
        {
          testSuiteId: 1,
          testPointsItems: [
            { testCaseId: 1 }, // No lastRunId/lastResultId
          ],
        },
      ];
      const fetchStrategy = jest.fn();

      // Act
      const result = await (resultDataProvider as any).fetchAllResultDataBase(
        testData,
        mockProjectName,
        false,
        fetchStrategy
      );

      // Assert
      expect(fetchStrategy).not.toHaveBeenCalled();
      expect(result).toEqual([]);
    });
  });

  describe('fetchResultDataBase', () => {
    it('should fetch result data for a point', async () => {
      // Arrange
      const point = { lastRunId: 100, lastResultId: 200 };
      const mockResultData = {
        testCase: { id: 1 },
        iterationDetails: [],
      };
      const fetchResultMethod = jest.fn().mockResolvedValue(mockResultData);
      const createResponseObject = jest.fn().mockReturnValue({ id: 1 });

      // Act
      const result = await (resultDataProvider as any).fetchResultDataBase(
        mockProjectName,
        '1',
        point,
        fetchResultMethod,
        createResponseObject
      );

      // Assert
      expect(fetchResultMethod).toHaveBeenCalled();
      expect(result).toBeDefined();
    });
  });

  describe('getCombinedResultsSummary', () => {
    it('should return combined results summary with expected content controls', async () => {
      const mockTestSuites = {
        value: [
          {
            id: 1,
            name: 'Root Suite',
            children: [{ id: 2, name: 'Child Suite 1', parentSuite: { id: 1 } }],
          },
        ],
        count: 1,
      };

      const mockTestPoints = {
        value: [
          {
            testCaseReference: { id: 1, name: 'Test Case 1' },
            configuration: { name: 'Config 1' },
            results: {
              outcome: 'passed',
              lastTestRunId: 100,
              lastResultId: 200,
              lastResultDetails: { dateCompleted: '2023-01-01', runBy: { displayName: 'Test User' } },
            },
          },
        ],
        count: 1,
      };

      const mockTestCases = {
        value: [
          {
            workItem: {
              id: 1,
              workItemFields: [{ key: 'Steps', value: '<steps>...</steps>' }],
            },
          },
        ],
      };

      const mockResult = {
        testCase: { id: 1, name: 'Test Case 1' },
        testSuite: { id: 2, name: 'Child Suite 1' },
        iterationDetails: [
          {
            actionResults: [
              { stepIdentifier: '1', outcome: 'Passed', errorMessage: '', actionPath: 'path1' },
            ],
            attachments: [],
          },
        ],
        testCaseRevision: 1,
        failureType: null,
        resolutionState: null,
        comment: null,
      };

      // Setup mocks for API calls
      (TFSServices.getItemContent as jest.Mock)
        .mockResolvedValueOnce(mockTestSuites) // fetchTestSuites
        .mockResolvedValueOnce(mockTestPoints) // fetchTestPoints
        .mockResolvedValueOnce(mockTestCases) // fetchTestCasesBySuiteId
        .mockResolvedValueOnce(mockResult) // fetchResult
        .mockResolvedValueOnce({ value: [] }) // fetchResult - attachments
        .mockResolvedValueOnce({ fields: {} }); // fetchResult - wiByRevision

      const mockTestStepParserHelper = (resultDataProvider as any).testStepParserHelper;
      mockTestStepParserHelper.parseTestSteps.mockResolvedValueOnce([
        {
          stepId: 1,
          stepPosition: '1',
          action: 'Do something',
          expected: 'Something happens',
          isSharedStepTitle: false,
        },
      ]);

      // Act
      const result = await resultDataProvider.getCombinedResultsSummary(
        mockTestPlanId,
        mockProjectName,
        undefined,
        true
      );

      // Assert
      expect(result.combinedResults.length).toBeGreaterThan(0);
      expect(result.combinedResults[0]).toHaveProperty(
        'contentControl',
        'test-group-summary-content-control'
      );
      expect(result.combinedResults[1]).toHaveProperty(
        'contentControl',
        'test-result-summary-content-control'
      );
      expect(result.combinedResults[2]).toHaveProperty(
        'contentControl',
        'detailed-test-result-content-control'
      );
    });
  });

  describe('fetchResultDataBase - shared step mapping', () => {
    it('should map parsed steps into actionResults, filter missing stepPosition, and sort by stepPosition', async () => {
      const point = { testCaseId: 1, lastRunId: 10, lastResultId: 20 };
      const fetchResultMethod = jest.fn().mockResolvedValue({
        testCase: { id: 1, name: 'TC' },
        stepsResultXml: '<steps></steps>',
        iterationDetails: [
          {
            actionResults: [
              { stepIdentifier: '2', actionPath: 'p2', sharedStepModel: { id: 5, revision: 7 } },
              { stepIdentifier: '999', actionPath: 'px' },
              { stepIdentifier: '1', actionPath: 'p1' },
            ],
          },
        ],
      });

      const createResponseObject = (resultData: any) => ({ iteration: resultData.iterationDetails[0] });

      const helper = (resultDataProvider as any).testStepParserHelper;
      helper.parseTestSteps.mockImplementationOnce(async (_xml: any, map: Map<number, number>) => {
        // cover sharedStepIdToRevisionLookupMap population
        expect(map.get(5)).toBe(7);
        return [
          { stepId: 1, stepPosition: '1', action: 'A1', expected: 'E1', isSharedStepTitle: false },
          { stepId: 2, stepPosition: '2', action: 'A2', expected: 'E2', isSharedStepTitle: true },
        ];
      });

      const res = await (resultDataProvider as any).fetchResultDataBase(
        mockProjectName,
        'suite1',
        point,
        fetchResultMethod,
        createResponseObject,
        []
      );

      const actionResults = res.iteration.actionResults;
      // 999 should be filtered out (no stepPosition)
      expect(actionResults).toHaveLength(2);
      // sorted by stepPosition numeric
      expect(actionResults[0]).toEqual(expect.objectContaining({ stepIdentifier: '1', action: 'A1' }));
      expect(actionResults[1]).toEqual(
        expect.objectContaining({ stepIdentifier: '2', action: 'A2', isSharedStepTitle: true })
      );
    });

    it('should fall back to parsed test steps when actionResults are missing for latest run', async () => {
      const point = { testCaseId: 1, lastRunId: 10, lastResultId: 20 };
      const fetchResultMethod = jest.fn().mockResolvedValue({
        testCase: { id: 1, name: 'TC' },
        stepsResultXml: '<steps></steps>',
        iterationDetails: [{}],
      });

      const createResponseObject = (resultData: any) => ({ iteration: resultData.iterationDetails[0] });

      const helper = (resultDataProvider as any).testStepParserHelper;
      helper.parseTestSteps.mockResolvedValueOnce([
        { stepId: 172, stepPosition: '1', action: 'Step 1', expected: 'Expected 1', isSharedStepTitle: false },
        { stepId: 173, stepPosition: '2', action: 'Step 2', expected: 'Expected 2', isSharedStepTitle: false },
      ]);

      const res = await (resultDataProvider as any).fetchResultDataBase(
        mockProjectName,
        'suite1',
        point,
        fetchResultMethod,
        createResponseObject,
        []
      );

      const actionResults = res.iteration.actionResults;
      expect(actionResults).toHaveLength(2);
      expect(actionResults[0]).toEqual(
        expect.objectContaining({
          stepIdentifier: '172',
          stepPosition: '1',
          action: 'Step 1',
          expected: 'Expected 1',
          outcome: 'Unspecified',
        })
      );
      expect(actionResults[1]).toEqual(
        expect.objectContaining({
          stepIdentifier: '173',
          stepPosition: '2',
          action: 'Step 2',
          expected: 'Expected 2',
          outcome: 'Unspecified',
        })
      );
    });

    it('should create synthetic iteration and fall back to parsed test steps when iterationDetails are missing', async () => {
      const point = { testCaseId: 1, lastRunId: 10, lastResultId: 20 };
      const fetchResultMethod = jest.fn().mockResolvedValue({
        testCase: { id: 1, name: 'TC' },
        stepsResultXml: '<steps></steps>',
        iterationDetails: [],
      });

      const createResponseObject = (resultData: any) => ({ iteration: resultData.iterationDetails[0] });

      const helper = (resultDataProvider as any).testStepParserHelper;
      helper.parseTestSteps.mockResolvedValueOnce([
        { stepId: 301, stepPosition: '1', action: 'S1', expected: 'E1', isSharedStepTitle: false },
      ]);

      const res = await (resultDataProvider as any).fetchResultDataBase(
        mockProjectName,
        'suite1',
        point,
        fetchResultMethod,
        createResponseObject,
        []
      );

      expect(res.iteration).toBeDefined();
      expect(res.iteration.actionResults).toHaveLength(1);
      expect(res.iteration.actionResults[0]).toEqual(
        expect.objectContaining({
          stepIdentifier: '301',
          stepPosition: '1',
          action: 'S1',
          expected: 'E1',
          outcome: 'Unspecified',
        })
      );
    });

    it('should fall back to parsed test steps when actionResults is an empty array', async () => {
      const point = { testCaseId: 1, lastRunId: 10, lastResultId: 20 };
      const fetchResultMethod = jest.fn().mockResolvedValue({
        testCase: { id: 1, name: 'TC' },
        stepsResultXml: '<steps></steps>',
        iterationDetails: [{ actionResults: [] }],
      });

      const createResponseObject = (resultData: any) => ({ iteration: resultData.iterationDetails[0] });

      const helper = (resultDataProvider as any).testStepParserHelper;
      helper.parseTestSteps.mockResolvedValueOnce([
        { stepId: 11, stepPosition: '1', action: 'A1', expected: 'E1', isSharedStepTitle: false },
        { stepId: 22, stepPosition: '2', action: 'A2', expected: 'E2', isSharedStepTitle: false },
      ]);

      const res = await (resultDataProvider as any).fetchResultDataBase(
        mockProjectName,
        'suite1',
        point,
        fetchResultMethod,
        createResponseObject,
        []
      );

      expect(helper.parseTestSteps).toHaveBeenCalledTimes(1);
      expect(res.iteration.actionResults).toHaveLength(2);
      expect(res.iteration.actionResults[0]).toEqual(
        expect.objectContaining({
          stepIdentifier: '11',
          stepPosition: '1',
          action: 'A1',
          expected: 'E1',
          outcome: 'Unspecified',
        })
      );
      expect(res.iteration.actionResults[1]).toEqual(
        expect.objectContaining({
          stepIdentifier: '22',
          stepPosition: '2',
          action: 'A2',
          expected: 'E2',
          outcome: 'Unspecified',
        })
      );
    });
  });

  describe('getTestReporterFlatResults', () => {
    it('should return flat rows with logical step numbering for fallback-generated action results', async () => {
      jest.spyOn(resultDataProvider as any, 'fetchTestPlanName').mockResolvedValueOnce('Plan 12');
      jest.spyOn(resultDataProvider as any, 'fetchTestSuites').mockResolvedValueOnce([
        {
          testSuiteId: 200,
          suiteId: 200,
          suiteName: 'suite 2.1',
          parentSuiteId: 100,
          parentSuiteName: 'Rel2',
          suitePath: 'Root/Rel2/suite 2.1',
          testGroupName: 'suite 2.1',
        },
      ]);

      const testData = [
        {
          testSuiteId: 200,
          suiteId: 200,
          suiteName: 'suite 2.1',
          parentSuiteId: 100,
          parentSuiteName: 'Rel2',
          suitePath: 'Root/Rel2/suite 2.1',
          testGroupName: 'suite 2.1',
          testPointsItems: [
            {
              testCaseId: 17,
              testCaseName: 'TC 17',
              outcome: 'Unspecified',
              lastRunId: 99,
              lastResultId: 88,
              testPointId: 501,
              lastResultDetails: {
                dateCompleted: '2026-02-01T10:00:00.000Z',
                outcome: 'Unspecified',
                runBy: { displayName: 'tester user' },
              },
            },
          ],
          testCasesItems: [
            {
              workItem: {
                id: 17,
                workItemFields: [{ key: 'Steps', value: '<steps></steps>' }],
              },
            },
          ],
        },
      ];
      jest.spyOn(resultDataProvider as any, 'fetchTestData').mockResolvedValueOnce(testData);
      jest.spyOn(resultDataProvider as any, 'fetchAllResultDataTestReporter').mockResolvedValueOnce([
        {
          testCaseId: 17,
          lastRunId: 99,
          lastResultId: 88,
          executionDate: '2026-02-01T10:00:00.000Z',
          testCaseResult: { resultMessage: '' },
          customFields: { 'Custom.SubSystem': 'SYS' },
          runBy: 'tester user',
          iteration: {
            actionResults: [
              {
                stepIdentifier: '172',
                stepPosition: '1',
                actionPath: '1',
                action: 'fallback action',
                expected: 'fallback expected',
                outcome: 'Unspecified',
                errorMessage: '',
                isSharedStepTitle: false,
              },
            ],
          },
        },
      ]);

      const result = await resultDataProvider.getTestReporterFlatResults(
        mockTestPlanId,
        mockProjectName,
        undefined,
        [],
        false
      );

      expect(result.planId).toBe(mockTestPlanId);
      expect(result.planName).toBe('Plan 12');
      expect(result.rows).toHaveLength(1);
      expect(result.rows[0]).toEqual(
        expect.objectContaining({
          planId: mockTestPlanId,
          planName: 'Plan 12',
          suiteId: 200,
          suiteName: 'suite 2.1',
          parentSuiteId: 100,
          parentSuiteName: 'Rel2',
          suitePath: 'Root/Rel2/suite 2.1',
          testCaseId: 17,
          testCaseName: 'TC 17',
          testRunId: 99,
          testPointId: 501,
          stepOutcome: 'Unspecified',
          stepStepIdentifier: '1',
        })
      );
    });
  });

  describe('getCombinedResultsSummary - appendix branches', () => {
    it('should use mapAttachmentsUrl when stepAnalysis.generateRunAttachments is enabled and stepExecution.runAttachmentMode != planOnly', async () => {
      jest
        .spyOn(resultDataProvider as any, 'fetchTestSuites')
        .mockResolvedValueOnce([{ testSuiteId: '1', testGroupName: 'Group 1' }]);

      jest.spyOn(resultDataProvider as any, 'fetchTestPoints').mockResolvedValueOnce([
        {
          testCaseId: 1,
          testCaseName: 'TC 1',
          testCaseUrl: 'http://example.com/1',
          configurationName: 'Cfg',
          outcome: 'passed',
          lastRunId: 10,
          lastResultId: 20,
          lastResultDetails: { dateCompleted: '2023-01-01T00:00:00.000Z', runBy: { displayName: 'User 1' } },
        },
      ]);

      jest.spyOn(resultDataProvider as any, 'fetchTestData').mockResolvedValueOnce([]);

      // ensure all predicates in stepAnalysis/filter can be true (comment false, attachments true, analysisAttachments true)
      const runResults = [
        {
          comment: '',
          iteration: { attachments: [{ actionPath: 'p', name: 'n', downloadUrl: 'd' }] },
          analysisAttachments: [{ id: 1 }],
        },
      ];
      jest.spyOn(resultDataProvider as any, 'fetchAllResultData').mockResolvedValueOnce(runResults);
      jest.spyOn(resultDataProvider as any, 'alignStepsWithIterations').mockReturnValueOnce([]);

      const mapSpy = jest
        .spyOn(resultDataProvider as any, 'mapAttachmentsUrl')
        .mockReturnValueOnce(runResults as any)
        .mockReturnValueOnce(runResults as any);

      jest
        .spyOn(resultDataProvider as any, 'mapStepResultsForExecutionAppendix')
        .mockReturnValueOnce(new Map());

      const res = await resultDataProvider.getCombinedResultsSummary(
        mockTestPlanId,
        mockProjectName,
        undefined,
        false,
        false,
        null,
        false,
        { isEnabled: true, generateAttachments: { isEnabled: true, runAttachmentMode: 'runOnly' } },
        { isEnabled: true, generateRunAttachments: { isEnabled: true } },
        false
      );

      expect(mapSpy).toHaveBeenCalled();
      expect(res.combinedResults.some((x: any) => x.contentControl === 'appendix-a-content-control')).toBe(
        true
      );
      expect(res.combinedResults.some((x: any) => x.contentControl === 'appendix-b-content-control')).toBe(
        true
      );
    });
  });

  describe('alignStepsWithIterationsTestReporter - step-level rows', () => {
    it('should emit step-level fields when includeSteps/stepRunStatus/testStepComment are selected', () => {
      const testData = [
        {
          testGroupName: 'G',
          testPointsItems: [
            {
              testCaseId: 123,
              testCaseName: 'TC',
              testCaseUrl: 'u',
              lastRunId: 10,
              lastResultId: 20,
            },
          ],
          testCasesItems: [
            { workItem: { id: 123, workItemFields: [{ key: 'Steps', value: '<steps></steps>' }] } },
          ],
        },
      ];
      const iterations = [
        {
          testCaseId: 123,
          lastRunId: 10,
          lastResultId: 20,
          iteration: {
            actionResults: [
              {
                stepIdentifier: '1',
                stepPosition: '1',
                action: 'A',
                expected: 'E',
                outcome: 'Unspecified',
                isSharedStepTitle: false,
                errorMessage: 'err',
              },
            ],
          },
          testCaseResult: 'Failed',
          comment: 'c',
          runBy: { displayName: 'u' },
          failureType: 'ft',
          executionDate: 'd',
          configurationName: 'cfg',
          relatedRequirements: [],
          relatedBugs: [],
          relatedCRs: [],
          customFields: {},
        },
      ];

      const res = (resultDataProvider as any).alignStepsWithIterationsTestReporter(
        testData,
        iterations,
        [
          'includeSteps@stepsRunProperties',
          'stepRunStatus@stepsRunProperties',
          'testStepComment@stepsRunProperties',
        ],
        true
      );

      expect(res).toHaveLength(1);
      expect(res[0]).toEqual(
        expect.objectContaining({
          stepNo: '1',
          stepAction: 'A',
          stepExpected: 'E',
          stepStatus: 'Not Run',
          stepComments: 'err',
        })
      );
    });

    it('should omit step fields when no @stepsRunProperties are selected', () => {
      const testData = [
        {
          testGroupName: 'G',
          testPointsItems: [
            { testCaseId: 123, testCaseName: 'TC', testCaseUrl: 'u', lastRunId: 10, lastResultId: 20 },
          ],
          testCasesItems: [
            { workItem: { id: 123, workItemFields: [{ key: 'Steps', value: '<steps></steps>' }] } },
          ],
        },
      ];
      const iterations = [
        {
          testCaseId: 123,
          lastRunId: 10,
          lastResultId: 20,
          iteration: {
            actionResults: [{ stepIdentifier: '1', stepPosition: '1', action: 'A', expected: 'E' }],
          },
          testCaseResult: 'Passed',
          comment: '',
          runBy: { displayName: 'u' },
          failureType: '',
          executionDate: 'd',
          configurationName: 'cfg',
          relatedRequirements: [],
          relatedBugs: [],
          relatedCRs: [],
          customFields: {},
        },
      ];

      const res = (resultDataProvider as any).alignStepsWithIterationsTestReporter(
        testData,
        iterations,
        [],
        true
      );
      expect(res).toHaveLength(1);
      expect(res[0].stepNo).toBeUndefined();
    });
  });

  describe('fetchOpenPcrData', () => {
    it('should populate both trace maps using linked work items', async () => {
      const testItems = [
        {
          testId: 1,
          testName: 'T1',
          testCaseUrl: 'u1',
          runStatus: 'Passed',
        },
      ];
      const linked = [
        {
          testId: 1,
          testName: 'T1',
          testCaseUrl: 'u1',
          runStatus: 'Passed',
          linkItems: [
            {
              pcrId: 10,
              workItemType: 'Bug',
              title: 'B10',
              severity: '2',
              pcrUrl: 'p10',
            },
          ],
        },
      ];

      jest.spyOn(resultDataProvider as any, 'fetchLinkedWi').mockResolvedValueOnce(linked);

      const openPcrToTestCaseTraceMap = new Map<string, string[]>();
      const testCaseToOpenPcrTraceMap = new Map<string, string[]>();

      await (resultDataProvider as any).fetchOpenPcrData(
        testItems,
        mockProjectName,
        openPcrToTestCaseTraceMap,
        testCaseToOpenPcrTraceMap
      );

      expect(openPcrToTestCaseTraceMap.size).toBe(1);
      expect(testCaseToOpenPcrTraceMap.size).toBe(1);
    });
  });
});
