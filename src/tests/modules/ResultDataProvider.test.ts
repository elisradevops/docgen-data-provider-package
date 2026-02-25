import { TFSServices } from '../../helpers/tfs';
import ResultDataProvider from '../../modules/ResultDataProvider';
import logger from '../../utils/logger';
import Utils from '../../utils/testStepParserHelper';
import axios from 'axios';
import * as XLSX from 'xlsx';

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
  const buildWorkbookBuffer = (rows: any[][]): Buffer => {
    const worksheet = XLSX.utils.aoa_to_sheet(rows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
    return XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' }) as Buffer;
  };

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

      it('should include pointAsOfTimestamp when lastUpdatedDate is available', () => {
        const testPoint = {
          testCaseReference: { id: 1, name: 'Test Case 1' },
          lastUpdatedDate: '2025-01-01T12:34:56Z',
        };

        const result = (resultDataProvider as any).mapTestPoint(testPoint, mockProjectName);

        expect(result).toEqual(
          expect.objectContaining({
            pointAsOfTimestamp: '2025-01-01T12:34:56.000Z',
          })
        );
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
      expect(TFSServices.getItemContent).toHaveBeenCalledWith(
        expect.stringContaining(
          `/_apis/testplan/Plans/${mockTestPlanId}/Suites/${mockSuiteId}/TestCase?witFields=Microsoft.VSTS.TCM.Steps,System.Rev`
        ),
        mockToken
      );
    });
  });

  describe('resolveSuiteTestCaseRevision', () => {
    it('should resolve System.Rev from workItemFields', () => {
      const revision = (resultDataProvider as any).resolveSuiteTestCaseRevision({
        workItem: {
          workItemFields: [{ key: 'System.Rev', value: '12' }],
        },
      });

      expect(revision).toBe(12);
    });

    it('should resolve System.Rev case-insensitively from workItem fields map', () => {
      const revision = (resultDataProvider as any).resolveSuiteTestCaseRevision({
        workItem: {
          fields: { 'system.rev': 14 },
        },
      });

      expect(revision).toBe(14);
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
        'https://example.com/points/2?witFields=Microsoft.VSTS.TCM.Steps,System.Rev&includePointDetails=true',
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
    it('should fetch MEWP scoped test data from selected suites', async () => {
      jest.spyOn(resultDataProvider as any, 'fetchTestPlanName').mockResolvedValueOnce('Plan A');
      const scopedSpy = jest
        .spyOn(resultDataProvider as any, 'fetchMewpScopedTestData')
        .mockResolvedValueOnce([]);
      jest.spyOn(resultDataProvider as any, 'fetchMewpL2Requirements').mockResolvedValueOnce([]);

      await (resultDataProvider as any).getMewpL2CoverageFlatResults(
        '123',
        mockProjectName,
        [1],
        undefined
      );

      expect(scopedSpy).toHaveBeenCalledWith('123', mockProjectName, [1]);
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
          baseKey: 'SR1001',
          areaPath: 'MEWP\\Customer Requirements\\Level 2',
          title: 'Covered requirement',
          responsibility: 'ESUK',
          linkedTestCaseIds: [101],
        },
        {
          workItemId: 5002,
          requirementId: 'SR1002',
          baseKey: 'SR1002',
          areaPath: 'MEWP\\Customer Requirements\\Level 2',
          title: 'Referenced from non-linked step text',
          responsibility: 'IL',
          linkedTestCaseIds: [],
        },
        {
          workItemId: 5003,
          requirementId: 'SR1003',
          baseKey: 'SR1003',
          areaPath: 'MEWP\\Customer Requirements\\Level 2',
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
                expected: '<b>S</b><b>R</b> 1 0 0 1',
                outcome: 'Passed',
              },
              { action: 'Validate SR1001 failed flow', expected: 'SR1001', outcome: 'Failed' },
              { action: '', expected: 'S R 1 0 0 1', outcome: 'Unspecified' },
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
            expected: 'SR1002',
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
          columnOrder: expect.arrayContaining(['L2 REQ ID', 'L2 REQ Title', 'L2 Run Status']),
        })
      );

      const covered = result.rows.find((row: any) => row['L2 REQ ID'] === '5001');
      const inferredByStepText = result.rows.find((row: any) => row['L2 REQ ID'] === '5002');
      const uncovered = result.rows.find((row: any) => row['L2 REQ ID'] === '5003');

      expect(covered).toEqual(
        expect.objectContaining({
          'L2 REQ Title': 'Covered requirement',
          'L2 SubSystem': '',
          'L2 Run Status': 'Fail',
          'Bug ID': '',
          'L3 REQ ID': '',
          'L4 REQ ID': '',
        })
      );
      expect(inferredByStepText).toEqual(
        expect.objectContaining({
          'L2 REQ Title': 'Referenced from non-linked step text',
          'L2 SubSystem': '',
          'L2 Run Status': 'Not Run',
        })
      );
      expect(uncovered).toEqual(
        expect.objectContaining({
          'L2 REQ Title': 'Not covered by any test case',
          'L2 SubSystem': '',
          'L2 Run Status': 'Not Run',
        })
      );
    });

    it('should extract SR ids from HTML/spacing and return unique ids per step text', () => {
      const text =
        '<b>S</b><b>R</b> 0 0 0 1; SR0002; S R 0 0 0 3; SR0002; &lt;b&gt;SR&lt;/b&gt;0004';
      const codes = (resultDataProvider as any).extractRequirementCodesFromText(text);
      expect([...codes].sort()).toEqual(['SR0001', 'SR0002', 'SR0003', 'SR0004']);
    });

    it('should keep only clean SR tokens and ignore noisy version/VVRM fragments', () => {
      const extract = (text: string) =>
        [...((resultDataProvider as any).extractRequirementCodesFromText(text) as Set<string>)].sort();

      expect(extract('SR12413; SR24513; SR25135 VVRM2425')).toEqual(['SR12413', 'SR24513']);
      expect(extract('SR12413; SR12412; SR12413-V3.24')).toEqual(['SR12412', 'SR12413']);
      expect(extract('SR12413; SR12412; SR12413 V3.24')).toEqual(['SR12412', 'SR12413']);
      expect(extract('SR12413, SR12412, SR12413-V3.24')).toEqual(['SR12412', 'SR12413']);

      const extractWithSuffix = (text: string) =>
        [
          ...((resultDataProvider as any).extractRequirementCodesFromExpectedText(
            text,
            true
          ) as Set<string>),
        ].sort();
      expect(extractWithSuffix('SR0095-2,3; SR0100-1,2,3,4')).toEqual([
        'SR0095-2',
        'SR0095-3',
        'SR0100-1',
        'SR0100-2',
        'SR0100-3',
        'SR0100-4',
      ]);
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
          baseKey: 'SR2001',
          areaPath: 'MEWP\\Customer Requirements\\Level 2',
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

      const row = result.rows.find((item: any) => item['L2 REQ ID'] === '7001');
      expect(parseSpy).not.toHaveBeenCalled();
      expect(row).toEqual(
        expect.objectContaining({
          'L2 Run Status': 'Not Run',
        })
      );
    });

    it('should not infer requirement id from unrelated SR text in non-identifier fields', () => {
      const requirementId = (resultDataProvider as any).extractMewpRequirementIdentifier(
        {
          'System.Description': 'random text with SR9999 that is unrelated',
          'Custom.CustomerId': 'customer id unknown',
          'System.Title': 'Requirement without explicit SR code',
        }
      );

      expect(requirementId).toBe('');
    });

    it('should derive responsibility from Custom.SAPWBS when present', () => {
      const responsibility = (resultDataProvider as any).deriveMewpResponsibility({
        'Custom.SAPWBS': 'IL',
        'System.AreaPath': 'MEWP\\ESUK',
      });

      expect(responsibility).toBe('IL');
    });

    it('should derive responsibility from AreaPath suffix ATP/ATP\\\\ESUK when SAPWBS is empty', () => {
      const esuk = (resultDataProvider as any).deriveMewpResponsibility({
        'Custom.SAPWBS': '',
        'System.AreaPath': 'MEWP\\Customer Requirements\\Level 2\\ATP\\ESUK',
      });
      const il = (resultDataProvider as any).deriveMewpResponsibility({
        'Custom.SAPWBS': '',
        'System.AreaPath': 'MEWP\\Customer Requirements\\Level 2\\ATP',
      });

      expect(esuk).toBe('ESUK');
      expect(il).toBe('IL');
    });

    it('should derive responsibility from Area Path alias when System.AreaPath is missing', () => {
      const esuk = (resultDataProvider as any).deriveMewpResponsibility({
        'Custom.SAPWBS': '',
        'Area Path': 'MEWP\\Customer Requirements\\Level 2\\ATP\\ESUK',
      });
      const il = (resultDataProvider as any).deriveMewpResponsibility({
        'Custom.SAPWBS': '',
        'Area Path': 'MEWP\\Customer Requirements\\Level 2\\ATP',
      });

      expect(esuk).toBe('ESUK');
      expect(il).toBe('IL');
    });

    it('should derive test-case responsibility from testCasesItems area-path fields', async () => {
      const fetchByIdsSpy = jest
        .spyOn(resultDataProvider as any, 'fetchWorkItemsByIds')
        .mockResolvedValue([]);

      const map = await (resultDataProvider as any).buildMewpTestCaseResponsibilityMap(
        [
          {
            testCasesItems: [
              {
                workItem: {
                  id: 101,
                  workItemFields: [{ name: 'Area Path', value: 'MEWP\\Customer Requirements\\Level 2\\ATP' }],
                },
              },
              {
                workItem: {
                  id: 102,
                  workItemFields: [
                    { referenceName: 'System.AreaPath', value: 'MEWP\\Customer Requirements\\Level 2\\ATP\\ESUK' },
                  ],
                },
              },
            ],
            testPointsItems: [],
          },
        ],
        mockProjectName
      );

      expect(map.get(101)).toBe('IL');
      expect(map.get(102)).toBe('ESUK');
      expect(fetchByIdsSpy).not.toHaveBeenCalled();
    });

    it('should derive test-case responsibility from AreaPath even when SAPWBS exists on test-case payload', async () => {
      jest.spyOn(resultDataProvider as any, 'fetchWorkItemsByIds').mockResolvedValue([]);

      const map = await (resultDataProvider as any).buildMewpTestCaseResponsibilityMap(
        [
          {
            testCasesItems: [
              {
                workItem: {
                  id: 201,
                  workItemFields: [
                    { referenceName: 'Custom.SAPWBS', value: 'ESUK' },
                    { referenceName: 'System.AreaPath', value: 'MEWP\\Customer Requirements\\Level 2\\ATP' },
                  ],
                },
              },
            ],
            testPointsItems: [],
          },
        ],
        mockProjectName
      );

      expect(map.get(201)).toBe('IL');
    });

    it('should zip bug rows with L3/L4 pairs and avoid cross-product duplication', () => {
      const requirements = [
        {
          requirementId: 'SR5303',
          baseKey: 'SR5303',
          title: 'Req 5303',
          subSystem: 'Power',
          responsibility: 'ESUK',
          linkedTestCaseIds: [101],
        },
      ];

      const requirementIndex = new Map([
        [
          'SR5303',
          new Map([
            [
              101,
              {
                passed: 0,
                failed: 1,
                notRun: 0,
              },
            ],
          ]),
        ],
      ]);

      const observedTestCaseIdsByRequirement = new Map<string, Set<number>>([
        ['SR5303', new Set([101])],
      ]);

      const linkedRequirementsByTestCase = new Map([
        [
          101,
          {
            baseKeys: new Set(['SR5303']),
            fullCodes: new Set(['SR5303']),
            bugIds: new Set([10003, 20003]),
          },
        ],
      ]);

      const externalBugsByTestCase = new Map([
        [
          101,
          [
            { id: 10003, title: 'Bug 10003', responsibility: 'Elisra', requirementBaseKey: 'SR5303' },
            { id: 20003, title: 'Bug 20003', responsibility: 'ESUK', requirementBaseKey: 'SR5303' },
          ],
        ],
      ]);

      const l3l4ByBaseKey = new Map([
        [
          'SR5303',
          [
            { l3Id: '9003', l3Title: 'L3 9003', l4Id: '9103', l4Title: 'L4 9103' },
          ],
        ],
      ]);

      const rows = (resultDataProvider as any).buildMewpCoverageRows(
        requirements,
        requirementIndex,
        observedTestCaseIdsByRequirement,
        linkedRequirementsByTestCase,
        l3l4ByBaseKey,
        externalBugsByTestCase
      );

      expect(rows).toHaveLength(2);
      expect(rows.map((row: any) => row['Bug ID'])).toEqual([10003, 20003]);
      expect(rows[0]).toEqual(
        expect.objectContaining({
          'L3 REQ ID': '9003',
          'L4 REQ ID': '9103',
        })
      );
      expect(rows[1]).toEqual(
        expect.objectContaining({
          'L3 REQ ID': '',
          'L4 REQ ID': '',
        })
      );
    });

    it('should drop standalone L3 row when same L3 has one or more L4 links', () => {
      const requirements = [
        {
          requirementId: 'SR5310',
          baseKey: 'SR5310',
          title: 'Req 5310',
          subSystem: 'Power',
          responsibility: 'ESUK',
          linkedTestCaseIds: [111],
        },
      ];

      const requirementIndex = new Map([
        [
          'SR5310',
          new Map([
            [
              111,
              {
                passed: 1,
                failed: 0,
                notRun: 0,
              },
            ],
          ]),
        ],
      ]);

      const observedTestCaseIdsByRequirement = new Map<string, Set<number>>([
        ['SR5310', new Set([111])],
      ]);

      const linkedRequirementsByTestCase = new Map([
        [
          111,
          {
            baseKeys: new Set(['SR5310']),
            fullCodes: new Set(['SR5310']),
            bugIds: new Set(),
          },
        ],
      ]);

      const l3l4ByBaseKey = new Map([
        [
          'SR5310',
          [
            { l3Id: '9003', l3Title: 'L3 9003', l4Id: '', l4Title: '' },
            { l3Id: '9003', l3Title: 'L3 9003', l4Id: '9103', l4Title: 'L4 9103' },
            { l3Id: '9003', l3Title: 'L3 9003', l4Id: '9104', l4Title: 'L4 9104' },
          ],
        ],
      ]);

      const rows = (resultDataProvider as any).buildMewpCoverageRows(
        requirements,
        requirementIndex,
        observedTestCaseIdsByRequirement,
        linkedRequirementsByTestCase,
        l3l4ByBaseKey,
        new Map()
      );

      expect(rows).toHaveLength(2);
      expect(rows[0]).toEqual(
        expect.objectContaining({
          'L3 REQ ID': '9003',
          'L4 REQ ID': '9103',
        })
      );
      expect(rows[1]).toEqual(
        expect.objectContaining({
          'L3 REQ ID': '9003',
          'L4 REQ ID': '9104',
        })
      );
    });

    it('should not emit bug rows from ADO-linked bug ids when external bugs source is empty', () => {
      const requirements = [
        {
          requirementId: 'SR5304',
          baseKey: 'SR5304',
          title: 'Req 5304',
          subSystem: 'Power',
          responsibility: 'ESUK',
          linkedTestCaseIds: [101],
        },
      ];
      const requirementIndex = new Map([
        [
          'SR5304',
          new Map([
            [
              101,
              {
                passed: 0,
                failed: 1,
                notRun: 0,
              },
            ],
          ]),
        ],
      ]);
      const observedTestCaseIdsByRequirement = new Map<string, Set<number>>([
        ['SR5304', new Set([101])],
      ]);
      const linkedRequirementsByTestCase = new Map([
        [
          101,
          {
            baseKeys: new Set(['SR5304']),
            fullCodes: new Set(['SR5304']),
            bugIds: new Set([55555]), // must be ignored in MEWP coverage mode
          },
        ],
      ]);

      const rows = (resultDataProvider as any).buildMewpCoverageRows(
        requirements,
        requirementIndex,
        observedTestCaseIdsByRequirement,
        linkedRequirementsByTestCase,
        new Map(),
        new Map()
      );

      expect(rows).toHaveLength(1);
      expect(rows[0]['L2 Run Status']).toBe('Fail');
      expect(rows[0]['Bug ID']).toBe('');
      expect(rows[0]['Bug Title']).toBe('');
      expect(rows[0]['Bug Responsibility']).toBe('');
    });

    it('should fallback bug responsibility from requirement when external bug row has unknown responsibility', () => {
      const requirements = [
        {
          requirementId: 'SR5305',
          baseKey: 'SR5305',
          title: 'Req 5305',
          subSystem: 'Auth',
          responsibility: 'IL',
          linkedTestCaseIds: [202],
        },
      ];
      const requirementIndex = new Map([
        [
          'SR5305',
          new Map([
            [
              202,
              {
                passed: 0,
                failed: 1,
                notRun: 0,
              },
            ],
          ]),
        ],
      ]);
      const observedTestCaseIdsByRequirement = new Map<string, Set<number>>([
        ['SR5305', new Set([202])],
      ]);
      const linkedRequirementsByTestCase = new Map([
        [
          202,
          {
            baseKeys: new Set(['SR5305']),
            fullCodes: new Set(['SR5305']),
            bugIds: new Set(),
          },
        ],
      ]);
      const externalBugsByTestCase = new Map([
        [
          202,
          [
            {
              id: 99001,
              title: 'External bug without SAPWBS',
              responsibility: 'Unknown',
              requirementBaseKey: 'SR5305',
            },
          ],
        ],
      ]);

      const rows = (resultDataProvider as any).buildMewpCoverageRows(
        requirements,
        requirementIndex,
        observedTestCaseIdsByRequirement,
        linkedRequirementsByTestCase,
        new Map(),
        externalBugsByTestCase
      );

      expect(rows).toHaveLength(1);
      expect(rows[0]['Bug ID']).toBe(99001);
      expect(rows[0]['Bug Responsibility']).toBe('Elisra');
    });

    it('should fallback bug responsibility from test case mapping when external bug row is unknown', () => {
      const requirements = [
        {
          requirementId: 'SR5310',
          baseKey: 'SR5310',
          title: 'Req 5310',
          subSystem: 'Power',
          responsibility: '',
          linkedTestCaseIds: [101],
        },
      ];

      const requirementIndex = new Map([
        [
          'SR5310',
          new Map([
            [
              101,
              {
                passed: 0,
                failed: 1,
                notRun: 0,
              },
            ],
          ]),
        ],
      ]);

      const observedTestCaseIdsByRequirement = new Map<string, Set<number>>([
        ['SR5310', new Set([101])],
      ]);

      const linkedRequirementsByTestCase = new Map([
        [
          101,
          {
            baseKeys: new Set(['SR5310']),
            fullCodes: new Set(['SR5310']),
            bugIds: new Set([10003]),
          },
        ],
      ]);

      const externalBugsByTestCase = new Map([
        [
          101,
          [
            {
              id: 10003,
              title: 'Bug 10003',
              responsibility: 'Unknown',
              requirementBaseKey: 'SR5310',
            },
          ],
        ],
      ]);

      const rows = (resultDataProvider as any).buildMewpCoverageRows(
        requirements,
        requirementIndex,
        observedTestCaseIdsByRequirement,
        linkedRequirementsByTestCase,
        new Map(),
        externalBugsByTestCase,
        new Map(),
        new Map([[101, 'IL']])
      );

      expect(rows).toHaveLength(1);
      expect(rows[0]['Bug Responsibility']).toBe('Elisra');
    });
  });

  describe('getMewpInternalValidationFlatResults', () => {
    it('should skip test cases with no in-scope expected requirements and no linked requirements', async () => {
      jest.spyOn(resultDataProvider as any, 'fetchTestPlanName').mockResolvedValueOnce('Plan A');
      jest.spyOn(resultDataProvider as any, 'fetchTestSuites').mockResolvedValueOnce([{ testSuiteId: 1 }]);
      jest.spyOn(resultDataProvider as any, 'fetchTestData').mockResolvedValueOnce([
        {
          testPointsItems: [
            { testCaseId: 101, testCaseName: 'TC 101' },
            { testCaseId: 102, testCaseName: 'TC 102' },
          ],
          testCasesItems: [
            {
              workItem: {
                id: 101,
                workItemFields: [{ key: 'Steps', value: '<steps></steps>' }],
              },
            },
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
          requirementId: 'SR0001',
          baseKey: 'SR0001',
          title: 'Req 1',
          responsibility: 'ESUK',
          linkedTestCaseIds: [101],
          areaPath: 'MEWP\\Customer Requirements\\Level 2',
        },
      ]);
      jest.spyOn(resultDataProvider as any, 'buildLinkedRequirementsByTestCase').mockResolvedValueOnce(
        new Map([[101, { baseKeys: new Set(['SR0001']), fullCodes: new Set(['SR0001']) }]])
      );
      jest
        .spyOn((resultDataProvider as any).testStepParserHelper, 'parseTestSteps')
        .mockResolvedValueOnce([
          {
            stepId: '1',
            stepPosition: '1',
            action: '',
            expected: 'SR0001',
            isSharedStepTitle: false,
          },
        ])
        .mockResolvedValueOnce([
          {
            stepId: '1',
            stepPosition: '1',
            action: '',
            expected: '',
            isSharedStepTitle: false,
          },
        ]);

      const result = await (resultDataProvider as any).getMewpInternalValidationFlatResults(
        '123',
        mockProjectName,
        [1]
      );

      expect(result.rows).toHaveLength(2);
      expect(result.rows).toEqual(
        expect.arrayContaining([
          expect.objectContaining({
            'Test Case ID': 101,
            'Validation Status': 'Pass',
            'Mentioned but Not Linked': '',
            'Linked but Not Mentioned': '',
          }),
          expect.objectContaining({
            'Test Case ID': 102,
            'Validation Status': 'Pass',
            'Mentioned but Not Linked': '',
            'Linked but Not Mentioned': '',
          }),
        ])
      );
    });

    it('should ignore non-L2 SR mentions and still report only real L2 discrepancies', async () => {
      jest.spyOn(resultDataProvider as any, 'fetchTestPlanName').mockResolvedValueOnce('Plan A');
      jest.spyOn(resultDataProvider as any, 'fetchTestSuites').mockResolvedValueOnce([{ testSuiteId: 1 }]);
      jest.spyOn(resultDataProvider as any, 'fetchTestData').mockResolvedValueOnce([
        {
          testPointsItems: [{ testCaseId: 101, testCaseName: 'TC 101' }],
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
          workItemId: 5001,
          requirementId: 'SR0001',
          baseKey: 'SR0001',
          title: 'Req 1',
          responsibility: 'ESUK',
          linkedTestCaseIds: [101],
          areaPath: 'MEWP\\Customer Requirements\\Level 2',
        },
      ]);
      jest.spyOn(resultDataProvider as any, 'buildLinkedRequirementsByTestCase').mockResolvedValueOnce(
        new Map()
      );
      jest.spyOn((resultDataProvider as any).testStepParserHelper, 'parseTestSteps').mockResolvedValueOnce([
        {
          stepId: '1',
          stepPosition: '1',
          action: '',
          expected: 'SR0001; SR9999',
          isSharedStepTitle: false,
        },
      ]);

      const result = await (resultDataProvider as any).getMewpInternalValidationFlatResults(
        '123',
        mockProjectName,
        [1]
      );

      expect(result.rows).toHaveLength(1);
      expect(result.rows[0]).toEqual(
        expect.objectContaining({
          'Test Case ID': 101,
          'Mentioned but Not Linked': 'Step 1: SR0001',
          'Linked but Not Mentioned': '',
          'Validation Status': 'Fail',
        })
      );
    });

    it('should emit Direction A rows only for specifically mentioned child requirements', async () => {
      jest.spyOn(resultDataProvider as any, 'fetchTestPlanName').mockResolvedValueOnce('Plan A');
      jest.spyOn(resultDataProvider as any, 'fetchTestSuites').mockResolvedValueOnce([{ testSuiteId: 1 }]);
      jest.spyOn(resultDataProvider as any, 'fetchTestData').mockResolvedValueOnce([
        {
          testPointsItems: [{ testCaseId: 101, testCaseName: 'TC 101' }],
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
          workItemId: 5001,
          requirementId: 'SR0001',
          baseKey: 'SR0001',
          title: 'Req parent',
          responsibility: 'ESUK',
          linkedTestCaseIds: [101],
          areaPath: 'MEWP\\Customer Requirements\\Level 2',
        },
        {
          workItemId: 5002,
          requirementId: 'SR0001-1',
          baseKey: 'SR0001',
          title: 'Req child',
          responsibility: 'ESUK',
          linkedTestCaseIds: [],
          areaPath: 'MEWP\\Customer Requirements\\Level 2\\MOP',
        },
      ]);
      jest.spyOn(resultDataProvider as any, 'buildLinkedRequirementsByTestCase').mockResolvedValueOnce(
        new Map([[101, { baseKeys: new Set(['SR0001']), fullCodes: new Set(['SR0001']) }]])
      );
      jest.spyOn((resultDataProvider as any).testStepParserHelper, 'parseTestSteps').mockResolvedValueOnce([
        {
          stepId: '2',
          stepPosition: '2',
          action: '',
          expected: 'SR0001-1',
          isSharedStepTitle: false,
        },
      ]);

      const result = await (resultDataProvider as any).getMewpInternalValidationFlatResults(
        '123',
        mockProjectName,
        [1]
      );

      expect(result.rows).toEqual(
        expect.arrayContaining([
          expect.objectContaining({
            'Test Case ID': 101,
            'Mentioned but Not Linked': expect.stringContaining('Step 2: SR0001-1'),
            'Validation Status': 'Fail',
          }),
        ])
      );
    });

    it('should keep explicit child IDs when multiple specifically mentioned children are missing', async () => {
      jest.spyOn(resultDataProvider as any, 'fetchTestPlanName').mockResolvedValueOnce('Plan A');
      jest.spyOn(resultDataProvider as any, 'fetchTestSuites').mockResolvedValueOnce([{ testSuiteId: 1 }]);
      jest.spyOn(resultDataProvider as any, 'fetchTestData').mockResolvedValueOnce([
        {
          testPointsItems: [{ testCaseId: 111, testCaseName: 'TC 111' }],
          testCasesItems: [
            {
              workItem: {
                id: 111,
                workItemFields: [{ key: 'Steps', value: '<steps></steps>' }],
              },
            },
          ],
        },
      ]);
      jest.spyOn(resultDataProvider as any, 'fetchMewpL2Requirements').mockResolvedValueOnce([
        {
          workItemId: 5101,
          requirementId: 'SR0054-1',
          baseKey: 'SR0054',
          title: 'Req child 1',
          responsibility: 'ESUK',
          linkedTestCaseIds: [],
          areaPath: 'MEWP\\Customer Requirements\\Level 2',
        },
        {
          workItemId: 5102,
          requirementId: 'SR0054-2',
          baseKey: 'SR0054',
          title: 'Req child 2',
          responsibility: 'ESUK',
          linkedTestCaseIds: [],
          areaPath: 'MEWP\\Customer Requirements\\Level 2',
        },
      ]);
      jest
        .spyOn(resultDataProvider as any, 'buildLinkedRequirementsByTestCase')
        .mockResolvedValueOnce(new Map([[111, { baseKeys: new Set<string>(), fullCodes: new Set<string>() }]]));
      jest.spyOn((resultDataProvider as any).testStepParserHelper, 'parseTestSteps').mockResolvedValueOnce([
        {
          stepId: '4',
          stepPosition: '4',
          action: '',
          expected: 'SR0054-1; SR0054-2',
          isSharedStepTitle: false,
        },
      ]);

      const result = await (resultDataProvider as any).getMewpInternalValidationFlatResults(
        '123',
        mockProjectName,
        [1]
      );

      expect(result.rows).toHaveLength(1);
      expect(result.rows[0]).toEqual(
        expect.objectContaining({
          'Test Case ID': 111,
          'Mentioned but Not Linked': 'Step 4: SR0054-1; SR0054-2',
          'Linked but Not Mentioned': '',
          'Validation Status': 'Fail',
        })
      );
    });

    it('should pass when a base SR mention is fully covered by its only linked child', async () => {
      jest.spyOn(resultDataProvider as any, 'fetchTestPlanName').mockResolvedValueOnce('Plan A');
      jest.spyOn(resultDataProvider as any, 'fetchTestSuites').mockResolvedValueOnce([{ testSuiteId: 1 }]);
      jest.spyOn(resultDataProvider as any, 'fetchTestData').mockResolvedValueOnce([
        {
          testPointsItems: [{ testCaseId: 402, testCaseName: 'TC 402 - Single child covered' }],
          testCasesItems: [
            {
              workItem: {
                id: 402,
                workItemFields: [{ key: 'Steps', value: '<steps id=\"mock-steps-tc-402\"></steps>' }],
              },
            },
          ],
        },
      ]);
      jest.spyOn(resultDataProvider as any, 'fetchMewpL2Requirements').mockResolvedValueOnce([
        {
          workItemId: 9101,
          requirementId: 'SR0054',
          baseKey: 'SR0054',
          title: 'SR0054 parent',
          responsibility: 'ESUK',
          linkedTestCaseIds: [],
          areaPath: 'MEWP\\Customer Requirements\\Level 2',
        },
        {
          workItemId: 9102,
          requirementId: 'SR0054-1',
          baseKey: 'SR0054',
          title: 'SR0054 child 1',
          responsibility: 'ESUK',
          linkedTestCaseIds: [],
          areaPath: 'MEWP\\Customer Requirements\\Level 2',
        },
      ]);
      jest.spyOn(resultDataProvider as any, 'buildLinkedRequirementsByTestCase').mockResolvedValueOnce(
        new Map([
          [
            402,
            {
              baseKeys: new Set(['SR0054']),
              fullCodes: new Set(['SR0054-1']),
            },
          ],
        ])
      );
      jest.spyOn((resultDataProvider as any).testStepParserHelper, 'parseTestSteps').mockResolvedValueOnce([
        {
          stepId: '3',
          stepPosition: '3',
          action: 'Validate family root',
          expected: 'SR0054',
          isSharedStepTitle: false,
        },
      ]);

      const result = await (resultDataProvider as any).getMewpInternalValidationFlatResults(
        '123',
        mockProjectName,
        [1]
      );

      expect(result.rows).toHaveLength(1);
      expect(result.rows[0]).toEqual(
        expect.objectContaining({
          'Test Case ID': 402,
          'Mentioned but Not Linked': '',
          'Linked but Not Mentioned': '',
          'Validation Status': 'Pass',
        })
      );
    });

    it('should support cross-test-case family coverage when siblings are linked on different test cases', async () => {
      jest.spyOn(resultDataProvider as any, 'fetchTestPlanName').mockResolvedValueOnce('Plan A');
      jest.spyOn(resultDataProvider as any, 'fetchTestSuites').mockResolvedValueOnce([{ testSuiteId: 1 }]);
      jest.spyOn(resultDataProvider as any, 'fetchTestData').mockResolvedValueOnce([
        {
          testPointsItems: [
            { testCaseId: 501, testCaseName: 'TC 501 - sibling 1' },
            { testCaseId: 502, testCaseName: 'TC 502 - sibling 2' },
          ],
          testCasesItems: [
            {
              workItem: {
                id: 501,
                workItemFields: [{ key: 'Steps', value: '<steps id=\"mock-steps-tc-501\"></steps>' }],
              },
            },
            {
              workItem: {
                id: 502,
                workItemFields: [{ key: 'Steps', value: '<steps id=\"mock-steps-tc-502\"></steps>' }],
              },
            },
          ],
        },
      ]);
      jest.spyOn(resultDataProvider as any, 'fetchMewpL2Requirements').mockResolvedValueOnce([
        {
          workItemId: 9301,
          requirementId: 'SR0054-1',
          baseKey: 'SR0054',
          title: 'SR0054 child 1',
          responsibility: 'ESUK',
          linkedTestCaseIds: [],
          areaPath: 'MEWP\\Customer Requirements\\Level 2',
        },
        {
          workItemId: 9302,
          requirementId: 'SR0054-2',
          baseKey: 'SR0054',
          title: 'SR0054 child 2',
          responsibility: 'ESUK',
          linkedTestCaseIds: [],
          areaPath: 'MEWP\\Customer Requirements\\Level 2',
        },
      ]);
      jest.spyOn(resultDataProvider as any, 'buildLinkedRequirementsByTestCase').mockResolvedValueOnce(
        new Map([
          [
            501,
            {
              baseKeys: new Set(['SR0054']),
              fullCodes: new Set(['SR0054-1']),
            },
          ],
          [
            502,
            {
              baseKeys: new Set(['SR0054']),
              fullCodes: new Set(['SR0054-2']),
            },
          ],
        ])
      );
      jest
        .spyOn((resultDataProvider as any).testStepParserHelper, 'parseTestSteps')
        .mockResolvedValueOnce([
          {
            stepId: '1',
            stepPosition: '1',
            action: 'Parent mention on first test case',
            expected: 'SR0054',
            isSharedStepTitle: false,
          },
        ])
        .mockResolvedValueOnce([
          {
            stepId: '1',
            stepPosition: '1',
            action: 'Parent mention on second test case',
            expected: 'SR0054',
            isSharedStepTitle: false,
          },
        ]);

      const result = await (resultDataProvider as any).getMewpInternalValidationFlatResults(
        '123',
        mockProjectName,
        [1]
      );

      const byTestCase = new Map<number, any>(
        result.rows.map((row: any) => [Number(row['Test Case ID']), row])
      );
      expect(byTestCase.get(501)).toEqual(
        expect.objectContaining({
          'Mentioned but Not Linked': '',
          'Linked but Not Mentioned': '',
          'Validation Status': 'Pass',
        })
      );
      expect(byTestCase.get(502)).toEqual(
        expect.objectContaining({
          'Mentioned but Not Linked': '',
          'Linked but Not Mentioned': '',
          'Validation Status': 'Pass',
        })
      );
    });

    it('should group linked-but-not-mentioned requirements by SR family', async () => {
      jest.spyOn(resultDataProvider as any, 'fetchTestPlanName').mockResolvedValueOnce('Plan A');
      jest.spyOn(resultDataProvider as any, 'fetchTestSuites').mockResolvedValueOnce([{ testSuiteId: 1 }]);
      jest.spyOn(resultDataProvider as any, 'fetchTestData').mockResolvedValueOnce([
        {
          testPointsItems: [{ testCaseId: 403, testCaseName: 'TC 403 - Linked only' }],
          testCasesItems: [
            {
              workItem: {
                id: 403,
                workItemFields: [{ key: 'Steps', value: '<steps id=\"mock-steps-tc-403\"></steps>' }],
              },
            },
          ],
        },
      ]);
      jest.spyOn(resultDataProvider as any, 'fetchMewpL2Requirements').mockResolvedValueOnce([
        {
          workItemId: 9201,
          requirementId: 'SR0054-1',
          baseKey: 'SR0054',
          title: 'SR0054 child 1',
          responsibility: 'ESUK',
          linkedTestCaseIds: [],
          areaPath: 'MEWP\\Customer Requirements\\Level 2',
        },
        {
          workItemId: 9202,
          requirementId: 'SR0054-2',
          baseKey: 'SR0054',
          title: 'SR0054 child 2',
          responsibility: 'ESUK',
          linkedTestCaseIds: [],
          areaPath: 'MEWP\\Customer Requirements\\Level 2',
        },
        {
          workItemId: 9203,
          requirementId: 'SR0100-1',
          baseKey: 'SR0100',
          title: 'SR0100 child 1',
          responsibility: 'ESUK',
          linkedTestCaseIds: [],
          areaPath: 'MEWP\\Customer Requirements\\Level 2',
        },
      ]);
      jest.spyOn(resultDataProvider as any, 'buildLinkedRequirementsByTestCase').mockResolvedValueOnce(
        new Map([
          [
            403,
            {
              baseKeys: new Set(['SR0054', 'SR0100']),
              fullCodes: new Set(['SR0054-1', 'SR0054-2', 'SR0100-1']),
            },
          ],
        ])
      );
      jest.spyOn((resultDataProvider as any).testStepParserHelper, 'parseTestSteps').mockResolvedValueOnce([
        {
          stepId: '1',
          stepPosition: '1',
          action: 'Action only',
          expected: '',
          isSharedStepTitle: false,
        },
      ]);

      const result = await (resultDataProvider as any).getMewpInternalValidationFlatResults(
        '123',
        mockProjectName,
        [1]
      );

      expect(result.rows).toHaveLength(1);
      expect(String(result.rows[0]['Mentioned but Not Linked'] || '')).toBe('');
      expect(String(result.rows[0]['Linked but Not Mentioned'] || '')).toContain('SR0054');
      expect(String(result.rows[0]['Linked but Not Mentioned'] || '')).not.toContain('SR0054-1');
      expect(String(result.rows[0]['Linked but Not Mentioned'] || '')).not.toContain('SR0054-2');
      expect(String(result.rows[0]['Linked but Not Mentioned'] || '')).toContain('SR0100-1');
    });

    it('should report only base SR when an entire mentioned family is uncovered', async () => {
      jest.spyOn(resultDataProvider as any, 'fetchTestPlanName').mockResolvedValueOnce('Plan A');
      jest.spyOn(resultDataProvider as any, 'fetchTestSuites').mockResolvedValueOnce([{ testSuiteId: 1 }]);
      jest.spyOn(resultDataProvider as any, 'fetchTestData').mockResolvedValueOnce([
        {
          testPointsItems: [{ testCaseId: 401, testCaseName: 'TC 401 - Family uncovered' }],
          testCasesItems: [
            {
              workItem: {
                id: 401,
                workItemFields: [{ key: 'Steps', value: '<steps id=\"mock-steps-tc-401\"></steps>' }],
              },
            },
          ],
        },
      ]);
      jest.spyOn(resultDataProvider as any, 'fetchMewpL2Requirements').mockResolvedValueOnce([
        {
          workItemId: 9001,
          requirementId: 'SR0054-1',
          baseKey: 'SR0054',
          title: 'SR0054 child 1',
          responsibility: 'ESUK',
          linkedTestCaseIds: [],
          areaPath: 'MEWP\\Customer Requirements\\Level 2',
        },
        {
          workItemId: 9002,
          requirementId: 'SR0054-2',
          baseKey: 'SR0054',
          title: 'SR0054 child 2',
          responsibility: 'ESUK',
          linkedTestCaseIds: [],
          areaPath: 'MEWP\\Customer Requirements\\Level 2',
        },
      ]);
      jest.spyOn(resultDataProvider as any, 'buildLinkedRequirementsByTestCase').mockResolvedValueOnce(
        new Map([[401, { baseKeys: new Set<string>(), fullCodes: new Set<string>() }]])
      );
      jest.spyOn((resultDataProvider as any).testStepParserHelper, 'parseTestSteps').mockResolvedValueOnce([
        {
          stepId: '3',
          stepPosition: '3',
          action: 'Validate family root',
          expected: 'SR0054',
          isSharedStepTitle: false,
        },
      ]);

      const result = await (resultDataProvider as any).getMewpInternalValidationFlatResults(
        '123',
        mockProjectName,
        [1]
      );

      expect(result.rows).toHaveLength(1);
      expect(result.rows[0]).toEqual(
        expect.objectContaining({
          'Test Case ID': 401,
          'Mentioned but Not Linked': 'Step 3: SR0054',
          'Linked but Not Mentioned': '',
          'Validation Status': 'Fail',
        })
      );
      expect(String(result.rows[0]['Mentioned but Not Linked'] || '')).not.toContain('SR0054-1');
      expect(String(result.rows[0]['Mentioned but Not Linked'] || '')).not.toContain('SR0054-2');
    });

    it('should not duplicate Direction A discrepancy when same requirement is repeated in multiple steps', async () => {
      jest.spyOn(resultDataProvider as any, 'fetchTestPlanName').mockResolvedValueOnce('Plan A');
      jest.spyOn(resultDataProvider as any, 'fetchTestSuites').mockResolvedValueOnce([{ testSuiteId: 1 }]);
      jest.spyOn(resultDataProvider as any, 'fetchTestData').mockResolvedValueOnce([
        {
          testPointsItems: [{ testCaseId: 101, testCaseName: 'TC 101' }],
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
          workItemId: 5001,
          requirementId: 'SR0001',
          baseKey: 'SR0001',
          title: 'Req 1',
          responsibility: 'ESUK',
          linkedTestCaseIds: [],
          areaPath: 'MEWP\\Customer Requirements\\Level 2',
        },
      ]);
      jest
        .spyOn(resultDataProvider as any, 'buildLinkedRequirementsByTestCase')
        .mockResolvedValueOnce(new Map());
      jest.spyOn((resultDataProvider as any).testStepParserHelper, 'parseTestSteps').mockResolvedValueOnce([
        {
          stepId: '1',
          stepPosition: '1',
          action: '',
          expected: 'SR0001',
          isSharedStepTitle: false,
        },
        {
          stepId: '2',
          stepPosition: '2',
          action: '',
          expected: 'SR0001',
          isSharedStepTitle: false,
        },
      ]);

      const result = await (resultDataProvider as any).getMewpInternalValidationFlatResults(
        '123',
        mockProjectName,
        [1]
      );

      expect(result.rows).toHaveLength(1);
      expect(result.rows[0]).toEqual(
        expect.objectContaining({
          'Test Case ID': 101,
          'Mentioned but Not Linked': 'Step 1: SR0001',
          'Linked but Not Mentioned': '',
          'Validation Status': 'Fail',
        })
      );
    });

    it('should not flag linked child as Direction B when parent family is mentioned in Expected Result', async () => {
      jest.spyOn(resultDataProvider as any, 'fetchTestPlanName').mockResolvedValueOnce('Plan A');
      jest.spyOn(resultDataProvider as any, 'fetchTestSuites').mockResolvedValueOnce([{ testSuiteId: 1 }]);
      jest.spyOn(resultDataProvider as any, 'fetchTestData').mockResolvedValueOnce([
        {
          testPointsItems: [{ testCaseId: 301, testCaseName: 'TC 301' }],
          testCasesItems: [
            {
              workItem: {
                id: 301,
                workItemFields: [{ key: 'Steps', value: '<steps></steps>' }],
              },
            },
          ],
        },
      ]);
      jest.spyOn(resultDataProvider as any, 'fetchMewpL2Requirements').mockResolvedValueOnce([
        {
          workItemId: 7001,
          requirementId: 'SR0054',
          baseKey: 'SR0054',
          title: 'Parent 0054',
          responsibility: 'ESUK',
          linkedTestCaseIds: [],
          areaPath: 'MEWP\\Customer Requirements\\Level 2',
        },
      ]);
      jest.spyOn(resultDataProvider as any, 'buildLinkedRequirementsByTestCase').mockResolvedValueOnce(
        new Map([
          [
            301,
            {
              baseKeys: new Set(['SR0054']),
              fullCodes: new Set(['SR0054-1']),
            },
          ],
        ])
      );
      jest.spyOn((resultDataProvider as any).testStepParserHelper, 'parseTestSteps').mockResolvedValueOnce([
        {
          stepId: '1',
          stepPosition: '1',
          action: '',
          expected: 'SR0054',
          isSharedStepTitle: false,
        },
      ]);

      const result = await (resultDataProvider as any).getMewpInternalValidationFlatResults(
        '123',
        mockProjectName,
        [1]
      );

      expect(result.rows).toHaveLength(1);
      expect(result.rows[0]).toEqual(
        expect.objectContaining({
          'Test Case ID': 301,
          'Linked but Not Mentioned': '',
        })
      );
    });

    it('should produce one detailed row per test case with correct bidirectional discrepancies', async () => {
      const mockDetailedStepsByTestCase = new Map<number, any[]>([
        [
          201,
          [
            {
              stepId: '1',
              stepPosition: '1',
              action: 'Validate parent SR0511 and SR0095 siblings',
              expected: '<b>sr0511</b>; SR0095-2,3; VVRM-05',
              isSharedStepTitle: false,
            },
            {
              stepId: '2',
              stepPosition: '2',
              action: 'Noisy requirement-like token should be ignored',
              expected: 'SR0511-V3.24',
              isSharedStepTitle: false,
            },
            {
              stepId: '3',
              stepPosition: '3',
              action: 'Regression note',
              expected: 'Verification note only',
              isSharedStepTitle: false,
            },
          ],
        ],
        [
          202,
          [
            {
              stepId: '1',
              stepPosition: '1',
              action: 'Linked SR0200 exists but is not cleanly mentioned in expected result',
              expected: 'VVRM-22; SR0200 V3.1',
              isSharedStepTitle: false,
            },
            {
              stepId: '2',
              stepPosition: '2',
              action: 'Execution notes',
              expected: 'Notes without SR requirement token',
              isSharedStepTitle: false,
            },
          ],
        ],
        [
          203,
          [
            {
              stepId: '1',
              stepPosition: '1',
              action: 'Primary requirement validation for SR0100-1',
              expected: '<i>SR0100-1</i>',
              isSharedStepTitle: false,
            },
            {
              stepId: '3',
              stepPosition: '3',
              action: 'Repeated mention should not create duplicate mismatch',
              expected: 'SR0100-1; SR0100-1',
              isSharedStepTitle: false,
            },
          ],
        ],
      ]);
      const mockLinkedRequirementsByTestCase = new Map<number, any>([
        [
          201,
          {
            baseKeys: new Set(['SR0511', 'SR0095', 'SR8888']),
            fullCodes: new Set(['SR0511', 'SR0095-2', 'SR8888']),
          },
        ],
        [
          202,
          {
            baseKeys: new Set(['SR0200']),
            fullCodes: new Set(['SR0200']),
          },
        ],
        [
          203,
          {
            baseKeys: new Set(['SR0100']),
            fullCodes: new Set(['SR0100-1']),
          },
        ],
      ]);

      jest.spyOn(resultDataProvider as any, 'fetchTestPlanName').mockResolvedValueOnce('Plan A');
      jest.spyOn(resultDataProvider as any, 'fetchTestSuites').mockResolvedValueOnce([{ testSuiteId: 1 }]);
      jest.spyOn(resultDataProvider as any, 'fetchTestData').mockResolvedValueOnce([
        {
          testPointsItems: [
            { testCaseId: 201, testCaseName: 'TC 201 - Mixed discrepancies' },
            { testCaseId: 202, testCaseName: 'TC 202 - Link only' },
            { testCaseId: 203, testCaseName: 'TC 203 - Fully valid' },
          ],
          testCasesItems: [
            {
              workItem: {
                id: 201,
                workItemFields: [{ key: 'Steps', value: '<steps id="mock-steps-tc-201"></steps>' }],
              },
            },
            {
              workItem: {
                id: 202,
                workItemFields: [{ key: 'Steps', value: '<steps id="mock-steps-tc-202"></steps>' }],
              },
            },
            {
              workItem: {
                id: 203,
                workItemFields: [{ key: 'Steps', value: '<steps id="mock-steps-tc-203"></steps>' }],
              },
            },
          ],
        },
      ]);
      jest.spyOn(resultDataProvider as any, 'fetchMewpL2Requirements').mockResolvedValueOnce([
        {
          workItemId: 6001,
          requirementId: 'SR0511',
          baseKey: 'SR0511',
          title: 'Parent 0511',
          responsibility: 'ESUK',
          linkedTestCaseIds: [201],
          areaPath: 'MEWP\\Customer Requirements\\Level 2',
        },
        {
          workItemId: 6002,
          requirementId: 'SR0511-1',
          baseKey: 'SR0511',
          title: 'Child 0511-1',
          responsibility: 'ESUK',
          linkedTestCaseIds: [],
          areaPath: 'MEWP\\Customer Requirements\\Level 2\\MOP',
        },
        {
          workItemId: 6003,
          requirementId: 'SR0511-2',
          baseKey: 'SR0511',
          title: 'Child 0511-2',
          responsibility: 'ESUK',
          linkedTestCaseIds: [],
          areaPath: 'MEWP\\Customer Requirements\\Level 2\\MOP',
        },
        {
          workItemId: 6004,
          requirementId: 'SR0095-2',
          baseKey: 'SR0095',
          title: 'SR0095 child 2',
          responsibility: 'ESUK',
          linkedTestCaseIds: [],
          areaPath: 'MEWP\\Customer Requirements\\Level 2\\MOP',
        },
        {
          workItemId: 6005,
          requirementId: 'SR0095-3',
          baseKey: 'SR0095',
          title: 'SR0095 child 3',
          responsibility: 'ESUK',
          linkedTestCaseIds: [],
          areaPath: 'MEWP\\Customer Requirements\\Level 2\\MOP',
        },
        {
          workItemId: 6006,
          requirementId: 'SR0200',
          baseKey: 'SR0200',
          title: 'SR0200 standalone',
          responsibility: 'ESUK',
          linkedTestCaseIds: [202],
          areaPath: 'MEWP\\Customer Requirements\\Level 2',
        },
        {
          workItemId: 6007,
          requirementId: 'SR0100-1',
          baseKey: 'SR0100',
          title: 'SR0100 child 1',
          responsibility: 'ESUK',
          linkedTestCaseIds: [203],
          areaPath: 'MEWP\\Customer Requirements\\Level 2\\MOP',
        },
      ]);
      jest
        .spyOn(resultDataProvider as any, 'buildLinkedRequirementsByTestCase')
        .mockResolvedValueOnce(mockLinkedRequirementsByTestCase);
      jest
        .spyOn((resultDataProvider as any).testStepParserHelper, 'parseTestSteps')
        .mockImplementation(async (...args: unknown[]) => {
          const stepsXml = String(args?.[0] || '');
          const testCaseMatch = /mock-steps-tc-(\d+)/i.exec(stepsXml);
          const testCaseId = Number(testCaseMatch?.[1] || 0);
          return mockDetailedStepsByTestCase.get(testCaseId) || [];
        });

      const result = await (resultDataProvider as any).getMewpInternalValidationFlatResults(
        '123',
        mockProjectName,
        [1]
      );

      expect(result.rows).toHaveLength(3);
      const byTestCase = new Map<number, any>(result.rows.map((row: any) => [row['Test Case ID'], row]));
      expect(new Set(result.rows.map((row: any) => row['Test Case ID']))).toEqual(new Set([201, 202, 203]));

      expect(byTestCase.get(201)).toEqual(
        expect.objectContaining({
          'Test Case Title': 'TC 201 - Mixed discrepancies',
          'Mentioned but Not Linked': 'Step 1: SR0095-3; SR0511',
          'Linked but Not Mentioned': 'SR8888',
          'Validation Status': 'Fail',
        })
      );
      expect(String(byTestCase.get(201)['Mentioned but Not Linked'] || '')).not.toContain('VVRM');

      expect(byTestCase.get(202)).toEqual(
        expect.objectContaining({
          'Test Case Title': 'TC 202 - Link only',
          'Mentioned but Not Linked': '',
          'Linked but Not Mentioned': 'SR0200',
          'Validation Status': 'Fail',
        })
      );

      expect(byTestCase.get(203)).toEqual(
        expect.objectContaining({
          'Test Case Title': 'TC 203 - Fully valid',
          'Mentioned but Not Linked': '',
          'Linked but Not Mentioned': '',
          'Validation Status': 'Pass',
        })
      );

      expect((resultDataProvider as any).testStepParserHelper.parseTestSteps).toHaveBeenCalledTimes(3);
      const parseCalls = ((resultDataProvider as any).testStepParserHelper.parseTestSteps as jest.Mock).mock.calls;
      const parsedCaseIds = parseCalls
        .map(([xml]: [string]) => {
          const match = /mock-steps-tc-(\d+)/i.exec(String(xml || ''));
          return Number(match?.[1] || 0);
        })
        .filter((id: number) => Number.isFinite(id) && id > 0);
      expect(new Set(parsedCaseIds)).toEqual(new Set([201, 202, 203]));
    });

    it('should parse TC-0042 mixed expected text and keep only valid SR requirement codes', async () => {
      jest.spyOn(resultDataProvider as any, 'fetchTestPlanName').mockResolvedValueOnce('Plan A');
      jest.spyOn(resultDataProvider as any, 'fetchTestSuites').mockResolvedValueOnce([{ testSuiteId: 1 }]);
      jest.spyOn(resultDataProvider as any, 'fetchTestData').mockResolvedValueOnce([
        {
          testPointsItems: [{ testCaseId: 42, testCaseName: 'TC-0042' }],
          testCasesItems: [
            {
              workItem: {
                id: 42,
                workItemFields: [{ key: 'Steps', value: '<steps id="mock-steps-tc-0042"></steps>' }],
              },
            },
          ],
        },
      ]);
      jest.spyOn(resultDataProvider as any, 'fetchMewpL2Requirements').mockResolvedValueOnce([
        {
          workItemId: 7001,
          requirementId: 'SR0036',
          baseKey: 'SR0036',
          title: 'SR0036',
          responsibility: 'ESUK',
          linkedTestCaseIds: [],
          areaPath: 'MEWP\\Customer Requirements\\Level 2',
        },
        {
          workItemId: 7002,
          requirementId: 'SR0215',
          baseKey: 'SR0215',
          title: 'SR0215',
          responsibility: 'ESUK',
          linkedTestCaseIds: [],
          areaPath: 'MEWP\\Customer Requirements\\Level 2',
        },
        {
          workItemId: 7003,
          requirementId: 'SR0539',
          baseKey: 'SR0539',
          title: 'SR0539',
          responsibility: 'ESUK',
          linkedTestCaseIds: [],
          areaPath: 'MEWP\\Customer Requirements\\Level 2',
        },
        {
          workItemId: 7004,
          requirementId: 'SR0348',
          baseKey: 'SR0348',
          title: 'SR0348',
          responsibility: 'ESUK',
          linkedTestCaseIds: [],
          areaPath: 'MEWP\\Customer Requirements\\Level 2',
        },
        {
          workItemId: 7005,
          requirementId: 'SR0027',
          baseKey: 'SR0027',
          title: 'SR0027',
          responsibility: 'ESUK',
          linkedTestCaseIds: [],
          areaPath: 'MEWP\\Customer Requirements\\Level 2',
        },
        {
          workItemId: 7006,
          requirementId: 'SR0041',
          baseKey: 'SR0041',
          title: 'SR0041',
          responsibility: 'ESUK',
          linkedTestCaseIds: [],
          areaPath: 'MEWP\\Customer Requirements\\Level 2',
        },
        {
          workItemId: 7007,
          requirementId: 'SR0550',
          baseKey: 'SR0550',
          title: 'SR0550',
          responsibility: 'ESUK',
          linkedTestCaseIds: [],
          areaPath: 'MEWP\\Customer Requirements\\Level 2',
        },
        {
          workItemId: 7008,
          requirementId: 'SR0550-2',
          baseKey: 'SR0550',
          title: 'SR0550-2',
          responsibility: 'ESUK',
          linkedTestCaseIds: [],
          areaPath: 'MEWP\\Customer Requirements\\Level 2\\MOP',
        },
        {
          workItemId: 7009,
          requirementId: 'SR0817',
          baseKey: 'SR0817',
          title: 'SR0817',
          responsibility: 'ESUK',
          linkedTestCaseIds: [],
          areaPath: 'MEWP\\Customer Requirements\\Level 2',
        },
        {
          workItemId: 7010,
          requirementId: 'SR0818',
          baseKey: 'SR0818',
          title: 'SR0818',
          responsibility: 'ESUK',
          linkedTestCaseIds: [],
          areaPath: 'MEWP\\Customer Requirements\\Level 2',
        },
        {
          workItemId: 7011,
          requirementId: 'SR0859',
          baseKey: 'SR0859',
          title: 'SR0859',
          responsibility: 'ESUK',
          linkedTestCaseIds: [],
          areaPath: 'MEWP\\Customer Requirements\\Level 2',
        },
      ]);
      jest
        .spyOn(resultDataProvider as any, 'buildLinkedRequirementsByTestCase')
        .mockResolvedValueOnce(
          new Map([
            [
              42,
              {
                baseKeys: new Set([
                  'SR0036',
                  'SR0215',
                  'SR0539',
                  'SR0348',
                  'SR0041',
                  'SR0550',
                  'SR0817',
                  'SR0818',
                  'SR0859',
                ]),
                fullCodes: new Set([
                  'SR0036',
                  'SR0215',
                  'SR0539',
                  'SR0348',
                  'SR0041',
                  'SR0550',
                  'SR0550-2',
                  'SR0817',
                  'SR0818',
                  'SR0859',
                ]),
              },
            ],
          ])
        );
      jest.spyOn((resultDataProvider as any).testStepParserHelper, 'parseTestSteps').mockResolvedValueOnce([
        {
          stepId: '1',
          stepPosition: '1',
          action:
            'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.',
          expected:
            'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. (SR0036; SR0215; VVRM-1; SR0817-V3.2; SR0818-V3.3; SR0859-V3.4)',
          isSharedStepTitle: false,
        },
        {
          stepId: '2',
          stepPosition: '2',
          action:
            'Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.',
          expected:
            'Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. (SR0539; SR0348; VVRM-1)',
          isSharedStepTitle: false,
        },
        {
          stepId: '3',
          stepPosition: '3',
          action:
            'Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur.',
          expected:
            'Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. (SR0027, SR0036; SR0041; VVRM-2)',
          isSharedStepTitle: false,
        },
        {
          stepId: '4',
          stepPosition: '4',
          action:
            'Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.',
          expected:
            'Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum. (SR0550-2)',
          isSharedStepTitle: false,
        },
        {
          stepId: '5',
          stepPosition: '5',
          action: 'Nemo enim ipsam voluptatem quia voluptas sit aspernatur aut odit aut fugit.',
          expected: 'Nemo enim ipsam voluptatem quia voluptas sit aspernatur aut odit aut fugit. (VVRM-1)',
          isSharedStepTitle: false,
        },
        {
          stepId: '6',
          stepPosition: '6',
          action:
            'Neque porro quisquam est qui dolorem ipsum quia dolor sit amet consectetur adipisci velit.',
          expected:
            'Neque porro quisquam est qui dolorem ipsum quia dolor sit amet consectetur adipisci velit. (SR0041; SR0550; VVRM-3)',
          isSharedStepTitle: false,
        },
      ]);

      const result = await (resultDataProvider as any).getMewpInternalValidationFlatResults(
        '123',
        mockProjectName,
        [1]
      );

      expect(result.rows).toHaveLength(1);
      expect(result.rows[0]).toEqual(
        expect.objectContaining({
          'Test Case ID': 42,
          'Test Case Title': 'TC-0042',
          'Mentioned but Not Linked': 'Step 3: SR0027',
          'Linked but Not Mentioned': 'SR0817; SR0818; SR0859',
          'Validation Status': 'Fail',
        })
      );
      expect(String(result.rows[0]['Mentioned but Not Linked'] || '')).not.toContain('VVRM');
      expect(String(result.rows[0]['Linked but Not Mentioned'] || '')).not.toContain('VVRM');
    });

    it('should fallback to work-item fields for steps XML when suite payload has no workItemFields', async () => {
      jest.spyOn(resultDataProvider as any, 'fetchTestPlanName').mockResolvedValueOnce('Plan A');
      jest.spyOn(resultDataProvider as any, 'fetchMewpScopedTestData').mockResolvedValueOnce([
        {
          testPointsItems: [{ testCaseId: 501, testCaseName: 'TC 501' }],
          testCasesItems: [
            {
              workItem: {
                id: 501,
                workItemFields: [],
              },
            },
          ],
        },
      ]);
      jest.spyOn(resultDataProvider as any, 'fetchMewpL2Requirements').mockResolvedValueOnce([
        {
          workItemId: 9001,
          requirementId: 'SR0501',
          baseKey: 'SR0501',
          title: 'Req 501',
          responsibility: 'ESUK',
          linkedTestCaseIds: [501],
          areaPath: 'MEWP\\Customer Requirements\\Level 2',
        },
      ]);
      jest.spyOn(resultDataProvider as any, 'buildLinkedRequirementsByTestCase').mockResolvedValueOnce(
        new Map([
          [
            501,
            {
              baseKeys: new Set(['SR0501']),
              fullCodes: new Set(['SR0501']),
            },
          ],
        ])
      );
      jest.spyOn(resultDataProvider as any, 'fetchWorkItemsByIds').mockResolvedValueOnce([
        {
          id: 501,
          fields: {
            'Microsoft.VSTS.TCM.Steps':
              '<steps><step id="2" type="ActionStep"><parameterizedString isformatted="true">Action</parameterizedString><parameterizedString isformatted="true">SR0501</parameterizedString></step></steps>',
          },
        },
      ]);
      jest.spyOn((resultDataProvider as any).testStepParserHelper, 'parseTestSteps').mockResolvedValueOnce([
        {
          stepId: '1',
          stepPosition: '1',
          action: 'Action',
          expected: 'SR0501',
          isSharedStepTitle: false,
        },
      ]);

      const result = await (resultDataProvider as any).getMewpInternalValidationFlatResults(
        '123',
        mockProjectName,
        [1]
      );

      expect((resultDataProvider as any).fetchWorkItemsByIds).toHaveBeenCalled();
      expect(result.rows).toHaveLength(1);
      expect(result.rows[0]).toEqual(
        expect.objectContaining({
          'Test Case ID': 501,
          'Mentioned but Not Linked': '',
          'Linked but Not Mentioned': '',
          'Validation Status': 'Pass',
        })
      );
    });
  });

  describe('buildLinkedRequirementsByTestCase', () => {
    it('should map linked requirements only for supported test-case requirement relation types', async () => {
      const requirements = [
        {
          workItemId: 7001,
          requirementId: 'SR0054-1',
          baseKey: 'SR0054',
          linkedTestCaseIds: [],
        },
      ];
      const testData = [
        {
          testCasesItems: [{ workItem: { id: 1001 } }],
          testPointsItems: [],
        },
      ];

      const fetchByIdsSpy = jest
        .spyOn(resultDataProvider as any, 'fetchWorkItemsByIds')
        .mockImplementation(async (...args: any[]) => {
          const ids = Array.isArray(args?.[1]) ? args[1] : [];
          const includeRelations = !!args?.[2];
          if (includeRelations) {
            return [
              {
                id: 1001,
                relations: [
                  {
                    rel: 'Microsoft.VSTS.Common.TestedBy-Reverse',
                    url: 'https://dev.azure.com/org/project/_apis/wit/workItems/7001',
                  },
                  {
                    rel: 'System.LinkTypes.Related',
                    url: 'https://dev.azure.com/org/project/_apis/wit/workItems/7001',
                  },
                ],
              },
            ];
          }
          return ids.map((id) => ({
            id,
            fields: {
              'System.WorkItemType': id === 7001 ? 'Requirement' : 'Test Case',
            },
          }));
        });

      const linked = await (resultDataProvider as any).buildLinkedRequirementsByTestCase(
        requirements,
        testData,
        mockProjectName
      );

      expect(fetchByIdsSpy).toHaveBeenCalled();
      expect(linked.get(1001)?.baseKeys?.has('SR0054')).toBe(true);
      expect(linked.get(1001)?.fullCodes?.has('SR0054-1')).toBe(true);
    });

    it('should ignore unsupported relation types when linking test case to requirements', async () => {
      const requirements = [
        {
          workItemId: 7002,
          requirementId: 'SR0099-1',
          baseKey: 'SR0099',
          linkedTestCaseIds: [],
        },
      ];
      const testData = [
        {
          testCasesItems: [{ workItem: { id: 1002 } }],
          testPointsItems: [],
        },
      ];

      jest.spyOn(resultDataProvider as any, 'fetchWorkItemsByIds').mockImplementation(async (...args: any[]) => {
        const ids = Array.isArray(args?.[1]) ? args[1] : [];
        const includeRelations = !!args?.[2];
        if (includeRelations) {
          return [
            {
              id: 1002,
              relations: [
                {
                  rel: 'System.LinkTypes.Related',
                  url: 'https://dev.azure.com/org/project/_apis/wit/workItems/7002',
                },
              ],
            },
          ];
        }
        return ids.map((id) => ({
          id,
          fields: {
            'System.WorkItemType': id === 7002 ? 'Requirement' : 'Test Case',
          },
        }));
      });

      const linked = await (resultDataProvider as any).buildLinkedRequirementsByTestCase(
        requirements,
        testData,
        mockProjectName
      );

      expect(linked.get(1002)?.baseKeys?.has('SR0099')).toBe(false);
      expect(linked.get(1002)?.fullCodes?.has('SR0099-1')).toBe(false);
    });
  });

  describe('MEWP release snapshot scoping', () => {
    it('should scope only to selected suites and avoid cross-Rel fallback selection', async () => {
      const suites = [
        { testSuiteId: 10, suiteName: 'Rel10 / Validation' },
        { testSuiteId: 11, suiteName: 'Rel11 / Validation' },
      ];
      const rawTestData = [
        {
          suiteName: 'Rel10 / Validation',
          testPointsItems: [
            { testCaseId: 501, lastRunId: 100, lastResultId: 200 },
            { testCaseId: 502, lastRunId: 300, lastResultId: 400 },
          ],
          testCasesItems: [{ workItem: { id: 501 } }, { workItem: { id: 502 } }],
        },
        {
          suiteName: 'Rel11 / Validation',
          testPointsItems: [
            { testCaseId: 501, lastRunId: 0, lastResultId: 0 },
            { testCaseId: 502, lastRunId: 500, lastResultId: 600 },
          ],
          testCasesItems: [{ workItem: { id: 501 } }, { workItem: { id: 502 } }],
        },
      ];

      const fetchSuitesSpy = jest
        .spyOn(resultDataProvider as any, 'fetchTestSuites')
        .mockResolvedValueOnce(suites);
      const fetchDataSpy = jest
        .spyOn(resultDataProvider as any, 'fetchTestData')
        .mockResolvedValueOnce(rawTestData);

      const scoped = await (resultDataProvider as any).fetchMewpScopedTestData(
        '123',
        mockProjectName,
        [11]
      );

      expect(fetchSuitesSpy).toHaveBeenCalledTimes(1);
      expect(fetchSuitesSpy).toHaveBeenCalledWith('123', mockProjectName, [11], true);
      expect(fetchDataSpy).toHaveBeenCalledTimes(1);
      expect(scoped).toEqual(rawTestData);
    });
  });

  describe('MEWP external ingestion validation/parsing', () => {
    const validBugsSource = {
      name: 'bugs.xlsx',
      url: 'https://minio.local/mewp-external-ingestion/MEWP/mewp-external-ingestion/bugs/bugs.xlsx',
      sourceType: 'mewpExternalIngestion',
    };
    const validL3L4Source = {
      name: 'l3l4.xlsx',
      url: 'https://minio.local/mewp-external-ingestion/MEWP/mewp-external-ingestion/l3l4/l3l4.xlsx',
      sourceType: 'mewpExternalIngestion',
    };

    it('should validate external files when required columns exist in A3 header row', async () => {
      const bugsBuffer = buildWorkbookBuffer([
        ['', '', '', '', ''],
        ['', '', '', '', ''],
        ['Elisra_SortIndex', 'SR', 'TargetWorkItemId', 'Title', 'TargetState'],
        ['101', 'SR0001', '9001', 'Bug one', 'Active'],
      ]);
      const l3l4Buffer = buildWorkbookBuffer([
        ['', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', ''],
        [
          'SR',
          'AREA 34',
          'TargetWorkItemId Level 3',
          'TargetTitleLevel3',
          'TargetStateLevel 3',
          'TargetWorkItemIdLevel 4',
          'TargetTitleLevel4',
          'TargetStateLevel 4',
        ],
        ['SR0001', 'Level 3', '7001', 'Req L3', 'Active', '', '', ''],
      ]);
      const axiosSpy = jest
        .spyOn(axios, 'get')
        .mockResolvedValueOnce({ data: bugsBuffer, headers: {} } as any)
        .mockResolvedValueOnce({ data: l3l4Buffer, headers: {} } as any);

      const result = await (resultDataProvider as any).validateMewpExternalFiles({
        externalBugsFile: validBugsSource,
        externalL3L4File: validL3L4Source,
      });

      expect(axiosSpy).toHaveBeenCalledTimes(2);
      expect(result.valid).toBe(true);
      expect(result.bugs).toEqual(
        expect.objectContaining({
          valid: true,
          headerRow: 'A3',
          matchedRequiredColumns: 5,
        })
      );
      expect(result.l3l4).toEqual(
        expect.objectContaining({
          valid: true,
          headerRow: 'A3',
          matchedRequiredColumns: 8,
        })
      );
    });

    it('should accept A1 fallback header row for backward compatibility', async () => {
      const bugsBuffer = buildWorkbookBuffer([
        ['Elisra_SortIndex', 'SR', 'TargetWorkItemId', 'Title', 'TargetState'],
        ['101', 'SR0001', '9001', 'Bug one', 'Active'],
      ]);
      jest.spyOn(axios, 'get').mockResolvedValueOnce({ data: bugsBuffer, headers: {} } as any);

      const result = await (resultDataProvider as any).validateMewpExternalFiles({
        externalBugsFile: validBugsSource,
      });

      expect(result.valid).toBe(true);
      expect(result.bugs).toEqual(
        expect.objectContaining({
          valid: true,
          headerRow: 'A1',
        })
      );
    });

    it('should fail validation when required columns are missing', async () => {
      const invalidBuffer = buildWorkbookBuffer([
        ['', '', ''],
        ['', '', ''],
        ['Elisra_SortIndex', 'TargetWorkItemId', 'TargetState'],
        ['101', '9001', 'Active'],
      ]);
      jest.spyOn(axios, 'get').mockResolvedValueOnce({ data: invalidBuffer, headers: {} } as any);

      const result = await (resultDataProvider as any).validateMewpExternalFiles({
        externalBugsFile: validBugsSource,
      });

      expect(result.valid).toBe(false);
      expect(result.bugs).toEqual(
        expect.objectContaining({
          valid: false,
          matchedRequiredColumns: 3,
          missingRequiredColumns: expect.arrayContaining(['SR', 'Title']),
        })
      );
    });

    it('should reject files from non-dedicated bucket/object path', async () => {
      const result = await (resultDataProvider as any).validateMewpExternalFiles({
        externalBugsFile: {
          name: 'bugs.xlsx',
          url: 'https://minio.local/mewp-external-ingestion/MEWP/other-prefix/bugs.xlsx',
          sourceType: 'mewpExternalIngestion',
        },
      });

      expect(result.valid).toBe(false);
      expect(result.bugs?.message || '').toContain('Invalid object path');
    });

    it('should filter external bugs by SR/state and dedupe by bug+requirement', async () => {
      const mewpExternalTableUtils = (resultDataProvider as any).mewpExternalTableUtils;
      jest.spyOn(mewpExternalTableUtils, 'loadExternalTableRows').mockResolvedValueOnce([
        {
          Elisra_SortIndex: '101',
          SR: 'SR0001',
          TargetWorkItemId: '9001',
          Title: 'Bug one',
          TargetState: 'Active',
          SAPWBS: 'ESUK',
        },
        {
          Elisra_SortIndex: '101',
          SR: 'SR0001',
          TargetWorkItemId: '9001',
          Title: 'Bug one duplicate',
          TargetState: 'Active',
        },
        {
          Elisra_SortIndex: '101',
          SR: '',
          TargetWorkItemId: '9002',
          Title: 'Missing SR',
          TargetState: 'Active',
        },
        {
          Elisra_SortIndex: '101',
          SR: 'SR0002',
          TargetWorkItemId: '9003',
          Title: 'Closed bug',
          TargetState: 'Closed',
        },
      ]);

      const map = await (resultDataProvider as any).loadExternalBugsByTestCase(validBugsSource);
      const bugs = map.get(101) || [];

      expect(bugs).toHaveLength(1);
      expect(bugs[0]).toEqual(
        expect.objectContaining({
          id: 9001,
          requirementBaseKey: 'SR0001',
          responsibility: 'ESUK',
        })
      );
    });

    it('should resolve external bug responsibility from AreaPath columns when SAPWBS is empty', async () => {
      const mewpExternalTableUtils = (resultDataProvider as any).mewpExternalTableUtils;
      jest.spyOn(mewpExternalTableUtils, 'loadExternalTableRows').mockResolvedValueOnce([
        {
          Elisra_SortIndex: '101',
          SR: 'SR0001',
          TargetWorkItemId: '9011',
          Title: 'Bug from ATP\\ESUK path',
          TargetState: 'Active',
          SAPWBS: '',
          AreaPath: 'MEWP\\Customer Requirements\\Level 2\\ATP\\ESUK',
        },
        {
          Elisra_SortIndex: '102',
          SR: 'SR0002',
          TargetWorkItemId: '9012',
          Title: 'Bug from ATP path',
          TargetState: 'Active',
          SAPWBS: '',
          'System.AreaPath': 'MEWP\\Customer Requirements\\Level 2\\ATP',
        },
      ]);

      const map = await (resultDataProvider as any).loadExternalBugsByTestCase(validBugsSource);
      const bugs101 = map.get(101) || [];
      const bugs102 = map.get(102) || [];

      expect(bugs101).toHaveLength(1);
      expect(bugs102).toHaveLength(1);
      expect(bugs101[0].responsibility).toBe('ESUK');
      expect(bugs102[0].responsibility).toBe('Elisra');
    });

    it('should require Elisra_SortIndex and ignore rows that only provide WorkItemId', async () => {
      const mewpExternalTableUtils = (resultDataProvider as any).mewpExternalTableUtils;
      jest.spyOn(mewpExternalTableUtils, 'loadExternalTableRows').mockResolvedValueOnce([
        {
          WorkItemId: '101',
          SR: 'SR0001',
          TargetWorkItemId: '9010',
          Title: 'Bug without Elisra_SortIndex',
          TargetState: 'Active',
        },
      ]);

      const map = await (resultDataProvider as any).loadExternalBugsByTestCase(validBugsSource);
      expect(map.size).toBe(0);
    });

    it('should parse external L3/L4 file using AREA 34 semantics and terminal-state filtering', async () => {
      const mewpExternalTableUtils = (resultDataProvider as any).mewpExternalTableUtils;
      jest.spyOn(mewpExternalTableUtils, 'loadExternalTableRows').mockResolvedValueOnce([
        {
          SR: 'SR0001',
          'AREA 34': 'Level 4',
          'TargetWorkItemId Level 3': '7001',
          TargetTitleLevel3: 'L4 From Level3 Column',
          'TargetStateLevel 3': 'Active',
        },
        {
          SR: 'SR0001',
          'AREA 34': 'Level 3',
          'TargetWorkItemId Level 3': '7002',
          TargetTitleLevel3: 'L3 Requirement',
          'TargetStateLevel 3': 'Active',
          'TargetWorkItemIdLevel 4': '7003',
          TargetTitleLevel4: 'L4 Requirement',
          'TargetStateLevel 4': 'Closed',
        },
      ]);

      const map = await (resultDataProvider as any).loadExternalL3L4ByBaseKey(validL3L4Source);
      expect(map.get('SR0001')).toEqual([
        { l3Id: '', l3Title: '', l4Id: '7001', l4Title: 'L4 From Level3 Column' },
        { l3Id: '7002', l3Title: 'L3 Requirement', l4Id: '', l4Title: '' },
      ]);
    });

    it('should emit paired L3+L4 when AREA 34 is Level 4 and both level columns are present', async () => {
      const mewpExternalTableUtils = (resultDataProvider as any).mewpExternalTableUtils;
      jest.spyOn(mewpExternalTableUtils, 'loadExternalTableRows').mockResolvedValueOnce([
        {
          SR: 'SR0001',
          'AREA 34': 'Level 4',
          'TargetWorkItemId Level 3': '7401',
          TargetTitleLevel3: 'L3 In Level4 Row',
          'TargetStateLevel 3': 'Active',
          'TargetWorkItemIdLevel 4': '8401',
          TargetTitleLevel4: 'L4 In Level4 Row',
          'TargetStateLevel 4': 'Active',
        },
      ]);

      const map = await (resultDataProvider as any).loadExternalL3L4ByBaseKey(validL3L4Source);
      expect(map.get('SR0001')).toEqual([
        { l3Id: '7401', l3Title: 'L3 In Level4 Row', l4Id: '8401', l4Title: 'L4 In Level4 Row' },
      ]);
    });

    it('should exclude external open L3/L4 rows when SAPWBS resolves to ESUK', async () => {
      const mewpExternalTableUtils = (resultDataProvider as any).mewpExternalTableUtils;
      jest.spyOn(mewpExternalTableUtils, 'loadExternalTableRows').mockResolvedValueOnce([
        {
          SR: 'SR0001',
          'AREA 34': 'Level 3',
          'TargetWorkItemId Level 3': '7101',
          TargetTitleLevel3: 'L3 ESUK',
          'TargetStateLevel 3': 'Active',
          'TargetSapWbsLevel 3': 'ESUK',
          'TargetWorkItemIdLevel 4': '7201',
          TargetTitleLevel4: 'L4 IL',
          'TargetStateLevel 4': 'Active',
          'TargetSapWbsLevel 4': 'IL',
        },
        {
          SR: 'SR0001',
          'AREA 34': 'Level 3',
          'TargetWorkItemId Level 3': '7102',
          TargetTitleLevel3: 'L3 IL',
          'TargetStateLevel 3': 'Active',
          'TargetSapWbsLevel 3': 'IL',
        },
      ]);

      const map = await (resultDataProvider as any).loadExternalL3L4ByBaseKey(validL3L4Source);
      expect(map.get('SR0001')).toEqual([
        { l3Id: '', l3Title: '', l4Id: '7201', l4Title: 'L4 IL' },
        { l3Id: '7102', l3Title: 'L3 IL', l4Id: '', l4Title: '' },
      ]);
    });

    it('should fallback L3/L4 SAPWBS exclusion from SR-mapped requirement when row SAPWBS is empty', async () => {
      const mewpExternalTableUtils = (resultDataProvider as any).mewpExternalTableUtils;
      jest.spyOn(mewpExternalTableUtils, 'loadExternalTableRows').mockResolvedValueOnce([
        {
          SR: 'SR0001',
          'AREA 34': 'Level 3',
          'TargetWorkItemId Level 3': '7301',
          TargetTitleLevel3: 'L3 From ESUK Requirement',
          'TargetStateLevel 3': 'Active',
          'TargetSapWbsLevel 3': '',
        },
        {
          SR: 'SR0002',
          'AREA 34': 'Level 3',
          'TargetWorkItemId Level 3': '7302',
          TargetTitleLevel3: 'L3 From IL Requirement',
          'TargetStateLevel 3': 'Active',
          'TargetSapWbsLevel 3': '',
        },
      ]);

      const map = await (resultDataProvider as any).loadExternalL3L4ByBaseKey(
        validL3L4Source,
        new Map([
          ['SR0001', 'ESUK'],
          ['SR0002', 'IL'],
        ])
      );

      expect(map.has('SR0001')).toBe(false);
      expect(map.get('SR0002')).toEqual([
        { l3Id: '7302', l3Title: 'L3 From IL Requirement', l4Id: '', l4Title: '' },
      ]);
    });

    it('should resolve bug responsibility from AreaPath when SAPWBS is empty', () => {
      const fromEsukAreaPath = (resultDataProvider as any).resolveBugResponsibility({
        'Custom.SAPWBS': '',
        'System.AreaPath': 'MEWP\\Customer Requirements\\Level 2\\ATP\\ESUK',
      });
      const fromIlAreaPath = (resultDataProvider as any).resolveBugResponsibility({
        'Custom.SAPWBS': '',
        'System.AreaPath': 'MEWP\\Customer Requirements\\Level 2\\ATP',
      });
      const unknown = (resultDataProvider as any).resolveBugResponsibility({
        'Custom.SAPWBS': '',
        'System.AreaPath': 'MEWP\\Other\\Area',
      });

      expect(fromEsukAreaPath).toBe('ESUK');
      expect(fromIlAreaPath).toBe('Elisra');
      expect(unknown).toBe('Unknown');
    });

    it('should handle 1000 external bug rows and keep only in-scope parsed items', async () => {
      const rows: any[] = [];
      for (let i = 1; i <= 1000; i++) {
        rows.push({
          Elisra_SortIndex: String(4000 + (i % 20)),
          SR: `SR${String(6000 + (i % 50)).padStart(4, '0')}`,
          TargetWorkItemId: String(900000 + i),
          Title: `Bug ${i}`,
          TargetState: i % 10 === 0 ? 'Closed' : 'Active',
          SAPWBS: i % 2 === 0 ? 'ESUK' : 'IL',
        });
      }

      const mewpExternalTableUtils = (resultDataProvider as any).mewpExternalTableUtils;
      jest.spyOn(mewpExternalTableUtils, 'loadExternalTableRows').mockResolvedValueOnce(rows);

      const startedAt = Date.now();
      const map = await (resultDataProvider as any).loadExternalBugsByTestCase({
        name: 'bulk-bugs.xlsx',
        url: 'https://minio.local/mewp-external-ingestion/MEWP/mewp-external-ingestion/bugs/bulk-bugs.xlsx',
        sourceType: 'mewpExternalIngestion',
      });
      const elapsedMs = Date.now() - startedAt;

      const totalParsed = [...map.values()].reduce((sum, items) => sum + (items?.length || 0), 0);
      expect(totalParsed).toBe(900); // every 10th row is closed and filtered out
      expect(elapsedMs).toBeLessThan(5000);
    });

    it('should handle 1000 external L3/L4 rows and map all active links', async () => {
      const rows: any[] = [];
      for (let i = 1; i <= 1000; i++) {
        rows.push({
          SR: `SR${String(7000 + (i % 25)).padStart(4, '0')}`,
          'AREA 34': i % 3 === 0 ? 'Level 4' : 'Level 3',
          'TargetWorkItemId Level 3': String(800000 + i),
          TargetTitleLevel3: `L3/L4 Title ${i}`,
          'TargetStateLevel 3': i % 11 === 0 ? 'Resolved' : 'Active',
          'TargetWorkItemIdLevel 4': String(810000 + i),
          TargetTitleLevel4: `L4 Title ${i}`,
          'TargetStateLevel 4': i % 13 === 0 ? 'Closed' : 'Active',
        });
      }

      const mewpExternalTableUtils = (resultDataProvider as any).mewpExternalTableUtils;
      jest.spyOn(mewpExternalTableUtils, 'loadExternalTableRows').mockResolvedValueOnce(rows);

      const startedAt = Date.now();
      const map = await (resultDataProvider as any).loadExternalL3L4ByBaseKey({
        name: 'bulk-l3l4.xlsx',
        url: 'https://minio.local/mewp-external-ingestion/MEWP/mewp-external-ingestion/l3l4/bulk-l3l4.xlsx',
        sourceType: 'mewpExternalIngestion',
      });
      const elapsedMs = Date.now() - startedAt;

      const totalLinks = [...map.values()].reduce((sum, items) => sum + (items?.length || 0), 0);
      expect(totalLinks).toBeGreaterThan(700);
      expect(elapsedMs).toBeLessThan(5000);
    });
  });

  describe('MEWP high-volume requirement token parsing', () => {
    it('should parse 1000 expected-result requirement tokens with noisy fragments', () => {
      const tokens: string[] = [];
      for (let i = 1; i <= 1000; i++) {
        const code = `SR${String(10000 + i)}`;
        tokens.push(code);
        if (i % 100 === 0) tokens.push(`${code} V3.24`);
        if (i % 125 === 0) tokens.push(`${code} VVRM2425`);
      }
      const sourceText = tokens.join('; ');

      const startedAt = Date.now();
      const codes = (resultDataProvider as any).extractRequirementCodesFromText(sourceText) as Set<string>;
      const elapsedMs = Date.now() - startedAt;

      expect(codes.size).toBe(1000);
      expect(codes.has('SR10001')).toBe(true);
      expect(codes.has('SR11000')).toBe(true);
      expect(elapsedMs).toBeLessThan(3000);
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

    it('should fetch no-run test case by suite test-case revision when provided', async () => {
      const point = {
        testCaseId: '123',
        testCaseName: 'TC 123',
        outcome: 'passed',
        suiteTestCase: {
          workItem: {
            id: 123,
            rev: 9,
            workItemFields: [{ key: 'Microsoft.VSTS.TCM.Steps', value: '<steps></steps>' }],
          },
        },
        testSuite: { id: '1', name: 'Suite' },
      };

      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce({
        id: 123,
        rev: 9,
        fields: {
          'System.State': 'Active',
          'System.CreatedDate': '2024-01-01T00:00:00',
          'Microsoft.VSTS.TCM.Priority': 1,
          'System.Title': 'Title 123',
          'Microsoft.VSTS.TCM.Steps': '<steps></steps>',
        },
        relations: null,
      });

      const res = await (resultDataProvider as any).fetchResultDataBasedOnWiBase(
        mockProjectName,
        '0',
        '0',
        true,
        [],
        false,
        point,
        false,
        true
      );

      expect(TFSServices.getItemContent).toHaveBeenCalledWith(
        expect.stringContaining('/_apis/wit/workItems/123/revisions/9?$expand=all'),
        mockToken
      );
      const calledUrls = (TFSServices.getItemContent as jest.Mock).mock.calls.map((args: any[]) => String(args[0]));
      expect(calledUrls.some((url: string) => url.includes('?asOf='))).toBe(false);
      expect(res).toEqual(expect.objectContaining({ testCaseRevision: 9 }));
    });

    it('should ignore point asOf timestamp when runless asOf mode is disabled', async () => {
      const point = {
        testCaseId: '123',
        testCaseName: 'TC 123',
        outcome: 'passed',
        pointAsOfTimestamp: '2025-01-01T12:34:56Z',
        suiteTestCase: {
          workItem: {
            id: 123,
            rev: 9,
            workItemFields: [{ key: 'Microsoft.VSTS.TCM.Steps', value: '<steps></steps>' }],
          },
        },
        testSuite: { id: '1', name: 'Suite' },
      };

      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce({
        id: 123,
        rev: 9,
        fields: {
          'System.State': 'Active',
          'System.CreatedDate': '2024-01-01T00:00:00',
          'Microsoft.VSTS.TCM.Priority': 1,
          'System.Title': 'Title 123',
          'Microsoft.VSTS.TCM.Steps': '<steps></steps>',
        },
        relations: null,
      });

      const res = await (resultDataProvider as any).fetchResultDataBasedOnWiBase(
        mockProjectName,
        '0',
        '0',
        true,
        [],
        false,
        point
      );

      const calledUrls = (TFSServices.getItemContent as jest.Mock).mock.calls.map((args: any[]) => String(args[0]));
      expect(calledUrls.some((url: string) => url.includes('/_apis/wit/workItems/123?asOf='))).toBe(false);
      expect(calledUrls.some((url: string) => url.includes('/_apis/wit/workItems/123/revisions/9'))).toBe(true);
      expect(res).toEqual(expect.objectContaining({ testCaseRevision: 9 }));
    });

    it('should fetch no-run test case by asOf timestamp when pointAsOfTimestamp is available', async () => {
      (TFSServices.getItemContent as jest.Mock).mockReset();
      const point = {
        testCaseId: '123',
        testCaseName: 'TC 123',
        outcome: 'passed',
        pointAsOfTimestamp: '2025-01-01T12:34:56Z',
        suiteTestCase: {
          workItem: {
            id: 123,
            rev: 9,
            workItemFields: [{ key: 'Microsoft.VSTS.TCM.Steps', value: '<steps></steps>' }],
          },
        },
        testSuite: { id: '1', name: 'Suite' },
      };

      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce({
        id: 123,
        rev: 6,
        fields: {
          'System.State': 'Active',
          'System.CreatedDate': '2024-01-01T00:00:00',
          'Microsoft.VSTS.TCM.Priority': 1,
          'System.Title': 'Title 123',
          'Microsoft.VSTS.TCM.Steps': '<steps></steps>',
        },
        relations: null,
      });

      const res = await (resultDataProvider as any).fetchResultDataBasedOnWiBase(
        mockProjectName,
        '0',
        '0',
        true,
        [],
        false,
        point,
        false,
        true
      );

      const calledUrls = (TFSServices.getItemContent as jest.Mock).mock.calls.map((args: any[]) => String(args[0]));
      expect(calledUrls.some((url: string) => url.includes('/_apis/wit/workItems/123?asOf='))).toBe(true);
      expect(calledUrls.some((url: string) => url.includes('/revisions/9'))).toBe(false);
      expect(res).toEqual(expect.objectContaining({ testCaseRevision: 6 }));
    });

    it('should fallback from asOf snapshot without steps to revision snapshot with steps', async () => {
      (TFSServices.getItemContent as jest.Mock).mockReset();
      const point = {
        testCaseId: '777',
        testCaseName: 'TC 777',
        outcome: 'Not Run',
        pointAsOfTimestamp: '2025-03-01T00:00:00Z',
        suiteTestCase: {
          workItem: {
            id: 777,
            workItemFields: [{ key: 'System.Rev', value: '21' }],
          },
        },
        testSuite: { id: '1', name: 'Suite' },
      };

      (TFSServices.getItemContent as jest.Mock)
        .mockResolvedValueOnce({
          id: 777,
          rev: 18,
          fields: {
            'System.State': 'Design',
            'System.CreatedDate': '2024-01-01T00:00:00',
            'Microsoft.VSTS.TCM.Priority': 1,
            'System.Title': 'TC 777',
          },
          relations: [],
        })
        .mockResolvedValueOnce({
          id: 777,
          rev: 21,
          fields: {
            'System.State': 'Design',
            'System.CreatedDate': '2024-01-01T00:00:00',
            'Microsoft.VSTS.TCM.Priority': 1,
            'System.Title': 'TC 777',
            'Microsoft.VSTS.TCM.Steps': '<steps></steps>',
          },
          relations: [],
        });

      const res = await (resultDataProvider as any).fetchResultDataBasedOnWiBase(
        mockProjectName,
        '0',
        '0',
        true,
        [],
        false,
        point,
        false,
        true
      );

      const calledUrls = (TFSServices.getItemContent as jest.Mock).mock.calls.map((args: any[]) => String(args[0]));
      expect(calledUrls.some((url: string) => url.includes('/_apis/wit/workItems/777?asOf='))).toBe(true);
      expect(calledUrls.some((url: string) => url.includes('/_apis/wit/workItems/777/revisions/21'))).toBe(true);
      expect(res).toEqual(
        expect.objectContaining({
          testCaseRevision: 21,
          stepsResultXml: '<steps></steps>',
        })
      );
    });

    it('should fallback to suite revision when asOf fetch fails', async () => {
      (TFSServices.getItemContent as jest.Mock).mockReset();
      const point = {
        testCaseId: '456',
        testCaseName: 'TC 456',
        outcome: 'Not Run',
        pointAsOfTimestamp: '2025-02-01T00:00:00Z',
        suiteTestCase: {
          workItem: {
            id: 456,
            workItemFields: [{ key: 'System.Rev', value: '11' }],
          },
        },
        testSuite: { id: '1', name: 'Suite' },
      };

      (TFSServices.getItemContent as jest.Mock)
        .mockRejectedValueOnce(new Error('asOf failed'))
        .mockResolvedValueOnce({
          id: 456,
          rev: 11,
          fields: {
            'System.State': 'Active',
            'System.CreatedDate': '2024-01-01T00:00:00',
            'Microsoft.VSTS.TCM.Priority': 1,
            'System.Title': 'Title 456',
            'Microsoft.VSTS.TCM.Steps': '<steps></steps>',
          },
          relations: [],
        });

      const res = await (resultDataProvider as any).fetchResultDataBasedOnWiBase(
        mockProjectName,
        '0',
        '0',
        true,
        [],
        false,
        point,
        false,
        true
      );

      const calledUrls = (TFSServices.getItemContent as jest.Mock).mock.calls.map((args: any[]) => String(args[0]));
      expect(calledUrls.some((url: string) => url.includes('/_apis/wit/workItems/456?asOf='))).toBe(true);
      expect(calledUrls.some((url: string) => url.includes('/_apis/wit/workItems/456/revisions/11'))).toBe(true);
      expect(res).toEqual(expect.objectContaining({ testCaseRevision: 11 }));
    });

    it('should resolve no-run revision from System.Rev in suite test-case fields', async () => {
      (TFSServices.getItemContent as jest.Mock).mockReset();
      const point = {
        testCaseId: '321',
        testCaseName: 'TC 321',
        outcome: 'Not Run',
        suiteTestCase: {
          workItem: {
            id: 321,
            workItemFields: [
              { key: 'System.Rev', value: '13' },
              { key: 'Microsoft.VSTS.TCM.Steps', value: '<steps></steps>' },
            ],
          },
        },
        testSuite: { id: '1', name: 'Suite' },
      };

      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce({
        id: 321,
        rev: 13,
        fields: {
          'System.State': 'Design',
          'System.CreatedDate': '2024-05-01T00:00:00',
          'Microsoft.VSTS.TCM.Priority': 1,
          'System.Title': 'Title 321',
          'Microsoft.VSTS.TCM.Steps': '<steps></steps>',
        },
        relations: [],
      });

      const res = await (resultDataProvider as any).fetchResultDataBasedOnWiBase(
        mockProjectName,
        '0',
        '0',
        true,
        [],
        false,
        point
      );

      expect(TFSServices.getItemContent).toHaveBeenCalledWith(
        expect.stringContaining('/_apis/wit/workItems/321/revisions/13?$expand=all'),
        mockToken
      );
      expect(res).toEqual(expect.objectContaining({ testCaseRevision: 13 }));
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

    it('should keep points without run/result IDs when test reporter mode is enabled', async () => {
      const testData = [
        {
          testSuiteId: 1,
          testPointsItems: [{ testCaseId: 10, lastRunId: 101, lastResultId: 201 }, { testCaseId: 11 }],
        },
      ];
      const fetchStrategy = jest
        .fn()
        .mockResolvedValueOnce({ testCaseId: 10 })
        .mockResolvedValueOnce({ testCaseId: 11 });

      const result = await (resultDataProvider as any).fetchAllResultDataBase(
        testData,
        mockProjectName,
        true,
        fetchStrategy
      );

      expect(fetchStrategy).toHaveBeenCalledTimes(2);
      expect(fetchStrategy).toHaveBeenNthCalledWith(
        2,
        mockProjectName,
        1,
        expect.objectContaining({ testCaseId: 11 })
      );
      expect(result).toEqual([{ testCaseId: 10 }, { testCaseId: 11 }]);
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

    it('should call fetch method with runId/resultId as 0 when point has no run history', async () => {
      const point = { testCaseId: 15, lastRunId: undefined, lastResultId: undefined };
      const fetchResultMethod = jest.fn().mockResolvedValue({
        testCase: { id: 15, name: 'TC 15' },
        testSuite: { name: 'S' },
        iterationDetails: [],
      });
      const createResponseObject = jest.fn().mockReturnValue({ id: 15 });

      await (resultDataProvider as any).fetchResultDataBase(
        mockProjectName,
        'suite-no-runs',
        point,
        fetchResultMethod,
        createResponseObject
      );

      expect(fetchResultMethod).toHaveBeenCalledWith(mockProjectName, '0', '0');
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

    it('should return test-level row with empty step fields when suite has no run history', async () => {
      jest.spyOn(resultDataProvider as any, 'fetchTestPlanName').mockResolvedValueOnce('Plan 12');
      jest.spyOn(resultDataProvider as any, 'fetchTestSuites').mockResolvedValueOnce([
        {
          testSuiteId: 300,
          suiteId: 300,
          suiteName: 'suite no runs',
          parentSuiteId: 100,
          parentSuiteName: 'Rel3',
          suitePath: 'Root/Rel3/suite no runs',
          testGroupName: 'suite no runs',
        },
      ]);

      jest.spyOn(resultDataProvider as any, 'fetchTestData').mockResolvedValueOnce([
        {
          testSuiteId: 300,
          suiteId: 300,
          suiteName: 'suite no runs',
          parentSuiteId: 100,
          parentSuiteName: 'Rel3',
          suitePath: 'Root/Rel3/suite no runs',
          testGroupName: 'suite no runs',
          testPointsItems: [
            {
              testCaseId: 55,
              testCaseName: 'TC 55',
              outcome: 'Not Run',
              testPointId: 9001,
              lastRunId: undefined,
              lastResultId: undefined,
              lastResultDetails: undefined,
            },
          ],
          testCasesItems: [
            {
              workItem: {
                id: 55,
                workItemFields: [{ key: 'System.Rev', value: 4 }],
              },
            },
          ],
        },
      ]);

      jest.spyOn(resultDataProvider as any, 'fetchAllResultDataTestReporter').mockResolvedValueOnce([
        {
          testCaseId: 55,
          testCase: { id: 55, name: 'TC 55' },
          testSuite: { name: 'suite no runs' },
          executionDate: '',
          testCaseResult: { resultMessage: 'Not Run', url: '' },
          customFields: {},
          runBy: '',
          iteration: undefined,
          lastRunId: undefined,
          lastResultId: undefined,
        },
      ]);

      const result = await resultDataProvider.getTestReporterFlatResults(
        mockTestPlanId,
        mockProjectName,
        undefined,
        [],
        false
      );

      expect(result.rows).toHaveLength(1);
      expect(result.rows[0]).toEqual(
        expect.objectContaining({
          testCaseId: 55,
          testRunId: undefined,
          testPointId: 9001,
          stepOutcome: undefined,
          stepStepIdentifier: '',
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
