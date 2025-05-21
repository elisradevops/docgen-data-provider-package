import { TFSServices } from '../../helpers/tfs';
import ResultDataProvider from '../ResultDataProvider';
import logger from '../../utils/logger';
import Utils from '../../utils/testStepParserHelper';

// Mock dependencies
jest.mock('../../helpers/tfs');
jest.mock('../../utils/logger');
jest.mock('../../utils/testStepParserHelper');
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

    describe('setRunStatus', () => {
      it('should return empty for shared step titles with Unspecified outcome', () => {
        // Arrange
        const actionResult = { outcome: 'Unspecified', isSharedStepTitle: true };

        // Act
        const result = (resultDataProvider as any).setRunStatus(actionResult);

        // Assert
        expect(result).toBe('');
      });

      it('should return "Not Run" for Unspecified outcome on regular steps', () => {
        // Arrange
        const actionResult = { outcome: 'Unspecified', isSharedStepTitle: false };

        // Act
        const result = (resultDataProvider as any).setRunStatus(actionResult);

        // Assert
        expect(result).toBe('Not Run');
      });

      it('should return the outcome for non-Unspecified outcomes', () => {
        // Arrange
        const actionResult = { outcome: 'Failed', isSharedStepTitle: false };

        // Act
        const result = (resultDataProvider as any).setRunStatus(actionResult);

        // Assert
        expect(result).toBe('Failed');
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
        const result = (resultDataProvider as any).mapTestPoint(testPoint);

        // Assert
        expect(result).toEqual({
          testCaseId: 1,
          testCaseName: 'Test Case 1',
          configurationName: 'Config 1',
          outcome: 'passed',
          lastRunId: 100,
          lastResultId: 200,
          lastResultDetails: { dateCompleted: '2023-01-01', runBy: { displayName: 'Test User' } },
        });
      });

      it('should handle missing fields', () => {
        // Arrange
        const testPoint = {
          testCaseReference: { id: 1, name: 'Test Case 1' },
          // No configuration or results
        };

        // Act
        const result = (resultDataProvider as any).mapTestPoint(testPoint);

        // Assert
        expect(result).toEqual({
          testCaseId: 1,
          testCaseName: 'Test Case 1',
          configurationName: undefined,
          outcome: 'Not Run',
          lastRunId: undefined,
          lastResultId: undefined,
          lastResultDetails: undefined,
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

  describe('getCombinedResultsSummary', () => {
    it('should combine all results into expected format', async () => {
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
});
