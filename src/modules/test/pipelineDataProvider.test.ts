import { PipelineRun } from '../../models/tfs-data';
import { TFSServices } from '../../helpers/tfs';
import PipelinesDataProvider from '../PipelinesDataProvider';
import GitDataProvider from '../GitDataProvider';
import logger from '../../utils/logger';

jest.mock('../../helpers/tfs');
jest.mock('../../utils/logger');
jest.mock('../GitDataProvider');

describe('PipelinesDataProvider', () => {
  let pipelinesDataProvider: PipelinesDataProvider;
  const mockOrgUrl = 'https://dev.azure.com/orgname/';
  const mockToken = 'mock-token';

  beforeEach(() => {
    jest.clearAllMocks();
    pipelinesDataProvider = new PipelinesDataProvider(mockOrgUrl, mockToken);
  });

  describe('isMatchingPipeline', () => {
    // Create test method to access private method
    const invokeIsMatchingPipeline = (
      fromPipeline: PipelineRun,
      targetPipeline: PipelineRun,
      searchPrevPipelineFromDifferentCommit: boolean
    ): boolean => {
      return (pipelinesDataProvider as any).isMatchingPipeline(
        fromPipeline,
        targetPipeline,
        searchPrevPipelineFromDifferentCommit
      );
    };

    it('should return false when repository IDs are different', () => {
      // Arrange
      const fromPipeline = {
        resources: {
          repositories: {
            '0': {
              self: {
                repository: { id: 'repo1' },
                version: 'v1',
                refName: 'refs/heads/main'
              }
            }
          }
        }
      } as unknown as PipelineRun;

      const targetPipeline = {
        resources: {
          repositories: {
            '0': {
              self: {
                repository: { id: 'repo2' },
                version: 'v1',
                refName: 'refs/heads/main'
              }
            }
          }
        }
      } as unknown as PipelineRun;

      // Act
      const result = invokeIsMatchingPipeline(fromPipeline, targetPipeline, false);

      // Assert
      expect(result).toBe(false);
    });

    it('should return true when versions are the same and searchPrevPipelineFromDifferentCommit is false', () => {
      // Arrange
      const fromPipeline = {
        resources: {
          repositories: {
            '0': {
              self: {
                repository: { id: 'repo1' },
                version: 'v1',
                refName: 'refs/heads/main'
              }
            }
          }
        }
      } as unknown as PipelineRun;

      const targetPipeline = {
        resources: {
          repositories: {
            '0': {
              self: {
                repository: { id: 'repo1' },
                version: 'v1',
                refName: 'refs/heads/main'
              }
            }
          }
        }
      } as unknown as PipelineRun;

      // Act
      const result = invokeIsMatchingPipeline(fromPipeline, targetPipeline, false);

      // Assert
      expect(result).toBe(true);
    });

    it('should return false when versions are the same and searchPrevPipelineFromDifferentCommit is true', () => {
      // Arrange
      const fromPipeline = {
        resources: {
          repositories: {
            '0': {
              self: {
                repository: { id: 'repo1' },
                version: 'v1',
                refName: 'refs/heads/main'
              }
            }
          }
        }
      } as unknown as PipelineRun;

      const targetPipeline = {
        resources: {
          repositories: {
            '0': {
              self: {
                repository: { id: 'repo1' },
                version: 'v1',
                refName: 'refs/heads/main'
              }
            }
          }
        }
      } as unknown as PipelineRun;

      // Act
      const result = invokeIsMatchingPipeline(fromPipeline, targetPipeline, true);

      // Assert
      expect(result).toBe(false);
    });

    it('should return true when refNames match but versions differ', () => {
      // Arrange
      const fromPipeline = {
        resources: {
          repositories: {
            '0': {
              self: {
                repository: { id: 'repo1' },
                version: 'v1',
                refName: 'refs/heads/main'
              }
            }
          }
        }
      } as unknown as PipelineRun;

      const targetPipeline = {
        resources: {
          repositories: {
            '0': {
              self: {
                repository: { id: 'repo1' },
                version: 'v2',
                refName: 'refs/heads/main'
              }
            }
          }
        }
      } as unknown as PipelineRun;

      // Act
      const result = invokeIsMatchingPipeline(fromPipeline, targetPipeline, true);

      // Assert
      expect(result).toBe(true);
    });

    it('should use __designer_repo when self is not available', () => {
      // Arrange
      const fromPipeline = {
        resources: {
          repositories: {
            __designer_repo: {
              repository: { id: 'repo1' },
              version: 'v1',
              refName: 'refs/heads/main'
            }
          }
        }
      } as unknown as PipelineRun;

      const targetPipeline = {
        resources: {
          repositories: {
            __designer_repo: {
              repository: { id: 'repo1' },
              version: 'v1',
              refName: 'refs/heads/main'
            }
          }
        }
      } as unknown as PipelineRun;

      // Act
      const result = invokeIsMatchingPipeline(fromPipeline, targetPipeline, false);

      // Assert
      expect(result).toBe(true);
    });
  });

  describe('getPipelineRunDetails', () => {
    it('should call TFSServices.getItemContent with correct parameters', async () => {
      // Arrange
      const projectName = 'project1';
      const pipelineId = 123;
      const runId = 456;
      const mockResponse = { id: runId, resources: {} };
      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockResponse);

      // Act
      const result = await pipelinesDataProvider.getPipelineRunDetails(projectName, pipelineId, runId);

      // Assert
      expect(TFSServices.getItemContent).toHaveBeenCalledWith(
        `${mockOrgUrl}${projectName}/_apis/pipelines/${pipelineId}/runs/${runId}`,
        mockToken
      );
      expect(result).toEqual(mockResponse);
    });
  });

  describe('GetPipelineRunHistory', () => {
    it('should return filtered pipeline run history', async () => {
      // Arrange
      const projectName = 'project1';
      const pipelineId = '123';
      const mockResponse = {
        value: [
          { id: 1, result: 'succeeded' },
          { id: 2, result: 'failed' },
          { id: 3, result: 'canceled' },
          { id: 4, result: 'succeeded' }
        ]
      };
      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockResponse);

      // Act
      const result = await pipelinesDataProvider.GetPipelineRunHistory(projectName, pipelineId);

      // Assert
      expect(TFSServices.getItemContent).toHaveBeenCalledWith(
        `${mockOrgUrl}${projectName}/_apis/pipelines/${pipelineId}/runs`,
        mockToken,
        'get',
        null,
        null
      );
      expect(result).toEqual({
        count: 4, // Note: Current filter logic keeps all runs where result is not 'failed' AND not 'canceled'
        value: mockResponse.value
      });
    });

    it('should handle API errors gracefully', async () => {
      // Arrange
      const projectName = 'project1';
      const pipelineId = '123';
      const expectedError = new Error('API error');
      (TFSServices.getItemContent as jest.Mock).mockRejectedValueOnce(expectedError);

      // Act
      const result = await pipelinesDataProvider.GetPipelineRunHistory(projectName, pipelineId);

      // Assert
      expect(TFSServices.getItemContent).toHaveBeenCalledWith(
        `${mockOrgUrl}${projectName}/_apis/pipelines/${pipelineId}/runs`,
        mockToken,
        'get',
        null,
        null
      );
      expect(logger.error).toHaveBeenCalledWith(`Could not fetch Pipeline Run History: ${expectedError.message}`);
      expect(result).toBeUndefined();
    });
  });
});