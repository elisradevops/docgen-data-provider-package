import { PipelineRun, Repository, ResourceRepository } from '../../models/tfs-data';
import { TFSServices } from '../../helpers/tfs';
import PipelinesDataProvider from '../../modules/PipelinesDataProvider';
import GitDataProvider from '../../modules/GitDataProvider';
import logger from '../../utils/logger';

jest.mock('../../helpers/tfs');
jest.mock('../../utils/logger');
jest.mock('../../modules/GitDataProvider');

describe('PipelinesDataProvider', () => {
  let pipelinesDataProvider: PipelinesDataProvider;
  const mockOrgUrl = 'https://dev.azure.com/orgname/';
  const mockToken = 'mock-token';

  beforeEach(() => {
    jest.clearAllMocks();
    pipelinesDataProvider = new PipelinesDataProvider(mockOrgUrl, mockToken);
  });

  describe('GetBuildWorkItems', () => {
    it('should return work item references associated with a single build', async () => {
      (TFSServices.getItemContent as jest.Mock).mockResolvedValue({
        value: [{ id: '1', url: 'https://example.test/wi/1' }],
      });

      const result = await pipelinesDataProvider.GetBuildWorkItems('project1', 123);

      expect(TFSServices.getItemContent).toHaveBeenCalledWith(
        'https://dev.azure.com/orgname/project1/_apis/build/builds/123/workitems?$top=2000&api-version=6.0',
        mockToken,
        'get'
      );
      expect(result).toEqual([{ id: '1', url: 'https://example.test/wi/1' }]);
    });

    it('should encode project names in the build work items URL', async () => {
      (TFSServices.getItemContent as jest.Mock).mockResolvedValue({});

      const result = await pipelinesDataProvider.GetBuildWorkItems('Project With Spaces', 123);

      expect(TFSServices.getItemContent).toHaveBeenCalledWith(
        'https://dev.azure.com/orgname/Project%20With%20Spaces/_apis/build/builds/123/workitems?$top=2000&api-version=6.0',
        mockToken,
        'get'
      );
      expect(result).toEqual([]);
    });
  });

  describe('isMatchingPipeline', () => {
    // Create test method to access private method
    const invokeIsMatchingPipeline = (
      fromPipeline: PipelineRun,
      targetPipeline: PipelineRun
    ): boolean => {
      return (pipelinesDataProvider as any).isMatchingPipeline(fromPipeline, targetPipeline);
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
                refName: 'refs/heads/main',
              },
            },
          },
        },
      } as unknown as PipelineRun;

      const targetPipeline = {
        resources: {
          repositories: {
            '0': {
              self: {
                repository: { id: 'repo2' },
                version: 'v1',
                refName: 'refs/heads/main',
              },
            },
          },
        },
      } as unknown as PipelineRun;

      // Act
      const result = invokeIsMatchingPipeline(fromPipeline, targetPipeline);

      // Assert
      expect(result).toBe(false);
    });

    it('should return true when versions are the same and refNames match', () => {
      // Arrange
      const fromPipeline = {
        resources: {
          repositories: {
            '0': {
              self: {
                repository: { id: 'repo1' },
                version: 'v1',
                refName: 'refs/heads/main',
              },
            },
          },
        },
      } as unknown as PipelineRun;

      const targetPipeline = {
        resources: {
          repositories: {
            '0': {
              self: {
                repository: { id: 'repo1' },
                version: 'v1',
                refName: 'refs/heads/main',
              },
            },
          },
        },
      } as unknown as PipelineRun;

      // Act
      const result = invokeIsMatchingPipeline(fromPipeline, targetPipeline);

      // Assert
      expect(result).toBe(true);
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
                refName: 'refs/heads/main',
              },
            },
          },
        },
      } as unknown as PipelineRun;

      const targetPipeline = {
        resources: {
          repositories: {
            '0': {
              self: {
                repository: { id: 'repo1' },
                version: 'v2',
                refName: 'refs/heads/main',
              },
            },
          },
        },
      } as unknown as PipelineRun;

      // Act
      const result = invokeIsMatchingPipeline(fromPipeline, targetPipeline);

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
              refName: 'refs/heads/main',
            },
          },
        },
      } as unknown as PipelineRun;

      const targetPipeline = {
        resources: {
          repositories: {
            __designer_repo: {
              repository: { id: 'repo1' },
              version: 'v1',
              refName: 'refs/heads/main',
            },
          },
        },
      } as unknown as PipelineRun;

      // Act
      const result = invokeIsMatchingPipeline(fromPipeline, targetPipeline);

      // Assert
      expect(result).toBe(true);
    });

    it('should resolve self from real ADO Runs API format (repositories.self, not repositories[0].self)', () => {
      // Real ADO Pipelines Runs API returns repositories as a named-key object:
      // { self: {...}, DevOpsTemplates: {...} } — NOT an array.
      const fromPipeline = {
        resources: {
          repositories: {
            self: {
              repository: { id: 'repo1' },
              version: 'v1',
              refName: 'refs/heads/main',
            },
          },
        },
      } as unknown as PipelineRun;

      const targetPipeline = {
        resources: {
          repositories: {
            self: {
              repository: { id: 'repo1' },
              version: 'v2',
              refName: 'refs/heads/main',
            },
          },
        },
      } as unknown as PipelineRun;

      const result = invokeIsMatchingPipeline(fromPipeline, targetPipeline);
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
          { id: 4, result: 'succeeded' },
        ],
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
        count: 2,
        value: mockResponse.value.filter((r) => r.result !== 'failed' && r.result !== 'canceled'),
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
      expect(logger.error).toHaveBeenCalledWith(
        `Could not fetch Pipeline Run History: ${expectedError.message}`
      );
      expect(result).toBeUndefined();
    });

    it('should return response when value is undefined', async () => {
      // Arrange
      const projectName = 'project1';
      const pipelineId = '123';
      const mockResponse = { count: 0 };
      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockResponse);

      // Act
      const result = await pipelinesDataProvider.GetPipelineRunHistory(projectName, pipelineId);

      // Assert
      expect(result).toEqual(mockResponse);
    });
  });

  describe('getPipelineBuildByBuildId', () => {
    it('should fetch pipeline build by build ID', async () => {
      // Arrange
      const projectName = 'project1';
      const buildId = 123;
      const mockResponse = { id: buildId, buildNumber: '20231201.1' };
      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockResponse);

      // Act
      const result = await pipelinesDataProvider.getPipelineBuildByBuildId(projectName, buildId);

      // Assert
      expect(TFSServices.getItemContent).toHaveBeenCalledWith(
        `${mockOrgUrl}${projectName}/_apis/build/builds/${buildId}`,
        mockToken,
        'get'
      );
      expect(result).toEqual(mockResponse);
    });
  });

  describe('TriggerBuildById', () => {
    it('should trigger a build with parameters', async () => {
      // Arrange
      const projectName = 'project1';
      const buildDefId = '456';
      const parameters = '{"Test":"123"}';
      const mockResponse = { id: 789, status: 'queued' };
      (TFSServices.postRequest as jest.Mock).mockResolvedValueOnce(mockResponse);

      // Act
      const result = await pipelinesDataProvider.TriggerBuildById(projectName, buildDefId, parameters);

      // Assert
      expect(TFSServices.postRequest).toHaveBeenCalledWith(
        `${mockOrgUrl}${projectName}/_apis/build/builds?api-version=5.0`,
        mockToken,
        'post',
        {
          definition: { id: buildDefId },
          parameters: parameters,
        },
        null
      );
      expect(result).toEqual(mockResponse);
    });
  });

  describe('GetArtifactByBuildId', () => {
    it('should return empty response when no artifacts exist', async () => {
      // Arrange
      const projectName = 'project1';
      const buildId = '123';
      const artifactName = 'drop';
      const mockResponse = { count: 0 };
      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockResponse);

      // Act
      const result = await pipelinesDataProvider.GetArtifactByBuildId(projectName, buildId, artifactName);

      // Assert
      expect(result).toEqual(mockResponse);
    });

    it('should download artifact when it exists', async () => {
      // Arrange
      const projectName = 'project1';
      const buildId = '123';
      const artifactName = 'drop';
      const mockArtifactsResponse = { count: 1 };
      const mockArtifactResponse = {
        resource: { downloadUrl: 'https://example.com/download' },
      };
      const mockDownloadResult = { data: Buffer.from('zip content') };

      (TFSServices.getItemContent as jest.Mock)
        .mockResolvedValueOnce(mockArtifactsResponse)
        .mockResolvedValueOnce(mockArtifactResponse);
      (TFSServices.downloadZipFile as jest.Mock).mockResolvedValueOnce(mockDownloadResult);

      // Act
      const result = await pipelinesDataProvider.GetArtifactByBuildId(projectName, buildId, artifactName);

      // Assert
      expect(TFSServices.downloadZipFile).toHaveBeenCalledWith('https://example.com/download', mockToken);
      expect(result).toEqual(mockDownloadResult);
    });

    it('should throw error when artifact fetch fails', async () => {
      // Arrange
      const projectName = 'project1';
      const buildId = '123';
      const artifactName = 'drop';
      const mockError = new Error('Artifact not found');
      (TFSServices.getItemContent as jest.Mock).mockRejectedValueOnce(mockError);

      // Act & Assert
      await expect(
        pipelinesDataProvider.GetArtifactByBuildId(projectName, buildId, artifactName)
      ).rejects.toThrow();
      expect(logger.error).toHaveBeenCalled();
    });
  });

  describe('getPipelineStageName', () => {
    it('should return the matching Stage record', async () => {
      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce({
        records: [
          { type: 'Stage', name: 'Deploy', state: 'completed', result: 'succeeded' },
          { type: 'Job', name: 'Job1' },
        ],
      });

      const record = await (pipelinesDataProvider as any).getPipelineStageName(123, 'project1', 'Deploy');
      expect(record).toEqual(expect.objectContaining({ type: 'Stage', name: 'Deploy' }));
    });

    it('should return undefined when no matching stage exists', async () => {
      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce({
        records: [{ type: 'Stage', name: 'Build', state: 'completed', result: 'succeeded' }],
      });

      const record = await (pipelinesDataProvider as any).getPipelineStageName(123, 'project1', 'Deploy');
      expect(record).toBeUndefined();
    });

    it('should return undefined on fetch error and log', async () => {
      (TFSServices.getItemContent as jest.Mock).mockRejectedValueOnce(new Error('boom'));

      const record = await (pipelinesDataProvider as any).getPipelineStageName(123, 'project1', 'Deploy');
      expect(record).toBeUndefined();
      expect(logger.error).toHaveBeenCalled();
    });
  });

  describe('isStageSuccessful', () => {
    it('should return false when stage is missing', async () => {
      jest.spyOn(pipelinesDataProvider as any, 'getPipelineStageName').mockResolvedValueOnce(undefined);
      await expect(
        (pipelinesDataProvider as any).isStageSuccessful({ id: 1 }, 'project1', 'Deploy')
      ).resolves.toBeUndefined();
    });

    it('should return false when stage is not completed', async () => {
      jest
        .spyOn(pipelinesDataProvider as any, 'getPipelineStageName')
        .mockResolvedValueOnce({ state: 'inProgress', result: 'succeeded' });
      await expect(
        (pipelinesDataProvider as any).isStageSuccessful({ id: 1 }, 'project1', 'Deploy')
      ).resolves.toBe(false);
    });

    it('should return false when stage result is not succeeded', async () => {
      jest
        .spyOn(pipelinesDataProvider as any, 'getPipelineStageName')
        .mockResolvedValueOnce({ state: 'completed', result: 'failed' });
      await expect(
        (pipelinesDataProvider as any).isStageSuccessful({ id: 1 }, 'project1', 'Deploy')
      ).resolves.toBe(false);
    });

    it('should return true when stage is completed and succeeded', async () => {
      jest
        .spyOn(pipelinesDataProvider as any, 'getPipelineStageName')
        .mockResolvedValueOnce({ state: 'completed', result: 'succeeded' });
      await expect(
        (pipelinesDataProvider as any).isStageSuccessful({ id: 1 }, 'project1', 'Deploy')
      ).resolves.toBe(true);
    });
  });

  describe('GetReleaseByReleaseId', () => {
    it('should fetch release by ID', async () => {
      // Arrange
      const projectName = 'project1';
      const releaseId = 123;
      const mockResponse = { id: releaseId, name: 'Release-1' };
      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockResponse);

      // Act
      const result = await pipelinesDataProvider.GetReleaseByReleaseId(projectName, releaseId);

      // Assert
      expect(result).toEqual(mockResponse);
    });

    it('should replace dev.azure.com with vsrm.dev.azure.com for release URL', async () => {
      // Arrange
      const projectName = 'project1';
      const releaseId = 123;
      const mockResponse = { id: releaseId };
      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockResponse);

      // Act
      await pipelinesDataProvider.GetReleaseByReleaseId(projectName, releaseId);

      // Assert
      expect(TFSServices.getItemContent).toHaveBeenCalledWith(
        expect.stringContaining('vsrm.dev.azure.com'),
        mockToken,
        'get',
        null,
        null
      );
    });
  });

  describe('GetReleaseHistory', () => {
    it('should fetch release history for a definition', async () => {
      // Arrange
      const projectName = 'project1';
      const definitionId = '456';
      const mockResponse = { value: [{ id: 1 }, { id: 2 }] };
      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockResponse);

      // Act
      const result = await pipelinesDataProvider.GetReleaseHistory(projectName, definitionId);

      // Assert
      expect(TFSServices.getItemContent).toHaveBeenCalledWith(
        expect.stringContaining(`definitionId=${definitionId}`),
        mockToken,
        'get',
        null,
        null
      );
      expect(result).toEqual(mockResponse);
    });
  });

  describe('GetAllReleaseHistory', () => {
    it('should fetch all releases with pagination', async () => {
      // Arrange
      const projectName = 'project1';
      const definitionId = '456';
      const mockResponse1 = {
        data: { value: [{ id: 1 }, { id: 2 }] },
        headers: { 'x-ms-continuationtoken': 'token123' },
      };
      const mockResponse2 = {
        data: { value: [{ id: 3 }] },
        headers: {},
      };
      (TFSServices.getItemContentWithHeaders as jest.Mock)
        .mockResolvedValueOnce(mockResponse1)
        .mockResolvedValueOnce(mockResponse2);

      // Act
      const result = await pipelinesDataProvider.GetAllReleaseHistory(projectName, definitionId);

      // Assert
      expect(result.count).toBe(3);
      expect(result.value).toHaveLength(3);
    });

    it('should support x-ms-continuation-token header and default value when data is missing', async () => {
      const projectName = 'project1';
      const definitionId = '456';

      (TFSServices.getItemContentWithHeaders as jest.Mock)
        .mockResolvedValueOnce({
          data: undefined,
          headers: { 'x-ms-continuation-token': 'token123' },
        })
        .mockResolvedValueOnce({
          data: { value: [{ id: 1 }] },
          headers: {},
        });

      const result = await pipelinesDataProvider.GetAllReleaseHistory(projectName, definitionId);

      expect(result.count).toBe(1);
      expect(result.value).toEqual([{ id: 1 }]);
    });

    it('should continue paging for reversed release ranges until the lower release id is loaded', async () => {
      const projectName = 'project1';
      const definitionId = '456';

      (TFSServices.getItemContentWithHeaders as jest.Mock)
        .mockResolvedValueOnce({
          data: { value: [{ id: 20 }, { id: 19 }] },
          headers: { 'x-ms-continuationtoken': 'page-2' },
        })
        .mockResolvedValueOnce({
          data: { value: [{ id: 14 }, { id: 13 }] },
          headers: {},
        });

      const result = await pipelinesDataProvider.GetAllReleaseHistory(projectName, definitionId, {
        fromId: 20,
        toId: 14,
      });

      expect(result.value.map((r: any) => r.id)).toEqual([20, 19, 14, 13]);
      expect(TFSServices.getItemContentWithHeaders).toHaveBeenCalledTimes(2);
    });

    it('should throw when release history pagination fails', async () => {
      // Arrange
      const projectName = 'project1';
      const definitionId = '456';
      (TFSServices.getItemContentWithHeaders as jest.Mock).mockRejectedValueOnce(new Error('API Error'));

      // Assert
      await expect(pipelinesDataProvider.GetAllReleaseHistory(projectName, definitionId)).rejects.toThrow(
        'API Error'
      );
      expect(logger.error).toHaveBeenCalled();
    });
  });

  describe('findPreviousSuccessfulRelease', () => {
    const releaseCandidate = (id: number, status = 'succeeded') => ({
      id,
      status: 'active',
      environments: [{ status }],
      releaseDefinition: { id: 456 },
    });

    it('should find previous successful release on a later Release API page', async () => {
      const projectName = 'project1';
      const definitionId = '456';

      (TFSServices.getItemContentWithHeaders as jest.Mock)
        .mockResolvedValueOnce({
          data: { value: [] },
          headers: { 'x-ms-continuationtoken': 'next-page' },
        })
        .mockResolvedValueOnce({
          data: { value: [releaseCandidate(80)] },
          headers: {},
        });

      const result = await pipelinesDataProvider.findPreviousSuccessfulRelease(
        projectName,
        definitionId,
        100
      );

      expect(result).toBe(80);
      expect(TFSServices.getItemContentWithHeaders).toHaveBeenCalledWith(
        expect.stringContaining('continuationToken=next-page'),
        mockToken,
        'get',
        null,
        null
      );
    });

    it('should query release history with expanded environments and ignore non-successful releases', async () => {
      const projectName = 'project1';
      const definitionId = '456';

      (TFSServices.getItemContentWithHeaders as jest.Mock).mockResolvedValueOnce({
        data: {
          value: [
            releaseCandidate(99, 'failed'),
            releaseCandidate(98, 'rejected'),
            releaseCandidate(97, 'succeeded'),
          ],
        },
        headers: {},
      });

      const result = await pipelinesDataProvider.findPreviousSuccessfulRelease(
        projectName,
        definitionId,
        100
      );

      const url = (TFSServices.getItemContentWithHeaders as jest.Mock).mock.calls[0][0];
      expect(url).toContain(`definitionId=${definitionId}`);
      expect(url).toContain('queryOrder=descending');
      expect(url).toContain('$top=200');
      expect(url).toContain('$expand=environments');
      expect(result).toBe(97);
    });

    it('should retry release discovery with api-version 6.0 when 7.1 is unsupported', async () => {
      const unsupportedError: any = new Error('The requested resource does not support api-version 7.1');
      unsupportedError.response = { status: 404 };

      (TFSServices.getItemContentWithHeaders as jest.Mock)
        .mockRejectedValueOnce(unsupportedError)
        .mockResolvedValueOnce({
          data: { value: [releaseCandidate(90)] },
          headers: {},
        });

      const result = await pipelinesDataProvider.findPreviousSuccessfulRelease('project1', '456', 100);

      expect(result).toBe(90);
      expect((TFSServices.getItemContentWithHeaders as jest.Mock).mock.calls[0][0]).toContain(
        'api-version=7.1'
      );
      expect((TFSServices.getItemContentWithHeaders as jest.Mock).mock.calls[1][0]).toContain(
        'api-version=6.0'
      );
    });

    it('should retry release discovery when unsupported api-version is reported in Axios response data', async () => {
      const unsupportedError: any = new Error('Request failed with status code 404');
      unsupportedError.response = {
        status: 404,
        data: { message: 'The requested resource does not support api-version 7.1' },
      };

      (TFSServices.getItemContentWithHeaders as jest.Mock)
        .mockRejectedValueOnce(unsupportedError)
        .mockResolvedValueOnce({
          data: { value: [releaseCandidate(89)] },
          headers: {},
        });

      const result = await pipelinesDataProvider.findPreviousSuccessfulRelease('project1', '456', 100);

      expect(result).toBe(89);
      expect((TFSServices.getItemContentWithHeaders as jest.Mock).mock.calls[1][0]).toContain(
        'api-version=6.0'
      );
    });

    it('should not retry api-version 6.0 for ordinary invalid release definition errors', async () => {
      const invalidDefinitionError: any = new Error('Release definition 456 was not found');
      invalidDefinitionError.response = { status: 404 };
      (TFSServices.getItemContentWithHeaders as jest.Mock).mockRejectedValueOnce(invalidDefinitionError);

      await expect(
        pipelinesDataProvider.findPreviousSuccessfulRelease('project1', '456', 100)
      ).rejects.toThrow('Release definition 456 was not found');

      expect(TFSServices.getItemContentWithHeaders).toHaveBeenCalledTimes(1);
      expect((TFSServices.getItemContentWithHeaders as jest.Mock).mock.calls[0][0]).toContain(
        'api-version=7.1'
      );
    });

    it('should throw release discovery permission errors without retrying api-version 6.0', async () => {
      const permissionError: any = new Error('Forbidden');
      permissionError.response = { status: 403 };
      (TFSServices.getItemContentWithHeaders as jest.Mock).mockRejectedValueOnce(permissionError);

      await expect(
        pipelinesDataProvider.findPreviousSuccessfulRelease('project1', '456', 100)
      ).rejects.toThrow('Forbidden');

      expect(TFSServices.getItemContentWithHeaders).toHaveBeenCalledTimes(1);
      expect((TFSServices.getItemContentWithHeaders as jest.Mock).mock.calls[0][0]).toContain(
        'api-version=7.1'
      );
    });

    it('should throw when a later release discovery page fails', async () => {
      (TFSServices.getItemContentWithHeaders as jest.Mock)
        .mockResolvedValueOnce({
          data: { value: [] },
          headers: { 'x-ms-continuationtoken': 'next-page' },
        })
        .mockRejectedValueOnce(new Error('page failed'));

      await expect(
        pipelinesDataProvider.findPreviousSuccessfulRelease('project1', '456', 100)
      ).rejects.toThrow('page failed');
    });

    it('should throw when previous release discovery exceeds the defensive page limit', async () => {
      (TFSServices.getItemContentWithHeaders as jest.Mock).mockImplementation(() =>
        Promise.resolve({
          data: { value: [] },
          headers: { 'x-ms-continuationtoken': 'next-page' },
        })
      );

      await expect(
        pipelinesDataProvider.findPreviousSuccessfulRelease('project1', '456', 100)
      ).rejects.toThrow('Release discovery exceeded 50 pages');
    });
  });

  describe('findLatestSuccessfulRelease', () => {
    const releaseCandidate = (id: number, status = 'succeeded') => ({
      id,
      status: 'active',
      environments: [{ status }],
      releaseDefinition: { id: 456 },
    });

    it('should find latest successful release on a later Release API page', async () => {
      (TFSServices.getItemContentWithHeaders as jest.Mock)
        .mockResolvedValueOnce({
          data: { value: [releaseCandidate(100, 'failed')] },
          headers: { 'x-ms-continuationtoken': 'next-page' },
        })
        .mockResolvedValueOnce({
          data: { value: [releaseCandidate(90)] },
          headers: {},
        });

      const result = await pipelinesDataProvider.findLatestSuccessfulRelease('project1', '456');

      expect(result).toBe(90);
      expect(TFSServices.getItemContentWithHeaders).toHaveBeenCalledWith(
        expect.stringContaining('continuationToken=next-page'),
        mockToken,
        'get',
        null,
        null
      );
    });

    it('should retry latest release discovery with api-version 6.0 only for unsupported api-version errors', async () => {
      const unsupportedError: any = new Error('The requested resource does not support api-version 7.1');
      unsupportedError.response = { status: 404 };

      (TFSServices.getItemContentWithHeaders as jest.Mock)
        .mockRejectedValueOnce(unsupportedError)
        .mockResolvedValueOnce({
          data: { value: [releaseCandidate(101)] },
          headers: {},
        });

      const result = await pipelinesDataProvider.findLatestSuccessfulRelease('project1', '456');

      expect(result).toBe(101);
      expect((TFSServices.getItemContentWithHeaders as jest.Mock).mock.calls[0][0]).toContain(
        'api-version=7.1'
      );
      expect((TFSServices.getItemContentWithHeaders as jest.Mock).mock.calls[1][0]).toContain(
        'api-version=6.0'
      );
    });

    it('should retry latest release discovery when unsupported api-version is reported in Axios response data', async () => {
      const unsupportedError: any = new Error('Request failed with status code 404');
      unsupportedError.response = {
        status: 404,
        data: { message: 'The requested resource does not support api-version 7.1' },
      };

      (TFSServices.getItemContentWithHeaders as jest.Mock)
        .mockRejectedValueOnce(unsupportedError)
        .mockResolvedValueOnce({
          data: { value: [releaseCandidate(102)] },
          headers: {},
        });

      const result = await pipelinesDataProvider.findLatestSuccessfulRelease('project1', '456');

      expect(result).toBe(102);
      expect((TFSServices.getItemContentWithHeaders as jest.Mock).mock.calls[1][0]).toContain(
        'api-version=6.0'
      );
    });

    it('should not retry latest release discovery for ordinary 404 errors', async () => {
      const invalidDefinitionError: any = new Error('Release definition 456 was not found');
      invalidDefinitionError.response = { status: 404 };
      (TFSServices.getItemContentWithHeaders as jest.Mock).mockRejectedValueOnce(invalidDefinitionError);

      await expect(
        pipelinesDataProvider.findLatestSuccessfulRelease('project1', '456')
      ).rejects.toThrow('Release definition 456 was not found');

      expect(TFSServices.getItemContentWithHeaders).toHaveBeenCalledTimes(1);
    });

    it('should throw when latest release discovery exceeds the defensive page limit', async () => {
      (TFSServices.getItemContentWithHeaders as jest.Mock).mockImplementation(() =>
        Promise.resolve({
          data: { value: [] },
          headers: { 'x-ms-continuationtoken': 'next-page' },
        })
      );

      await expect(
        pipelinesDataProvider.findLatestSuccessfulRelease('project1', '456')
      ).rejects.toThrow('Release discovery exceeded 50 pages');
    });
  });

  describe('GetAllPipelines', () => {
    it('should fetch all pipelines for a project', async () => {
      // Arrange
      const projectName = 'project1';
      const mockResponse = { value: [{ id: 1 }, { id: 2 }], count: 2 };
      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockResponse);

      // Act
      const result = await pipelinesDataProvider.GetAllPipelines(projectName);

      // Assert
      expect(TFSServices.getItemContent).toHaveBeenCalledWith(
        `${mockOrgUrl}${projectName}/_apis/pipelines?$top=2000`,
        mockToken,
        'get',
        null,
        null
      );
      expect(result).toEqual(mockResponse);
    });
  });

  describe('GetAllReleaseDefenitions', () => {
    it('should fetch all release definitions', async () => {
      // Arrange
      const projectName = 'project1';
      const mockResponse = { value: [{ id: 1 }, { id: 2 }] };
      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockResponse);

      // Act
      const result = await pipelinesDataProvider.GetAllReleaseDefenitions(projectName);

      // Assert
      expect(TFSServices.getItemContent).toHaveBeenCalledWith(
        expect.stringContaining('vsrm.dev.azure.com'),
        mockToken,
        'get',
        null,
        null
      );
      expect(result).toEqual(mockResponse);
    });
  });

  describe('GetRecentReleaseArtifactInfo', () => {
    it('should return empty array when no releases exist', async () => {
      // Arrange
      const projectName = 'project1';
      const mockResponse = { value: [] };
      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockResponse);

      // Act
      const result = await pipelinesDataProvider.GetRecentReleaseArtifactInfo(projectName);

      // Assert
      expect(result).toEqual([]);
    });

    it('should return artifact info from most recent release', async () => {
      // Arrange
      const projectName = 'project1';
      const mockReleasesResponse = { value: [{ id: 123 }] };
      const mockReleaseResponse = {
        artifacts: [
          {
            definitionReference: {
              definition: { name: 'artifact1' },
              version: { name: '1.0.0' },
            },
          },
        ],
      };
      (TFSServices.getItemContent as jest.Mock)
        .mockResolvedValueOnce(mockReleasesResponse)
        .mockResolvedValueOnce(mockReleaseResponse);

      // Act
      const result = await pipelinesDataProvider.GetRecentReleaseArtifactInfo(projectName);

      // Assert
      expect(result).toEqual([{ artifactName: 'artifact1', artifactVersion: '1.0.0' }]);
    });
  });

  describe('getPipelineResourcePipelinesFromObject', () => {
    it('should return empty array when no pipeline resources exist', async () => {
      // Arrange
      const inPipeline = {
        resources: {},
      } as unknown as PipelineRun;

      // Act
      const result = await pipelinesDataProvider.getPipelineResourcePipelinesFromObject(inPipeline);

      // Assert
      expect(result).toEqual([]);
    });

    it('should extract pipeline resources from pipeline object', async () => {
      // Arrange
      const inPipeline = {
        resources: {
          pipelines: {
            myPipeline: {
              pipeline: {
                id: 123,
                url: 'https://dev.azure.com/org/project/_apis/pipelines/123?revision=1',
              },
              runId: 789,
            },
          },
        },
      } as unknown as PipelineRun;

      const mockBuildResponse = {
        id: 789,
        definition: { id: 123, type: 'build' },
        buildNumber: '20231201.1',
        project: { name: 'project' },
        repository: { type: 'TfsGit' },
      };

      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockBuildResponse); // fixed-url attempt

      // Act
      const result = await pipelinesDataProvider.getPipelineResourcePipelinesFromObject(inPipeline);

      // Assert
      expect(result).toHaveLength(1);
      expect((result as any[])[0]).toEqual({
        name: 'myPipeline',
        buildId: 789,
        definitionId: 123,
        buildNumber: '20231201.1',
        teamProject: 'project',
        provider: 'TfsGit',
      });
    });

    it('should handle errors when fetching pipeline resources', async () => {
      // Arrange
      const inPipeline = {
        resources: {
          pipelines: {
            myPipeline: {
              pipeline: {
                id: 123,
                url: 'https://dev.azure.com/org/project/_apis/pipelines/123?revision=1',
              },
              runId: 789,
            },
          },
        },
      } as unknown as PipelineRun;

      (TFSServices.getItemContent as jest.Mock).mockRejectedValueOnce(new Error('API Error'));

      // Act
      const result = await pipelinesDataProvider.getPipelineResourcePipelinesFromObject(inPipeline);

      // Assert
      expect(result).toEqual([]);
    });

    it('should skip resource when fixed-url returns a build from a different project/pipeline', async () => {
      const inPipeline = {
        resources: {
          pipelines: {
            myPipeline: {
              pipeline: {
                id: 123,
                name: 'ExpectedPipeline',
                url: 'https://dev.azure.com/org/project/_apis/pipelines/123?revision=1',
              },
            },
          },
        },
      } as unknown as PipelineRun;

      const mockBuildResponse = {
        id: 789,
        definition: { id: 999, name: 'DifferentPipeline', type: 'build' },
        buildNumber: '20231201.1',
        project: { name: 'DifferentProject' },
        repository: { type: 'TfsGit' },
      };

      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockBuildResponse);

      const result = await pipelinesDataProvider.getPipelineResourcePipelinesFromObject(inPipeline);

      expect(result).toEqual([]);
      expect(TFSServices.getItemContent).toHaveBeenCalledTimes(1);
    });
  });

  describe('getPipelineResourceRepositoriesFromObject', () => {
    it('should return empty array when no repository resources exist', async () => {
      // Arrange
      const inPipeline = {
        resources: {},
      } as unknown as PipelineRun;
      const mockGitDataProvider = {} as GitDataProvider;

      // Act
      const result = await pipelinesDataProvider.getPipelineResourceRepositoriesFromObject(
        inPipeline,
        mockGitDataProvider
      );

      // Assert
      expect(result).toEqual([]);
    });

    it('should extract repository resources from pipeline object', async () => {
      // Arrange
      const inPipeline = {
        resources: {
          repositories: {
            self: {
              repository: { id: 'repo-123', type: 'azureReposGit' },
              version: 'abc123',
            },
          },
        },
      } as unknown as PipelineRun;

      const mockRepo = {
        id: 'repo-123',
        name: 'MyRepo',
        url: 'https://dev.azure.com/org/project/_git/MyRepo',
        // no project field → falls back to repo.url
      };

      const mockGitDataProvider = {
        GetGitRepoFromRepoId: jest.fn().mockResolvedValue(mockRepo),
      } as unknown as GitDataProvider;

      // Act
      const result = await pipelinesDataProvider.getPipelineResourceRepositoriesFromObject(
        inPipeline,
        mockGitDataProvider
      );

      // Assert
      expect(result).toHaveLength(1);
      expect((result as any[])[0]).toEqual({
        repoName: 'MyRepo',
        repoSha1: 'abc123',
        url: 'https://dev.azure.com/org/project/_git/MyRepo',
      });
    });

    it('should use project-name URL when repo response includes project.name (on-prem TFS WI fix)', async () => {
      // On-prem TFS canonicalizes repo.url to a project-UUID form, but commitsbatch only
      // resolves workItems when the URL uses the project name. This test verifies the URL
      // is rewritten to project-name form when project.name is available.
      const inPipeline = {
        resources: {
          repositories: {
            self: {
              repository: { id: 'b671b0fa-1111-2222-3333-444444444444', type: 'TfsGit' },
              version: 'abc123',
            },
          },
        },
      } as unknown as PipelineRun;

      const mockRepo = {
        id: 'b671b0fa-1111-2222-3333-444444444444',
        name: 'eden1',
        url: 'http://elis-tfs:8080/tfs/ElisraCollection/5d662fe5-uuid/_apis/git/repositories/b671b0fa-1111-2222-3333-444444444444',
        project: { id: '5d662fe5-uuid', name: 'TestProject-CMMI' },
      };

      const mockGitDataProvider = {
        GetGitRepoFromRepoId: jest.fn().mockResolvedValue(mockRepo),
      } as unknown as GitDataProvider;

      // Act
      const result = await pipelinesDataProvider.getPipelineResourceRepositoriesFromObject(
        inPipeline,
        mockGitDataProvider
      );

      // Assert — URL must use project NAME, not UUID
      expect(result).toHaveLength(1);
      expect((result as any[])[0].url).toContain('TestProject-CMMI');
      expect((result as any[])[0].url).not.toContain('5d662fe5-uuid');
      expect((result as any[])[0].url).toContain('b671b0fa-1111-2222-3333-444444444444');
    });

    it('should encode project name when rebuilding repository API URL', async () => {
      const inPipeline = {
        resources: {
          repositories: {
            self: {
              repository: { id: 'repo-123', type: 'TfsGit' },
              version: 'abc123',
            },
          },
        },
      } as unknown as PipelineRun;

      const mockRepo = {
        id: 'repo-123',
        name: 'MyRepo',
        url: 'http://elis-tfs:8080/tfs/ElisraCollection/project-id/_apis/git/repositories/repo-123',
        project: { id: 'project-id', name: 'Project With Spaces' },
      };

      const mockGitDataProvider = {
        GetGitRepoFromRepoId: jest.fn().mockResolvedValue(mockRepo),
      } as unknown as GitDataProvider;

      const result = await pipelinesDataProvider.getPipelineResourceRepositoriesFromObject(
        inPipeline,
        mockGitDataProvider
      );

      expect((result as any[])[0].url).toBe(
        'https://dev.azure.com/orgname/Project%20With%20Spaces/_apis/git/repositories/repo-123'
      );
    });

    it('should skip non-azureReposGit repositories', async () => {
      // Arrange
      const inPipeline = {
        resources: {
          repositories: {
            external: {
              repository: { id: 'repo-123', type: 'GitHub' },
              version: 'abc123',
            },
          },
        },
      } as unknown as PipelineRun;

      const mockGitDataProvider = {
        GetGitRepoFromRepoId: jest.fn(),
      } as unknown as GitDataProvider;

      // Act
      const result = await pipelinesDataProvider.getPipelineResourceRepositoriesFromObject(
        inPipeline,
        mockGitDataProvider
      );

      // Assert
      expect(result).toHaveLength(0);
      expect(mockGitDataProvider.GetGitRepoFromRepoId).not.toHaveBeenCalled();
    });
  });

  describe('findPreviousPipeline', () => {
    const targetPipelineRun = {
      resources: {
        repositories: {
          '0': {
            self: {
              repository: { id: 'repo1' },
              version: 'target-sha',
              refName: 'refs/heads/main',
            },
          },
        },
      },
    } as unknown as PipelineRun;

    const buildCandidate = (id: number, branchName: string, repoId = 'repo1') => ({
      id,
      status: 'completed',
      result: 'succeeded',
      sourceBranch: branchName,
      sourceVersion: `sha-${id}`,
      repository: { id: repoId },
      definition: { id: 123 },
    });

    it('should return undefined when no pipeline runs exist', async () => {
      // Arrange
      const teamProject = 'project1';
      const pipelineId = '123';
      const toPipelineRunId = 100;
      const targetPipeline = {} as PipelineRun;
      (TFSServices.getItemContentWithHeaders as jest.Mock).mockResolvedValueOnce({
        data: { value: [] },
        headers: {},
      });
      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce({});

      // Act
      const result = await pipelinesDataProvider.findPreviousPipeline(
        teamProject,
        pipelineId,
        toPipelineRunId,
        targetPipeline
      );

      // Assert
      expect(result).toBeUndefined();
    });

    it('should skip invalid runs and return first matching previous pipeline', async () => {
      const teamProject = 'project1';
      const pipelineId = '123';
      const toPipelineRunId = 100;
      const targetPipeline = {
        resources: {
          repositories: { '0': { self: { repository: { id: 'r' }, version: 'v2', refName: 'main' } } },
        },
      } as any;

      (TFSServices.getItemContentWithHeaders as jest.Mock)
        .mockResolvedValueOnce({
          data: { value: [] },
          headers: {},
        })
        .mockResolvedValueOnce({
          data: { value: [] },
          headers: {},
        });
      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce({
        value: [
          { id: 100, result: 'succeeded' },
          { id: 99, result: 'succeeded' },
        ],
      });

      jest.spyOn(pipelinesDataProvider as any, 'getPipelineRunDetails').mockResolvedValueOnce({
        resources: {
          repositories: { '0': { self: { repository: { id: 'r' }, version: 'v1', refName: 'main' } } },
        },
      });
      jest.spyOn(pipelinesDataProvider as any, 'isMatchingPipeline').mockReturnValueOnce(true);

      const res = await pipelinesDataProvider.findPreviousPipeline(
        teamProject,
        pipelineId,
        toPipelineRunId,
        targetPipeline
      );
      expect(res).toBe(99);
    });

    it('should skip when fromStage provided but stage is not successful', async () => {
      const teamProject = 'project1';
      const pipelineId = '123';
      const toPipelineRunId = 100;
      const targetPipeline = {} as PipelineRun;

      (TFSServices.getItemContentWithHeaders as jest.Mock).mockResolvedValueOnce({
        data: { value: [] },
        headers: {},
      });
      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce({
        value: [{ id: 99, result: 'succeeded' }],
      });
      jest.spyOn(pipelinesDataProvider as any, 'isStageSuccessful').mockResolvedValueOnce(false);
      const detailsSpy = jest.spyOn(pipelinesDataProvider as any, 'getPipelineRunDetails');

      const res = await pipelinesDataProvider.findPreviousPipeline(
        teamProject,
        pipelineId,
        toPipelineRunId,
        targetPipeline,
        'Deploy'
      );

      expect(res).toBeUndefined();
      expect(detailsSpy).not.toHaveBeenCalled();
      expect(TFSServices.getItemContentWithHeaders).not.toHaveBeenCalled();
    });

    it('should skip when pipeline details do not include repositories', async () => {
      const teamProject = 'project1';
      const pipelineId = '123';
      const toPipelineRunId = 100;
      const targetPipeline = {} as PipelineRun;

      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce({
        value: [{ id: 99, result: 'succeeded' }],
      });
      jest
        .spyOn(pipelinesDataProvider as any, 'getPipelineRunDetails')
        .mockResolvedValueOnce({ resources: {} });

      const res = await pipelinesDataProvider.findPreviousPipeline(
        teamProject,
        pipelineId,
        toPipelineRunId,
        targetPipeline
      );

      expect(res).toBeUndefined();
    });

    it('should find matching previous pipeline', async () => {
      // Arrange
      const teamProject = 'project1';
      const pipelineId = '123';
      const toPipelineRunId = 100;
      const targetPipeline = {
        resources: {
          repositories: {
            '0': {
              self: {
                repository: { id: 'repo1' },
                version: 'v2',
                refName: 'refs/heads/main',
              },
            },
          },
        },
      } as unknown as PipelineRun;

      const mockRunHistory = {
        value: [{ id: 99, result: 'succeeded' }],
      };

      const mockPipelineDetails = {
        resources: {
          repositories: {
            '0': {
              self: {
                repository: { id: 'repo1' },
                version: 'v1',
                refName: 'refs/heads/main',
              },
            },
          },
        },
      };

      (TFSServices.getItemContentWithHeaders as jest.Mock)
        .mockResolvedValueOnce({
          data: { value: [] },
          headers: {},
        })
        .mockResolvedValueOnce({
          data: { value: [] },
          headers: {},
        });
      (TFSServices.getItemContent as jest.Mock)
        .mockResolvedValueOnce(mockRunHistory)
        .mockResolvedValueOnce(mockPipelineDetails);

      // Act
      const result = await pipelinesDataProvider.findPreviousPipeline(
        teamProject,
        pipelineId,
        toPipelineRunId,
        targetPipeline
      );

      // Assert
      expect(result).toBe(99);
    });

    it('should find previous successful build on a later Builds API page', async () => {
      (TFSServices.getItemContentWithHeaders as jest.Mock)
        .mockResolvedValueOnce({
          data: { value: [] },
          headers: { 'x-ms-continuationtoken': 'next-page' },
        })
        .mockResolvedValueOnce({
          data: { value: [buildCandidate(80, 'refs/heads/main')] },
          headers: {},
        });

      const result = await pipelinesDataProvider.findPreviousPipeline(
        'project1',
        '123',
        100,
        targetPipelineRun
      );

      expect(result).toBe(80);
      expect(TFSServices.getItemContentWithHeaders).toHaveBeenCalledWith(
        expect.stringContaining('continuationToken=next-page'),
        mockToken,
        'get',
        null,
        null
      );
    });

    it('should prefer a same-branch successful build before trying cross-branch fallback', async () => {
      (TFSServices.getItemContentWithHeaders as jest.Mock).mockResolvedValueOnce({
        data: { value: [buildCandidate(90, 'refs/heads/main')] },
        headers: {},
      });

      const result = await pipelinesDataProvider.findPreviousPipeline(
        'project1',
        '123',
        100,
        targetPipelineRun
      );

      expect(result).toBe(90);
      expect(TFSServices.getItemContentWithHeaders).toHaveBeenCalledTimes(1);
      expect((TFSServices.getItemContentWithHeaders as jest.Mock).mock.calls[0][0]).toContain(
        'branchName=refs%2Fheads%2Fmain'
      );
    });

    it('should fall back to a different branch only when same-branch discovery fails', async () => {
      (TFSServices.getItemContentWithHeaders as jest.Mock)
        .mockResolvedValueOnce({
          data: { value: [] },
          headers: {},
        })
        .mockResolvedValueOnce({
          data: { value: [buildCandidate(95, 'refs/heads/release')] },
          headers: {},
        });

      const result = await pipelinesDataProvider.findPreviousPipeline(
        'project1',
        '123',
        100,
        targetPipelineRun
      );

      expect(result).toBe(95);
      expect(TFSServices.getItemContentWithHeaders).toHaveBeenCalledTimes(2);
      expect((TFSServices.getItemContentWithHeaders as jest.Mock).mock.calls[0][0]).toContain(
        'branchName=refs%2Fheads%2Fmain'
      );
      expect((TFSServices.getItemContentWithHeaders as jest.Mock).mock.calls[1][0]).not.toContain(
        'branchName='
      );
    });

    it('should query completed successful builds and ignore non-previous candidates', async () => {
      (TFSServices.getItemContentWithHeaders as jest.Mock)
        .mockResolvedValueOnce({
          data: { value: [buildCandidate(100, 'refs/heads/main'), buildCandidate(101, 'refs/heads/main')] },
          headers: {},
        })
        .mockResolvedValueOnce({
          data: { value: [] },
          headers: {},
        });
      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce({});

      const result = await pipelinesDataProvider.findPreviousPipeline(
        'project1',
        '123',
        100,
        targetPipelineRun
      );

      const url = (TFSServices.getItemContentWithHeaders as jest.Mock).mock.calls[0][0];
      expect(url).toContain('definitions=123');
      expect(url).toContain('resultFilter=succeeded');
      expect(url).toContain('statusFilter=completed');
      expect(url).toContain('queryOrder=finishTimeDescending');
      expect(result).toBeUndefined();
    });

    it('should throw and not try cross-branch fallback when same-branch Builds API fails', async () => {
      (TFSServices.getItemContentWithHeaders as jest.Mock).mockRejectedValueOnce(new Error('same branch failed'));

      await expect(
        pipelinesDataProvider.findPreviousPipeline('project1', '123', 100, targetPipelineRun)
      ).rejects.toThrow('same branch failed');

      expect(TFSServices.getItemContentWithHeaders).toHaveBeenCalledTimes(1);
      expect((TFSServices.getItemContentWithHeaders as jest.Mock).mock.calls[0][0]).toContain(
        'branchName=refs%2Fheads%2Fmain'
      );
      expect(TFSServices.getItemContent).not.toHaveBeenCalled();
    });

    it('should throw when cross-branch Builds API fallback fails after same-branch no-match', async () => {
      (TFSServices.getItemContentWithHeaders as jest.Mock)
        .mockResolvedValueOnce({
          data: { value: [] },
          headers: {},
        })
        .mockRejectedValueOnce(new Error('fallback failed'));

      await expect(
        pipelinesDataProvider.findPreviousPipeline('project1', '123', 100, targetPipelineRun)
      ).rejects.toThrow('fallback failed');

      expect(TFSServices.getItemContentWithHeaders).toHaveBeenCalledTimes(2);
      expect((TFSServices.getItemContentWithHeaders as jest.Mock).mock.calls[1][0]).not.toContain(
        'branchName='
      );
    });

    it('should throw when a later Builds API page fails', async () => {
      (TFSServices.getItemContentWithHeaders as jest.Mock)
        .mockResolvedValueOnce({
          data: { value: [] },
          headers: { 'x-ms-continuationtoken': 'next-page' },
        })
        .mockRejectedValueOnce(new Error('build page failed'));

      await expect(
        pipelinesDataProvider.findPreviousPipeline('project1', '123', 100, targetPipelineRun)
      ).rejects.toThrow('build page failed');
    });

    it('should throw when previous build discovery exceeds the defensive page limit', async () => {
      (TFSServices.getItemContentWithHeaders as jest.Mock).mockImplementation(() =>
        Promise.resolve({
          data: { value: [] },
          headers: { 'x-ms-continuationtoken': 'next-page' },
        })
      );

      await expect(
        pipelinesDataProvider.findPreviousPipeline('project1', '123', 100, targetPipelineRun)
      ).rejects.toThrow('Pipeline discovery exceeded 50 pages');
    });
  });

  describe('private helper methods', () => {
    it('tryGetTeamProjectFromAzureDevOpsUrl should extract project segment before _apis', () => {
      const fn = (pipelinesDataProvider as any).tryGetTeamProjectFromAzureDevOpsUrl.bind(
        pipelinesDataProvider
      );
      expect(fn('https://dev.azure.com/org/Test/_apis/pipelines/123?revision=1')).toBe('Test');
      expect(fn('https://dev.azure.com/_apis/pipelines/123?revision=1')).toBeUndefined();
      expect(fn('https://dev.azure.com/org/_apis/pipelines/123?revision=1')).toBe('org');
      expect(fn('not-a-url')).toBeUndefined();
    });

    it('tryBuildBuildApiUrlFromPipelinesApiUrl should rewrite pipelines url to builds url and drop query', () => {
      const fn = (pipelinesDataProvider as any).tryBuildBuildApiUrlFromPipelinesApiUrl.bind(
        pipelinesDataProvider
      );
      expect(fn('https://dev.azure.com/org/Test/_apis/pipelines/123?revision=1')).toBe(
        'https://dev.azure.com/org/Test/_apis/build/builds/123'
      );
      expect(fn('https://dev.azure.com/org/Test/_apis/build/builds/123')).toBeUndefined();
      expect(fn('not-a-url')).toBeUndefined();
    });

    it('tryParseRunIdFromUrl should parse run id from pipelines runs URL and build id from builds URL', () => {
      const fn = (pipelinesDataProvider as any).tryParseRunIdFromUrl.bind(pipelinesDataProvider);
      expect(fn('https://dev.azure.com/org/Test/_apis/pipelines/10/runs/555')).toBe(555);
      expect(fn('https://dev.azure.com/org/Test/_apis/build/builds/777')).toBe(777);
      expect(fn('https://dev.azure.com/org/Test/_apis/pipelines/10/runs/not-a-number')).toBeUndefined();
      expect(fn('not-a-url')).toBeUndefined();
    });

    it('normalizeBranchName should normalize branch to refs/heads form', () => {
      const fn = (pipelinesDataProvider as any).normalizeBranchName.bind(pipelinesDataProvider);
      expect(fn('main')).toBe('refs/heads/main');
      expect(fn('heads/main')).toBe('refs/heads/main');
      expect(fn('refs/heads/main')).toBe('refs/heads/main');
      expect(fn('')).toBeUndefined();
      expect(fn(undefined)).toBeUndefined();
    });

    it('normalizeProjectName should resolve GUID project id to name and cache it', async () => {
      const guid = '009c6fae-b000-47fe-994e-be3354b78fbc';
      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce({ name: 'ResolvedProjectName' });

      const fn = (pipelinesDataProvider as any).normalizeProjectName.bind(pipelinesDataProvider);
      const first = await fn(guid);
      const second = await fn(guid);

      expect(first).toBe('ResolvedProjectName');
      expect(second).toBe('ResolvedProjectName');
      expect(TFSServices.getItemContent).toHaveBeenCalledTimes(1);
      expect(TFSServices.getItemContent).toHaveBeenCalledWith(
        `${mockOrgUrl}_apis/projects/${encodeURIComponent(guid)}?api-version=6.0`,
        mockToken,
        'get',
        null,
        null,
        false
      );
    });

    it('normalizeProjectName should return raw GUID when project resolution fails', async () => {
      const guid = '009c6fae-b000-47fe-994e-be3354b78fbc';
      (TFSServices.getItemContent as jest.Mock).mockRejectedValueOnce(new Error('boom'));
      const fn = (pipelinesDataProvider as any).normalizeProjectName.bind(pipelinesDataProvider);
      const result = await fn(guid);
      expect(result).toBe(guid);
    });

    it('findBuildByBuildNumber should prefer matching definition name when provided', async () => {
      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce({
        value: [
          { id: 1, definition: { name: 'Other' } },
          { id: 2, definition: { name: 'Expected' } },
        ],
      });

      const fn = (pipelinesDataProvider as any).findBuildByBuildNumber.bind(pipelinesDataProvider);
      const result = await fn('project1', '20251225.2', 'main', 'Expected');
      expect(result?.id).toBe(2);
    });

    it('tryGetBuildByIdWithFallback should fall back to non-project-scoped URL when project-scoped lookup fails', async () => {
      jest
        .spyOn(pipelinesDataProvider as any, 'getPipelineBuildByBuildId')
        .mockRejectedValueOnce(new Error('not found'));
      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce({ id: 123 });

      const fn = (pipelinesDataProvider as any).tryGetBuildByIdWithFallback.bind(pipelinesDataProvider);
      const result = await fn('project1', 123);

      expect(result).toEqual({ id: 123 });
      expect(TFSServices.getItemContent).toHaveBeenCalledWith(
        `${mockOrgUrl}_apis/build/builds/123`,
        mockToken,
        'get',
        null,
        null
      );
    });

    it('getPipelineResourcePipelinesFromObject should skip resource when pipelines url cannot be converted to builds url', async () => {
      const inPipeline = {
        resources: {
          pipelines: {
            myPipeline: {
              pipeline: {
                id: 123,
                url: 'https://dev.azure.com/org/project/_apis/build/builds/123',
              },
            },
          },
        },
      } as unknown as PipelineRun;

      const result = await pipelinesDataProvider.getPipelineResourcePipelinesFromObject(inPipeline);
      expect(result).toEqual([]);
      expect(TFSServices.getItemContent).not.toHaveBeenCalled();
    });
  });

  describe('isInvalidPipelineRun', () => {
    const invokeIsInvalidPipelineRun = (
      pipelineRun: any,
      toPipelineRunId: number,
      fromStage: string
    ): boolean => {
      return (pipelinesDataProvider as any).isInvalidPipelineRun(pipelineRun, toPipelineRunId, fromStage);
    };

    it('should return true when pipeline run id >= toPipelineRunId', () => {
      expect(invokeIsInvalidPipelineRun({ id: 100, result: 'succeeded' }, 100, '')).toBe(true);
      expect(invokeIsInvalidPipelineRun({ id: 101, result: 'succeeded' }, 100, '')).toBe(true);
    });

    it('should return true for canceled/failed/canceling results', () => {
      expect(invokeIsInvalidPipelineRun({ id: 99, result: 'canceled' }, 100, '')).toBe(true);
      expect(invokeIsInvalidPipelineRun({ id: 99, result: 'failed' }, 100, '')).toBe(true);
      expect(invokeIsInvalidPipelineRun({ id: 99, result: 'canceling' }, 100, '')).toBe(true);
    });

    it('should return true for unknown result without fromStage', () => {
      expect(invokeIsInvalidPipelineRun({ id: 99, result: 'unknown' }, 100, '')).toBe(true);
    });

    it('should return false for valid pipeline run with fromStage', () => {
      expect(invokeIsInvalidPipelineRun({ id: 99, result: 'unknown' }, 100, 'Deploy')).toBe(false);
    });

    it('should return false for succeeded result', () => {
      expect(invokeIsInvalidPipelineRun({ id: 99, result: 'succeeded' }, 100, '')).toBe(false);
    });
  });
});
