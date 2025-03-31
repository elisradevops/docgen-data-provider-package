import axios from 'axios';
import { TFSServices } from '../../helpers/tfs';
import GitDataProvider from '../GitDataProvider';
import logger from '../../utils/logger';

jest.mock('../../helpers/tfs');
jest.mock('../../utils/logger');

describe('GitDataProvider - GetCommitForPipeline', () => {
  let gitDataProvider: GitDataProvider;
  const mockOrgUrl = 'https://dev.azure.com/orgname/';
  const mockToken = 'mock-token';
  const mockProjectId = 'project-123';
  const mockBuildId = 456;

  beforeEach(() => {
    jest.clearAllMocks();
    gitDataProvider = new GitDataProvider(mockOrgUrl, mockToken);
  });

  it('should return the sourceVersion from build information', async () => {
    // Arrange
    const mockCommitSha = 'abc123def456';
    const mockResponse = {
      id: mockBuildId,
      sourceVersion: mockCommitSha,
      status: 'completed'
    };

    (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockResponse);

    // Act
    const result = await gitDataProvider.GetCommitForPipeline(mockProjectId, mockBuildId);

    // Assert
    expect(TFSServices.getItemContent).toHaveBeenCalledWith(
      `${mockOrgUrl}${mockProjectId}/_apis/build/builds/${mockBuildId}`,
      mockToken,
      'get'
    );
    expect(result).toBe(mockCommitSha);
  });

  it('should throw an error if the API call fails', async () => {
    // Arrange
    const expectedError = new Error('API call failed');
    (TFSServices.getItemContent as jest.Mock).mockRejectedValueOnce(expectedError);

    // Act & Assert
    await expect(gitDataProvider.GetCommitForPipeline(mockProjectId, mockBuildId))
      .rejects.toThrow('API call failed');

    expect(TFSServices.getItemContent).toHaveBeenCalledWith(
      `${mockOrgUrl}${mockProjectId}/_apis/build/builds/${mockBuildId}`,
      mockToken,
      'get'
    );
  });

  it('should return undefined if the response does not contain sourceVersion', async () => {
    // Arrange
    const mockResponse = {
      id: mockBuildId,
      status: 'completed'
      // No sourceVersion property
    };

    (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockResponse);

    // Act
    const result = await gitDataProvider.GetCommitForPipeline(mockProjectId, mockBuildId);

    // Assert
    expect(TFSServices.getItemContent).toHaveBeenCalledWith(
      `${mockOrgUrl}${mockProjectId}/_apis/build/builds/${mockBuildId}`,
      mockToken,
      'get'
    );
    expect(result).toBeUndefined();
  });

  it('should correctly construct URL with given project ID and build ID', async () => {
    // Arrange
    const customProjectId = 'custom-project';
    const customBuildId = 789;
    const mockCommitSha = 'xyz789abc';
    const mockResponse = { sourceVersion: mockCommitSha };

    (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockResponse);

    // Act
    await gitDataProvider.GetCommitForPipeline(customProjectId, customBuildId);

    // Assert
    expect(TFSServices.getItemContent).toHaveBeenCalledWith(
      `${mockOrgUrl}${customProjectId}/_apis/build/builds/${customBuildId}`,
      mockToken,
      'get'
    );
  });

  it('should handle different organization URLs correctly', async () => {
    // Arrange
    const altOrgUrl = 'https://dev.azure.com/different-org/';
    const altGitDataProvider = new GitDataProvider(altOrgUrl, mockToken);
    const mockResponse = { sourceVersion: 'commit-sha' };

    (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockResponse);

    // Act
    await altGitDataProvider.GetCommitForPipeline(mockProjectId, mockBuildId);

    // Assert
    expect(TFSServices.getItemContent).toHaveBeenCalledWith(
      `${altOrgUrl}${mockProjectId}/_apis/build/builds/${mockBuildId}`,
      mockToken,
      'get'
    );
  });
});
describe('GitDataProvider - GetTeamProjectGitReposList', () => {
  let gitDataProvider: GitDataProvider;
  const mockOrgUrl = 'https://dev.azure.com/orgname/';
  const mockToken = 'mock-token';
  const mockTeamProject = 'project-123';

  beforeEach(() => {
    jest.clearAllMocks();
    gitDataProvider = new GitDataProvider(mockOrgUrl, mockToken);
  });

  it('should return sorted repositories when API call succeeds', async () => {
    // Arrange
    const mockRepos = {
      value: [
        { id: 'repo2', name: 'ZRepo' },
        { id: 'repo1', name: 'ARepo' },
        { id: 'repo3', name: 'MRepo' }
      ]
    };
    (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockRepos);

    // Act
    const result = await gitDataProvider.GetTeamProjectGitReposList(mockTeamProject);

    // Assert
    expect(TFSServices.getItemContent).toHaveBeenCalledWith(
      `${mockOrgUrl}/${mockTeamProject}/_apis/git/repositories`,
      mockToken,
      'get'
    );
    expect(result).toHaveLength(3);
    expect(result[0].name).toBe('ARepo');
    expect(result[1].name).toBe('MRepo');
    expect(result[2].name).toBe('ZRepo');
    expect(logger.debug).toHaveBeenCalledWith(
      expect.stringContaining(`fetching repos list for team project - ${mockTeamProject}`)
    );
  });

  it('should return empty array when no repositories exist', async () => {
    // Arrange
    const mockEmptyRepos = { value: [] };
    (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockEmptyRepos);

    // Act
    const result = await gitDataProvider.GetTeamProjectGitReposList(mockTeamProject);

    // Assert
    expect(result).toEqual([]);
  });

  it('should handle API errors appropriately', async () => {
    // Arrange
    const mockError = new Error('API Error');
    (TFSServices.getItemContent as jest.Mock).mockRejectedValueOnce(mockError);

    // Act & Assert
    await expect(gitDataProvider.GetTeamProjectGitReposList(mockTeamProject))
      .rejects.toThrow('API Error');
  });
});

describe('GitDataProvider - GetFileFromGitRepo', () => {
  let gitDataProvider: GitDataProvider;
  const mockOrgUrl = 'https://dev.azure.com/orgname/';
  const mockToken = 'mock-token';
  const mockProjectName = 'project-123';
  const mockRepoId = 'repo-456';
  const mockFileName = 'README.md';
  const mockVersion = { version: 'main', versionType: 'branch' };

  beforeEach(() => {
    jest.clearAllMocks();
    gitDataProvider = new GitDataProvider(mockOrgUrl, mockToken);
  });

  it('should return file content when file exists', async () => {
    // Arrange
    const mockContent = 'This is a test readme file';
    const mockResponse = { content: mockContent };
    (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockResponse);

    // Act
    const result = await gitDataProvider.GetFileFromGitRepo(
      mockProjectName,
      mockRepoId,
      mockFileName,
      mockVersion
    );

    // Assert
    expect(TFSServices.getItemContent).toHaveBeenCalledWith(
      expect.stringContaining(`${mockOrgUrl}${mockProjectName}/_apis/git/repositories/${mockRepoId}/items`),
      mockToken,
      'get',
      {},
      {},
      false
    );
    expect(result).toBe(mockContent);
  });

  it('should handle special characters in version by encoding them', async () => {
    // Arrange
    const specialVersion = { version: 'feature/branch#123', versionType: 'branch' };
    const mockResponse = { content: 'content' };
    (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockResponse);

    // Act
    await gitDataProvider.GetFileFromGitRepo(
      mockProjectName,
      mockRepoId,
      mockFileName,
      specialVersion
    );

    // Assert
    expect(TFSServices.getItemContent).toHaveBeenCalledWith(
      expect.stringContaining('versionDescriptor.version=feature%2Fbranch%23123'),
      expect.anything(),
      expect.anything(),
      expect.anything(),
      expect.anything(),
      expect.anything()
    );
  });

  it('should use custom gitRepoUrl if provided', async () => {
    // Arrange
    const mockCustomUrl = 'https://custom.git.url';
    const mockResponse = { content: 'content' };
    (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockResponse);

    // Act
    await gitDataProvider.GetFileFromGitRepo(
      mockProjectName,
      mockRepoId,
      mockFileName,
      mockVersion,
      mockCustomUrl
    );

    // Assert
    expect(TFSServices.getItemContent).toHaveBeenCalledWith(
      expect.stringContaining(mockCustomUrl),
      expect.anything(),
      expect.anything(),
      expect.anything(),
      expect.anything(),
      expect.anything()
    );
  });

  it('should return undefined when file does not exist', async () => {
    // Arrange
    (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce({});

    // Act
    const result = await gitDataProvider.GetFileFromGitRepo(
      mockProjectName,
      mockRepoId,
      mockFileName,
      mockVersion
    );

    // Assert
    expect(result).toBeUndefined();
  });

  it('should log warning and return undefined when error occurs', async () => {
    // Arrange
    const mockError = new Error('File not found');
    (TFSServices.getItemContent as jest.Mock).mockRejectedValueOnce(mockError);

    // Act
    const result = await gitDataProvider.GetFileFromGitRepo(
      mockProjectName,
      mockRepoId,
      mockFileName,
      mockVersion
    );

    // Assert
    expect(logger.warn).toHaveBeenCalledWith(
      expect.stringContaining(`File ${mockFileName} could not be read: ${mockError.message}`)
    );
    expect(result).toBeUndefined();
  });
});

describe('GitDataProvider - CheckIfItemExist', () => {
  let gitDataProvider: GitDataProvider;
  const mockOrgUrl = 'https://dev.azure.com/orgname/';
  const mockToken = 'mock-token';
  const mockGitApiUrl = 'https://dev.azure.com/orgname/project/_apis/git/repositories/repo-id';
  const mockItemPath = 'path/to/file.txt';
  const mockVersion = { version: 'main', versionType: 'branch' };

  beforeEach(() => {
    jest.clearAllMocks();
    gitDataProvider = new GitDataProvider(mockOrgUrl, mockToken);
  });

  it('should return true when item exists', async () => {
    // Arrange
    const mockResponse = { path: mockItemPath, content: 'content' };
    (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockResponse);

    // Act
    const result = await gitDataProvider.CheckIfItemExist(
      mockGitApiUrl,
      mockItemPath,
      mockVersion
    );

    // Assert
    expect(TFSServices.getItemContent).toHaveBeenCalledWith(
      expect.stringContaining(`${mockGitApiUrl}/items?path=${mockItemPath}`),
      mockToken,
      'get',
      {},
      {},
      false
    );
    expect(result).toBe(true);
  });

  it('should return false when item does not exist', async () => {
    // Arrange
    (TFSServices.getItemContent as jest.Mock).mockRejectedValueOnce(new Error('Not found'));

    // Act
    const result = await gitDataProvider.CheckIfItemExist(
      mockGitApiUrl,
      mockItemPath,
      mockVersion
    );

    // Assert
    expect(result).toBe(false);
  });

  it('should return false when API returns null', async () => {
    // Arrange
    (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(null);

    // Act
    const result = await gitDataProvider.CheckIfItemExist(
      mockGitApiUrl,
      mockItemPath,
      mockVersion
    );

    // Assert
    expect(result).toBe(false);
  });

  it('should handle special characters in version', async () => {
    // Arrange
    const specialVersion = { version: 'feature/branch#123', versionType: 'branch' };
    const mockResponse = { path: mockItemPath };
    (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockResponse);

    // Act
    await gitDataProvider.CheckIfItemExist(
      mockGitApiUrl,
      mockItemPath,
      specialVersion
    );

    // Assert
    expect(TFSServices.getItemContent).toHaveBeenCalledWith(
      expect.stringContaining('versionDescriptor.version=feature%2Fbranch%23123'),
      expect.anything(),
      expect.anything(),
      expect.anything(),
      expect.anything(),
      expect.anything()
    );
  });
});

describe('GitDataProvider - GetPullRequestsInCommitRangeWithoutLinkedItems', () => {
  let gitDataProvider: GitDataProvider;
  const mockOrgUrl = 'https://dev.azure.com/orgname/';
  const mockToken = 'mock-token';
  const mockProjectId = 'project-123';
  const mockRepoId = 'repo-456';

  beforeEach(() => {
    jest.clearAllMocks();
    gitDataProvider = new GitDataProvider(mockOrgUrl, mockToken);
  });

  it('should return filtered pull requests matching commit ids', async () => {
    // Arrange
    const mockCommits = {
      value: [
        { commitId: 'commit-1' },
        { commitId: 'commit-2' }
      ]
    };

    const mockPullRequests = {
      count: 3,
      value: [
        {
          pullRequestId: 101,
          title: 'PR 1',
          createdBy: { displayName: 'User 1' },
          creationDate: '2023-01-01',
          closedDate: '2023-01-02',
          description: 'Description 1',
          lastMergeCommit: { commitId: 'commit-1' }
        },
        {
          pullRequestId: 102,
          title: 'PR 2',
          createdBy: { displayName: 'User 2' },
          creationDate: '2023-02-01',
          closedDate: '2023-02-02',
          description: 'Description 2',
          lastMergeCommit: { commitId: 'commit-3' } // Not in our commit range
        },
        {
          pullRequestId: 103,
          title: 'PR 3',
          createdBy: { displayName: 'User 3' },
          creationDate: '2023-03-01',
          closedDate: '2023-03-02',
          description: 'Description 3',
          lastMergeCommit: { commitId: 'commit-2' }
        }
      ]
    };

    (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockPullRequests);

    // Act
    const result = await gitDataProvider.GetPullRequestsInCommitRangeWithoutLinkedItems(
      mockProjectId,
      mockRepoId,
      mockCommits
    );

    // Assert
    expect(TFSServices.getItemContent).toHaveBeenCalledWith(
      expect.stringContaining(`${mockOrgUrl}${mockProjectId}/_apis/git/repositories/${mockRepoId}/pullrequests`),
      mockToken,
      'get'
    );
    expect(result).toHaveLength(2);
    expect(result[0].pullRequestId).toBe(101);
    expect(result[1].pullRequestId).toBe(103);
    expect(logger.info).toHaveBeenCalledWith(
      expect.stringContaining('filtered in commit range 2 pullrequests')
    );
  });

  it('should return empty array when no matching pull requests', async () => {
    // Arrange
    const mockCommits = {
      value: [
        { commitId: 'commit-999' } // Not matching any PRs
      ]
    };

    const mockPullRequests = {
      count: 2,
      value: [
        {
          pullRequestId: 101,
          lastMergeCommit: { commitId: 'commit-1' }
        },
        {
          pullRequestId: 102,
          lastMergeCommit: { commitId: 'commit-2' }
        }
      ]
    };

    (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockPullRequests);

    // Act
    const result = await gitDataProvider.GetPullRequestsInCommitRangeWithoutLinkedItems(
      mockProjectId,
      mockRepoId,
      mockCommits
    );

    // Assert
    expect(result).toHaveLength(0);
  });

  it('should handle API errors appropriately', async () => {
    // Arrange
    const mockCommits = { value: [{ commitId: 'commit-1' }] };
    const mockError = new Error('API Error');
    (TFSServices.getItemContent as jest.Mock).mockRejectedValueOnce(mockError);

    // Act & Assert
    await expect(gitDataProvider.GetPullRequestsInCommitRangeWithoutLinkedItems(
      mockProjectId,
      mockRepoId,
      mockCommits
    )).rejects.toThrow('API Error');
  });
});

describe('GitDataProvider - GetRepoReferences', () => {
  let gitDataProvider: GitDataProvider;
  const mockOrgUrl = 'https://dev.azure.com/orgname/';
  const mockToken = 'mock-token';
  const mockProjectId = 'project-123';
  const mockRepoId = 'repo-456';

  beforeEach(() => {
    jest.clearAllMocks();
    gitDataProvider = new GitDataProvider(mockOrgUrl, mockToken);
  });

  it('should return formatted tags when gitObjectType is "tag"', async () => {
    // Arrange
    const mockTags = {
      count: 2,
      value: [
        { name: 'refs/tags/v1.0.0', objectId: 'tag-1' },
        { name: 'refs/tags/v2.0.0', objectId: 'tag-2' }
      ]
    };
    (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockTags);

    // Act
    const result = await gitDataProvider.GetRepoReferences(
      mockProjectId,
      mockRepoId,
      'tag'
    );

    // Assert
    expect(TFSServices.getItemContent).toHaveBeenCalledWith(
      `${mockOrgUrl}${mockProjectId}/_apis/git/repositories/${mockRepoId}/refs/tags?api-version=5.1`,
      mockToken,
      'get'
    );
    expect(result).toHaveLength(2);
    expect(result[0].name).toBe('v1.0.0');
    expect(result[0].value).toBe('refs/tags/v1.0.0');
    expect(result[1].name).toBe('v2.0.0');
    expect(result[1].value).toBe('refs/tags/v2.0.0');
  });

  it('should return formatted branches when gitObjectType is "branch"', async () => {
    // Arrange
    const mockBranches = {
      count: 2,
      value: [
        { name: 'refs/heads/main', objectId: 'branch-1' },
        { name: 'refs/heads/develop', objectId: 'branch-2' }
      ]
    };
    (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockBranches);

    // Act
    const result = await gitDataProvider.GetRepoReferences(
      mockProjectId,
      mockRepoId,
      'branch'
    );

    // Assert
    expect(TFSServices.getItemContent).toHaveBeenCalledWith(
      `${mockOrgUrl}${mockProjectId}/_apis/git/repositories/${mockRepoId}/refs/heads?api-version=5.1`,
      mockToken,
      'get'
    );
    expect(result).toHaveLength(2);
    expect(result[0].name).toBe('main');
    expect(result[0].value).toBe('refs/heads/main');
    expect(result[1].name).toBe('develop');
    expect(result[1].value).toBe('refs/heads/develop');
  });

  it('should throw error for unsupported git object type', async () => {
    // Act & Assert
    await expect(gitDataProvider.GetRepoReferences(
      mockProjectId,
      mockRepoId,
      'invalid-type'
    )).rejects.toThrow('Unsupported git object type: invalid-type');
  });

  it('should return empty array when no references exist', async () => {
    // Arrange
    const mockEmptyRefs = { count: 0, value: [] };
    (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockEmptyRefs);

    // Act
    const result = await gitDataProvider.GetRepoReferences(
      mockProjectId,
      mockRepoId,
      'branch'
    );

    // Assert
    expect(result).toEqual([]);
  });
});

describe('GitDataProvider - GetJsonFileFromGitRepo', () => {
  let gitDataProvider: GitDataProvider;
  const mockOrgUrl = 'https://dev.azure.com/orgname/';
  const mockToken = 'mock-token';
  const mockProjectName = 'project-123';
  const mockRepoName = 'repo-456';
  const mockFilePath = 'config/settings.json';

  beforeEach(() => {
    jest.clearAllMocks();
    gitDataProvider = new GitDataProvider(mockOrgUrl, mockToken);
  });

  it('should parse and return JSON content when file exists', async () => {
    // Arrange
    const mockJsonContent = { setting1: 'value1', setting2: 'value2' };
    const mockResponse = { content: JSON.stringify(mockJsonContent) };
    (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockResponse);

    // Act
    const result = await gitDataProvider.GetJsonFileFromGitRepo(
      mockProjectName,
      mockRepoName,
      mockFilePath
    );

    // Assert
    expect(TFSServices.getItemContent).toHaveBeenCalledWith(
      `${mockOrgUrl}${mockProjectName}/_apis/git/repositories/${mockRepoName}/items?path=${mockFilePath}&includeContent=true`,
      mockToken,
      'get'
    );
    expect(result).toEqual(mockJsonContent);
  });

  it('should throw an error for invalid JSON content', async () => {
    // Arrange
    const mockInvalidJson = { content: '{ invalid json' };
    (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockInvalidJson);

    // Act & Assert
    await expect(gitDataProvider.GetJsonFileFromGitRepo(
      mockProjectName,
      mockRepoName,
      mockFilePath
    )).rejects.toThrow(SyntaxError);
  });

  it('should handle API errors appropriately', async () => {
    // Arrange
    const mockError = new Error('API Error');
    (TFSServices.getItemContent as jest.Mock).mockRejectedValueOnce(mockError);

    // Act & Assert
    await expect(gitDataProvider.GetJsonFileFromGitRepo(
      mockProjectName,
      mockRepoName,
      mockFilePath
    )).rejects.toThrow('API Error');
  });
});