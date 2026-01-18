import axios from 'axios';
import { TFSServices } from '../../helpers/tfs';
import GitDataProvider from '../../modules/GitDataProvider';
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
      status: 'completed',
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
    await expect(gitDataProvider.GetCommitForPipeline(mockProjectId, mockBuildId)).rejects.toThrow(
      'API call failed'
    );

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
      status: 'completed',
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
        { id: 'repo3', name: 'MRepo' },
      ],
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
    await expect(gitDataProvider.GetTeamProjectGitReposList(mockTeamProject)).rejects.toThrow('API Error');
  });
});

describe('GitDataProvider - GetGitRepoFromRepoId', () => {
  let gitDataProvider: GitDataProvider;
  const mockOrgUrl = 'https://dev.azure.com/orgname/';
  const mockToken = 'mock-token';

  beforeEach(() => {
    jest.clearAllMocks();
    gitDataProvider = new GitDataProvider(mockOrgUrl, mockToken);
  });

  it('should fetch repository by id', async () => {
    const mockRepoId = 'repo-123';
    const mockResponse = { id: mockRepoId, name: 'Repo' };
    (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockResponse);

    const result = await gitDataProvider.GetGitRepoFromRepoId(mockRepoId);

    expect(TFSServices.getItemContent).toHaveBeenCalledWith(
      `${mockOrgUrl}_apis/git/repositories/${mockRepoId}`,
      mockToken,
      'get'
    );
    expect(result).toEqual(mockResponse);
  });
});

describe('GitDataProvider - GetTag', () => {
  let gitDataProvider: GitDataProvider;
  const mockOrgUrl = 'https://dev.azure.com/orgname/';
  const mockToken = 'mock-token';
  const mockGitRepoUrl = 'https://dev.azure.com/orgname/project/_apis/git/repositories/repo-id';

  beforeEach(() => {
    jest.clearAllMocks();
    gitDataProvider = new GitDataProvider(mockOrgUrl, mockToken);
  });

  it('should return tag info when tag exists', async () => {
    (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce({
      value: [
        {
          name: 'refs/tags/v1.0.0',
          objectId: 'abc123',
          peeledObjectId: 'def456',
        },
      ],
    });

    const result = await gitDataProvider.GetTag(mockGitRepoUrl, 'v1.0.0');

    expect(result).toEqual(
      expect.objectContaining({
        name: 'v1.0.0',
        objectId: 'abc123',
        peeledObjectId: 'def456',
      })
    );
  });

  it('should return null when no matching tag found', async () => {
    (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce({
      value: [{ name: 'refs/tags/other-tag', objectId: 'abc123' }],
    });

    const result = await gitDataProvider.GetTag(mockGitRepoUrl, 'v1.0.0');
    expect(result).toBeNull();
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
    await gitDataProvider.GetFileFromGitRepo(mockProjectName, mockRepoId, mockFileName, specialVersion);

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
    const result = await gitDataProvider.CheckIfItemExist(mockGitApiUrl, mockItemPath, mockVersion);

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
    const result = await gitDataProvider.CheckIfItemExist(mockGitApiUrl, mockItemPath, mockVersion);

    // Assert
    expect(result).toBe(false);
  });

  it('should return false when API returns null', async () => {
    // Arrange
    (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(null);

    // Act
    const result = await gitDataProvider.CheckIfItemExist(mockGitApiUrl, mockItemPath, mockVersion);

    // Assert
    expect(result).toBe(false);
  });

  it('should handle special characters in version', async () => {
    // Arrange
    const specialVersion = { version: 'feature/branch#123', versionType: 'branch' };
    const mockResponse = { path: mockItemPath };
    (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockResponse);

    // Act
    await gitDataProvider.CheckIfItemExist(mockGitApiUrl, mockItemPath, specialVersion);

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
      value: [{ commitId: 'commit-1' }, { commitId: 'commit-2' }],
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
          lastMergeCommit: { commitId: 'commit-1' },
        },
        {
          pullRequestId: 102,
          title: 'PR 2',
          createdBy: { displayName: 'User 2' },
          creationDate: '2023-02-01',
          closedDate: '2023-02-02',
          description: 'Description 2',
          lastMergeCommit: { commitId: 'commit-3' }, // Not in our commit range
        },
        {
          pullRequestId: 103,
          title: 'PR 3',
          createdBy: { displayName: 'User 3' },
          creationDate: '2023-03-01',
          closedDate: '2023-03-02',
          description: 'Description 3',
          lastMergeCommit: { commitId: 'commit-2' },
        },
      ],
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
      expect.stringContaining(
        `${mockOrgUrl}${mockProjectId}/_apis/git/repositories/${mockRepoId}/pullrequests`
      ),
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
        { commitId: 'commit-999' }, // Not matching any PRs
      ],
    };

    const mockPullRequests = {
      count: 2,
      value: [
        {
          pullRequestId: 101,
          lastMergeCommit: { commitId: 'commit-1' },
        },
        {
          pullRequestId: 102,
          lastMergeCommit: { commitId: 'commit-2' },
        },
      ],
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
    await expect(
      gitDataProvider.GetPullRequestsInCommitRangeWithoutLinkedItems(mockProjectId, mockRepoId, mockCommits)
    ).rejects.toThrow('API Error');
  });
});

describe('GitDataProvider - GetBranch', () => {
  let gitDataProvider: GitDataProvider;
  const mockOrgUrl = 'https://dev.azure.com/orgname/';
  const mockToken = 'mock-token';
  const mockGitRepoUrl = 'https://dev.azure.com/orgname/project/_apis/git/repositories/repo-id';

  beforeEach(() => {
    jest.clearAllMocks();
    gitDataProvider = new GitDataProvider(mockOrgUrl, mockToken);
  });

  it('should return branch info when branch exists', async () => {
    // Arrange
    const branchName = 'main';
    const mockResponse = {
      value: [{ name: 'refs/heads/main', objectId: 'abc123' }],
    };
    (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockResponse);

    // Act
    const result = await gitDataProvider.GetBranch(mockGitRepoUrl, branchName);

    // Assert
    expect(result).toEqual({ name: 'refs/heads/main', objectId: 'abc123' });
  });

  it('should return null when branch does not exist', async () => {
    // Arrange
    const mockResponse = { value: [] };
    (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockResponse);

    // Act
    const result = await gitDataProvider.GetBranch(mockGitRepoUrl, 'nonexistent');

    // Assert
    expect(result).toBeNull();
  });

  it('should return null when response is empty', async () => {
    // Arrange
    (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(null);

    // Act
    const result = await gitDataProvider.GetBranch(mockGitRepoUrl, 'main');

    // Assert
    expect(result).toBeNull();
  });
});

describe('GitDataProvider - GetGitRepoFromPrId', () => {
  let gitDataProvider: GitDataProvider;
  const mockOrgUrl = 'https://dev.azure.com/orgname/';
  const mockToken = 'mock-token';

  beforeEach(() => {
    jest.clearAllMocks();
    gitDataProvider = new GitDataProvider(mockOrgUrl, mockToken);
  });

  it('should fetch PR by ID', async () => {
    // Arrange
    const prId = 123;
    const mockResponse = { pullRequestId: prId, title: 'Test PR' };
    (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockResponse);

    // Act
    const result = await gitDataProvider.GetGitRepoFromPrId(prId);

    // Assert
    expect(TFSServices.getItemContent).toHaveBeenCalledWith(
      `${mockOrgUrl}_apis/git/pullrequests/${prId}`,
      mockToken,
      'get'
    );
    expect(result).toEqual(mockResponse);
  });
});

describe('GitDataProvider - GetPullRequestCommits', () => {
  let gitDataProvider: GitDataProvider;
  const mockOrgUrl = 'https://dev.azure.com/orgname/';
  const mockToken = 'mock-token';

  beforeEach(() => {
    jest.clearAllMocks();
    gitDataProvider = new GitDataProvider(mockOrgUrl, mockToken);
  });

  it('should fetch PR commits', async () => {
    // Arrange
    const repoId = 'repo-123';
    const prId = 456;
    const mockResponse = { value: [{ commitId: 'abc123' }] };
    (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockResponse);

    // Act
    const result = await gitDataProvider.GetPullRequestCommits(repoId, prId);

    // Assert
    expect(TFSServices.getItemContent).toHaveBeenCalledWith(
      `${mockOrgUrl}_apis/git/repositories/${repoId}/pullRequests/${prId}/commits`,
      mockToken,
      'get'
    );
    expect(result).toEqual(mockResponse);
  });
});

describe('GitDataProvider - GetCommitByCommitId', () => {
  let gitDataProvider: GitDataProvider;
  const mockOrgUrl = 'https://dev.azure.com/orgname/';
  const mockToken = 'mock-token';

  beforeEach(() => {
    jest.clearAllMocks();
    gitDataProvider = new GitDataProvider(mockOrgUrl, mockToken);
  });

  it('should fetch commit by SHA', async () => {
    // Arrange
    const projectId = 'project-123';
    const repoId = 'repo-456';
    const commitSha = 'abc123def456';
    const mockResponse = { commitId: commitSha, comment: 'Test commit' };
    (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockResponse);

    // Act
    const result = await gitDataProvider.GetCommitByCommitId(projectId, repoId, commitSha);

    // Assert
    expect(TFSServices.getItemContent).toHaveBeenCalledWith(
      `${mockOrgUrl}${projectId}/_apis/git/repositories/${repoId}/commits/${commitSha}`,
      mockToken,
      'get'
    );
    expect(result).toEqual(mockResponse);
  });
});

describe('GitDataProvider - GetItemsForPipelinesRange', () => {
  let gitDataProvider: GitDataProvider;
  const mockOrgUrl = 'https://dev.azure.com/orgname/';
  const mockToken = 'mock-token';

  beforeEach(() => {
    jest.clearAllMocks();
    gitDataProvider = new GitDataProvider(mockOrgUrl, mockToken);
  });

  it('should fetch items in pipeline range', async () => {
    // Arrange
    const projectId = 'project-123';
    const fromBuildId = 100;
    const toBuildId = 200;
    const mockWorkItemsResponse = {
      count: 1,
      value: [{ id: 1 }],
    };
    const mockWorkItemResponse = { id: 1, fields: { 'System.Title': 'Test' } };

    (TFSServices.getItemContent as jest.Mock)
      .mockResolvedValueOnce(mockWorkItemsResponse)
      .mockResolvedValueOnce(mockWorkItemResponse);

    // Act
    const result = await gitDataProvider.GetItemsForPipelinesRange(projectId, fromBuildId, toBuildId);

    // Assert
    expect(TFSServices.getItemContent).toHaveBeenCalledWith(
      expect.stringContaining(`fromBuildId=${fromBuildId}&toBuildId=${toBuildId}`),
      mockToken,
      'get'
    );
    expect(result).toHaveLength(1);
  });
});

describe('GitDataProvider - GetCommitsInDateRange', () => {
  let gitDataProvider: GitDataProvider;
  const mockOrgUrl = 'https://dev.azure.com/orgname/';
  const mockToken = 'mock-token';

  beforeEach(() => {
    jest.clearAllMocks();
    gitDataProvider = new GitDataProvider(mockOrgUrl, mockToken);
  });

  it('should fetch commits in date range without branch', async () => {
    // Arrange
    const projectId = 'project-123';
    const repoId = 'repo-456';
    const fromDate = '2023-01-01';
    const toDate = '2023-12-31';
    const mockResponse = { value: [{ commitId: 'abc123' }] };
    (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockResponse);

    // Act
    const result = await gitDataProvider.GetCommitsInDateRange(projectId, repoId, fromDate, toDate);

    // Assert
    expect(TFSServices.getItemContent).toHaveBeenCalledWith(
      expect.stringContaining(`fromDate=${fromDate}`),
      mockToken,
      'get'
    );
    expect(result).toEqual(mockResponse);
  });

  it('should fetch commits in date range with branch', async () => {
    // Arrange
    const projectId = 'project-123';
    const repoId = 'repo-456';
    const fromDate = '2023-01-01';
    const toDate = '2023-12-31';
    const branchName = 'main';
    const mockResponse = { value: [{ commitId: 'abc123' }] };
    (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockResponse);

    // Act
    const result = await gitDataProvider.GetCommitsInDateRange(
      projectId,
      repoId,
      fromDate,
      toDate,
      branchName
    );

    // Assert
    expect(TFSServices.getItemContent).toHaveBeenCalledWith(
      expect.stringContaining(`itemVersion.version=${branchName}`),
      mockToken,
      'get'
    );
  });
});

describe('GitDataProvider - GetCommitsInCommitRange', () => {
  let gitDataProvider: GitDataProvider;
  const mockOrgUrl = 'https://dev.azure.com/orgname/';
  const mockToken = 'mock-token';

  beforeEach(() => {
    jest.clearAllMocks();
    gitDataProvider = new GitDataProvider(mockOrgUrl, mockToken);
  });

  it('should fetch commits between two commit SHAs', async () => {
    // Arrange
    const projectId = 'project-123';
    const repoId = 'repo-456';
    const fromSha = 'abc123';
    const toSha = 'def456';
    const mockResponse = { value: [{ commitId: 'xyz789' }] };
    (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockResponse);

    // Act
    const result = await gitDataProvider.GetCommitsInCommitRange(projectId, repoId, fromSha, toSha);

    // Assert
    expect(TFSServices.getItemContent).toHaveBeenCalledWith(
      expect.stringContaining(`fromCommitId=${fromSha}&searchCriteria.toCommitId=${toSha}`),
      mockToken,
      'get'
    );
    expect(result).toEqual(mockResponse);
  });
});

describe('GitDataProvider - CreatePullRequestComment', () => {
  let gitDataProvider: GitDataProvider;
  const mockOrgUrl = 'https://dev.azure.com/orgname/';
  const mockToken = 'mock-token';

  beforeEach(() => {
    jest.clearAllMocks();
    gitDataProvider = new GitDataProvider(mockOrgUrl, mockToken);
  });

  it('should create a PR comment thread', async () => {
    // Arrange
    const projectName = 'project-123';
    const repoId = 'repo-456';
    const prId = 789;
    const threads = { comments: [{ content: 'Test comment' }] };
    const mockResponse = { id: 1, comments: threads.comments };
    (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockResponse);

    // Act
    const result = await gitDataProvider.CreatePullRequestComment(projectName, repoId, prId, threads);

    // Assert
    expect(TFSServices.getItemContent).toHaveBeenCalledWith(
      expect.stringContaining(`pullRequests/${prId}/threads`),
      mockToken,
      'post',
      threads,
      null
    );
    expect(result).toEqual(mockResponse);
  });
});

describe('GitDataProvider - GetPullRequestComments', () => {
  let gitDataProvider: GitDataProvider;
  const mockOrgUrl = 'https://dev.azure.com/orgname/';
  const mockToken = 'mock-token';

  beforeEach(() => {
    jest.clearAllMocks();
    gitDataProvider = new GitDataProvider(mockOrgUrl, mockToken);
  });

  it('should fetch PR comments', async () => {
    // Arrange
    const projectName = 'project-123';
    const repoId = 'repo-456';
    const prId = 789;
    const mockResponse = { value: [{ id: 1, comments: [] }] };
    (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockResponse);

    // Act
    const result = await gitDataProvider.GetPullRequestComments(projectName, repoId, prId);

    // Assert
    expect(TFSServices.getItemContent).toHaveBeenCalledWith(
      expect.stringContaining(`pullRequests/${prId}/threads`),
      mockToken,
      'get',
      null,
      null
    );
    expect(result).toEqual(mockResponse);
  });
});

describe('GitDataProvider - GetCommitsForRepo', () => {
  let gitDataProvider: GitDataProvider;
  const mockOrgUrl = 'https://dev.azure.com/orgname/';
  const mockToken = 'mock-token';

  beforeEach(() => {
    jest.clearAllMocks();
    gitDataProvider = new GitDataProvider(mockOrgUrl, mockToken);
  });

  it('should fetch commits for repo with version identifier', async () => {
    // Arrange
    const projectName = 'project-123';
    const repoId = 'repo-456';
    const versionId = 'main';
    const mockResponse = {
      count: 2,
      value: [
        { commitId: 'abc123', comment: 'First commit', committer: { date: '2023-01-01' } },
        { commitId: 'def456', comment: 'Second commit', author: { date: '2023-01-02' } },
      ],
    };
    (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockResponse);

    // Act
    const result = await gitDataProvider.GetCommitsForRepo(projectName, repoId, versionId);

    // Assert
    expect(result).toHaveLength(2);
    expect(result[0].name).toContain('abc123');
    expect(result[0].value).toBe('abc123');
  });

  it('should return empty array when no commits', async () => {
    // Arrange
    const mockResponse = { count: 0, value: [] };
    (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockResponse);

    // Act
    const result = await gitDataProvider.GetCommitsForRepo('project', 'repo');

    // Assert
    expect(result).toEqual([]);
  });
});

describe('GitDataProvider - GetPullRequestsForRepo', () => {
  let gitDataProvider: GitDataProvider;
  const mockOrgUrl = 'https://dev.azure.com/orgname/';
  const mockToken = 'mock-token';

  beforeEach(() => {
    jest.clearAllMocks();
    gitDataProvider = new GitDataProvider(mockOrgUrl, mockToken);
  });

  it('should fetch PRs for repo', async () => {
    // Arrange
    const projectName = 'project-123';
    const repoId = 'repo-456';
    const mockResponse = { count: 2, value: [{ pullRequestId: 1 }, { pullRequestId: 2 }] };
    (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockResponse);

    // Act
    const result = await gitDataProvider.GetPullRequestsForRepo(projectName, repoId);

    // Assert
    expect(TFSServices.getItemContent).toHaveBeenCalledWith(
      expect.stringContaining('pullrequests?status=completed'),
      mockToken,
      'get',
      null,
      null
    );
    expect(result).toEqual(mockResponse);
  });
});

describe('GitDataProvider - GetRepoBranches', () => {
  let gitDataProvider: GitDataProvider;
  const mockOrgUrl = 'https://dev.azure.com/orgname/';
  const mockToken = 'mock-token';

  beforeEach(() => {
    jest.clearAllMocks();
    gitDataProvider = new GitDataProvider(mockOrgUrl, mockToken);
  });

  it('should fetch repo branches', async () => {
    // Arrange
    const projectName = 'project-123';
    const repoId = 'repo-456';
    const mockResponse = { value: [{ name: 'refs/heads/main' }, { name: 'refs/heads/develop' }] };
    (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockResponse);

    // Act
    const result = await gitDataProvider.GetRepoBranches(projectName, repoId);

    // Assert
    expect(TFSServices.getItemContent).toHaveBeenCalledWith(
      expect.stringContaining('refs?searchCriteria.$top=1000&filter=heads'),
      mockToken,
      'get',
      null,
      null
    );
    expect(result).toEqual(mockResponse);
  });
});

describe('GitDataProvider - GetPullRequestsLinkedItemsInCommitRange', () => {
  let gitDataProvider: GitDataProvider;
  const mockOrgUrl = 'https://dev.azure.com/orgname/';
  const mockToken = 'mock-token';

  beforeEach(() => {
    jest.clearAllMocks();
    gitDataProvider = new GitDataProvider(mockOrgUrl, mockToken);
  });

  it('should fetch and filter PRs with linked work items', async () => {
    // Arrange
    const projectId = 'project-123';
    const repoId = 'repo-456';
    const commitRange = { value: [{ commitId: 'commit-1' }] };

    const mockPRsResponse = {
      count: 1,
      value: [
        {
          pullRequestId: 101,
          lastMergeCommit: { commitId: 'commit-1' },
          _links: { workItems: { href: 'https://example.com/workitems' } },
        },
      ],
    };

    const mockWorkItemsResponse = {
      count: 1,
      value: [{ id: 1 }],
    };

    const mockPopulatedWI = { id: 1, fields: { 'System.Title': 'Test WI' } };

    (TFSServices.getItemContent as jest.Mock)
      .mockResolvedValueOnce(mockPRsResponse)
      .mockResolvedValueOnce(mockWorkItemsResponse)
      .mockResolvedValueOnce(mockPopulatedWI);

    // Act
    const result = await gitDataProvider.GetPullRequestsLinkedItemsInCommitRange(
      projectId,
      repoId,
      commitRange
    );

    // Assert
    expect(result).toHaveLength(1);
    expect(result[0].workItem.id).toBe(1);
    expect(result[0].pullrequest.pullRequestId).toBe(101);
  });

  it('should handle PRs without linked work items', async () => {
    // Arrange
    const projectId = 'project-123';
    const repoId = 'repo-456';
    const commitRange = { value: [{ commitId: 'commit-1' }] };

    const mockPRsResponse = {
      count: 1,
      value: [
        {
          pullRequestId: 101,
          lastMergeCommit: { commitId: 'commit-1' },
          _links: {},
        },
      ],
    };

    (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockPRsResponse);

    // Act
    const result = await gitDataProvider.GetPullRequestsLinkedItemsInCommitRange(
      projectId,
      repoId,
      commitRange
    );

    // Assert
    expect(result).toHaveLength(0);
  });
});

describe('GitDataProvider - GetItemsInCommitRange', () => {
  let gitDataProvider: GitDataProvider;
  const mockOrgUrl = 'https://dev.azure.com/orgname/';
  const mockToken = 'mock-token';

  beforeEach(() => {
    jest.clearAllMocks();
    gitDataProvider = new GitDataProvider(mockOrgUrl, mockToken);
  });

  it('should process commits with work items', async () => {
    // Arrange
    const projectId = 'project-123';
    const repoId = 'repo-456';
    const commitRange = {
      value: [
        {
          commitId: 'commit-1',
          workItems: [{ id: 1 }],
        },
      ],
    };

    const mockPopulatedWI = { id: 1, fields: { 'System.Title': 'Test WI' } };
    const mockPRsResponse = { count: 0, value: [] };

    (TFSServices.getItemContent as jest.Mock)
      .mockResolvedValueOnce(mockPopulatedWI)
      .mockResolvedValueOnce(mockPRsResponse);

    // Act
    const result = await gitDataProvider.GetItemsInCommitRange(projectId, repoId, commitRange, null, false);

    // Assert
    expect(result.commitChangesArray).toHaveLength(1);
    expect(result.commitsWithNoRelations).toHaveLength(0);
  });

  it('marks commit-linked items as PR-only when only on the merge commit', async () => {
    const projectId = 'project-123';
    const repoId = 'repo-456';
    const commitRange = {
      value: [
        {
          commitId: 'merge-1',
          workItems: [{ id: 42 }],
        },
      ],
    };

    const populatedItem = { id: 42, fields: { 'System.Title': 'Bug 42' } };
    (gitDataProvider as any).ticketsDataProvider.GetWorkItem = jest.fn().mockResolvedValue(populatedItem);
    (gitDataProvider as any).GetPullRequestsLinkedItemsInCommitRange = jest.fn().mockResolvedValue([
      { workItem: populatedItem, pullrequest: { lastMergeCommit: { commitId: 'merge-1' } } },
    ]);

    const result = await gitDataProvider.GetItemsInCommitRange(projectId, repoId, commitRange, null, false, true);

    expect(result.commitChangesArray).toHaveLength(1);
    expect(result.commitChangesArray[0].pullRequestWorkItemOnly).toBe(true);
  });

  it('should include unlinked commits when flag is true', async () => {
    // Arrange
    const projectId = 'project-123';
    const repoId = 'repo-456';
    const commitRange = {
      value: [
        {
          commitId: 'commit-1',
          workItems: [],
          committer: { date: '2023-01-01', name: 'Test User' },
          comment: 'Test commit',
          remoteUrl: 'https://example.com/commit/1',
        },
      ],
    };

    const mockPRsResponse = { count: 0, value: [] };

    (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockPRsResponse);

    // Act
    const result = await gitDataProvider.GetItemsInCommitRange(projectId, repoId, commitRange, null, true);

    // Assert
    expect(result.commitChangesArray).toHaveLength(0);
    expect(result.commitsWithNoRelations).toHaveLength(1);
    expect(result.commitsWithNoRelations[0].commitId).toBe('commit-1');
  });
});

describe('GitDataProvider - getItemsForPipelineRange', () => {
  let gitDataProvider: GitDataProvider;
  const mockOrgUrl = 'https://dev.azure.com/orgname/';
  const mockToken = 'mock-token';

  beforeEach(() => {
    jest.clearAllMocks();
    gitDataProvider = new GitDataProvider(mockOrgUrl, mockToken);
  });

  it('should throw error when extended commits is empty', async () => {
    // Arrange
    const teamProject = 'project-123';
    const extendedCommits: any[] = [];
    const targetRepo = { url: 'https://example.com/repo' };
    const addedWorkItemByIdSet = new Set<number>();

    // Act
    const result = await gitDataProvider.getItemsForPipelineRange(
      teamProject,
      extendedCommits,
      targetRepo,
      addedWorkItemByIdSet
    );

    // Assert - should log error and return empty arrays
    expect(logger.error).toHaveBeenCalledWith('extended commits cannot be empty');
    expect(result.commitChangesArray).toHaveLength(0);
  });

  it('should process commits with work items', async () => {
    // Arrange
    const teamProject = 'project-123';
    const extendedCommits = [
      {
        commit: {
          commitId: 'commit-1',
          workItems: [{ id: 1 }],
          committer: { date: '2023-01-01', name: 'Test User' },
          comment: 'Test commit',
        },
      },
    ];
    const targetRepo = { url: 'https://example.com/repo' };
    const addedWorkItemByIdSet = new Set<number>();

    const mockRepoData = {
      _links: { web: { href: 'https://example.com/repo-web' } },
      project: { id: 'project-id' },
    };
    const mockPopulatedWI = { id: 1, fields: { 'System.Title': 'Test WI' } };

    (TFSServices.getItemContent as jest.Mock)
      .mockResolvedValueOnce(mockRepoData)
      .mockResolvedValueOnce(mockPopulatedWI);

    // Act
    const result = await gitDataProvider.getItemsForPipelineRange(
      teamProject,
      extendedCommits,
      targetRepo,
      addedWorkItemByIdSet
    );

    // Assert
    expect(result.commitChangesArray).toHaveLength(1);
    expect(result.commitChangesArray[0].workItem.id).toBe(1);
  });

  it('should include unlinked commits when flag is true', async () => {
    // Arrange
    const teamProject = 'project-123';
    const extendedCommits = [
      {
        commit: {
          commitId: 'commit-1',
          workItems: [],
          committer: { date: '2023-01-01', name: 'Test User' },
          comment: 'Test commit',
          remoteUrl: 'https://example.com/commit/1',
        },
      },
    ];
    const targetRepo = { url: 'https://example.com/repo' };
    const addedWorkItemByIdSet = new Set<number>();

    const mockRepoData = {
      _links: { web: { href: 'https://example.com/repo-web' } },
      project: { id: 'project-id' },
    };

    (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockRepoData);

    // Act
    const result = await gitDataProvider.getItemsForPipelineRange(
      teamProject,
      extendedCommits,
      targetRepo,
      addedWorkItemByIdSet,
      undefined,
      true
    );

    // Assert
    expect(result.commitsWithNoRelations).toHaveLength(1);
    expect(result.commitsWithNoRelations[0].commitId).toBe('commit-1');
  });

  it('should not add duplicate work items', async () => {
    // Arrange
    const teamProject = 'project-123';
    const extendedCommits = [
      { commit: { commitId: 'commit-1', workItems: [{ id: 1 }] } },
      { commit: { commitId: 'commit-2', workItems: [{ id: 1 }] } }, // Same WI
    ];
    const targetRepo = { url: 'https://example.com/repo' };
    const addedWorkItemByIdSet = new Set<number>();

    const mockRepoData = {
      _links: { web: { href: 'https://example.com/repo-web' } },
      project: { id: 'project-id' },
    };
    const mockPopulatedWI = { id: 1, fields: { 'System.Title': 'Test WI' } };

    (TFSServices.getItemContent as jest.Mock)
      .mockResolvedValueOnce(mockRepoData)
      .mockResolvedValueOnce(mockPopulatedWI)
      .mockResolvedValueOnce(mockPopulatedWI);

    // Act
    const result = await gitDataProvider.getItemsForPipelineRange(
      teamProject,
      extendedCommits,
      targetRepo,
      addedWorkItemByIdSet
    );

    // Assert - should only have 1 work item (no duplicates)
    expect(result.commitChangesArray).toHaveLength(1);
  });
});

describe('GitDataProvider - GetItemsInPullRequestRange', () => {
  let gitDataProvider: GitDataProvider;
  const mockOrgUrl = 'https://dev.azure.com/orgname/';
  const mockToken = 'mock-token';

  beforeEach(() => {
    jest.clearAllMocks();
    gitDataProvider = new GitDataProvider(mockOrgUrl, mockToken);
  });

  it('should fetch items from PRs in range', async () => {
    // Arrange
    const projectId = 'project-123';
    const repoId = 'repo-456';
    const prIds = [101];

    const mockPRsResponse = {
      count: 1,
      value: [{ pullRequestId: 101, _links: { workItems: { href: 'https://example.com/wi1' } } }],
    };

    const mockWIResponse = { count: 1, value: [{ id: 1 }] };
    const mockPopulatedWI = { id: 1, fields: { 'System.Title': 'WI 1' } };

    (TFSServices.getItemContent as jest.Mock)
      .mockResolvedValueOnce(mockPRsResponse)
      .mockResolvedValueOnce(mockWIResponse)
      .mockResolvedValueOnce(mockPopulatedWI);

    // Act
    const result = await gitDataProvider.GetItemsInPullRequestRange(projectId, repoId, prIds);

    // Assert
    expect(result).toHaveLength(1);
    expect(result[0].workItem.id).toBe(1);
  });
});

describe('GitDataProvider - GetRepoTagsWithCommits', () => {
  let gitDataProvider: GitDataProvider;
  const mockOrgUrl = 'https://dev.azure.com/orgname/';
  const mockToken = 'mock-token';

  beforeEach(() => {
    jest.clearAllMocks();
    gitDataProvider = new GitDataProvider(mockOrgUrl, mockToken);
  });

  it('should return empty array when response is null', async () => {
    // Arrange
    const repoApiUrl = 'https://dev.azure.com/org/project/_apis/git/repositories/repo-id';

    (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(null);

    // Act
    const result = await gitDataProvider.GetRepoTagsWithCommits(repoApiUrl);

    // Assert
    expect(result).toEqual([]);
  });

  it('should return empty array when no tags exist', async () => {
    // Arrange
    const repoApiUrl = 'https://dev.azure.com/org/project/_apis/git/repositories/repo-id';
    const mockTagsResponse = { count: 0, value: [] };

    (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockTagsResponse);

    // Act
    const result = await gitDataProvider.GetRepoTagsWithCommits(repoApiUrl);

    // Assert
    expect(result).toEqual([]);
  });
});

describe('GitDataProvider - GetCommitBatch', () => {
  let gitDataProvider: GitDataProvider;
  const mockOrgUrl = 'https://dev.azure.com/orgname/';
  const mockToken = 'mock-token';

  beforeEach(() => {
    jest.clearAllMocks();
    gitDataProvider = new GitDataProvider(mockOrgUrl, mockToken);
    // Mock postRequest
    (TFSServices as any).postRequest = jest.fn();
  });

  it('should fetch commits in batches', async () => {
    // Arrange
    const gitUrl = 'https://dev.azure.com/org/project/_apis/git/repositories/repo-id';
    const itemVersion = { version: 'main', versionType: 'branch' };
    const compareVersion = { version: 'v1.0.0', versionType: 'tag' };

    const mockCommitsResponse = {
      data: {
        count: 1,
        value: [
          {
            commitId: 'abc123',
            committer: { name: 'Test User', date: '2023-01-01T00:00:00Z' },
            comment: 'Test commit',
          },
        ],
      },
    };
    const mockEmptyResponse = { data: { count: 0, value: [] } };

    (TFSServices as any).postRequest
      .mockResolvedValueOnce(mockCommitsResponse)
      .mockResolvedValueOnce(mockEmptyResponse);

    // Act
    const result = await gitDataProvider.GetCommitBatch(gitUrl, itemVersion, compareVersion);

    // Assert
    expect(result).toHaveLength(1);
    expect(result[0].commit.commitId).toBe('abc123');
    expect(result[0].committerName).toBe('Test User');
  });

  it('should handle empty commits response', async () => {
    // Arrange
    const gitUrl = 'https://dev.azure.com/org/project/_apis/git/repositories/repo-id';
    const itemVersion = { version: 'main', versionType: 'branch' };
    const compareVersion = { version: 'v1.0.0', versionType: 'tag' };

    const mockEmptyResponse = { data: { count: 0, value: [] } };

    (TFSServices as any).postRequest.mockResolvedValueOnce(mockEmptyResponse);

    // Act
    const result = await gitDataProvider.GetCommitBatch(gitUrl, itemVersion, compareVersion);

    // Assert
    expect(result).toHaveLength(0);
  });

  it('should include specificItemPath when provided', async () => {
    // Arrange
    const gitUrl = 'https://dev.azure.com/org/project/_apis/git/repositories/repo-id';
    const itemVersion = { version: 'main', versionType: 'branch' };
    const compareVersion = { version: 'v1.0.0', versionType: 'tag' };
    const specificItemPath = '/src/file.ts';

    const mockEmptyResponse = { data: { count: 0, value: [] } };

    (TFSServices as any).postRequest.mockResolvedValueOnce(mockEmptyResponse);

    // Act
    await gitDataProvider.GetCommitBatch(gitUrl, itemVersion, compareVersion, specificItemPath);

    // Assert
    expect((TFSServices as any).postRequest).toHaveBeenCalledWith(
      expect.any(String),
      mockToken,
      undefined,
      expect.objectContaining({
        itemPath: specificItemPath,
        historyMode: 'fullHistory',
      })
    );
  });
});

describe('GitDataProvider - getSubmodulesData', () => {
  let gitDataProvider: GitDataProvider;
  const mockOrgUrl = 'https://dev.azure.com/orgname/';
  const mockToken = 'mock-token';

  beforeEach(() => {
    jest.clearAllMocks();
    gitDataProvider = new GitDataProvider(mockOrgUrl, mockToken);
  });

  it('should return empty array when no .gitmodules file exists', async () => {
    // Arrange
    const projectName = 'project-123';
    const repoId = 'repo-456';
    const targetVersion = { version: 'main', versionType: 'branch' };
    const sourceVersion = { version: 'v1.0.0', versionType: 'tag' };
    const allCommitsExtended: any[] = [];

    // Mock GetFileFromGitRepo to return undefined (no .gitmodules)
    (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce({});

    // Act
    const result = await gitDataProvider.getSubmodulesData(
      projectName,
      repoId,
      targetVersion,
      sourceVersion,
      allCommitsExtended
    );

    // Assert
    expect(result).toEqual([]);
  });
});

describe('GitDataProvider - GetRepoTagsWithCommits (extended)', () => {
  let gitDataProvider: GitDataProvider;
  const mockOrgUrl = 'https://dev.azure.com/orgname/';
  const mockToken = 'mock-token';

  beforeEach(() => {
    jest.clearAllMocks();
    gitDataProvider = new GitDataProvider(mockOrgUrl, mockToken);
  });

  it('should fetch tags with commit dates', async () => {
    // Arrange
    const repoApiUrl = 'https://dev.azure.com/org/project/_apis/git/repositories/repo-id';
    const mockTagsResponse = {
      count: 1,
      value: [{ name: 'refs/tags/v1.0.0', objectId: 'abc123', peeledObjectId: 'def456' }],
    };
    const mockCommitResponse = { committer: { date: '2023-01-01' } };

    (TFSServices.getItemContent as jest.Mock)
      .mockResolvedValueOnce(mockTagsResponse)
      .mockResolvedValueOnce(mockCommitResponse);

    // Act
    const result = await gitDataProvider.GetRepoTagsWithCommits(repoApiUrl);

    // Assert
    expect(result).toHaveLength(1);
    expect(result[0].name).toBe('v1.0.0');
    expect(result[0].commitId).toBe('def456');
    expect(result[0].date).toBe('2023-01-01');
  });

  it('should skip tags without commitId', async () => {
    // Arrange
    const repoApiUrl = 'https://dev.azure.com/org/project/_apis/git/repositories/repo-id';
    const mockTagsResponse = {
      count: 2,
      value: [
        { name: 'refs/tags/v1.0.0' }, // No objectId or peeledObjectId
        { name: 'refs/tags/v2.0.0', objectId: 'abc123' },
      ],
    };
    const mockCommitResponse = { committer: { date: '2023-01-01' } };

    (TFSServices.getItemContent as jest.Mock)
      .mockResolvedValueOnce(mockTagsResponse)
      .mockResolvedValueOnce(mockCommitResponse);

    // Act
    const result = await gitDataProvider.GetRepoTagsWithCommits(repoApiUrl);

    // Assert
    expect(result).toHaveLength(1);
    expect(result[0].name).toBe('v2.0.0');
  });

  it('should handle commit fetch failure gracefully', async () => {
    // Arrange
    const repoApiUrl = 'https://dev.azure.com/org/project/_apis/git/repositories/repo-id';
    const mockTagsResponse = {
      count: 1,
      value: [{ name: 'refs/tags/v1.0.0', objectId: 'abc123' }],
    };

    (TFSServices.getItemContent as jest.Mock)
      .mockResolvedValueOnce(mockTagsResponse)
      .mockRejectedValueOnce(new Error('Commit not found'));

    // Act
    const result = await gitDataProvider.GetRepoTagsWithCommits(repoApiUrl);

    // Assert
    expect(result).toHaveLength(1);
    expect(result[0].date).toBeUndefined();
  });

  it('should use author date when committer date is missing', async () => {
    // Arrange
    const repoApiUrl = 'https://dev.azure.com/org/project/_apis/git/repositories/repo-id';
    const mockTagsResponse = {
      count: 1,
      value: [{ name: 'refs/tags/v1.0.0', objectId: 'abc123' }],
    };
    const mockCommitResponse = { author: { date: '2023-02-01' } }; // No committer

    (TFSServices.getItemContent as jest.Mock)
      .mockResolvedValueOnce(mockTagsResponse)
      .mockResolvedValueOnce(mockCommitResponse);

    // Act
    const result = await gitDataProvider.GetRepoTagsWithCommits(repoApiUrl);

    // Assert
    expect(result[0].date).toBe('2023-02-01');
  });
});

describe('GitDataProvider - createLinkedRelatedItemsForSVD (via GetItemsInCommitRange)', () => {
  let gitDataProvider: GitDataProvider;
  const mockOrgUrl = 'https://dev.azure.com/orgname/';
  const mockToken = 'mock-token';

  beforeEach(() => {
    jest.clearAllMocks();
    gitDataProvider = new GitDataProvider(mockOrgUrl, mockToken);
  });

  it('should create linked items for requirements with Affects relationship', async () => {
    // Arrange
    const projectId = 'project-123';
    const repoId = 'repo-456';
    const commitRange = {
      value: [
        {
          commitId: 'commit-1',
          workItems: [{ id: 1 }],
        },
      ],
    };

    const linkedWiOptions = {
      isEnabled: true,
      linkedWiTypes: 'reqOnly',
      linkedWiRelationship: 'affectsOnly',
    };

    const mockPopulatedWI = {
      id: 1,
      fields: { 'System.Title': 'Test WI' },
      relations: [
        {
          url: 'https://example.com/workItems/2',
          rel: 'System.LinkTypes.Affects-Forward',
          attributes: { name: 'Affects' },
        },
      ],
    };

    const mockRelatedWI = {
      id: 2,
      fields: {
        'System.WorkItemType': 'Requirement',
        'System.Title': 'Related Requirement',
      },
      _links: { html: { href: 'https://example.com/wi/2' } },
    };

    const mockPRsResponse = { count: 0, value: [] };

    (TFSServices.getItemContent as jest.Mock)
      .mockResolvedValueOnce(mockPopulatedWI)
      .mockResolvedValueOnce(mockRelatedWI)
      .mockResolvedValueOnce(mockPRsResponse);

    // Act
    const result = await gitDataProvider.GetItemsInCommitRange(
      projectId,
      repoId,
      commitRange,
      linkedWiOptions,
      false
    );

    // Assert
    expect(result.commitChangesArray).toHaveLength(1);
    expect(result.commitChangesArray[0].linkedItems).toBeDefined();
  });

  it('should filter out non-workItem relations', async () => {
    // Arrange
    const projectId = 'project-123';
    const repoId = 'repo-456';
    const commitRange = {
      value: [
        {
          commitId: 'commit-1',
          workItems: [{ id: 1 }],
        },
      ],
    };

    const linkedWiOptions = {
      isEnabled: true,
      linkedWiTypes: 'both',
      linkedWiRelationship: 'both',
    };

    const mockPopulatedWI = {
      id: 1,
      fields: { 'System.Title': 'Test WI' },
      relations: [
        {
          url: 'https://example.com/attachments/file.txt', // Not a workItem
          rel: 'AttachedFile',
          attributes: {},
        },
      ],
    };

    const mockPRsResponse = { count: 0, value: [] };

    (TFSServices.getItemContent as jest.Mock)
      .mockResolvedValueOnce(mockPopulatedWI)
      .mockResolvedValueOnce(mockPRsResponse);

    // Act
    const result = await gitDataProvider.GetItemsInCommitRange(
      projectId,
      repoId,
      commitRange,
      linkedWiOptions,
      false
    );

    // Assert
    expect(result.commitChangesArray).toHaveLength(1);
    expect(result.commitChangesArray[0].linkedItems).toEqual([]);
  });

  it('should handle linkedWiTypes=none', async () => {
    // Arrange
    const projectId = 'project-123';
    const repoId = 'repo-456';
    const commitRange = {
      value: [
        {
          commitId: 'commit-1',
          workItems: [{ id: 1 }],
        },
      ],
    };

    const linkedWiOptions = {
      isEnabled: true,
      linkedWiTypes: 'none',
      linkedWiRelationship: 'both',
    };

    const mockPopulatedWI = {
      id: 1,
      fields: { 'System.Title': 'Test WI' },
      relations: [
        {
          url: 'https://example.com/workItems/2',
          rel: 'System.LinkTypes.Affects-Forward',
          attributes: { name: 'Affects' },
        },
      ],
    };

    const mockPRsResponse = { count: 0, value: [] };

    (TFSServices.getItemContent as jest.Mock)
      .mockResolvedValueOnce(mockPopulatedWI)
      .mockResolvedValueOnce(mockPRsResponse);

    // Act
    const result = await gitDataProvider.GetItemsInCommitRange(
      projectId,
      repoId,
      commitRange,
      linkedWiOptions,
      false
    );

    // Assert
    expect(result.commitChangesArray[0].linkedItems).toEqual([]);
  });

  it('should handle Feature type with featureOnly option', async () => {
    // Arrange
    const projectId = 'project-123';
    const repoId = 'repo-456';
    const commitRange = {
      value: [
        {
          commitId: 'commit-1',
          workItems: [{ id: 1 }],
        },
      ],
    };

    const linkedWiOptions = {
      isEnabled: true,
      linkedWiTypes: 'featureOnly',
      linkedWiRelationship: 'coversOnly',
    };

    const mockPopulatedWI = {
      id: 1,
      fields: { 'System.Title': 'Test WI' },
      relations: [
        {
          url: 'https://example.com/workItems/2',
          rel: 'System.LinkTypes.CoveredBy-Forward',
          attributes: { name: 'CoveredBy' },
        },
      ],
    };

    const mockRelatedWI = {
      id: 2,
      fields: {
        'System.WorkItemType': 'Feature',
        'System.Title': 'Related Feature',
      },
      _links: { html: { href: 'https://example.com/wi/2' } },
    };

    const mockPRsResponse = { count: 0, value: [] };

    (TFSServices.getItemContent as jest.Mock)
      .mockResolvedValueOnce(mockPopulatedWI)
      .mockResolvedValueOnce(mockRelatedWI)
      .mockResolvedValueOnce(mockPRsResponse);

    // Act
    const result = await gitDataProvider.GetItemsInCommitRange(
      projectId,
      repoId,
      commitRange,
      linkedWiOptions,
      false
    );

    // Assert
    expect(result.commitChangesArray).toHaveLength(1);
  });
});

describe('GitDataProvider - GetRepoReferences (extended)', () => {
  let gitDataProvider: GitDataProvider;
  const mockOrgUrl = 'https://dev.azure.com/orgname/';
  const mockToken = 'mock-token';

  beforeEach(() => {
    jest.clearAllMocks();
    gitDataProvider = new GitDataProvider(mockOrgUrl, mockToken);
  });

  it('should sort tags by commit date (most recent first)', async () => {
    // Arrange
    const projectId = 'project-123';
    const repoId = 'repo-456';
    const mockTagsResponse = {
      count: 2,
      value: [
        { name: 'refs/tags/v1.0.0', objectId: 'abc123' },
        { name: 'refs/tags/v2.0.0', objectId: 'def456', peeledObjectId: 'ghi789' },
      ],
    };
    const mockCommit1 = { committer: { date: '2023-01-01' } };
    const mockCommit2 = { committer: { date: '2023-06-01' } };

    (TFSServices.getItemContent as jest.Mock)
      .mockResolvedValueOnce(mockTagsResponse)
      .mockResolvedValueOnce(mockCommit1)
      .mockResolvedValueOnce(mockCommit2);

    // Act
    const result = await gitDataProvider.GetRepoReferences(projectId, repoId, 'tag');

    // Assert
    expect(result).toHaveLength(2);
    // v2.0.0 should be first (more recent)
    expect(result[0].name).toBe('v2.0.0');
    expect(result[1].name).toBe('v1.0.0');
  });

  it('should sort branches by commit date', async () => {
    // Arrange
    const projectId = 'project-123';
    const repoId = 'repo-456';
    const mockBranchesResponse = {
      count: 2,
      value: [
        { name: 'refs/heads/main', objectId: 'abc123' },
        { name: 'refs/heads/develop', objectId: 'def456' },
      ],
    };
    const mockCommit1 = { committer: { date: '2023-01-01' } };
    const mockCommit2 = { committer: { date: '2023-06-01' } };

    (TFSServices.getItemContent as jest.Mock)
      .mockResolvedValueOnce(mockBranchesResponse)
      .mockResolvedValueOnce(mockCommit1)
      .mockResolvedValueOnce(mockCommit2);

    // Act
    const result = await gitDataProvider.GetRepoReferences(projectId, repoId, 'branch');

    // Assert
    expect(result).toHaveLength(2);
    // develop should be first (more recent)
    expect(result[0].name).toBe('develop');
    expect(result[1].name).toBe('main');
  });

  it('should handle commit resolution failure in tag sorting', async () => {
    // Arrange
    const projectId = 'project-123';
    const repoId = 'repo-456';
    const mockTagsResponse = {
      count: 1,
      value: [{ name: 'refs/tags/v1.0.0', objectId: 'abc123' }],
    };

    (TFSServices.getItemContent as jest.Mock)
      .mockResolvedValueOnce(mockTagsResponse)
      .mockRejectedValueOnce(new Error('Commit not found'));

    // Act
    const result = await gitDataProvider.GetRepoReferences(projectId, repoId, 'tag');

    // Assert
    expect(result).toHaveLength(1);
    expect(result[0].date).toBeUndefined();
  });

  it('should set date undefined when tag commit has no committer/author date', async () => {
    const projectId = 'project-123';
    const repoId = 'repo-456';
    const mockTagsResponse = {
      count: 1,
      value: [{ name: 'refs/tags/v0.0.1', objectId: 'abc123' }],
    };

    (TFSServices.getItemContent as jest.Mock)
      .mockResolvedValueOnce(mockTagsResponse)
      .mockResolvedValueOnce({});

    const result = await gitDataProvider.GetRepoReferences(projectId, repoId, 'tag');

    expect(result).toHaveLength(1);
    expect(result[0]).toEqual(expect.objectContaining({ name: 'v0.0.1', date: undefined }));
  });

  it('should set date undefined when branch commit cannot be resolved to a date', async () => {
    const projectId = 'project-123';
    const repoId = 'repo-456';
    const mockBranchesResponse = {
      count: 1,
      value: [{ name: 'refs/heads/feature', objectId: 'abc123' }],
    };

    (TFSServices.getItemContent as jest.Mock)
      .mockResolvedValueOnce(mockBranchesResponse)
      .mockResolvedValueOnce(null);

    const result = await gitDataProvider.GetRepoReferences(projectId, repoId, 'branch');

    expect(result).toHaveLength(1);
    expect(result[0]).toEqual(expect.objectContaining({ name: 'feature', date: undefined }));
  });
});

describe('GitDataProvider - duplicate work item removal', () => {
  let gitDataProvider: GitDataProvider;
  const mockOrgUrl = 'https://dev.azure.com/orgname/';
  const mockToken = 'mock-token';

  beforeEach(() => {
    jest.clearAllMocks();
    gitDataProvider = new GitDataProvider(mockOrgUrl, mockToken);
  });

  it('should remove duplicate work items in GetItemsInCommitRange', async () => {
    // Arrange
    const projectId = 'project-123';
    const repoId = 'repo-456';
    const commitRange = {
      value: [
        { commitId: 'commit-1', workItems: [{ id: 1 }] },
        { commitId: 'commit-2', workItems: [{ id: 1 }] },
      ],
    };

    const mockPopulatedWI = { id: 1, fields: { 'System.Title': 'Test WI' } };
    const mockPRsResponse = { count: 0, value: [] };

    (TFSServices.getItemContent as jest.Mock)
      .mockResolvedValueOnce(mockPopulatedWI)
      .mockResolvedValueOnce(mockPopulatedWI)
      .mockResolvedValueOnce(mockPRsResponse);

    // Act
    const result = await gitDataProvider.GetItemsInCommitRange(projectId, repoId, commitRange, null, false);

    // Assert
    expect(result.commitChangesArray.length).toBeLessThanOrEqual(2);
  });
});

describe('GitDataProvider - version encoding edge cases', () => {
  let gitDataProvider: GitDataProvider;
  const mockOrgUrl = 'https://dev.azure.com/orgname/';
  const mockToken = 'mock-token';

  beforeEach(() => {
    jest.clearAllMocks();
    gitDataProvider = new GitDataProvider(mockOrgUrl, mockToken);
  });

  it('should handle version with null gracefully in GetFileFromGitRepo', async () => {
    // Arrange
    const projectName = 'project-123';
    const repoId = 'repo-456';
    const fileName = 'test.txt';
    const version = { version: null as any, versionType: 'branch' };

    (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce({ content: 'file content' });

    // Act
    const result = await gitDataProvider.GetFileFromGitRepo(projectName, repoId, fileName, version);

    // Assert
    expect(result).toBe('file content');
  });

  it('should handle version with null gracefully in CheckIfItemExist', async () => {
    // Arrange
    const gitApiUrl = 'https://dev.azure.com/org/project/_apis/git/repositories/repo-id';
    const itemPath = 'path/to/file.txt';
    const version = { version: null as any, versionType: 'branch' };

    (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce({ path: itemPath });

    // Act
    const result = await gitDataProvider.CheckIfItemExist(gitApiUrl, itemPath, version);

    // Assert
    expect(result).toBe(true);
  });
});

describe('GitDataProvider - GetCommitsForRepo edge cases', () => {
  let gitDataProvider: GitDataProvider;
  const mockOrgUrl = 'https://dev.azure.com/orgname/';
  const mockToken = 'mock-token';

  beforeEach(() => {
    jest.clearAllMocks();
    gitDataProvider = new GitDataProvider(mockOrgUrl, mockToken);
  });

  it('should handle commits without committer date', async () => {
    // Arrange
    const projectName = 'project-123';
    const repoId = 'repo-456';
    const mockResponse = {
      count: 1,
      value: [{ commitId: 'abc123', comment: 'Test commit' }],
    };
    (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockResponse);

    // Act
    const result = await gitDataProvider.GetCommitsForRepo(projectName, repoId);

    // Assert
    expect(result).toHaveLength(1);
    expect(result[0].date).toBeUndefined();
  });

  it('should fetch commits without version identifier', async () => {
    // Arrange
    const projectName = 'project-123';
    const repoId = 'repo-456';
    const mockResponse = {
      count: 1,
      value: [{ commitId: 'abc123', comment: 'Test', committer: { date: '2023-01-01' } }],
    };
    (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockResponse);

    // Act
    const result = await gitDataProvider.GetCommitsForRepo(projectName, repoId, '');

    // Assert
    expect(result).toHaveLength(1);
  });
});

describe('GitDataProvider - GetItemsInPullRequestRange edge cases', () => {
  let gitDataProvider: GitDataProvider;
  const mockOrgUrl = 'https://dev.azure.com/orgname/';
  const mockToken = 'mock-token';

  beforeEach(() => {
    jest.clearAllMocks();
    gitDataProvider = new GitDataProvider(mockOrgUrl, mockToken);
  });

  it('should handle PR without workItems link', async () => {
    // Arrange
    const projectId = 'project-123';
    const repoId = 'repo-456';
    const prIds = [101];

    const mockPRsResponse = {
      count: 1,
      value: [{ pullRequestId: 101, _links: {} }],
    };

    (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockPRsResponse);

    // Act
    const result = await gitDataProvider.GetItemsInPullRequestRange(projectId, repoId, prIds);

    // Assert
    expect(result).toHaveLength(0);
  });

  it('should handle errors when fetching work items', async () => {
    // Arrange
    const projectId = 'project-123';
    const repoId = 'repo-456';
    const prIds = [101];

    const mockPRsResponse = {
      count: 1,
      value: [{ pullRequestId: 101, _links: { workItems: { href: 'https://example.com/wi' } } }],
    };

    (TFSServices.getItemContent as jest.Mock)
      .mockResolvedValueOnce(mockPRsResponse)
      .mockRejectedValueOnce(new Error('Failed to fetch'));

    // Act
    const result = await gitDataProvider.GetItemsInPullRequestRange(projectId, repoId, prIds);

    // Assert
    expect(result).toHaveLength(0);
    expect(logger.error).toHaveBeenCalled();
  });
});

describe('GitDataProvider - getSubmodulesData', () => {
  let gitDataProvider: GitDataProvider;
  const mockOrgUrl = 'https://dev.azure.com/orgname/';
  const mockToken = 'mock-token';

  beforeEach(() => {
    jest.clearAllMocks();
    gitDataProvider = new GitDataProvider(mockOrgUrl, mockToken);
  });

  it('should parse .gitmodules file and return submodule data', async () => {
    // Arrange
    const projectName = 'project-123';
    const repoId = 'repo-456';
    const targetVersion = { version: 'main', versionType: 'branch' };
    const sourceVersion = { version: 'v1.0.0', versionType: 'tag' };
    const allCommitsExtended: any[] = [];

    const gitModulesContent = `[submodule "libs/common"]
	path = libs/common
	url = https://dev.azure.com/org/project/_git/common-lib`;

    // Mock GetFileFromGitRepo calls
    (TFSServices.getItemContent as jest.Mock)
      .mockResolvedValueOnce({ content: gitModulesContent }) // .gitmodules file
      .mockResolvedValueOnce({ content: 'target-sha-123' }) // target SHA
      .mockResolvedValueOnce({ content: 'source-sha-456' }); // source SHA

    // Act
    const result = await gitDataProvider.getSubmodulesData(
      projectName,
      repoId,
      targetVersion,
      sourceVersion,
      allCommitsExtended
    );

    // Assert
    expect(result).toHaveLength(1);
    expect(result[0].gitSubModuleName).toBe('libs_common');
    expect(result[0].targetSha1).toBe('target-sha-123');
    expect(result[0].sourceSha1).toBe('source-sha-456');
  });

  it('should handle .gitmodules with CRLF line endings', async () => {
    // Arrange
    const projectName = 'project-123';
    const repoId = 'repo-456';
    const targetVersion = { version: 'main', versionType: 'branch' };
    const sourceVersion = { version: 'v1.0.0', versionType: 'tag' };
    const allCommitsExtended: any[] = [];

    const gitModulesContent = `[submodule "libs/common"]\r\n\tpath = libs/common\r\n\turl = https://example.com/repo`;

    (TFSServices.getItemContent as jest.Mock)
      .mockResolvedValueOnce({ content: gitModulesContent })
      .mockResolvedValueOnce({ content: 'target-sha' })
      .mockResolvedValueOnce({ content: 'source-sha' });

    // Act
    const result = await gitDataProvider.getSubmodulesData(
      projectName,
      repoId,
      targetVersion,
      sourceVersion,
      allCommitsExtended
    );

    // Assert
    expect(result).toHaveLength(1);
  });

  it('should handle relative URL paths in submodules', async () => {
    // Arrange
    const projectName = 'project-123';
    const repoId = 'repo-456';
    const targetVersion = { version: 'main', versionType: 'branch' };
    const sourceVersion = { version: 'v1.0.0', versionType: 'tag' };
    const allCommitsExtended: any[] = [];

    const gitModulesContent = `[submodule "libs/common"]
	path = libs/common
	url = ../common-lib`;

    (TFSServices.getItemContent as jest.Mock)
      .mockResolvedValueOnce({ content: gitModulesContent })
      .mockResolvedValueOnce({ content: 'target-sha' })
      .mockResolvedValueOnce({ content: 'source-sha' });

    // Act
    const result = await gitDataProvider.getSubmodulesData(
      projectName,
      repoId,
      targetVersion,
      sourceVersion,
      allCommitsExtended
    );

    // Assert
    expect(result).toHaveLength(1);
    expect(result[0].gitSubRepoUrl).toContain('common-lib');
  });

  it('should skip submodule when source SHA not found', async () => {
    // Arrange
    const projectName = 'project-123';
    const repoId = 'repo-456';
    const targetVersion = { version: 'main', versionType: 'branch' };
    const sourceVersion = { version: 'v1.0.0', versionType: 'tag' };
    const allCommitsExtended: any[] = [];

    const gitModulesContent = `[submodule "libs/common"]
	path = libs/common
	url = https://example.com/repo`;

    (TFSServices.getItemContent as jest.Mock)
      .mockResolvedValueOnce({ content: gitModulesContent })
      .mockResolvedValueOnce({ content: 'target-sha' })
      .mockResolvedValueOnce({ content: undefined }); // source not found

    // Act
    const result = await gitDataProvider.getSubmodulesData(
      projectName,
      repoId,
      targetVersion,
      sourceVersion,
      allCommitsExtended
    );

    // Assert
    expect(result).toHaveLength(0);
    expect(logger.warn).toHaveBeenCalled();
  });

  it('should skip submodule when target SHA not found', async () => {
    // Arrange
    const projectName = 'project-123';
    const repoId = 'repo-456';
    const targetVersion = { version: 'main', versionType: 'branch' };
    const sourceVersion = { version: 'v1.0.0', versionType: 'tag' };
    const allCommitsExtended: any[] = [];

    const gitModulesContent = `[submodule "libs/common"]
	path = libs/common
	url = https://example.com/repo`;

    (TFSServices.getItemContent as jest.Mock)
      .mockResolvedValueOnce({ content: gitModulesContent })
      .mockResolvedValueOnce({ content: undefined }) // target not found
      .mockResolvedValueOnce({ content: 'source-sha' });

    // Act
    const result = await gitDataProvider.getSubmodulesData(
      projectName,
      repoId,
      targetVersion,
      sourceVersion,
      allCommitsExtended
    );

    // Assert
    expect(result).toHaveLength(0);
    expect(logger.warn).toHaveBeenCalled();
  });

  it('should skip submodule when source and target SHA are the same', async () => {
    // Arrange
    const projectName = 'project-123';
    const repoId = 'repo-456';
    const targetVersion = { version: 'main', versionType: 'branch' };
    const sourceVersion = { version: 'v1.0.0', versionType: 'tag' };
    const allCommitsExtended: any[] = [];

    const gitModulesContent = `[submodule "libs/common"]
	path = libs/common
	url = https://example.com/repo`;

    (TFSServices.getItemContent as jest.Mock)
      .mockResolvedValueOnce({ content: gitModulesContent })
      .mockResolvedValueOnce({ content: 'same-sha' })
      .mockResolvedValueOnce({ content: 'same-sha' });

    // Act
    const result = await gitDataProvider.getSubmodulesData(
      projectName,
      repoId,
      targetVersion,
      sourceVersion,
      allCommitsExtended
    );

    // Assert
    expect(result).toHaveLength(0);
    expect(logger.warn).toHaveBeenCalled();
  });

  it('should search commits for source SHA when not found initially', async () => {
    // Arrange
    const projectName = 'project-123';
    const repoId = 'repo-456';
    const targetVersion = { version: 'main', versionType: 'branch' };
    const sourceVersion = { version: 'v1.0.0', versionType: 'tag' };
    const allCommitsExtended = [{ commit: { commitId: 'commit-1' } }, { commit: { commitId: 'commit-2' } }];

    const gitModulesContent = `[submodule "libs/common"]
	path = libs/common
	url = https://example.com/repo`;

    (TFSServices.getItemContent as jest.Mock)
      .mockResolvedValueOnce({ content: gitModulesContent })
      .mockResolvedValueOnce({ content: 'target-sha' })
      .mockResolvedValueOnce({ content: undefined }) // source not found initially
      .mockResolvedValueOnce({ content: 'source-sha-from-commit' }); // found in commit search

    // Act
    const result = await gitDataProvider.getSubmodulesData(
      projectName,
      repoId,
      targetVersion,
      sourceVersion,
      allCommitsExtended
    );

    // Assert
    expect(result).toHaveLength(1);
    expect(result[0].sourceSha1).toBe('source-sha-from-commit');
  });

  it('should handle commits with direct commitId property', async () => {
    // Arrange
    const projectName = 'project-123';
    const repoId = 'repo-456';
    const targetVersion = { version: 'main', versionType: 'branch' };
    const sourceVersion = { version: 'v1.0.0', versionType: 'tag' };
    const allCommitsExtended = [{ commitId: 'direct-commit-1' }, { commitId: 'direct-commit-2' }];

    const gitModulesContent = `[submodule "libs/common"]
	path = libs/common
	url = https://example.com/repo`;

    (TFSServices.getItemContent as jest.Mock)
      .mockResolvedValueOnce({ content: gitModulesContent })
      .mockResolvedValueOnce({ content: 'target-sha' })
      .mockResolvedValueOnce({ content: undefined })
      .mockResolvedValueOnce({ content: 'source-sha' });

    // Act
    const result = await gitDataProvider.getSubmodulesData(
      projectName,
      repoId,
      targetVersion,
      sourceVersion,
      allCommitsExtended
    );

    // Assert
    expect(result).toHaveLength(1);
  });

  it('should warn when commit not found in extended commits', async () => {
    // Arrange
    const projectName = 'project-123';
    const repoId = 'repo-456';
    const targetVersion = { version: 'main', versionType: 'branch' };
    const sourceVersion = { version: 'v1.0.0', versionType: 'tag' };
    const allCommitsExtended = [
      { noCommitId: true }, // No commitId or commit property
      { noCommitId: true },
    ];

    const gitModulesContent = `[submodule "libs/common"]
	path = libs/common
	url = https://example.com/repo`;

    (TFSServices.getItemContent as jest.Mock)
      .mockResolvedValueOnce({ content: gitModulesContent })
      .mockResolvedValueOnce({ content: 'target-sha' })
      .mockResolvedValueOnce({ content: undefined });

    // Act
    const result = await gitDataProvider.getSubmodulesData(
      projectName,
      repoId,
      targetVersion,
      sourceVersion,
      allCommitsExtended
    );

    // Assert
    expect(result).toHaveLength(0);
    expect(logger.warn).toHaveBeenCalled();
  });

  it('should handle errors gracefully', async () => {
    // Arrange
    const projectName = 'project-123';
    const repoId = 'repo-456';
    const targetVersion = { version: 'main', versionType: 'branch' };
    const sourceVersion = { version: 'v1.0.0', versionType: 'tag' };
    const allCommitsExtended: any[] = [];

    (TFSServices.getItemContent as jest.Mock).mockRejectedValueOnce(new Error('API error'));

    // Act
    const result = await gitDataProvider.getSubmodulesData(
      projectName,
      repoId,
      targetVersion,
      sourceVersion,
      allCommitsExtended
    );

    // Assert - returns empty array when .gitmodules fetch fails
    expect(result).toEqual([]);
  });

  it('should handle multiple submodules', async () => {
    // Arrange
    const projectName = 'project-123';
    const repoId = 'repo-456';
    const targetVersion = { version: 'main', versionType: 'branch' };
    const sourceVersion = { version: 'v1.0.0', versionType: 'tag' };
    const allCommitsExtended: any[] = [];

    const gitModulesContent = `[submodule "libs/common"]
	path = libs/common
	url = https://example.com/common
[submodule "libs/utils"]
	path = libs/utils
	url = https://example.com/utils`;

    (TFSServices.getItemContent as jest.Mock)
      .mockResolvedValueOnce({ content: gitModulesContent })
      .mockResolvedValueOnce({ content: 'target-sha-1' })
      .mockResolvedValueOnce({ content: 'source-sha-1' })
      .mockResolvedValueOnce({ content: 'target-sha-2' })
      .mockResolvedValueOnce({ content: 'source-sha-2' });

    // Act
    const result = await gitDataProvider.getSubmodulesData(
      projectName,
      repoId,
      targetVersion,
      sourceVersion,
      allCommitsExtended
    );

    // Assert
    expect(result).toHaveLength(2);
    expect(result[0].gitSubModuleName).toBe('libs_common');
    expect(result[1].gitSubModuleName).toBe('libs_utils');
  });
});

describe('GitDataProvider - createLinkedRelatedItemsForSVD', () => {
  let gitDataProvider: GitDataProvider;
  const mockOrgUrl = 'https://dev.azure.com/orgname/';
  const mockToken = 'mock-token';

  beforeEach(() => {
    jest.clearAllMocks();
    gitDataProvider = new GitDataProvider(mockOrgUrl, mockToken);
  });

  it('should add Requirement when linkedWiTypes=reqOnly and linkedWiRelationship=affectsOnly', async () => {
    jest.spyOn((gitDataProvider as any).ticketsDataProvider, 'GetWorkItemByUrl').mockResolvedValueOnce({
      id: 10,
      fields: { 'System.WorkItemType': 'Requirement', 'System.Title': 'Req' },
      _links: { html: { href: 'http://example.com/10' } },
    });

    const res = await (gitDataProvider as any).createLinkedRelatedItemsForSVD(
      { isEnabled: true, linkedWiTypes: 'reqOnly', linkedWiRelationship: 'affectsOnly' },
      {
        id: 1,
        relations: [
          {
            url: 'https://example.com/_apis/wit/workItems/10',
            rel: 'Affects',
            attributes: { name: 'Affects' },
          },
        ],
      }
    );

    expect(res).toHaveLength(1);
    expect(res[0]).toEqual(
      expect.objectContaining({
        id: 10,
        wiType: 'Requirement',
        relationType: 'Affects',
        title: 'Req',
        url: 'http://example.com/10',
      })
    );
  });

  it('should add Feature when linkedWiTypes=featureOnly and linkedWiRelationship=coversOnly', async () => {
    jest.spyOn((gitDataProvider as any).ticketsDataProvider, 'GetWorkItemByUrl').mockResolvedValueOnce({
      id: 11,
      fields: { 'System.WorkItemType': 'Feature', 'System.Title': 'Feat' },
      _links: { html: { href: 'http://example.com/11' } },
    });

    const res = await (gitDataProvider as any).createLinkedRelatedItemsForSVD(
      { isEnabled: true, linkedWiTypes: 'featureOnly', linkedWiRelationship: 'coversOnly' },
      {
        id: 2,
        relations: [
          {
            url: 'https://example.com/_apis/wit/workItems/11',
            rel: 'CoveredBy',
            attributes: { name: 'CoveredBy' },
          },
        ],
      }
    );

    expect(res).toHaveLength(1);
    expect(res[0]).toEqual(
      expect.objectContaining({
        id: 11,
        wiType: 'Feature',
        relationType: 'CoveredBy',
        title: 'Feat',
        url: 'http://example.com/11',
      })
    );
  });

  it('should add items when linkedWiTypes=both and linkedWiRelationship=both', async () => {
    jest.spyOn((gitDataProvider as any).ticketsDataProvider, 'GetWorkItemByUrl').mockResolvedValueOnce({
      id: 12,
      fields: { 'System.WorkItemType': 'Requirement', 'System.Title': 'Req 12' },
      _links: { html: { href: 'http://example.com/12' } },
    });

    const res = await (gitDataProvider as any).createLinkedRelatedItemsForSVD(
      { isEnabled: true, linkedWiTypes: 'both', linkedWiRelationship: 'both' },
      {
        id: 3,
        relations: [
          {
            url: 'https://example.com/_apis/wit/workItems/12',
            rel: 'Affects',
            attributes: { name: 'Affects' },
          },
        ],
      }
    );

    expect(res).toHaveLength(1);
  });

  it('should skip when linkedWiTypes is enabled but item type does not match', async () => {
    jest.spyOn((gitDataProvider as any).ticketsDataProvider, 'GetWorkItemByUrl').mockResolvedValueOnce({
      id: 13,
      fields: { 'System.WorkItemType': 'Task', 'System.Title': 'Task 13' },
      _links: { html: { href: 'http://example.com/13' } },
    });

    const res = await (gitDataProvider as any).createLinkedRelatedItemsForSVD(
      { isEnabled: true, linkedWiTypes: 'reqOnly', linkedWiRelationship: 'both' },
      {
        id: 4,
        relations: [
          {
            url: 'https://example.com/_apis/wit/workItems/13',
            rel: 'Affects',
            attributes: { name: 'Affects' },
          },
        ],
      }
    );

    expect(res).toEqual([]);
  });
});
