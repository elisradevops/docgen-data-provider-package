import { TFSServices } from '../../helpers/tfs';
import MangementDataProvider from '../MangementDataProvider';

jest.mock('../../helpers/tfs');

describe('MangementDataProvider', () => {
  let managementDataProvider: MangementDataProvider;
  const mockOrgUrl = 'https://dev.azure.com/organization/';
  const mockToken = 'mock-token';

  beforeEach(() => {
    jest.clearAllMocks();
    managementDataProvider = new MangementDataProvider(mockOrgUrl, mockToken);
  });

  describe('GetCllectionLinkTypes', () => {
    it('should return collection link types when API call succeeds', async () => {
      // Arrange
      const mockResponse = {
        value: [
          { id: 'link-1', name: 'Child' },
          { id: 'link-2', name: 'Related' }
        ]
      };
      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockResponse);

      // Act
      const result = await managementDataProvider.GetCllectionLinkTypes();

      // Assert
      expect(TFSServices.getItemContent).toHaveBeenCalledWith(
        `${mockOrgUrl}_apis/wit/workitemrelationtypes`,
        mockToken,
        'get',
        null,
        null
      );
      expect(result).toEqual(mockResponse);
    });

    it('should throw an error when the API call fails', async () => {
      // Arrange
      const expectedError = new Error('API call failed');
      (TFSServices.getItemContent as jest.Mock).mockRejectedValueOnce(expectedError);

      // Act & Assert
      await expect(managementDataProvider.GetCllectionLinkTypes())
        .rejects.toThrow('API call failed');

      expect(TFSServices.getItemContent).toHaveBeenCalledWith(
        `${mockOrgUrl}_apis/wit/workitemrelationtypes`,
        mockToken,
        'get',
        null,
        null
      );
    });
  });

  describe('GetProjects', () => {
    it('should return projects when API call succeeds', async () => {
      // Arrange
      const mockResponse = {
        count: 2,
        value: [
          { id: 'project-1', name: 'Project One' },
          { id: 'project-2', name: 'Project Two' }
        ]
      };
      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockResponse);

      // Act
      const result = await managementDataProvider.GetProjects();

      // Assert
      expect(TFSServices.getItemContent).toHaveBeenCalledWith(
        `${mockOrgUrl}_apis/projects?$top=1000`,
        mockToken
      );
      expect(result).toEqual(mockResponse);
    });

    it('should throw an error when the API call fails', async () => {
      // Arrange
      const expectedError = new Error('Projects API call failed');
      (TFSServices.getItemContent as jest.Mock).mockRejectedValueOnce(expectedError);

      // Act & Assert
      await expect(managementDataProvider.GetProjects())
        .rejects.toThrow('Projects API call failed');

      expect(TFSServices.getItemContent).toHaveBeenCalledWith(
        `${mockOrgUrl}_apis/projects?$top=1000`,
        mockToken
      );
    });
  });

  describe('GetProjectByName', () => {
    it('should return a project when project exists with given name', async () => {
      // Arrange
      const mockProjects = {
        count: 2,
        value: [
          { id: 'project-1', name: 'Project One' },
          { id: 'project-2', name: 'Project Two' }
        ]
      };
      const expectedProject = { id: 'project-2', name: 'Project Two' };

      // Mock GetProjects to return our mock data
      jest.spyOn(managementDataProvider, 'GetProjects').mockResolvedValueOnce(mockProjects);

      // Act
      const result = await managementDataProvider.GetProjectByName('Project Two');

      // Assert
      expect(managementDataProvider.GetProjects).toHaveBeenCalledTimes(1);
      expect(result).toEqual(expectedProject);
    });

    it('should return empty object when project does not exist with given name', async () => {
      // Arrange
      const mockProjects = {
        count: 2,
        value: [
          { id: 'project-1', name: 'Project One' },
          { id: 'project-2', name: 'Project Two' }
        ]
      };

      // Mock GetProjects to return our mock data
      jest.spyOn(managementDataProvider, 'GetProjects').mockResolvedValueOnce(mockProjects);

      // Act
      const result = await managementDataProvider.GetProjectByName('Non-Existent Project');

      // Assert
      expect(managementDataProvider.GetProjects).toHaveBeenCalledTimes(1);
      expect(result).toEqual({});
    });

    it('should return empty object and log error when GetProjects throws', async () => {
      // Arrange
      const expectedError = new Error('Projects API call failed');

      // Mock GetProjects to throw an error
      jest.spyOn(managementDataProvider, 'GetProjects').mockRejectedValueOnce(expectedError);

      // Mock console.log to capture the error
      const consoleLogSpy = jest.spyOn(console, 'log');

      // Act
      const result = await managementDataProvider.GetProjectByName('Any Project');

      // Assert
      expect(managementDataProvider.GetProjects).toHaveBeenCalledTimes(1);
      expect(consoleLogSpy).toHaveBeenCalledWith(expectedError);
      expect(result).toEqual({});
    });
  });

  describe('GetProjectByID', () => {
    it('should return a project when API call succeeds', async () => {
      // Arrange
      const projectId = 'project-123';
      const mockResponse = { id: projectId, name: 'Project 123' };
      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockResponse);

      // Act
      const result = await managementDataProvider.GetProjectByID(projectId);

      // Assert
      expect(TFSServices.getItemContent).toHaveBeenCalledWith(
        `${mockOrgUrl}_apis/projects/${projectId}`,
        mockToken
      );
      expect(result).toEqual(mockResponse);
    });

    it('should throw an error when the API call fails', async () => {
      // Arrange
      const projectId = 'project-123';
      const expectedError = new Error('Project API call failed');
      (TFSServices.getItemContent as jest.Mock).mockRejectedValueOnce(expectedError);

      // Act & Assert
      await expect(managementDataProvider.GetProjectByID(projectId))
        .rejects.toThrow('Project API call failed');

      expect(TFSServices.getItemContent).toHaveBeenCalledWith(
        `${mockOrgUrl}_apis/projects/${projectId}`,
        mockToken
      );
    });
  });

  describe('GetUserProfile', () => {
    it('should return user profile when API call succeeds', async () => {
      // Arrange
      const mockResponse = {
        id: 'user-123',
        displayName: 'Test User',
        emailAddress: 'test@example.com'
      };
      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockResponse);

      // Act
      const result = await managementDataProvider.GetUserProfile();

      // Assert
      expect(TFSServices.getItemContent).toHaveBeenCalledWith(
        `${mockOrgUrl}_api/_common/GetUserProfile?__v=5`,
        mockToken
      );
      expect(result).toEqual(mockResponse);
    });

    it('should throw an error when the API call fails', async () => {
      // Arrange
      const expectedError = new Error('User profile API call failed');
      (TFSServices.getItemContent as jest.Mock).mockRejectedValueOnce(expectedError);

      // Act & Assert
      await expect(managementDataProvider.GetUserProfile())
        .rejects.toThrow('User profile API call failed');

      expect(TFSServices.getItemContent).toHaveBeenCalledWith(
        `${mockOrgUrl}_api/_common/GetUserProfile?__v=5`,
        mockToken
      );
    });
  });

  it('should initialize with the provided organization URL and token', () => {
    // Arrange
    const customOrgUrl = 'https://dev.azure.com/custom-org/';
    const customToken = 'custom-token';

    // Act
    const provider = new MangementDataProvider(customOrgUrl, customToken);

    // Assert
    expect(provider.orgUrl).toBe(customOrgUrl);
    expect(provider.token).toBe(customToken);
  });
});
describe('MangementDataProvider - Additional Tests', () => {
  let managementDataProvider: MangementDataProvider;
  const mockOrgUrl = 'https://dev.azure.com/organization/';
  const mockToken = 'mock-token';

  beforeEach(() => {
    jest.clearAllMocks();
    managementDataProvider = new MangementDataProvider(mockOrgUrl, mockToken);
  });

  describe('GetProjectByName - Edge Cases', () => {
    it('should return empty object when projects list is empty', async () => {
      // Arrange
      const mockEmptyProjects = {
        count: 0,
        value: []
      };
      jest.spyOn(managementDataProvider, 'GetProjects').mockResolvedValueOnce(mockEmptyProjects);

      // Act
      const result = await managementDataProvider.GetProjectByName('Any Project');

      // Assert
      expect(managementDataProvider.GetProjects).toHaveBeenCalledTimes(1);
      expect(result).toEqual({});
    });

    it('should handle projects response with missing value property', async () => {
      // Arrange
      const mockInvalidProjects = {
        count: 0
        // no value property
      };
      jest.spyOn(managementDataProvider, 'GetProjects').mockResolvedValueOnce(mockInvalidProjects);

      // Mock console.log to capture the error
      const consoleLogSpy = jest.spyOn(console, 'log');

      // Act
      const result = await managementDataProvider.GetProjectByName('Any Project');

      // Assert
      expect(result).toEqual({});
      expect(consoleLogSpy).toHaveBeenCalled();
    });

    it('should be case sensitive when matching project names', async () => {
      // Arrange
      const mockProjects = {
        count: 1,
        value: [
          { id: 'project-1', name: 'Project One' }
        ]
      };
      jest.spyOn(managementDataProvider, 'GetProjects').mockResolvedValueOnce(mockProjects);

      // Act
      const result = await managementDataProvider.GetProjectByName('project one'); // lowercase vs Project One

      // Assert
      expect(result).toEqual({});
    });
  });

  describe('GetProjectByID - Edge Cases', () => {
    it('should handle project IDs with special characters', async () => {
      // Arrange
      const projectId = 'project/with-special_chars%123';
      const mockResponse = { id: projectId, name: 'Special Project' };
      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockResponse);

      // Act
      const result = await managementDataProvider.GetProjectByID(projectId);

      // Assert
      expect(TFSServices.getItemContent).toHaveBeenCalledWith(
        `${mockOrgUrl}_apis/projects/${projectId}`,
        mockToken
      );
      expect(result).toEqual(mockResponse);
    });
  });

  describe('Url Construction', () => {
    it('should handle orgUrl with trailing slash properly', async () => {
      // Arrange
      const orgUrlWithTrailingSlash = 'https://dev.azure.com/organization/';
      const provider = new MangementDataProvider(orgUrlWithTrailingSlash, mockToken);
      const mockResponse = { value: [] };
      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockResponse);

      // Act
      await provider.GetCllectionLinkTypes();

      // Assert
      expect(TFSServices.getItemContent).toHaveBeenCalledWith(
        `${orgUrlWithTrailingSlash}_apis/wit/workitemrelationtypes`,
        mockToken,
        'get',
        null,
        null
      );
    });

    it('should handle orgUrl without trailing slash properly', async () => {
      // Arrange
      const orgUrlWithoutTrailingSlash = 'https://dev.azure.com/organization';
      const provider = new MangementDataProvider(orgUrlWithoutTrailingSlash, mockToken);
      const mockResponse = { value: [] };
      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockResponse);

      // Act
      await provider.GetCllectionLinkTypes();

      // Assert
      expect(TFSServices.getItemContent).toHaveBeenCalledWith(
        `${orgUrlWithoutTrailingSlash}_apis/wit/workitemrelationtypes`,
        mockToken,
        'get',
        null,
        null
      );
    });
  });

  describe('GetUserProfile - Edge Cases', () => {
    it('should handle unusual profile data structure', async () => {
      // Arrange
      const unusualProfileData = {
        // Missing typical fields like displayName
        id: 'user-123',
        // Contains unexpected fields
        unusualField: 'unusual value',
        nestedData: {
          someProperty: 'some value'
        }
      };
      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(unusualProfileData);

      // Act
      const result = await managementDataProvider.GetUserProfile();

      // Assert
      expect(result).toEqual(unusualProfileData);
    });
  });
});