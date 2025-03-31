import { TFSServices } from '../../helpers/tfs';
import logger from '../../utils/logger';
import JfrogDataProvider from '../JfrogDataProvider';

jest.mock('../../helpers/tfs');
jest.mock('../../utils/logger');

describe('JfrogDataProvider', () => {
    let jfrogDataProvider: JfrogDataProvider;
    const mockOrgUrl = 'https://dev.azure.com/organization/';
    const mockTfsToken = 'mock-tfs-token';
    const mockJfrogToken = 'mock-jfrog-token';

    beforeEach(() => {
        jest.clearAllMocks();
        jfrogDataProvider = new JfrogDataProvider(mockOrgUrl, mockTfsToken, mockJfrogToken);
    });

    describe('getServiceConnectionUrlByConnectionId', () => {
        it('should fetch service connection URL with correct parameters', async () => {
            // Arrange
            const mockTeamProject = 'test-project';
            const mockConnectionId = 'connection-123';
            const mockResponse = { url: 'https://jfrog.example.com' };

            (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockResponse);

            // Act
            const result = await jfrogDataProvider.getServiceConnectionUrlByConnectionId(
                mockTeamProject,
                mockConnectionId
            );

            // Assert
            expect(TFSServices.getItemContent).toHaveBeenCalledWith(
                `${mockOrgUrl}${mockTeamProject}/_apis/serviceendpoint/endpoints/${mockConnectionId}?api-version=6`,
                mockTfsToken
            );
            expect(logger.debug).toHaveBeenCalledWith(
                `service connection url "${mockResponse.url}"`
            );
            expect(result).toBe(mockResponse.url);
        });

        it('should propagate errors when API call fails', async () => {
            // Arrange
            const mockTeamProject = 'test-project';
            const mockConnectionId = 'connection-123';
            const mockError = new Error('API error');

            (TFSServices.getItemContent as jest.Mock).mockRejectedValueOnce(mockError);

            // Act & Assert
            await expect(
                jfrogDataProvider.getServiceConnectionUrlByConnectionId(mockTeamProject, mockConnectionId)
            ).rejects.toThrow('API error');
        });
    });

    describe('getCiDataFromJfrog', () => {
        it('should fetch JFrog data with token and format URLs correctly', async () => {
            // Arrange
            const mockJfrogUrl = 'https://jfrog.example.com';
            const mockBuildName = 'build-name';
            const mockBuildVersion = 'version-1.0';
            const mockResponse = {
                buildInfo: {
                    url: 'https://ci.example.com/build/123'
                }
            };

            (TFSServices.getJfrogRequest as jest.Mock).mockResolvedValueOnce(mockResponse);

            // Act
            const result = await jfrogDataProvider.getCiDataFromJfrog(
                mockJfrogUrl,
                mockBuildName,
                mockBuildVersion
            );

            // Assert
            expect(TFSServices.getJfrogRequest).toHaveBeenCalledWith(
                `${mockJfrogUrl}/api/build/${mockBuildName}/${mockBuildVersion}`,
                { 'Authorization': `Bearer ${mockJfrogToken}` }
            );
            expect(logger.info).toHaveBeenCalledWith(
                `Querying Jfrog using url ${mockJfrogUrl}/api/build/${mockBuildName}/${mockBuildVersion}`
            );
            expect(logger.debug).toHaveBeenCalledWith(
                `CI Url from JFROG: ${mockResponse.buildInfo.url}`
            );
            expect(result).toBe(mockResponse.buildInfo.url);
        });

        it('should handle build names and versions that already start with /', async () => {
            // Arrange
            const mockJfrogUrl = 'https://jfrog.example.com';
            const mockBuildName = '/build-name';
            const mockBuildVersion = '/version-1.0';
            const mockResponse = {
                buildInfo: {
                    url: 'https://ci.example.com/build/123'
                }
            };

            (TFSServices.getJfrogRequest as jest.Mock).mockResolvedValueOnce(mockResponse);

            // Act
            const result = await jfrogDataProvider.getCiDataFromJfrog(
                mockJfrogUrl,
                mockBuildName,
                mockBuildVersion
            );

            // Assert
            expect(TFSServices.getJfrogRequest).toHaveBeenCalledWith(
                `${mockJfrogUrl}/api/build${mockBuildName}${mockBuildVersion}`,
                { 'Authorization': `Bearer ${mockJfrogToken}` }
            );
            expect(result).toBe(mockResponse.buildInfo.url);
        });

        it('should make request without token when jfrogToken is empty', async () => {
            // Arrange
            const mockJfrogUrl = 'https://jfrog.example.com';
            const mockBuildName = 'build-name';
            const mockBuildVersion = 'version-1.0';
            const mockResponse = {
                buildInfo: {
                    url: 'https://ci.example.com/build/123'
                }
            };

            // Create provider with empty jfrog token
            const jfrogDataProviderNoToken = new JfrogDataProvider(mockOrgUrl, mockTfsToken, '');
            (TFSServices.getJfrogRequest as jest.Mock).mockResolvedValueOnce(mockResponse);

            // Act
            const result = await jfrogDataProviderNoToken.getCiDataFromJfrog(
                mockJfrogUrl,
                mockBuildName,
                mockBuildVersion
            );

            // Assert
            expect(TFSServices.getJfrogRequest).toHaveBeenCalledWith(
                `${mockJfrogUrl}/api/build/${mockBuildName}/${mockBuildVersion}`
            );
            expect(result).toBe(mockResponse.buildInfo.url);
        });

        it('should log and rethrow errors from JFrog API', async () => {
            // Arrange
            const mockJfrogUrl = 'https://jfrog.example.com';
            const mockBuildName = 'build-name';
            const mockBuildVersion = 'version-1.0';
            const mockError = new Error('JFrog API error');

            (TFSServices.getJfrogRequest as jest.Mock).mockRejectedValueOnce(mockError);

            // Act & Assert
            await expect(
                jfrogDataProvider.getCiDataFromJfrog(mockJfrogUrl, mockBuildName, mockBuildVersion)
            ).rejects.toThrow('JFrog API error');

            expect(logger.error).toHaveBeenCalledWith(
                `Error occurred during querying JFrog using: JFrog API error`
            );
        });
    });
});