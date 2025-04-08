import { TFSServices } from '../../helpers/tfs';
import TicketsDataProvider from '../TicketsDataProvider';
import logger from '../../utils/logger';
import { Helper } from '../../helpers/helper';
import { QueryType } from '../../models/tfs-data';

jest.mock('../../helpers/tfs');
jest.mock('../../utils/logger');
jest.mock('../../helpers/helper');

describe('TicketsDataProvider', () => {
  let ticketsDataProvider: TicketsDataProvider;
  const mockOrgUrl = 'https://dev.azure.com/organization/';
  const mockToken = 'mock-token';
  const mockProject = 'test-project';

  beforeEach(() => {
    jest.clearAllMocks();
    ticketsDataProvider = new TicketsDataProvider(mockOrgUrl, mockToken);
  });

  describe('FetchImageAsBase64', () => {
    it('should fetch image as base64', async () => {
      // Arrange
      const mockUrl = 'https://example.com/image.jpg';
      const mockBase64 = 'base64-encoded-image';
      (TFSServices.fetchAzureDevOpsImageAsBase64 as jest.Mock).mockResolvedValueOnce(mockBase64);

      // Act
      const result = await ticketsDataProvider.FetchImageAsBase64(mockUrl);

      // Assert
      expect(TFSServices.fetchAzureDevOpsImageAsBase64).toHaveBeenCalledWith(
        mockUrl, mockToken, 'get', null
      );
      expect(result).toBe(mockBase64);
    });
  });

  describe('GetWorkItem', () => {
    it('should fetch work item with correct URL', async () => {
      // Arrange
      const mockId = '123';
      const mockWorkItem = { id: 123, fields: { 'System.Title': 'Test Work Item' } };
      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockWorkItem);

      // Act
      const result = await ticketsDataProvider.GetWorkItem(mockProject, mockId);

      // Assert
      expect(TFSServices.getItemContent).toHaveBeenCalledWith(
        `${mockOrgUrl}${mockProject}/_apis/wit/workitems/${mockId}?$expand=All`,
        mockToken
      );
      expect(result).toEqual(mockWorkItem);
    });
  });

  describe('GetLinksByIds', () => {
    it('should retrieve links by ids', async () => {
      // Arrange
      const mockIds = [1, 2];
      const mockWorkItems = [
        { id: 1, fields: { 'System.Title': 'Item 1' } },
        { id: 2, fields: { 'System.Title': 'Item 2' } }
      ];
      const mockLinksMap = new Map();
      mockLinksMap.set('1', { id: '1', rels: ['3'] });
      mockLinksMap.set('2', { id: '2', rels: [] });

      const mockRelatedItems = [{ id: 3, fields: { 'System.Title': 'Related Item' } }];
      const mockTraceItem = { id: '1', title: 'Item 1', url: 'url', customerId: 'customer', links: [] };

      jest.spyOn(ticketsDataProvider, 'PopulateWorkItemsByIds').mockResolvedValueOnce(mockWorkItems);
      jest.spyOn(ticketsDataProvider, 'GetRelationsIds').mockResolvedValueOnce(mockLinksMap);
      jest.spyOn(ticketsDataProvider, 'GetParentLink')
        .mockResolvedValueOnce(mockTraceItem)
        .mockResolvedValueOnce({ id: '2', title: 'Item 2', url: 'url', customerId: 'customer', links: [] });
      jest.spyOn(ticketsDataProvider, 'PopulateWorkItemsByIds').mockResolvedValueOnce(mockRelatedItems);
      jest.spyOn(ticketsDataProvider, 'GetLinks').mockResolvedValueOnce([]);

      // Act
      const result = await ticketsDataProvider.GetLinksByIds(mockProject, mockIds);

      // Assert
      expect(result.length).toBe(2);
      expect(ticketsDataProvider.PopulateWorkItemsByIds).toHaveBeenCalledWith(mockIds, mockProject);
      expect(ticketsDataProvider.GetRelationsIds).toHaveBeenCalledWith(mockWorkItems);
    });
  });

  describe('GetSharedQueries', () => {
    it('should fetch STD shared queries with correct URL', async () => {
      // Arrange
      const mockPath = '';
      const mockDocType = 'STD';
      const mockQueries = { name: 'Query 1' };
      const mockResponse = { reqTestTree: {}, testReqTree: {} };

      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockQueries);
      jest.spyOn(ticketsDataProvider as any, 'fetchLinkedQueries').mockResolvedValueOnce(mockResponse);

      // Act
      const result = await ticketsDataProvider.GetSharedQueries(mockProject, mockPath, mockDocType);

      // Assert
      expect(TFSServices.getItemContent).toHaveBeenCalledWith(
        `${mockOrgUrl}${mockProject}/_apis/wit/queries/Shared%20Queries?$depth=2&$expand=all`,
        mockToken
      );
      expect(result).toEqual(mockResponse);
    });

    it('should fetch SVD shared queries and call fetchAnyQueries', async () => {
      // Arrange
      const mockPath = 'Custom Path';
      const mockDocType = 'SVD';
      const mockQueries = { name: 'Query 1' };
      const mockResponse = { systemOverviewQueryTree: {}, knownBugsQueryTree: {} };

      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockQueries);
      jest.spyOn(ticketsDataProvider as any, 'fetchAnyQueries').mockResolvedValueOnce(mockResponse);

      // Act
      const result = await ticketsDataProvider.GetSharedQueries(mockProject, mockPath, mockDocType);

      // Assert
      expect(TFSServices.getItemContent).toHaveBeenCalledWith(
        `${mockOrgUrl}${mockProject}/_apis/wit/queries/${mockPath}?$depth=2&$expand=all`,
        mockToken
      );
      expect(result).toEqual(mockResponse);
    });

    it('should handle errors', async () => {
      // Arrange
      const mockPath = '';
      const mockError = new Error('API error');

      (TFSServices.getItemContent as jest.Mock).mockImplementationOnce(() => {
        return Promise.reject(mockError);
      });

      // Act & Assert
      await expect(ticketsDataProvider.GetSharedQueries(mockProject, mockPath))
        .rejects.toThrow('API error');
      expect(logger.error).toHaveBeenCalled();
    });
  });

  describe('GetQueryResultsFromWiql', () => {
    it('should handle OneHop query with table format', async () => {
      // Arrange
      const mockWiqlHref = 'https://example.com/wiql';
      const mockTestCaseMap = new Map<number, Set<any>>();
      const mockQueryResult = {
        queryType: QueryType.OneHop,
        columns: [],
        workItemRelations: [{ source: null, target: { id: 1, url: 'url' } }]
      };
      const mockTableResult = {
        sourceTargetsMap: new Map(),
        sortingSourceColumnsMap: new Map(),
        sortingTargetsColumnsMap: new Map()
      };

      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockQueryResult);
      jest.spyOn(ticketsDataProvider as any, 'parseDirectLinkedQueryResultForTableFormat')
        .mockResolvedValueOnce(mockTableResult);

      // Act
      const result = await ticketsDataProvider.GetQueryResultsFromWiql(
        mockWiqlHref, true, mockTestCaseMap
      );

      // Assert
      expect(TFSServices.getItemContent).toHaveBeenCalledWith(mockWiqlHref, mockToken);
      expect(result).toEqual(mockTableResult);
    });

    it('should throw error when wiqlHref is empty', async () => {
      // Arrange
      const mockTestCaseMap = new Map<number, Set<any>>();

      // Act & Assert
      const result = await ticketsDataProvider.GetQueryResultsFromWiql('', false, mockTestCaseMap);
      expect(logger.error).toHaveBeenCalled();
      expect(result).toBeUndefined();
    });
  });

  describe('GetModeledQuery', () => {
    it('should structure query list correctly', () => {
      // Arrange
      const mockQueryList = [
        {
          name: 'Query 1',
          _links: { wiql: 'http://example.com/wiql1' },
          id: 'q1'
        },
        {
          name: 'Query 2',
          _links: { wiql: null },
          id: 'q2'
        }
      ];

      // Act
      const result = ticketsDataProvider.GetModeledQuery(mockQueryList);

      // Assert
      expect(result).toEqual([
        { queryName: 'Query 1', wiql: 'http://example.com/wiql1', id: 'q1' },
        { queryName: 'Query 2', wiql: null, id: 'q2' }
      ]);
    });
  });

  describe('PopulateWorkItemsByIds', () => {
    it('should fetch work items in batches of 200', async () => {
      // Arrange
      const mockIds = Array.from({ length: 250 }, (_, i) => i + 1);
      const mockResponse1 = { value: mockIds.slice(0, 200).map(id => ({ id })) };
      const mockResponse2 = { value: mockIds.slice(200).map(id => ({ id })) };

      (TFSServices.getItemContent as jest.Mock)
        .mockResolvedValueOnce(mockResponse1)
        .mockResolvedValueOnce(mockResponse2);

      // Act
      const result = await ticketsDataProvider.PopulateWorkItemsByIds(mockIds, mockProject);

      // Assert
      expect(TFSServices.getItemContent).toHaveBeenCalledTimes(2);
      expect(result.length).toBe(250);
    });

    it('should handle errors and return empty array', async () => {
      // Arrange
      const mockIds = [1, 2, 3];
      const mockError = new Error('API error');

      (TFSServices.getItemContent as jest.Mock).mockRejectedValueOnce(mockError);

      // Act
      const result = await ticketsDataProvider.PopulateWorkItemsByIds(mockIds, mockProject);

      // Assert
      expect(result).toEqual([]);
      expect(logger.error).toHaveBeenCalled();
    });
  });

  describe('GetIterationsByTeamName', () => {
    it('should fetch iterations with team name specified', async () => {
      // Arrange
      const mockTeamName = 'test-team';
      const mockIterations = ['iteration1', 'iteration2'];

      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockIterations);

      // Act
      const result = await ticketsDataProvider.GetIterationsByTeamName(mockProject, mockTeamName);

      // Assert
      expect(TFSServices.getItemContent).toHaveBeenCalledWith(
        `${mockOrgUrl}${mockProject}/${mockTeamName}/_apis/work/teamsettings/iterations`,
        mockToken,
        'get'
      );
      expect(result).toEqual(mockIterations);
    });

    it('should fetch iterations without team name', async () => {
      // Arrange
      const mockTeamName = '';
      const mockIterations = ['iteration1', 'iteration2'];

      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockIterations);

      // Act
      const result = await ticketsDataProvider.GetIterationsByTeamName(mockProject, mockTeamName);

      // Assert
      expect(TFSServices.getItemContent).toHaveBeenCalledWith(
        `${mockOrgUrl}${mockProject}/_apis/work/teamsettings/iterations`,
        mockToken,
        'get'
      );
      expect(result).toEqual(mockIterations);
    });
  });

  describe('CreateNewWorkItem', () => {
    it('should create work item with correct parameters', async () => {
      // Arrange
      const mockWiBody = [{ op: 'add', path: '/fields/System.Title', value: 'New Item' }];
      const mockWiType = 'Bug';
      const mockByPass = true;
      const mockResponse = { id: 123, fields: { 'System.Title': 'New Item' } };

      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockResponse);

      // Act
      const result = await ticketsDataProvider.CreateNewWorkItem(
        mockProject, mockWiBody, mockWiType, mockByPass
      );

      // Assert
      expect(TFSServices.getItemContent).toHaveBeenCalledWith(
        `${mockOrgUrl}${mockProject}/_apis/wit/workitems/$${mockWiType}?bypassRules=true`,
        mockToken,
        'POST',
        mockWiBody,
        {
          'Content-Type': 'application/json-patch+json'
        }
      );
      expect(result).toEqual(mockResponse);
    });
  });

  describe('UpdateWorkItem', () => {
    it('should update work item with correct parameters', async () => {
      // Arrange
      const mockWiBody = [{ op: 'add', path: '/fields/System.Title', value: 'Updated Item' }];
      const mockWorkItemId = 123;
      const mockByPass = true;
      const mockResponse = { id: 123, fields: { 'System.Title': 'Updated Item' } };

      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockResponse);

      // Act
      const result = await ticketsDataProvider.UpdateWorkItem(
        mockProject, mockWiBody, mockWorkItemId, mockByPass
      );

      // Assert
      expect(TFSServices.getItemContent).toHaveBeenCalledWith(
        `${mockOrgUrl}${mockProject}/_apis/wit/workitems/${mockWorkItemId}?bypassRules=true`,
        mockToken,
        'patch',
        mockWiBody,
        {
          'Content-Type': 'application/json-patch+json'
        }
      );
      expect(result).toEqual(mockResponse);
    });
  });

  describe('GetWorkitemAttachments', () => {
    it('should return attachments for work item', async () => {
      // Arrange
      const mockId = '123';
      const mockWorkItem = {
        relations: [
          {
            rel: 'AttachedFile',
            url: 'https://example.com/attachment/1',
            attributes: { name: 'file.txt' }
          }
        ]
      };

      // Mock the TFSServices.getItemContent directly to return our mock data
      // This is what will be called by the new TicketsDataProvider instance
      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockWorkItem);

      // Act
      const result = await ticketsDataProvider.GetWorkitemAttachments(mockProject, mockId);

      // Assert
      expect(TFSServices.getItemContent).toHaveBeenCalledWith(
        `${mockOrgUrl}${mockProject}/_apis/wit/workitems/${mockId}?$expand=All`,
        mockToken
      );
      expect(result.length).toBe(1);

      // Check that the downloadUrl was added correctly
      expect(result[0]).toHaveProperty('downloadUrl', 'https://example.com/attachment/1/file.txt');
      expect(result[0].rel).toBe('AttachedFile');
    });

    it('should handle work item with no relations', async () => {
      // Arrange
      const mockId = '123';
      const mockWorkItem = { relations: null };

      jest.spyOn(ticketsDataProvider, 'GetWorkItem').mockResolvedValueOnce(mockWorkItem);

      // Act
      const result = await ticketsDataProvider.GetWorkitemAttachments(mockProject, mockId);

      // Assert
      expect(result).toEqual([]);
    });

    it('should filter out non-attachment relations', async () => {
      // Arrange
      const mockId = '123';
      const mockWorkItem = {
        relations: [
          {
            rel: 'Parent',
            url: 'https://example.com/parent/1',
            attributes: { name: 'parent' }
          },
          {
            rel: 'AttachedFile',
            url: 'https://example.com/attachment/1',
            attributes: { name: 'file.txt' }
          }
        ]
      };

      // Mock TFSServices.getItemContent directly instead of GetWorkItem
      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockWorkItem);

      // Act
      const result = await ticketsDataProvider.GetWorkitemAttachments(mockProject, mockId);

      // Assert
      expect(TFSServices.getItemContent).toHaveBeenCalledWith(
        `${mockOrgUrl}${mockProject}/_apis/wit/workitems/${mockId}?$expand=All`,
        mockToken
      );
      expect(result.length).toBe(1);
      expect(result[0].rel).toBe('AttachedFile');
      // Only the AttachedFile relation should be in the result
      expect(result.some(item => item.rel === 'Parent')).toBe(false);
    });
  });
});