import { TFSServices } from '../../helpers/tfs';
import TicketsDataProvider from '../../modules/TicketsDataProvider';
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

  describe('isLinkSideAllowedByTypeOrId', () => {
    it('should return true when source types include allowed types', async () => {
      const wiql = "SELECT * FROM WorkItemLinks WHERE Source.[System.WorkItemType] IN ('Bug')";
      const result = await (ticketsDataProvider as any).isLinkSideAllowedByTypeOrId(
        {},
        wiql,
        'Source',
        [],
        new Map()
      );
      expect(result).toBe(true);
    });

    it('should return false when allowedTypes is empty but field is not present', async () => {
      const wiql = 'SELECT * FROM WorkItems';
      const result = await (ticketsDataProvider as any).isLinkSideAllowedByTypeOrId(
        {},
        wiql,
        'Source',
        [],
        new Map()
      );
      expect(result).toBe(false);
    });

    it('should support equality operator and reject disallowed types', async () => {
      const wiql = "SELECT * FROM WorkItemLinks WHERE Source.[System.WorkItemType] = 'Bug'";
      const result = await (ticketsDataProvider as any).isLinkSideAllowedByTypeOrId(
        {},
        wiql,
        'Source',
        ['Epic', 'Feature'],
        new Map()
      );
      expect(result).toBe(false);
    });

    it('should support IN operator when all types are allowed', async () => {
      const wiql = "SELECT * FROM WorkItemLinks WHERE Target.[System.WorkItemType] IN ('Requirement', 'Bug')";
      const result = await (ticketsDataProvider as any).isLinkSideAllowedByTypeOrId(
        {},
        wiql,
        'Target',
        ['Requirement', 'Bug'],
        new Map()
      );
      expect(result).toBe(true);
    });

    it('should return false when no types are found in WIQL and no ids exist', async () => {
      const wiql = "SELECT * FROM WorkItemLinks WHERE Target.[System.AreaPath] = 'A'";
      const result = await (ticketsDataProvider as any).isLinkSideAllowedByTypeOrId(
        {},
        wiql,
        'Target',
        ['Bug'],
        new Map()
      );
      expect(result).toBe(false);
    });
  });

  describe('fetchSystemRequirementQueries', () => {
    it('should include Task in allowed types', async () => {
      const structureSpy = jest
        .spyOn(ticketsDataProvider as any, 'structureFetchedQueries')
        .mockResolvedValue({ tree1: { id: 't1' }, tree2: null });

      await (ticketsDataProvider as any).fetchSystemRequirementQueries({ hasChildren: false }, []);

      expect(structureSpy.mock.calls[0][3]).toEqual(['epic', 'feature', 'requirement', 'task']);
    });
  });

  describe('findChildFolderByPossibleNames', () => {
    it('should return null when parent is missing or possibleNames is empty', async () => {
      await expect(
        (ticketsDataProvider as any).findChildFolderByPossibleNames(null, ['a'])
      ).resolves.toBeNull();
      await expect(
        (ticketsDataProvider as any).findChildFolderByPossibleNames({ hasChildren: false }, [])
      ).resolves.toBeNull();
    });

    it('should prefer exact match over partial match and return the enriched folder', async () => {
      const parent: any = {
        id: 'p',
        isFolder: true,
        name: 'Parent',
        hasChildren: true,
        children: [
          { id: 'c1', isFolder: true, name: 'Requirement - Test', hasChildren: false },
          { id: 'c2', isFolder: true, name: 'requirement to test case', hasChildren: false },
        ],
      };

      const ensureSpy = jest
        .spyOn(ticketsDataProvider as any, 'ensureQueryChildren')
        .mockImplementation(async (node: any) => node);

      const result = await (ticketsDataProvider as any).findChildFolderByPossibleNames(parent, [
        'requirement - test',
        'req',
      ]);
      expect(result).toBe(parent.children[0]);
      expect(ensureSpy).toHaveBeenCalled();
    });

    it('should return first partial match when no exact match exists', async () => {
      const parent: any = {
        id: 'p',
        isFolder: true,
        name: 'Parent',
        hasChildren: true,
        children: [
          { id: 'c1', isFolder: true, name: 'Some Req Folder', hasChildren: false },
          { id: 'c2', isFolder: true, name: 'Other', hasChildren: false },
        ],
      };

      jest
        .spyOn(ticketsDataProvider as any, 'ensureQueryChildren')
        .mockImplementation(async (node: any) => node);

      const result = await (ticketsDataProvider as any).findChildFolderByPossibleNames(parent, ['req']);
      expect(result).toBe(parent.children[0]);
    });

    it('should BFS into nested folders and ignore non-folder children and visited duplicates', async () => {
      const leaf: any = { id: 'leaf', isFolder: true, name: 'Deep Match', hasChildren: false, children: [] };
      const parent: any = {
        id: 'p',
        isFolder: true,
        name: 'Parent',
        hasChildren: true,
        children: [
          { id: 'a', isFolder: true, name: 'A', hasChildren: true },
          { id: 'b', isFolder: true, name: 'B', hasChildren: true },
        ],
      };

      jest.spyOn(ticketsDataProvider as any, 'ensureQueryChildren').mockImplementation(async (node: any) => {
        if (node?.id === 'a') {
          return {
            ...node,
            children: [{ id: 'not-folder', isFolder: false, name: 'NF' }, leaf, leaf],
          };
        }
        if (node?.id === 'b') {
          return { ...node, children: [leaf] };
        }
        return node;
      });

      const found = await (ticketsDataProvider as any).findChildFolderByPossibleNames(parent, ['deep match']);
      expect(found).toEqual(leaf);
    });
  });

  describe('fetchWithAncestorFallback', () => {
    it('should short-circuit when validator passes on the first candidate', async () => {
      const root: any = { id: 'root', name: 'Root', hasChildren: false, children: [] };
      const starting: any = { id: 'start', name: 'Start', hasChildren: false, children: [] };

      jest.spyOn(ticketsDataProvider as any, 'ensureQueryChildren').mockImplementation(async (n: any) => n);
      jest.spyOn(ticketsDataProvider as any, 'buildFallbackChain').mockResolvedValueOnce([starting, root]);

      const fetcher = jest.fn().mockResolvedValueOnce({ ok: true });
      const validator = jest.fn().mockReturnValueOnce(true);

      const res = await (ticketsDataProvider as any).fetchWithAncestorFallback(
        root,
        starting,
        fetcher,
        'ctx',
        validator
      );
      expect(res.usedFolder).toBe(starting);
      expect(fetcher).toHaveBeenCalledTimes(1);
      expect(validator).toHaveBeenCalledWith({ ok: true });
    });

    it('should fall back through all candidates and return last attempt when validator never passes', async () => {
      const root: any = { id: 'root', name: 'Root', hasChildren: false, children: [] };
      const a: any = { id: 'a', name: 'A', hasChildren: false, children: [] };
      const b: any = { id: 'b', name: 'B', hasChildren: false, children: [] };

      jest.spyOn(ticketsDataProvider as any, 'ensureQueryChildren').mockImplementation(async (n: any) => n);
      jest.spyOn(ticketsDataProvider as any, 'buildFallbackChain').mockResolvedValueOnce([a, b, root]);

      const fetcher = jest
        .fn()
        .mockResolvedValueOnce(null)
        .mockResolvedValueOnce({})
        .mockResolvedValueOnce({ last: true });
      const validator = jest.fn().mockReturnValue(false);

      const res = await (ticketsDataProvider as any).fetchWithAncestorFallback(
        root,
        a,
        fetcher,
        'ctx',
        validator
      );
      expect(res.usedFolder).toBe(root);
      expect(res.result).toEqual({ last: true });
      expect(fetcher).toHaveBeenCalledTimes(3);
    });
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
      expect(TFSServices.fetchAzureDevOpsImageAsBase64).toHaveBeenCalledWith(mockUrl, mockToken, 'get', null);
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
        { id: 2, fields: { 'System.Title': 'Item 2' } },
      ];
      const mockLinksMap = new Map();
      mockLinksMap.set('1', { id: '1', rels: ['3'] });
      mockLinksMap.set('2', { id: '2', rels: [] });

      const mockRelatedItems = [{ id: 3, fields: { 'System.Title': 'Related Item' } }];
      const mockTraceItem = { id: '1', title: 'Item 1', url: 'url', customerId: 'customer', links: [] };

      jest.spyOn(ticketsDataProvider, 'PopulateWorkItemsByIds').mockResolvedValueOnce(mockWorkItems);
      jest.spyOn(ticketsDataProvider, 'GetRelationsIds').mockResolvedValueOnce(mockLinksMap);
      jest
        .spyOn(ticketsDataProvider, 'GetParentLink')
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
      const mockQueries = { name: 'Query 1' } as any;
      const mockBranchesResponse = {
        reqToTest: { result: { reqTestTree: {} }, usedFolder: {} },
        testToReq: { result: { testReqTree: {} }, usedFolder: {} },
        mom: { result: { linkedMomTree: {} }, usedFolder: {} },
      } as any;

      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockQueries);
      jest.spyOn(ticketsDataProvider as any, 'ensureQueryChildren').mockResolvedValueOnce(mockQueries);
      jest
        .spyOn(ticketsDataProvider as any, 'getDocTypeRoot')
        .mockResolvedValueOnce({ root: mockQueries, found: true });
      jest
        .spyOn(ticketsDataProvider as any, 'fetchDocTypeBranches')
        .mockResolvedValueOnce(mockBranchesResponse);

      // Act
      const result = await ticketsDataProvider.GetSharedQueries(mockProject, mockPath, mockDocType);

      // Assert
      expect(TFSServices.getItemContent).toHaveBeenCalledWith(
        `${mockOrgUrl}${mockProject}/_apis/wit/queries/Shared%20Queries?$depth=2&$expand=all`,
        mockToken
      );
      expect(result).toEqual({
        reqTestQueries: { reqTestTree: {}, testReqTree: {} },
        linkedMomQueries: { linkedMomTree: {} },
      });
    });

    it('should fetch SVD shared queries and call fetchAnyQueries', async () => {
      // Arrange
      const mockPath = 'Custom Path';
      const mockDocType = 'SVD';
      const mockQueries = { name: 'Query 1' } as any;
      const mockBranchesResponse = {
        systemOverview: { result: {}, usedFolder: {} },
        knownBugs: { result: {}, usedFolder: {} },
      } as any;

      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockQueries);
      jest.spyOn(ticketsDataProvider as any, 'ensureQueryChildren').mockResolvedValueOnce(mockQueries);
      jest
        .spyOn(ticketsDataProvider as any, 'getDocTypeRoot')
        .mockResolvedValueOnce({ root: mockQueries, found: true });
      jest
        .spyOn(ticketsDataProvider as any, 'fetchDocTypeBranches')
        .mockResolvedValueOnce(mockBranchesResponse);

      // Act
      const result = await ticketsDataProvider.GetSharedQueries(mockProject, mockPath, mockDocType);

      // Assert
      expect(TFSServices.getItemContent).toHaveBeenCalledWith(
        `${mockOrgUrl}${mockProject}/_apis/wit/queries/${mockPath}?$depth=2&$expand=all`,
        mockToken
      );
      expect(result).toEqual({
        systemOverviewQueryTree: {},
        knownBugsQueryTree: {},
      });
    });

    it('should handle errors', async () => {
      // Arrange
      const mockPath = '';
      const mockError = new Error('API error');

      (TFSServices.getItemContent as jest.Mock).mockImplementationOnce(() => {
        return Promise.reject(mockError);
      });

      // Act & Assert
      await expect(ticketsDataProvider.GetSharedQueries(mockProject, mockPath)).rejects.toThrow('API error');
      expect(logger.error).toHaveBeenCalled();
    });

    it('should fall back to reqToTestResult.testReqTree when testToReqResult is missing (std)', async () => {
      const mockQueries = { name: 'QueriesRoot' } as any;
      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockQueries);
      jest.spyOn(ticketsDataProvider as any, 'ensureQueryChildren').mockResolvedValueOnce(mockQueries);
      jest
        .spyOn(ticketsDataProvider as any, 'getDocTypeRoot')
        .mockResolvedValueOnce({ root: mockQueries, found: true });

      jest.spyOn(ticketsDataProvider as any, 'fetchDocTypeBranches').mockResolvedValueOnce({
        reqToTest: {
          result: {
            reqTestTree: { a: 1 },
            testReqTree: { fallback: true },
          },
          usedFolder: { name: 'req-to-test' },
        },
        // missing testToReq
        mom: { result: { linkedMomTree: null }, usedFolder: { name: 'mom' } },
      });

      const res = await ticketsDataProvider.GetSharedQueries(mockProject, '', 'std');
      expect(res).toEqual({
        reqTestQueries: { reqTestTree: { a: 1 }, testReqTree: { fallback: true } },
        linkedMomQueries: { linkedMomTree: null },
      });
    });

    it('should fall back to openPcrToTest.TestToOpenPcrTree when testToOpenPcr branch is missing (str)', async () => {
      const mockQueries = { name: 'QueriesRoot' } as any;
      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockQueries);
      jest.spyOn(ticketsDataProvider as any, 'ensureQueryChildren').mockResolvedValueOnce(mockQueries);
      jest
        .spyOn(ticketsDataProvider as any, 'getDocTypeRoot')
        .mockResolvedValueOnce({ root: mockQueries, found: true });

      jest.spyOn(ticketsDataProvider as any, 'fetchDocTypeBranches').mockResolvedValueOnce({
        reqToTest: { result: { reqTestTree: null }, usedFolder: { name: 'req-to-test' } },
        testToReq: { result: { testReqTree: null }, usedFolder: { name: 'test-to-req' } },
        openPcrToTest: {
          result: {
            OpenPcrToTestTree: { ok: true },
            TestToOpenPcrTree: { fromOpenPcr: true },
          },
          usedFolder: { name: 'open-pcr-to-test' },
        },
        // missing testToOpenPcr
      });

      const res = await ticketsDataProvider.GetSharedQueries(mockProject, '', 'str');
      expect(res).toEqual({
        reqTestTrees: { reqTestTree: null, testReqTree: null },
        openPcrTestTrees: {
          OpenPcrToTestTree: { ok: true },
          TestToOpenPcrTree: { fromOpenPcr: true },
        },
      });
    });

    it('should execute std validators via fetchDocTypeBranches and accept fallback when first folder yields no results', async () => {
      const queries: any = {
        id: 'root',
        name: 'Root',
        isFolder: true,
        hasChildren: true,
        children: [
          { id: 'req', name: 'Requirement - Test', isFolder: true, hasChildren: false },
          { id: 'test', name: 'Test - Requirement', isFolder: true, hasChildren: false },
          { id: 'mom', name: 'MOM', isFolder: true, hasChildren: false },
        ],
      };

      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(queries);
      jest.spyOn(ticketsDataProvider as any, 'ensureQueryChildren').mockImplementation(async (x: any) => x);
      jest
        .spyOn(ticketsDataProvider as any, 'getDocTypeRoot')
        .mockResolvedValueOnce({ root: queries, found: true });

      // Force two candidates so validators run with null and then with data
      jest
        .spyOn(ticketsDataProvider as any, 'buildFallbackChain')
        .mockResolvedValueOnce([queries.children[0], queries])
        .mockResolvedValueOnce([queries.children[1], queries])
        .mockResolvedValueOnce([queries.children[2], queries]);

      const fetchReqTestSpy = jest
        .spyOn(ticketsDataProvider as any, 'fetchLinkedReqTestQueries')
        .mockResolvedValueOnce(null)
        .mockResolvedValueOnce({ reqTestTree: { isValidQuery: true } })
        .mockResolvedValueOnce(null)
        .mockResolvedValueOnce({ testReqTree: { isValidQuery: true } });

      const fetchMomSpy = jest
        .spyOn(ticketsDataProvider as any, 'fetchLinkedMomQueries')
        .mockResolvedValueOnce(null)
        .mockResolvedValueOnce({ linkedMomTree: { isValidQuery: true } });

      const res = await ticketsDataProvider.GetSharedQueries(mockProject, '', 'std');

      expect(fetchReqTestSpy).toHaveBeenCalled();
      expect(fetchMomSpy).toHaveBeenCalled();
      expect(res.reqTestQueries.reqTestTree).toEqual({ isValidQuery: true });
      expect(res.reqTestQueries.testReqTree).toEqual({ isValidQuery: true });
      expect(res.linkedMomQueries.linkedMomTree).toEqual({ isValidQuery: true });
    });
  });

  describe('hasAnyQueryTree', () => {
    it('should return true for objects with isValidQuery/wiql/queryType', () => {
      expect((ticketsDataProvider as any).hasAnyQueryTree({ isValidQuery: true })).toBe(true);
      expect((ticketsDataProvider as any).hasAnyQueryTree({ wiql: 'x' })).toBe(true);
      expect((ticketsDataProvider as any).hasAnyQueryTree({ queryType: 'Flat' })).toBe(true);
    });

    it('should return true for roots/children containers and nested objects, otherwise false', () => {
      expect((ticketsDataProvider as any).hasAnyQueryTree({ roots: [{}] })).toBe(true);
      expect((ticketsDataProvider as any).hasAnyQueryTree({ children: [{}] })).toBe(true);
      expect((ticketsDataProvider as any).hasAnyQueryTree({ a: { b: [{ wiql: 'y' }] } })).toBe(true);
      expect((ticketsDataProvider as any).hasAnyQueryTree({ a: { b: [] } })).toBe(false);
    });
  });

  describe('matchesAreaPathCondition', () => {
    it('should require both source+target leaf matches when filters are provided', () => {
      const wiql =
        "SELECT * FROM WorkItemLinks WHERE Source.[System.AreaPath] = 'Test CMMI\\System' AND Target.[System.AreaPath] = 'Test CMMI\\Software'";
      const ok = (ticketsDataProvider as any).matchesAreaPathCondition(wiql, 'system', 'software');
      expect(ok).toBe(true);

      const badTarget = (ticketsDataProvider as any).matchesAreaPathCondition(wiql, 'system', 'nope');
      expect(badTarget).toBe(false);
    });

    it('should return false when filter is provided but area paths are not present for that owner', () => {
      const wiql = "SELECT * FROM WorkItemLinks WHERE Target.[System.AreaPath] = 'A\\B\\Leaf'";
      const ok = (ticketsDataProvider as any).matchesAreaPathCondition(wiql, 'leaf', 'leaf');
      expect(ok).toBe(false);
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
        workItemRelations: [{ source: null, target: { id: 1, url: 'url' } }],
      };
      const mockTableResult = {
        sourceTargetsMap: new Map(),
        sortingSourceColumnsMap: new Map(),
        sortingTargetsColumnsMap: new Map(),
      };

      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockQueryResult);
      jest
        .spyOn(ticketsDataProvider as any, 'parseDirectLinkedQueryResultForTableFormat')
        .mockResolvedValueOnce(mockTableResult);

      // Act
      const result = await ticketsDataProvider.GetQueryResultsFromWiql(mockWiqlHref, true, mockTestCaseMap);

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

  describe('parseDirectLinkedQueryResultForTableFormat', () => {
    it('should throw when workItemRelations is an empty array', async () => {
      await expect(
        (ticketsDataProvider as any).parseDirectLinkedQueryResultForTableFormat(
          { columns: [], workItemRelations: [], queryType: QueryType.OneHop } as any,
          new Map()
        )
      ).rejects.toThrow('No related work items were found');
    });

    it('should build sourceTargetsMap, include root links, and dedupe testCase->related items', async () => {
      const testCaseToRelatedWiMap = new Map<number, Set<any>>();
      const workItemRelations = [
        { source: null, target: { id: 1 } },
        { source: { id: 1 }, target: { id: 2 } },
        { source: { id: 1 }, target: { id: 2 } },
      ];

      const fetchSpy = jest
        .spyOn(ticketsDataProvider as any, 'fetchWIForQueryResult')
        .mockImplementation(async (...args: any[]) => {
          const _rel = args[0];
          const columnsToShowMap = args[1] as Map<string, string>;
          // Validate CustomerRequirementId -> Customer ID mapping is prepared
          expect(columnsToShowMap.get('CustomerRequirementId')).toBe('Customer ID');
          const id = _rel?.target?.id ?? _rel?.source?.id;
          if (id === 1) {
            return { id: 1, fields: { 'System.WorkItemType': 'Test Case' } };
          }
          if (id === 2) {
            return { id: 2, fields: { 'System.WorkItemType': 'Bug' } };
          }
          return { id, fields: { 'System.WorkItemType': 'Task' } };
        });

      const res = await (ticketsDataProvider as any).parseDirectLinkedQueryResultForTableFormat(
        {
          queryType: QueryType.OneHop,
          columns: [{ referenceName: 'CustomerRequirementId', name: 'CustomerRequirementId' }],
          workItemRelations,
        } as any,
        testCaseToRelatedWiMap
      );

      expect(fetchSpy).toHaveBeenCalled();
      expect(res.sourceTargetsMap).toBeInstanceOf(Map);

      const keys = Array.from(res.sourceTargetsMap.keys());
      expect(keys.some((k: any) => k.id === 1)).toBe(true);

      const relatedSet = testCaseToRelatedWiMap.get(1);
      expect(relatedSet).toBeDefined();
      expect(Array.from(relatedSet || [])).toHaveLength(1);
      expect(Array.from(relatedSet || [])[0]).toEqual(expect.objectContaining({ id: 2 }));
    });

    it('should throw when relation.target is missing', async () => {
      jest
        .spyOn(ticketsDataProvider as any, 'fetchWIForQueryResult')
        .mockResolvedValueOnce({ id: 1, fields: { 'System.WorkItemType': 'Test Case' } });

      await expect(
        (ticketsDataProvider as any).parseDirectLinkedQueryResultForTableFormat(
          {
            queryType: QueryType.OneHop,
            columns: [],
            workItemRelations: [{ source: { id: 1 }, target: null }],
          } as any,
          new Map()
        )
      ).rejects.toThrow('Target relation is missing');
    });
  });

  describe('isFlatQueryAllowedByTypeOrId', () => {
    it('should accept when allowedTypes is empty and WIQL references [System.WorkItemType]', async () => {
      const wiql = "SELECT * FROM WorkItems WHERE [System.WorkItemType] = 'Bug'";
      const res = await (ticketsDataProvider as any).isFlatQueryAllowedByTypeOrId({}, wiql, [], new Map());
      expect(res).toBe(true);
    });

    it('should return false when allowedTypes is provided but no types are found and no ids exist', async () => {
      const wiql = "SELECT * FROM WorkItems WHERE [System.Title] <> ''";
      const res = await (ticketsDataProvider as any).isFlatQueryAllowedByTypeOrId(
        {},
        wiql,
        ['Bug'],
        new Map()
      );
      expect(res).toBe(false);
    });

    it('should reject when WIQL contains a type outside allowedTypes', async () => {
      const wiql = "SELECT * FROM WorkItems WHERE [System.WorkItemType] IN ('Bug','Task')";
      const res = await (ticketsDataProvider as any).isFlatQueryAllowedByTypeOrId(
        {},
        wiql,
        ['Bug'],
        new Map()
      );
      expect(res).toBe(false);
    });
  });

  describe('matchesFlatAreaCondition', () => {
    it('should return true when filter is empty', () => {
      const wiql = 'SELECT * FROM WorkItems';
      const res = (ticketsDataProvider as any).matchesFlatAreaCondition(wiql, '');
      expect(res).toBe(true);
    });

    it('should return false when no [System.AreaPath] is present in WIQL', () => {
      const wiql = "SELECT * FROM WorkItems WHERE [System.Title] <> ''";
      const res = (ticketsDataProvider as any).matchesFlatAreaCondition(wiql, 'X');
      expect(res).toBe(false);
    });

    it('should match by leaf segment of area path (case-insensitive)', () => {
      const wiql = "SELECT * FROM WorkItems WHERE [System.AreaPath] = 'A\\B\\Leaf'";
      const res = (ticketsDataProvider as any).matchesFlatAreaCondition(wiql, 'leaf');
      expect(res).toBe(true);
    });
  });

  describe('filterFieldsByColumns', () => {
    it('should keep fields included by columns map and always include System.WorkItemType and System.Title', () => {
      const item: any = {
        fields: {
          'System.Title': 'T',
          'System.WorkItemType': 'Bug',
          'Custom.Keep': 1,
          'Custom.Drop': 2,
        },
      };
      const columnsToFilterMap = new Map<string, string>([['Custom.Keep', 'Keep']]);
      const resultedRefNameMap = new Map<string, string>();

      (ticketsDataProvider as any).filterFieldsByColumns(item, columnsToFilterMap, resultedRefNameMap);

      expect(Object.keys(item.fields).sort()).toEqual(['Custom.Keep', 'System.Title', 'System.WorkItemType']);
      expect(resultedRefNameMap.get('Custom.Keep')).toBe('Keep');
      expect(resultedRefNameMap.get('System.Title')).toBe('System.Title');
      expect(resultedRefNameMap.get('System.WorkItemType')).toBe('System.WorkItemType');
    });

    it('should log and throw when item.fields is invalid', () => {
      const item: any = { fields: undefined };
      const columnsToFilterMap = new Map<string, string>();
      const resultedRefNameMap = new Map<string, string>();

      expect(() =>
        (ticketsDataProvider as any).filterFieldsByColumns(item, columnsToFilterMap, resultedRefNameMap)
      ).toThrow();
      expect(logger.error).toHaveBeenCalledWith(expect.stringContaining('Cannot filter columns'));
    });
  });

  describe('GetWorkItemTypeList', () => {
    it('should fetch work item types and attempt icon download with fallback accepts', async () => {
      (TFSServices.getItemContent as jest.Mock).mockReset();
      if (!(TFSServices as any).fetchAzureDevOpsImageAsBase64) {
        (TFSServices as any).fetchAzureDevOpsImageAsBase64 = jest.fn();
      }
      (TFSServices.fetchAzureDevOpsImageAsBase64 as jest.Mock).mockReset();

      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce({
        value: [
          {
            name: 'Bug',
            referenceName: 'Microsoft.VSTS.WorkItemTypes.Bug',
            color: 'ff0000',
            icon: { id: 'i', url: 'http://example.com/icon' },
            states: [],
          },
        ],
      });

      // First accept fails, second succeeds
      (TFSServices.fetchAzureDevOpsImageAsBase64 as jest.Mock)
        .mockRejectedValueOnce(new Error('svg fail'))
        .mockResolvedValueOnce('data:image/png;base64,xxx');

      const res = await ticketsDataProvider.GetWorkItemTypeList(mockProject);
      expect(res).toHaveLength(1);
      expect(res[0]).toEqual(
        expect.objectContaining({
          name: 'Bug',
          icon: expect.objectContaining({ dataUrl: 'data:image/png;base64,xxx' }),
        })
      );
      expect(logger.warn).toHaveBeenCalledWith(expect.stringContaining('Failed to download icon'));
    });
  });

  describe('GetModeledQuery', () => {
    it('should structure query list correctly', () => {
      // Arrange
      const mockQueryList = [
        {
          name: 'Query 1',
          _links: { wiql: 'http://example.com/wiql1' },
          id: 'q1',
        },
        {
          name: 'Query 2',
          _links: { wiql: null },
          id: 'q2',
        },
      ];

      // Act
      const result = ticketsDataProvider.GetModeledQuery(mockQueryList);

      // Assert
      expect(result).toEqual([
        { queryName: 'Query 1', wiql: 'http://example.com/wiql1', id: 'q1' },
        { queryName: 'Query 2', wiql: null, id: 'q2' },
      ]);
    });
  });

  describe('PopulateWorkItemsByIds', () => {
    it('should fetch work items in batches of 200', async () => {
      // Arrange
      const mockIds = Array.from({ length: 250 }, (_, i) => i + 1);
      const mockResponse1 = { value: mockIds.slice(0, 200).map((id) => ({ id })) };
      const mockResponse2 = { value: mockIds.slice(200).map((id) => ({ id })) };

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
        mockProject,
        mockWiBody,
        mockWiType,
        mockByPass
      );

      // Assert
      expect(TFSServices.getItemContent).toHaveBeenCalledWith(
        `${mockOrgUrl}${mockProject}/_apis/wit/workitems/$${mockWiType}?bypassRules=true`,
        mockToken,
        'POST',
        mockWiBody,
        {
          'Content-Type': 'application/json-patch+json',
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
        mockProject,
        mockWiBody,
        mockWorkItemId,
        mockByPass
      );

      // Assert
      expect(TFSServices.getItemContent).toHaveBeenCalledWith(
        `${mockOrgUrl}${mockProject}/_apis/wit/workitems/${mockWorkItemId}?bypassRules=true`,
        mockToken,
        'patch',
        mockWiBody,
        {
          'Content-Type': 'application/json-patch+json',
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
            attributes: { name: 'file.txt' },
          },
        ],
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

    it('should return [] when relations are missing', async () => {
      jest
        .spyOn(ticketsDataProvider, 'GetWorkItem')
        .mockResolvedValueOnce({ id: 1, relations: undefined } as any);
      const res = await ticketsDataProvider.GetWorkitemAttachments(mockProject, '1');
      expect(res).toEqual([]);
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

    it('should log and return [] when GetWorkItem throws', async () => {
      jest.spyOn(TicketsDataProvider.prototype, 'GetWorkItem').mockRejectedValueOnce(new Error('boom'));
      const res = await ticketsDataProvider.GetWorkitemAttachments(mockProject, '1');
      expect(res).toEqual([]);
      expect(logger.error).toHaveBeenCalled();
    });

    it('should filter out non-attachment relations', async () => {
      // Arrange
      const mockId = '123';
      const mockWorkItem = {
        relations: [
          {
            rel: 'Parent',
            url: 'https://example.com/parent/1',
            attributes: { name: 'parent' },
          },
          {
            rel: 'AttachedFile',
            url: 'https://example.com/attachment/1',
            attributes: { name: 'file.txt' },
          },
        ],
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
      expect(result.some((item) => item.rel === 'Parent')).toBe(false);
    });
  });

  describe('GetWorkItemByUrl', () => {
    it('should fetch work item by URL', async () => {
      // Arrange
      const mockUrl = 'https://dev.azure.com/org/project/_apis/wit/workitems/123';
      const mockWorkItem = { id: 123, fields: { 'System.Title': 'Test' } };
      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockWorkItem);

      // Act
      const result = await ticketsDataProvider.GetWorkItemByUrl(mockUrl);

      // Assert
      expect(TFSServices.getItemContent).toHaveBeenCalledWith(mockUrl, mockToken);
      expect(result).toEqual(mockWorkItem);
    });
  });

  describe('GetParentLink', () => {
    it('should return trace with work item info', async () => {
      // Arrange
      const mockWi = {
        id: '123',
        fields: {
          'System.Title': 'Test Work Item',
          'System.CustomerId': 'CUST-001',
        },
      };

      // Act
      const result = await ticketsDataProvider.GetParentLink(mockProject, mockWi);

      // Assert
      expect(result.id).toBe('123');
      expect(result.title).toBe('Test Work Item');
      expect(result.customerId).toBe('CUST-001');
      expect(result.url).toContain(mockProject);
    });

    it('should handle work item without customerId', async () => {
      // Arrange
      const mockWi = {
        id: '123',
        fields: {
          'System.Title': 'Test Work Item',
        },
      };

      // Act
      const result = await ticketsDataProvider.GetParentLink(mockProject, mockWi);

      // Assert
      expect(result.id).toBe('123');
      expect(result.customerId).toBeUndefined();
    });

    it('should handle null work item', async () => {
      // Act
      const result = await ticketsDataProvider.GetParentLink(mockProject, null);

      // Assert
      expect(result.id).toBeUndefined();
    });
  });

  describe('GetRelationsIds', () => {
    it('should return a map from work items', async () => {
      // Arrange
      const mockIds = [
        {
          id: 1,
          relations: [{ rel: 'Parent', url: 'https://example.com/workitems/2' }],
        },
      ];

      // Act
      const result = await ticketsDataProvider.GetRelationsIds(mockIds);

      // Assert - result is a Map
      expect(result).toBeInstanceOf(Map);
    });

    it('should handle empty array', async () => {
      // Arrange
      const mockIds: any[] = [];

      // Act
      const result = await ticketsDataProvider.GetRelationsIds(mockIds);

      // Assert
      expect(result.size).toBe(0);
    });
  });

  describe('GetLinks', () => {
    it('should get links from work item relations', async () => {
      // Arrange
      const mockWi = {
        relations: [{ rel: 'Parent', url: 'https://example.com/workitems/2' }],
      };
      const mockLinks = [
        {
          id: '2',
          fields: {
            'System.Title': 'Parent Item',
            'System.Description': 'Parent Description',
          },
        },
      ];

      // Act
      const result = await ticketsDataProvider.GetLinks(mockProject, mockWi, mockLinks);

      // Assert
      expect(result).toHaveLength(1);
      expect(result[0].id).toBe('2');
      expect(result[0].title).toBe('Parent Item');
      expect(result[0].type).toBe('Parent');
    });
  });

  describe('GetFieldsByType', () => {
    it('should fetch fields for a work item type', async () => {
      // Arrange
      const mockItemType = 'Bug';
      const mockResponse = {
        value: [
          { name: 'Priority', referenceName: 'Microsoft.VSTS.Common.Priority' },
          { name: 'Severity', referenceName: 'Microsoft.VSTS.Common.Severity' },
          { name: 'ID', referenceName: 'System.Id' },
          { name: 'Title', referenceName: 'System.Title' },
        ],
      };
      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockResponse);

      // Act
      const result = await ticketsDataProvider.GetFieldsByType(mockProject, mockItemType);

      // Assert
      expect(TFSServices.getItemContent).toHaveBeenCalledWith(
        `${mockOrgUrl}${mockProject}/_apis/wit/workitemtypes/${mockItemType}/fields`,
        mockToken
      );
      // Should filter out ID and Title
      expect(result).toHaveLength(2);
      expect(result[0].text).toBe('Priority (Bug)');
      expect(result[0].key).toBe('Microsoft.VSTS.Common.Priority');
    });

    it('should throw error when API fails', async () => {
      // Arrange
      const mockError = new Error('API Error');
      (TFSServices.getItemContent as jest.Mock).mockRejectedValueOnce(mockError);

      // Act & Assert
      await expect(ticketsDataProvider.GetFieldsByType(mockProject, 'Bug')).rejects.toThrow();
      expect(logger.error).toHaveBeenCalled();
    });
  });

  describe('GetQueryResultsFromWiql - additional cases', () => {
    it('should handle Tree query type', async () => {
      // Arrange
      const mockWiqlHref = 'https://example.com/wiql';
      const mockTestCaseMap = new Map<number, Set<any>>();
      const mockQueryResult = {
        queryType: 'tree',
        workItemRelations: [{ source: null, target: { id: 1, url: 'https://example.com/wi/1' }, rel: null }],
      };
      const mockWiResponse = {
        fields: { 'System.Title': 'Test', 'System.Description': 'Desc' },
        _links: { html: { href: 'https://example.com' } },
      };

      (TFSServices.getItemContent as jest.Mock)
        .mockResolvedValueOnce(mockQueryResult)
        .mockResolvedValueOnce(mockWiResponse);

      // Act
      const result = await ticketsDataProvider.GetQueryResultsFromWiql(mockWiqlHref, false, mockTestCaseMap);

      // Assert
      expect(result).toBeDefined();
    });

    it('should handle Flat query type with table format', async () => {
      // Arrange
      const mockWiqlHref = 'https://example.com/wiql';
      const mockTestCaseMap = new Map<number, Set<any>>();
      const mockQueryResult = {
        queryType: 'flat',
        columns: [{ referenceName: 'System.Title', name: 'Title' }],
        workItems: [{ id: 1, url: 'https://example.com/wi/1' }],
      };
      const mockWiResponse = {
        id: 1,
        fields: { 'System.Title': 'Test Item' },
      };

      (TFSServices.getItemContent as jest.Mock)
        .mockResolvedValueOnce(mockQueryResult)
        .mockResolvedValueOnce(mockWiResponse);

      // Act
      const result = await ticketsDataProvider.GetQueryResultsFromWiql(mockWiqlHref, true, mockTestCaseMap);

      // Assert
      expect(result).toBeDefined();
    });

    it('should handle Flat query type without table format', async () => {
      // Arrange
      const mockWiqlHref = 'https://example.com/wiql';
      const mockTestCaseMap = new Map<number, Set<any>>();
      const mockQueryResult = {
        queryType: 'flat',
        workItems: [{ id: 1, url: 'https://example.com/wi/1' }],
      };
      const mockWiResponse = {
        fields: { 'System.Title': 'Test', 'System.Description': 'Desc' },
        _links: { html: { href: 'https://example.com' } },
      };

      (TFSServices.getItemContent as jest.Mock)
        .mockResolvedValueOnce(mockQueryResult)
        .mockResolvedValueOnce(mockWiResponse);

      // Act
      const result = await ticketsDataProvider.GetQueryResultsFromWiql(mockWiqlHref, false, mockTestCaseMap);

      // Assert
      expect(result).toBeDefined();
    });

    it('should return undefined when queryType is not supported', async () => {
      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce({ queryType: 'Unknown' });
      const res = await ticketsDataProvider.GetQueryResultsFromWiql(
        'https://example.com/wiql',
        false,
        new Map()
      );
      expect(res).toBeUndefined();
    });
  });

  describe('getRequirementTypeFieldRefs', () => {
    it('should prioritize known ref and dedupe candidates', async () => {
      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce({
        value: [
          { name: 'Requirement Type', referenceName: 'Microsoft.VSTS.CMMI.RequirementType' },
          { name: 'Requirement_Type', referenceName: 'Custom.RequirementType' },
          { name: 'Requirement Type', referenceName: 'Microsoft.VSTS.CMMI.RequirementType' },
          { name: 'Other', referenceName: 'Custom.Other' },
        ],
      });

      const res = await (ticketsDataProvider as any).getRequirementTypeFieldRefs('project1');
      expect(res[0]).toBe('Microsoft.VSTS.CMMI.RequirementType');
      expect(res).toContain('Custom.RequirementType');
      expect(res.filter((x: string) => x === 'Microsoft.VSTS.CMMI.RequirementType')).toHaveLength(1);
    });

    it('should fall back to known ref when API call throws', async () => {
      (TFSServices.getItemContent as jest.Mock).mockRejectedValueOnce(new Error('boom'));
      const res = await (ticketsDataProvider as any).getRequirementTypeFieldRefs('project1');
      expect(res).toEqual(['Microsoft.VSTS.CMMI.RequirementType']);
    });
  });

  describe('findQueryFolderByName', () => {
    it('should return null when input is invalid', async () => {
      await expect((ticketsDataProvider as any).findQueryFolderByName(null, 'x')).resolves.toBeNull();
      await expect((ticketsDataProvider as any).findQueryFolderByName({ id: 'r' }, '')).resolves.toBeNull();
    });

    it('should BFS and return a folder by name; using ensureQueryChildren when node hasChildren', async () => {
      const root: any = {
        id: 'root',
        name: 'Root',
        isFolder: true,
        hasChildren: true,
        children: [
          { id: 'a', name: 'A', isFolder: true, hasChildren: true },
          { id: 'b', name: 'B', isFolder: true, hasChildren: false, children: [] },
        ],
      };

      jest.spyOn(ticketsDataProvider as any, 'ensureQueryChildren').mockImplementation(async (node: any) => {
        if (node.id === 'a') {
          node.children = [{ id: 'target', name: 'TargetFolder', isFolder: true, hasChildren: false }];
        }
        return node;
      });

      const found = await (ticketsDataProvider as any).findQueryFolderByName(root, 'targetfolder');
      expect(found).toEqual(expect.objectContaining({ id: 'target' }));
    });

    it('should return null when folder is not found', async () => {
      const root: any = { id: 'root', name: 'Root', isFolder: true, hasChildren: false, children: [] };
      const found = await (ticketsDataProvider as any).findQueryFolderByName(root, 'missing');
      expect(found).toBeNull();
    });
  });

  describe('findChildFolderByName', () => {
    it('should return null when parent or childName is invalid', async () => {
      await expect((ticketsDataProvider as any).findChildFolderByName(null, 'x')).resolves.toBeNull();
      await expect((ticketsDataProvider as any).findChildFolderByName({ id: 1 }, '')).resolves.toBeNull();
    });

    it('should return null when parent has no children after enrichment', async () => {
      jest
        .spyOn(ticketsDataProvider as any, 'ensureQueryChildren')
        .mockResolvedValueOnce({ id: 'p', hasChildren: true, children: [] });

      const res = await (ticketsDataProvider as any).findChildFolderByName({ id: 'p' }, 'child');
      expect(res).toBeNull();
    });

    it('should return matching child folder by case-insensitive exact name', async () => {
      const parentWithChildren: any = {
        id: 'p',
        hasChildren: true,
        children: [
          { id: 'c1', isFolder: true, name: 'Match', hasChildren: false },
          { id: 'c2', isFolder: true, name: 'Other', hasChildren: false },
        ],
      };
      jest.spyOn(ticketsDataProvider as any, 'ensureQueryChildren').mockResolvedValueOnce(parentWithChildren);

      const res = await (ticketsDataProvider as any).findChildFolderByName({ id: 'p' }, 'match');
      expect(res).toEqual(parentWithChildren.children[0]);
    });
  });

  describe('buildFallbackChain + findPathToNode', () => {
    it('should use findPathToNode when startingFolder has id', async () => {
      const root: any = { id: 'root', name: 'Root', hasChildren: true, children: [] };
      const mid: any = { id: 'mid', name: 'Mid', hasChildren: true, children: [] };
      const start: any = { id: 'start', name: 'Start', hasChildren: false, children: [] };
      root.children = [mid];
      mid.children = [start];

      jest.spyOn(ticketsDataProvider as any, 'ensureQueryChildren').mockImplementation(async (n: any) => n);

      const path = await (ticketsDataProvider as any).findPathToNode(root, 'start');
      expect(path?.map((n: any) => n.id)).toEqual(['root', 'mid', 'start']);

      const chain = await (ticketsDataProvider as any).buildFallbackChain(root, start);
      expect(chain.map((n: any) => n.id)).toEqual(['start', 'mid', 'root']);
    });

    it('should fall back to startingFolder when findPathToNode returns null', async () => {
      const root: any = { id: 'root', name: 'Root', hasChildren: false, children: [] };
      const start: any = { id: 'start', name: 'Start', hasChildren: false, children: [] };

      jest.spyOn(ticketsDataProvider as any, 'findPathToNode').mockResolvedValueOnce(null);
      const chain = await (ticketsDataProvider as any).buildFallbackChain(root, start);
      expect(chain.map((n: any) => n.id)).toEqual(['start', 'root']);
    });

    it('should handle startingFolder without id and always include rootQueries', async () => {
      const root: any = { id: 'root', name: 'Root', hasChildren: false, children: [] };
      const start: any = { name: 'StartNoId', hasChildren: false, children: [] };
      const chain = await (ticketsDataProvider as any).buildFallbackChain(root, start);
      expect(chain).toHaveLength(2);
      expect(chain[0]).toBe(start);
      expect(chain[1]).toBe(root);
    });

    it('should return null from findPathToNode when a cycle is detected', async () => {
      const node: any = { id: 'x', name: 'X', hasChildren: true, children: [] };
      node.children = [node];
      jest.spyOn(ticketsDataProvider as any, 'ensureQueryChildren').mockImplementation(async (n: any) => n);
      const path = await (ticketsDataProvider as any).findPathToNode(node, 'missing');
      expect(path).toBeNull();
    });
  });

  describe('structureFetchedQueries', () => {
    it('should skip excluded folders', async () => {
      const res = await (ticketsDataProvider as any).structureFetchedQueries(
        { isFolder: true, name: 'SkipMe', hasChildren: true },
        false,
        null,
        ['Epic'],
        ['Feature'],
        undefined,
        undefined,
        false,
        ['skipme'],
        false
      );
      expect(res).toEqual({ tree1: null, tree2: null });
    });

    it('should fetch children when hasChildren=true but children are missing', async () => {
      jest
        .spyOn(ticketsDataProvider as any, 'matchesSourceTargetConditionAsync')
        .mockResolvedValueOnce(true)
        .mockResolvedValueOnce(false);
      jest
        .spyOn(ticketsDataProvider as any, 'buildQueryNode')
        .mockImplementation((rq: any, parentId: any) => ({
          id: rq.id,
          pId: parentId,
        }));

      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce({
        id: 'root',
        name: 'Root',
        hasChildren: false,
        isFolder: false,
        queryType: 'oneHop',
        wiql: "SELECT * FROM WorkItemLinks WHERE Source.[System.WorkItemType] = 'Epic' AND Target.[System.WorkItemType] = 'Feature'",
      });

      const res = await (ticketsDataProvider as any).structureFetchedQueries(
        { id: 'root', url: 'https://example.com/q', hasChildren: true, children: undefined },
        false,
        null,
        ['Epic'],
        ['Feature']
      );

      expect(TFSServices.getItemContent).toHaveBeenCalledWith(
        'https://example.com/q?$depth=2&$expand=all',
        mockToken
      );
      expect(res.tree1).toEqual({ id: 'root', pId: 'root' });
      expect(res.tree2).toBeNull();
    });

    it('should build tree nodes for flat queries when includeFlatQueries is enabled (tree1 only)', async () => {
      jest
        .spyOn(ticketsDataProvider as any, 'buildQueryNode')
        .mockImplementation((rq: any, parentId: any) => ({ id: rq.id, pId: parentId }));

      const wiql =
        "SELECT * FROM WorkItems WHERE [System.WorkItemType] = 'Bug' AND [System.AreaPath] = 'A\\Leaf'";
      const res = await (ticketsDataProvider as any).structureFetchedQueries(
        { id: 'q', isFolder: false, hasChildren: false, queryType: 'flat', wiql },
        false,
        null,
        ['Bug'],
        ['Task'],
        'leaf',
        'missing',
        false,
        [],
        true
      );

      expect(res.tree1).toEqual({ id: 'q', pId: null });
      expect(res.tree2).toBeNull();
    });

    it('should build tree2 for reverse source/target in oneHop queries', async () => {
      jest
        .spyOn(ticketsDataProvider as any, 'matchesSourceTargetConditionAsync')
        .mockResolvedValueOnce(false)
        .mockResolvedValueOnce(true);
      jest
        .spyOn(ticketsDataProvider as any, 'buildQueryNode')
        .mockImplementation((rq: any, parentId: any) => ({ id: rq.id, pId: parentId }));

      const wiql =
        "SELECT * FROM WorkItemLinks WHERE Source.[System.WorkItemType] = 'Feature' AND Target.[System.WorkItemType] = 'Epic'";
      const res = await (ticketsDataProvider as any).structureFetchedQueries(
        { id: 'q2', isFolder: false, hasChildren: false, queryType: 'oneHop', wiql },
        false,
        null,
        ['Epic'],
        ['Feature']
      );

      expect(res.tree1).toBeNull();
      expect(res.tree2).toEqual({ id: 'q2', pId: null });
    });

    describe('ID-only work item type fallback', () => {
      const projectHref = `${mockOrgUrl}MyProject/_apis/wit/wiql/123`;
      const allowedTypes = ['epic', 'feature', 'requirement'];

      const makeLeafQuery = (overrides: any = {}) => ({
        id: 'q1',
        name: 'Query 1',
        isFolder: false,
        hasChildren: false,
        queryType: 'oneHop',
        wiql: '',
        _links: { wiql: { href: projectHref } },
        ...overrides,
      });

      it('should include a tree query when Source has [System.Id] and no Source type filter, and the fetched work item type is allowed', async () => {
        (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce({
          fields: { 'System.WorkItemType': 'Epic' },
        });

        const queryNode = makeLeafQuery({
          queryType: 'tree',
          wiql: `
            SELECT * FROM WorkItemLinks
            WHERE
              ([Source].[System.Id] = 123)
              AND ([Target].[System.WorkItemType] IN ('Requirement','Epic','Feature'))
          `,
        });

        const res = await (ticketsDataProvider as any).structureFetchedQueries(
          queryNode,
          false,
          null,
          allowedTypes,
          [],
          undefined,
          undefined,
          true
        );

        expect(res.tree1).not.toBeNull();
        expect(TFSServices.getItemContent).toHaveBeenCalledTimes(1);
      });

      it('should exclude a tree query when Source has [System.Id] and no Source type filter, but the fetched work item type is not allowed', async () => {
        (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce({
          fields: { 'System.WorkItemType': 'Bug' },
        });

        const queryNode = makeLeafQuery({
          queryType: 'tree',
          wiql: `
            SELECT * FROM WorkItemLinks
            WHERE
              ([Source].[System.Id] = 123)
              AND ([Target].[System.WorkItemType] IN ('Requirement','Epic','Feature'))
          `,
        });

        const res = await (ticketsDataProvider as any).structureFetchedQueries(
          queryNode,
          false,
          null,
          allowedTypes,
          [],
          undefined,
          undefined,
          true
        );

        expect(res.tree1).toBeNull();
        expect(TFSServices.getItemContent).toHaveBeenCalledTimes(1);
      });

      it('should include a oneHop query when Source has [System.Id] and no Source type filter, and the fetched work item type is allowed', async () => {
        (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce({
          fields: { 'System.WorkItemType': 'Requirement' },
        });

        const queryNode = makeLeafQuery({
          queryType: 'oneHop',
          wiql: `
            SELECT * FROM WorkItemLinks
            WHERE
              ([Source].[System.Id] = 123)
              AND ([Target].[System.WorkItemType] = 'Feature')
          `,
        });

        const res = await (ticketsDataProvider as any).structureFetchedQueries(
          queryNode,
          false,
          null,
          allowedTypes,
          allowedTypes
        );

        expect(res.tree1).not.toBeNull();
        // reverse evaluation should reuse the cached lookup
        expect(TFSServices.getItemContent).toHaveBeenCalledTimes(1);
      });

      it('should include a flat query when [System.Id] exists and no [System.WorkItemType] filter exists, and the fetched work item type is allowed', async () => {
        (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce({
          fields: { 'System.WorkItemType': 'Feature' },
        });

        const queryNode = makeLeafQuery({
          queryType: 'flat',
          wiql: `
            SELECT * FROM WorkItems
            WHERE [System.Id] = 123
          `,
        });

        const res = await (ticketsDataProvider as any).structureFetchedQueries(
          queryNode,
          false,
          null,
          allowedTypes,
          [],
          undefined,
          undefined,
          false,
          [],
          true
        );

        expect(res.tree1).not.toBeNull();
        expect(TFSServices.getItemContent).toHaveBeenCalledTimes(1);
      });

      it('should not trigger the ID fallback when no [System.Id] filter exists', async () => {
        const queryNode = makeLeafQuery({
          queryType: 'tree',
          wiql: `
            SELECT * FROM WorkItemLinks
            WHERE
              ([Target].[System.WorkItemType] IN ('Requirement','Epic','Feature'))
          `,
        });

        const res = await (ticketsDataProvider as any).structureFetchedQueries(
          queryNode,
          false,
          null,
          allowedTypes,
          [],
          undefined,
          undefined,
          true
        );

        expect(res.tree1).toBeNull();
        expect(TFSServices.getItemContent).not.toHaveBeenCalled();
      });
    });
  });

  describe('structureAllQueryPath', () => {
    it('should return sysOverview+knownBugs nodes for a flat Bug leaf query', async () => {
      const leaf: any = {
        id: 'q1',
        name: 'BugQuery',
        hasChildren: false,
        isFolder: false,
        queryType: 'flat',
        columns: [],
        wiql: "SELECT * FROM WorkItems WHERE [System.WorkItemType] = 'Bug'",
        _links: { wiql: { href: 'http://example.com/wiql' } },
      };

      const res = await (ticketsDataProvider as any).structureAllQueryPath(leaf, 'p');
      expect(res.tree1).toEqual(expect.objectContaining({ id: 'q1', pId: 'p', isValidQuery: true }));
      expect(res.tree2).toEqual(expect.objectContaining({ id: 'q1', pId: 'p', isValidQuery: true }));
    });

    it('should return sysOverview only for a flat non-bug leaf query', async () => {
      const leaf: any = {
        id: 'q2',
        name: 'OtherQuery',
        hasChildren: false,
        isFolder: false,
        queryType: 'flat',
        columns: [],
        wiql: "SELECT * FROM WorkItems WHERE [System.WorkItemType] = 'Task'",
        _links: { wiql: { href: 'http://example.com/wiql' } },
      };

      const res = await (ticketsDataProvider as any).structureAllQueryPath(leaf, 'p');
      expect(res.tree1).toEqual(expect.objectContaining({ id: 'q2', pId: 'p', isValidQuery: true }));
      expect(res.tree2).toBeNull();
    });

    it('should return null trees for a folder leaf', async () => {
      const leafFolder: any = { id: 'f', name: 'Folder', hasChildren: false, isFolder: true };
      const res = await (ticketsDataProvider as any).structureAllQueryPath(leafFolder, 'p');
      expect(res).toEqual({ tree1: null, tree2: null });
    });

    it('should fetch query children from url when children are missing', async () => {
      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce({
        id: 'q3',
        name: 'Fetched',
        hasChildren: false,
        isFolder: false,
        queryType: 'flat',
        columns: [],
        wiql: "SELECT * FROM WorkItems WHERE [System.WorkItemType] = 'Bug'",
        _links: { wiql: { href: 'http://example.com/wiql' } },
      });

      const res = await (ticketsDataProvider as any).structureAllQueryPath({
        id: 'q3',
        url: 'https://example.com/q3',
        name: 'NeedFetch',
        hasChildren: true,
        children: undefined,
      });

      expect(TFSServices.getItemContent).toHaveBeenCalledWith(
        'https://example.com/q3?$depth=2&$expand=all',
        mockToken
      );
      expect(res.tree1).toEqual(expect.objectContaining({ id: 'q3', pId: 'q3' }));
    });
  });

  describe('parseTreeQueryResult', () => {
    it('should build roots, skip non-hierarchy links, and dedupe children', async () => {
      const initSpy = jest
        .spyOn(ticketsDataProvider as any, 'initTreeQueryResultItem')
        .mockImplementation(async (item: any, allItems: any) => {
          allItems[item.id] = {
            id: item.id,
            title: `T${item.id}`,
            description: '',
            htmlUrl: `h${item.id}`,
            children: [],
          };
        });

      const workItemRelations = [
        { source: null, target: { id: 1, url: 'u1' }, rel: null },
        { source: null, target: { id: 1, url: 'u1' }, rel: null },
        {
          source: { id: 1, url: 'u1' },
          target: { id: 2, url: 'u2' },
          rel: 'System.LinkTypes.Hierarchy-Forward',
        },
        {
          source: { id: 1, url: 'u1' },
          target: { id: 2, url: 'u2' },
          rel: 'System.LinkTypes.Hierarchy-Forward',
        },
        {
          source: { id: 1, url: 'u1' },
          target: { id: 3, url: 'u3' },
          rel: 'System.LinkTypes.Related',
        },
      ];

      const res = await (ticketsDataProvider as any).parseTreeQueryResult({ workItemRelations } as any);

      expect(initSpy).toHaveBeenCalled();
      expect(res.roots).toHaveLength(1);
      expect(res.roots[0].id).toBe(1);
      expect(res.roots[0].children).toHaveLength(1);
      expect(res.roots[0].children[0].id).toBe(2);
      expect(Object.keys(res.allItems)).toEqual(expect.arrayContaining(['1', '2', '3']));
    });

    it('should warn and initialize missing parent/child during hierarchy link processing when init no-ops first pass', async () => {
      const warnSpy = jest.spyOn(logger, 'warn');

      const callCountById = new Map<number, number>();
      jest
        .spyOn(ticketsDataProvider as any, 'initTreeQueryResultItem')
        .mockImplementation(async (item: any, allItems: any) => {
          const count = (callCountById.get(item.id) || 0) + 1;
          callCountById.set(item.id, count);
          // First time: do nothing (simulate unexpected missing init)
          if (count === 1) return;
          // Second time: actually initialize
          allItems[item.id] = {
            id: item.id,
            title: `T${item.id}`,
            description: '',
            htmlUrl: `h${item.id}`,
            children: [],
          };
        });

      const workItemRelations = [
        {
          source: { id: 10, url: 'u10' },
          target: { id: 11, url: 'u11' },
          rel: 'System.LinkTypes.Hierarchy-Forward',
        },
      ];

      const res = await (ticketsDataProvider as any).parseTreeQueryResult({ workItemRelations } as any);

      expect(warnSpy).toHaveBeenCalled();
      expect(res.roots).toHaveLength(0);
      expect(res.allItems[10]).toBeDefined();
      expect(res.allItems[11]).toBeDefined();
      expect(res.allItems[10].children).toHaveLength(1);
      expect(res.allItems[10].children[0].id).toBe(11);
    });
  });

  describe('GetQueryResultsByWiqlHref', () => {
    it('should fetch and model query results', async () => {
      // Arrange
      const mockWiqlHref = 'https://example.com/wiql';
      const mockResults = {
        workItems: [{ id: 1, url: 'https://example.com/wi/1' }],
      };
      const mockWiResponse = {
        id: 1,
        fields: { 'System.Title': 'Test' },
        _links: { html: { href: 'https://example.com' } },
      };

      (TFSServices.getItemContent as jest.Mock)
        .mockResolvedValueOnce(mockResults)
        .mockResolvedValueOnce(mockWiResponse);

      // Act
      const result = await ticketsDataProvider.GetQueryResultsByWiqlHref(mockWiqlHref, mockProject);

      // Assert
      expect(TFSServices.getItemContent).toHaveBeenCalledWith(mockWiqlHref, mockToken);
    });
  });

  describe('GetWorkItemByUrl', () => {
    it('should fetch work item by URL', async () => {
      // Arrange
      const mockUrl = 'https://dev.azure.com/org/project/_apis/wit/workitems/123';

      // Act
      const result = await ticketsDataProvider.GetWorkItemByUrl(mockUrl);

      // Assert
      expect(TFSServices.getItemContent).toHaveBeenCalled();
    });
  });

  describe('GetSharedQueries - docType branches', () => {
    it('should handle STR docType', async () => {
      // Arrange
      const mockQueries = {
        children: [{ name: 'STR', isFolder: true, children: [] }],
      };
      (TFSServices.getItemContent as jest.Mock).mockResolvedValue(mockQueries);

      // Act
      const result = await ticketsDataProvider.GetSharedQueries(mockProject, '', 'str');

      // Assert
      expect(result).toBeDefined();
    });

    it('should handle test-reporter docType', async () => {
      // Arrange
      const mockQueries = {
        children: [{ name: 'Test Reporter', isFolder: true, children: [] }],
      };
      (TFSServices.getItemContent as jest.Mock).mockResolvedValue(mockQueries);

      // Act
      const result = await ticketsDataProvider.GetSharedQueries(mockProject, '', 'test-reporter');

      // Assert
      expect(result).toBeDefined();
    });

    it('should handle SRS docType', async () => {
      // Arrange
      const mockQueries = {
        children: [{ name: 'SRS', isFolder: true, children: [] }],
      };
      (TFSServices.getItemContent as jest.Mock).mockResolvedValue(mockQueries);

      // Act
      const result = await ticketsDataProvider.GetSharedQueries(mockProject, '', 'srs');

      // Assert
      expect(result).toBeDefined();
    });

    it('should handle unknown docType', async () => {
      // Arrange
      const mockQueries = { children: [] };
      (TFSServices.getItemContent as jest.Mock).mockResolvedValue(mockQueries);

      // Act
      const result = await ticketsDataProvider.GetSharedQueries(mockProject, '', 'unknown');

      // Assert
      expect(result).toBeUndefined();
    });

    it('should handle error in GetSharedQueries', async () => {
      // Arrange
      (TFSServices.getItemContent as jest.Mock).mockRejectedValueOnce(new Error('API Error'));

      // Act & Assert
      await expect(ticketsDataProvider.GetSharedQueries(mockProject, '', 'std')).rejects.toThrow('API Error');
      expect(logger.error).toHaveBeenCalled();
    });
  });

  describe('PopulateWorkItemsByIds', () => {
    it('should populate work items by IDs', async () => {
      // Arrange
      const mockIds = [1, 2, 3];
      const mockResponse = {
        value: [
          { id: 1, fields: { 'System.Title': 'Item 1' } },
          { id: 2, fields: { 'System.Title': 'Item 2' } },
          { id: 3, fields: { 'System.Title': 'Item 3' } },
        ],
      };
      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce(mockResponse);

      // Act
      const result = await ticketsDataProvider.PopulateWorkItemsByIds(mockIds, mockProject);

      // Assert
      expect(result).toEqual(mockResponse.value);
    });

    it('should handle empty IDs array', async () => {
      // Arrange
      const mockIds: number[] = [];

      // Act
      const result = await ticketsDataProvider.PopulateWorkItemsByIds(mockIds, mockProject);

      // Assert
      expect(result).toEqual([]);
    });
  });

  describe('GetRelationsIds - edge cases', () => {
    it('should return a Map from work items with relations', async () => {
      // Arrange
      const mockIds = [
        {
          id: 1,
          relations: [{ rel: 'Parent', url: 'https://example.com/workitems/2' }],
        },
      ];

      // Act
      const result = await ticketsDataProvider.GetRelationsIds(mockIds);

      // Assert
      expect(result).toBeInstanceOf(Map);
    });
  });

  describe('GetLinks - edge cases', () => {
    it('should handle empty relations', async () => {
      // Arrange
      const mockWi = { relations: [] };
      const mockLinks: any[] = [];

      // Act
      const result = await ticketsDataProvider.GetLinks(mockProject, mockWi, mockLinks);

      // Assert
      expect(result).toHaveLength(0);
    });
  });

  describe('GetSharedQueries - path variations', () => {
    it('should use path in URL when provided', async () => {
      // Arrange
      const mockPath = 'My Queries/Test';
      const mockQueries = { children: [] };
      (TFSServices.getItemContent as jest.Mock).mockResolvedValue(mockQueries);

      // Act
      await ticketsDataProvider.GetSharedQueries(mockProject, mockPath, '');

      // Assert
      expect(TFSServices.getItemContent).toHaveBeenCalledWith(expect.stringContaining(mockPath), mockToken);
    });
  });

  describe('GetQueryResultById', () => {
    it('should fetch query result by ID', async () => {
      // Arrange
      const mockQueryId = 'query-123';
      const mockQuery = {
        _links: { wiql: { href: 'https://example.com/wiql' } },
      };
      const mockWiqlResult = { workItems: [{ id: 1 }] };
      const mockWiResponse = {
        fields: { 'System.Title': 'Test' },
        _links: { html: { href: 'https://example.com' } },
      };

      (TFSServices.getItemContent as jest.Mock)
        .mockResolvedValueOnce(mockQuery)
        .mockResolvedValueOnce(mockWiqlResult)
        .mockResolvedValueOnce(mockWiResponse);

      // Act
      const result = await ticketsDataProvider.GetQueryResultById(mockProject, mockQueryId);

      // Assert
      expect(TFSServices.getItemContent).toHaveBeenCalled();
    });
  });

  describe('GetModeledQueryResults', () => {
    it('should model Flat query results including System.Id and System.AssignedTo displayName', async () => {
      const results: any = {
        asOf: 'now',
        queryResultType: 'workItem',
        queryType: QueryType.Flat,
        workItems: [{ id: '1', url: 'https://example.com/wi/1' }],
        columns: [
          { name: 'ID', referenceName: 'System.Id', url: 'u1' },
          { name: 'Assigned To', referenceName: 'System.AssignedTo', url: 'u2' },
          { name: 'Custom', referenceName: 'Custom.Field', url: 'u3' },
        ],
      };

      jest.spyOn(TicketsDataProvider.prototype, 'GetWorkItem').mockResolvedValueOnce({
        id: 1,
        url: 'https://example.com/wi/1',
        fields: {
          'System.AssignedTo': { displayName: 'Bob' },
          'Custom.Field': 'X',
          'System.Id': 1,
        },
        relations: [{ rel: 'AttachedFile', url: 'https://example.com/a', attributes: { name: 'a.txt' } }],
      } as any);

      const modeled = await ticketsDataProvider.GetModeledQueryResults(results, mockProject);

      expect(modeled.queryType).toBe(QueryType.Flat);
      expect(modeled.workItems).toHaveLength(1);
      expect(modeled.workItems[0].attachments).toBeDefined();
      // System.Id branch
      expect(modeled.workItems[0].fields[0]).toEqual(expect.objectContaining({ name: 'ID', value: '1' }));
      // System.AssignedTo branch
      // Note: implementation does not set fields[i].name for AssignedTo branch
      expect(modeled.workItems[0].fields[1]).toEqual(expect.objectContaining({ value: 'Bob' }));
      // default field branch
      expect(modeled.workItems[0].fields[2]).toEqual(expect.objectContaining({ name: 'Custom', value: 'X' }));
    });

    it('should model non-Flat query results using workItemRelations and set Source when present', async () => {
      const results: any = {
        asOf: 'now',
        queryResultType: 'workItemLink',
        queryType: QueryType.OneHop,
        workItemRelations: [
          {
            source: { id: 10 },
            target: { id: 20 },
          },
        ],
        columns: [
          { name: 'Assigned To', referenceName: 'System.AssignedTo', url: 'u2' },
          { name: 'Title', referenceName: 'System.Title', url: 'u3' },
        ],
      };

      jest.spyOn(TicketsDataProvider.prototype, 'GetWorkItem').mockResolvedValueOnce({
        id: 20,
        url: 'https://example.com/wi/20',
        fields: {
          // Ensure AssignedTo branch falls through to the default path
          'System.AssignedTo': null,
          'System.Title': 'T20',
        },
        relations: null,
      } as any);

      const modeled = await ticketsDataProvider.GetModeledQueryResults(results, mockProject);

      expect(modeled.queryType).toBe(QueryType.OneHop);
      expect(modeled.workItems).toHaveLength(1);
      expect(modeled.workItems[0].Source).toBe(10);
      expect(modeled.workItems[0].url).toBe('https://example.com/wi/20');
    });
  });

  describe('GetCategorizedRequirementsByType', () => {
    it('should categorize requirement items from queryResult.workItems and include priority=1 in precedence category', async () => {
      const wiqlHref = `${mockOrgUrl}project1/_apis/wit/wiql/123`;

      // Other tests in this file may set a persistent mockResolvedValue on getItemContent.
      // clearAllMocks() does not reset implementations, so ensure a clean slate.
      (TFSServices.getItemContent as jest.Mock).mockReset();

      (TFSServices.getItemContent as jest.Mock)
        // query result
        .mockResolvedValueOnce({ workItems: [{ id: 10 }, { id: 11 }, { id: 12 }, { id: 13 }] })
        // getRequirementTypeFieldRefs(project)
        .mockResolvedValueOnce({
          value: [
            { name: 'Requirement Type', referenceName: 'Microsoft.VSTS.CMMI.RequirementType' },
            { name: 'Requirement_Type', referenceName: 'Custom.RequirementType' },
          ],
        })
        // WI 10 requirement security, priority=1
        .mockResolvedValueOnce({
          id: 10,
          fields: {
            'System.WorkItemType': 'Requirement',
            'System.Title': 'R10',
            'System.Description': 'D10',
            'Microsoft.VSTS.CMMI.RequirementType': 'Security',
            'Microsoft.VSTS.Common.Priority': 1,
          },
          _links: { html: { href: 'h10' } },
        })
        // WI 11 requirement unknown type -> Other Requirements
        .mockResolvedValueOnce({
          id: 11,
          fields: {
            'System.WorkItemType': 'Requirement',
            'System.Title': 'R11',
            'System.Description': 'D11',
            'Microsoft.VSTS.CMMI.RequirementType': 'UnknownType',
          },
          _links: { html: { href: 'h11' } },
        })
        // WI 12 not a requirement -> skipped
        .mockResolvedValueOnce({
          id: 12,
          fields: {
            'System.WorkItemType': 'Bug',
            'System.Title': 'B12',
          },
          _links: { html: { href: 'h12' } },
        })
        // WI 13 throws -> warn branch
        .mockRejectedValueOnce(new Error('boom'));

      const res = await ticketsDataProvider.GetCategorizedRequirementsByType(wiqlHref);

      expect(res.totalCount).toBe(4);
      expect(res.categories['Security and Privacy Requirements']).toHaveLength(1);
      expect(res.categories['Precedence and Criticality of Requirements']).toHaveLength(1);
      expect(res.categories['Other Requirements']).toHaveLength(1);
      expect(logger.warn).toHaveBeenCalledWith(expect.stringContaining('Could not fetch work item 13'));
    });

    it('should extract IDs from workItemRelations for OneHop queries', async () => {
      const wiqlHref = `${mockOrgUrl}project1/_apis/wit/wiql/rel`;
      (TFSServices.getItemContent as jest.Mock).mockReset();

      (TFSServices.getItemContent as jest.Mock)
        .mockResolvedValueOnce({
          workItemRelations: [
            { source: { id: 10 }, target: { id: 20 } },
            { source: { id: 10 }, target: { id: 21 } },
          ],
        })
        .mockResolvedValueOnce({ value: [] })
        .mockResolvedValueOnce({
          id: 10,
          fields: { 'System.WorkItemType': 'Requirement', 'System.Title': 'R10' },
          _links: { html: { href: 'h10' } },
        })
        .mockResolvedValueOnce({
          id: 20,
          fields: { 'System.WorkItemType': 'Bug', 'System.Title': 'B20' },
          _links: { html: { href: 'h20' } },
        })
        .mockResolvedValueOnce({
          id: 21,
          fields: { 'System.WorkItemType': 'Requirement', 'System.Title': 'R21' },
          _links: { html: { href: 'h21' } },
        });

      const res = await ticketsDataProvider.GetCategorizedRequirementsByType(wiqlHref);
      expect(res.totalCount).toBe(3);
    });

    it('should return empty when query result has no workItems and no workItemRelations', async () => {
      const wiqlHref = `${mockOrgUrl}project1/_apis/wit/wiql/empty`;
      (TFSServices.getItemContent as jest.Mock).mockReset();
      (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce({});

      const res = await ticketsDataProvider.GetCategorizedRequirementsByType(wiqlHref);
      expect(res).toEqual({ categories: {}, totalCount: 0 });
      expect(logger.warn).toHaveBeenCalledWith(
        expect.stringContaining('No work items found in query result')
      );
    });
  });
});
