import { TFSServices } from '../../helpers/tfs';
import TicketsDataProvider from '../../modules/TicketsDataProvider';

jest.mock('../../helpers/tfs');
jest.mock('../../utils/logger', () => ({
  __esModule: true,
  default: {
    debug: jest.fn(),
    info: jest.fn(),
    warn: jest.fn(),
    error: jest.fn(),
  },
}));

describe('TicketsDataProvider historical queries', () => {
  const orgUrl = 'https://dev.azure.com/org/';
  const token = 'pat';
  const project = 'team-project';
  let provider: TicketsDataProvider;

  beforeEach(() => {
    jest.clearAllMocks();
    provider = new TicketsDataProvider(orgUrl, token);
  });

  it('GetHistoricalQueries flattens shared query tree into a sorted list with explicit api-version', async () => {
    (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce({
      id: 'root',
      name: 'Shared Queries',
      isFolder: true,
      children: [
        {
          id: 'folder-b',
          name: 'B',
          isFolder: true,
          children: [{ id: 'q-2', name: 'Second Query', isFolder: false }],
        },
        {
          id: 'folder-a',
          name: 'A',
          isFolder: true,
          children: [{ id: 'q-1', name: 'First Query', isFolder: false }],
        },
      ],
    });

    const result = await provider.GetHistoricalQueries(project);

    expect(TFSServices.getItemContent).toHaveBeenCalledWith(
      expect.stringContaining(`/${project}/_apis/wit/queries/Shared%20Queries?$depth=2&$expand=all&api-version=7.1`),
      token,
    );
    expect(result).toEqual([
      { id: 'q-1', queryName: 'First Query', path: 'Shared Queries/A' },
      { id: 'q-2', queryName: 'Second Query', path: 'Shared Queries/B' },
    ]);
  });

  it('GetHistoricalQueries supports legacy response shape and falls back to default api-version', async () => {
    (TFSServices.getItemContent as jest.Mock).mockImplementation(async (url: string) => {
      if (url.includes('api-version=7.1') || url.includes('api-version=5.1')) {
        throw {
          response: {
            status: 400,
            data: { message: 'The requested api-version is not supported.' },
          },
        };
      }
      if (url.includes('/_apis/wit/queries/Shared%20Queries')) {
        return {
          value: [
            {
              id: 'q-legacy',
              name: 'Legacy Query',
              isFolder: false,
            },
          ],
        };
      }
      throw new Error(`unexpected URL: ${url}`);
    });

    const result = await provider.GetHistoricalQueries(project);

    expect(result).toEqual([{ id: 'q-legacy', queryName: 'Legacy Query', path: 'Shared Queries' }]);
  });

  it('GetHistoricalQueries retries with 5.1 when 7.1 returns 500', async () => {
    (TFSServices.getItemContent as jest.Mock).mockImplementation(async (url: string) => {
      if (url.includes('api-version=7.1')) {
        throw {
          response: {
            status: 500,
            data: { message: 'Internal Server Error' },
          },
        };
      }
      if (url.includes('api-version=5.1')) {
        return {
          id: 'root',
          name: 'Shared Queries',
          isFolder: true,
          children: [{ id: 'q-51', name: 'V5 Query', isFolder: false }],
        };
      }
      throw new Error(`unexpected URL: ${url}`);
    });

    const result = await provider.GetHistoricalQueries(project);

    expect(result).toEqual([{ id: 'q-51', queryName: 'V5 Query', path: 'Shared Queries' }]);
    expect(TFSServices.getItemContent).toHaveBeenCalledWith(
      expect.stringContaining(`/${project}/_apis/wit/queries/Shared%20Queries?$depth=2&$expand=all&api-version=7.1`),
      token,
    );
    expect(TFSServices.getItemContent).toHaveBeenCalledWith(
      expect.stringContaining(`/${project}/_apis/wit/queries/Shared%20Queries?$depth=2&$expand=all&api-version=5.1`),
      token,
    );
  });

  it('GetHistoricalQueries treats "Shared Queries" alias as shared root', async () => {
    (TFSServices.getItemContent as jest.Mock).mockResolvedValueOnce({
      id: 'root',
      name: 'Shared Queries',
      isFolder: true,
      children: [],
    });

    await provider.GetHistoricalQueries(project, 'Shared Queries');

    expect(TFSServices.getItemContent).toHaveBeenCalledWith(
      expect.stringContaining(`/${project}/_apis/wit/queries/Shared%20Queries?$depth=2&$expand=all&api-version=7.1`),
      token,
    );
  });

  it('GetHistoricalQueryResults executes WIQL with ASOF and returns as-of snapshot rows', async () => {
    const asOfIso = '2026-01-01T10:00:00.000Z';
    (TFSServices.getItemContent as jest.Mock).mockImplementation(
      async (url: string, _pat: string, method?: string, data?: any) => {
        if (url.includes('/_apis/wit/queries/q-1') && url.includes('api-version=7.1')) {
          return { name: 'Historical Q', wiql: 'SELECT [System.Id] FROM WorkItems' };
        }
        if (url.includes('/_apis/wit/wiql?') && url.includes('api-version=7.1') && method === 'post') {
          expect(String(data?.query || '')).toContain(`ASOF '${asOfIso}'`);
          return { workItems: [{ id: 101 }, { id: 102 }] };
        }
        if (url.includes('/_apis/wit/workitemsbatch') && url.includes('api-version=7.1') && method === 'post') {
          expect(data.asOf).toBe(asOfIso);
          return {
            value: [
              {
                id: 101,
                rev: 3,
                fields: {
                  'System.WorkItemType': 'Requirement',
                  'System.Title': 'Req title',
                  'System.State': 'Active',
                  'System.AreaPath': 'Proj\\Area',
                  'System.IterationPath': 'Proj\\Iter',
                  'System.ChangedDate': '2025-12-30T10:00:00Z',
                },
                relations: [],
              },
              {
                id: 102,
                rev: 8,
                fields: {
                  'System.WorkItemType': 'Bug',
                  'System.Title': 'Bug title',
                  'System.State': 'Closed',
                  'System.AreaPath': 'Proj\\Area',
                  'System.IterationPath': 'Proj\\Iter',
                  'System.ChangedDate': '2025-12-31T10:00:00Z',
                },
                relations: [],
              },
            ],
          };
        }
        throw new Error(`unexpected URL: ${url}`);
      },
    );

    const result = await provider.GetHistoricalQueryResults('q-1', project, asOfIso);

    const deprecatedWiqlByIdCallUsed = (TFSServices.getItemContent as jest.Mock).mock.calls.some((call) =>
      String(call[0]).includes('/_apis/wit/wiql/q-1'),
    );
    expect(deprecatedWiqlByIdCallUsed).toBe(false);
    expect(result.queryName).toBe('Historical Q');
    expect(result.asOf).toBe(asOfIso);
    expect(result.total).toBe(2);
    expect(result.rows[0]).toEqual(
      expect.objectContaining({
        id: 101,
        workItemType: 'Requirement',
        title: 'Req title',
        versionId: 3,
      }),
    );
  });

  it('GetHistoricalQueryResults falls back to api-version 5.1 and chunks workitemsbatch by 200 IDs', async () => {
    const asOfIso = '2026-01-01T10:00:00.000Z';
    const allIds = Array.from({ length: 205 }, (_, idx) => idx + 1);
    (TFSServices.getItemContent as jest.Mock).mockImplementation(
      async (url: string, _pat: string, method?: string, data?: any) => {
        if (url.includes('/_apis/wit/queries/q-fallback') && url.includes('api-version=7.1')) {
          throw {
            response: {
              status: 400,
              data: { message: 'The requested api-version is not supported.' },
            },
          };
        }
        if (url.includes('/_apis/wit/queries/q-fallback') && url.includes('api-version=5.1')) {
          return { name: 'Fallback Q', wiql: 'SELECT [System.Id] FROM WorkItems' };
        }
        if (url.includes('/_apis/wit/wiql?') && url.includes('api-version=5.1') && method === 'post') {
          expect(String(data?.query || '')).toContain(`ASOF '${asOfIso}'`);
          return { workItems: allIds.map((id) => ({ id })) };
        }
        if (url.includes('/_apis/wit/workitemsbatch') && url.includes('api-version=5.1') && method === 'post') {
          return {
            value: (Array.isArray(data?.ids) ? data.ids : []).map((id: number) => ({
              id,
              rev: 1,
              fields: {
                'System.WorkItemType': 'Bug',
                'System.Title': `Bug ${id}`,
                'System.State': 'Active',
                'System.AreaPath': 'Proj\\Area',
                'System.IterationPath': 'Proj\\Iter',
                'System.ChangedDate': '2025-12-31T10:00:00Z',
              },
              relations: [],
            })),
          };
        }
        throw new Error(`unexpected URL: ${url}`);
      },
    );

    const result = await provider.GetHistoricalQueryResults('q-fallback', project, asOfIso);

    expect(result.total).toBe(205);
    const batchCalls = (TFSServices.getItemContent as jest.Mock).mock.calls.filter(
      (call) =>
        String(call[0]).includes('/_apis/wit/workitemsbatch') &&
        String(call[0]).includes('api-version=5.1') &&
        String(call[2]).toLowerCase() === 'post',
    );
    expect(batchCalls).toHaveLength(2);
    expect(batchCalls[0][3].ids).toHaveLength(200);
    expect(batchCalls[1][3].ids).toHaveLength(5);
  });

  it('GetHistoricalQueryResults falls back to per-item retrieval when workitemsbatch fails', async () => {
    const asOfIso = '2026-01-01T10:00:00.000Z';
    (TFSServices.getItemContent as jest.Mock).mockImplementation(
      async (url: string, _pat: string, method?: string, data?: any) => {
        if (url.includes('/_apis/wit/queries/q-batch-fallback') && url.includes('api-version=7.1')) {
          return { name: 'Batch Fallback Q', wiql: 'SELECT [System.Id] FROM WorkItems' };
        }
        if (url.includes('/_apis/wit/wiql?') && url.includes('api-version=7.1') && method === 'post') {
          expect(String(data?.query || '')).toContain(`ASOF '${asOfIso}'`);
          return { workItems: [{ id: 101 }, { id: 102 }] };
        }
        if (url.includes('/_apis/wit/workitemsbatch') && url.includes('api-version=7.1') && method === 'post') {
          throw {
            response: {
              status: 500,
              data: { message: 'workitemsbatch failed' },
            },
          };
        }
        if (url.includes('/_apis/wit/workitems/101') && url.includes('api-version=7.1')) {
          expect(url).toContain('$expand=Relations');
          expect(url).toContain(`asOf=${encodeURIComponent(asOfIso)}`);
          expect(url).not.toContain('fields=');
          return {
            id: 101,
            rev: 3,
            fields: {
              'System.WorkItemType': 'Requirement',
              'System.Title': 'Req 101',
              'System.State': 'Active',
              'System.AreaPath': 'Proj\\Area',
              'System.IterationPath': 'Proj\\Iter',
              'System.ChangedDate': '2025-12-30T10:00:00Z',
            },
            relations: [],
          };
        }
        if (url.includes('/_apis/wit/workitems/102') && url.includes('api-version=7.1')) {
          expect(url).toContain('$expand=Relations');
          expect(url).toContain(`asOf=${encodeURIComponent(asOfIso)}`);
          expect(url).not.toContain('fields=');
          return {
            id: 102,
            rev: 4,
            fields: {
              'System.WorkItemType': 'Bug',
              'System.Title': 'Bug 102',
              'System.State': 'Closed',
              'System.AreaPath': 'Proj\\Area',
              'System.IterationPath': 'Proj\\Iter',
              'System.ChangedDate': '2025-12-31T10:00:00Z',
            },
            relations: [],
          };
        }
        throw new Error(`unexpected URL: ${url}`);
      },
    );

    const result = await provider.GetHistoricalQueryResults('q-batch-fallback', project, asOfIso);

    expect(result.total).toBe(2);
    expect(result.rows.map((row: any) => row.id)).toEqual([101, 102]);
  });

  it('GetHistoricalQueryResults falls back to WIQL-by-id when inline WIQL fails', async () => {
    const asOfIso = '2026-01-01T10:00:00.000Z';
    (TFSServices.getItemContent as jest.Mock).mockImplementation(
      async (url: string, _pat: string, method?: string, data?: any) => {
        if (url.includes('/_apis/wit/queries/q-inline-fallback') && url.includes('api-version=7.1')) {
          return { name: 'Inline Fallback Q', wiql: 'SELECT [System.Id] FROM WorkItems' };
        }
        if (url.includes('/_apis/wit/wiql?') && url.includes('api-version=7.1') && method === 'post') {
          expect(String(data?.query || '')).toContain(`ASOF '${asOfIso}'`);
          throw {
            response: {
              status: 500,
              data: { message: 'inline wiql failed' },
            },
          };
        }
        if (url.includes('/_apis/wit/wiql/q-inline-fallback') && url.includes('api-version=7.1')) {
          expect(url).toContain(`asOf=${encodeURIComponent(asOfIso)}`);
          return { workItems: [{ id: 3001 }] };
        }
        if (url.includes('/_apis/wit/workitemsbatch') && url.includes('api-version=7.1') && method === 'post') {
          return {
            value: [
              {
                id: 3001,
                rev: 1,
                fields: {
                  'System.WorkItemType': 'Requirement',
                  'System.Title': 'Req 3001',
                  'System.State': 'Active',
                  'System.AreaPath': 'Proj\\Area',
                  'System.IterationPath': 'Proj\\Iter',
                  'System.ChangedDate': '2025-12-31T10:00:00Z',
                },
                relations: [],
              },
            ],
          };
        }
        throw new Error(`unexpected URL: ${url}`);
      },
    );

    const result = await provider.GetHistoricalQueryResults('q-inline-fallback', project, asOfIso);

    expect(result.total).toBe(1);
    expect(result.rows[0]).toEqual(
      expect.objectContaining({
        id: 3001,
        workItemType: 'Requirement',
        title: 'Req 3001',
      }),
    );
  });

  it('CompareHistoricalQueryResults marks Added/Deleted/Changed/No changes using noise-control fields', async () => {
    const baselineIso = '2025-12-22T17:08:00.000Z';
    const compareIso = '2025-12-28T08:57:00.000Z';

    const baselineBatch = {
      value: [
        {
          id: 11,
          rev: 2,
          fields: {
            'System.WorkItemType': 'Requirement',
            'System.Title': 'Req A',
            'System.State': 'Active',
            'System.Description': 'Old desc',
            'Elisra.TestPhase': 'FAT',
            'System.ChangedDate': baselineIso,
          },
          relations: [],
        },
        {
          id: 23,
          rev: 1,
          fields: {
            'System.WorkItemType': 'Test Case',
            'System.Title': 'Case B',
            'System.State': 'Active',
            'System.Description': 'Case desc',
            'Microsoft.VSTS.TCM.Steps': '<steps>1</steps>',
            'Elisra.TestPhase': 'FAT',
            'System.ChangedDate': baselineIso,
          },
          relations: [{ id: 'l-1' }],
        },
        {
          id: 58,
          rev: 2,
          fields: {
            'System.WorkItemType': 'Bug',
            'System.Title': 'Deleted bug',
            'System.State': 'New',
            'System.Description': 'x',
            'System.ChangedDate': baselineIso,
          },
          relations: [],
        },
        {
          id: 813,
          rev: 3,
          fields: {
            'System.WorkItemType': 'Bug',
            'System.Title': 'No Change bug',
            'System.State': 'Active',
            'System.Description': 'same',
            'System.ChangedDate': baselineIso,
          },
          relations: [],
        },
      ],
    };

    const compareBatch = {
      value: [
        {
          id: 11,
          rev: 20,
          fields: {
            'System.WorkItemType': 'Requirement',
            'System.Title': 'Req A',
            'System.State': 'Active',
            'System.Description': 'New desc',
            'Elisra.TestPhase': 'FAT; ATP',
            'System.ChangedDate': compareIso,
          },
          relations: [],
        },
        {
          id: 23,
          rev: 3,
          fields: {
            'System.WorkItemType': 'Test Case',
            'System.Title': 'Case B',
            'System.State': 'Active',
            'System.Description': 'Case desc',
            'Microsoft.VSTS.TCM.Steps': '<steps>2</steps>',
            'Elisra.TestPhase': 'FAT',
            'System.ChangedDate': compareIso,
          },
          relations: [{ id: 'l-1' }, { id: 'l-2' }],
        },
        {
          id: 814,
          rev: 1,
          fields: {
            'System.WorkItemType': 'Bug',
            'System.Title': 'Added bug',
            'System.State': 'New',
            'System.Description': 'new',
            'System.ChangedDate': compareIso,
          },
          relations: [],
        },
        {
          id: 813,
          rev: 9,
          fields: {
            'System.WorkItemType': 'Bug',
            'System.Title': 'No Change bug',
            'System.State': 'Active',
            'System.Description': 'same',
            'System.ChangedDate': compareIso,
          },
          relations: [],
        },
      ],
    };

    (TFSServices.getItemContent as jest.Mock).mockImplementation(
      async (url: string, _pat: string, method?: string, data?: any) => {
        if (url.includes('/_apis/wit/queries/q-compare') && url.includes('api-version=7.1')) {
          return { name: 'Compare Query', wiql: 'SELECT [System.Id] FROM WorkItems' };
        }
        if (url.includes('/_apis/wit/wiql?') && url.includes('api-version=7.1') && method === 'post') {
          const query = String(data?.query || '');
          if (query.includes(baselineIso)) {
            return { workItems: [{ id: 11 }, { id: 23 }, { id: 58 }, { id: 813 }] };
          }
          if (query.includes(compareIso)) {
            return { workItems: [{ id: 11 }, { id: 23 }, { id: 814 }, { id: 813 }] };
          }
        }
        if (url.includes('/_apis/wit/workitemsbatch') && method === 'post' && data?.asOf === baselineIso) {
          return baselineBatch;
        }
        if (url.includes('/_apis/wit/workitemsbatch') && method === 'post' && data?.asOf === compareIso) {
          return compareBatch;
        }
        throw new Error(`unexpected URL: ${url}`);
      },
    );

    const result = await provider.CompareHistoricalQueryResults(
      'q-compare',
      project,
      baselineIso,
      compareIso,
    );

    const byId = new Map<number, any>(result.rows.map((row: any) => [row.id, row]));
    expect(byId.get(11)?.compareStatus).toBe('Changed');
    expect(byId.get(11)?.changedFields).toEqual(expect.arrayContaining(['Description', 'Test Phase']));
    expect(byId.get(23)?.compareStatus).toBe('Changed');
    expect(byId.get(23)?.changedFields).toEqual(expect.arrayContaining(['Steps', 'Related Link Count']));
    expect(byId.get(58)?.compareStatus).toBe('Deleted');
    expect(byId.get(814)?.compareStatus).toBe('Added');
    expect(byId.get(813)?.compareStatus).toBe('No changes');
    expect(result.summary).toEqual({
      addedCount: 1,
      deletedCount: 1,
      changedCount: 2,
      noChangeCount: 1,
      updatedCount: 2,
    });
  });
});
