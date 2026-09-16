import { TFSServices } from '../../helpers/tfs';
import TicketsDataProvider from '../../modules/TicketsDataProvider';
import logger from '../../utils/logger';

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
      expect.stringContaining(
        `/${project}/_apis/wit/queries/Shared%20Queries?$depth=2&$expand=all&api-version=7.1`,
      ),
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
      expect.stringContaining(
        `/${project}/_apis/wit/queries/Shared%20Queries?$depth=2&$expand=all&api-version=7.1`,
      ),
      token,
    );
    expect(TFSServices.getItemContent).toHaveBeenCalledWith(
      expect.stringContaining(
        `/${project}/_apis/wit/queries/Shared%20Queries?$depth=2&$expand=all&api-version=5.1`,
      ),
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
      expect.stringContaining(
        `/${project}/_apis/wit/queries/Shared%20Queries?$depth=2&$expand=all&api-version=7.1`,
      ),
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
        if (
          url.includes('/_apis/wit/workitemsbatch') &&
          url.includes('api-version=7.1') &&
          method === 'post'
        ) {
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
        if (
          url.includes('/_apis/wit/workitemsbatch') &&
          url.includes('api-version=5.1') &&
          method === 'post'
        ) {
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
    // Two chunks (200 + 5 ids), two batched passes per chunk (fields-only,
    // then $expand=Relations-only) — never both in the same request, since
    // ADO rejects that combination.
    expect(batchCalls).toHaveLength(4);
    batchCalls.forEach((call) => {
      const payload = call[3];
      expect(payload.fields && payload.$expand).toBeUndefined();
      expect(payload.errorPolicy).toBe('Omit');
    });
    const fieldsCalls = batchCalls.filter((call) => Array.isArray(call[3]?.fields));
    const relationsCalls = batchCalls.filter((call) => call[3]?.$expand === 'Relations');
    expect(fieldsCalls).toHaveLength(2);
    expect(relationsCalls).toHaveLength(2);
    const byLength = (a: number, b: number) => a - b;
    expect(fieldsCalls.map((call) => call[3].ids.length).sort(byLength)).toEqual([5, 200]);
    expect(relationsCalls.map((call) => call[3].ids.length).sort(byLength)).toEqual([5, 200]);
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
        if (
          url.includes('/_apis/wit/workitemsbatch') &&
          url.includes('api-version=7.1') &&
          method === 'post'
        ) {
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

  it('GetHistoricalQueryResults skips work items that do not exist at as-of time and logs warning', async () => {
    const asOfIso = '2026-01-01T10:00:00.000Z';
    (TFSServices.getItemContent as jest.Mock).mockImplementation(
      async (url: string, _pat: string, method?: string, data?: any) => {
        if (url.includes('/_apis/wit/queries/q-missing-asof') && url.includes('api-version=7.1')) {
          return { name: 'Missing AsOf Q', wiql: 'SELECT [System.Id] FROM WorkItems' };
        }
        if (url.includes('/_apis/wit/wiql?') && url.includes('api-version=7.1') && method === 'post') {
          expect(String(data?.query || '')).toContain(`ASOF '${asOfIso}'`);
          return { workItems: [{ id: 101 }, { id: 102 }] };
        }
        if (url.includes('/_apis/wit/workitemsbatch') && method === 'post') {
          throw {
            response: {
              status: 500,
              data: { message: 'workitemsbatch failed' },
            },
          };
        }
        if (url.includes('/_apis/wit/workitems/101')) {
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
        if (url.includes('/_apis/wit/workitems/102')) {
          throw {
            response: {
              status: 404,
              data: {
                message:
                  'The work item 102 does not exist at time 12/31/2025 11:56:00 PM. It might have been deleted.',
              },
            },
          };
        }
        throw new Error(`unexpected URL: ${url}`);
      },
    );

    const result = await provider.GetHistoricalQueryResults('q-missing-asof', project, asOfIso);

    expect(result.total).toBe(1);
    expect(result.skippedWorkItemsCount).toBe(1);
    expect(result.rows.map((row: any) => row.id)).toEqual([101]);
    expect(
      (logger.warn as jest.Mock).mock.calls.some((call) =>
        String(call[0]).includes('skipping work item 102'),
      ),
    ).toBe(true);
  });

  it('GetHistoricalQueryResults returns empty result when all work items are missing at as-of time', async () => {
    const asOfIso = '2026-01-01T10:00:00.000Z';
    (TFSServices.getItemContent as jest.Mock).mockImplementation(
      async (url: string, _pat: string, method?: string, data?: any) => {
        if (url.includes('/_apis/wit/queries/q-all-missing-asof') && url.includes('api-version=7.1')) {
          return { name: 'All Missing AsOf Q', wiql: 'SELECT [System.Id] FROM WorkItems' };
        }
        if (url.includes('/_apis/wit/wiql?') && method === 'post') {
          expect(String(data?.query || '')).toContain(`ASOF '${asOfIso}'`);
          return { workItems: [{ id: 201 }, { id: 202 }] };
        }
        if (url.includes('/_apis/wit/workitemsbatch') && method === 'post') {
          throw {
            response: {
              status: 500,
              data: { message: 'workitemsbatch failed' },
            },
          };
        }
        if (url.includes('/_apis/wit/workitems/201') || url.includes('/_apis/wit/workitems/202')) {
          const id = url.includes('/201') ? 201 : 202;
          throw {
            response: {
              status: 404,
              data: {
                message: `The work item ${id} does not exist at time 12/31/2025 11:56:00 PM.`,
              },
            },
          };
        }
        throw new Error(`unexpected URL: ${url}`);
      },
    );

    const result = await provider.GetHistoricalQueryResults('q-all-missing-asof', project, asOfIso);

    expect(result.total).toBe(0);
    expect(result.rows).toEqual([]);
    expect(result.skippedWorkItemsCount).toBe(2);
  });

  it('GetHistoricalQueryResults still throws for non-missing historical errors', async () => {
    const asOfIso = '2026-01-01T10:00:00.000Z';
    (TFSServices.getItemContent as jest.Mock).mockImplementation(
      async (url: string, _pat: string, method?: string, data?: any) => {
        if (url.includes('/_apis/wit/queries/q-auth-error') && url.includes('api-version=7.1')) {
          return { name: 'Auth Error Q', wiql: 'SELECT [System.Id] FROM WorkItems' };
        }
        if (url.includes('/_apis/wit/wiql?') && method === 'post') {
          expect(String(data?.query || '')).toContain(`ASOF '${asOfIso}'`);
          return { workItems: [{ id: 901 }] };
        }
        if (url.includes('/_apis/wit/workitemsbatch') && method === 'post') {
          throw {
            response: {
              status: 500,
              data: { message: 'workitemsbatch failed' },
            },
          };
        }
        if (url.includes('/_apis/wit/workitems/901')) {
          throw {
            response: {
              status: 401,
              data: { message: 'Unauthorized' },
            },
          };
        }
        throw new Error(`unexpected URL: ${url}`);
      },
    );

    await expect(provider.GetHistoricalQueryResults('q-auth-error', project, asOfIso)).rejects.toEqual(
      expect.objectContaining({
        response: expect.objectContaining({ status: 401 }),
      }),
    );
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
        if (
          url.includes('/_apis/wit/workitemsbatch') &&
          url.includes('api-version=7.1') &&
          method === 'post'
        ) {
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

  it('CompareHistoricalQueryResults supports missing work items on one side and reports Added/Deleted', async () => {
    const baselineIso = '2025-12-20T00:00:00.000Z';
    const compareIso = '2025-12-30T00:00:00.000Z';

    (TFSServices.getItemContent as jest.Mock).mockImplementation(
      async (url: string, _pat: string, method?: string, data?: any) => {
        if (url.includes('/_apis/wit/queries/q-compare-missing-side') && url.includes('api-version=7.1')) {
          return { name: 'Compare Missing Side', wiql: 'SELECT [System.Id] FROM WorkItems' };
        }
        if (url.includes('/_apis/wit/wiql?') && method === 'post') {
          const query = String(data?.query || '');
          if (query.includes(baselineIso)) {
            return { workItems: [{ id: 1001 }, { id: 1002 }, { id: 1003 }] };
          }
          if (query.includes(compareIso)) {
            return { workItems: [{ id: 1001 }, { id: 1002 }, { id: 1003 }] };
          }
        }
        if (url.includes('/_apis/wit/workitemsbatch') && method === 'post') {
          throw {
            response: {
              status: 500,
              data: { message: 'workitemsbatch failed' },
            },
          };
        }
        if (url.includes('/_apis/wit/workitems/1001')) {
          return {
            id: 1001,
            rev: 1,
            fields: {
              'System.WorkItemType': 'Requirement',
              'System.Title': 'Stable Item',
              'System.State': 'Active',
              'System.ChangedDate': compareIso,
            },
            relations: [],
          };
        }
        if (
          url.includes(`asOf=${encodeURIComponent(baselineIso)}`) &&
          url.includes('/_apis/wit/workitems/1002')
        ) {
          throw {
            response: {
              status: 404,
              data: { message: 'The work item 1002 does not exist at time 12/20/2025 12:00:00 AM.' },
            },
          };
        }
        if (
          url.includes(`asOf=${encodeURIComponent(compareIso)}`) &&
          url.includes('/_apis/wit/workitems/1002')
        ) {
          return {
            id: 1002,
            rev: 4,
            fields: {
              'System.WorkItemType': 'Bug',
              'System.Title': 'Added Later',
              'System.State': 'New',
              'System.ChangedDate': compareIso,
            },
            relations: [],
          };
        }
        if (
          url.includes(`asOf=${encodeURIComponent(baselineIso)}`) &&
          url.includes('/_apis/wit/workitems/1003')
        ) {
          return {
            id: 1003,
            rev: 2,
            fields: {
              'System.WorkItemType': 'Bug',
              'System.Title': 'Deleted Later',
              'System.State': 'Active',
              'System.ChangedDate': baselineIso,
            },
            relations: [],
          };
        }
        if (
          url.includes(`asOf=${encodeURIComponent(compareIso)}`) &&
          url.includes('/_apis/wit/workitems/1003')
        ) {
          throw {
            response: {
              status: 404,
              data: { message: 'The work item 1003 does not exist at time 12/30/2025 12:00:00 AM.' },
            },
          };
        }
        throw new Error(`unexpected URL: ${url}`);
      },
    );

    const result = await provider.CompareHistoricalQueryResults(
      'q-compare-missing-side',
      project,
      baselineIso,
      compareIso,
    );

    const byId = new Map<number, any>(result.rows.map((row: any) => [row.id, row]));
    expect(byId.get(1001)?.compareStatus).toBe('No changes');
    expect(byId.get(1002)?.compareStatus).toBe('Added');
    expect(byId.get(1003)?.compareStatus).toBe('Deleted');
    expect(result.skippedWorkItems).toEqual(
      expect.objectContaining({ baselineCount: 1, compareToCount: 1, totalDistinct: 2 }),
    );
  });

  it('CompareHistoricalQueryResults excludes work items missing at both dates and can return empty set', async () => {
    const baselineIso = '2025-01-01T00:00:00.000Z';
    const compareIso = '2025-01-02T00:00:00.000Z';

    (TFSServices.getItemContent as jest.Mock).mockImplementation(
      async (url: string, _pat: string, method?: string, data?: any) => {
        if (url.includes('/_apis/wit/queries/q-compare-all-missing') && url.includes('api-version=7.1')) {
          return { name: 'Compare All Missing', wiql: 'SELECT [System.Id] FROM WorkItems' };
        }
        if (url.includes('/_apis/wit/wiql?') && method === 'post') {
          const query = String(data?.query || '');
          if (query.includes(baselineIso) || query.includes(compareIso)) {
            return { workItems: [{ id: 777 }] };
          }
        }
        if (url.includes('/_apis/wit/workitemsbatch') && method === 'post') {
          throw {
            response: {
              status: 500,
              data: { message: 'workitemsbatch failed' },
            },
          };
        }
        if (url.includes('/_apis/wit/workitems/777')) {
          throw {
            response: {
              status: 404,
              data: { message: 'The work item 777 does not exist at time 01/01/2025 12:00:00 AM.' },
            },
          };
        }
        throw new Error(`unexpected URL: ${url}`);
      },
    );

    const result = await provider.CompareHistoricalQueryResults(
      'q-compare-all-missing',
      project,
      baselineIso,
      compareIso,
    );

    expect(result.rows).toEqual([]);
    expect(result.summary).toEqual({
      addedCount: 0,
      deletedCount: 0,
      changedCount: 0,
      noChangeCount: 0,
      updatedCount: 0,
    });
    expect(result.skippedWorkItems).toEqual(
      expect.objectContaining({ baselineCount: 1, compareToCount: 1, totalDistinct: 1 }),
    );
  });

  it('CompareHistoricalQueryResults merges relations from the separate $expand pass into relatedLinkCount', async () => {
    const baselineIso = '2026-01-05T00:00:00.000Z';
    const compareIso = '2026-01-10T00:00:00.000Z';

    // Fields-only and relations-only responses are disjoint (as ADO actually
    // returns them for each pass) to prove the merge-by-id, not just that a
    // duplicated mock happens to satisfy both reads.
    (TFSServices.getItemContent as jest.Mock).mockImplementation(
      async (url: string, _pat: string, method?: string, data?: any) => {
        if (url.includes('/_apis/wit/queries/q-merge-relations') && url.includes('api-version=7.1')) {
          return { name: 'Merge Relations Q', wiql: 'SELECT [System.Id] FROM WorkItems' };
        }
        if (url.includes('/_apis/wit/wiql?') && method === 'post') {
          return { workItems: [{ id: 501 }] };
        }
        if (url.includes('/_apis/wit/workitemsbatch') && method === 'post') {
          expect(data.fields && data.$expand).toBeUndefined();
          expect(data.errorPolicy).toBe('Omit');
          const isRelationsPass = data.$expand === 'Relations';
          const isBaseline = data.asOf === baselineIso;
          if (isRelationsPass) {
            return {
              value: [{ id: 501, relations: isBaseline ? [{ id: 'l-1' }] : [{ id: 'l-1' }, { id: 'l-2' }] }],
            };
          }
          return {
            value: [
              {
                id: 501,
                rev: isBaseline ? 1 : 2,
                fields: {
                  'System.WorkItemType': 'Test Case',
                  'System.Title': 'Merge Case',
                  'System.State': 'Active',
                  'Microsoft.VSTS.TCM.Steps': '<steps>same</steps>',
                  'System.ChangedDate': isBaseline ? baselineIso : compareIso,
                },
              },
            ],
          };
        }
        throw new Error(`unexpected URL: ${url}`);
      },
    );

    const result = await provider.CompareHistoricalQueryResults(
      'q-merge-relations',
      project,
      baselineIso,
      compareIso,
    );

    const row = result.rows.find((r: any) => r.id === 501);
    expect(row?.compareStatus).toBe('Changed');
    expect(row?.changedFields).toEqual(['Related Link Count']);
    const diff = row?.differences.find((d: any) => d.field === 'Related Link Count');
    expect(diff).toEqual({ field: 'Related Link Count', baseline: '1', compareTo: '2' });
  });

  it('GetHistoricalQueryResults treats a workitemsbatch errorPolicy Omit as skipped without a per-item fallback', async () => {
    const asOfIso = '2026-01-15T00:00:00.000Z';

    (TFSServices.getItemContent as jest.Mock).mockImplementation(
      async (url: string, _pat: string, method?: string, data?: any) => {
        if (url.includes('/_apis/wit/queries/q-omit') && url.includes('api-version=7.1')) {
          return { name: 'Omit Q', wiql: 'SELECT [System.Id] FROM WorkItems' };
        }
        if (url.includes('/_apis/wit/wiql?') && method === 'post') {
          return { workItems: [{ id: 601 }, { id: 602 }] };
        }
        if (url.includes('/_apis/wit/workitemsbatch') && method === 'post') {
          expect(data.errorPolicy).toBe('Omit');
          // Work item 602 is inaccessible/deleted and is omitted by ADO from
          // `value` rather than failing the whole batch.
          if (data.$expand === 'Relations') {
            return { value: [{ id: 601, relations: [] }] };
          }
          return {
            value: [
              {
                id: 601,
                rev: 1,
                fields: {
                  'System.WorkItemType': 'Bug',
                  'System.Title': 'Present Bug',
                  'System.State': 'Active',
                  'System.ChangedDate': asOfIso,
                },
              },
            ],
          };
        }
        // A call here would mean the per-item fallback ran despite the batch
        // succeeding, which defeats the point of errorPolicy: 'Omit'.
        throw new Error(`unexpected URL: ${url}`);
      },
    );

    const result = await provider.GetHistoricalQueryResults('q-omit', project, asOfIso);

    expect(result.total).toBe(1);
    expect(result.skippedWorkItemsCount).toBe(1);
    expect(result.rows.map((row: any) => row.id)).toEqual([601]);
    expect(
      (logger.warn as jest.Mock).mock.calls.some((call) => String(call[0]).includes('omitted 1')),
    ).toBe(true);
  });

  it('GetHistoricalQueryResults drops only the specific field ADO names in TF51535, keeping Custom.TestPhase when only Elisra.TestPhase is undefined for this project', async () => {
    const asOfIso = '2026-02-01T00:00:00.000Z';
    // 205 ids -> 2 chunks (200 + 5), both dispatched with the full field list
    // near-simultaneously — this proves the retry is decided per-call, not
    // from the shared override flag, since both chunks fail before either
    // one sets it.
    const allIds = Array.from({ length: 205 }, (_, idx) => idx + 1);

    (TFSServices.getItemContent as jest.Mock).mockImplementation(
      async (url: string, _pat: string, method?: string, data?: any) => {
        if (url.includes('/_apis/wit/queries/q-unknown-field') && url.includes('api-version=7.1')) {
          return { name: 'Unknown Field Q', wiql: 'SELECT [System.Id] FROM WorkItems' };
        }
        if (url.includes('/_apis/wit/wiql?') && method === 'post') {
          return { workItems: allIds.map((id) => ({ id })) };
        }
        if (url.includes('/_apis/wit/workitemsbatch') && method === 'post') {
          if (data?.$expand === 'Relations') {
            return {
              value: (Array.isArray(data?.ids) ? data.ids : []).map((id: number) => ({ id, relations: [] })),
            };
          }
          if (Array.isArray(data?.fields) && data.fields.includes('Elisra.TestPhase')) {
            throw {
              response: {
                status: 400,
                data: { message: "TF51535: Cannot find field 'Elisra.TestPhase'." },
              },
            };
          }
          // Custom.TestPhase is the real field on this project and must
          // still be requested — only the unrecognized alias should be
          // dropped, not every optional field.
          expect(data.fields).toEqual(expect.arrayContaining(['Custom.TestPhase']));
          expect(data.fields).not.toEqual(expect.arrayContaining(['Elisra.TestPhase']));
          return {
            value: (Array.isArray(data?.ids) ? data.ids : []).map((id: number) => ({
              id,
              rev: 1,
              fields: {
                'System.WorkItemType': 'Test Case',
                'System.Title': `Test ${id}`,
                'System.State': 'Active',
                'System.ChangedDate': asOfIso,
                'Custom.TestPhase': 'FAT',
              },
            })),
          };
        }
        throw new Error(`unexpected URL: ${url}`);
      },
    );

    const result = await provider.GetHistoricalQueryResults('q-unknown-field', project, asOfIso);

    // Every id came back — nothing fell to the 2xN per-item fallback, which
    // would have thrown on any call not covered by the handler above.
    expect(result.total).toBe(205);
    const fieldsCallsWithElisraAlias = (TFSServices.getItemContent as jest.Mock).mock.calls.filter(
      (call) => Array.isArray(call[3]?.fields) && call[3].fields.includes('Elisra.TestPhase'),
    );
    // Both chunks start with the full field list before the override is set,
    // so both legitimately fail once and both must retry — not just one.
    expect(fieldsCallsWithElisraAlias).toHaveLength(2);
    expect(
      (logger.warn as jest.Mock).mock.calls.some((call) =>
        String(call[0]).includes("field 'Elisra.TestPhase' not defined for this project"),
      ),
    ).toBe(true);
  });
});
