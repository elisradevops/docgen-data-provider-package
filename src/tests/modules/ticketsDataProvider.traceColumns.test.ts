import { TFSServices } from '../../helpers/tfs';
import TicketsDataProvider from '../../modules/TicketsDataProvider';
import logger from '../../utils/logger';

jest.mock('../../helpers/tfs');
jest.mock('../../utils/logger');

const mockGetItemContent = TFSServices.getItemContent as jest.Mock;

// Minimal OneHop query result shape
const makeQueryResult = (columns: any[], relations: any[]) => ({
  columns,
  workItemRelations: relations,
  queryType: 'oneHop',
});

// Root relation (source = the root WI, no .source field)
const rootRelation = (targetId: number, targetUrl: string) => ({
  source: null,
  target: { id: targetId, url: targetUrl },
});

// Link relation (source WI → target WI)
const linkRelation = (sourceId: number, targetId: number, targetUrl: string) => ({
  source: { id: sourceId },
  target: { id: targetId, url: targetUrl },
});

const REQ_COLS = [
  { referenceName: 'System.Id', name: 'ID' },
  { referenceName: 'System.Title', name: 'Title' },
  { referenceName: 'Custom.CustomerRequirementId', name: 'CustomerRequirementId' },
  { referenceName: 'Microsoft.VSTS.Common.Priority', name: 'Priority' },
];

const TC_COLS = [
  { referenceName: 'System.Id', name: 'ID' },
  { referenceName: 'System.Title', name: 'Title' },
  { referenceName: 'Microsoft.VSTS.TCM.AutomationStatus', name: 'Automation Status' },
  { referenceName: 'Microsoft.VSTS.Common.Priority', name: 'Priority' },
];

const REQ_FIELDS_SCHEMA = {
  value: [
    { referenceName: 'System.Id' },
    { referenceName: 'System.Title' },
    { referenceName: 'Custom.CustomerRequirementId' },
    { referenceName: 'Microsoft.VSTS.Common.Priority' },
  ],
};

const TC_FIELDS_SCHEMA = {
  value: [
    { referenceName: 'System.Id' },
    { referenceName: 'System.Title' },
    { referenceName: 'Microsoft.VSTS.TCM.AutomationStatus' },
    { referenceName: 'Microsoft.VSTS.Common.Priority' },
    // Note: Custom.CustomerRequirementId is NOT in TC schema
  ],
};

describe('TicketsDataProvider.GetTraceColumnsByType', () => {
  let provider: TicketsDataProvider;

  beforeEach(() => {
    jest.clearAllMocks();
    provider = new TicketsDataProvider('https://dev.azure.com/org/', 'mock-token');
  });

  it('splits columns correctly: CustomerRequirementId in Requirement only, not Test Case', async () => {
    const reqRelations = [
      rootRelation(1, 'https://ado/wi/1'),
      linkRelation(1, 101, 'https://ado/wi/101'),
    ];
    const reqQueryResult = makeQueryResult(REQ_COLS, reqRelations);

    mockGetItemContent
      .mockResolvedValueOnce(reqQueryResult)                          // reqTest wiql fetch
      .mockResolvedValueOnce({ fields: { 'System.WorkItemType': 'Requirement' } }) // source WI sample
      .mockResolvedValueOnce({ fields: { 'System.WorkItemType': 'Test Case' } })   // target WI sample
      .mockResolvedValueOnce(REQ_FIELDS_SCHEMA)                       // Requirement type fields
      .mockResolvedValueOnce(TC_FIELDS_SCHEMA);                       // Test Case type fields

    const result = await provider.GetTraceColumnsByType(
      'https://ado/wiql/reqtest',
      undefined,
      'my-project',
    );

    const reqRefs = result['req-test']!.Requirement.map((c: any) => c.referenceName);
    const tcRefs = result['req-test']!['Test Case'].map((c: any) => c.referenceName);

    expect(reqRefs).toContain('Custom.CustomerRequirementId');
    expect(tcRefs).not.toContain('Custom.CustomerRequirementId');
    expect(reqRefs).toContain('Microsoft.VSTS.Common.Priority');
    expect(tcRefs).toContain('Microsoft.VSTS.Common.Priority');
  });

  it('resolves each query independently: req-test and test-req have separate column sets', async () => {
    const reqRelations = [rootRelation(1, 'https://ado/wi/1'), linkRelation(1, 101, 'https://ado/wi/101')];
    const testRelations = [rootRelation(101, 'https://ado/wi/101'), linkRelation(101, 1, 'https://ado/wi/1')];

    const extraCol = { referenceName: 'System.AreaPath', name: 'Area Path' };
    const tcSchemaWithAreaPath = { value: [...TC_FIELDS_SCHEMA.value, { referenceName: 'System.AreaPath' }] };

    mockGetItemContent
      // req-test query processing
      .mockResolvedValueOnce(makeQueryResult(REQ_COLS, reqRelations))          // reqTest wiql
      .mockResolvedValueOnce({ fields: { 'System.WorkItemType': 'Requirement' } })  // source sample
      .mockResolvedValueOnce({ fields: { 'System.WorkItemType': 'Test Case' } })    // target sample
      .mockResolvedValueOnce(REQ_FIELDS_SCHEMA)                                     // Req schema
      .mockResolvedValueOnce(tcSchemaWithAreaPath)                                  // TC schema
      // test-req query processing (srcIsReq=false: source=TC, target=Req)
      .mockResolvedValueOnce(makeQueryResult([...TC_COLS, extraCol], testRelations)) // testReq wiql
      .mockResolvedValueOnce({ fields: { 'System.WorkItemType': 'Test Case' } })     // source sample (TC)
      .mockResolvedValueOnce({ fields: { 'System.WorkItemType': 'Requirement' } })   // target sample (Req)
      // Promise.all([fetchTypeFieldSet(reqType), fetchTypeFieldSet(tcType)]): reqType=Requirement first
      .mockResolvedValueOnce(REQ_FIELDS_SCHEMA)                                      // Req schema (first in Promise.all)
      .mockResolvedValueOnce(tcSchemaWithAreaPath);                                  // TC schema (second in Promise.all)

    const result = await provider.GetTraceColumnsByType(
      'https://ado/wiql/reqtest',
      'https://ado/wiql/testreq',
      'my-project',
    );

    // req-test: columns from REQ_COLS only — no extraCol
    const reqTestTcRefs = result['req-test']!['Test Case'].map((c: any) => c.referenceName);
    expect(reqTestTcRefs).not.toContain('System.AreaPath');
    // test-req: columns from TC_COLS+extraCol — System.AreaPath present
    const testReqTcRefs = result['test-req']!['Test Case'].map((c: any) => c.referenceName);
    expect(testReqTcRefs).toContain('System.AreaPath');
  });

  it('raw display names passed through without rename (no CustomerRequirementId → Customer ID)', async () => {
    const reqRelations = [rootRelation(1, 'https://ado/wi/1'), linkRelation(1, 101, 'https://ado/wi/101')];

    mockGetItemContent
      .mockResolvedValueOnce(makeQueryResult(REQ_COLS, reqRelations))
      .mockResolvedValueOnce({ fields: { 'System.WorkItemType': 'Requirement' } })
      .mockResolvedValueOnce({ fields: { 'System.WorkItemType': 'Test Case' } })
      .mockResolvedValueOnce(REQ_FIELDS_SCHEMA)
      .mockResolvedValueOnce(TC_FIELDS_SCHEMA);

    const result = await provider.GetTraceColumnsByType('https://ado/wiql/reqtest', undefined, 'proj');

    const custCol = result['req-test']!.Requirement.find((c: any) => c.referenceName === 'Custom.CustomerRequirementId');
    expect(custCol?.name).toBe('CustomerRequirementId'); // raw ADO name, NOT 'Customer ID'
  });

  it('gracefully degrades when WIT type cannot be sampled (no relations) — returns merged columns both sides', async () => {
    const emptyRelations: any[] = [];

    mockGetItemContent.mockResolvedValueOnce(makeQueryResult(REQ_COLS, emptyRelations));

    const result = await provider.GetTraceColumnsByType('https://ado/wiql/reqtest', undefined, 'proj');

    const reqRefs = result['req-test']!.Requirement.map((c: any) => c.referenceName);
    const tcRefs = result['req-test']!['Test Case'].map((c: any) => c.referenceName);
    // Both sides get this query's declared columns as fallback (no sampling possible)
    expect(reqRefs).toContain('Custom.CustomerRequirementId');
    expect(tcRefs).toContain('Custom.CustomerRequirementId');
  });

  it('returns empty arrays when wiql fetch fails', async () => {
    mockGetItemContent.mockRejectedValueOnce(new Error('Network error'));

    const result = await provider.GetTraceColumnsByType('https://ado/wiql/reqtest', undefined, 'proj');

    expect(result).toEqual({ 'req-test': { Requirement: [], 'Test Case': [] } });
    expect((logger.error as jest.Mock)).toHaveBeenCalled();
  });

  it('returns empty arrays when schema fetch fails', async () => {
    const reqRelations = [rootRelation(1, 'https://ado/wi/1'), linkRelation(1, 101, 'https://ado/wi/101')];

    mockGetItemContent
      .mockResolvedValueOnce(makeQueryResult(REQ_COLS, reqRelations))
      .mockResolvedValueOnce({ fields: { 'System.WorkItemType': 'Requirement' } })
      .mockResolvedValueOnce({ fields: { 'System.WorkItemType': 'Test Case' } })
      .mockRejectedValueOnce(new Error('ADO 404')); // schema fetch fails

    const result = await provider.GetTraceColumnsByType('https://ado/wiql/reqtest', undefined, 'proj');

    expect(result).toEqual({ 'req-test': { Requirement: [], 'Test Case': [] } });
  });

  it('handles undefined schema value gracefully (no crash on null response)', async () => {
    const reqRelations = [rootRelation(1, 'https://ado/wi/1'), linkRelation(1, 101, 'https://ado/wi/101')];

    mockGetItemContent
      .mockResolvedValueOnce(makeQueryResult(REQ_COLS, reqRelations))
      .mockResolvedValueOnce({ fields: { 'System.WorkItemType': 'Requirement' } })
      .mockResolvedValueOnce({ fields: { 'System.WorkItemType': 'Test Case' } })
      .mockResolvedValueOnce(undefined)   // schema returns undefined (null guard fix)
      .mockResolvedValueOnce(undefined);

    const result = await provider.GetTraceColumnsByType('https://ado/wiql/reqtest', undefined, 'proj');

    // Both field sets are empty → both sides empty after intersection
    expect(result['req-test']!.Requirement).toEqual([]);
    expect(result['req-test']!['Test Case']).toEqual([]);
  });
});
