import { TFSServices } from '../helpers/tfs';
import { Workitem, QueryTree } from '../models/tfs-data';
import { Helper, Links, Trace, Relations } from '../helpers/helper';
import { Query } from '../models/tfs-data';
import { QueryType } from '../models/tfs-data';
import { QueryAllTypes } from '../models/tfs-data';
import { Column } from '../models/tfs-data';
import { value } from '../models/tfs-data';

import logger from '../utils/logger';
const pLimit = require('p-limit');

type FallbackFetchOutcome = {
  result: any;
  usedFolder: any;
};

type DocTypeBranchConfig = {
  id: string;
  label: string;
  // Candidate folder names (case-insensitive). If none resolve, fallback starts from `fallbackStart`.
  folderNames?: string[];
  fetcher: (folder: any) => Promise<any>;
  // Optional validator to determine whether the fetch produced any usable queries.
  validator?: (result: any) => boolean;
  // Optional explicit starting folder for the fallback chain.
  fallbackStart?: any;
};

export default class TicketsDataProvider {
  orgUrl: string = '';
  token: string = '';
  queriesList: Array<any> = new Array<any>();
  private limit = pLimit(10);

  constructor(orgUrl: string, token: string) {
    this.orgUrl = orgUrl;
    this.token = token;
  }

  /**
   * Extracts the Azure DevOps project name from a WIQL URL.
   *
   * @param wiqlHref - Full WIQL href (may or may not start with orgUrl)
   * @returns The project name if present; otherwise null
   */
  private getProjectFromWiqlHref(wiqlHref: string): string | null {
    if (!wiqlHref) return null;
    const href = wiqlHref.startsWith(this.orgUrl) ? wiqlHref.substring(this.orgUrl.length) : wiqlHref;
    const beforeApis = href.split('/_apis')[0] || '';
    const segs = beforeApis.split('/').filter(Boolean);
    return segs[0] || null;
  }

  /**
   * Retrieves reference names of fields on the Requirement work item type whose display name
   * indicates a "requirement type" field, ordered by priority.
   * Priority: Microsoft.VSTS.CMMI.RequirementType first (when present), then other matches
   * in the order returned by the API. Falls back to the known ref when no matches are found.
   *
   * @param project - Azure DevOps project name
   * @returns Ordered list of candidate reference names to inspect on each work item
   */
  private async getRequirementTypeFieldRefs(project: string): Promise<string[]> {
    const result: string[] = [];
    try {
      const fieldsUrl = `${this.orgUrl}${project}/_apis/wit/workitemtypes/Requirement/fields`;
      const fieldsResp = await TFSServices.getItemContent(fieldsUrl, this.token);
      const fieldsArr = Array.isArray(fieldsResp?.value) ? fieldsResp.value : [];
      const candidates = fieldsArr
        .filter((f: any) => {
          const nm = String(f?.name || '')
            .toLowerCase()
            .replace(/_/g, ' ');
          return nm.includes('requirement type');
        })
        .map((f: any) => f?.referenceName)
        .filter((x: any) => typeof x === 'string' && x.length > 0);

      const knownRef = 'Microsoft.VSTS.CMMI.RequirementType';
      const unique = new Set<string>();
      if (candidates.includes(knownRef)) {
        result.push(knownRef);
        unique.add(knownRef);
      }
      for (const c of candidates) {
        if (!unique.has(c)) {
          unique.add(c);
          result.push(c);
        }
      }
    } catch (e) {}
    if (result.length === 0) {
      result.push('Microsoft.VSTS.CMMI.RequirementType');
    }
    return result;
  }

  async FetchImageAsBase64(url: string): Promise<string> {
    let image = await TFSServices.fetchAzureDevOpsImageAsBase64(url, this.token, 'get', null);
    return image;
  }

  async GetWorkItem(project: string, id: string): Promise<any> {
    let url = `${this.orgUrl}${project}/_apis/wit/workitems/${id}?$expand=All`;
    return TFSServices.getItemContent(url, this.token);
  }

  async GetWorkItemByUrl(url: string): Promise<any> {
    return TFSServices.getItemContent(url, this.token);
  }

  async GetLinksByIds(project: string, ids: any) {
    var trace: Array<Trace> = new Array<Trace>();
    let wis = await this.PopulateWorkItemsByIds(ids, project);
    let linksMap: any = await this.GetRelationsIds(wis);

    let relations;
    for (let i = 0; i < wis.length; i++) {
      let traceItem: Trace; //= new Trace();
      traceItem = await this.GetParentLink(project, wis[i]);

      if (linksMap.get(wis[i].id)?.rels.length > 0) {
        relations = await this.PopulateWorkItemsByIds(linksMap.get(wis[i].id).rels, project);
        traceItem.links = await this.GetLinks(project, wis[i], relations);
      }
      trace.push(traceItem);
    }
    return trace;
  }

  async GetParentLink(project: string, wi: any) {
    let trace: Trace = new Trace();
    if (wi != null) {
      trace.id = wi.id;
      trace.title = wi.fields['System.Title'];
      trace.url = this.orgUrl + project + '/_workitems/edit/' + wi.id;
      if (wi.fields['System.CustomerId'] != null && wi.fields['System.CustomerId'] != undefined) {
        trace.customerId = wi.fields['System.CustomerId'];
      }
    }
    return trace;
  }
  async GetRelationsIds(ids: any) {
    let rel = new Map<string, Relations>();
    try {
      for (let i = 0; i < ids.length; i++) {
        var link = new Relations();
        link.id = ids[i].id;
        if (ids[i].relations != null)
          for (let j = 0; j < ids[i].relations.length; j++) {
            if (ids[i].relations[j].rel != 'AttachedFile') {
              let index = ids[i].relations[j].url.lastIndexOf('/');
              let id = ids[i].relations[j].url.substring(index + 1);
              link.rels.push(id);
            }
          }
        rel.set(ids[i].id, link);
      }
    } catch (e) {}
    return rel;
  }
  async GetLinks(project: string, wi: any, links: any) {
    var linkList: Array<Links> = new Array<Links>();
    for (let i = 0; i < wi.relations.length; i++) {
      for (let j = 0; j < links.length; j++) {
        let index = wi.relations[i].url.lastIndexOf('/');
        let linkId = wi.relations[i].url.substring(index + 1);
        if (linkId == links[j].id) {
          var link = new Links();
          link.type = wi.relations[i].rel;
          link.id = links[j].id;
          link.title = links[j].fields['System.Title'];
          link.description = links[j].fields['System.Description'];
          link.url = this.orgUrl + project + '/_workitems/edit/' + linkId;
          linkList.push(link);
          break;
        }
      }
    }
    return linkList;
  }
  /**
   * Getting shared queries
   * @param project project name
   * @param path query path
   * @param docType document type
   * @returns
   */
  async GetSharedQueries(project: string, path: string, docType: string = ''): Promise<any> {
    let url;
    try {
      if (path === '')
        url = `${this.orgUrl}${project}/_apis/wit/queries/Shared%20Queries?$depth=2&$expand=all`;
      else url = `${this.orgUrl}${project}/_apis/wit/queries/${path}?$depth=2&$expand=all`;
      let queries: any = await TFSServices.getItemContent(url, this.token);
      logger.debug(`doctype: ${docType}`);
      const normalizedDocType = (docType || '').toLowerCase();
      const queriesWithChildren = await this.ensureQueryChildren(queries);

      switch (normalizedDocType) {
        case 'std': {
          const { root: stdRoot, found: stdRootFound } = await this.getDocTypeRoot(
            queriesWithChildren,
            'std'
          );
          logger.debug(`[GetSharedQueries][std] using ${stdRootFound ? 'dedicated folder' : 'root queries'}`);
          // Each branch describes the dedicated folder names, the fetch routine, and how to validate results.
          const stdBranches = await this.fetchDocTypeBranches(queriesWithChildren, stdRoot, [
            {
              id: 'reqToTest',
              label: '[GetSharedQueries][std][req-to-test]',
              folderNames: [
                'requirement - test',
                'requirement to test case',
                'requirement to test',
                'req to test',
              ],
              fetcher: (folder: any) => this.fetchLinkedReqTestQueries(folder, false),
              validator: (result: any) => this.hasAnyQueryTree(result?.reqTestTree),
            },
            {
              id: 'testToReq',
              label: '[GetSharedQueries][std][test-to-req]',
              folderNames: [
                'test - requirement',
                'test to requirement',
                'test case to requirement',
                'test to req',
              ],
              fetcher: (folder: any) => this.fetchLinkedReqTestQueries(folder, true),
              validator: (result: any) => this.hasAnyQueryTree(result?.testReqTree),
            },
            {
              id: 'mom',
              label: '[GetSharedQueries][std][mom]',
              folderNames: ['linked mom', 'mom'],
              fetcher: (folder: any) => this.fetchLinkedMomQueries(folder),
              validator: (result: any) => this.hasAnyQueryTree(result?.linkedMomTree),
            },
          ]);

          const reqToTestResult = stdBranches['reqToTest'];
          const testToReqResult = stdBranches['testToReq'];
          const momResult = stdBranches['mom'];

          const reqTestQueries = {
            reqTestTree: reqToTestResult?.result?.reqTestTree ?? null,
            testReqTree: testToReqResult?.result?.testReqTree ?? reqToTestResult?.result?.testReqTree ?? null,
          };

          const linkedMomQueries = {
            linkedMomTree: momResult?.result?.linkedMomTree ?? null,
          };

          return { reqTestQueries, linkedMomQueries };
        }
        case 'str': {
          const { root: strRoot, found: strRootFound } = await this.getDocTypeRoot(
            queriesWithChildren,
            'str'
          );
          logger.debug(`[GetSharedQueries][str] using ${strRootFound ? 'dedicated folder' : 'root queries'}`);
          const strBranches = await this.fetchDocTypeBranches(queriesWithChildren, strRoot, [
            {
              id: 'reqToTest',
              label: '[GetSharedQueries][str][req-to-test]',
              folderNames: [
                'requirement - test',
                'requirement to test case',
                'requirement to test',
                'req to test',
              ],
              fetcher: (folder: any) => this.fetchLinkedReqTestQueries(folder, false),
              validator: (result: any) => this.hasAnyQueryTree(result?.reqTestTree),
            },
            {
              id: 'testToReq',
              label: '[GetSharedQueries][str][test-to-req]',
              folderNames: [
                'test - requirement',
                'test to requirement',
                'test case to requirement',
                'test to req',
              ],
              fetcher: (folder: any) => this.fetchLinkedReqTestQueries(folder, true),
              validator: (result: any) => this.hasAnyQueryTree(result?.testReqTree),
            },
            {
              id: 'openPcrToTest',
              label: '[GetSharedQueries][str][open-pcr-to-test]',
              folderNames: ['open pcr to test case', 'open pcr to test', 'open pcr - test', 'open pcr'],
              fetcher: (folder: any) => this.fetchLinkedOpenPcrTestQueries(folder, false),
              validator: (result: any) => this.hasAnyQueryTree(result?.OpenPcrToTestTree),
            },
            {
              id: 'testToOpenPcr',
              label: '[GetSharedQueries][str][test-to-open-pcr]',
              folderNames: ['test case to open pcr', 'test to open pcr', 'test - open pcr', 'open pcr'],
              fetcher: (folder: any) => this.fetchLinkedOpenPcrTestQueries(folder, true),
              validator: (result: any) => this.hasAnyQueryTree(result?.TestToOpenPcrTree),
            },
          ]);

          const strReqToTest = strBranches['reqToTest'];
          const strTestToReq = strBranches['testToReq'];
          const strOpenPcrToTest = strBranches['openPcrToTest'];
          const strTestToOpenPcr = strBranches['testToOpenPcr'];

          const reqTestTrees = {
            reqTestTree: strReqToTest?.result?.reqTestTree ?? null,
            testReqTree: strTestToReq?.result?.testReqTree ?? strReqToTest?.result?.testReqTree ?? null,
          };

          const openPcrTestTrees = {
            OpenPcrToTestTree: strOpenPcrToTest?.result?.OpenPcrToTestTree ?? null,
            TestToOpenPcrTree:
              strTestToOpenPcr?.result?.TestToOpenPcrTree ??
              strOpenPcrToTest?.result?.TestToOpenPcrTree ??
              null,
          };

          return { reqTestTrees, openPcrTestTrees };
        }
        case 'test-reporter': {
          const { root: testReporterRoot, found: testReporterFound } = await this.getDocTypeRoot(
            queriesWithChildren,
            'test-reporter'
          );
          logger.debug(
            `[GetSharedQueries][test-reporter] using ${
              testReporterFound ? 'dedicated folder' : 'root queries'
            }`
          );
          const testReporterBranches = await this.fetchDocTypeBranches(
            queriesWithChildren,
            testReporterRoot,
            [
              {
                id: 'testReporter',
                label: '[GetSharedQueries][test-reporter]',
                folderNames: ['test reporter', 'test-reporter'],
                fetcher: (folder: any) => this.fetchTestReporterQueries(folder),
                validator: (result: any) => this.hasAnyQueryTree(result?.testAssociatedTree),
              },
            ]
          );
          const testReporterFetch = testReporterBranches['testReporter'];
          return testReporterFetch?.result ?? { testAssociatedTree: null };
        }
        case 'srs':
          return await this.fetchSrsQueries(queriesWithChildren);
        case 'svd': {
          const { root: svdRoot, found } = await this.getDocTypeRoot(queriesWithChildren, 'svd');
          if (!found) {
            logger.debug('[GetSharedQueries][svd] dedicated folder not found, using fallback tree');
          }
          const svdBranches = await this.fetchDocTypeBranches(queriesWithChildren, svdRoot, [
            {
              id: 'systemOverview',
              label: '[GetSharedQueries][svd][system-overview]',
              folderNames: ['system overview'],
              fetcher: async (folder: any) => {
                const { tree1 } = await this.structureAllQueryPath(folder);
                return tree1;
              },
              validator: (result: any) => !!result,
            },
            {
              id: 'knownBugs',
              label: '[GetSharedQueries][svd][known-bugs]',
              folderNames: ['known bugs', 'known bug'],
              fetcher: async (folder: any) => {
                const { tree2 } = await this.structureAllQueryPath(folder);
                return tree2;
              },
              validator: (result: any) => !!result,
            },
          ]);

          const systemOverviewFetch = svdBranches['systemOverview'];
          const knownBugsFetch = svdBranches['knownBugs'];

          return {
            systemOverviewQueryTree: systemOverviewFetch?.result ?? null,
            knownBugsQueryTree: knownBugsFetch?.result ?? null,
          };
        }
        default:
          break;
      }
    } catch (err: any) {
      logger.error(`Error occurred during fetching shared queries: ${err.message}`);
      throw err;
    }
  }

  /**
   * Fetches the fields of a specific work item type.
   *
   * @param project - The project name.
   * @param itemType - The work item type.
   * @returns An array of objects containing the field name and reference name.
   */

  async GetFieldsByType(project: string, itemType: string) {
    try {
      let url = `${this.orgUrl}${project}/_apis/wit/workitemtypes/${itemType}/fields`;
      const { value: fields } = await TFSServices.getItemContent(url, this.token);
      return (
        fields
          // filter out the fields that are not relevant for the user
          .filter(
            (field: any) =>
              field.name !== 'ID' &&
              field.name !== 'Title' &&
              field.name !== 'Description' &&
              field.name !== 'Work Item Type' &&
              field.name !== 'Steps'
          )
          .map((field: any) => {
            return {
              text: `${field.name} (${itemType})`,
              key: field.referenceName,
            };
          })
      );
    } catch (err: any) {
      logger.error(`Error occurred during fetching fields by type: ${err.message}`);
      throw err;
    }
  }

  /**
   * fetches linked queries
   * @param queries fetched queries
   * @param onlyTestReq get only test req
   * @returns ReqTestTree and TestReqTree
   */
  private async fetchLinkedReqTestQueries(queries: any, onlyTestReq: boolean = false) {
    const { tree1: reqTestTree, tree2: testReqTree } = await this.structureFetchedQueries(
      queries,
      onlyTestReq,
      null,
      ['Requirement'],
      ['Test Case']
    );
    return { reqTestTree, testReqTree };
  }

  /**
   * fetches linked mom queries
   * @param queries fetched queries
   * @returns linkedMomTree
   */

  private async fetchLinkedMomQueries(queries: any) {
    const { tree1: linkedMomTree } = await this.structureFetchedQueries(
      queries,
      false,
      null,
      ['Test Case'],
      [
        'Task',
        'Bug',
        'Code Review Request',
        'Change Request',
        'Code Review Response',
        'Epic',
        'Feature',
        'User Story',
        'Feedback Request',
        'Feedback Response',
        'Issue',
        'Risk',
        'Review',
        'Test Plan',
        'Test Suite',
      ]
    );
    return { linkedMomTree };
  }

  private hasAnyQueryTree(result: any): boolean {
    const inspect = (value: any): boolean => {
      if (!value) {
        return false;
      }

      if (Array.isArray(value)) {
        return value.some(inspect);
      }

      if (typeof value === 'object') {
        if (value.isValidQuery || value.wiql || value.queryType) {
          return true;
        }

        if ('roots' in value && Array.isArray(value.roots) && value.roots.length > 0) {
          return true;
        }

        if ('children' in value && Array.isArray(value.children) && value.children.length > 0) {
          return true;
        }

        return Object.values(value).some(inspect);
      }

      return false;
    };

    return inspect(result);
  }

  /**
   * Fetches and structures linked queries related to open PCR (Problem Change Request) tests.
   *
   * This method retrieves and organizes the relationships between "Test Case" and
   * other entities such as "Bug" and "Change Request" into two tree structures.
   *
   * @param queries - The input queries to be processed and structured.
   * @param onlySourceSide - A flag indicating whether to process only the source side of the queries.
   *                          Defaults to `false`.
   * @returns An object containing two tree structures:
   *          - `OpenPcrToTestTree`: The tree representing the relationship from Open PCR to Test Case.
   *          - `TestToOpenPcrTree`: The tree representing the relationship from Test Case to Open PCR.
   */
  private async fetchLinkedOpenPcrTestQueries(queries: any, onlySourceSide: boolean = false) {
    const { tree1: OpenPcrToTestTree, tree2: TestToOpenPcrTree } = await this.structureFetchedQueries(
      queries,
      onlySourceSide,
      null,
      ['Bug', 'Change Request'],
      ['Test Case']
    );
    return { OpenPcrToTestTree, TestToOpenPcrTree };
  }

  /**
   * fetches test reporter queries
   * @param queries  fetched queries
   * @returns
   */
  private async fetchTestReporterQueries(queries: any) {
    const { tree1: tree1, tree2: testAssociatedTree } = await this.structureFetchedQueries(
      queries,
      true,
      null,
      ['Requirement', 'Bug', 'Change Request'],
      ['Test Case']
    );
    return { testAssociatedTree };
  }

  /**
   * Fetches System Requirements queries and structures them into a single tree. :)
   *
   * Behavior:
   * - Includes oneHop queries.
   * - Includes tree queries (includeTreeQueries = true).
   * - Includes flat queries (includeFlatQueries = true) matching Epic/Feature/Requirement types.
   * - Skips folders whose names match any of the provided excludedFolderNames (case-insensitive).
   *
   * @param queries - Root or folder query node to search under.
   * @param excludedFolderNames - Folder names to exclude from traversal.
   * @returns An object containing `systemRequirementsQueryTree`.
   */
  private async fetchSystemRequirementQueries(queries: any, excludedFolderNames: string[] = []) {
    const { tree1: systemRequirementsQueryTree } = await this.structureFetchedQueries(
      queries,
      false,
      null,
      ['Epic', 'Feature', 'Requirement'],
      [],
      undefined,
      undefined,
      true, // Enable processing of both tree and direct link queries, including flat queries
      excludedFolderNames,
      true
    );
    return { systemRequirementsQueryTree };
  }

  private async fetchSrsQueries(rootQueries: any) {
    const srsFolder = await this.findQueryFolderByName(rootQueries, 'srs');
    if (!srsFolder) {
      const systemRequirementsQueries = await this.fetchSystemRequirementQueries(rootQueries);
      const { SystemToSoftwareRequirementsTree, SoftwareToSystemRequirementsTree } =
        await this.fetchLinkedRequirementsTraceQueries(rootQueries);
      return {
        systemRequirementsQueries,
        systemToSoftwareRequirementsQueries: SystemToSoftwareRequirementsTree,
        softwareToSystemRequirementsQueries: SoftwareToSystemRequirementsTree,
      };
    }

    const srsFolderWithChildren = await this.ensureQueryChildren(srsFolder);
    const systemRequirementsQueries = await this.fetchSystemRequirementQueries(srsFolderWithChildren, [
      'System to Software',
      'Software to System',
    ]);

    const systemToSoftwareFolder = await this.findChildFolderByName(
      srsFolderWithChildren,
      'System to Software'
    );
    const softwareToSystemFolder = await this.findChildFolderByName(
      srsFolderWithChildren,
      'Software to System'
    );

    const systemToSoftwareRequirementsQueries = await this.fetchRequirementsTraceQueriesForFolder(
      systemToSoftwareFolder
    );
    const softwareToSystemRequirementsQueries = await this.fetchRequirementsTraceQueriesForFolder(
      softwareToSystemFolder
    );

    return {
      systemRequirementsQueries,
      systemToSoftwareRequirementsQueries,
      softwareToSystemRequirementsQueries,
    };
  }

  /**
   * Fetches and structures linked queries related to requirements traceability with area path filtering.
   *
   * This method retrieves and organizes the bidirectional relationships between
   * Epic/Features/Requirements into two tree structures for traceability analysis.
   * - Tree 1: Sources from area path containing "System" → Targets from area path containing "Software"
   * - Tree 2: Sources from area path containing "Software" → Targets from area path containing "System" (reverse)
   *
   * @param queries - The input queries to be processed and structured.
   * @param onlySourceSide - A flag indicating whether to process only the source side of the queries.
   *                          Defaults to `false`.
   * @returns An object containing two tree structures:
   *          - `SystemToSoftwareRequirementsTree`: The tree representing System → Software Requirements traceability.
   *          - `SoftwareToSystemRequirementsTree`: The tree representing Software → System Requirements traceability.
   */
  private async fetchLinkedRequirementsTraceQueries(queries: any, onlySourceSide: boolean = false) {
    const { tree1: SystemToSoftwareRequirementsTree, tree2: SoftwareToSystemRequirementsTree } =
      await this.structureFetchedQueries(
        queries,
        onlySourceSide,
        null,
        ['Epic', 'Feature', 'Requirement'],
        ['Epic', 'Feature', 'Requirement'],
        'sys', // Source area filter for tree1: System area paths
        'soft' // Target area filter for tree1: Software area paths (tree2 will be reversed automatically)
      );
    return { SystemToSoftwareRequirementsTree, SoftwareToSystemRequirementsTree };
  }

  private async fetchRequirementsTraceQueriesForFolder(folder: any) {
    if (!folder) {
      return null;
    }

    const { tree1 } = await this.structureFetchedQueries(
      folder,
      false,
      null,
      ['Epic', 'Feature', 'Requirement'],
      ['Epic', 'Feature', 'Requirement'],
      undefined,
      undefined,
      true
    );
    return tree1;
  }

  private async findQueryFolderByName(rootQuery: any, folderName: string): Promise<any | null> {
    if (!rootQuery || !folderName) {
      return null;
    }

    const normalizedName = folderName.toLowerCase();
    const queue: any[] = [rootQuery];

    while (queue.length > 0) {
      const current = queue.shift();
      if (!current) {
        continue;
      }

      if (current.isFolder && (current.name || '').toLowerCase() === normalizedName) {
        return current;
      }

      if (current.hasChildren) {
        const currentWithChildren = await this.ensureQueryChildren(current);
        if (currentWithChildren?.children?.length) {
          queue.push(...currentWithChildren.children);
        }
      }
    }

    return null;
  }

  private async findChildFolderByName(parent: any, childName: string): Promise<any | null> {
    if (!parent || !childName) {
      return null;
    }

    const parentWithChildren = await this.ensureQueryChildren(parent);
    if (!parentWithChildren?.children?.length) {
      return null;
    }

    const normalizedName = childName.toLowerCase();
    return (
      parentWithChildren.children.find(
        (child: any) => child.isFolder && (child.name || '').toLowerCase() === normalizedName
      ) || null
    );
  }

  /**
   * Performs a breadth-first walk starting at `parent` to locate the nearest folder whose
   * name matches any of the provided candidates (case-insensitive). Exact matches win; if none
   * are found the first partial match encountered is returned. When no candidates are located,
   * the method yields `null`.
   */
  private async findChildFolderByPossibleNames(parent: any, possibleNames: string[]): Promise<any | null> {
    if (!parent || !possibleNames?.length) {
      return null;
    }

    const normalizedNames = possibleNames.map((name) => name.toLowerCase());

    const isMatch = (candidate: string, value: string) => value === candidate;
    const isPartialMatch = (candidate: string, value: string) => value.includes(candidate);

    const tryMatch = (folder: any, matcher: (candidate: string, value: string) => boolean) => {
      const folderName = (folder?.name || '').toLowerCase();
      return normalizedNames.some((candidate) => matcher(candidate, folderName));
    };

    const parentWithChildren = await this.ensureQueryChildren(parent);
    if (!parentWithChildren?.children?.length) {
      return null;
    }

    const queue: any[] = [];
    const visited = new Set<string>();
    let partialCandidate: any = null;

    // Seed the queue with direct children so we prefer closer matches before walking deeper.
    for (const child of parentWithChildren.children) {
      if (!child?.isFolder) {
        continue;
      }
      const childId = child.id ?? `${child.name}-${Math.random()}`;
      queue.push(child);
      visited.add(childId);
    }

    const considerFolder = async (folder: any): Promise<any | null> => {
      if (tryMatch(folder, isMatch)) {
        return await this.ensureQueryChildren(folder);
      }

      if (!partialCandidate && tryMatch(folder, isPartialMatch)) {
        partialCandidate = await this.ensureQueryChildren(folder);
      }

      return null;
    };

    for (const child of queue) {
      const match = await considerFolder(child);
      if (match) {
        return match;
      }
    }

    while (queue.length > 0) {
      const current = queue.shift();
      if (!current) {
        continue;
      }

      const currentWithChildren = await this.ensureQueryChildren(current);
      if (!currentWithChildren) {
        continue;
      }

      const match = await considerFolder(currentWithChildren);
      if (match) {
        return match;
      }

      // Breadth-first expansion so we climb the hierarchy gradually.
      if (currentWithChildren.children?.length) {
        for (const child of currentWithChildren.children) {
          if (!child?.isFolder) {
            continue;
          }
          const childId = child.id ?? `${child.name}-${Math.random()}`;
          if (!visited.has(childId)) {
            visited.add(childId);
            queue.push(child);
          }
        }
      }
    }

    return partialCandidate;
  }

  /**
   * Executes `fetcher` against `startingFolder` and, if the validator deems the result empty,
   * climbs ancestor folders toward `rootQueries` until a satisfactory result is produced.
   * The first successful folder short-circuits the search; otherwise the final attempt is
   * returned to preserve legacy behavior.
   */
  private async fetchWithAncestorFallback(
    rootQueries: any,
    startingFolder: any,
    fetcher: (folder: any) => Promise<any>,
    logContext: string,
    validator?: (result: any) => boolean
  ): Promise<{ result: any; usedFolder: any }> {
    const rootWithChildren = await this.ensureQueryChildren(rootQueries);
    const candidates = await this.buildFallbackChain(rootWithChildren, startingFolder);
    const evaluate = validator ?? ((res: any) => this.hasAnyQueryTree(res));

    let lastResult: any = null;
    let lastFolder: any = startingFolder ?? rootWithChildren;

    for (const candidate of candidates) {
      const enrichedCandidate = await this.ensureQueryChildren(candidate);
      const candidateName = enrichedCandidate?.name ?? '<root>';
      logger.debug(`${logContext} trying folder: ${candidateName}`);
      lastResult = await fetcher(enrichedCandidate);
      lastFolder = enrichedCandidate;
      if (evaluate(lastResult)) {
        logger.debug(`${logContext} using folder: ${candidateName}`);
        return { result: lastResult, usedFolder: enrichedCandidate };
      }
      logger.debug(`${logContext} folder ${candidateName} produced no results, ascending`);
    }

    logger.debug(`${logContext} no folders yielded results, returning last attempt`);
    return { result: lastResult, usedFolder: lastFolder };
  }

  /**
   * Applies `fetchWithAncestorFallback` to each configured branch, resolving dedicated folders
   * when available and emitting a map keyed by branch id. Each outcome includes both the
   * resulting payload and the specific folder that satisfied the fallback chain.
   */
  private async fetchDocTypeBranches(
    queriesWithChildren: any,
    docRoot: any,
    branches: DocTypeBranchConfig[]
  ): Promise<Record<string, FallbackFetchOutcome>> {
    const results: Record<string, FallbackFetchOutcome> = {};
    const effectiveDocRoot = docRoot ?? queriesWithChildren;

    for (const branch of branches) {
      const fallbackStart = branch.fallbackStart ?? effectiveDocRoot;
      let startingFolder = fallbackStart;
      let startingName = startingFolder?.name ?? '<root>';

      // Attempt to locate a more specific child folder, falling back to the provided root if absent.
      if (branch.folderNames?.length && effectiveDocRoot) {
        const resolvedFolder = await this.findChildFolderByPossibleNames(
          effectiveDocRoot,
          branch.folderNames
        );
        if (resolvedFolder) {
          startingFolder = resolvedFolder;
          startingName = resolvedFolder?.name ?? '<root>';
        }
      }

      logger.debug(`${branch.label} starting folder: ${startingName}`);

      const fetchOutcome = await this.fetchWithAncestorFallback(
        queriesWithChildren,
        startingFolder,
        branch.fetcher,
        branch.label,
        branch.validator
      );

      logger.debug(`${branch.label} final folder: ${fetchOutcome.usedFolder?.name ?? '<root>'}`);

      results[branch.id] = fetchOutcome;
    }

    return results;
  }

  /**
   * Constructs an ordered list of folders to probe during fallback. The sequence starts at
   * `startingFolder` (if provided) and walks upward through ancestors to the root query tree,
   * ensuring no folder id appears twice.
   */
  private async buildFallbackChain(rootQueries: any, startingFolder: any): Promise<any[]> {
    const chain: any[] = [];
    const seen = new Set<string>();
    const pushUnique = (node: any) => {
      if (!node) {
        return;
      }
      const id = node.id ?? '__root__';
      if (seen.has(id)) {
        return;
      }
      seen.add(id);
      chain.push(node);
    };

    if (startingFolder?.id) {
      const path = await this.findPathToNode(rootQueries, startingFolder.id);
      if (path) {
        for (let i = path.length - 1; i >= 0; i--) {
          pushUnique(path[i]);
        }
      } else {
        pushUnique(startingFolder);
      }
    } else if (startingFolder) {
      pushUnique(startingFolder);
    }

    pushUnique(rootQueries);
    return chain;
  }

  /**
   * Recursively searches the query tree for the node with the provided id and returns the
   * path (root → target). Nodes are enriched with children on demand and a visited set guards
   * against cycles within malformed data.
   */
  private async findPathToNode(
    currentNode: any,
    targetId: string,
    visited: Set<string> = new Set<string>()
  ): Promise<any[] | null> {
    if (!currentNode) {
      return null;
    }

    const currentId = currentNode.id ?? '__root__';
    if (visited.has(currentId)) {
      return null;
    }
    visited.add(currentId);

    if (currentNode.id === targetId) {
      return [currentNode];
    }

    const enrichedNode = await this.ensureQueryChildren(currentNode);
    const children = enrichedNode?.children;
    if (!children?.length) {
      return null;
    }

    for (const child of children) {
      const path = await this.findPathToNode(child, targetId, visited);
      if (path) {
        return [enrichedNode, ...path];
      }
    }

    return null;
  }

  private async getDocTypeRoot(
    rootQueries: any,
    docTypeName: string
  ): Promise<{ root: any; found: boolean }> {
    if (!rootQueries) {
      return { root: rootQueries, found: false };
    }

    const docTypeFolder = await this.findQueryFolderByName(rootQueries, docTypeName);
    if (docTypeFolder) {
      const folderWithChildren = await this.ensureQueryChildren(docTypeFolder);
      return { root: folderWithChildren, found: true };
    }

    return { root: rootQueries, found: false };
  }

  private async ensureQueryChildren(node: any): Promise<any> {
    if (!node || !node.hasChildren || node.children) {
      return node;
    }

    if (!node.url) {
      return node;
    }

    const queryUrl = `${node.url}?$depth=2&$expand=all`;
    const refreshedNode = await TFSServices.getItemContent(queryUrl, this.token);
    Object.assign(node, refreshedNode);
    return node;
  }

  async GetQueryResultsFromWiql(
    wiqlHref: string = '',
    displayAsTable: boolean = false,
    testCaseToRelatedWiMap: Map<number, Set<any>>
  ): Promise<any> {
    try {
      if (!wiqlHref) {
        throw new Error('Incorrect WIQL Link');
      }
      // Remember to add customer id if needed
      const queryResult: QueryTree = await TFSServices.getItemContent(wiqlHref, this.token);
      if (!queryResult) {
        throw new Error('Query result failed');
      }

      switch (queryResult.queryType) {
        case QueryType.OneHop:
          return displayAsTable
            ? await this.parseDirectLinkedQueryResultForTableFormat(queryResult, testCaseToRelatedWiMap)
            : await this.parseTreeQueryResult(queryResult);
        case QueryType.Tree:
          return await this.parseTreeQueryResult(queryResult);
        case QueryType.Flat:
          return displayAsTable
            ? await this.parseFlatQueryResultForTableFormat(queryResult)
            : await this.parseFlatQueryResult(queryResult);
        default:
          break;
      }
    } catch (err: any) {
      logger.error(`Could not fetch query results for ${wiqlHref}: ${err.message}`);
    }
  }

  private async parseDirectLinkedQueryResultForTableFormat(
    queryResult: QueryTree,
    testCaseToRelatedWiMap: Map<number, Set<any>>
  ) {
    const { columns, workItemRelations } = queryResult;

    if (workItemRelations?.length === 0) {
      throw new Error('No related work items were found');
    }

    const columnsToShowMap: Map<string, string> = new Map();

    const columnSourceMap: Map<string, string> = new Map();
    const columnTargetsMap: Map<string, string> = new Map();

    //The map is copied because there are different fields for each WI type,
    //Need to consider both ways (Req->Test; Test->Req)
    columns.forEach((column: any) => {
      const { referenceName, name } = column;
      if (name === 'CustomerRequirementId') {
        columnsToShowMap.set(referenceName, 'Customer ID');
      } else {
        columnsToShowMap.set(referenceName, name);
      }
    });

    // Initialize maps
    const sourceTargetsMap: Map<any, any[]> = new Map();
    const lookupMap: Map<number, any> = new Map();

    if (workItemRelations) {
      // Step 1: Collect all unique work item IDs that need to be fetched
      const sourceIds = new Set<number>();
      const targetIds = new Set<number>();

      for (const relation of workItemRelations) {
        if (!relation.source) {
          // Root link - target is actually the source
          sourceIds.add(relation.target.id);
        } else {
          sourceIds.add(relation.source.id);
          if (relation.target) {
            targetIds.add(relation.target.id);
          }
        }
      }

      // Step 2: Fetch all work items in parallel with concurrency limit
      const allSourcePromises = Array.from(sourceIds).map((id) =>
        this.limit(() => {
          const relation = workItemRelations.find(
            (r) => (!r.source && r.target.id === id) || r.source?.id === id
          );
          return this.fetchWIForQueryResult(relation, columnsToShowMap, columnSourceMap, true);
        })
      );

      const allTargetPromises = Array.from(targetIds).map((id) =>
        this.limit(() => {
          const relation = workItemRelations.find((r) => r.target?.id === id);
          return this.fetchWIForQueryResult(relation, columnsToShowMap, columnTargetsMap, true);
        })
      );

      // Wait for all fetches to complete in parallel (with concurrency control)
      const [sourceWorkItems, targetWorkItems] = await Promise.all([
        Promise.all(allSourcePromises),
        Promise.all(allTargetPromises),
      ]);

      // Build lookup maps
      const sourceWorkItemMap = new Map<number, any>();
      sourceWorkItems.forEach((wi) => {
        sourceWorkItemMap.set(wi.id, wi);
        if (!lookupMap.has(wi.id)) {
          lookupMap.set(wi.id, wi);
        }
      });

      const targetWorkItemMap = new Map<number, any>();
      targetWorkItems.forEach((wi) => {
        targetWorkItemMap.set(wi.id, wi);
        if (!lookupMap.has(wi.id)) {
          lookupMap.set(wi.id, wi);
        }
      });

      // Step 3: Build the sourceTargetsMap using the fetched work items
      for (const relation of workItemRelations) {
        if (!relation.source) {
          // Root link
          const wi = sourceWorkItemMap.get(relation.target.id);
          if (wi && !sourceTargetsMap.has(wi)) {
            sourceTargetsMap.set(wi, []);
          }
          continue;
        }

        if (!relation.target) {
          throw new Error('Target relation is missing');
        }

        const sourceWorkItem = sourceWorkItemMap.get(relation.source.id);
        if (!sourceWorkItem) {
          throw new Error('Source relation has no mapping');
        }

        const targetWi = targetWorkItemMap.get(relation.target.id);
        if (!targetWi) {
          throw new Error('Target work item not found');
        }

        // In case if source is a test case
        this.mapTestCaseToRelatedItem(sourceWorkItem, targetWi, testCaseToRelatedWiMap);

        // In case of target is a test case
        this.mapTestCaseToRelatedItem(targetWi, sourceWorkItem, testCaseToRelatedWiMap);

        const targets: any = sourceTargetsMap.get(sourceWorkItem) || [];
        targets.push(targetWi);
        sourceTargetsMap.set(sourceWorkItem, targets);
      }
    }

    columnsToShowMap.clear();
    return {
      sourceTargetsMap,
      sortingSourceColumnsMap: columnSourceMap,
      sortingTargetsColumnsMap: columnTargetsMap,
    };
  }

  private mapTestCaseToRelatedItem(
    sourceWi: any,
    targetWi: any,
    testCaseToRelatedItemMap: Map<number, Set<any>>
  ) {
    if (sourceWi.fields['System.WorkItemType'] == 'Test Case') {
      if (!testCaseToRelatedItemMap.has(sourceWi.id)) {
        testCaseToRelatedItemMap.set(sourceWi.id, new Set());
      }
      const relatedItemsSet = testCaseToRelatedItemMap.get(sourceWi.id);
      if (relatedItemsSet) {
        // Check if there's already an item with the same ID
        const alreadyExists = [...relatedItemsSet].some((reqItem) => reqItem.id === targetWi.id);
        if (!alreadyExists) {
          relatedItemsSet.add(targetWi);
        }
      }
    }
  }

  private async parseFlatQueryResultForTableFormat(queryResult: QueryTree) {
    const { columns, workItems } = queryResult;

    if (workItems?.length === 0) {
      throw new Error('No work items were found');
    }

    const columnsToShowMap: Map<string, string> = new Map();

    const fieldsToIncludeMap: Map<string, string> = new Map();

    //The map is copied because there are different fields for each WI type,
    columns.forEach((column: any) => {
      const { referenceName, name } = column;
      if (name === 'CustomerRequirementId') {
        columnsToShowMap.set(referenceName, 'Customer ID');
      } else {
        columnsToShowMap.set(referenceName, name);
      }
    });

    // Fetch all work items in parallel with concurrency limit
    const wiSet: Set<any> = new Set();
    if (workItems) {
      const fetchPromises = workItems.map((workItem) =>
        this.limit(() => this.fetchWIForQueryResult(workItem, columnsToShowMap, fieldsToIncludeMap, false))
      );

      const fetchedWorkItems = await Promise.all(fetchPromises);
      fetchedWorkItems.forEach((wi) => wiSet.add(wi));
    }

    columnsToShowMap.clear();
    return {
      fetchedWorkItems: [...wiSet],
      fieldsToIncludeMap,
    };
  }

  private async parseTreeQueryResult(queryResult: QueryTree) {
    const { workItemRelations } = queryResult;
    if (!workItemRelations) return null;

    logger.debug(`parseTreeQueryResult: Processing ${workItemRelations.length} workItemRelations`);

    const allItems: Record<number, any> = {};
    const rootOrder: number[] = [];
    const rootSet = new Set<number>(); // track roots for dedupe

    // Initialize ALL nodes from workItemRelations (not just hierarchy ones)
    // This ensures nodes with non-hierarchy links are also available
    for (const rel of workItemRelations) {
      const t = rel.target;
      if (!allItems[t.id]) await this.initTreeQueryResultItem(t, allItems);

      // Also initialize source nodes if they exist
      if (rel.source && !allItems[rel.source.id]) {
        await this.initTreeQueryResultItem(rel.source, allItems);
      }

      if (rel.rel === null && rel.source === null) {
        if (!rootSet.has(t.id)) {
          rootSet.add(t.id);
          rootOrder.push(t.id);
        }
      }
    }
    logger.debug(
      `parseTreeQueryResult: Found ${rootOrder.length} roots, ${Object.keys(allItems).length} total nodes`
    );

    // Attach only forward hierarchy edges; dedupe children by id per parent
    let hierarchyCount = 0;
    let skippedNonHierarchy = 0;
    for (const rel of workItemRelations) {
      if (!rel.source) continue;
      const linkType = (rel.rel || '').toLowerCase();
      if (!linkType.includes('hierarchy-forward')) {
        skippedNonHierarchy++;
        continue; // skip reverse/others
      }

      const parentId = rel.source.id;
      const childId = rel.target.id;

      // Nodes should already be initialized, but double-check
      if (!allItems[parentId]) {
        logger.warn(`Parent ${parentId} not found, initializing now`);
        await this.initTreeQueryResultItem(rel.source, allItems);
      }
      if (!allItems[childId]) {
        logger.warn(`Child ${childId} not found, initializing now`);
        await this.initTreeQueryResultItem(rel.target, allItems);
      }

      const parent = allItems[parentId];
      parent._childrenSet ||= new Set<number>();
      if (!parent._childrenSet.has(childId)) {
        parent._childrenSet.add(childId);
        parent.children.push(allItems[childId]);
        hierarchyCount++;
      }
      // If this child was previously considered a root, remove it from roots
      if (rootSet.has(childId)) {
        rootSet.delete(childId);
      }
    }
    logger.debug(
      `parseTreeQueryResult: ${hierarchyCount} hierarchy links, ${skippedNonHierarchy} non-hierarchy links skipped`
    );

    // Return roots in original order, excluding those that became children
    const roots = rootOrder.filter((id) => rootSet.has(id)).map((id) => allItems[id]);
    logger.debug(
      `parseTreeQueryResult: Returning ${roots.length} roots with ${Object.keys(allItems).length} total items`
    );

    // Optional: clean helper sets
    for (const id in allItems) delete allItems[id]._childrenSet;

    return {
      roots, // your parsed tree (what you currently return)
      workItemRelations, // all relations for link-driven rendering
      allItems, // all fetched items (including those not in hierarchy)
    };
  }

  private async initTreeQueryResultItem(item: any, allItems: any) {
    const urlWi = `${item.url}?fields=System.Description,System.Title,Microsoft.VSTS.TCM.ReproSteps,Microsoft.VSTS.CMMI.Symptom`;
    const wi = await TFSServices.getItemContent(urlWi, this.token);
    // need to fetch the WI with only the the title, the web URL and the description
    allItems[item.id] = {
      id: item.id,
      title: wi.fields['System.Title'] || '',
      description: wi.fields['Microsoft.VSTS.CMMI.Symptom'] ?? wi.fields['System.Description'] ?? '',
      htmlUrl: wi._links.html.href,
      children: [],
    };
  }

  private async initFlatQueryResultItem(item: any, workItemMap: Map<number, any>) {
    const urlWi = `${item.url}?fields=System.Description,System.Title,Microsoft.VSTS.TCM.ReproSteps,Microsoft.VSTS.CMMI.Symptom`;
    const wi = await TFSServices.getItemContent(urlWi, this.token);
    // need to fetch the WI with only the the title, the web URL and the description
    workItemMap.set(item.id, {
      id: item.id,
      title: wi.fields['System.Title'] || '',
      description: wi.fields['Microsoft.VSTS.CMMI.Symptom'] ?? wi.fields['System.Description'] ?? '',
      htmlUrl: wi._links.html.href,
    });
  }

  private async parseFlatQueryResult(queryResult: QueryTree) {
    const { workItems } = queryResult;
    if (!workItems) {
      logger.warn(`No work items were found for this requested query`);
      return null;
    }

    try {
      const workItemsResultMap: Map<number, any> = new Map();
      for (const wi of workItems) {
        if (!workItemsResultMap.has(wi.id)) {
          await this.initFlatQueryResultItem(wi, workItemsResultMap);
        }
      }
      return [...workItemsResultMap.values()];
    } catch (error: any) {
      logger.error(`could not parse requested flat query: ${error.message}`);
    }
  }

  private async fetchWIForQueryResult(
    receivedObject: any,
    columnMap: Map<string, string>,
    resultedRefNameMap: Map<string, string>,
    isRelation: boolean
  ) {
    const url = isRelation ? `${receivedObject.target.url}` : `${receivedObject.url}`;
    const wi: any = await TFSServices.getItemContent(url, this.token);
    if (!wi) {
      throw new Error(`WI ${isRelation ? receivedObject.target.id : receivedObject.id} not found`);
    }

    this.filterFieldsByColumns(wi, columnMap, resultedRefNameMap);
    return wi;
  }

  // get queris structured
  GetModeledQuery(list: Array<any>): Array<any> {
    let queryListObject: Array<any> = [];
    list.forEach((query) => {
      let newObj = {
        queryName: query.name,
        wiql: query._links.wiql != null ? query._links.wiql : null,
        id: query.id,
      };
      queryListObject.push(newObj);
    });
    return queryListObject;
  }
  // gets query results
  async GetQueryResultsByWiqlHref(wiqlHref: string, project: string): Promise<any> {
    try {
      let results: any = await TFSServices.getItemContent(wiqlHref, this.token);
      let modeledResult = await this.GetModeledQueryResults(results, project);
      if (modeledResult.queryType == 'tree') {
        let levelResults: Array<Workitem> = Helper.LevelBuilder(
          modeledResult,
          modeledResult.workItems[0].fields[0].value
        );
        return levelResults;
      }
      return modeledResult.workItems;
    } catch (e) {}
  }
  // gets query results
  async GetQueryResultsByWiqlString(wiql: string, projectName: string): Promise<any> {
    let res;
    let url = `${this.orgUrl}${projectName}/_apis/wit/wiql?$top=2147483646&expand=all`;
    try {
      res = await TFSServices.getItemContent(url, this.token, 'post', {
        query: wiql,
      });
    } catch (error) {
      console.log(error);
      return [];
    }
    return res;
  }
  async GetQueryResultById(query: string, project: string): Promise<any> {
    var url = `${this.orgUrl}${project}/_apis/wit/queries/${query}`;
    let querie: any = await TFSServices.getItemContent(url, this.token);
    var wiql = querie._links.wiql;
    return await this.GetQueryResultsByWiqlHref(wiql.href, project);
  }

  async PopulateWorkItemsByIds(workItemsArray: any[] = [], projectName: string = ''): Promise<any[]> {
    let url = `${this.orgUrl}${projectName}/_apis/wit/workitemsbatch`;
    let res: any[] = [];
    let divByMax = Math.floor(workItemsArray.length / 200);
    let modulusByMax = workItemsArray.length % 200;
    //iterating
    for (let i = 0; i < divByMax; i++) {
      let from = i * 200;
      let to = (i + 1) * 200;
      let currentIds = workItemsArray.slice(from, to);
      try {
        let subRes = await TFSServices.getItemContent(url, this.token, 'post', {
          $expand: 'Relations',
          ids: currentIds,
        });
        res = [...res, ...subRes.value];
      } catch (error) {
        logger.error(`error populating workitems array`);
        logger.error(JSON.stringify(error));
        return [];
      }
    }
    //compliting the rimainder
    if (modulusByMax !== 0) {
      try {
        let currentIds = workItemsArray.slice(workItemsArray.length - modulusByMax, workItemsArray.length);
        let subRes = await TFSServices.getItemContent(url, this.token, 'post', {
          $expand: 'Relations',
          ids: currentIds,
        });
        res = [...res, ...subRes.value];
      } catch (error) {
        logger.error(`error populating workitems array`);
        logger.error(JSON.stringify(error));
        return [];
      }
    } //if

    return res;
  }

  async GetModeledQueryResults(results: any, project: string) {
    let ticketsDataProvider = new TicketsDataProvider(this.orgUrl, this.token);
    var queryResult: Query = new Query();
    queryResult.asOf = results.asOf;
    queryResult.queryResultType = results.queryResultType;
    queryResult.queryType = results.queryType;
    if (results.queryType == QueryType.Flat) {
      //     //Flat Query
      //TODo: attachment
      //TODO:check if wi.relations exist
      //TODO: add attachment to any list from 1
      for (var j = 0; j < results.workItems.length; j++) {
        let wi = await ticketsDataProvider.GetWorkItem(project, results.workItems[j].id);

        queryResult.workItems[j] = new Workitem();
        queryResult.workItems[j].url = results.workItems[j].url;
        queryResult.workItems[j].fields = new Array(results.columns.length);
        if (wi.relations != null) {
          queryResult.workItems[j].attachments = wi.relations;
        }
        var rel = new QueryAllTypes();
        for (var i = 0; i < results.columns.length; i++) {
          queryResult.columns[i] = new Column();
          queryResult.workItems[j].fields[i] = new value();
          queryResult.columns[i].name = results.columns[i].name;
          queryResult.columns[i].referenceName = results.columns[i].referenceName;
          queryResult.columns[i].url = results.columns[i].url;
          if (results.columns[i].referenceName.toUpperCase() == 'SYSTEM.ID') {
            queryResult.workItems[j].fields[i].value = wi.id.toString();
            queryResult.workItems[j].fields[i].name = 'ID';
          } else if (
            results.columns[i].referenceName.toUpperCase() == 'SYSTEM.ASSIGNEDTO' &&
            wi.fields[results.columns[i].referenceName] != null
          )
            queryResult.workItems[j].fields[i].value =
              wi.fields[results.columns[i].referenceName].displayName;
          else {
            let s: string = wi.fields[results.columns[i].referenceName];
            queryResult.workItems[j].fields[i].value = wi.fields[results.columns[i].referenceName];
            queryResult.workItems[j].fields[i].name = results.columns[i].name;
          }
        }
      }
    } //Tree Query
    else {
      this.BuildColumns(results, queryResult);
      for (var j = 0; j < results.workItemRelations.length; j++) {
        if (results.workItemRelations[j].target != null) {
          let wiT = await ticketsDataProvider.GetWorkItem(project, results.workItemRelations[j].target.id);
          // var rel = new QueryAllTypes();
          queryResult.workItems[j] = new Workitem();
          queryResult.workItems[j].url = wiT.url;
          queryResult.workItems[j].fields = new Array(results.columns.length);
          if (wiT.relations != null) {
            queryResult.workItems[j].attachments = wiT.relations;
          }
          // rel.q = queryResult;
          for (i = 0; i < queryResult.columns.length; i++) {
            //..  rel.q.workItems[j].fields[i] = new value();
            queryResult.workItems[j].fields[i] = new value();
            queryResult.workItems[j].fields[i].name = queryResult.columns[i].name;
            if (
              results.columns[i].referenceName.toUpperCase() == 'SYSTEM.ASSIGNEDTO' &&
              wiT.fields[results.columns[i].referenceName] != null
            )
              queryResult.workItems[j].fields[i].value =
                wiT.fields[results.columns[i].referenceName].displayName;
            else queryResult.workItems[j].fields[i].value = wiT.fields[queryResult.columns[i].referenceName];
            //}
          }
          if (results.workItemRelations[j].source != null)
            queryResult.workItems[j].Source = results.workItemRelations[j].source.id;
        }
      }
    }
    return queryResult;
  }

  //Build columns
  BuildColumns(results: any, queryResult: Query) {
    for (var i = 0; i < results.columns.length; i++) {
      queryResult.columns[i] = new Column();
      queryResult.columns[i].name = results.columns[i].name;
      queryResult.columns[i].referenceName = results.columns[i].referenceName;
      queryResult.columns[i].url = results.columns[i].url;
    }
  }

  async GetIterationsByTeamName(projectName: string, teamName: string): Promise<any[]> {
    let res: any;
    let url;
    if (teamName) {
      url = `${this.orgUrl}${projectName}/${teamName}/_apis/work/teamsettings/iterations`;
    } else {
      url = `${this.orgUrl}${projectName}/_apis/work/teamsettings/iterations`;
    }
    res = await TFSServices.getItemContent(url, this.token, 'get');
    return res;
  } //GetIterationsByTeamName

  async CreateNewWorkItem(projectName: string, wiBody: any, wiType: string, byPass: boolean) {
    let url = `${this.orgUrl}${projectName}/_apis/wit/workitems/$${wiType}?bypassRules=${String(
      byPass
    ).toString()}`;
    return TFSServices.getItemContent(url, this.token, 'POST', wiBody, {
      'Content-Type': 'application/json-patch+json',
    });
  } //CreateNewWorkItem

  async GetWorkitemAttachments(project: string, id: string) {
    let attachmentList: Array<any> = [];
    let ticketsDataProvider = new TicketsDataProvider(this.orgUrl, this.token);
    try {
      let wi = await ticketsDataProvider.GetWorkItem(project, id);
      if (!wi?.relations) return [];
      await Promise.all(
        wi.relations.map(async (relation: any) => {
          if (relation.rel == 'AttachedFile') {
            let attachment = JSON.parse(JSON.stringify(relation));
            attachment.downloadUrl = `${relation.url}/${relation.attributes.name}`;
            attachmentList.push(attachment);
          }
        })
      );
      return attachmentList;
    } catch (e) {
      logger.error(`error fetching attachments for work item ${id}`);
      logger.error(`${JSON.stringify(e)}`);
      return [];
    }
  }

  //Get work item attachments
  async GetWorkitemAttachmentsJSONData(project: string, attachmentId: string) {
    let wiuRL = `${this.orgUrl}${project}/_apis/wit/attachments/${attachmentId}`;
    let attachment = await TFSServices.getItemContent(wiuRL, this.token);
    return attachment;
  }

  //Update work item
  async UpdateWorkItem(projectName: string, wiBody: any, workItemId: number, byPass: boolean) {
    let res: any;
    let url: string = `${this.orgUrl}${projectName}/_apis/wit/workitems/${workItemId}?bypassRules=${String(
      byPass
    ).toString()}`;
    res = await TFSServices.getItemContent(url, this.token, 'patch', wiBody, {
      'Content-Type': 'application/json-patch+json',
    });
    return res;
  } //CreateNewWorkItem

  private async structureAllQueryPath(rootQuery: any, parentId: any = null): Promise<any> {
    try {
      if (!rootQuery.hasChildren) {
        if (!rootQuery.isFolder) {
          let sysOverviewNode = null;
          let knownBugsNode = null;
          if (rootQuery.queryType === 'flat' && this.matchesBugCondition(rootQuery.wiql)) {
            //Add this to the known bugs query tree
            knownBugsNode = {
              id: rootQuery.id,
              pId: parentId,
              value: rootQuery.name,
              title: rootQuery.name,
              queryType: rootQuery.queryType,
              columns: rootQuery.columns,
              wiql: rootQuery._links.wiql ?? undefined,
              isValidQuery: true,
            };
          }
          sysOverviewNode = {
            id: rootQuery.id,
            pId: parentId,
            value: rootQuery.name,
            title: rootQuery.name,
            queryType: rootQuery.queryType,
            columns: rootQuery.columns,
            wiql: rootQuery._links.wiql ?? undefined,
            isValidQuery: true,
          };

          return { tree1: sysOverviewNode, tree2: knownBugsNode };
        } else {
          return { tree1: null, tree2: null };
        }
      }

      if (!rootQuery.children) {
        const queryUrl = `${rootQuery.url}?$depth=2&$expand=all`;
        const currentQuery = await TFSServices.getItemContent(queryUrl, this.token);
        return currentQuery
          ? await this.structureAllQueryPath(currentQuery, currentQuery.id)
          : { tree1: null, tree2: null };
      }

      // Process children recursively
      const childResults = await Promise.all(
        rootQuery.children.map((child: any) => this.structureAllQueryPath(child, rootQuery.id))
      );

      // Build tree
      const sysOverviewNodeTreeChildren = childResults
        .map((res: any) => res.tree1)
        .filter((child: any) => child !== null);
      const sysOverviewNode =
        sysOverviewNodeTreeChildren.length > 0
          ? {
              id: rootQuery.id,
              pId: parentId,
              value: rootQuery.name,
              title: rootQuery.name,
              children: sysOverviewNodeTreeChildren,
              columns: rootQuery.columns,
            }
          : null;

      const knownBugsTreeChildren = childResults
        .map((res: any) => res.tree2)
        .filter((child: any) => child !== null);
      const knownBugs =
        knownBugsTreeChildren.length > 0
          ? {
              id: rootQuery.id,
              pId: parentId,
              value: rootQuery.name,
              title: rootQuery.name,
              children: knownBugsTreeChildren,
              columns: rootQuery.columns,
            }
          : null;

      return { tree1: sysOverviewNode, tree2: knownBugs };
    } catch (err: any) {
      logger.error(
        `Error occurred while constructing the query list ${err.message} with query ${JSON.stringify(
          rootQuery
        )}`
      );
      throw err;
    }
  }

  /**
   * Recursively structures fetched queries into two hierarchical trees (tree1 and tree2)
   * by matching WIQL against allowed Source/Target types and optional area filters.
   * Supports leaf queries of type:
   * - oneHop (always)
   * - tree (when includeTreeQueries === true)
   * - flat (when includeFlatQueries === true)
   *
   * @param rootQuery - The root query object to process. It may contain children or be a leaf node.
   * @param onlyTestReq - A boolean flag that, when true, suppresses adding matching queries to tree1.
   * @param parentId - The ID of the parent node, used to maintain the hierarchy. Defaults to `null`.
   * @param sources - Allowed Source work item types to match in WIQL.
   * @param targets - Allowed Target work item types to match in WIQL.
   * @param sourceAreaFilter - Optional area path filter for the Source side (leaf name substring match).
   * @param targetAreaFilter - Optional area path filter for the Target side (leaf name substring match).
   * @param includeTreeQueries - Include 'tree' queries in addition to 'oneHop'. Defaults to `false`.
   * @param excludedFolderNames - Optional list of folder names to skip entirely (case-insensitive exact match).
   * @param includeFlatQueries - Include 'flat' queries (matched by [System.WorkItemType] and [System.AreaPath]). Defaults to `false`.
   * @returns A promise resolving to an object with `tree1` and `tree2` nodes, or `null` for each when none match.
   * @throws Logs an error if an exception occurs during processing.
   */
  private async structureFetchedQueries(
    rootQuery: any,
    onlyTestReq: boolean,
    parentId: any = null,
    sources: string[],
    targets: string[],
    sourceAreaFilter?: string,
    targetAreaFilter?: string,
    includeTreeQueries: boolean = false,
    excludedFolderNames: string[] = [],
    includeFlatQueries: boolean = false
  ): Promise<any> {
    try {
      const shouldSkipFolder =
        rootQuery?.isFolder &&
        excludedFolderNames.some(
          (folderName) => folderName.toLowerCase() === (rootQuery.name || '').toLowerCase()
        );

      if (shouldSkipFolder) {
        return { tree1: null, tree2: null };
      }

      if (!rootQuery.hasChildren) {
        const isLeafCandidate =
          !rootQuery.isFolder &&
          (rootQuery.queryType === 'oneHop' ||
            (includeTreeQueries && rootQuery.queryType === 'tree') ||
            (includeFlatQueries && rootQuery.queryType === 'flat'));
        if (isLeafCandidate) {
          const wiql = rootQuery.wiql;
          let tree1Node = null;
          let tree2Node = null;

          if (rootQuery.queryType === 'flat' && includeFlatQueries) {
            const allTypes = Array.from(new Set([...(sources || []), ...(targets || [])]));
            const typesOk = this.matchesFlatWorkItemTypeCondition(wiql, allTypes);

            if (typesOk) {
              const allowTree1 =
                !onlyTestReq &&
                (sourceAreaFilter ? this.matchesFlatAreaCondition(wiql, sourceAreaFilter || '') : true);
              const allowTree2 = targetAreaFilter
                ? this.matchesFlatAreaCondition(wiql, targetAreaFilter || '')
                : true;

              if (allowTree1) {
                tree1Node = this.buildQueryNode(rootQuery, parentId);
              }

              if (allowTree2) {
                tree2Node = this.buildQueryNode(rootQuery, parentId);
              }
            }
          } else {
            if (!onlyTestReq && this.matchesSourceTargetCondition(wiql, sources, targets)) {
              const matchesAreaPath =
                sourceAreaFilter || targetAreaFilter
                  ? this.matchesAreaPathCondition(wiql, sourceAreaFilter || '', targetAreaFilter || '')
                  : true;

              if (matchesAreaPath) {
                tree1Node = this.buildQueryNode(rootQuery, parentId);
              }
            }
            if (this.matchesSourceTargetCondition(wiql, targets, sources)) {
              const matchesReverseAreaPath =
                sourceAreaFilter || targetAreaFilter
                  ? this.matchesAreaPathCondition(wiql, targetAreaFilter || '', sourceAreaFilter || '')
                  : true;

              if (matchesReverseAreaPath) {
                tree2Node = this.buildQueryNode(rootQuery, parentId);
              }
            }
          }
          return { tree1: tree1Node, tree2: tree2Node };
        } else {
          return { tree1: null, tree2: null };
        }
      }
      // If the query has children, but they are not loaded, fetch them
      if (!rootQuery.children) {
        const queryUrl = `${rootQuery.url}?$depth=2&$expand=all`;
        const currentQuery = await TFSServices.getItemContent(queryUrl, this.token);
        return currentQuery
          ? await this.structureFetchedQueries(
              currentQuery,
              onlyTestReq,
              currentQuery.id,
              sources,
              targets,
              sourceAreaFilter,
              targetAreaFilter,
              includeTreeQueries,
              excludedFolderNames,
              includeFlatQueries
            )
          : { tree1: null, tree2: null };
      }

      // Process children recursively
      const childResults = await Promise.all(
        rootQuery.children.map((child: any) =>
          this.structureFetchedQueries(
            child,
            onlyTestReq,
            rootQuery.id,
            sources,
            targets,
            sourceAreaFilter,
            targetAreaFilter,
            includeTreeQueries,
            excludedFolderNames,
            includeFlatQueries
          )
        )
      );

      // Build tree1
      const tree1Children = childResults.map((res: any) => res.tree1).filter((child: any) => child !== null);
      const tree1Node =
        tree1Children.length > 0
          ? {
              id: rootQuery.id,
              pId: parentId,
              value: rootQuery.name,
              title: rootQuery.name,
              children: tree1Children,
            }
          : null;

      // Build tree2
      const tree2Children = childResults.map((res: any) => res.tree2).filter((child: any) => child !== null);
      const tree2Node =
        tree2Children.length > 0
          ? {
              id: rootQuery.id,
              value: rootQuery.name,
              pId: parentId,
              title: rootQuery.name,
              children: tree2Children,
            }
          : null;

      return { tree1: tree1Node, tree2: tree2Node };
    } catch (err: any) {
      logger.error(
        `Error occurred while constructing the query list ${err.message} with query ${JSON.stringify(
          rootQuery
        )}`
      );
      logger.error(`Error stack ${err.message}`);
    }
  }

  /**
   * Checks if WIQL matches area path conditions for System/Software requirements filtering
   * @param wiql - The WIQL string to evaluate
   * @param sourceAreaFilter - Area path filter for source (e.g., "System" or "Software")
   * @param targetAreaFilter - Area path filter for target (e.g., "Software" or "System")
   * @returns Boolean indicating if the WIQL matches the area path conditions
   */
  private matchesAreaPathCondition(
    wiql: string,
    sourceAreaFilter: string,
    targetAreaFilter: string
  ): boolean {
    const wiqlLower = (wiql || '').toLowerCase();
    const srcFilter = (sourceAreaFilter || '').toLowerCase().trim();
    const tgtFilter = (targetAreaFilter || '').toLowerCase().trim();

    const extractAreaPaths = (owner: 'source' | 'target'): string[] => {
      const re = new RegExp(`${owner}\\.\\[system\\.areapath\\][^']*'([^']+)'`, 'gi');
      const results: string[] = [];
      let match: RegExpExecArray | null;
      while ((match = re.exec(wiqlLower)) !== null) {
        // match[1] is the quoted area path value already in lowercase (because wiqlLower)
        results.push(match[1]);
      }
      return results;
    };

    const getLeaf = (p: string): string => {
      const parts = p.split(/[\\/]/); // supports both \\ and /
      return parts[parts.length - 1] || p;
    };

    const sourceAreaPaths = extractAreaPaths('source');
    const targetAreaPaths = extractAreaPaths('target');

    // Match only against the sub-area (leaf) name, e.g. Test CMMI\\Software_Requirement -> Software_Requirement
    const hasSourceAreaPath = !srcFilter || sourceAreaPaths.some((p) => getLeaf(p).includes(srcFilter));
    const hasTargetAreaPath = !tgtFilter || targetAreaPaths.some((p) => getLeaf(p).includes(tgtFilter));

    return hasSourceAreaPath && hasTargetAreaPath;
  }

  /**
   * Determines whether the given WIQL (Work Item Query Language) string matches the specified
   * source and target conditions. It checks if the WIQL contains references to the specified
   * source and target work item types.
   *
   * Supports both equality (=) and IN operators:
   * - Source.[System.WorkItemType] = 'Epic'
   * - Source.[System.WorkItemType] IN ('Epic', 'Feature', 'Requirement')
   *
   * @param wiql - The WIQL string to evaluate.
   * @param source - An array of source work item types to check for in the WIQL.
   * @param target - An array of target work item types to check for in the WIQL.
   * @returns A boolean indicating whether the WIQL includes at least one valid source work item type
   *          and at least one valid target work item type.
   */
  private matchesSourceTargetCondition(wiql: string, source: string[], target: string[]): boolean {
    const isSourceIncluded = this.matchesWorkItemTypeCondition(wiql, 'Source', source);
    const isTargetIncluded = this.matchesWorkItemTypeCondition(wiql, 'Target', target);
    return isSourceIncluded && isTargetIncluded;
  }

  /**
   * Helper method to check if a WIQL contains valid work item types for a given context (Source/Target).
   * Supports both = and IN operators.
   *
   * @param wiql - The WIQL string to evaluate
   * @param context - Either 'Source' or 'Target'
   * @param allowedTypes - Array of allowed work item types
   * @returns true if all work item types in the WIQL are in the allowedTypes array
   */
  private matchesWorkItemTypeCondition(
    wiql: string,
    context: 'Source' | 'Target',
    allowedTypes: string[]
  ): boolean {
    // If allowedTypes is empty, accept any work item type (for backward compatibility)
    if (allowedTypes.length === 0) {
      return wiql.includes(`${context}.[System.WorkItemType]`);
    }

    const fieldPattern = `${context}.\\[System.WorkItemType\\]`;

    // Pattern for equality: Source.[System.WorkItemType] = 'Epic'
    const equalityRegex = new RegExp(`${fieldPattern}\\s*=\\s*'([^']+)'`, 'gi');

    // Pattern for IN operator: Source.[System.WorkItemType] IN ('Epic', 'Feature', 'Requirement')
    const inRegex = new RegExp(`${fieldPattern}\\s+IN\\s*\\(([^)]+)\\)`, 'gi');

    const foundTypes = new Set<string>();

    // Extract types from equality operators
    let match;
    while ((match = equalityRegex.exec(wiql)) !== null) {
      foundTypes.add(match[1].trim());
    }

    // Extract types from IN operators
    while ((match = inRegex.exec(wiql)) !== null) {
      const typesString = match[1];
      // Extract all quoted values from the IN clause
      const typeMatches = typesString.matchAll(/'([^']+)'/g);
      for (const typeMatch of typeMatches) {
        foundTypes.add(typeMatch[1].trim());
      }
    }

    // If no work item types found in WIQL, return false
    if (foundTypes.size === 0) {
      return false;
    }

    // Check if all found types are in the allowedTypes array
    for (const type of foundTypes) {
      if (!allowedTypes.includes(type)) {
        // Found a type that's not in the allowed list - reject this query
        return false;
      }
    }

    // All found types are valid
    return true;
  }

  // Build a normalized node object for tree outputs
  private buildQueryNode(rootQuery: any, parentId: any) {
    return {
      id: rootQuery.id,
      pId: parentId,
      value: rootQuery.name,
      title: rootQuery.name,
      queryType: rootQuery.queryType,
      columns: rootQuery.columns,
      wiql: rootQuery._links.wiql ?? undefined,
      isValidQuery: true,
    };
  }

  /**
   * Matches flat query WIQL against allowed work item types.
   * Accept when at least one type is present and all found types are within the allowed set.
   */
  private matchesFlatWorkItemTypeCondition(wiql: string, allowedTypes: string[]): boolean {
    // If allowedTypes is empty, accept any work item type reference
    if (!allowedTypes || allowedTypes.length === 0) {
      return /\[System\.WorkItemType\]/i.test(wiql || '');
    }

    const wiqlStr = String(wiql || '');
    const eqRe = /\[System\.WorkItemType\]\s*=\s*'([^']+)'/gi;
    const inRe = /\[System\.WorkItemType\]\s+IN\s*\(([^)]+)\)/gi;

    const found = new Set<string>();
    let m: RegExpExecArray | null;
    while ((m = eqRe.exec(wiqlStr)) !== null) {
      found.add(m[1].trim().toLowerCase());
    }
    while ((m = inRe.exec(wiqlStr)) !== null) {
      const inner = m[1];
      for (const mm of inner.matchAll(/'([^']+)'/g)) {
        found.add(String(mm[1]).trim().toLowerCase());
      }
    }

    if (found.size === 0) return false;

    const allowed = new Set(allowedTypes.map((t) => String(t).toLowerCase()));
    for (const t of found) {
      if (!allowed.has(t)) return false;
    }
    return true;
  }

  /**
   * Matches flat query WIQL against an area path filter by checking any referenced [System.AreaPath].
   * Compares only the leaf segment of the path and performs a case-insensitive substring match.
   */
  private matchesFlatAreaCondition(wiql: string, areaFilter: string): boolean {
    const filter = String(areaFilter || '')
      .trim()
      .toLowerCase();
    if (!filter) return true;

    const wiqlLower = String(wiql || '').toLowerCase();
    // Capture any quoted value that appears in an expression mentioning [System.AreaPath]
    const re = /\[system\.areapath\][^']*'([^']+)'/gi;
    const paths: string[] = [];
    let match: RegExpExecArray | null;
    while ((match = re.exec(wiqlLower)) !== null) {
      paths.push(match[1]);
    }

    if (paths.length === 0) return false;

    const leaf = (p: string) => {
      const parts = p.split(/[\\/]/);
      return parts[parts.length - 1] || p;
    };

    return paths.some((p) => leaf(p).includes(filter));
  }

  // check if the query is a bug query
  private matchesBugCondition(wiql: string): boolean {
    return wiql.includes(`[System.WorkItemType] = 'Bug'`);
  }

  // Filter the fields based on the columns to filter map
  private filterFieldsByColumns(
    item: any,
    columnsToFilterMap: Map<string, string>,
    resultedRefNameMap: Map<string, string>
  ) {
    try {
      const parsedFields: any = {};
      for (const fieldName of Object.keys(item.fields)) {
        const value = item.fields[fieldName];

        // Always include System.WorkItemType and System.Title
        if (
          columnsToFilterMap.has(fieldName) ||
          fieldName === 'System.WorkItemType' ||
          fieldName === 'System.Title'
        ) {
          // If it's not in the map but we're including it anyway, add it to the resulted map
          if (!columnsToFilterMap.has(fieldName)) {
            resultedRefNameMap.set(fieldName, fieldName);
          } else {
            resultedRefNameMap.set(fieldName, columnsToFilterMap.get(fieldName) || '');
          }
          parsedFields[fieldName] = value;
        }
      }
      item.fields = { ...parsedFields };
    } catch (err: any) {
      logger.error(`Cannot filter columns: ${err.message}`);
      throw err;
    }
  }

  public async GetWorkItemTypeList(project: string) {
    try {
      let url = `${this.orgUrl}${project}/_apis/wit/workitemtypes?api-version=5.1`;
      const { value: workItemTypes } = await TFSServices.getItemContent(url, this.token);
      const workItemTypesWithIcons = await Promise.all(
        workItemTypes.map(async (workItemType: any) => {
          let iconDataUrl: string | null = null;
          const iconUrl = workItemType?.icon?.url;

          if (iconUrl) {
            const acceptHeaders = ['image/svg+xml', 'image/png'];
            for (const accept of acceptHeaders) {
              try {
                iconDataUrl = await TFSServices.fetchAzureDevOpsImageAsBase64(
                  iconUrl,
                  this.token,
                  'get',
                  {},
                  { Accept: accept }
                );
                if (iconDataUrl) break;
              } catch (error: any) {
                logger.warn(
                  `Failed to download icon (${accept}) for work item type ${
                    workItemType?.name ?? 'unknown'
                  }: ${error?.message || error}`
                );
              }
            }
          }

          const iconPayload = workItemType.icon
            ? { ...workItemType.icon, dataUrl: iconDataUrl }
            : iconDataUrl
            ? { id: undefined, url: undefined, dataUrl: iconDataUrl }
            : workItemType.icon;

          return {
            name: workItemType.name,
            referenceName: workItemType.referenceName,
            color: workItemType.color,
            icon: iconPayload,
            states: workItemType.states,
          };
        })
      );

      return workItemTypesWithIcons;
    } catch (err: any) {
      logger.error(`Error occurred during fetching work item types: ${err.message}`);
      throw err;
    }
  }

  /**
   * Fetches query results and categorizes Requirement work items by their Requirement Type value.
   *
   * Behavior:
   * - Detects Requirement items by checking if `System.WorkItemType` contains "requirement" (case-insensitive).
   * - Resolves candidate Requirement Type fields by calling the Work Item Type Fields API for the
   *   current project (Requirement type), selecting all fields whose display name (lowercased, underscores
   *   normalized to spaces) includes "requirement type". If available, `Microsoft.VSTS.CMMI.RequirementType`
   *   is prioritized first. If no candidates are found, falls back to that known reference name.
   * - For each Requirement, uses the first non-empty value among the discovered fields as the type value.
   * - Maps the value to a standardized category header; when unmapped or empty, places the item under
   *   "Other Requirements".
   * - When `Microsoft.VSTS.Common.Priority` equals 1, the item is also added to
   *   "Precedence and Criticality of Requirements".
   *
   * @param wiqlHref - The WIQL query URL to execute.
   * @returns An object containing the categorized requirements and total processed count.
   */
  async GetCategorizedRequirementsByType(wiqlHref: string): Promise<any> {
    try {
      if (!wiqlHref) {
        throw new Error('Incorrect WIQL Link');
      }

      logger.debug('Fetching query results for categorization');
      const queryResult: QueryTree = await TFSServices.getItemContent(wiqlHref, this.token);

      if (!queryResult) {
        throw new Error('Query result failed');
      }

      // Get work item IDs from the query result
      let workItemIds: number[] = [];

      if (queryResult.workItems && Array.isArray(queryResult.workItems)) {
        // Extract IDs from the workItems array
        workItemIds = queryResult.workItems.map((wi: any) => wi.id).filter((id: number) => id);
      } else if (queryResult.workItemRelations && Array.isArray(queryResult.workItemRelations)) {
        // Extract IDs from workItemRelations (for OneHop queries)
        const idSet = new Set<number>();
        queryResult.workItemRelations.forEach((rel: any) => {
          if (rel.source?.id) idSet.add(rel.source.id);
          if (rel.target?.id) idSet.add(rel.target.id);
        });
        workItemIds = Array.from(idSet);
      } else {
        logger.warn('No work items found in query result');
        return { categories: {}, totalCount: 0 };
      }

      if (workItemIds.length === 0) {
        logger.warn('No work item IDs extracted from query result');
        return { categories: {}, totalCount: 0 };
      }

      // Define the mapping from requirement type keys to standard headers
      const typeToHeaderMap: Record<string, string> = {
        Adaptation: 'Adaptation Requirements',
        'Computer Resource': 'Computer Resource Requirements',
        'System Environment': 'CSCI Environment Requirements',
        Constraints: 'Design and Implementation Constraints',
        'Design Constrains': 'Design and Implementation Constraints',
        'Physical Constraints': 'Design and Implementation Constraints',
        'External Interface': 'External Interfaces Requirements',
        'Internal Data': 'Internal Data Requirements',
        'Internal Interface': 'Internal Interfaces Requirements',
        Logistics: 'Logistics-Related Requirements',
        Packaging: 'Packaging Requirements',
        'Human Factors': 'Personnel-Related Requirements',
        Safety: 'Safety Requirements',
        Security: 'Security and Privacy Requirements',
        'Security and Privacy': 'Security and Privacy Requirements',
        'Quality of Service': 'Software Quality Factors',
        Reliability: 'Software Quality Factors',
        'System Quality Factors': 'Software Quality Factors',
        Training: 'Training-Related Requirements',
      };

      // Define the desired order of categories
      const categoryOrder = [
        'External Interfaces Requirements',
        'Internal Interfaces Requirements',
        'Internal Data Requirements',
        'Adaptation Requirements',
        'Safety Requirements',
        'Security and Privacy Requirements',
        'CSCI Environment Requirements',
        'Computer Resource Requirements',
        'Software Quality Factors',
        'Design and Implementation Constraints',
        'Personnel-Related Requirements',
        'Training-Related Requirements',
        'Logistics-Related Requirements',
        'Other Requirements',
        'Packaging Requirements',
        'Precedence and Criticality of Requirements',
      ];

      // Initialize all categories as empty arrays (for consistent ordering)
      const categorizedRequirements: Record<string, any[]> = {};
      categoryOrder.forEach((category) => {
        categorizedRequirements[category] = [];
      });

      const project = this.getProjectFromWiqlHref(wiqlHref);
      const reqTypeFieldRefNames: string[] = project
        ? await this.getRequirementTypeFieldRefs(project)
        : ['Microsoft.VSTS.CMMI.RequirementType'];

      // Process each work item
      for (const workItemId of workItemIds) {
        try {
          const wiUrl = `${this.orgUrl}_apis/wit/workitems/${workItemId}?$expand=All`;
          const fullWi = await TFSServices.getItemContent(wiUrl, this.token);

          const workItemType = fullWi.fields['System.WorkItemType'];
          if (!(typeof workItemType === 'string' && /requirement/i.test(workItemType))) {
            continue;
          }

          let trimmedType = '';
          for (const refName of reqTypeFieldRefNames) {
            const val = fullWi.fields ? fullWi.fields[refName] : undefined;
            if (val != null && String(val).trim() !== '') {
              trimmedType = String(val).trim();
              break;
            }
          }

          const categoryHeader = typeToHeaderMap[trimmedType]
            ? typeToHeaderMap[trimmedType]
            : 'Other Requirements';

          const requirementItem = {
            id: workItemId,
            title: fullWi.fields['System.Title'] || '',
            description:
              fullWi.fields['Microsoft.VSTS.CMMI.Symptom'] || fullWi.fields['System.Description'] || '',
            htmlUrl: fullWi._links?.html?.href || '',
          };

          categorizedRequirements[categoryHeader].push(requirementItem);

          const priority = fullWi.fields['Microsoft.VSTS.Common.Priority'];
          if (priority === 1) {
            categorizedRequirements['Precedence and Criticality of Requirements'].push(requirementItem);
          }
        } catch (err: any) {
          logger.warn(`Could not fetch work item ${workItemId}: ${err.message}`);
        }
      }

      // Sort items within each category by ID and remove empty categories
      const finalCategories: Record<string, any[]> = {};
      categoryOrder.forEach((category) => {
        const items = categorizedRequirements[category];
        if (items && items.length > 0) {
          finalCategories[category] = items.sort((a, b) => a.id - b.id);
        }
      });

      logger.debug(
        `Categorized ${workItemIds.length} work items into ${Object.keys(finalCategories).length} categories`
      );

      return {
        categories: finalCategories,
        totalCount: workItemIds.length,
      };
    } catch (err: any) {
      logger.error(`Could not fetch categorized requirements: ${err.message}`);
      throw err;
    }
  }

  /**
   * Helper method to flatten a tree structure into a flat array of work items
   */
  private flattenTreeToWorkItems(roots: any[]): any[] {
    const result: any[] = [];

    const traverse = (node: any) => {
      if (!node) return;

      result.push({
        id: node.id,
        title: node.title,
        description: node.description,
        htmlUrl: node.htmlUrl,
        url: node.htmlUrl, // Some nodes might use url instead
      });

      if (Array.isArray(node.children)) {
        node.children.forEach(traverse);
      }
    };

    roots.forEach(traverse);
    return result;
  }
}
