import { TFSServices } from '../helpers/tfs';
import logger from '../utils/logger';

export type AccessStatus = 'ok' | 'denied' | 'notFound' | 'error';

/** Outcome of one read the credential attempted. `count` is how many items it could SEE. */
export interface AccessArea {
  status: AccessStatus;
  httpStatus?: number;
  count?: number;
}

export interface ProjectAccess {
  project: AccessArea;
  repositories: AccessArea;
  workItems: AccessArea;
  builds: AccessArea;
  releases: AccessArea;
  testPlans: AccessArea;
}

const PROBE_TIMEOUT_MS = 10_000;

/**
 * Turns a failed read into an access outcome. 401/403 are a denial; 404 is "not found" (it is also what a
 * server without that API answers, so it is not reported as a denial); anything else (a timeout, a 5xx) is
 * an error that says nothing about permissions.
 */
export const accessOutcomeOfError = (error: any): AccessArea => {
  // TFSServices rethrows Azure DevOps' "could not be found" answers as a plain Error without a status
  // ("File not found or insufficient permissions: <url>"); that is a 404 in effect.
  if (/^File not found or insufficient permissions/i.test(String(error?.message || ''))) {
    return { status: 'notFound', httpStatus: 404 };
  }
  const httpStatus = Number(error?.response?.status ?? error?.status);
  if (httpStatus === 401 || httpStatus === 403) return { status: 'denied', httpStatus };
  if (httpStatus === 404) return { status: 'notFound', httpStatus };
  return { status: 'error', ...(Number.isFinite(httpStatus) ? { httpStatus } : {}) };
};

export default class MangementDataProvider {
  orgUrl: string = '';
  token: string = '';

  constructor(orgUrl: string, token: string) {
    this.orgUrl = orgUrl;
    this.token = token;
  }

  async GetCllectionLinkTypes() {
    let url: string = `${this.orgUrl}_apis/wit/workitemrelationtypes`;
    let res: any = await TFSServices.getItemContent(url, this.token, 'get', null, null);
    return res;
  }

  //get all projects
  async GetProjects(): Promise<any> {
    let projectUrl: string = `${this.orgUrl}_apis/projects?$top=1000`;
    let projects: any = await TFSServices.getItemContent(projectUrl, this.token);
    return projects;
  }

  // get project by  name return project object
  async GetProjectByName(projectName: string): Promise<any> {
    try {
      let projects: any = await this.GetProjects();
      // Safely handle cases where the projects payload is missing or malformed
      if (!projects || !Array.isArray(projects.value)) {
        logger.error('Projects response does not contain a valid value array');
        return {};
      }
      for (let i = 0; i < projects.value.length; i++) {
        if (projects.value[i].name === projectName) return projects.value[i];
      }
      return {};
    } catch (err) {
      logger.error(err);
      return {};
    }
  }

  // get project by id return project object
  async GetProjectByID(projectID: string): Promise<any> {
    let projectUrl: string = `${this.orgUrl}_apis/projects/${projectID}`;
    let project: any = await TFSServices.getItemContent(projectUrl, this.token);
    return project;
  }

  async GetUserProfile(): Promise<any> {
    let url: string = `${this.orgUrl}_api/_common/GetUserProfile?__v=5`;
    return TFSServices.getItemContent(url, this.token);
  }

  // Check if organization URL is valid and optionally validate PAT
  // Without token: checks if organization URL exists
  // With token: checks both URL validity AND PAT validity
  async CheckOrgUrlValidity(token?: string): Promise<any> {
    let url: string = `${this.orgUrl}_apis/connectionData`;
    // Use provided token or empty string for URL-only validation
    return TFSServices.getItemContent(url, token || '', 'get', null, null, false);
  }

  // Resolves the AD domain/account for an on-prem, Windows-integrated-auth
  // identity via the documented Identities REST API. Works with both PAT
  // and bearer tokens — TFSServices.getItemContent/applyAuth already
  // auto-detects which one it was given. printError=false: failure here is
  // the EXPECTED, common case (any cloud/Entra-backed org), not a real error.
  async GetIdentityById(identityId: string): Promise<any> {
    const url = `${this.orgUrl}_apis/identities?identityIds=${encodeURIComponent(
      identityId
    )}&queryMembership=None&api-version=6.0`;
    return TFSServices.getItemContent(url, this.token, 'get', null, null, false);
  }

  /**
   * What THIS credential can actually see in a project, measured with the kind of reads the SVD and STD
   * make. Effective access rather than permission bits: Azure DevOps often answers a reader without access
   * with an empty or shorter list instead of a 403, so the counts matter as much as the status. Read-only,
   * parallel and bounded; a 403 here is an expected outcome, so nothing is logged as an error.
   */
  async ProbeProjectAccess(projectName: string, timeoutMs: number = PROBE_TIMEOUT_MS): Promise<ProjectAccess> {
    const read = (url: string, method = 'get', body: any = {}, headers: any = {}) =>
      TFSServices.getItemContent(url, this.token, method, body, headers, false);
    const org = this.orgUrl;
    const vsrm = (url: string) => (url.startsWith('https://dev.azure.com') ? url.replace('https://dev.azure.com', 'https://vsrm.dev.azure.com') : url);

    const probe = async (run: () => Promise<AccessArea>): Promise<AccessArea> => {
      let timer: NodeJS.Timeout | undefined;
      try {
        return await Promise.race([
          run(),
          new Promise<AccessArea>((resolve) => {
            timer = setTimeout(() => resolve({ status: 'error' }), timeoutMs);
          }),
        ]);
      } catch (error: any) {
        return accessOutcomeOfError(error);
      } finally {
        if (timer) clearTimeout(timer);
      }
    };
    const listCount = (data: any): number => (Array.isArray(data?.value) ? data.value.length : Number(data?.count) || 0);

    const [project, repositories, workItems, builds, releases, testPlans] = await Promise.all([
      // Asked for directly (name or id): a project list is paged and may not contain it on the first page.
      probe(async () => {
        await read(`${org}_apis/projects/${encodeURIComponent(String(projectName))}`);
        return { status: 'ok' };
      }),
      probe(async () => ({ status: 'ok', count: listCount(await read(`${org}${projectName}/_apis/git/repositories`)) })),
      probe(async () => {
        const result = await read(
          `${org}${projectName}/_apis/wit/wiql?$top=1`,
          'post',
          { query: 'Select [System.Id] From WorkItems Where [System.TeamProject] = @project' },
          { 'Content-Type': 'application/json' }
        );
        return { status: 'ok', count: Array.isArray(result?.workItems) ? result.workItems.length : 0 };
      }),
      probe(async () => ({ status: 'ok', count: listCount(await read(`${org}${projectName}/_apis/build/definitions?$top=200`)) })),
      probe(async () => ({ status: 'ok', count: listCount(await read(vsrm(`${org}${projectName}/_apis/release/definitions?$top=200`))) })),
      probe(async () => ({ status: 'ok', count: listCount(await read(`${org}${projectName}/_apis/test/plans?$top=200`)) })),
    ]);
    return { project, repositories, workItems, builds, releases, testPlans };
  }
}
