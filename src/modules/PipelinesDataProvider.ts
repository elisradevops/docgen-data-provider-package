import { PipelineRun, Repository, ResourceRepository } from '../models/tfs-data';
import { TFSServices } from '../helpers/tfs';

import logger from '../utils/logger';
import GitDataProvider from './GitDataProvider';

export default class PipelinesDataProvider {
  orgUrl: string = '';
  token: string = '';
  private projectNameByIdCache: Map<string, string> = new Map();

  constructor(orgUrl: string, token: string) {
    this.orgUrl = orgUrl;
    this.token = token;
  }

  public async findPreviousPipeline(
    teamProject: string,
    pipelineId: string,
    toPipelineRunId: number,
    targetPipeline: any,
    searchPrevPipelineFromDifferentCommit: boolean,
    fromStage: string = ''
  ) {
    const pipelineRuns = await this.GetPipelineRunHistory(teamProject, pipelineId);
    if (!pipelineRuns.value) {
      return undefined;
    }

    for (const pipelineRun of pipelineRuns.value) {
      if (this.isInvalidPipelineRun(pipelineRun, toPipelineRunId, fromStage)) {
        continue;
      }

      if (fromStage && !(await this.isStageSuccessful(pipelineRun, teamProject, fromStage))) {
        continue;
      }

      const fromPipeline = await this.getPipelineRunDetails(teamProject, Number(pipelineId), pipelineRun.id);
      if (!fromPipeline?.resources?.repositories) {
        continue;
      }

      if (this.isMatchingPipeline(fromPipeline, targetPipeline, searchPrevPipelineFromDifferentCommit)) {
        return pipelineRun.id;
      }
    }
    return undefined;
  }

  /**
   * Determines if a pipeline run is invalid based on various conditions.
   *
   * @param pipelineRun - The pipeline run object to evaluate.
   * @param toPipelineRunId - The pipeline run ID to compare against.
   * @param fromStage - The stage from which the pipeline run originated.
   * @returns `true` if the pipeline run is considered invalid, `false` otherwise.
   */
  private isInvalidPipelineRun(pipelineRun: any, toPipelineRunId: number, fromStage: string): boolean {
    return (
      pipelineRun.id >= toPipelineRunId ||
      ['canceled', 'failed', 'canceling'].includes(pipelineRun.result) ||
      (pipelineRun.result === 'unknown' && !fromStage) ||
      (pipelineRun.result !== 'succeeded' && !fromStage)
    );
  }

  /**
   * Checks if a specific stage in a pipeline run was successful.
   *
   * @param pipelineRun - The pipeline run object containing details of the run.
   * @param teamProject - The name of the team project.
   * @param fromStage - The name of the stage to check.
   * @returns A promise that resolves to a boolean indicating whether the stage was successful.
   */
  private async isStageSuccessful(
    pipelineRun: any,
    teamProject: string,
    fromStage: string
  ): Promise<boolean> {
    const fromPipelineStage = await this.getPipelineStageName(pipelineRun, teamProject, fromStage);
    return (
      fromPipelineStage && fromPipelineStage.state === 'completed' && fromPipelineStage.result === 'succeeded'
    );
  }

  /**
   * Determines if two pipelines match based on their repository and version information.
   *
   * @param fromPipeline - The source pipeline to compare.
   * @param targetPipeline - The target pipeline to compare against.
   * @param searchPrevPipelineFromDifferentCommit - A flag indicating whether to search for a previous pipeline from a different commit.
   * @returns `true` if the pipelines match based on the repository and version criteria; otherwise, `false`.
   */
  private isMatchingPipeline(
    fromPipeline: PipelineRun,
    targetPipeline: PipelineRun,
    searchPrevPipelineFromDifferentCommit: boolean
  ): boolean {
    if (!fromPipeline?.resources?.repositories || !targetPipeline?.resources?.repositories) {
      return false;
    }

    const fromRepo =
      fromPipeline.resources.repositories[0]?.self || fromPipeline.resources.repositories.__designer_repo;
    const targetRepo =
      targetPipeline.resources.repositories[0]?.self || targetPipeline.resources.repositories.__designer_repo;

    if (!fromRepo?.repository?.id || !targetRepo?.repository?.id) {
      return false;
    }

    if (fromRepo.repository.id !== targetRepo.repository.id) {
      return false;
    }

    if (fromRepo.version === targetRepo.version) {
      return !searchPrevPipelineFromDifferentCommit;
    }

    return fromRepo.refName === targetRepo.refName;
  }

  private tryGetTeamProjectFromAzureDevOpsUrl(url?: string): string | undefined {
    if (!url) return undefined;
    try {
      const parsed = new URL(url);
      const parts = parsed.pathname.split('/').filter(Boolean);
      const apiIndex = parts.findIndex((p) => p === '_apis');
      if (apiIndex <= 0) return undefined;
      return parts[apiIndex - 1];
    } catch {
      return undefined;
    }
  }

  /**
   * Returns `true` when the input looks like an Azure DevOps GUID identifier.
   */
  private isGuidLike(value?: string): boolean {
    if (!value) return false;
    return /^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/.test(
      String(value).trim()
    );
  }

  /**
   * Normalizes a project identifier into a project name suitable for `{project}` URL path segments.
   *
   * Azure DevOps APIs often accept both project names and IDs, but some endpoints (especially on ADO Server)
   * behave differently or return empty results when a GUID is used in the URL path. This method converts
   * a GUID into its canonical project name via `/_apis/projects/{id}` and caches the mapping.
   *
   * @param projectNameOrId Project name (e.g. "Test CMMI") or project GUID.
   * @returns The project name, or the original input when resolution fails.
   */
  private async normalizeProjectName(projectNameOrId?: string): Promise<string | undefined> {
    if (!projectNameOrId) return undefined;
    const raw = String(projectNameOrId).trim();
    if (!raw) return undefined;
    if (!this.isGuidLike(raw)) return raw;

    const cached = this.projectNameByIdCache.get(raw);
    if (cached) return cached;

    // ADO supports querying projects at the collection/org root.
    const url = `${this.orgUrl}_apis/projects/${encodeURIComponent(raw)}?api-version=6.0`;
    try {
      const project = await TFSServices.getItemContent(url, this.token, 'get', null, null, false);
      const resolvedName = String(project?.name || '').trim();
      if (resolvedName) {
        this.projectNameByIdCache.set(raw, resolvedName);
        return resolvedName;
      }
      return raw;
    } catch (err: any) {
      return raw;
    }
  }

  /**
   * Attempts to extract a run/build id from an Azure DevOps URL.
   *
   * Supports URLs shaped like:
   * - `.../_apis/pipelines/{pipelineId}/runs/{runId}`
   * - `.../_apis/build/builds/{buildId}`
   */
  private tryParseRunIdFromUrl(url?: string): number | undefined {
    if (!url) return undefined;
    try {
      const parsed = new URL(url);
      const parts = parsed.pathname.split('/').filter(Boolean);
      const runsIndex = parts.findIndex((p) => p === 'runs');
      if (runsIndex >= 0 && parts[runsIndex + 1]) {
        const runId = Number(parts[runsIndex + 1]);
        return Number.isFinite(runId) ? runId : undefined;
      }
      const buildsIndex = parts.findIndex((p) => p === 'builds');
      if (buildsIndex >= 0 && parts[buildsIndex + 1]) {
        const buildId = Number(parts[buildsIndex + 1]);
        return Number.isFinite(buildId) ? buildId : undefined;
      }
      return undefined;
    } catch {
      return undefined;
    }
  }

  /**
   * Normalizes a branch name into a `refs/...` form accepted by the Builds API.
   */
  private normalizeBranchName(branch?: string): string | undefined {
    if (!branch) return undefined;
    const b = String(branch).trim();
    if (!b) return undefined;
    if (b.startsWith('refs/')) return b;
    if (b.startsWith('heads/')) return `refs/${b}`;
    return `refs/heads/${b}`;
  }

  /**
   * Searches for a build by build number without restricting by definition id.
   *
   * Used as a fallback when the `resources.pipelines[alias].pipeline.id` is not a stable build definition id
   * (some ADO instances return a run/build id or a pipeline revision-related id instead).
   *
   * @param projectName Team project name.
   * @param buildNumber Build number / run name (e.g. "20251225.2", "1.0.56").
   * @param branch Optional branch filter.
   * @param expectedDefinitionName Optional build definition name to disambiguate results (typically YAML `source`).
   */
  private async findBuildByBuildNumber(
    projectName: string,
    buildNumber: string,
    branch?: string,
    expectedDefinitionName?: string
  ): Promise<any | undefined> {
    const bn = String(buildNumber || '').trim();
    if (!bn) return undefined;
    const normalizedBranch = this.normalizeBranchName(branch);
    let url = `${this.orgUrl}${projectName}/_apis/build/builds?buildNumber=${encodeURIComponent(
      bn
    )}&$top=20&queryOrder=finishTimeDescending&api-version=6.0`;
    if (normalizedBranch) {
      url += `&branchName=${encodeURIComponent(normalizedBranch)}`;
    }
    try {
      const res: any = await TFSServices.getItemContent(url, this.token, 'get', null, null);
      const value: any[] = res?.value || [];
      const filtered = expectedDefinitionName
        ? value.filter((b: any) => String(b?.definition?.name || '') === String(expectedDefinitionName))
        : value;
      return filtered[0] ?? value[0];
    } catch (err: any) {
      logger.error(
        `Error resolving build by buildNumber (no definition) (${projectName}/${bn}): ${err?.message || err}`
      );
      return undefined;
    }
  }

  /**
   * Retrieves a build by id, with an additional fallback for ADO instances that allow the non-project-scoped route.
   *
   * @param projectName Optional team project name.
   * @param buildId Build id.
   */
  private async tryGetBuildByIdWithFallback(
    projectName: string | undefined,
    buildId: number
  ): Promise<any | undefined> {
    if (projectName) {
      try {
        return await this.getPipelineBuildByBuildId(projectName, buildId);
      } catch (e1: any) {
      }
    }

    // Some ADO instances allow build lookup without the {project} path segment.
    try {
      const url = `${this.orgUrl}_apis/build/builds/${buildId}`;
      return await TFSServices.getItemContent(url, this.token, 'get', null, null);
    } catch (e2: any) {
      return undefined;
    }
  }

  /**
   * Searches for a build by definition id + build number.
   *
   * @param projectName Team project name.
   * @param definitionId Build definition id (classic build definition id).
   * @param buildNumber Build number / run name.
   * @param branch Optional branch filter.
   */
  private async findBuildByDefinitionAndBuildNumber(
    projectName: string,
    definitionId: number,
    buildNumber: string,
    branch?: string
  ): Promise<any | undefined> {
    const bn = String(buildNumber || '').trim();
    if (!bn) return undefined;
    const normalizedBranch = this.normalizeBranchName(branch);
    let url = `${this.orgUrl}${projectName}/_apis/build/builds?definitions=${encodeURIComponent(
      String(definitionId)
    )}&buildNumber=${encodeURIComponent(bn)}&$top=1&queryOrder=finishTimeDescending&api-version=6.0`;
    if (normalizedBranch) {
      url += `&branchName=${encodeURIComponent(normalizedBranch)}`;
    }
    try {
      const res: any = await TFSServices.getItemContent(url, this.token, 'get', null, null);
      const value: any[] = res?.value || [];
      return value[0];
    } catch (err: any) {
      logger.error(
        `Error resolving build by buildNumber for definition ${definitionId} (${projectName}/${bn}): ${err.message}`
      );
      return undefined;
    }
  }

  /**
   * Resolves a pipelines API run id by comparing a desired run name against the pipeline run history.
   *
   * Useful when Builds API lookups don't return results even though the upstream run exists.
   *
   * @param projectName Team project name.
   * @param pipelineId Pipeline id.
   * @param runName Run "name" from ADO UI (e.g. "20251225.2").
   */
  private async findRunIdByPipelineRunName(
    projectName: string,
    pipelineId: number,
    runName: string
  ): Promise<number | undefined> {
    const desired = String(runName || '').trim();
    if (!desired) return undefined;
    try {
      const history = await this.GetPipelineRunHistory(projectName, String(pipelineId));
      const runs: any[] = history?.value || [];
      const match = runs.find((r: any) => {
        const name = String(r?.name || '').trim();
        const id = String(r?.id || '').trim();
        return name === desired || name === `#${desired}` || id === desired;
      });
      if (match?.id) {
        return Number(match.id);
      }
      return undefined;
    } catch (e: any) {
      logger.error(
        `Error resolving runId by runName for pipeline ${pipelineId} (${projectName}/${desired}): ${e?.message || e}`
      );
      return undefined;
    }
  }

  /**
   * Attempts to infer a run/build id for a pipeline resource from the run payload fields.
   *
   * Resolution order:
   * 1) `resource.runId` (preferred)
   * 2) numeric `resource.version` (only if it's an integer string/number)
   * 3) parse from `resource.pipeline.url` if it contains `/runs/{id}` or `/builds/{id}`
   */
  private inferRunIdFromPipelineResource(resource: any, pipelineUrl?: string): number | undefined {
    const explicit = Number(resource?.runId);
    if (Number.isFinite(explicit)) return explicit;

    const version = resource?.version;
    if (typeof version === 'number' && Number.isFinite(version)) return version;
    if (typeof version === 'string' && /^\d+$/.test(version)) return Number(version);

    const parsed = this.tryParseRunIdFromUrl(pipelineUrl);
    return Number.isFinite(parsed) ? Number(parsed) : undefined;
  }

  /**
   * Resolves a pipeline resource run id from a non-numeric `version` (run name/build number).
   *
   * Fall back order (when `runId` isn't present):
   * 1) Builds API: definitionId + buildNumber
   * 2) Builds API: buildNumber-only (optionally filtered by YAML `source`)
   * 3) Pipelines API: run history match by `name` / `#name`
   * 4) Heuristic: treat `pipelineIdCandidate` as buildId and fetch that build
   */
  private async resolveRunIdFromVersion(params: {
    projectName: string | undefined;
    pipelineIdCandidate: number;
    version: unknown;
    branch?: string;
    source?: unknown;
  }): Promise<number | undefined> {
    const { projectName, pipelineIdCandidate, version, branch, source } = params;

    if (!Number.isFinite(pipelineIdCandidate)) return undefined;

    if (!projectName) {
      const buildById = await this.tryGetBuildByIdWithFallback(undefined, pipelineIdCandidate);
      return buildById?.id ? Number(buildById.id) : undefined;
    }

    if (typeof version !== 'string' || !String(version).trim()) return undefined;
    const buildNumber = String(version).trim();

    const buildByNumber = await this.findBuildByDefinitionAndBuildNumber(
      projectName,
      pipelineIdCandidate,
      buildNumber,
      branch
    );
    if (buildByNumber?.id) return Number(buildByNumber.id);

    const buildByNumberAny = await this.findBuildByBuildNumber(
      projectName,
      buildNumber,
      branch,
      typeof source === 'string' ? source : undefined
    );
    if (buildByNumberAny?.id) return Number(buildByNumberAny.id);

    const runIdByName = await this.findRunIdByPipelineRunName(projectName, pipelineIdCandidate, buildNumber);
    if (runIdByName) return runIdByName;

    const buildById = await this.tryGetBuildByIdWithFallback(projectName, pipelineIdCandidate);
    return buildById?.id ? Number(buildById.id) : undefined;
  }

  private isSupportedResourcePipelineBuild(buildResponse: any): boolean {
    const definitionType = buildResponse?.definition?.type;
    const repoType = buildResponse?.repository?.type;
    return (
      !!buildResponse &&
      definitionType === 'build' &&
      (repoType === 'TfsGit' || repoType === 'azureReposGit')
    );
  }

  /**
   * Extracts and resolves pipeline resources (`resources.pipelines`) from a pipeline run.
   *
   * Azure DevOps represents pipeline dependencies as "pipeline resources". Those resources do not always include
   * a concrete `runId`, so this method resolves them into build-backed objects that can be used for recursion
   * (e.g., release notes / SVD traversal).
   *
   * @param inPipeline Pipeline run payload that contains `resources.pipelines`.
   * @returns A collection of normalized pipeline resource objects (array), or an empty Set when there are no resources.
   *
   * Complexity: O(p) for p pipeline resources (network I/O dominates).
   *
   * Returned object shape:
   * - `name`: resource alias
   * - `buildId`: resolved build/run id
   * - `definitionId`: build definition id (classic build definition id)
   * - `buildNumber`: build number / run name
   * - `teamProject`: resolved project name
   * - `provider`: repo provider type (e.g., `TfsGit`)
   */
  public async getPipelineResourcePipelinesFromObject(inPipeline: PipelineRun) {
    const resourcePipelinesByKey: Map<string, any> = new Map();

    if (!inPipeline?.resources?.pipelines) {
      return new Set();
    }
    const pipelineEntries = Object.entries(inPipeline.resources.pipelines);

    await Promise.all(
      pipelineEntries.map(async ([resourcePipelineAlias, resource]) => {
        const resourcePipelineObj = (resource as any)?.pipeline;
        const pipelineIdCandidate = Number(resourcePipelineObj?.id);
        const source = (resource as any)?.source;
        const version = (resource as any)?.version;
        const branch = (resource as any)?.branch;

        const rawProjectName =
          (resource as any)?.project?.name ||
          this.tryGetTeamProjectFromAzureDevOpsUrl(resourcePipelineObj?.url) ||
          this.tryGetTeamProjectFromAzureDevOpsUrl(inPipeline?.url);
        const projectName = await this.normalizeProjectName(rawProjectName);

        if (!Number.isFinite(pipelineIdCandidate)) {
          return;
        }

        let runId: number | undefined = this.inferRunIdFromPipelineResource(resource, resourcePipelineObj?.url);
        if (typeof runId !== 'number' || !Number.isFinite(runId)) {
          runId = await this.resolveRunIdFromVersion({
            projectName,
            pipelineIdCandidate,
            version,
            branch,
            source,
          });
        }

        if (typeof runId !== 'number' || !Number.isFinite(runId)) {
          return;
        }

        let buildResponse: any;
        try {
          buildResponse = await this.tryGetBuildByIdWithFallback(projectName, runId);
        } catch (err: any) {
          logger.error(`Error fetching pipeline ${resourcePipelineAlias} run ${runId} : ${err.message}`);
        }
        if (this.isSupportedResourcePipelineBuild(buildResponse)) {
          const definitionId = buildResponse.definition?.id ?? pipelineIdCandidate;
          const buildId = buildResponse.id ?? runId;
          const resourcePipelineToAdd = {
            name: resourcePipelineAlias,
            buildId,
            definitionId,
            buildNumber: buildResponse.buildNumber ?? (resource as any)?.runName ?? (resource as any)?.version,
            teamProject: buildResponse.project?.name ?? projectName,
            provider: buildResponse.repository?.type,
          };
          const key = `${resourcePipelineToAdd.teamProject}:${resourcePipelineToAdd.definitionId}:${resourcePipelineToAdd.buildId}:${resourcePipelineToAdd.name}`;
          if (!resourcePipelinesByKey.has(key)) resourcePipelinesByKey.set(key, resourcePipelineToAdd);
        }
      })
    );
    return [...resourcePipelinesByKey.values()];
  }

  /**
   * Retrieves a set of resource repositories from a given pipeline object.
   *
   * @param inPipeline - The pipeline run object containing resource information.
   * @param gitDataProviderInstance - An instance of GitDataProvider to fetch repository details.
   * @returns A promise that resolves to an array of unique resource repositories.
   */
  public async getPipelineResourceRepositoriesFromObject(
    inPipeline: PipelineRun,
    gitDataProviderInstance: GitDataProvider
  ): Promise<ResourceRepository[]> {
    const resourceRepositoriesById: Map<string, ResourceRepository> = new Map();

    const repositories = inPipeline?.resources?.repositories;
    if (!repositories) {
      return [];
    }

    for (const prop in repositories) {
      const resourceRepo = repositories[prop];
      if (resourceRepo?.repository?.type !== 'azureReposGit') {
        continue;
      }
      const repoId = resourceRepo.repository.id;
      if (!repoId) continue;

      const repo: Repository = await gitDataProviderInstance.GetGitRepoFromRepoId(repoId);
      const resourceRepository: ResourceRepository = {
        repoName: repo.name,
        repoSha1: resourceRepo.version,
        url: repo.url,
      };
      const key = String(repoId);
      if (!resourceRepositoriesById.has(key)) {
        resourceRepositoriesById.set(key, resourceRepository);
      }
    }

    return [...resourceRepositoriesById.values()];
  }

  /**
   * Retrieves the details of a specific pipeline build by its build ID.
   *
   * @param projectName - The name of the project that contains the pipeline.
   * @param buildId - The unique identifier of the build to retrieve.
   * @returns A promise that resolves to the content of the build details.
   */
  async getPipelineBuildByBuildId(projectName: string, buildId: number) {
    let url = `${this.orgUrl}${projectName}/_apis/build/builds/${buildId}`;
    return TFSServices.getItemContent(url, this.token, 'get');
  } //GetCommitForPipeline

  /**
   * Retrieves the details of a specific pipeline run.
   *
   * @param projectName - The name of the project containing the pipeline.
   * @param pipelineId - The ID of the pipeline.
   * @param runId - The ID of the pipeline run.
   * @returns A promise that resolves to the content of the pipeline run.
   */
  async getPipelineRunDetails(projectName: string, pipelineId: number, runId: number): Promise<PipelineRun> {
    let url = `${this.orgUrl}${projectName}/_apis/pipelines/${pipelineId}/runs/${runId}`;
    return TFSServices.getItemContent(url, this.token);
  }

  async TriggerBuildById(projectName: string, buildDefanitionId: string, parameter: any) {
    let data = {
      definition: {
        id: buildDefanitionId,
      },
      parameters: parameter, //'{"Test":"123"}'
    };
    logger.info(JSON.stringify(data));
    let url = `${this.orgUrl}${projectName}/_apis/build/builds?api-version=5.0`;
    let res = await TFSServices.postRequest(url, this.token, 'post', data, null);
    return res;
  }

  /**
   * Retrieves an artifact by build ID from a specified project.
   *
   * @param {string} projectName - The name of the project.
   * @param {string} buildId - The ID of the build.
   * @param {string} artifactName - The name of the artifact to retrieve.
   * @returns {Promise<any>} A promise that resolves to the artifact data.
   * @throws Will throw an error if the retrieval process fails.
   *
   * @example
   * const artifact = await GetArtifactByBuildId('MyProject', '12345', 'MyArtifact');
   * console.log(artifact);
   */
  async GetArtifactByBuildId(projectName: string, buildId: string, artifactName: string): Promise<any> {
    try {
      logger.info(
        `Get artifactory from project ${projectName},BuildId ${buildId} artifact name ${artifactName}`
      );
      logger.info(`Check if build ${buildId} have artifact`);
      let url = `${this.orgUrl}${projectName}/_apis/build/builds/${buildId}/artifacts`;
      let response = await TFSServices.getItemContent(url, this.token, 'Get', null, null);
      if (response.count == 0) {
        logger.info(`No artifact for build ${buildId} was published `);
        return response;
      }
      url = `${this.orgUrl}${projectName}/_apis/build/builds/${buildId}/artifacts?artifactName=${artifactName}`;
      let res = await TFSServices.getItemContent(url, this.token, 'Get', null, null);
      logger.info(`Url for download :${res.resource.downloadUrl}`);
      let result = await TFSServices.downloadZipFile(res.resource.downloadUrl, this.token);
      return result;
    } catch (err) {
      logger.error(`Error : ${err}`);
      throw new Error(String(err));
    }
  }

  /**
   * Retrieves a release by its release ID for a given project.
   *
   * @param projectName - The name of the project.
   * @param releaseId - The ID of the release to retrieve.
   * @returns A promise that resolves to the release data.
   */
  async GetReleaseByReleaseId(projectName: string, releaseId: number): Promise<any> {
    let url = `${this.orgUrl}${projectName}/_apis/release/releases/${releaseId}`;
    if (url.startsWith('https://dev.azure.com')) {
      url = url.replace('https://dev.azure.com', 'https://vsrm.dev.azure.com');
    }
    return TFSServices.getItemContent(url, this.token, 'get', null, null);
  }

  /**
   * Retrieves the run history of a specified pipeline within a project.
   *
   * @param projectName - The name of the project containing the pipeline.
   * @param pipelineId - The ID of the pipeline to retrieve the run history for.
   * @returns An object containing the count of successful runs and an array of successful run details.
   * @throws Will log an error message if the pipeline run history could not be fetched.
   */
  async GetPipelineRunHistory(projectName: string, pipelineId: string) {
    try {
      let url: string = `${this.orgUrl}${projectName}/_apis/pipelines/${pipelineId}/runs`;
      let res: any = await TFSServices.getItemContent(url, this.token, 'get', null, null);
      //Filter successful builds only
      let { value } = res;
      if (value) {
        const successfulRunHistory = value.filter(
          (run: any) => run.result !== 'failed' && run.result !== 'canceled'
        );
        return { count: successfulRunHistory.length, value: successfulRunHistory };
      }
      return res;
    } catch (err: any) {
      logger.error(`Could not fetch Pipeline Run History: ${err.message}`);
    }
  }

  async GetReleaseHistory(projectName: string, definitionId: string) {
    let url: string = `${this.orgUrl}${projectName}/_apis/release/releases?definitionId=${definitionId}&$top=200`;
    if (url.startsWith('https://dev.azure.com')) {
      url = url.replace('https://dev.azure.com', 'https://vsrm.dev.azure.com');
    }
    let res: any = await TFSServices.getItemContent(url, this.token, 'get', null, null);
    return res;
  }

  /**
   * Fetch all releases for a definition using continuation tokens.
   */
  async GetAllReleaseHistory(projectName: string, definitionId: string) {
    let baseUrl: string = `${this.orgUrl}${projectName}/_apis/release/releases?definitionId=${definitionId}&api-version=6.0`;
    if (baseUrl.startsWith('https://dev.azure.com')) {
      baseUrl = baseUrl.replace('https://dev.azure.com', 'https://vsrm.dev.azure.com');
    }

    const all: any[] = [];
    let continuationToken: string | undefined = undefined;
    let page = 0;
    do {
      let url = `${baseUrl}&$top=200`;
      if (continuationToken) {
        url += `&continuationToken=${encodeURIComponent(continuationToken)}`;
      }
      try {
        const { data, headers } = await TFSServices.getItemContentWithHeaders(url, this.token, 'get', null, null);
        const { value = [] } = data || {};
        all.push(...value);
        // Azure DevOps returns continuation token header for next page
        continuationToken = headers?.['x-ms-continuationtoken'] || headers?.['x-ms-continuation-token'] || undefined;
        page++;
        logger.debug(`GetAllReleaseHistory: fetched page ${page}, cumulative ${all.length} releases`);
      } catch (err: any) {
        logger.error(`GetAllReleaseHistory failed: ${err.message}`);
        break;
      }
    } while (continuationToken);

    return { count: all.length, value: all };
  }

  async GetAllPipelines(projectName: string) {
    let url: string = `${this.orgUrl}${projectName}/_apis/pipelines?$top=2000`;
    let res: any = await TFSServices.getItemContent(url, this.token, 'get', null, null);
    return res;
  }

  async GetAllReleaseDefenitions(projectName: string) {
    let url: string = `${this.orgUrl}${projectName}/_apis/release/definitions?$top=2000`;
    if (url.startsWith('https://dev.azure.com')) {
      url = url.replace('https://dev.azure.com', 'https://vsrm.dev.azure.com');
    }
    let res: any = await TFSServices.getItemContent(url, this.token, 'get', null, null);
    return res;
  }

  async GetRecentReleaseArtifactInfo(projectName: string) {
    let artifactInfo: any[] = [];
    let url: string = `${this.orgUrl}${projectName}/_apis/release/releases?$top=1&api-version=6.0`;
    if (url.startsWith('https://dev.azure.com')) {
      url = url.replace('https://dev.azure.com', 'https://vsrm.dev.azure.com');
    }
    let res: any = await TFSServices.getItemContent(url, this.token);
    const { value: releases } = res;
    if (releases && releases.length > 0) {
      const releaseId = releases[0].id;
      let url: string = `${this.orgUrl}${projectName}/_apis/release/releases/${releaseId}?api-version=6.0`;
      const releaseResponse = await TFSServices.getItemContent(url, this.token);
      if (releaseResponse) {
        const { artifacts } = releaseResponse;
        for (const artifact of artifacts) {
          const { definition, version } = artifact.definitionReference;
          artifactInfo.push({ artifactName: definition.name, artifactVersion: version.name });
        }
      }
    }
    return artifactInfo;
  }

  /**
   * Get stage name
   * @param pipelineRunId requested pipeline run id
   * @param teamProject requested team project
   * @param stageName stage name to search for in the pipeline
   * @returns
   */
  private async getPipelineStageName(pipelineRunId: number, teamProject: string, stageName: string) {
    let url = `${this.orgUrl}${teamProject}/_apis/build/builds/${pipelineRunId}/timeline?api-version=6.0`;
    try {
      const getPipelineLogsResponse = await TFSServices.getItemContent(url, this.token, 'get');

      const { records } = getPipelineLogsResponse;
      for (const record of records) {
        if (record.type === 'Stage' && record.name === stageName) {
          return record;
        }
      }
    } catch (err: any) {
      logger.error(`Error fetching pipeline ${pipelineRunId} with url ${url} : ${err.message}`);
      return undefined;
    }
  }
}
