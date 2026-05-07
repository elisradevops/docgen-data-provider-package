import { PipelineRun, Repository, ResourceRepository } from '../models/tfs-data';
import { TFSServices } from '../helpers/tfs';

import logger from '../utils/logger';
import GitDataProvider from './GitDataProvider';
const pLimit = require('p-limit');
const MAX_DISCOVERY_PAGES = 50;

type PreviousDiscoveryResult =
  | { status: 'found'; id: number }
  | { status: 'not_found' }
  | { status: 'failed'; error: unknown };

export default class PipelinesDataProvider {
  orgUrl: string = '';
  token: string = '';
  private projectNameByIdCache: Map<string, string> = new Map();

  constructor(orgUrl: string, token: string) {
    this.orgUrl = orgUrl;
    this.token = token;
  }

  /**
   * Returns work item references associated with one completed build.
   *
   * Used by template-only first-run SVD baseline generation, where there is no
   * previous successful build to compare against and the current target build's
   * linked work items become the baseline content.
   */
  public async GetBuildWorkItems(teamProject: string, buildId: number): Promise<any[]> {
    const encodedProject = encodeURIComponent(teamProject);
    const url = `${this.orgUrl}${encodedProject}/_apis/build/builds/${buildId}/workitems?$top=2000&api-version=6.0`;
    const result = await TFSServices.getItemContent(url, this.token, 'get');
    return Array.isArray(result?.value) ? result.value : [];
  }

  /**
   * Resolves the previous pipeline run used as the source side of an SVD pipeline range.
   *
   * For regular pipeline ranges, the Builds API is used first because it supports paging,
   * completed/succeeded filters, and branch filtering. Stage-specific discovery keeps the
   * older Pipelines Runs path because stage status must be checked from run details.
   *
   * @param teamProject Azure DevOps project name.
   * @param pipelineId Pipeline/build definition id.
   * @param toPipelineRunId Target run/build id. Candidates must be older than this id.
   * @param targetPipeline Full target pipeline run details, used for repository/branch matching.
   * @param fromStage Optional stage name. When set, only previous runs with this successful stage match.
   */
  public async findPreviousPipeline(
    teamProject: string,
    pipelineId: string,
    toPipelineRunId: number,
    targetPipeline: any,
    fromStage: string = ''
  ) {
    if (!fromStage) {
      const previousBuildId = await this.findPreviousSuccessfulBuild(
        teamProject,
        pipelineId,
        toPipelineRunId,
        targetPipeline
      );
      if (previousBuildId) {
        return previousBuildId;
      }
    }

    const pipelineRuns = await this.GetPipelineRunHistory(teamProject, pipelineId);
    if (!pipelineRuns?.value) {
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

      if (this.isMatchingPipeline(fromPipeline, targetPipeline)) {
        return pipelineRun.id;
      }
    }
    return undefined;
  }

  /**
   * Finds the previous successful completed build for a definition.
   *
   * Discovery order:
   * 1. Same-branch search (preferred): pages the Builds API filtered to succeeded builds on the
   *    same branch as the target run.
   * 2. Ancestry-walk fallback: if no same-branch result, finds the merge-base between the
   *    target commit and the default branch, then returns the latest default-branch build whose
   *    sourceVersion is an ancestor of that merge-base. Useful for feature-branch first builds
   *    that have never had a prior same-branch success.
   * 3. Returns undefined if neither step finds a candidate (caller falls through to baseline).
   *
   * @returns Previous build id, or undefined when no valid candidate is found.
   */
  public async findPreviousSuccessfulBuild(
    teamProject: string,
    definitionId: string,
    toBuildId: number,
    targetPipeline: any
  ): Promise<number | undefined> {
    const targetRepo = this.getPrimaryPipelineRepository(targetPipeline);
    const targetBranch = this.normalizeBranchName(targetRepo?.refName);

    if (targetBranch) {
      const sameBranchResult = await this.findPreviousSuccessfulBuildPage(
        teamProject,
        definitionId,
        toBuildId,
        targetPipeline,
        targetBranch
      );
      if (sameBranchResult.status === 'failed') {
        throw sameBranchResult.error;
      }
      if (sameBranchResult.status === 'found') {
        return sameBranchResult.id;
      }
    }

    const ancestryId = await this.findAncestryFallbackBuild(
      teamProject,
      definitionId,
      toBuildId,
      targetPipeline
    );
    if (ancestryId !== undefined) {
      return ancestryId;
    }

    return undefined;
  }

  /**
   * Pages the Builds API until a valid previous successful build is found.
   *
   * The API query already asks for completed/succeeded builds ordered by finish time, but
   * candidates are validated locally as well to protect callers from incomplete API data,
   * mocks, or future response-shape differences.
   */
  private async findPreviousSuccessfulBuildPage(
    teamProject: string,
    definitionId: string,
    toBuildId: number,
    targetPipeline: any,
    branchName?: string
  ): Promise<PreviousDiscoveryResult> {
    const targetRepo = this.getPrimaryPipelineRepository(targetPipeline);
    let continuationToken: string | undefined = undefined;
    let pageCount = 0;
    do {
      let url = `${this.orgUrl}${teamProject}/_apis/build/builds?definitions=${encodeURIComponent(
        String(definitionId)
      )}&resultFilter=succeeded&statusFilter=completed&queryOrder=finishTimeDescending&$top=200&api-version=6.0`;
      if (branchName) {
        url += `&branchName=${encodeURIComponent(branchName)}`;
      }
      if (continuationToken) {
        url += `&continuationToken=${encodeURIComponent(continuationToken)}`;
      }

      try {
        const { data, headers } = await TFSServices.getItemContentWithHeaders(
          url,
          this.token,
          'get',
          null,
          null
        );
        pageCount++;
        const builds: any[] = data?.value || [];
        const match = builds.find((build: any) =>
          this.isMatchingPreviousBuild(
            build,
            targetRepo,
            toBuildId,
            branchName
          )
        );
        if (match?.id) {
          return { status: 'found', id: Number(match.id) };
        }
        continuationToken = this.getContinuationToken(headers);
        if (continuationToken && pageCount >= MAX_DISCOVERY_PAGES) {
          return {
            status: 'failed',
            error: new Error(`Pipeline discovery exceeded ${MAX_DISCOVERY_PAGES} pages`),
          };
        }
      } catch (err: unknown) {
        logger.warn(`Could not fetch previous successful builds: ${this.getErrorMessage(err)}`);
        return { status: 'failed', error: err };
      }
    } while (continuationToken);

    return { status: 'not_found' };
  }

  /**
   * Ancestry-walk fallback for findPreviousSuccessfulBuild.
   *
   * Used when no same-branch successful build exists. Resolves the merge-base between the
   * target commit and the repo's default branch, then finds the latest default-branch build
   * whose sourceVersion is an ancestor of that merge-base.
   *
   * Returns undefined (never throws) so the caller can fall through to baseline SVD mode.
   */
  private async findAncestryFallbackBuild(
    teamProject: string,
    definitionId: string,
    toBuildId: number,
    targetPipeline: any
  ): Promise<number | undefined> {
    try {
      const targetRepo = this.getPrimaryPipelineRepository(targetPipeline);
      const repoId = targetRepo?.repository?.id;
      const targetSha = targetRepo?.version;

      if (!repoId || !targetSha) {
        return undefined;
      }

      const defaultBranch = await this.getRepoDefaultBranch(teamProject, repoId);
      if (!defaultBranch) {
        return undefined;
      }
      const normalizedDefault = this.normalizeBranchName(defaultBranch);
      if (!normalizedDefault) {
        return undefined;
      }

      const mergeBase = await this.getMergeBase(teamProject, repoId, defaultBranch, targetSha);
      if (!mergeBase) {
        return undefined;
      }

      logger.debug(
        `[ancestry] target=${targetSha.substring(0, 7)} defaultBranch=${defaultBranch} mergeBase=${mergeBase.substring(0, 7)}`
      );

      let continuationToken: string | undefined;
      let pageCount = 0;

      do {
        // encodeURIComponent(teamProject) is used here (and in getRepoDefaultBranch / getMergeBase)
        // for consistency within the ancestry helpers. findPreviousSuccessfulBuildPage uses a bare
        // teamProject segment — both forms are accepted by ADO, but they should be unified in a
        // future cleanup pass.
        let url =
          `${this.orgUrl}${encodeURIComponent(teamProject)}/_apis/build/builds` +
          `?definitions=${encodeURIComponent(String(definitionId))}` +
          `&resultFilter=succeeded&statusFilter=completed` +
          `&queryOrder=finishTimeDescending&$top=200&api-version=6.0` +
          `&branchName=${encodeURIComponent(normalizedDefault)}`;
        if (continuationToken) {
          url += `&continuationToken=${encodeURIComponent(continuationToken)}`;
        }

        const { data, headers } = await TFSServices.getItemContentWithHeaders(
          url,
          this.token,
          'get',
          null,
          null
        );
        pageCount++;
        const builds: any[] = data?.value || [];

        for (const build of builds) {
          if (!this.isMatchingPreviousBuild(build, targetRepo, toBuildId, normalizedDefault)) {
            continue;
          }
          const candidateSha: string | undefined = build.sourceVersion;
          if (!candidateSha) continue;

          const isAncestor = await this.isCommitAncestorOf(
            teamProject,
            repoId,
            candidateSha,
            mergeBase
          );
          if (isAncestor) {
            logger.debug(
              `[ancestry] selected build ${build.id} sourceVersion=${candidateSha.substring(0, 7)}`
            );
            return Number(build.id);
          }
        }

        continuationToken = this.getContinuationToken(headers);
        if (continuationToken && pageCount >= MAX_DISCOVERY_PAGES) {
          logger.warn(`[ancestry] fallback exceeded ${MAX_DISCOVERY_PAGES} pages without match`);
          return undefined;
        }
      } while (continuationToken);

      return undefined;
    } catch (err: unknown) {
      logger.warn(`[ancestry] fallback failed: ${this.getErrorMessage(err)}`);
      return undefined;
    }
  }

  /**
   * Returns the default branch name (e.g. "refs/heads/main") for a Git repository.
   * Returns undefined if the repository cannot be fetched.
   */
  private async getRepoDefaultBranch(
    teamProject: string,
    repoId: string
  ): Promise<string | undefined> {
    const url = `${this.orgUrl}${encodeURIComponent(teamProject)}/_apis/git/repositories/${repoId}?api-version=6.0`;
    try {
      const result = await TFSServices.getItemContent(url, this.token, 'get', null, null, false);
      return typeof result?.defaultBranch === 'string' && result.defaultBranch
        ? result.defaultBranch
        : undefined;
    } catch {
      return undefined;
    }
  }

  /**
   * Returns the merge-base commit SHA between a branch tip and a target commit SHA.
   *
   * Uses the ADO Git diffs/commits API:
   *   baseVersion=<branchShortName> (branch type), targetVersion=<commitSha> (commit type)
   * The returned commonCommit is the merge-base.
   */
  private async getMergeBase(
    teamProject: string,
    repoId: string,
    defaultBranch: string,
    targetSha: string
  ): Promise<string | undefined> {
    const branchShort = defaultBranch.replace(/^refs\/heads\//, '');
    const url =
      `${this.orgUrl}${encodeURIComponent(teamProject)}/_apis/git/repositories/${repoId}/diffs/commits` +
      `?baseVersion=${encodeURIComponent(branchShort)}&baseVersionType=branch` +
      `&targetVersion=${encodeURIComponent(targetSha)}&targetVersionType=commit` +
      `&$top=1&api-version=6.0`;
    try {
      const result = await TFSServices.getItemContent(url, this.token, 'get', null, null, false);
      return typeof result?.commonCommit === 'string' && result.commonCommit
        ? result.commonCommit
        : undefined;
    } catch {
      return undefined;
    }
  }

  /**
   * Returns true when candidateSha is an ancestor of (or equal to) targetSha.
   *
   * Uses the ADO Git diffs/commits API:
   *   baseVersion=candidateSha (commit), targetVersion=targetSha (commit)
   * If candidateSha is an ancestor of targetSha, it IS the common ancestor of the two,
   * so commonCommit === candidateSha.
   */
  private async isCommitAncestorOf(
    teamProject: string,
    repoId: string,
    candidateSha: string,
    targetSha: string
  ): Promise<boolean> {
    if (candidateSha === targetSha) return true;
    const url =
      `${this.orgUrl}${encodeURIComponent(teamProject)}/_apis/git/repositories/${repoId}/diffs/commits` +
      `?baseVersion=${encodeURIComponent(candidateSha)}&baseVersionType=commit` +
      `&targetVersion=${encodeURIComponent(targetSha)}&targetVersionType=commit` +
      `&$top=1&api-version=6.0`;
    try {
      const result = await TFSServices.getItemContent(url, this.token, 'get', null, null, false);
      return result?.commonCommit === candidateSha;
    } catch {
      return false;
    }
  }

  private getContinuationToken(headers: any): string | undefined {
    return headers?.['x-ms-continuationtoken'] || headers?.['x-ms-continuation-token'] || undefined;
  }

  /**
   * Extracts the primary repository from a pipeline run details response.
   *
   * The ADO Pipelines Runs API returns resources.repositories as a named-key object where
   * 'self' is the pipeline's own checkout (e.g. { self: {...}, AppRepo: {...} }). Some older
   * in-process representations wrap it as an array ({ 0: { self: {...} } }) and classic/designer
   * pipelines use __designer_repo.
   */
  private getPrimaryPipelineRepository(pipeline: any): any {
    const repositories = pipeline?.resources?.repositories;
    if (!repositories) return undefined;
    if (repositories.self) return repositories.self;
    if (repositories.__designer_repo) return repositories.__designer_repo;
    // resources.repositories is a plain named-key object (not an array), so [0] is always
    // undefined. Fall back to the first value for pipelines using a custom checkout alias.
    // Some older ADO pipeline formats wrap the repo under a nested .self key; unwrap if present.
    const first = Object.values(repositories)[0] as any;
    return first?.self ?? first ?? undefined;
  }

  /**
   * Validates a Builds API candidate against the target run.
   *
   * A candidate must be older than the target, completed successfully, from the same
   * repository, and optionally from the required branch.
   */
  private isMatchingPreviousBuild(
    build: any,
    targetRepo: any,
    toBuildId: number,
    requiredBranch?: string
  ): boolean {
    const buildId = Number(build?.id);
    if (!Number.isFinite(buildId) || buildId >= toBuildId) return false;
    if (build?.status && build.status !== 'completed') return false;
    if (build?.result && build.result !== 'succeeded') return false;

    const buildRepoId = build?.repository?.id;
    const targetRepoId = targetRepo?.repository?.id;
    if (!buildRepoId || !targetRepoId || buildRepoId !== targetRepoId) return false;

    if (requiredBranch && build?.sourceBranch !== requiredBranch) return false;

    return true;
  }

  /**
   * Finds the previous successful release/deployment for an SVD release range.
   *
   * Release discovery pages the Release List API and expands environments so candidates can
   * be validated by deployment status. A matching release must be older than the target,
   * active, and have at least one succeeded environment.
   *
   * @returns Previous release id, or undefined when no valid candidate is found.
   */
  public async findPreviousSuccessfulRelease(
    projectName: string,
    definitionId: string,
    toReleaseId: number
  ): Promise<number | undefined> {
    let result = await this.findSuccessfulReleasePage(
      projectName,
      definitionId,
      '7.1',
      (release) => this.isPreviousSuccessfulRelease(release, toReleaseId),
      'previous'
    );
    if (result.status === 'failed' && this.isUnsupportedApiVersionError(result.error)) {
      result = await this.findSuccessfulReleasePage(
        projectName,
        definitionId,
        '6.0',
        (release) => this.isPreviousSuccessfulRelease(release, toReleaseId),
        'previous'
      );
    }
    if (result.status === 'failed') {
      throw result.error;
    }
    return result.status === 'found' ? result.id : undefined;
  }

  /**
   * Finds the latest successful release/deployment for an SVD release range.
   *
   * Used when the release template omits `toReleaseId`. The same candidate rule is used as
   * previous-release discovery: active release with at least one succeeded environment.
   *
   * @returns Latest release id, or undefined when no valid candidate is found.
   */
  public async findLatestSuccessfulRelease(
    projectName: string,
    definitionId: string
  ): Promise<number | undefined> {
    let result = await this.findSuccessfulReleasePage(
      projectName,
      definitionId,
      '7.1',
      (release) => this.isSuccessfulRelease(release),
      'latest'
    );
    if (result.status === 'failed' && this.isUnsupportedApiVersionError(result.error)) {
      result = await this.findSuccessfulReleasePage(
        projectName,
        definitionId,
        '6.0',
        (release) => this.isSuccessfulRelease(release),
        'latest'
      );
    }
    if (result.status === 'failed') {
      throw result.error;
    }
    return result.status === 'found' ? result.id : undefined;
  }

  private async findSuccessfulReleasePage(
    projectName: string,
    definitionId: string,
    apiVersion: string,
    isMatch: (release: any) => boolean,
    discoveryLabel: 'previous' | 'latest'
  ): Promise<PreviousDiscoveryResult> {
    let baseUrl: string = `${this.orgUrl}${projectName}/_apis/release/releases?definitionId=${encodeURIComponent(
      String(definitionId)
    )}&queryOrder=descending&$top=200&$expand=environments&api-version=${apiVersion}`;
    if (baseUrl.startsWith('https://dev.azure.com')) {
      baseUrl = baseUrl.replace('https://dev.azure.com', 'https://vsrm.dev.azure.com');
    }

    let continuationToken: string | undefined = undefined;
    let pageCount = 0;
    do {
      let url = baseUrl;
      if (continuationToken) {
        url += `&continuationToken=${encodeURIComponent(continuationToken)}`;
      }

      try {
        const { data, headers } = await TFSServices.getItemContentWithHeaders(
          url,
          this.token,
          'get',
          null,
          null
        );
        pageCount++;
        const releases: any[] = data?.value || [];
        const match = releases.find(isMatch);
        if (match?.id) {
          return { status: 'found', id: Number(match.id) };
        }
        continuationToken = this.getContinuationToken(headers);
        if (continuationToken && pageCount >= MAX_DISCOVERY_PAGES) {
          return {
            status: 'failed',
            error: new Error(`Release discovery exceeded ${MAX_DISCOVERY_PAGES} pages`),
          };
        }
      } catch (err: unknown) {
        logger.warn(`Could not fetch ${discoveryLabel} successful releases: ${this.getErrorMessage(err)}`);
        return { status: 'failed', error: err };
      }
    } while (continuationToken);

    return { status: 'not_found' };
  }

  private isUnsupportedApiVersionError(err: unknown): boolean {
    const error = err as any;
    const responseData = error?.response?.data;
    const message = [
      error?.message,
      responseData?.message,
      typeof responseData === 'string' ? responseData : JSON.stringify(responseData || ''),
    ]
      .join(' ')
      .toLowerCase();
    return (
      message.includes('api-version') &&
      (message.includes('unsupported') ||
        message.includes('not support') ||
        message.includes('not supported'))
    );
  }

  private getErrorMessage(err: unknown): string {
    return err instanceof Error ? err.message : String(err);
  }

  /**
   * Validates a release candidate for auto-discovery.
   *
   * A release is considered successful for this SVD range purpose when at least one
   * deployment environment succeeded. This mirrors the existing release artifact flow,
   * where a release may contain multiple environments but still provide usable artifacts.
   */
  private isPreviousSuccessfulRelease(release: any, toReleaseId: number): boolean {
    const releaseId = Number(release?.id);
    if (!Number.isFinite(releaseId) || releaseId >= toReleaseId) return false;
    return this.isSuccessfulRelease(release);
  }

  private isSuccessfulRelease(release: any): boolean {
    const releaseId = Number(release?.id);
    if (!Number.isFinite(releaseId)) return false;
    if (release?.status && String(release.status).toLowerCase() !== 'active') return false;

    const environments = Array.isArray(release?.environments) ? release.environments : [];
    return environments.some((environment: any) => {
      return String(environment?.status || '').toLowerCase() === 'succeeded';
    });
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
   * Determines if two pipelines match based on their repository and branch.
   *
   * @param fromPipeline - The source pipeline to compare.
   * @param targetPipeline - The target pipeline to compare against.
   * @returns `true` if the pipelines share the same repository and ref; otherwise, `false`.
   */
  private isMatchingPipeline(
    fromPipeline: PipelineRun,
    targetPipeline: PipelineRun
  ): boolean {
    if (!fromPipeline?.resources?.repositories || !targetPipeline?.resources?.repositories) {
      return false;
    }

    const fromRepo =
      fromPipeline.resources.repositories.self ||
      fromPipeline.resources.repositories[0]?.self ||
      fromPipeline.resources.repositories.__designer_repo;
    const targetRepo =
      targetPipeline.resources.repositories.self ||
      targetPipeline.resources.repositories[0]?.self ||
      targetPipeline.resources.repositories.__designer_repo;

    if (!fromRepo?.repository?.id || !targetRepo?.repository?.id) {
      return false;
    }

    if (fromRepo.repository.id !== targetRepo.repository.id) {
      return false;
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

  private tryBuildBuildApiUrlFromPipelinesApiUrl(url?: string): string | undefined {
    if (!url) return undefined;
    try {
      const parsed = new URL(url);
      parsed.search = '';
      const before = parsed.pathname;
      const after = before.replace('/_apis/pipelines/', '/_apis/build/builds/');
      if (after === before) return undefined;
      parsed.pathname = after;
      return parsed.toString();
    } catch {
      return undefined;
    }
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
      } catch (e1: any) {}
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
        `Error resolving runId by runName for pipeline ${pipelineId} (${projectName}/${desired}): ${
          e?.message || e
        }`
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
      !!buildResponse && definitionType === 'build' && (repoType === 'TfsGit' || repoType === 'azureReposGit')
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
      logger.debug('getPipelineResourcePipelinesFromObject: no pipeline resources on run');
      return [];
    }
    const pipelineEntries = Object.entries(inPipeline.resources.pipelines);
    logger.debug(
      `getPipelineResourcePipelinesFromObject: resolving ${pipelineEntries.length} pipeline resources`
    );

    const concurrencyLimit = pLimit(8);
    await Promise.all(
      pipelineEntries.map(([resourcePipelineAlias, resource]) => concurrencyLimit(async () => {
        const resourcePipelineObj = (resource as any)?.pipeline;
        const pipelineIdCandidate = Number(resourcePipelineObj?.id);

        const rawProjectName =
          (resource as any)?.project?.name ||
          this.tryGetTeamProjectFromAzureDevOpsUrl(resourcePipelineObj?.url) ||
          this.tryGetTeamProjectFromAzureDevOpsUrl(inPipeline?.url);
        const projectName = await this.normalizeProjectName(rawProjectName);

        if (!Number.isFinite(pipelineIdCandidate)) {
          logger.warn(
            `getPipelineResourcePipelinesFromObject: resource ${resourcePipelineAlias} missing numeric pipeline id`
          );
          return;
        }

        let buildResponse: any;
        const fixedUrl = this.tryBuildBuildApiUrlFromPipelinesApiUrl(resourcePipelineObj?.url);
        if (fixedUrl) {
          try {
            const fixedBuildResponse = await TFSServices.getItemContent(
              fixedUrl,
              this.token,
              'get',
              null,
              null
            );
            if (this.isSupportedResourcePipelineBuild(fixedBuildResponse)) {
              buildResponse = fixedBuildResponse;
            }
          } catch (err: any) {
            logger.error(
              `Error fetching pipeline ${resourcePipelineAlias} via fixed build url ${fixedUrl} : ${err.message}`
            );
          }
        } else {
          logger.warn(
            `getPipelineResourcePipelinesFromObject: could not convert pipelines url to builds url for ${resourcePipelineAlias}: ${String(
              resourcePipelineObj?.url || ''
            )}`
          );
        }

        // PowerShell-style logic: do not fall back to runId/buildNumber/run-history resolution.
        if (!this.isSupportedResourcePipelineBuild(buildResponse)) {
          logger.debug(
            `getPipelineResourcePipelinesFromObject: skipping ${resourcePipelineAlias} because resolved build is missing/unsupported (project=${String(
              projectName || ''
            )}, pipelineIdCandidate=${pipelineIdCandidate})`
          );
          return;
        }

        if (this.isSupportedResourcePipelineBuild(buildResponse)) {
          const buildNumber = String(buildResponse.buildNumber || '').trim();
          if (!buildNumber) {
            logger.warn(
              `Resource pipeline ${resourcePipelineAlias} resolved to buildId=${String(
                buildResponse?.id
              )} but buildNumber is missing, skipping`
            );
            return;
          }

          const expectedProjectName = String(projectName || '').trim();
          const actualProjectName = String(buildResponse?.project?.name || '').trim();
          if (
            expectedProjectName &&
            actualProjectName &&
            expectedProjectName.toLowerCase() !== actualProjectName.toLowerCase()
          ) {
            logger.warn(
              `Resource pipeline ${resourcePipelineAlias} resolved to buildId=${String(
                buildResponse?.id
              )} but project mismatch expected=${expectedProjectName} actual=${actualProjectName}, skipping`
            );
            return;
          }

          const expectedPipelineName = String(resourcePipelineObj?.name || '').trim();
          const actualPipelineName = String(buildResponse?.definition?.name || '').trim();
          if (
            expectedPipelineName &&
            actualPipelineName &&
            expectedPipelineName.toLowerCase() !== actualPipelineName.toLowerCase()
          ) {
            logger.warn(
              `Resource pipeline ${resourcePipelineAlias} resolved to buildId=${String(
                buildResponse?.id
              )} but pipeline name mismatch expected=${expectedPipelineName} actual=${actualPipelineName}, skipping`
            );
            return;
          }
          const definitionId = buildResponse.definition?.id ?? pipelineIdCandidate;
          const buildId = buildResponse.id;
          logger.debug(
            `getPipelineResourcePipelinesFromObject: resolved ${resourcePipelineAlias} -> ${String(
              buildResponse?.project?.name || projectName || ''
            )}/${definitionId} runId=${buildId} repoType=${String(buildResponse?.repository?.type || '')}`
          );
          const resourcePipelineToAdd = {
            name: resourcePipelineObj?.name ?? resourcePipelineAlias,
            buildId,
            definitionId,
            buildNumber,
            teamProject: buildResponse.project?.name ?? projectName,
            provider: buildResponse.repository?.type,
          };
          const key = `${resourcePipelineToAdd.teamProject}:${resourcePipelineToAdd.definitionId}:${resourcePipelineToAdd.buildId}:${resourcePipelineToAdd.name}`;
          if (!resourcePipelinesByKey.has(key)) resourcePipelinesByKey.set(key, resourcePipelineToAdd);
        }
      }))
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
  /**
   * Rebuilds a repository API URL using the project name rather than the UUID that TFS
   * returns in repo.url. On-prem TFS indexes WI-to-commit associations by project name, so
   * commitsbatch requests must use the name form to receive populated workItems arrays.
   */
  private buildRepoApiUrl(repo: Repository): string {
    const projectName = repo.project?.name;
    return projectName
      ? `${this.orgUrl}${encodeURIComponent(projectName)}/_apis/git/repositories/${repo.id}`
      : repo.url;
  }

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
      if (!['azureReposGit', 'TfsGit'].includes(resourceRepo?.repository?.type)) {
        continue;
      }
      const repoId = resourceRepo.repository.id;
      if (!repoId) continue;

      const repo: Repository = await gitDataProviderInstance.GetGitRepoFromRepoId(repoId);
      const rawProjectName = repo.project?.name || repo.project?.id;
      const resolvedProjectName = await this.normalizeProjectName(rawProjectName);
      if (resolvedProjectName && repo.project) {
        repo.project.name = resolvedProjectName;
      }
      const repoApiUrl = this.buildRepoApiUrl(repo);
      const resourceRepository: ResourceRepository = {
        repoName: repo.name,
        repoSha1: resourceRepo.version,
        url: repoApiUrl,
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
    let url = `${this.orgUrl}${projectName}/_apis/pipelines/${pipelineId}/runs/${runId}?$expand=resources`;
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

  /**
   * Fetches the first page of release history for a definition.
   *
   * Kept for backward compatibility with existing consumers. New range/discovery flows
   * should prefer GetAllReleaseHistory when full release history is required.
   */
  async GetReleaseHistory(projectName: string, definitionId: string) {
    let url: string = `${this.orgUrl}${projectName}/_apis/release/releases?definitionId=${definitionId}&$top=200`;
    if (url.startsWith('https://dev.azure.com')) {
      url = url.replace('https://dev.azure.com', 'https://vsrm.dev.azure.com');
    }
    let res: any = await TFSServices.getItemContent(url, this.token, 'get', null, null);
    return res;
  }

  /**
   * Fetches all releases for a definition using continuation tokens.
   *
   * This is used by SVD release range handling because the requested from/to releases may be
   * older than the first page returned by Azure DevOps.
   *
   * @param range When provided, pagination stops as soon as both fromId and toId are present in
   *   the accumulated results. ADO returns releases in descending id order, so once the smallest
   *   id seen is <= the lower requested id both endpoints are guaranteed to be loaded.
   */
  async GetAllReleaseHistory(projectName: string, definitionId: string, range?: { fromId: number; toId: number }) {
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
        const { data, headers } = await TFSServices.getItemContentWithHeaders(
          url,
          this.token,
          'get',
          null,
          null
        );
        const { value = [] } = data || {};
        all.push(...value);
        page++;
        logger.debug(`GetAllReleaseHistory: fetched page ${page}, cumulative ${all.length} releases`);

        // Stop-early: once the smallest id in the accumulated list is <= the lower
        // requested id, both endpoints of the range are present.
        if (range && all.length > 0) {
          const lowerBound = Math.min(range.fromId, range.toId);
          const minId = Math.min(...all.map((r: any) => Number(r.id)));
          if (minId <= lowerBound) {
            logger.debug(`GetAllReleaseHistory: stop-early at page ${page} (minId=${minId} <= lowerBound=${lowerBound})`);
            break;
          }
        }

        // Azure DevOps returns continuation token header for next page
        continuationToken = this.getContinuationToken(headers);
      } catch (err: any) {
        logger.error(`GetAllReleaseHistory failed: ${err.message}`);
        throw err;
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
