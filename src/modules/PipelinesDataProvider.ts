import { PipelineRun, Repository, ResourceRepository } from '../models/tfs-data';
import { TFSServices } from '../helpers/tfs';

import logger from '../utils/logger';
import GitDataProvider from './GitDataProvider';

export default class PipelinesDataProvider {
  orgUrl: string = '';
  token: string = '';

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
      if (!fromPipeline.resources.repositories) {
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
    const fromRepo =
      fromPipeline.resources.repositories[0].self || fromPipeline.resources.repositories.__designer_repo;
    const targetRepo =
      targetPipeline.resources.repositories[0].self || targetPipeline.resources.repositories.__designer_repo;

    if (fromRepo.repository.id !== targetRepo.repository.id) {
      return false;
    }

    if (fromRepo.version === targetRepo.version) {
      return !searchPrevPipelineFromDifferentCommit;
    }

    return fromRepo.refName === targetRepo.refName;
  }

  /**
   * Retrieves a set of pipeline resources from a given pipeline run object.
   *
   * @param inPipeline - The pipeline run object containing resources.
   * @returns A promise that resolves to an array of unique pipeline resource objects.
   *
   * The function performs the following steps:
   * 1. Initializes an empty set to store unique pipeline resources.
   * 2. Checks if the input pipeline has any resources of type pipelines.
   * 3. Iterates over each pipeline resource and processes it.
   * 4. Fixes the URL of the pipeline resource to match the build API format.
   * 5. Fetches the build details using the fixed URL.
   * 6. If the build response is valid and matches the criteria, adds the pipeline resource to the set.
   * 7. Returns an array of unique pipeline resources.
   *
   * The returned pipeline resource object contains the following properties:
   * - name: The alias name of the resource pipeline.
   * - buildId: The ID of the resource pipeline.
   * - definitionId: The ID of the build definition.
   * - buildNumber: The build number.
   * - teamProject: The name of the team project.
   * - provider: The type of repository provider.
   *
   * @throws Will log an error message if there is an issue fetching the pipeline resource.
   */
  public async getPipelineResourcePipelinesFromObject(inPipeline: PipelineRun) {
    const resourcePipelines: Set<any> = new Set();

    if (!inPipeline.resources.pipelines) {
      return resourcePipelines;
    }
    const pipelines = inPipeline.resources.pipelines;

    const pipelineEntries = Object.entries(pipelines);

    await Promise.all(
      pipelineEntries.map(async ([resourcePipelineAlias, resource]) => {
        const resourcePipelineObj = (resource as any).pipeline;
        const resourcePipelineName = resourcePipelineAlias;
        let urlBeforeFix = resourcePipelineObj.url;
        urlBeforeFix = urlBeforeFix.substring(0, urlBeforeFix.indexOf('?revision'));
        const fixedUrl = urlBeforeFix.replace('/_apis/pipelines/', '/_apis/build/builds/');
        let buildResponse: any;
        try {
          buildResponse = await TFSServices.getItemContent(fixedUrl, this.token, 'get');
        } catch (err: any) {
          logger.error(`Error fetching pipeline ${resourcePipelineName} : ${err.message}`);
        }
        if (
          buildResponse &&
          buildResponse.definition.type === 'build' &&
          buildResponse.repository.type === 'TfsGit'
        ) {
          let resourcePipelineToAdd = {
            name: resourcePipelineName,
            buildId: resourcePipelineObj.id,
            definitionId: buildResponse.definition.id,
            buildNumber: buildResponse.buildNumber,
            teamProject: buildResponse.project.name,
            provider: buildResponse.repository.type,
          };
          if (!resourcePipelines.has(resourcePipelineToAdd)) {
            resourcePipelines.add(resourcePipelineToAdd);
          }
        }
      })
    );

    return [...resourcePipelines];
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
  ) {
    const resourceRepositories: Set<any> = new Set();

    if (!inPipeline.resources.repositories) {
      return resourceRepositories;
    }
    const repositories = inPipeline.resources.repositories;
    for (const prop in repositories) {
      const resourceRepo = repositories[prop];
      if (resourceRepo.repository.type !== 'azureReposGit') {
        continue;
      }
      const repoId = resourceRepo.repository.id;

      const repo: Repository = await gitDataProviderInstance.GetGitRepoFromRepoId(repoId);
      const resourceRepository: ResourceRepository = {
        repoName: repo.name,
        repoSha1: resourceRepo.version,
        url: repo.url,
      };
      if (!resourceRepositories.has(resourceRepository)) {
        resourceRepositories.add(resourceRepository);
      }
    }
    return [...resourceRepositories];
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
          (run: any) => run.result !== 'failed' || run.result !== 'canceled'
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
