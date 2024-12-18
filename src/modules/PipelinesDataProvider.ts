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
    projectName: string,
    pipeLineId: string,
    toPipelineRunId: number,
    toPipeline: any,
    searchPrevPipelineFromDifferentCommit: boolean
  ) {
    // get pipeline runs:
    const pipelineRuns = await this.GetPipelineRunHistory(projectName, pipeLineId);

    for (const pipelineRun of pipelineRuns.value) {
      if (pipelineRun.id >= toPipelineRunId) {
        continue;
      }
      if (pipelineRun.result !== 'succeeded') {
        continue;
      }

      const fromPipeline = await this.getPipelineFromPipelineId(projectName, Number(pipeLineId));
      if (!fromPipeline.resources.repositories) {
        continue;
      }
      const fromPipelineRepositories = fromPipeline.resources.repositories;
      logger.debug(`from pipeline repositories ${JSON.stringify(fromPipelineRepositories)}`);
      const toPipelineRepositories = toPipeline.resources.repositories;
      logger.debug(`to pipeline repositories ${JSON.stringify(toPipelineRepositories)}`);

      const fromPipeLineSelfRepo =
        '__designer_repo' in fromPipelineRepositories
          ? fromPipelineRepositories['__designer_repo']
          : 'self' in fromPipelineRepositories
          ? fromPipelineRepositories['self']
          : undefined;
      const toPipeLineSelfRepo =
        '__designer_repo' in toPipelineRepositories
          ? toPipelineRepositories['__designer_repo']
          : 'self' in toPipelineRepositories
          ? toPipelineRepositories['self']
          : undefined;
      if (
        fromPipeLineSelfRepo.repository.id === toPipeLineSelfRepo.repository.id &&
        fromPipeLineSelfRepo.version === toPipeLineSelfRepo.version
      ) {
        if (searchPrevPipelineFromDifferentCommit) {
          continue;
        }
      }

      if (
        fromPipeLineSelfRepo.repository.id === toPipeLineSelfRepo.repository.id &&
        fromPipeLineSelfRepo.refName !== toPipeLineSelfRepo.refName
      ) {
        continue;
      }

      return pipelineRun.id;
    }
    return undefined;
  }

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

  async getPipelineFromPipelineId(projectName: string, buildId: number) {
    let url = `${this.orgUrl}${projectName}/_apis/build/builds/${buildId}`;
    return TFSServices.getItemContent(url, this.token, 'get');
  } //GetCommitForPipeline

  async getPipelineRunBuildById(projectName: string, pipelineId: number, runId: number) {
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

  async GetReleaseByReleaseId(projectName: string, releaseId: number): Promise<any> {
    let url = `${this.orgUrl}${projectName}/_apis/release/releases/${releaseId}`;
    if (url.startsWith('https://dev.azure.com')) {
      url = url.replace('https://dev.azure.com', 'https://vsrm.dev.azure.com');
    }
    return TFSServices.getItemContent(url, this.token, 'get', null, null);
  }

  async GetPipelineRunHistory(projectName: string, pipelineId: string) {
    let url: string = `${this.orgUrl}${projectName}/_apis/pipelines/${pipelineId}/runs`;
    let res: any = await TFSServices.getItemContent(url, this.token, 'get', null, null);
    //Filter successful builds only
    let { value } = res;
    if (value) {
      const successfulRunHistory = value.filter((run: any) => run.result === 'succeeded');
      return { count: successfulRunHistory.length, value: successfulRunHistory };
    }
    return res;
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
}
