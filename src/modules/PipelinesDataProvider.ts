import { TFSServices } from "../helpers/tfs";

import logger from "../utils/logger";

export default class PipelinesDataProvider {
  orgUrl: string = "";
  token: string = "";

  constructor(orgUrl: string, token: string) {
    this.orgUrl = orgUrl;
    this.token = token;
  }

  async getPipelineFromPipelineId(projectName: string, buildId: number) {
    let url = `${this.orgUrl}${projectName}/_apis/build/builds/${buildId}`;
    return TFSServices.getItemContent(url, this.token, "get");
  } //GetCommitForPipeline

  async TriggerBuildById(
    projectName: string,
    buildDefanitionId: string,
    parameter: any
  ) {
    let data = {
      definition: {
        id: buildDefanitionId,
      },
      parameters: parameter, //'{"Test":"123"}'
    };
    logger.info(JSON.stringify(data));
    let url = `${this.orgUrl}${projectName}/_apis/build/builds?api-version=5.0`;
    let res = await TFSServices.postRequest(
      url,
      this.token,
      "post",
      data,
      null
    );
    return res;
  }

  async GetArtifactByBuildId(
    projectName: string,
    buildId: string,
    artifactName: string
  ): Promise<any> {
    try {
      logger.info(
        `Get artifactory from project ${projectName},BuildId ${buildId} artifact name ${artifactName}`
      );
      logger.info(`Check if build ${buildId} have artifact`);
      let url = `${this.orgUrl}${projectName}/_apis/build/builds/${buildId}/artifacts`;
      let response = await TFSServices.getItemContent(
        url,
        this.token,
        "Get",
        null,
        null
      );
      if (response.count == 0) {
        logger.info(`No artifact for build ${buildId} was published `);
        return response;
      }
      url = `${this.orgUrl}${projectName}/_apis/build/builds/${buildId}/artifacts?artifactName=${artifactName}`;
      let res = await TFSServices.getItemContent(
        url,
        this.token,
        "Get",
        null,
        null
      );
      logger.info(`Url for download :${res.resource.downloadUrl}`);
      let result = await TFSServices.downloadZipFile(
        res.resource.downloadUrl,
        this.token
      );
      return result;
    } catch (err) {
      logger.error(`Error : ${err}`);
      throw new Error(String(err));
    }
  }

  async GetReleaseByReleaseId(
    projectName: string,
    releaseId: number
  ): Promise<any> {
    let url = `${this.orgUrl}${projectName}/_apis/release/releases/${releaseId}`;
    url = url.replace("dev.azure.com", "vsrm.dev.azure.com");
    return TFSServices.getItemContent(url, this.token, "get", null, null);
  }

  async GetPipelineRunHistory(projectName: string, pipelineId: string) {
    let url: string = `${this.orgUrl}${projectName}/_apis/pipelines/${pipelineId}/runs`;
    let res: any = await TFSServices.getItemContent(
      url,
      this.token,
      "get",
      null,
      null
    );
    return res;
  }

  async GetReleaseHistory(projectName: string, definitionId: string) {
    let url: string = `${this.orgUrl}${projectName}/_apis/release/releases?definitionId=${definitionId}&$top=2000`;
    url = url.replace("dev.azure.com", "vsrm.dev.azure.com");
    let res: any = await TFSServices.getItemContent(
      url,
      this.token,
      "get",
      null,
      null
    );
    return res;
  }

  async GetAllPipelines(projectName: string) {
    let url: string = `${this.orgUrl}${projectName}/_apis/pipelines?$top=2000`;
    let res: any = await TFSServices.getItemContent(
      url,
      this.token,
      "get",
      null,
      null
    );
    return res;
  }

  async GetAllReleaseDefenitions(projectName: string) {
    let url: string = `${this.orgUrl}${projectName}/_apis/release/definitions?$top=2000`;
    url = url.replace("dev.azure.com", "vsrm.dev.azure.com");
    let res: any = await TFSServices.getItemContent(
      url,
      this.token,
      "get",
      null,
      null
    );
    return res;
  }
}
