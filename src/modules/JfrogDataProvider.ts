import { TFSServices } from '../helpers/tfs';

import logger from '../utils/logger';

export default class JfrogDataProvider {
  orgUrl: string = '';
  tfsToken: string = '';
  jfrogToken: string = '';

  constructor(orgUrl: string, tfsToken: string, jfrogToken: string) {
    this.orgUrl = orgUrl;
    this.tfsToken = tfsToken;
    this.jfrogToken = jfrogToken;
  }

  async getServiceConnectionUrlByConnectionId(teamProject: string, connectionId: string) {
    let url = `${this.orgUrl}${teamProject}/_apis/serviceendpoint/endpoints/${connectionId}?api-version=7.1`;
    const serviceConnectionResponse = await TFSServices.getItemContent(url, this.tfsToken);
    if (!serviceConnectionResponse?.url) {
      throw new Error(`Service connection ${connectionId} returned no URL — identity may lack service-endpoint read permission`);
    }
    logger.debug(`service connection url ${JSON.stringify(serviceConnectionResponse.url)}`);
    return serviceConnectionResponse.url;
  }

  async getCiDataFromJfrog(jfrogUrl: string, buildName: string, buildVersion: string): Promise<string> {
    let jfrogHeader: any = {};
    try {
      if (!jfrogUrl) {
        throw new Error(
          `JFrog service connection URL is unresolved for build "${buildName}" — the caller identity may lack service-endpoint read permission`
        );
      }
      if (this.jfrogToken !== '') {
        jfrogHeader['Authorization'] = `Bearer ${this.jfrogToken}`;
      }
      if (!buildName.startsWith('/')) {
        const currentBuildName = buildName;
        buildName = '/' + currentBuildName;
      }

      if (!buildVersion.startsWith('/')) {
        const currentBuildVersion = buildVersion;
        buildVersion = '/' + currentBuildVersion;
      }
      let getCiRequestUrl = `${jfrogUrl}/api/build${buildName}${buildVersion}`;
      logger.info(`Querying Jfrog using url ${getCiRequestUrl}`);

      const getCiResponse =
        this.jfrogToken !== ''
          ? await TFSServices.getJfrogRequest(getCiRequestUrl, jfrogHeader)
          : await TFSServices.getJfrogRequest(getCiRequestUrl);
      if (!getCiResponse?.buildInfo?.url) {
        throw new Error(`JFrog response for build "${buildName}${buildVersion}" is missing buildInfo.url`);
      }
      logger.debug(`CI Url from JFROG: ${getCiResponse.buildInfo.url}`);
      return getCiResponse.buildInfo.url;
    } catch (err: any) {
      logger.error(`Error occurred during querying JFrog using: ${err.message}`);
      throw err;
    }
  }
}
