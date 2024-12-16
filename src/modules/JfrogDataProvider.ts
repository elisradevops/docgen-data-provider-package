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
    let url = `${this.orgUrl}${teamProject}/_apis/serviceendpoint/endpoints/${connectionId}?api-version=6`;
    const serviceConnectionResponse = await TFSServices.getItemContent(url, this.tfsToken);
    logger.debug(`service connection url ${JSON.stringify(serviceConnectionResponse.url)}`);
    return serviceConnectionResponse.url;
  }

  async getCiDataFromJfrog(jfrogUrl: string, buildName: string, buildVersion: string) {
    let jfrogHeader: any = {};
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

    try {
      const getCiResponse =
        this.jfrogToken !== ''
          ? await TFSServices.getJfrogRequest(getCiRequestUrl, jfrogHeader)
          : await TFSServices.getJfrogRequest(getCiRequestUrl);
      logger.debug(`CI Url from JFROG: ${getCiResponse.buildInfo.url}`);
    } catch (err: any) {
      logger.error('Error occurred during querying JFrog using');
    }
    return '';
  }
}
