import axios from 'axios';
import logger from '../utils/logger';
import * as request from 'request';

export class TFSServices {
  public static async downloadZipFile(url: string, pat: string): Promise<any> {
    let json;
    try {
      var d = {
        request: {
          method: 'getUser',
          auth: {
            username: '',
            password: `${pat}`,
          },
        },
      };
      let res = await request({
        url: url,
        headers: { 'Content-Type': 'application/zip' },
        auth: {
          username: '',
          password: pat,
        },
        encoding: null,
      });
      return res;
    } catch (e) {
      logger.error(`error download zip file , url : ${url}`);
      throw new Error(String(e));
    }
  }

  public static async getItemContent(
    url: string,
    pat: string,
    requestMethod: string = 'get',
    data: any = {},
    customHeaders: any = {}
  ): Promise<any> {
    let config: any = {
      headers: customHeaders,
      method: requestMethod,
      auth: {
        username: '',
        password: pat,
      },
      data: data,
    };
    let json;
    logger.silly(`making request:
    url: ${url}
    config: ${JSON.stringify(config)}`);
    try {
      let result = await axios(url, config);
      json = JSON.parse(JSON.stringify(result.data));
    } catch (e: any) {
      if (e.response) {
        // Log detailed error information including the URL
        logger.error(`Error making request to Azure DevOps at ${url}: ${e.message}`);
        logger.error(`Status: ${e.response.status}`);
        logger.error(`Response Data: ${JSON.stringify(e.response.data)}`);
      } else {
        // Handle other errors (network, etc.)
        logger.error(`Error making request to Azure DevOps at ${url}: ${e.message}`);
      }
    }
    return json;
  }

  public static async postRequest(
    url: string,
    pat: string,
    requestMethod: string = 'post',
    data: any,
    customHeaders: any = {}
  ): Promise<any> {
    let config: any = {
      headers: customHeaders,
      method: requestMethod,
      auth: {
        username: '',
        password: pat,
      },
      data: data,
    };
    let result;
    logger.silly(`making request:
    url: ${url}
    config: ${JSON.stringify(config)}`);
    try {
      result = await axios(url, config);
    } catch (e) {
      logger.error(`error making post request to azure devops`);
      console.log(e);
    }
    return result;
  }
}
