import axios from 'axios';
import logger from '../utils/logger';

export class TFSServices {
  public static async downloadZipFile(url: string, pat: string): Promise<any> {
    try {
      let res = await axios.request({
        url: url,
        headers: { 'Content-Type': 'application/zip' },
        auth: {
          username: '',
          password: pat,
        },
      });
      return res;
    } catch (e) {
      logger.error(`error download zip file , url : ${url}`);
      throw new Error(String(e));
    }
  }

  public static async fetchAzureDevOpsImageAsBase64(
    url: string,
    pat: string,
    requestMethod: string = 'get',
    data: any = {},
    customHeaders: any = {},
    printError: boolean = true
  ): Promise<any> {
    let config: any = {
      headers: customHeaders,
      method: requestMethod,
      auth: {
        username: '',
        password: pat,
      },
      data: data,
      responseType: 'arraybuffer', // Important for binary data
      timeout: 5000, // Set timeout to 5 seconds
    };
    let json;
    let attempts = 0;
    const maxAttempts = 3;

    logger.silly(`making request:
    url: ${url}
    config: ${JSON.stringify(config)}`);

    while (attempts < maxAttempts) {
      try {
        const response = await axios(url, config);

        // Convert binary data to Base64
        const base64String = Buffer.from(response.data, 'binary').toString('base64');
        const contentType = response.headers['content-type']; // e.g., "image/png; api-version=7.1"
        const mimeType = contentType.split(';')[0].trim(); // Extracts "image/png"
        return `data:${mimeType};base64,${base64String}`;
      } catch (e: any) {
        attempts++;
        if (e.message.includes('ETIMEDOUT') && attempts < maxAttempts) {
          logger.warn(`Request timed out. Retrying attempt ${attempts} of ${maxAttempts}...`);
          continue;
        }
        if (printError) {
          if (e.response) {
            logger.error(`Error fetching image from Azure DevOps at ${url}: ${e.message}`);
            logger.error(`Status: ${e.response.status}`);
            logger.error(`Response Data: ${JSON.stringify(e.response.data)}`);
          } else {
            logger.error(`Error fetching image from Azure DevOps at ${url}: ${e.message}`);
          }
        }
        throw e;
      }
    }
  }

  public static async getItemContent(
    url: string,
    pat: string,
    requestMethod: string = 'get',
    data: any = {},
    customHeaders: any = {},
    printError: boolean = true
  ): Promise<any> {
    let config: any = {
      headers: customHeaders,
      method: requestMethod,
      auth: {
        username: '',
        password: pat,
      },
      data: data,
      timeout: 2000, // Set timeout to 2 seconds
    };
    let json;
    let attempts = 0;
    const maxAttempts = 3;

    logger.silly(`making request:
    url: ${url}
    config: ${JSON.stringify(config)}`);

    while (attempts < maxAttempts) {
      try {
        let result = await axios(url, config);
        json = JSON.parse(JSON.stringify(result.data));
        return json;
      } catch (e: any) {
        attempts++;
        if (e.message.includes('ETIMEDOUT') || (e.message.includes('timeout') && attempts < maxAttempts)) {
          logger.warn(`Request timed out. Retrying attempt ${attempts} of ${maxAttempts}...`);
          continue;
        }
        if (printError) {
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
        throw e;
      }
    }
  }

  public static async getJfrogRequest(url: string, header?: any) {
    let config: any = {
      method: 'GET',
    };
    if (header) {
      config['headers'] = header;
    }

    let json;
    try {
      let result = await axios(url, config);
      json = JSON.parse(JSON.stringify(result.data));
    } catch (e: any) {
      if (e.response) {
        // Log detailed error information including the URL
        logger.error(`Error making request Jfrog at ${url}: ${e.message}`);
        logger.error(`Status: ${e.response.status}`);
        logger.error(`Response Data: ${JSON.stringify(e.response.data)}`);
      } else {
        // Handle other errors (network, etc.)
        logger.error(`Error making request to Jfrog at ${url}: ${e.message}`);
      }
      throw e;
    }
    return json;
  }

  public static async postRequest(
    url: string,
    pat: string,
    requestMethod: string = 'post',
    data: any,
    customHeaders: any = { headers: { 'Content-Type': 'application/json' } }
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
