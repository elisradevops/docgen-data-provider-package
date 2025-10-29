import axios, { AxiosInstance, AxiosRequestConfig } from 'axios';
import logger from '../utils/logger';

// Environment detection
const isNode = typeof window === 'undefined' && typeof process !== 'undefined' && process.versions?.node;

export class TFSServices {
  // Universal axios instance with environment-specific configuration
  private static axiosInstance: AxiosInstance = axios.create(
    isNode
      ? {
          // Node.js configuration with connection pooling and HTTPS handling
          httpAgent: new (require('http').Agent)({
            keepAlive: true,
            maxSockets: 50,
            keepAliveMsecs: 300000,
          }),
          httpsAgent: new (require('https').Agent)({
            keepAlive: true,
            maxSockets: 50,
            keepAliveMsecs: 300000,
            // Disable SSL certificate validation
            rejectUnauthorized: false,
          }),
          timeout: 30000,
        }
      : {
          // Browser configuration - browsers handle HTTPS automatically
          timeout: 30000,
          maxRedirects: 5,
        }
  );

  public static async downloadZipFile(url: string, pat: string): Promise<any> {
    try {
      const res = await this.axiosInstance.request({
        url: url,
        headers: { 'Content-Type': 'application/zip' },
        auth: { username: '', password: pat },
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
    const config: AxiosRequestConfig = {
      headers: customHeaders,
      method: requestMethod,
      auth: { username: '', password: pat },
      data: data,
      responseType: 'arraybuffer', // Important for binary data
    };

    return this.executeWithRetry(url, config, printError, (response) => {
      // Convert binary data to Base64
      const base64String = Buffer.from(response.data, 'binary').toString('base64');
      const contentType = response.headers['content-type'] || 'application/octet-stream';
      const mimeType = contentType.split(';')[0].trim();
      if (!mimeType.startsWith('image/')) {
        throw new Error(`Expected image content but received '${mimeType}'`);
      }
      return `data:${mimeType};base64,${base64String}`;
    });
  }

  public static async getItemContent(
    url: string,
    pat: string,
    requestMethod: string = 'get',
    data: any = {},
    customHeaders: any = {},
    printError: boolean = true
  ): Promise<any> {
    // Clean URL
    const cleanUrl = url.replace(/ /g, '%20');

    const config: AxiosRequestConfig = {
      headers: customHeaders,
      method: requestMethod,
      auth: { username: '', password: pat },
      data: data,
      timeout: requestMethod.toLocaleLowerCase() === 'get' ? 10000 : undefined, // More reasonable timeout
    };

    return this.executeWithRetry(cleanUrl, config, printError, (response) => {
      // Direct return of data without extra JSON parsing
      return response.data;
    });
  }

  public static async getItemContentWithHeaders(
    url: string,
    pat: string,
    requestMethod: string = 'get',
    data: any = {},
    customHeaders: any = {},
    printError: boolean = true
  ): Promise<{ data: any; headers: any }> {
    // Clean URL
    const cleanUrl = url.replace(/ /g, '%20');

    const config: AxiosRequestConfig = {
      headers: customHeaders,
      method: requestMethod,
      auth: { username: '', password: pat },
      data: data,
      timeout: requestMethod.toLocaleLowerCase() === 'get' ? 10000 : undefined,
    };

    return this.executeWithRetry(cleanUrl, config, printError, (response) => {
      return { data: response.data, headers: response.headers };
    });
  }

  public static async getJfrogRequest(url: string, header?: any) {
    const config: AxiosRequestConfig = {
      url: url,
      method: 'GET',
      headers: header,
    };

    try {
      const result = await this.axiosInstance.request(config);
      return result.data;
    } catch (e: any) {
      this.logDetailedError(e, url);
      throw e;
    }
  }

  public static async postRequest(
    url: string,
    pat: string,
    requestMethod: string = 'post',
    data: any,
    customHeaders: any = { headers: { 'Content-Type': 'application/json' } }
  ): Promise<any> {
    const config: AxiosRequestConfig = {
      url: url,
      headers: customHeaders,
      method: requestMethod,
      auth: { username: '', password: pat },
      data: data,
    };

    try {
      const result = await this.axiosInstance.request(config);
      return result;
    } catch (e: any) {
      this.logDetailedError(e, url);
      throw e;
    }
  }

  /**
   * Execute a request with intelligent retry logic
   */
  private static async executeWithRetry(
    url: string,
    config: AxiosRequestConfig,
    printError: boolean,
    responseProcessor: (response: any) => any
  ): Promise<any> {
    let attempts = 0;
    const maxAttempts = 3;
    const baseDelay = 500; // Start with 500ms delay

    while (true) {
      try {
        const result = await this.axiosInstance.request({ ...config, url });
        return responseProcessor(result);
      } catch (e: any) {
        attempts++;
        const errorMessage = this.getErrorMessage(e);

        // Handle not found errors
        if (errorMessage.includes('could not be found')) {
          logger.info(`File does not exist, or you do not have permissions to read it.`);
          throw new Error(`File not found or insufficient permissions: ${url}`);
        }

        // Check if we should retry
        if (attempts < maxAttempts && this.isRetryableError(e)) {
          // Calculate exponential backoff with jitter
          const jitter = Math.random() * 0.3 + 0.85; // Between 0.85 and 1.15
          const delay = Math.min(baseDelay * Math.pow(2, attempts - 1) * jitter, 5000);

          logger.warn(`Request failed. Retrying in ${Math.round(delay)}ms (${attempts}/${maxAttempts})`);
          await new Promise((resolve) => setTimeout(resolve, delay));
          continue;
        }

        // Log error if needed
        if (printError) {
          this.logDetailedError(e, url);
        }

        throw e;
      }
    }
  }

  /**
   * Check if an error is retryable
   */
  private static isRetryableError(error: any): boolean {
    // Network errors
    if (error.code === 'ECONNRESET' || error.code === 'ETIMEDOUT' || error.message.includes('timeout')) {
      return true;
    }

    // Server errors (5xx)
    if (error.response?.status >= 500) {
      return true;
    }

    // Rate limiting (429)
    if (error.response?.status === 429) {
      return true;
    }

    return false;
  }

  /**
   * Log detailed error information
   */
  private static logDetailedError(error: any, url: string): void {
    if (error.response) {
      logger.error(`Error for ${url}: ${error.message}`);
      logger.error(`Status: ${error.response.status}`);

      if (error.response.data) {
        if (typeof error.response.data === 'string') {
          logger.error(`Response: ${error.response.data.substring(0, 200)}`);
        } else {
          const dataMessage =
            error.response.data.message || JSON.stringify(error.response.data).substring(0, 200);
          logger.error(`Response: ${dataMessage}`);
        }
      }
    } else {
      logger.error(`Error for ${url}: ${error.message}`);
    }
  }

  private static getErrorMessage(error: any): string {
    if (error.response?.data?.message) {
      return JSON.stringify(error.response.data.message);
    } else if (error.response?.data) {
      return JSON.stringify(error.response.data);
    } else if (error.response) {
      return `HTTP ${error.response.status}`;
    } else if (error.message) {
      return error.message;
    } else {
      return 'Unknown error occurred';
    }
  }
}
