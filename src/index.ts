import MangementDataProvider from './modules/MangementDataProvider';
import TicketsDataProvider from './modules/TicketsDataProvider';
import GitDataProvider from './modules/GitDataProvider';
import PipelinesDataProvider from './modules/PipelinesDataProvider';
import TestDataProvider from './modules/TestDataProvider';

import logger from './utils/logger';
import ResultDataProvider from './modules/ResultDataProvider';
import JfrogDataProvider from './modules/JfrogDataProvider';

export default class DgDataProviderAzureDevOps {
  orgUrl: string = '';
  token: string = '';
  apiVersion: string;
  jfrogToken?: string;
  constructor(orgUrl: string, token: string, apiVersion?: string, jfrogToken?: string) {
    this.orgUrl = orgUrl;
    this.token = token;
    this.jfrogToken = jfrogToken;
    logger.info(`azure devops data provider initilized`);
  }

  async getMangementDataProvider(): Promise<MangementDataProvider> {
    return new MangementDataProvider(this.orgUrl, this.token);
  }
  async getTicketsDataProvider(): Promise<TicketsDataProvider> {
    return new TicketsDataProvider(this.orgUrl, this.token);
  }
  async getGitDataProvider(): Promise<GitDataProvider> {
    return new GitDataProvider(this.orgUrl, this.token);
  }
  async getPipelinesDataProvider(): Promise<PipelinesDataProvider> {
    return new PipelinesDataProvider(this.orgUrl, this.token);
  }
  async getTestDataProvider(): Promise<TestDataProvider> {
    return new TestDataProvider(this.orgUrl, this.token);
  }

  async getResultDataProvider(): Promise<ResultDataProvider> {
    return new ResultDataProvider(this.orgUrl, this.token);
  }

  async getJfrogDataProvider(): Promise<JfrogDataProvider> {
    return new JfrogDataProvider(this.orgUrl, this.token, this.jfrogToken || '');
  }
} //class
