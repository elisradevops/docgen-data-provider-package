import { TFSServices } from '../helpers/tfs';
import logger from '../utils/logger';

export default class MangementDataProvider {
  orgUrl: string = '';
  token: string = '';

  constructor(orgUrl: string, token: string) {
    this.orgUrl = orgUrl;
    this.token = token;
  }

  async GetCllectionLinkTypes() {
    let url: string = `${this.orgUrl}_apis/wit/workitemrelationtypes`;
    let res: any = await TFSServices.getItemContent(url, this.token, 'get', null, null);
    return res;
  }

  //get all projects
  async GetProjects(): Promise<any> {
    let projectUrl: string = `${this.orgUrl}_apis/projects?$top=1000`;
    let projects: any = await TFSServices.getItemContent(projectUrl, this.token);
    return projects;
  }

  // get project by  name return project object
  async GetProjectByName(projectName: string): Promise<any> {
    try {
      let projects: any = await this.GetProjects();
      for (let i = 0; i < projects.value.length; i++) {
        if (projects.value[i].name === projectName) return projects.value[i];
      }
      return {};
    } catch (err) {
      console.log(err);
      return {};
    }
  }

  // get project by id return project object
  async GetProjectByID(projectID: string): Promise<any> {
    let projectUrl: string = `${this.orgUrl}_apis/projects/${projectID}`;
    let project: any = await TFSServices.getItemContent(projectUrl, this.token);
    return project;
  }

  async GetUserProfile(): Promise<any> {
    let url: string = `${this.orgUrl}_api/_common/GetUserProfile?__v=5`;
    return TFSServices.getItemContent(url, this.token);
  }

  // Check if organization URL is valid and optionally validate PAT
  // Without token: checks if organization URL exists
  // With token: checks both URL validity AND PAT validity
  async CheckOrgUrlValidity(token?: string): Promise<any> {
    let url: string = `${this.orgUrl}_apis/connectionData`;
    // Use provided token or empty string for URL-only validation
    return TFSServices.getItemContent(url, token || '', 'get', null, null, false);
  }
}
