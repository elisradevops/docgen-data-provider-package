import { TFSServices } from '../helpers/tfs';
import TicketsDataProvider from './TicketsDataProvider';
import logger from '../utils/logger';
import { GitVersionDescriptor } from '../models/tfs-data';
export default class GitDataProvider {
  orgUrl: string = '';
  token: string = '';
  ticketsDataProvider: TicketsDataProvider;

  constructor(orgUrl: string, token: string) {
    this.orgUrl = orgUrl;
    this.token = token;
    this.ticketsDataProvider = new TicketsDataProvider(this.orgUrl, this.token);
  }
  async GetTeamProjectGitReposList(teamProject: string) {
    logger.debug(`fetching repos list for team project - ${teamProject}`);
    let url = `${this.orgUrl}/${teamProject}/_apis/git/repositories`;
    return TFSServices.getItemContent(url, this.token, 'get');
  } //GetGitRepoFromPrId

  async GetGitRepoFromRepoId(repoId: string) {
    logger.debug(`fetching repo data by id - ${repoId}`);
    let url = `${this.orgUrl}_apis/git/repositories/${repoId}`;
    return TFSServices.getItemContent(url, this.token, 'get');
  } //GetGitRepoFromPrId

  async GetJsonFileFromGitRepo(projectName: string, repoName: string, filePath: string) {
    let url = `${this.orgUrl}${projectName}/_apis/git/repositories/${repoName}/items?path=${filePath}&includeContent=true`;
    let res = await TFSServices.getItemContent(url, this.token, 'get');
    let jsonObject = JSON.parse(res.content);
    return jsonObject;
  } //GetJsonFileFromGitRepo

  async GetGitRepoFromPrId(pullRequestId: number) {
    let url = `${this.orgUrl}_apis/git/pullrequests/${pullRequestId}`;
    let res = await TFSServices.getItemContent(url, this.token, 'get');
    return res;
  } //GetGitRepoFromPrId

  async GetPullRequestCommits(repositoryId: string, pullRequestId: number) {
    let url = `${this.orgUrl}_apis/git/repositories/${repositoryId}/pullRequests/${pullRequestId}/commits`;
    let res = await TFSServices.getItemContent(url, this.token, 'get');
    return res;
  } //GetGitRepoFromPrId

  async GetPullRequestsLinkedItemsInCommitRange(
    projectId: string,
    repositoryId: string,
    commitRangeArray: any
  ) {
    let pullRequestsFilteredArray: any = [];
    let ChangeSetsArray: any = [];
    //get all pr's in git repo
    let url = `${this.orgUrl}${projectId}/_apis/git/repositories/${repositoryId}/pullrequests?status=completed&includeLinks=true&$top=2000}`;
    logger.debug(`request url: ${url}`);
    let pullRequestsArray = await TFSServices.getItemContent(url, this.token, 'get');
    logger.info(`got ${pullRequestsArray.count} pullrequests for repo: ${repositoryId}`);
    //iterate commit list to filter relavant pullrequests
    pullRequestsArray.value.forEach((pr: any) => {
      commitRangeArray.value.forEach((commit: any) => {
        if (pr.lastMergeCommit.commitId == commit.commitId) {
          pullRequestsFilteredArray.push(pr);
        }
      });
    });
    logger.info(
      `filtered in commit range ${pullRequestsFilteredArray.length} pullrequests for repo: ${repositoryId}`
    );
    //extract linked items and append them to result
    await Promise.all(
      pullRequestsFilteredArray.map(async (pr: any) => {
        let linkedItems: any = {};
        try {
          if (pr._links.workItems.href) {
            //get workitems linked to pr
            let url: string = pr._links.workItems.href;
            linkedItems = await TFSServices.getItemContent(url, this.token, 'get');
            logger.info(`got ${linkedItems.count} items linked to pr ${pr.pullRequestId}`);
            await Promise.all(
              linkedItems.value.map(async (item: any) => {
                let populatedItem = await this.ticketsDataProvider.GetWorkItem(projectId, item.id);
                let changeSet: any = {
                  workItem: populatedItem,
                  pullrequest: pr,
                };
                ChangeSetsArray.push(changeSet);
              })
            );
          }
        } catch (error) {
          logger.error(error);
        }
      })
    );
    return ChangeSetsArray;
  } //GetPullRequestsInCommitRange

  async GetItemsInCommitRange(projectId: string, repositoryId: string, commitRange: any) {
    //get all items linked to commits
    let res: any = [];
    let commitChangesArray: any = [];
    //extract linked items and append them to result
    for (const commit of commitRange.value) {
      if (commit.workItems) {
        for (const wi of commit.workItems) {
          let populatedItem = await this.ticketsDataProvider.GetWorkItem(projectId, wi.id);
          let changeSet: any = { workItem: populatedItem, commit: commit };
          commitChangesArray.push(changeSet);
        }
      }
    }
    //get all items and pr data from pr's in commit range - using the above function
    let pullRequestsChangesArray = await this.GetPullRequestsLinkedItemsInCommitRange(
      projectId,
      repositoryId,
      commitRange
    );
    //merge commit links with pr links
    logger.info(`got ${pullRequestsChangesArray.length} items from pr's and`);
    res = [...commitChangesArray, ...pullRequestsChangesArray];
    let workItemIds: any = [];
    for (let index = 0; index < res.length; index++) {
      if (workItemIds.includes(res[index].workItem.id)) {
        res.splice(index, 1);
        index--;
      } else {
        workItemIds.push(res[index].workItem.id);
      }
    }
    return res;
  } //GetItemsInCommitRange

  async GetPullRequestsInCommitRangeWithoutLinkedItems(
    projectId: string,
    repositoryId: string,
    commitRangeArray: any
  ) {
    let pullRequestsFilteredArray: any[] = [];

    // Extract the organization name from orgUrl
    let orgName = this.orgUrl.split('/').filter(Boolean).pop();

    // Get all PRs in the git repo
    let url = `${this.orgUrl}${projectId}/_apis/git/repositories/${repositoryId}/pullrequests?status=completed&includeLinks=true&$top=2000}`;
    logger.debug(`request url: ${url}`);
    let pullRequestsArray = await TFSServices.getItemContent(url, this.token, 'get');
    logger.info(`got ${pullRequestsArray.count} pullrequests for repo: ${repositoryId}`);

    // Iterate commit list to filter relevant pull requests
    pullRequestsArray.value.forEach((pr: any) => {
      commitRangeArray.value.forEach((commit: any) => {
        if (pr.lastMergeCommit.commitId == commit.commitId) {
          // Construct the pull request URL
          const prUrl = `https://dev.azure.com/${orgName}/${projectId}/_git/${repositoryId}/pullrequest/${pr.pullRequestId}`;

          // Extract only the desired properties from the PR
          const prFilteredData = {
            pullRequestId: pr.pullRequestId,
            createdBy: pr.createdBy.displayName,
            creationDate: pr.creationDate,
            closedDate: pr.closedDate,
            title: pr.title,
            description: pr.description,
            url: prUrl, // Use the constructed URL here
          };
          pullRequestsFilteredArray.push(prFilteredData);
        }
      });
    });
    logger.info(
      `filtered in commit range ${pullRequestsFilteredArray.length} pullrequests for repo: ${repositoryId}`
    );

    return pullRequestsFilteredArray;
  } // GetPullRequestsInCommitRangeWithoutLinkedItems

  async GetCommitByCommitId(projectId: string, repositoryId: string, commitSha: string) {
    let url = `${this.orgUrl}${projectId}/_apis/git/repositories/${repositoryId}/commits/${commitSha}`;
    return TFSServices.getItemContent(url, this.token, 'get');
  }

  async GetCommitForPipeline(projectId: string, buildId: number) {
    let url = `${this.orgUrl}${projectId}/_apis/build/builds/${buildId}`;
    let res = await TFSServices.getItemContent(url, this.token, 'get');
    return res.sourceVersion;
  } //GetCommitForPipeline

  //TODO: replace this....
  async GetItemsForPipelinesRange(projectId: string, fromBuildId: number, toBuildId: number) {
    let linkedItemsArray: any = [];
    let url = `${this.orgUrl}${projectId}/_apis/build/workitems?fromBuildId=${fromBuildId}&toBuildId=${toBuildId}&$top=2000`;
    let res = await TFSServices.getItemContent(url, this.token, 'get');
    logger.info(`recieved ${res.count} items in build range ${fromBuildId}-${toBuildId}`);
    await Promise.all(
      res.value.map(async (wi: any) => {
        let populatedItem = await this.ticketsDataProvider.GetWorkItem(projectId, wi.id);
        let changeSet: any = { workItem: populatedItem, build: toBuildId };
        linkedItemsArray.push(changeSet);
      })
    );
    return linkedItemsArray;
  } //GetCommitForPipeline

  async getItemsForPipelineRange(projectName: string, extendedCommits: any[], targetRepo: any) {
    let commitChangesArray: any[] = [];
    try {
      if (extendedCommits?.length === 0) {
        throw new Error('extended commits cannot be empty');
      }
      //First fetch the repo web url
      if (targetRepo.url) {
        const repoData = await TFSServices.getItemContent(targetRepo.url, this.token);
        const repoWebUrl = repoData._links?.web.href;
        if (repoWebUrl) {
          targetRepo['url'] = repoWebUrl;
        }
      }
      //Then extend the commit information with the related WIs
      for (const extendedCommit of extendedCommits) {
        const { commit } = extendedCommit;
        if (!commit.workItems) {
          throw new Error(`commit ${commit.commitId} does not have work items`);
        }
        for (const wi of commit.workItems) {
          const populatedWorkItem = await this.ticketsDataProvider.GetWorkItem(projectName, wi.id);
          let changeSet: any = { workItem: populatedWorkItem, commit: commit, targetRepo };
          commitChangesArray.push(changeSet);
        }
      }

      let workItemIds: any = [];
      for (let index = 0; index < commitChangesArray.length; index++) {
        if (workItemIds.includes(commitChangesArray[index].workItem.id)) {
          commitChangesArray.splice(index, 1);
          index--;
        } else {
          workItemIds.push(commitChangesArray[index].workItem.id);
        }
      }
    } catch (error: any) {
      logger.error(error.message);
    }

    return commitChangesArray;
  }

  async GetCommitsInDateRange(
    projectId: string,
    repositoryId: string,
    fromDate: string,
    toDate: string,
    branchName?: string
  ) {
    let url: string;
    if (typeof branchName !== 'undefined') {
      url = `${this.orgUrl}${projectId}/_apis/git/repositories/${repositoryId}/commits?searchCriteria.fromDate=${fromDate}&searchCriteria.toDate=${toDate}&searchCriteria.includeWorkItems=true&searchCriteria.$top=2000&searchCriteria.itemVersion.version=${branchName}`;
    } else {
      url = `${this.orgUrl}${projectId}/_apis/git/repositories/${repositoryId}/commits?searchCriteria.fromDate=${fromDate}&searchCriteria.toDate=${toDate}&searchCriteria.includeWorkItems=true&searchCriteria.$top=2000`;
    }
    return TFSServices.getItemContent(url, this.token, 'get');
  } //GetCommitsInDateRange

  async GetCommitsInCommitRange(projectId: string, repositoryId: string, fromSha: string, toSha: string) {
    let url = `${this.orgUrl}${projectId}/_apis/git/repositories/${repositoryId}/commits?searchCriteria.fromCommitId=${fromSha}&searchCriteria.toCommitId=${toSha}&searchCriteria.includeWorkItems=true&searchCriteria.$top=2000`;
    return TFSServices.getItemContent(url, this.token, 'get');
  } //GetCommitsInCommitRange doesen't work!!!

  async CreatePullRequestComment(projectName: string, repoID: string, pullRequestID: number, threads: any) {
    let url: string = `${this.orgUrl}${projectName}/_apis/git/repositories/${repoID}/pullRequests/${pullRequestID}/threads?api-version=5.0`;
    let res: any = await TFSServices.getItemContent(url, this.token, 'post', threads, null);
    return res;
  }

  async GetPullRequestComments(projectName: string, repoID: string, pullRequestID: number) {
    let url: string = `${this.orgUrl}${projectName}/_apis/git/repositories/${repoID}/pullRequests/${pullRequestID}/threads`;
    return TFSServices.getItemContent(url, this.token, 'get', null, null);
  }

  async GetCommitsForRepo(projectName: string, repoID: string, branchName?: string) {
    let url: string;
    if (typeof branchName !== 'undefined') {
      url = `${this.orgUrl}${projectName}/_apis/git/repositories/${repoID}/commits?searchCriteria.$top=2000&searchCriteria.itemVersion.version=${branchName}`;
    } else {
      url = `${this.orgUrl}${projectName}/_apis/git/repositories/${repoID}/commits?searchCriteria.$top=2000`;
    }
    let res: any = await TFSServices.getItemContent(url, this.token, 'get', null, null);
    return res;
  }

  async GetPullRequestsForRepo(projectName: string, repoID: string) {
    let url: string = `${this.orgUrl}${projectName}/_apis/git/repositories/${repoID}/pullrequests?status=completed&includeLinks=true&$top=2000}`;

    let res: any = await TFSServices.getItemContent(url, this.token, 'get', null, null);
    return res;
  }

  async GetItemsInPullRequestRange(projectId: string, repositoryId: string, pullRequestIDs: any) {
    let pullRequestsFilteredArray: any = [];
    let ChangeSetsArray: any = [];
    //get all pr's in git repo
    let url = `${this.orgUrl}${projectId}/_apis/git/repositories/${repositoryId}/pullrequests?status=completed&includeLinks=true&$top=2000}`;
    logger.debug(`request url: ${url}`);
    let pullRequestsArray = await TFSServices.getItemContent(url, this.token, 'get');
    logger.info(`got ${pullRequestsArray.count} pullrequests for repo: ${repositoryId}`);
    //iterate commit list to filter relavant pullrequests
    pullRequestsArray.value.forEach((pr: any) => {
      pullRequestIDs.forEach((prId: any) => {
        if (prId == pr.pullRequestId) {
          pullRequestsFilteredArray.push(pr);
        }
      });
    });
    logger.info(
      `filtered in prId range ${pullRequestsFilteredArray.length} pullrequests for repo: ${repositoryId}`
    );
    //extract linked items and append them to result
    await Promise.all(
      pullRequestsFilteredArray.map(async (pr: any) => {
        let linkedItems: any = {};
        try {
          if (pr._links.workItems.href) {
            //get workitems linked to pr
            let url: string = pr._links.workItems.href;
            linkedItems = await TFSServices.getItemContent(url, this.token, 'get');
            logger.info(`got ${linkedItems.count} items linked to pr ${pr.pullRequestId}`);
            await Promise.all(
              linkedItems.value.map(async (item: any) => {
                let populatedItem = await this.ticketsDataProvider.GetWorkItem(projectId, item.id);
                let changeSet: any = {
                  workItem: populatedItem,
                  pullrequest: pr,
                };
                ChangeSetsArray.push(changeSet);
              })
            );
          }
        } catch (error) {
          logger.error(error);
        }
      })
    );
    return ChangeSetsArray;
  }

  async GetRepoBranches(projectName: string, repoID: string) {
    let url: string = `${this.orgUrl}${projectName}/_apis/git/repositories/${repoID}/refs?searchCriteria.$top=1000&filter=heads`;
    let res: any = await TFSServices.getItemContent(url, this.token, 'get', null, null);
    return res;
  }

  async GetCommitBatch(
    projectName: string,
    repoID: string,
    itemVersion: GitVersionDescriptor,
    compareVersion: GitVersionDescriptor
  ) {
    const allCommitsExtended: any[] = [];
    let skipping = 0;
    let chunkSize = 100;
    let commitCounter = 0;
    let logToConsoleAfterCommits = 500;
    try {
      let body = {
        itemVersion,
        compareVersion,
        includeWorkItems: true,
      };

      let url = `${this.orgUrl}${projectName}/_apis/git/repositories/${repoID}/commitsBatch?api-version=5.0&$skip=${skipping}&$top=${chunkSize}`;
      let commitsResponse = await TFSServices.postRequest(url, this.token, undefined, body);
      let commits = commitsResponse.data;
      while (commits.count > 0) {
        for (const commit of commits.value) {
          try {
            if (commitCounter % logToConsoleAfterCommits === 0) {
              logger.debug(`commit number ${commitCounter + 1}`);
            }

            if (commit.workItems?.length > 0) {
              let extendedCommit: any = {};

              extendedCommit['commit'] = commit;

              let committerName = commit.committer.name;
              let commitDate = commit.committer.date.toString().slice(0, 10);
              extendedCommit['committerName'] = committerName;
              extendedCommit['commitDate'] = commitDate;

              allCommitsExtended.push(extendedCommit);
            }
          } catch (err: any) {
            const errMsg = `Cannot fetch commit batch: ${err.message}`;
            throw new Error(errMsg);
          } finally {
            commitCounter++;
          }
        }

        skipping += chunkSize;
        let url = `${this.orgUrl}${projectName}/_apis/git/repositories/${repoID}/commitsBatch?api-version=5.0&$skip=${skipping}&$top=${chunkSize}`;
        commitsResponse = await TFSServices.postRequest(url, this.token, undefined, body);
        commits = commitsResponse.data;
      }
    } catch (error: any) {
      logger.error(error.message);
    }
    return allCommitsExtended;
  }
}
