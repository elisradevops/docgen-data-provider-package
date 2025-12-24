import { TFSServices } from '../helpers/tfs';
import TicketsDataProvider from './TicketsDataProvider';
import logger from '../utils/logger';
import {
  createLinkedRelation,
  createRequirementRelation,
  GitVersionDescriptor,
  LinkedRelation,
  value,
} from '../models/tfs-data';
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
    const res = await TFSServices.getItemContent(url, this.token, 'get');
    return res.value && res.value.length > 0
      ? [...res.value].sort((a, b) => a.name.toLowerCase().localeCompare(b.name.toLowerCase()))
      : [];
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

  async GetTag(gitRepoUrl: string, tag: string) {
    const encodedTag = encodeURIComponent(tag);
    let url = `${gitRepoUrl}/refs/tags/${encodedTag}?peelTags=true&api-version=5.1`;
    const res = await TFSServices.getItemContent(url, this.token, 'get');
    if (res?.value && Array.isArray(res.value) && res.value.length > 0) {
      const match = res.value.find((r: any) => r.name.split('/').pop().toLowerCase() === tag.toLowerCase());
      if (match) {
        return {
          name: match.name.replace('refs/tags/', ''),
          objectId: match.objectId,
          url: url,
          peeledObjectId: match.peeledObjectId,
        };
      }
    }
    return null;
  } //GetTag

  async GetBranch(gitRepoUrl: string, branch: string) {
    const encodedBranch = encodeURIComponent(branch);
    let url = `${gitRepoUrl}/refs?filter=heads/${encodedBranch}&api-version=5.1`;
    const res = await TFSServices.getItemContent(url, this.token, 'get');
    if (res && res.value && Array.isArray(res.value) && res.value.length > 0) {
      const match = res.value.find(
        (r: any) => r.name.split('/').pop().toLowerCase() === branch.toLowerCase()
      );

      if (match) {
        return match;
      }
    }
    return null;
  } //GetBranch

  /**
   * Gets a file from a Git repository.
   *
   * @param projectName - The name of the project.
   * @param repoId - The ID of the repository.
   * @param fileName - The name of the file to retrieve.
   * @param version - The version descriptor for the file.
   * @param gitRepoUrl - Optional URL of the Git repository.
   * @returns The file content as a string.
   */
  async GetFileFromGitRepo(
    projectName: string,
    repoId: string,
    fileName: string,
    version: GitVersionDescriptor,
    gitRepoUrl: string = ''
  ) {
    // get a single tag
    let versionFix = '';
    try {
      versionFix = version.version.replace('/', '%2F').replace('#', '%23');
    } catch {
      versionFix = version.version;
    }
    try {
      let urlPrefix =
        gitRepoUrl !== '' ? gitRepoUrl : `${this.orgUrl}${projectName}/_apis/git/repositories/${repoId}`;

      let url =
        `${urlPrefix}/items` +
        `?path=${fileName}&download=true&includeContent=true&recursionLevel=none` +
        `&versionDescriptor.version=${versionFix}` +
        `&versionDescriptor.versionType=${version.versionType}` +
        `&api-version=5.1`;

      let res = await TFSServices.getItemContent(url, this.token, 'get', {}, {}, false);
      if (res && res.content) {
        const fileContent = res.content;
        // Assuming the file content is in plain text format
        return fileContent;
      }
      return undefined;
    } catch (err: any) {
      logger.warn(`File ${fileName} could not be read: ${err.message}`);
      return undefined;
    }
  }

  async CheckIfItemExist(
    gitApiUrl: string,
    itemPath: string,
    version: GitVersionDescriptor
  ): Promise<boolean> {
    let safePath = encodeURIComponent(itemPath).replace(/%2F/g, '/');
    let versionFixed = '';
    try {
      versionFixed = encodeURIComponent(version.version).replace(/%2F/g, '%2F').replace(/#/g, '%23');
    } catch {
      versionFixed = version.version;
    }

    let url =
      `${gitApiUrl}/items?path=${safePath}` +
      `&download=false&recursionLevel=none` +
      `&versionDescriptor.version=${versionFixed}` +
      `&versionDescriptor.versionType=${version.versionType}`;
    try {
      let res = await TFSServices.getItemContent(url, this.token, 'get', {}, {}, false);
      return res ? true : false;
    } catch {
      return false;
    }
  }

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

  async GetItemsInCommitRange(
    projectId: string,
    repositoryId: string,
    commitRange: any,
    linkedWiOptions: any,
    includeUnlinkedCommits: boolean = false
  ) {
    logger.info(
      `GetItemsInCommitRange: includeUnlinkedCommits=${includeUnlinkedCommits}, commits=${
        commitRange?.value?.length || 0
      }`
    );
    //get all items linked to commits
    let res: any = [];
    let commitChangesArray: any = [];
    let commitsWithNoRelations: any[] = [];
    //extract linked items and append them to result
    for (const commit of commitRange.value) {
      if (commit.workItems && commit.workItems.length > 0) {
        for (const wi of commit.workItems) {
          let populatedItem = await this.ticketsDataProvider.GetWorkItem(projectId, wi.id);
          let linkedItems: LinkedRelation[] = await this.createLinkedRelatedItemsForSVD(
            linkedWiOptions,
            populatedItem
          );
          let changeSet: any = { workItem: populatedItem, commit: commit, linkedItems };
          commitChangesArray.push(changeSet);
        }
      } else {
        // Handle commits with no linked work items
        if (includeUnlinkedCommits) {
          commitsWithNoRelations.push({
            commitId: commit.commitId,
            commitDate: commit.committer?.date,
            committer: commit.committer?.name,
            comment: commit.comment,
            url: commit.remoteUrl,
          });
        }
      }
    }
    logger.info(
      `GetItemsInCommitRange: produced ${commitChangesArray.length} linked changes and ${commitsWithNoRelations.length} unlinked commits`
    );
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
    return { commitChangesArray: res, commitsWithNoRelations };
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

  //
  async getItemsForPipelineRange(
    teamProject: string,
    extendedCommits: any[],
    targetRepo: any,
    addedWorkItemByIdSet: Set<number>,
    linkedWiOptions: any = undefined,
    includeUnlinkedCommits: boolean = false
  ) {
    logger.info(
      `getItemsForPipelineRange: includeUnlinkedCommits=${includeUnlinkedCommits}, extendedCommits=${
        extendedCommits?.length || 0
      }`
    );
    let commitChangesArray: any[] = [];
    let commitsWithNoRelations: any[] = [];
    try {
      if (extendedCommits?.length === 0) {
        throw new Error('extended commits cannot be empty');
      }
      logger.debug(
        `getItemsForPipelineRange: ${extendedCommits?.length} commits for ${JSON.stringify(targetRepo)}`
      );
      //First fetch the repo web url
      if (targetRepo.url) {
        const repoData = await TFSServices.getItemContent(targetRepo.url, this.token);
        const repoWebUrl = repoData._links?.web.href;
        const targetRepoProjectId = repoData.project?.id;
        if (repoWebUrl) {
          targetRepo['url'] = repoWebUrl;
          targetRepo['projectId'] = targetRepoProjectId ?? teamProject;
        }
      }
      //Then extend the commit information with the related WIs
      for (const extendedCommit of extendedCommits) {
        const { commit } = extendedCommit;
        if (!Array.isArray(commit.workItems) || commit.workItems.length === 0) {
          if (includeUnlinkedCommits) {
            commitsWithNoRelations.push({
              commitId: commit.commitId,
              commitDate: commit.committer?.date,
              committer: commit.committer?.name,
              comment: commit.comment,
              url: commit.remoteUrl,
            });
          }
          continue;
        }
        for (const wi of commit.workItems) {
          const populatedWorkItem = await this.ticketsDataProvider.GetWorkItem(
            targetRepo['projectId'],
            wi.id
          );
          let linkedItems: LinkedRelation[] = await this.createLinkedRelatedItemsForSVD(
            linkedWiOptions,
            populatedWorkItem
          );
          let changeSet: any = {
            workItem: populatedWorkItem,
            commit: commit,
            targetRepo,
            linkedItems,
          };
          if (!addedWorkItemByIdSet.has(wi.id)) {
            addedWorkItemByIdSet.add(wi.id);
            commitChangesArray.push(changeSet);
          }
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
    logger.info(
      `getItemsForPipelineRange: produced ${commitChangesArray.length} linked changes and ${commitsWithNoRelations.length} unlinked commits`
    );
    return { commitChangesArray, commitsWithNoRelations };
  }

  /**
   * Creates a list of linked related items for a given work item (SVD).
   *
   * This method processes the relations of a populated work item and filters them
   * based on the provided options for linked work item types and relationships.
   * It then creates and returns an array of linked relations that match the criteria.
   *
   * @param linkedWiOptions - Options for filtering linked work items. Contains:
   *   - `linkedWiTypes`: Specifies the types of work items to include. Possible values:
   *     - `'none'`: Do not include any linked work items.
   *     - `'reqOnly'`: Include only work items of type 'Requirement'.
   *     - `'featureOnly'`: Include only work items of type 'Feature'.
   *     - `'both'`: Include both 'Requirement' and 'Feature' work items.
   *   - `linkedWiRelationship`: Specifies the relationship types to include. Possible values:
   *     - `'affectsOnly'`: Include only relations with 'Affects' in their name.
   *     - `'coversOnly'`: Include only relations with 'CoveredBy' in their name.
   *     - `'both'`: Include both 'Affects' and 'CoveredBy' relations.
   * @param wi - The work item for which linked related items are being created.
   * @param populatedWorkItem - The fully populated work item containing its relations.
   * @returns A promise that resolves to an array of `LinkedRelation` objects representing
   *          the filtered and created linked related items.
   */
  private async createLinkedRelatedItemsForSVD(linkedWiOptions: any, populatedWorkItem: any) {
    let linkedItems: LinkedRelation[] = [];
    try {
      if (linkedWiOptions?.isEnabled) {
        const { linkedWiTypes, linkedWiRelationship } = linkedWiOptions;

        // linkedWiTypes = {'none','reqOnly', 'featureOnly', 'both'};
        // linkedWiRelationship = {'affectsOnly', 'coversOnly', 'both' };
        if (linkedWiTypes !== 'none') {
          logger.debug(`Adding linked work items for ${populatedWorkItem.id}`);
          if (populatedWorkItem.relations) {
            for (const relation of populatedWorkItem.relations) {
              if (!relation.url.includes('/workItems/')) continue;

              const relatedItemContent: any = await this.ticketsDataProvider.GetWorkItemByUrl(relation.url);
              const wiItemType = relatedItemContent.fields['System.WorkItemType'];
              const relName = relation.rel;

              const isRequirement = wiItemType === 'Requirement';
              const isFeature = wiItemType === 'Feature';

              const shouldAddLinkedItem =
                (linkedWiTypes === 'reqOnly' && isRequirement) ||
                (linkedWiTypes === 'featureOnly' && isFeature) ||
                (linkedWiTypes === 'both' && (isRequirement || isFeature));

              if (!shouldAddLinkedItem) continue;

              const linkedItem = createLinkedRelation(
                relatedItemContent.id,
                wiItemType,
                relatedItemContent.fields['System.Title'],
                relatedItemContent._links?.html?.href || '',
                relation.attributes['name'] || ''
              );

              const isAffectsRelation = relName.includes('Affects');
              const isCoveredByRelation = relName.includes('CoveredBy');

              const shouldAddRelation =
                (linkedWiRelationship === 'affectsOnly' && isAffectsRelation) ||
                (linkedWiRelationship === 'coversOnly' && isCoveredByRelation) ||
                (linkedWiRelationship === 'both' && (isAffectsRelation || isCoveredByRelation));

              if (shouldAddRelation) {
                linkedItems.push(linkedItem);
              }
            }
          }
        }
      }
    } catch (ex) {
      logger.error(`Error creating linked related items: ${ex}`);
    }

    return linkedItems;
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

  async GetCommitsForRepo(projectName: string, repoID: string, versionIdentifier?: string) {
    let url: string;
    if (typeof versionIdentifier !== 'undefined' || versionIdentifier !== '') {
      url = `${this.orgUrl}${projectName}/_apis/git/repositories/${repoID}/commits?searchCriteria.$top=2000&searchCriteria.itemVersion.version=${versionIdentifier}`;
    } else {
      url = `${this.orgUrl}${projectName}/_apis/git/repositories/${repoID}/commits?searchCriteria.$top=2000`;
    }
    let res: any = await TFSServices.getItemContent(url, this.token, 'get', null, null);
    if (res.count === 0) {
      return [];
    }

    return res.value.map((commit: any) => {
      const dateStr = commit?.committer?.date || commit?.author?.date || undefined;
      return {
        name: `${commit.commitId.slice(0, 7)} - ${commit.comment}`,
        value: commit.commitId,
        date: dateStr,
      };
    });
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

  async GetRepoReferences(projectId: string, repoId: string, gitObjectType: string) {
    let url: string = '';
    switch (gitObjectType) {
      case 'tag':
        // peelTags=true ensures annotated tags are dereferenced to their target commit (peeledObjectId)
        url = `${this.orgUrl}${projectId}/_apis/git/repositories/${repoId}/refs/tags?peelTags=true&api-version=5.1`;
        break;
      case 'branch':
        url = `${this.orgUrl}${projectId}/_apis/git/repositories/${repoId}/refs/heads?api-version=5.1`;
        break;
      default:
        throw new Error(`Unsupported git object type: ${gitObjectType}`);
    }

    const res = await TFSServices.getItemContent(url, this.token, 'get');

    if (res.count === 0) {
      return [];
    }

    // For tags, sort by the commit date that the tag points to (most recent first).
    if (gitObjectType === 'tag') {
      // Resolve each tag to its commitId and fetch commit metadata for sorting
      const taggedWithDates = await Promise.all(
        res.value.map(async (refItem: any) => {
          // For annotated tags, prefer peeledObjectId; for lightweight tags, objectId is the commit
          const commitId = refItem.peeledObjectId || refItem.objectId;
          let ts = 0;
          let dateStr: string | null = null;
          try {
            const commit = await this.GetCommitByCommitId(projectId, repoId, commitId);
            const d = commit?.committer?.date || commit?.author?.date;
            ts = d ? new Date(d).getTime() : 0;
            dateStr = d || null;
          } catch {
            // If commit cannot be resolved, leave timestamp as 0
          }
          return { refItem, ts, dateStr };
        })
      );

      taggedWithDates.sort((a: any, b: any) => b.ts - a.ts);
      return taggedWithDates.map(({ refItem, dateStr }: any) => ({
        name: refItem.name.replace('refs/heads/', '').replace('refs/tags/', ''),
        value: refItem.name,
        date: dateStr || undefined,
      }));
    }

    // For branches, sort by the tip commit date (most recent first)
    if (gitObjectType === 'branch') {
      const branchesWithDates = await Promise.all(
        res.value.map(async (refItem: any) => {
          const commitId = refItem.objectId; // branch tip sha
          let ts = 0;
          let dateStr: string | null = null;
          try {
            const commit = await this.GetCommitByCommitId(projectId, repoId, commitId);
            const d = commit?.committer?.date || commit?.author?.date;
            ts = d ? new Date(d).getTime() : 0;
            dateStr = d || null;
          } catch {
            // ignore and keep ts=0
          }
          return { refItem, ts, dateStr };
        })
      );

      branchesWithDates.sort((a: any, b: any) => b.ts - a.ts);
      return branchesWithDates.map(({ refItem, dateStr }: any) => ({
        name: refItem.name.replace('refs/heads/', '').replace('refs/tags/', ''),
        value: refItem.name,
        date: dateStr || undefined,
      }));
    }

    // Default: return as-is
    return res.value.map((refItem: any) => ({
      name: refItem.name.replace('refs/heads/', '').replace('refs/tags/', ''),
      value: refItem.name,
    }));
  }

  async GetRepoTagsWithCommits(repoApiUrl: string) {
    // repoApiUrl is expected to be the full API base for the repository, e.g.
    //   http://server/tfs/Collection/Project/_apis/git/repositories/{repoId}
    const url = `${repoApiUrl}/refs/tags?peelTags=true&api-version=5.1`;
    const res = await TFSServices.getItemContent(url, this.token, 'get');
    if (!res || res.count === 0 || !Array.isArray(res.value)) {
      return [];
    }

    const out: Array<{ name: string; commitId: string; date?: string }> = [];
    for (const refItem of res.value) {
      const commitId = refItem.peeledObjectId || refItem.objectId;
      if (!commitId) {
        continue;
      }
      let dateStr: string | undefined;
      try {
        const commitUrl = `${repoApiUrl}/commits/${commitId}`;
        const commit = await TFSServices.getItemContent(commitUrl, this.token, 'get');
        const d = commit?.committer?.date || commit?.author?.date;
        if (d) {
          dateStr = d;
        }
      } catch {
        // ignore resolution failures, keep date undefined
      }

      out.push({
        name: (refItem.name || '').replace('refs/heads/', '').replace('refs/tags/', ''),
        commitId,
        date: dateStr,
      });
    }
    return out;
  }

  async GetCommitBatch(
    gitUrl: string,
    itemVersion: GitVersionDescriptor,
    compareVersion: GitVersionDescriptor,
    specificItemPath: string = ''
  ) {
    const allCommitsExtended: any[] = [];
    let skipping = 0;
    let chunkSize = 100;
    let commitCounter = 0;
    let logToConsoleAfterCommits = 500;
    try {
      let body =
        specificItemPath === ''
          ? {
              itemVersion,
              compareVersion,
              includeWorkItems: true,
            }
          : {
              itemVersion,
              compareVersion,
              includeWorkItems: true,
              itemPath: specificItemPath,
              historyMode: 'fullHistory',
            };

      let url = `${gitUrl}/commitsbatch?$skip=${skipping}&$top=${chunkSize}&api-version=5.1`;
      let commitsResponse = await TFSServices.postRequest(url, this.token, undefined, body);
      let commits = commitsResponse.data;
      while (commits.count > 0) {
        for (const commit of commits.value) {
          try {
            if (commitCounter % logToConsoleAfterCommits === 0) {
              logger.debug(`commit number ${commitCounter + 1}`);
            }

            let extendedCommit: any = {};

            extendedCommit['commit'] = commit;

            let committerName = commit.committer.name;
            let commitDate = commit.committer.date.toString().slice(0, 10);
            extendedCommit['committerName'] = committerName;
            extendedCommit['commitDate'] = commitDate;

            allCommitsExtended.push(extendedCommit);
          } catch (err: any) {
            const errMsg = `Cannot fetch commit batch: ${err.message}`;
            throw new Error(errMsg);
          } finally {
            commitCounter++;
          }
        }

        skipping += chunkSize;
        let url = `${gitUrl}/commitsbatch?$skip=${skipping}&$top=${chunkSize}&api-version=5.1`;
        commitsResponse = await TFSServices.postRequest(url, this.token, undefined, body);
        commits = commitsResponse.data;
      }
    } catch (error: any) {
      logger.error(error.message);
    }
    return allCommitsExtended;
  }

  async getSubmodulesData(
    projectName: string,
    repoId: string,
    targetVersion: GitVersionDescriptor,
    sourceVersion: GitVersionDescriptor,
    allCommitsExtended: any[]
  ) {
    let submodules: any[] = [];
    try {
      const gitModulesFile = await this.GetFileFromGitRepo(projectName, repoId, '.gitmodules', targetVersion);
      let gitRepoUrl = `${this.orgUrl}${projectName}/_apis/git/repositories/${repoId}`;
      if (!gitModulesFile) {
        // No submodules found
        return submodules;
      }
      const gitModulesFileLines = gitModulesFile.includes('\r\n')
        ? gitModulesFile.split('\r\n')
        : gitModulesFile.split('\n');

      logger.info(`generating submodules data for ${repoId}`);

      let gitSubModuleName = '';
      let gitSubPointerPath = '';

      for (const gitModuleLine of gitModulesFileLines) {
        if (gitModuleLine.startsWith('[submodule')) {
          gitSubModuleName = gitModuleLine
            .replace('[submodule "', '')
            .replace('"]', '')
            .replace('/', '_')
            .trim();
          continue;
        }
        if (gitModuleLine.includes('path = ')) {
          gitSubPointerPath = gitModuleLine.replace('path = ', '').trim();
        }
        if (!gitModuleLine.includes('url = ')) {
          // If the line does not contain the URL, skip it
          continue;
        }

        let gitSubRepoUrl = gitModuleLine.replace('url = ', '').trim();
        if (gitSubRepoUrl.startsWith('../')) {
          const relativePaths = gitSubRepoUrl.match(/\.\.\//g);
          if (relativePaths.length > 0) {
            let gitRepoUrlSplitted = gitRepoUrl.split('/');
            let gitSubRepoUrlPrefix = gitRepoUrlSplitted
              .slice(0, gitRepoUrlSplitted.length - relativePaths.length)
              .join('/');
            gitSubRepoUrl = gitSubRepoUrlPrefix + '/' + gitSubRepoUrl.replace('../', '');
          }
        }

        let targetSha1 = await this.GetFileFromGitRepo(projectName, repoId, gitSubPointerPath, targetVersion);
        let sourceSha1 = await this.GetFileFromGitRepo(projectName, repoId, gitSubPointerPath, sourceVersion);
        if (!sourceSha1) {
          for (
            let checkCommitIndex = allCommitsExtended.length - 1;
            checkCommitIndex > 0;
            checkCommitIndex--
          ) {
            let checkCommit = null;
            const { commit } = allCommitsExtended[checkCommitIndex];
            if (commit) {
              checkCommit = commit.commitId;
            }
            //In case of only commit object
            else if (allCommitsExtended[checkCommitIndex].commitId) {
              checkCommit = allCommitsExtended[checkCommitIndex].commitId;
            }
            if (!checkCommit) {
              logger.warn(`commit not found for ${gitSubModuleName}`);
              continue;
            }
            sourceSha1 = await this.GetFileFromGitRepo(projectName, repoId, gitSubPointerPath, {
              ...sourceVersion,
              version: checkCommit,
            });
            if (sourceSha1) {
              //found
              break;
            }
          }
        }

        if (!sourceSha1) {
          logger.warn(
            `${gitSubModuleName} pointer not exist in source version ${sourceVersion.versionType} ${sourceVersion.version} in repository ${gitRepoUrl}`
          );
          continue;
        }

        if (!targetSha1) {
          logger.warn(
            `${gitSubModuleName} pointer not exist in target version ${targetVersion.versionType} ${targetVersion.version} in repository ${gitRepoUrl}`
          );
          continue;
        }
        if (sourceSha1 === targetSha1) {
          logger.warn(
            `${gitSubModuleName} pointer is the same in source and target version in repository ${gitRepoUrl}`
          );
          continue;
        }

        let subModule = {
          sourceSha1,
          targetSha1,
          gitSubRepoUrl,
          gitSubRepoName: decodeURIComponent(gitSubRepoUrl?.split('/').pop() || ''),
          gitSubModuleName,
        };
        submodules.push(subModule);
      }
    } catch (error: any) {
      logger.error(`Error in getSubmodulesData: ${error.message}`);
    } finally {
      return submodules;
    }
  }
}
