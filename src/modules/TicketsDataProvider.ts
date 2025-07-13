import { TFSServices } from '../helpers/tfs';
import { Workitem, QueryTree } from '../models/tfs-data';
import { Helper, Links, Trace, Relations } from '../helpers/helper';
import { Query } from '../models/tfs-data';
import { QueryType } from '../models/tfs-data';
import { QueryAllTypes } from '../models/tfs-data';
import { Column } from '../models/tfs-data';
import { value } from '../models/tfs-data';

import logger from '../utils/logger';

export default class TicketsDataProvider {
  orgUrl: string = '';
  token: string = '';
  queriesList: Array<any> = new Array<any>();

  constructor(orgUrl: string, token: string) {
    this.orgUrl = orgUrl;
    this.token = token;
  }

  async FetchImageAsBase64(url: string): Promise<string> {
    let image = await TFSServices.fetchAzureDevOpsImageAsBase64(url, this.token, 'get', null);
    return image;
  }

  async GetWorkItem(project: string, id: string): Promise<any> {
    let url = `${this.orgUrl}${project}/_apis/wit/workitems/${id}?$expand=All`;
    return TFSServices.getItemContent(url, this.token);
  }

  async GetWorkItemByUrl(url: string): Promise<any> {
    return TFSServices.getItemContent(url, this.token);
  }

  async GetLinksByIds(project: string, ids: any) {
    var trace: Array<Trace> = new Array<Trace>();
    let wis = await this.PopulateWorkItemsByIds(ids, project);
    let linksMap: any = await this.GetRelationsIds(wis);

    let relations;
    for (let i = 0; i < wis.length; i++) {
      let traceItem: Trace; //= new Trace();
      traceItem = await this.GetParentLink(project, wis[i]);

      if (linksMap.get(wis[i].id)?.rels.length > 0) {
        relations = await this.PopulateWorkItemsByIds(linksMap.get(wis[i].id).rels, project);
        traceItem.links = await this.GetLinks(project, wis[i], relations);
      }
      trace.push(traceItem);
    }
    return trace;
  }

  async GetParentLink(project: string, wi: any) {
    let trace: Trace = new Trace();
    if (wi != null) {
      trace.id = wi.id;
      trace.title = wi.fields['System.Title'];
      trace.url = this.orgUrl + project + '/_workitems/edit/' + wi.id;
      if (wi.fields['System.CustomerId'] != null && wi.fields['System.CustomerId'] != undefined) {
        trace.customerId = wi.fields['System.CustomerId'];
      }
    }
    return trace;
  }
  async GetRelationsIds(ids: any) {
    let rel = new Map<string, Relations>();
    try {
      for (let i = 0; i < ids.length; i++) {
        var link = new Relations();
        link.id = ids[i].id;
        if (ids[i].relations != null)
          for (let j = 0; j < ids[i].relations.length; j++) {
            if (ids[i].relations[j].rel != 'AttachedFile') {
              let index = ids[i].relations[j].url.lastIndexOf('/');
              let id = ids[i].relations[j].url.substring(index + 1);
              link.rels.push(id);
            }
          }
        rel.set(ids[i].id, link);
      }
    } catch (e) {}
    return rel;
  }
  async GetLinks(project: string, wi: any, links: any) {
    var linkList: Array<Links> = new Array<Links>();
    for (let i = 0; i < wi.relations.length; i++) {
      for (let j = 0; j < links.length; j++) {
        let index = wi.relations[i].url.lastIndexOf('/');
        let linkId = wi.relations[i].url.substring(index + 1);
        if (linkId == links[j].id) {
          var link = new Links();
          link.type = wi.relations[i].rel;
          link.id = links[j].id;
          link.title = links[j].fields['System.Title'];
          link.description = links[j].fields['System.Description'];
          link.url = this.orgUrl + project + '/_workitems/edit/' + linkId;
          linkList.push(link);
          break;
        }
      }
    }
    return linkList;
  }
  /**
   * Getting shared queries
   * @param project project name
   * @param path query path
   * @param docType document type
   * @returns
   */
  async GetSharedQueries(project: string, path: string, docType: string = ''): Promise<any> {
    let url;
    try {
      if (path === '')
        url = `${this.orgUrl}${project}/_apis/wit/queries/Shared%20Queries?$depth=2&$expand=all`;
      else url = `${this.orgUrl}${project}/_apis/wit/queries/${path}?$depth=2&$expand=all`;
      let queries: any = await TFSServices.getItemContent(url, this.token);
      logger.debug(`doctype: ${docType}`);
      switch (docType?.toLowerCase()) {
        case 'std':
          return await this.fetchLinkedReqTestQueries(queries, false);
        case 'str':
          const reqTestTrees = await this.fetchLinkedReqTestQueries(queries, false);
          const openPcrTestTrees = await this.fetchLinkedOpenPcrTestQueries(queries, false);
          return { reqTestTrees, openPcrTestTrees };
        case 'test-reporter':
          const testAssociatedTree = await this.fetchTestReporterQueries(queries);
          return { testAssociatedTree };
        case 'svd':
          return await this.fetchAnyQueries(queries);
        default:
          break;
      }
    } catch (err: any) {
      logger.error(`Error occurred during fetching shared queries: ${err.message}`);
      throw err;
    }
  }

  /**
   * Fetches the fields of a specific work item type.
   *
   * @param project - The project name.
   * @param itemType - The work item type.
   * @returns An array of objects containing the field name and reference name.
   */

  async GetFieldsByType(project: string, itemType: string) {
    try {
      let url = `${this.orgUrl}${project}/_apis/wit/workitemtypes/${itemType}/fields`;
      const { value: fields } = await TFSServices.getItemContent(url, this.token);
      return (
        fields
          // filter out the fields that are not relevant for the user
          .filter(
            (field: any) =>
              field.name !== 'ID' &&
              field.name !== 'Title' &&
              field.name !== 'Description' &&
              field.name !== 'Work Item Type' &&
              field.name !== 'Steps'
          )
          .map((field: any) => {
            return {
              text: `${field.name} (${itemType})`,
              key: field.referenceName,
            };
          })
      );
    } catch (err: any) {
      logger.error(`Error occurred during fetching fields by type: ${err.message}`);
      throw err;
    }
  }

  /**
   * fetches linked queries
   * @param queries fetched queries
   * @param onlyTestReq get only test req
   * @returns ReqTestTree and TestReqTree
   */
  private async fetchLinkedReqTestQueries(queries: any, onlyTestReq: boolean = false) {
    const { tree1: reqTestTree, tree2: testReqTree } = await this.structureFetchedQueries(
      queries,
      onlyTestReq,
      null,
      ['Requirement'],
      ['Test Case']
    );
    return { reqTestTree, testReqTree };
  }

  /**
   * Fetches and structures linked queries related to open PCR (Problem Change Request) tests.
   *
   * This method retrieves and organizes the relationships between "Test Case" and
   * other entities such as "Bug" and "Change Request" into two tree structures.
   *
   * @param queries - The input queries to be processed and structured.
   * @param onlySourceSide - A flag indicating whether to process only the source side of the queries.
   *                          Defaults to `false`.
   * @returns An object containing two tree structures:
   *          - `OpenPcrToTestTree`: The tree representing the relationship from Open PCR to Test Case.
   *          - `TestToOpenPcrTree`: The tree representing the relationship from Test Case to Open PCR.
   */
  private async fetchLinkedOpenPcrTestQueries(queries: any, onlySourceSide: boolean = false) {
    const { tree1: OpenPcrToTestTree, tree2: TestToOpenPcrTree } = await this.structureFetchedQueries(
      queries,
      onlySourceSide,
      null,
      ['Bug', 'Change Request'],
      ['Test Case']
    );
    return { OpenPcrToTestTree, TestToOpenPcrTree };
  }

  /**
   * fetches test reporter queries
   * @param queries  fetched queries
   * @returns
   */
  private async fetchTestReporterQueries(queries: any) {
    const { tree1: tree1, tree2: testAssociatedTree } = await this.structureFetchedQueries(
      queries,
      true,
      null,
      ['Requirement', 'Bug', 'Change Request'],
      ['Test Case']
    );
    return { testAssociatedTree };
  }

  private async fetchAnyQueries(queries: any) {
    const { tree1: systemOverviewQueryTree, tree2: knownBugsQueryTree } = await this.structureAllQueryPath(
      queries
    );
    return { systemOverviewQueryTree, knownBugsQueryTree };
  }

  async GetQueryResultsFromWiql(
    wiqlHref: string = '',
    displayAsTable: boolean = false,
    testCaseToRequirementMap: Map<number, Set<any>>
  ): Promise<any> {
    try {
      if (!wiqlHref) {
        throw new Error('Incorrect WIQL Link');
      }
      // Remember to add customer id if needed
      const queryResult: QueryTree = await TFSServices.getItemContent(wiqlHref, this.token);
      if (!queryResult) {
        throw new Error('Query result failed');
      }

      switch (queryResult.queryType) {
        case QueryType.OneHop:
          return displayAsTable
            ? await this.parseDirectLinkedQueryResultForTableFormat(queryResult, testCaseToRequirementMap)
            : await this.parseTreeQueryResult(queryResult);
        case QueryType.Tree:
          return await this.parseTreeQueryResult(queryResult);
        case QueryType.Flat:
          return displayAsTable
            ? await this.parseFlatQueryResultForTableFormat(queryResult)
            : await this.parseFlatQueryResult(queryResult);
        default:
          break;
      }
    } catch (err: any) {
      logger.error(`Could not fetch query results for ${wiqlHref}: ${err.message}`);
    }
  }

  private async parseDirectLinkedQueryResultForTableFormat(
    queryResult: QueryTree,
    testCaseToRequirementMap: Map<number, Set<any>>
  ) {
    const { columns, workItemRelations } = queryResult;

    if (workItemRelations?.length === 0) {
      throw new Error('No related work items were found');
    }

    const columnsToShowMap: Map<string, string> = new Map();

    const columnSourceMap: Map<string, string> = new Map();
    const columnTargetsMap: Map<string, string> = new Map();

    //The map is copied because there are different fields for each WI type,
    //Need to consider both ways (Req->Test; Test->Req)
    columns.forEach((column: any) => {
      const { referenceName, name } = column;
      if (name === 'CustomerRequirementId') {
        columnsToShowMap.set(referenceName, 'Customer ID');
      } else {
        columnsToShowMap.set(referenceName, name);
      }
    });

    // Initialize maps
    const sourceTargetsMap: Map<any, any[]> = new Map();
    const lookupMap: Map<number, any> = new Map();
    if (workItemRelations) {
      for (const relation of workItemRelations) {
        //if relation.Source is null and target has a valid value then the target is the source
        logger.debug(`DEBUG ONLY: relation.source: ${JSON.stringify(relation.source)}`);
        logger.debug(`DEBUG ONLY: relation.target: ${JSON.stringify(relation.target)}`);
        if (!relation.source) {
          // Root link
          const wi: any = await this.fetchWIForQueryResult(relation, columnsToShowMap, columnSourceMap, true);
          if (!lookupMap.has(wi.id)) {
            sourceTargetsMap.set(wi, []);
            lookupMap.set(wi.id, wi);
          }
          continue; // Move to the next relation
        }

        if (!relation.target) {
          throw new Error('Target relation is missing');
        }

        // Get relation source from lookup
        const sourceWorkItem = lookupMap.get(relation.source.id);
        if (!sourceWorkItem) {
          throw new Error('Source relation has no mapping');
        }

        const targetWi: any = await this.fetchWIForQueryResult(
          relation,
          columnsToShowMap,
          columnTargetsMap,
          true
        );
        //In case if source is a test case
        this.mapTestCaseToRequirement(sourceWorkItem, testCaseToRequirementMap, targetWi);

        //In case of target is a test case
        this.mapTestCaseToRequirement(targetWi, testCaseToRequirementMap, sourceWorkItem);
        const targets: any = sourceTargetsMap.get(sourceWorkItem) || [];
        logger.debug(`DEBUG ONLY: targets: ${JSON.stringify(targets)}`);
        targets.push(targetWi);
        logger.debug(`DEBUG ONLY: targets after push: ${JSON.stringify(targets)}`);
        sourceTargetsMap.set(sourceWorkItem, targets);
      }
    }

    columnsToShowMap.clear();
    return {
      sourceTargetsMap,
      sortingSourceColumnsMap: columnSourceMap,
      sortingTargetsColumnsMap: columnTargetsMap,
    };
  }

  private mapTestCaseToRequirement(
    testCaseItem: any,
    testCaseToRequirementMap: Map<number, Set<any>>,
    RequirementWi: any
  ) {
    if (testCaseItem.fields['System.WorkItemType'] == 'Test Case') {
      if (!testCaseToRequirementMap.has(testCaseItem.id)) {
        testCaseToRequirementMap.set(testCaseItem.id, new Set());
      }
      const requirementSet = testCaseToRequirementMap.get(testCaseItem.id);
      if (requirementSet) {
        // Check if there's already an item with the same ID
        const alreadyExists = [...requirementSet].some((reqItem) => reqItem.id === RequirementWi.id);

        if (!alreadyExists) {
          requirementSet.add(RequirementWi);
        }
      }
    }
  }

  private async parseFlatQueryResultForTableFormat(queryResult: QueryTree) {
    const { columns, workItems } = queryResult;

    if (workItems?.length === 0) {
      throw new Error('No work items were found');
    }

    const columnsToShowMap: Map<string, string> = new Map();

    const fieldsToIncludeMap: Map<string, string> = new Map();

    //The map is copied because there are different fields for each WI type,
    columns.forEach((column: any) => {
      const { referenceName, name } = column;
      if (name === 'CustomerRequirementId') {
        columnsToShowMap.set(referenceName, 'Customer ID');
      } else {
        columnsToShowMap.set(referenceName, name);
      }
    });

    // Initialize maps
    const wiSet: Set<any> = new Set();
    if (workItems) {
      for (const workItem of workItems) {
        const wi: any = await this.fetchWIForQueryResult(
          workItem,
          columnsToShowMap,
          fieldsToIncludeMap,
          false
        );
        wiSet.add(wi);
      }
    }

    columnsToShowMap.clear();
    return {
      fetchedWorkItems: [...wiSet],
      fieldsToIncludeMap,
    };
  }

  private async parseTreeQueryResult(queryResult: QueryTree) {
    const { workItemRelations } = queryResult;
    if (!workItemRelations) {
      logger.warn(`No work items were found for this requested query`);
      return null;
    }

    try {
      // Step 1: Identify roots and prepare a map of all nodes
      const allItems: any = {};
      const roots = [];
      for (const relation of workItemRelations) {
        const target = relation.target;
        // Initialize the node if it doesn't exist yet
        if (!allItems[target.id]) {
          await this.initTreeQueryResultItem(target, allItems);
        }

        if (relation.rel === null && relation.source === null) {
          roots.push(allItems[target.id]);
        } else if (relation.source) {
          // This means target is a child of source
          const parentId = relation.source.id;
          if (!allItems[parentId]) {
            const source = relation.source;
            await this.initTreeQueryResultItem(source, allItems);
          }
          // Add the target as a child of the parent
          allItems[parentId].children.push(allItems[target.id]);
        }
      }

      return roots;
    } catch (error: any) {
      logger.error(`could not parse requested tree query: ${error.message}`);
    }
  }

  private async initTreeQueryResultItem(item: any, allItems: any) {
    const urlWi = `${item.url}?fields=System.Description,System.Title,Microsoft.VSTS.TCM.ReproSteps,Microsoft.VSTS.CMMI.Symptom`;
    const wi = await TFSServices.getItemContent(urlWi, this.token);
    // need to fetch the WI with only the the title, the web URL and the description
    allItems[item.id] = {
      id: item.id,
      title: wi.fields['System.Title'] || '',
      description: wi.fields['Microsoft.VSTS.CMMI.Symptom'] ?? wi.fields['System.Description'] ?? '',
      htmlUrl: wi._links.html.href,
      children: [],
    };
  }

  private async initFlatQueryResultItem(item: any, workItemMap: Map<number, any>) {
    const urlWi = `${item.url}?fields=System.Description,System.Title,Microsoft.VSTS.TCM.ReproSteps,Microsoft.VSTS.CMMI.Symptom`;
    const wi = await TFSServices.getItemContent(urlWi, this.token);
    // need to fetch the WI with only the the title, the web URL and the description
    workItemMap.set(item.id, {
      id: item.id,
      title: wi.fields['System.Title'] || '',
      description: wi.fields['Microsoft.VSTS.CMMI.Symptom'] ?? wi.fields['System.Description'] ?? '',
      htmlUrl: wi._links.html.href,
    });
  }

  private async parseFlatQueryResult(queryResult: QueryTree) {
    const { workItems } = queryResult;
    if (!workItems) {
      logger.warn(`No work items were found for this requested query`);
      return null;
    }

    try {
      const workItemsResultMap: Map<number, any> = new Map();
      for (const wi of workItems) {
        if (!workItemsResultMap.has(wi.id)) {
          await this.initFlatQueryResultItem(wi, workItemsResultMap);
        }
      }
      return [...workItemsResultMap.values()];
    } catch (error: any) {
      logger.error(`could not parse requested flat query: ${error.message}`);
    }
  }

  private async fetchWIForQueryResult(
    receivedObject: any,
    columnMap: Map<string, string>,
    resultedRefNameMap: Map<string, string>,
    isRelation: boolean
  ) {
    const url = isRelation ? `${receivedObject.target.url}` : `${receivedObject.url}`;
    const wi: any = await TFSServices.getItemContent(url, this.token);
    if (!wi) {
      throw new Error(`WI ${isRelation ? receivedObject.target.id : receivedObject.id} not found`);
    }

    this.filterFieldsByColumns(wi, columnMap, resultedRefNameMap);
    return wi;
  }

  // get queris structured
  GetModeledQuery(list: Array<any>): Array<any> {
    let queryListObject: Array<any> = [];
    list.forEach((query) => {
      let newObj = {
        queryName: query.name,
        wiql: query._links.wiql != null ? query._links.wiql : null,
        id: query.id,
      };
      queryListObject.push(newObj);
    });
    return queryListObject;
  }
  // gets query results
  async GetQueryResultsByWiqlHref(wiqlHref: string, project: string): Promise<any> {
    try {
      let results: any = await TFSServices.getItemContent(wiqlHref, this.token);
      let modeledResult = await this.GetModeledQueryResults(results, project);
      if (modeledResult.queryType == 'tree') {
        let levelResults: Array<Workitem> = Helper.LevelBuilder(
          modeledResult,
          modeledResult.workItems[0].fields[0].value
        );
        return levelResults;
      }
      return modeledResult.workItems;
    } catch (e) {}
  }
  // gets query results
  async GetQueryResultsByWiqlString(wiql: string, projectName: string): Promise<any> {
    let res;
    let url = `${this.orgUrl}${projectName}/_apis/wit/wiql?$top=2147483646&expand=all`;
    try {
      res = await TFSServices.getItemContent(url, this.token, 'post', {
        query: wiql,
      });
    } catch (error) {
      console.log(error);
      return [];
    }
    return res;
  }
  async GetQueryResultById(query: string, project: string): Promise<any> {
    var url = `${this.orgUrl}${project}/_apis/wit/queries/${query}`;
    let querie: any = await TFSServices.getItemContent(url, this.token);
    var wiql = querie._links.wiql;
    return await this.GetQueryResultsByWiqlHref(wiql.href, project);
  }

  async PopulateWorkItemsByIds(workItemsArray: any[] = [], projectName: string = ''): Promise<any[]> {
    let url = `${this.orgUrl}${projectName}/_apis/wit/workitemsbatch`;
    let res: any[] = [];
    let divByMax = Math.floor(workItemsArray.length / 200);
    let modulusByMax = workItemsArray.length % 200;
    //iterating
    for (let i = 0; i < divByMax; i++) {
      let from = i * 200;
      let to = (i + 1) * 200;
      let currentIds = workItemsArray.slice(from, to);
      try {
        let subRes = await TFSServices.getItemContent(url, this.token, 'post', {
          $expand: 'Relations',
          ids: currentIds,
        });
        res = [...res, ...subRes.value];
      } catch (error) {
        logger.error(`error populating workitems array`);
        logger.error(JSON.stringify(error));
        return [];
      }
    }
    //compliting the rimainder
    if (modulusByMax !== 0) {
      try {
        let currentIds = workItemsArray.slice(workItemsArray.length - modulusByMax, workItemsArray.length);
        let subRes = await TFSServices.getItemContent(url, this.token, 'post', {
          $expand: 'Relations',
          ids: currentIds,
        });
        res = [...res, ...subRes.value];
      } catch (error) {
        logger.error(`error populating workitems array`);
        logger.error(JSON.stringify(error));
        return [];
      }
    } //if

    return res;
  }

  async GetModeledQueryResults(results: any, project: string) {
    let ticketsDataProvider = new TicketsDataProvider(this.orgUrl, this.token);
    var queryResult: Query = new Query();
    queryResult.asOf = results.asOf;
    queryResult.queryResultType = results.queryResultType;
    queryResult.queryType = results.queryType;
    if (results.queryType == QueryType.Flat) {
      //     //Flat Query
      //TODo: attachment
      //TODO:check if wi.relations exist
      //TODO: add attachment to any list from 1
      for (var j = 0; j < results.workItems.length; j++) {
        let wi = await ticketsDataProvider.GetWorkItem(project, results.workItems[j].id);

        queryResult.workItems[j] = new Workitem();
        queryResult.workItems[j].url = results.workItems[j].url;
        queryResult.workItems[j].fields = new Array(results.columns.length);
        if (wi.relations != null) {
          queryResult.workItems[j].attachments = wi.relations;
        }
        var rel = new QueryAllTypes();
        for (var i = 0; i < results.columns.length; i++) {
          queryResult.columns[i] = new Column();
          queryResult.workItems[j].fields[i] = new value();
          queryResult.columns[i].name = results.columns[i].name;
          queryResult.columns[i].referenceName = results.columns[i].referenceName;
          queryResult.columns[i].url = results.columns[i].url;
          if (results.columns[i].referenceName.toUpperCase() == 'SYSTEM.ID') {
            queryResult.workItems[j].fields[i].value = wi.id.toString();
            queryResult.workItems[j].fields[i].name = 'ID';
          } else if (
            results.columns[i].referenceName.toUpperCase() == 'SYSTEM.ASSIGNEDTO' &&
            wi.fields[results.columns[i].referenceName] != null
          )
            queryResult.workItems[j].fields[i].value =
              wi.fields[results.columns[i].referenceName].displayName;
          else {
            let s: string = wi.fields[results.columns[i].referenceName];
            queryResult.workItems[j].fields[i].value = wi.fields[results.columns[i].referenceName];
            queryResult.workItems[j].fields[i].name = results.columns[i].name;
          }
        }
      }
    } //Tree Query
    else {
      this.BuildColumns(results, queryResult);
      for (var j = 0; j < results.workItemRelations.length; j++) {
        if (results.workItemRelations[j].target != null) {
          let wiT = await ticketsDataProvider.GetWorkItem(project, results.workItemRelations[j].target.id);
          // var rel = new QueryAllTypes();
          queryResult.workItems[j] = new Workitem();
          queryResult.workItems[j].url = wiT.url;
          queryResult.workItems[j].fields = new Array(results.columns.length);
          if (wiT.relations != null) {
            queryResult.workItems[j].attachments = wiT.relations;
          }
          // rel.q = queryResult;
          for (i = 0; i < queryResult.columns.length; i++) {
            //..  rel.q.workItems[j].fields[i] = new value();
            queryResult.workItems[j].fields[i] = new value();
            queryResult.workItems[j].fields[i].name = queryResult.columns[i].name;
            if (
              results.columns[i].referenceName.toUpperCase() == 'SYSTEM.ASSIGNEDTO' &&
              wiT.fields[results.columns[i].referenceName] != null
            )
              queryResult.workItems[j].fields[i].value =
                wiT.fields[results.columns[i].referenceName].displayName;
            else queryResult.workItems[j].fields[i].value = wiT.fields[queryResult.columns[i].referenceName];
            //}
          }
          if (results.workItemRelations[j].source != null)
            queryResult.workItems[j].Source = results.workItemRelations[j].source.id;
        }
      }
    }
    return queryResult;
  }

  //Build columns
  BuildColumns(results: any, queryResult: Query) {
    for (var i = 0; i < results.columns.length; i++) {
      queryResult.columns[i] = new Column();
      queryResult.columns[i].name = results.columns[i].name;
      queryResult.columns[i].referenceName = results.columns[i].referenceName;
      queryResult.columns[i].url = results.columns[i].url;
    }
  }

  async GetIterationsByTeamName(projectName: string, teamName: string): Promise<any[]> {
    let res: any;
    let url;
    if (teamName) {
      url = `${this.orgUrl}${projectName}/${teamName}/_apis/work/teamsettings/iterations`;
    } else {
      url = `${this.orgUrl}${projectName}/_apis/work/teamsettings/iterations`;
    }
    res = await TFSServices.getItemContent(url, this.token, 'get');
    return res;
  } //GetIterationsByTeamName

  async CreateNewWorkItem(projectName: string, wiBody: any, wiType: string, byPass: boolean) {
    let url = `${this.orgUrl}${projectName}/_apis/wit/workitems/$${wiType}?bypassRules=${String(
      byPass
    ).toString()}`;
    return TFSServices.getItemContent(url, this.token, 'POST', wiBody, {
      'Content-Type': 'application/json-patch+json',
    });
  } //CreateNewWorkItem

  async GetWorkitemAttachments(project: string, id: string) {
    let attachmentList: Array<any> = [];
    let ticketsDataProvider = new TicketsDataProvider(this.orgUrl, this.token);
    try {
      let wi = await ticketsDataProvider.GetWorkItem(project, id);
      if (!wi?.relations) return [];
      await Promise.all(
        wi.relations.map(async (relation: any) => {
          if (relation.rel == 'AttachedFile') {
            let attachment = JSON.parse(JSON.stringify(relation));
            attachment.downloadUrl = `${relation.url}/${relation.attributes.name}`;
            attachmentList.push(attachment);
          }
        })
      );
      return attachmentList;
    } catch (e) {
      logger.error(`error fetching attachments for work item ${id}`);
      logger.error(`${JSON.stringify(e)}`);
      return [];
    }
  }

  //Get work item attachments
  async GetWorkitemAttachmentsJSONData(project: string, attachmentId: string) {
    let wiuRL = `${this.orgUrl}${project}/_apis/wit/attachments/${attachmentId}`;
    let attachment = await TFSServices.getItemContent(wiuRL, this.token);
    return attachment;
  }

  //Update work item
  async UpdateWorkItem(projectName: string, wiBody: any, workItemId: number, byPass: boolean) {
    let res: any;
    let url: string = `${this.orgUrl}${projectName}/_apis/wit/workitems/${workItemId}?bypassRules=${String(
      byPass
    ).toString()}`;
    res = await TFSServices.getItemContent(url, this.token, 'patch', wiBody, {
      'Content-Type': 'application/json-patch+json',
    });
    return res;
  } //CreateNewWorkItem

  private async structureAllQueryPath(rootQuery: any, parentId: any = null): Promise<any> {
    try {
      if (!rootQuery.hasChildren) {
        if (!rootQuery.isFolder) {
          let sysOverviewNode = null;
          let knownBugsNode = null;
          if (rootQuery.queryType === 'flat' && this.matchesBugCondition(rootQuery.wiql)) {
            //Add this to the known bugs query tree
            knownBugsNode = {
              id: rootQuery.id,
              pId: parentId,
              value: rootQuery.name,
              title: rootQuery.name,
              queryType: rootQuery.queryType,
              columns: rootQuery.columns,
              wiql: rootQuery._links.wiql ?? undefined,
              isValidQuery: true,
            };
          }
          sysOverviewNode = {
            id: rootQuery.id,
            pId: parentId,
            value: rootQuery.name,
            title: rootQuery.name,
            queryType: rootQuery.queryType,
            columns: rootQuery.columns,
            wiql: rootQuery._links.wiql ?? undefined,
            isValidQuery: true,
          };

          return { tree1: sysOverviewNode, tree2: knownBugsNode };
        } else {
          return { tree1: null, tree2: null };
        }
      }

      if (!rootQuery.children) {
        const queryUrl = `${rootQuery.url}?$depth=2&$expand=all`;
        const currentQuery = await TFSServices.getItemContent(queryUrl, this.token);
        return currentQuery
          ? await this.structureAllQueryPath(currentQuery, currentQuery.id)
          : { tree1: null, tree2: null };
      }

      // Process children recursively
      const childResults = await Promise.all(
        rootQuery.children.map((child: any) => this.structureAllQueryPath(child, rootQuery.id))
      );

      // Build tree
      const sysOverviewNodeTreeChildren = childResults
        .map((res: any) => res.tree1)
        .filter((child: any) => child !== null);
      const sysOverviewNode =
        sysOverviewNodeTreeChildren.length > 0
          ? {
              id: rootQuery.id,
              pId: parentId,
              value: rootQuery.name,
              title: rootQuery.name,
              children: sysOverviewNodeTreeChildren,
              columns: rootQuery.columns,
            }
          : null;

      const knownBugsTreeChildren = childResults
        .map((res: any) => res.tree2)
        .filter((child: any) => child !== null);
      const knownBugs =
        knownBugsTreeChildren.length > 0
          ? {
              id: rootQuery.id,
              pId: parentId,
              value: rootQuery.name,
              title: rootQuery.name,
              children: knownBugsTreeChildren,
              columns: rootQuery.columns,
            }
          : null;

      return { tree1: sysOverviewNode, tree2: knownBugs };
    } catch (err: any) {
      logger.error(
        `Error occurred while constructing the query list ${err.message} with query ${JSON.stringify(
          rootQuery
        )}`
      );
      throw err;
    }
  }

  /**
   * Recursively structures fetched queries into two hierarchical trees (tree1 and tree2)
   * based on specific conditions. It processes both leaf and non-leaf nodes, fetching
   * children if necessary, and builds the trees by matching source and target conditions.
   *
   * @param rootQuery - The root query object to process. It may contain children or be a leaf node.
   * @param onlyTestReq - A boolean flag indicating whether to exclude requirement-to-test-case queries.
   * @param parentId - The ID of the parent node, used to maintain the hierarchy. Defaults to `null`.
   * @param sources - An array of source strings used to match queries.
   * @param targets - An array of target strings used to match queries.
   * @returns A promise that resolves to an object containing two trees (`tree1` and `tree2`),
   *          or `null` for each tree if no valid nodes are found.
   * @throws Logs an error if an exception occurs during processing.
   */
  private async structureFetchedQueries(
    rootQuery: any,
    onlyTestReq: boolean,
    parentId: any = null,
    sources: string[],
    targets: string[]
  ): Promise<any> {
    try {
      if (!rootQuery.hasChildren) {
        if (!rootQuery.isFolder && rootQuery.queryType === 'oneHop') {
          const wiql = rootQuery.wiql;
          let tree1Node = null;
          let tree2Node = null;
          // Check if the query is a requirement to test case query
          if (!onlyTestReq && this.matchesSourceTargetCondition(wiql, sources, targets)) {
            tree1Node = {
              id: rootQuery.id,
              pId: parentId,
              value: rootQuery.name,
              title: rootQuery.name,
              queryType: rootQuery.queryType,
              columns: rootQuery.columns,
              wiql: rootQuery._links.wiql ?? undefined,
              isValidQuery: true,
            };
          }
          if (this.matchesSourceTargetCondition(wiql, targets, sources)) {
            tree2Node = {
              id: rootQuery.id,
              pId: parentId,
              value: rootQuery.name,
              title: rootQuery.name,
              queryType: rootQuery.queryType,
              columns: rootQuery.columns,
              wiql: rootQuery._links.wiql ?? undefined,
              isValidQuery: true,
            };
          }
          return { tree1: tree1Node, tree2: tree2Node };
        } else {
          return { tree1: null, tree2: null };
        }
      }
      // If the query has children, but they are not loaded, fetch them
      if (!rootQuery.children) {
        const queryUrl = `${rootQuery.url}?$depth=2&$expand=all`;
        const currentQuery = await TFSServices.getItemContent(queryUrl, this.token);
        return currentQuery
          ? await this.structureFetchedQueries(currentQuery, onlyTestReq, currentQuery.id, sources, targets)
          : { tree1: null, tree2: null };
      }

      // Process children recursively
      const childResults = await Promise.all(
        rootQuery.children.map((child: any) =>
          this.structureFetchedQueries(child, onlyTestReq, rootQuery.id, sources, targets)
        )
      );

      // Build tree1
      const tree1Children = childResults.map((res: any) => res.tree1).filter((child: any) => child !== null);
      const tree1Node =
        tree1Children.length > 0
          ? {
              id: rootQuery.id,
              pId: parentId,
              value: rootQuery.name,
              title: rootQuery.name,
              children: tree1Children,
            }
          : null;

      // Build tree2
      const tree2Children = childResults.map((res: any) => res.tree2).filter((child: any) => child !== null);
      const tree2Node =
        tree2Children.length > 0
          ? {
              id: rootQuery.id,
              value: rootQuery.name,
              pId: parentId,
              title: rootQuery.name,
              children: tree2Children,
            }
          : null;

      return { tree1: tree1Node, tree2: tree2Node };
    } catch (err: any) {
      logger.error(
        `Error occurred while constructing the query list ${err.message} with query ${JSON.stringify(
          rootQuery
        )}`
      );
      logger.error(`Error stack ${err.message}`);
    }
  }

  /**
   * Determines whether the given WIQL (Work Item Query Language) string matches the specified
   * source and target conditions. It checks if the WIQL contains references to the specified
   * source and target work item types.
   *
   * @param wiql - The WIQL string to evaluate.
   * @param source - An array of source work item types to check for in the WIQL.
   * @param target - An array of target work item types to check for in the WIQL.
   * @returns A boolean indicating whether the WIQL includes at least one source work item type
   *          and at least one target work item type.
   */
  private matchesSourceTargetCondition(wiql: string, source: string[], target: string[]): boolean {
    let isSourceIncluded = false;
    let isTargetIncluded = false;

    for (const src of source) {
      if (wiql.includes(`Source.[System.WorkItemType] = '${src}'`)) {
        isSourceIncluded = true;
        break;
      }
    }

    for (const tgt of target) {
      if (wiql.includes(`Target.[System.WorkItemType] = '${tgt}'`)) {
        isTargetIncluded = true;
        break;
      }
    }

    return isSourceIncluded && isTargetIncluded;
  }

  // check if the query is a bug query
  private matchesBugCondition(wiql: string): boolean {
    return wiql.includes(`[System.WorkItemType] = 'Bug'`);
  }

  // Filter the fields based on the columns to filter map
  private filterFieldsByColumns(
    item: any,
    columnsToFilterMap: Map<string, string>,
    resultedRefNameMap: Map<string, string>
  ) {
    try {
      const parsedFields: any = {};
      for (const fieldName of Object.keys(item.fields)) {
        const value = item.fields[fieldName];

        if (columnsToFilterMap.has(fieldName)) {
          resultedRefNameMap.set(fieldName, columnsToFilterMap.get(fieldName) || '');
          parsedFields[fieldName] = value;
        }
      }
      item.fields = { ...parsedFields };
    } catch (err: any) {
      logger.error(`Cannot filter columns: ${err.message}`);
      throw err;
    }
  }
}
