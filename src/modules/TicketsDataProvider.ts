import { TFSServices } from '../helpers/tfs';
import { Workitem } from '../models/tfs-data';
import { Helper, suiteData, Links, Trace, Relations } from '../helpers/helper';
import { Query, TestSteps } from '../models/tfs-data';
import { QueryType } from '../models/tfs-data';
import { QueryAllTypes } from '../models/tfs-data';
import { Column } from '../models/tfs-data';
import { value } from '../models/tfs-data';
import { TestCase } from '../models/tfs-data';
import * as xml2js from 'xml2js';

import logger from '../utils/logger';
import { log } from 'console';
import { pid, title } from 'process';

export default class TicketsDataProvider {
  orgUrl: string = '';
  token: string = '';
  queriesList: Array<any> = new Array<any>();

  constructor(orgUrl: string, token: string) {
    this.orgUrl = orgUrl;
    this.token = token;
  }

  async GetWorkItem(project: string, id: string): Promise<any> {
    let url = `${this.orgUrl}${project}/_apis/wit/workitems/${id}?$expand=All`;
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

      if (linksMap.get(wis[i].id).rels.length > 0) {
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
  // gets queries recursiv
  async GetSharedQueries(project: string, path: string): Promise<any> {
    let url;
    try {
      if (path == '')
        url = `${this.orgUrl}${project}/_apis/wit/queries/Shared%20Queries?$depth=2&$expand=all`;
      else url = `${this.orgUrl}${project}/_apis/wit/queries/${path}?$depth=2&$expand=all`;
      let queries: any = await TFSServices.getItemContent(url, this.token);
      const { tree1: reqTestTree, tree2: testReqTree } = await this.structureQueries(queries);
      return { reqTestTree, testReqTree };
    } catch (err: any) {
      logger.error(err.message);
      logger.error(`Error stack trace:  ${JSON.stringify(err.stack)}`);
    }
  }

  async GetQueryResultsFromWiqlHref(projectName: string, wiqlHref: string = ''): Promise<any> {
    try {
      if (!wiqlHref) {
        throw new Error('Incorrect WIQL Link');
      }

      // Remember to add customer id if needed
      const queryResult: any = await TFSServices.getItemContent(wiqlHref, this.token);
      if (!queryResult) {
        throw new Error('Query result failed');
      }

      const { columns, workItemRelations } = queryResult;

      if (workItemRelations.length === 0) {
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
      for (const relation of workItemRelations) {
        //if relation.Source is null and target has a valid value then the target is the source
        if (!relation.source) {
          // Root link
          const wi: any = await this.fetchWIByRelation(relation, columnsToShowMap, columnSourceMap);
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
        const sourceRelation = lookupMap.get(relation.source.id);
        if (!sourceRelation) {
          throw new Error('Source relation has no mapping');
        }

        const targetWi: any = await this.fetchWIByRelation(relation, columnsToShowMap, columnTargetsMap);
        const targets: any = sourceTargetsMap.get(sourceRelation) || [];
        targets.push(targetWi);
        sourceTargetsMap.set(sourceRelation, targets);
      }
      columnsToShowMap.clear();
      return {
        sourceTargetsMap,
        sortingSourceColumnsMap: columnSourceMap,
        sortingTargetsColumnsMap: columnTargetsMap,
      };
    } catch (err: any) {
      logger.error(err.message);
    }
  }

  private async fetchWIByRelation(
    relation: any,
    columnMap: Map<string, string>,
    resultedRefNameMap: Map<string, string>
  ) {
    const url = `${relation.target.url}`;
    const wi: any = await TFSServices.getItemContent(url, this.token);
    if (!wi) {
      throw new Error(`WI ${relation.target.id} not found`);
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
    let attachmentList: Array<string> = [];
    let ticketsDataProvider = new TicketsDataProvider(this.orgUrl, this.token);
    try {
      let wi = await ticketsDataProvider.GetWorkItem(project, id);
      if (!wi.relations) return [];
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

  async GetWorkitemAttachmentsJSONData(project: string, attachmentId: string) {
    let wiuRL = `${this.orgUrl}${project}/_apis/wit/attachments/${attachmentId}`;
    let attachment = await TFSServices.getItemContent(wiuRL, this.token);
    return attachment;
  }

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

  private async structureQueries(rootQuery: any, parentId: any = null): Promise<any> {
    try {
      if (!rootQuery.hasChildren) {
        if (!rootQuery.isFolder && rootQuery.queryType === 'oneHop') {
          const wiql = rootQuery.wiql;
          let tree1Node = null;
          let tree2Node = null;

          if (this.matchesReqTestCondition(wiql)) {
            tree1Node = {
              id: rootQuery.id,
              pId: parentId,
              value: rootQuery.name,
              title: rootQuery.name,
              queryType: rootQuery.queryType,
              wiql: rootQuery._links.wiql ?? undefined,
              isValidQuery: true,
            };
          }
          if (this.matchesTestReqCondition(wiql)) {
            tree2Node = {
              id: rootQuery.id,
              pId: parentId,
              value: rootQuery.name,
              title: rootQuery.name,
              queryType: rootQuery.queryType,
              wiql: rootQuery._links.wiql ?? undefined,
              isValidQuery: true,
            };
          }
          return { tree1: tree1Node, tree2: tree2Node };
        } else {
          return { tree1: null, tree2: null };
        }
      }

      if (!rootQuery.children) {
        const queryUrl = `${rootQuery.url}?$depth=2&$expand=all`;
        const currentQuery = await TFSServices.getItemContent(queryUrl, this.token);
        return currentQuery
          ? await this.structureQueries(currentQuery, currentQuery.id)
          : { tree1: null, tree2: null };
      }

      // Process children recursively
      const childResults = await Promise.all(
        rootQuery.children.map((child: any) => this.structureQueries(child, rootQuery.id))
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

  private matchesReqTestCondition(wiql: string): boolean {
    return (
      wiql.includes("Source.[System.WorkItemType] = 'Requirement'") &&
      wiql.includes("Target.[System.WorkItemType] = 'Test Case'")
    );
  }

  private matchesTestReqCondition(wiql: string): boolean {
    return (
      wiql.includes("Source.[System.WorkItemType] = 'Test Case'") &&
      wiql.includes("Target.[System.WorkItemType] = 'Requirement'")
    );
  }

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
      logger.error('Error Stack: ', err.stack);
    }
  }
}
