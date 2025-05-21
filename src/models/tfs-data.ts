export class QueryAllTypes {
  q: Query;
  source: number;
}

export class Query {
  queryType: QueryType;
  queryResultType: QueryResultType;
  asOf: Date;
  columns: Column[] = new Array<Column>();
  //sortColumns: Sortcolumn[]  =new Array<Sortcolumn>()
  workItems: Workitem[] = new Array<Workitem>();
}
export class QueryTree {
  queryType: QueryType;
  queryResultType: QueryResultType;
  asOf: string;
  columns: Column[];
  sortColumns?: SortColumn[];
  workItems?: WorkItemForQuery[];
  workItemRelations?: WorkItemRelation[];
}

export class Column {
  referenceName: string;
  name: string;
  url: string;
}

class SortColumn {
  field: Field;
  descending: boolean;
}

class Field {
  referenceName: string;
  name: string;
  url: string;
}

export class Workitem {
  id: number;
  url: string;
  parentId: number;
  fields: value[];
  Source: number = 0;
  attachments: any[];
  level: number;
}
export class value {
  name: string;
  value: string;
}
enum QueryResultType {
  WorkItem = 1,
  WorkItemLink = 2,
}
export enum QueryType {
  Flat = 'flat',
  Tree = 'tree',
  OneHop = 'oneHop',
}

class WorkItemForQuery {
  id: number;
  url: string;
}

class WorkItemRelation {
  rel: string;
  source: Source;
  target: Target;
}

class Workrelation {
  relation: WorkItemRelation;
  value: value;
}

class Source {
  id: number;
  url: string;
}

class Target {
  id: number;
  url: string;
}

export interface Comment {
  parentCommentId: number;
  content: string;
  commentType: number;
}

export interface RightFileEnd {
  line: number;
  offset: number;
}

export interface RightFileStart {
  line: number;
  offset: number;
}

export interface ThreadContext {
  filePath: string;
  leftFileEnd?: any;
  leftFileStart?: any;
  rightFileEnd: RightFileEnd;
  rightFileStart: RightFileStart;
}

export interface IterationContext {
  firstComparingIteration: number;
  secondComparingIteration: number;
}

export interface PullRequestThreadContext {
  changeTrackingId: number;
  iterationContext: IterationContext;
}

export interface RootObject {
  comments: Comment[];
  status: number;
  threadContext: ThreadContext;
  pullRequestThreadContext: PullRequestThreadContext;
}
export class TestCase {
  id: string;
  title: string;
  description: string;
  area: string;
  steps: TestSteps[];
  caseEvidenceAttachments: any[] = [];
  suit: string;
  url: string;
  relations: Relation[];

  constructor() {
    this.relations = [];
  }
}

type Relation = RequirementRelation | BugRelation | MomRelation;

export type LinkedRelation = {
  id: string;
  wiType: string;
  title: string;
  url: string;
  relationType: string;
};

type RequirementRelation = {
  type: 'requirement';
  id: string;
  title: string;
  customerId?: string;
};

type MomRelation = {
  type: string;
  id: string;
  title: string;
  url: string;
  status: string;
};

type BugRelation = {
  type: 'bug';
  id: string;
  title: string;
  severity?: string;
};

export type OpenPcrRequest = {
  openPcrMode: string;
  testToOpenPcrQuery: string;
  OpenPcrToTestQuery: string;
  includeCommonColumnsMode: string;
};

export function createLinkedRelation(
  id: string,
  wiType: string,
  title: string,
  url: string,
  relationType: string
): LinkedRelation {
  return { relationType, wiType, id, title, url };
}

export function createRequirementRelation(
  id: string,
  title: string,
  customerId?: string
): RequirementRelation {
  return { type: 'requirement', id, title, customerId };
}

export function createMomRelation(
  id: string,
  type: string,
  title: string,
  url: string,
  status: string
): MomRelation {
  return { type, id, title, url, status };
}

export function createBugRelation(id: string, title: string, severity?: string): BugRelation {
  return { type: 'bug', id, title, severity };
}

export class TestSteps {
  stepId: string;
  stepPosition: string;
  action: string;
  expected: string;
  isSharedStepTitle: boolean;
}

export interface GitVersionDescriptor {
  version: string;
  versionType: string;
}

export interface Pipeline {
  id: number;
  revision: number;
  name: string;
}

export interface Repository {
  id: string;
  name: string;
  url: string;
}

export interface ResourceRepository {
  repoName: string;
  repoSha1: string;
  url: string;
}

export interface Link {
  href: string;
}

export interface PipelineLinks {
  self: Link;
  web: Link;
  'pipeline.web': Link;
  pipeline: Link;
}

export interface PipelineResources {
  repositories: any;
  pipelines: any;
}

export interface PipelineRun {
  _links: PipelineLinks;
  pipeline: Pipeline;
  state: string;
  result: string;
  createdDate: string;
  finishedDate: string;
  url: string;
  resources: PipelineResources;
  id: number;
  name: string;
}

export interface Artifact {
  alias: string;
  isPrimary: boolean;
  type: string;
}
