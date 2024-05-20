
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
class QueryTree {
  queryType: QueryType;
  queryResultType: QueryResultType;
  asOf: Date;
  columns: Column[] = new Array<Column>();
  //sortColumns: Sortcolumn[]  =new Array<Sortcolumn>()
  workItems: Workitemrelation[] = new Array<Workitemrelation>();
}

export class Column {
  referenceName: string;
  name: string;
  url: string;
}

class Sortcolumn {
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
  level:number;
}
export class value {
  name: string;
  value: string;
}
enum QueryResultType {
  WorkItem = 1,
  WorkItemLink = 2
}
export enum QueryType {
  Flat = "flat",
  Tree = "tree",
  OneHop = "oneHop"
}

class Workitemrelation {
  rel: string;
  source: Source;
  target: Target;
}

class Workrelation {
  relation: Workitemrelation;
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
  suit: string;
  url: string;
  relations: Relation[];

  constructor() {
    this.relations = [];
  }
}

export type Relation = {
  id: string;
  title: string;
  customerId?: string;
}

export class TestSteps {
  action: String;
  expected: String;
}