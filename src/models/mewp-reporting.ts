export type MewpRunStatus = 'Pass' | 'Fail' | 'Not Run';

export interface MewpExternalFileRef {
  url?: string;
  text?: string;
  name?: string;
  bucketName?: string;
  objectName?: string;
  etag?: string;
  contentType?: string;
  sizeBytes?: number;
  sourceType?: 'mewpExternalIngestion' | 'generic';
}

export interface MewpCoverageRequestOptions {
  useRelFallback?: boolean;
  externalBugsFile?: MewpExternalFileRef | null;
  externalL3L4File?: MewpExternalFileRef | null;
}

export interface MewpInternalValidationRequestOptions {
  useRelFallback?: boolean;
}

export interface MewpRequirementStepSummary {
  passed: number;
  failed: number;
  notRun: number;
}

export interface MewpL2RequirementWorkItem {
  workItemId: number;
  requirementId: string;
  baseKey: string;
  title: string;
  subSystem: string;
  responsibility: string;
  linkedTestCaseIds: number[];
  relatedWorkItemIds: number[];
  areaPath: string;
}

export interface MewpL2RequirementFamily {
  workItemId?: number;
  requirementId: string;
  baseKey: string;
  title: string;
  subSystem: string;
  responsibility: string;
  linkedTestCaseIds: number[];
}

export interface MewpLinkedRequirementEntry {
  baseKeys: Set<string>;
  fullCodes: Set<string>;
  bugIds: Set<number>;
}

export type MewpLinkedRequirementsByTestCase = Map<number, MewpLinkedRequirementEntry>;
export type MewpRequirementIndex = Map<string, Map<number, MewpRequirementStepSummary>>;

export interface MewpBugLink {
  id: number;
  title: string;
  responsibility: string;
  requirementBaseKey?: string;
}

export interface MewpL3L4Link {
  id: string;
  title: string;
  level: 'L3' | 'L4';
}

export interface MewpL3L4Pair {
  l3Id: string;
  l3Title: string;
  l4Id: string;
  l4Title: string;
}

export interface MewpCoverageBugCell {
  id: number | '';
  title: string;
  responsibility: string;
}

export type MewpCoverageL3L4Cell = MewpL3L4Pair;

export interface MewpCoverageRow {
  'L2 REQ ID': string;
  'L2 REQ Title': string;
  'L2 SubSystem': string;
  'L2 Run Status': MewpRunStatus;
  'Bug ID': number | '';
  'Bug Title': string;
  'Bug Responsibility': string;
  'L3 REQ ID': string;
  'L3 REQ Title': string;
  'L4 REQ ID': string;
  'L4 REQ Title': string;
}

export interface MewpCoverageFlatPayload {
  sheetName: string;
  columnOrder: string[];
  rows: MewpCoverageRow[];
}

export interface MewpInternalValidationRow {
  'Test Case ID': number;
  'Test Case Title': string;
  'Mentioned but Not Linked': string;
  'Linked but Not Mentioned': string;
  'Validation Status': 'Pass' | 'Fail';
}

export interface MewpInternalValidationFlatPayload {
  sheetName: string;
  columnOrder: string[];
  rows: MewpInternalValidationRow[];
}

export interface MewpExternalTableValidationResult {
  tableType: 'bugs' | 'l3l4';
  sourceName: string;
  valid: boolean;
  headerRow: 'A3' | 'A1' | '';
  matchedRequiredColumns: number;
  totalRequiredColumns: number;
  missingRequiredColumns: string[];
  rowCount: number;
  message: string;
}

export interface MewpExternalFilesValidationResponse {
  valid: boolean;
  bugs?: MewpExternalTableValidationResult;
  l3l4?: MewpExternalTableValidationResult;
}
