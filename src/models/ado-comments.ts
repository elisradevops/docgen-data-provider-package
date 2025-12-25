export type AdoCommentFormat = 'html' | 'markdown' | 'text' | string;

export interface AdoIdentityRef {
  displayName?: string;
  uniqueName?: string;
  id?: string;
  descriptor?: string;
  url?: string;
  imageUrl?: string;
}

export interface AdoWorkItemComment {
  workItemId?: number;
  id?: number;
  version?: number;
  text?: string;
  renderedText?: string;
  format?: AdoCommentFormat;
  createdBy?: AdoIdentityRef;
  createdDate?: string;
  modifiedBy?: AdoIdentityRef;
  modifiedDate?: string;
  isDeleted?: boolean;
  url?: string;
}

export interface AdoWorkItemCommentsResponse {
  totalCount?: number;
  count?: number;
  comments?: AdoWorkItemComment[];
  continuationToken?: string;
}

