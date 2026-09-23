// src/shared/types/Task.ts
import { SharePointPerson } from "./SharePointPerson";

export interface Task {
  Id: number;
  Title: string;
  ChangeRequestId: number;
  ParentTaskId?: number;
  PublishedDocumentId?: number;
  PublishedDocument?: {
    Id: number;
    DocumentTitle?: string;
  };
  // Expanded lookup field
  ChangeRequest?: {
    Id: number;
    ChangeRequestNumber?: string;
    Title?: string;
    DraftDocumentName?: string;
    DocumentNumber?: string;
  };
  TaskType:
    | "CA Review"
    | "Document Review"
    | "Final Approval"
    | "CR Completion"
    | "CR Info Required"
    | "Change Authority Approval"
    | "Change Authority Review"
    | "Document Controller Review"
    | "Compliance Authority Review"
    | "Publish Document"
    | "Author Review"
    | "Document Change Process"
    | "Participant Task"
    | "Publishing Rejection Review"
    | "Document Review";
  Role?:
    | "Reviewer"
    | "Contributor"
    | "Change Authority"
    | "Compliance Authority"
    | "Release Authority"
    | "Author"
    | "Requestor";

  AssignedTo?: SharePointPerson;
  Status:
    | "Pending"
    | "In Progress"
    | "Approved"
    | "Rejected"
    | "Reassigned"
    | "Complete"
    | "Cancelled"
    | "Needs more info"
    | "On Hold"
    | "New"
    | "Author Approved"
    | "Marked as Minor Change"
    | "Marked for Document Obsoletion"
    | "Start Change Request";
  Created: Date;
  DueDate?: Date;
  CompletedDate?: string;
  Requestor?: SharePointPerson;
  Comments?: string;
  isPinned?: boolean;
}