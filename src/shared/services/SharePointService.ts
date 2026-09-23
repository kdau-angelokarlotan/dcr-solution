import { PnPSetup } from "./PnPSetup";
import { SharePointPerson } from "../types/SharePointPerson";
import { Document } from "../types/Document";
import { Department } from "../types/Department";
import { IChangeRequest } from "../types/ChangeRequest";
import "@pnp/sp/attachments";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import "@pnp/sp/folders";
import "@pnp/sp/files";
import "@pnp/sp/sites";
import "@pnp/sp/search";
import { Task } from "../types/Task";
import { Participant } from "../types/Participant";
import { MinorChange } from "../types/MinorChange";
import { LookupFieldItem } from "../types/LookupFieldItem";
import { ReadAcknowledgement } from "../types/ReadAcknowledgement";

const normalizeServerRelativePath = (path: string): string => {
  const trimmed = path.trim();

  // Strip origin from full URLs (handles spaces that break new URL())
  const originMatch = trimmed.match(/^https?:\/\/[^/]+(\/.*)/i);
  if (originMatch) {
    return originMatch[1];
  }

  if (trimmed.startsWith("/")) {
    return trimmed;
  }

  if (trimmed.startsWith("sites/") || trimmed.startsWith("teams/")) {
    return `/${trimmed}`;
  }

  return `/${trimmed.replace(/^\/+/, "")}`;
};

// Cached web server-relative URL (e.g. "/sites/DocumentChangeManagementDemo")
let _webServerRelativeUrl: string | undefined;

const getWebServerRelativeUrl = async (): Promise<string> => {
  if (_webServerRelativeUrl) return _webServerRelativeUrl;
  const sp = PnPSetup.getSP();
  const webInfo = await sp.web.select("ServerRelativeUrl")();
  const url = webInfo.ServerRelativeUrl;
  if (!_webServerRelativeUrl) {
    _webServerRelativeUrl = url;
  }
  return _webServerRelativeUrl;
};

/**
 * Ensure a path is fully server-relative by prepending the web's server-relative URL
 * if the path is only site-relative (e.g. "/Change Request Documents/...").
 */
const ensureServerRelativePath = async (path: string): Promise<string> => {
  const normalized = normalizeServerRelativePath(path);
  const webUrl = await getWebServerRelativeUrl();
  if (normalized.startsWith(webUrl)) return normalized;
  return `${webUrl}${normalized}`;
};

const getDepartments = async (): Promise<Department[]> => {
  try {
    const sp = PnPSetup.getSP();
    const departments = await sp.web.lists
      .getByTitle("Departments Configuration")
      .items.select(
        "Id",
        "Title",
        "ChangeAuthority/Id", // ← Select person fields
        "ChangeAuthority/Title",
        "ChangeAuthority/EMail",
      )
      .expand("ChangeAuthority") // ← CRITICAL: Must expand!
      .orderBy("Title", true)();

    return departments as Department[];
  } catch (error) {
    console.error("Error fetching departments:", error);
    throw error;
  }
};

const getAudienceGroups = async (): Promise<LookupFieldItem[]> => {
  try {
    const sp = PnPSetup.getSP();
    const items = await sp.web.lists
      .getByTitle("Audience Groups Configuration")
      .items.select("Id", "Title")
      .orderBy("Title", true)
      .top(100)();
    return items as LookupFieldItem[];
  } catch (error) {
    console.error("Error fetching audience groups:", error);
    throw error;
  }
};

const getChangeRequests = async (): Promise<IChangeRequest[]> => {
  try {
    const sp = PnPSetup.getSP();
    const items = await sp.web.lists
      .getByTitle("Change Requests")
      .items.select(
        "Id",
        "Title",
        "ChangeRequestNumber",
        "ScopeofChange",
        "Status",
        "Urgency",
        "NewDocument",
        "Classification",
        "DraftDocumentName",
        "ReviewPeriod",
        "ReadAcknowledgementRequired",
        "ReadDueDate",
        "Created",
        "Author0/Id",
        "Author0/Title",
        "Author0/EMail",
        "ChangeAuthority/Id",
        "ChangeAuthority/Title",
        "ChangeAuthority/EMail",
        "ReleaseAuthority/Id",
        "ReleaseAuthority/Title",
        "ReleaseAuthority/EMail",
        "DocumentType/Id",
        "DocumentType/Title",
        "Category/Id",
        "Category/Title",
        "Audience/Id",
        "Audience/Title",
        "ReadAudienceGroups/Id",
        "ReadAudienceGroups/Title",
        "BusinessFunction/Id",
        "BusinessFunction/Title",
        "CoreFunctionality/Id",
        "CoreFunctionality/Title",
        "VersionNumber",
        "DocumentNumber"
      )
      .expand(
        "Author0",
        "ChangeAuthority",
        "ReleaseAuthority",
        "DocumentType",
        "Category",
        "Audience",
        "ReadAudienceGroups",
        "BusinessFunction",
        "CoreFunctionality",
      )
      .orderBy("Created", false)
      .top(500)();

    return items as IChangeRequest[];
  } catch (error) {
    console.error("Error fetching change requests:", error);
    throw error;
  }
};

const getChangeRequestById = async (
  id: number,
): Promise<IChangeRequest | null> => {
  try {
    const sp = PnPSetup.getSP();
    const changeRequest = await sp.web.lists
      .getByTitle("Change Requests")
      .items.getById(id)
      .select(
        "Id",
        "Title",
        "ChangeRequestNumber",
        "ScopeofChange",
        "Status",
        "Urgency",
        "NewDocument",
        "Classification",
        "DraftDocumentName",
        "PublishedDate",
        "ReviewPeriod",
        "ReadAcknowledgementRequired",
        "ReadDueDate",
        "DownloadFormat",
        // Lookup fields
        "CoreFunctionality/Id",
        "CoreFunctionality/Title",
        "BusinessFunction/Id",
        "BusinessFunction/Title",
        "DocumentType/Id", // ← add
        "DocumentType/Title", // ← add
        "Category/Id", // ← add
        "Category/Title", // ← add
        "Audience/Id",
        "Audience/Title",
        "ReadAudienceGroups/Id",
        "ReadAudienceGroups/Title",
        // People fields
        "ChangeAuthority/Id",
        "ChangeAuthority/Title",
        "ChangeAuthority/EMail",
        "ReleaseAuthority/Id",
        "ReleaseAuthority/Title",
        "ReleaseAuthority/EMail",
        "Author0/Id",
        "Author0/Title",
        "Author0/EMail",
        "Author/Id",
        "Author/Title",
        "Author/EMail",
        "Reviewers/Id",
        "Reviewers/Title",
        "Reviewers/EMail",
        "Contributors/Id",
        "Contributors/Title",
        "Contributors/EMail",
        // Document reference
        "TargetDocumentId",
        "DraftDocumentUrl",
        "DraftFolderUrl",
        "VersionNumber",
        "DocumentNumber"
      )
      .expand(
        "CoreFunctionality",
        "BusinessFunction",
        "DocumentType", // ← add
        "Category", // ← add
        "Audience",
        "ReadAudienceGroups",
        "ChangeAuthority",
        "ReleaseAuthority",
        "Author0",
        "Author", // ← add (needed for Submitted By)
        "Reviewers",
        "Contributors",
      )();

    return changeRequest as IChangeRequest;
  } catch (error) {
    console.error(`Error fetching change request with ID ${id}:`, error);
    return null;
  }
};

const createChangeRequest = async (
  data: Record<string, unknown>,
): Promise<Record<string, unknown>> => {
  try {
    const sp = PnPSetup.getSP();

    const result = await sp.web.lists
      .getByTitle("Change Requests")
      .items.add(data);

    return result as unknown as Record<string, unknown>;
  } catch (error) {
    console.error("Error creating Change Request:", error);
    throw error;
  }
};

const updateChangeRequest = async (
  id: number,
  data: Record<string, unknown>,
): Promise<void> => {
  try {
    const sp = PnPSetup.getSP();
    await sp.web.lists
      .getByTitle("Change Requests")
      .items.getById(id)
      .update(data);
  } catch (error) {
    console.error("Error updating change request:", error);
    throw error;
  }
};

const getDocuments = async (): Promise<Document[]> => {
  try {
    const sp = PnPSetup.getSP();
    const documents = await sp.web.lists
      .getByTitle("Published Documents")
      .items.select(
        "Id",
        "DocumentTitle",
        "PublishedDate",
        "Modified",
        "FileRef", // ← ADD THIS
        "FileLeafRef", // ← ADD THIS
        "PublishedFileUrl",
        "DownloadFileUrl",
        "SourceFileUrl",
        "DownloadFormat",
        // Lookup fields
        "DocumentType/Id",
        "DocumentType/Title",
        "Category/Id",
        "Category/Title",
        "BusinessFunction/Id",
        "BusinessFunction/Title",
        "CoreFunctionality/Id",
        "CoreFunctionality/Title",
        "Audience/Id",
        "Audience/Title",
        // Choice field (no expand needed)
        "Classification",
        "VersionNumber",
        "DocumentNumber"
      )
      .expand(
        "DocumentType",
        "Category",
        "BusinessFunction",
        "CoreFunctionality",
        "Audience",
      )
      .filter("FSObjType eq 1 and PublishedFileUrl ne null")
      .top(500)
      .orderBy("DocumentTitle", true)();
    console.log(documents);
    return documents as unknown as Document[];
  } catch (error) {
    console.error("Error fetching documents:", error);
    throw error;
  }
};

const getDocumentById = async (id: number): Promise<Document | null> => {
  try {
    const sp = PnPSetup.getSP();
    const document = await sp.web.lists
      .getByTitle("Published Documents")
      .items.getById(id)
      .select(
        "Id",
        "DocumentTitle",

        "Classification",
        "PublishedFileUrl",
        "DownloadFileUrl",
        "SourceFileUrl",
        "DownloadFormat",
        // Expand lookup fields
        "DocumentType/Id",
        "DocumentType/Title",
        "Category/Id",
        "Category/Title",
        "Audience/Id",
        "Audience/Title",
        "BusinessFunction/Id",
        "BusinessFunction/Title",
        "CoreFunctionality/Id",
        "CoreFunctionality/Title",
        // Expand person fields
        "ChangeAuthority/Id",
        "ChangeAuthority/Title",
        "ChangeAuthority/EMail",
        "ReleaseAuthority/Id",
        "ReleaseAuthority/Title",
        "ReleaseAuthority/EMail",
        "Author0/Id",
        "Author0/Title",
        "Author0/EMail",
        "FileRef",
        "FileLeafRef",
        "VersionNumber",
        "DocumentNumber"
      )
      .expand(
        "Category",
        "Audience",
        "DocumentType",
        "BusinessFunction",
        "CoreFunctionality",
        "ChangeAuthority",
        "ReleaseAuthority",
        "Author0",
      )();
    console.log(document);
    return document as Document;
  } catch (error) {
    console.error(`Error fetching document with ID ${id}:`, error);
    throw error;
  }
};

export interface ContentSearchResult {
  title: string;
  path: string;
  snippet: string;
}

const searchDocumentContent = async (
  query: string,
): Promise<ContentSearchResult[]> => {
  try {
    const sp = PnPSetup.getSP();
    const webUrl = await getWebServerRelativeUrl();
    const libraryPath = `${window.location.origin}${webUrl}/Published Documents`;

    const results = await sp.search({
      Querytext: `${query}* path:"${libraryPath}"`,
      SelectProperties: ["Path", "Title", "HitHighlightedSummary"],
      RowLimit: 50,
      TrimDuplicates: false,
    });

    const primaryResults = results.PrimarySearchResults as Array<{
      Path?: string;
      Title?: string;
      HitHighlightedSummary?: string;
    }>;

    return primaryResults.map((r) => ({
      title: r.Title ?? "Untitled",
      path: r.Path ?? "",
      snippet: r.HitHighlightedSummary ?? "",
    }));
  } catch (error) {
    console.error("Error searching document content:", error);
    return [];
  }
};

const getTasks = async (userId: number): Promise<Task[]> => {
  try {
    const sp = PnPSetup.getSP();
    const tasks = await sp.web.lists
      .getByTitle("Tasks")
      .items.select(
        "Id",
        "Title",
        "Status",
        "TaskType",
        "Role",
        "DueDate",
        "Created",
        "Comments",
        "ChangeRequestId",
        "ChangeRequest/Id",
        "isPinned",
        "ChangeRequest/ChangeRequestNumber",
        "ChangeRequest/Title",
        "ChangeRequest/DocumentNumber",
        "AssignedTo/Id",
        "AssignedTo/Title",
        "AssignedTo/EMail",
        "Author/Id",
        "Author/EMail",
        "Requestor/Id",
        "Requestor/Title",
        "Requestor/EMail",
        "PublishedDocumentId",
      )
      .expand("AssignedTo", "Author", "Requestor", "ChangeRequest")
      .filter(
        `AssignedTo/Id eq ${userId} and (Status eq 'Pending' or Status eq 'In Progress')`,
      )();
    return tasks as Task[];
  } catch (error) {
    console.error("Error fetching tasks:", error);
    throw error;
  }
};

const getPendingReadAcknowledgements = async (
  userId: number,
): Promise<ReadAcknowledgement[]> => {
  try {
    const sp = PnPSetup.getSP();
    const items = await sp.web.lists
      .getByTitle("Read Acknowledgements")
      .items.select(
        "Id",
        "Title",
        "Acknowledged",
        "AcknowledgedDate",
        "DocumentVersion",
        "Person/Id",
        "Person/Title",
        "Person/EMail",
        "PublishedDocumentId/Id",
        "PublishedDocumentId/DocumentTitle",
        "PublishedDocumentId/PublishedFileUrl",
        "ReadRequirementsId/Id",
        "ReadRequirementsId/DueDate",
      )
      .expand("Person", "PublishedDocumentId", "ReadRequirementsId")
      .filter(`Person/Id eq ${userId} and Acknowledged eq 0`)
      .top(100)();

    return items as unknown as ReadAcknowledgement[];
  } catch (error) {
    console.error("Error fetching pending read acknowledgements:", error);
    return [];
  }
};

const acknowledgeDocument = async (
  acknowledgementId: number,
  documentVersion?: string,
): Promise<void> => {
  try {
    const sp = PnPSetup.getSP();
    await sp.web.lists
      .getByTitle("Read Acknowledgements")
      .items.getById(acknowledgementId)
      .update({
        Acknowledged: true,
        AcknowledgedDate: new Date().toISOString(),
        DocumentVersion: documentVersion ?? null,
      });
  } catch (error) {
    console.error("Error acknowledging document:", error);
    throw error;
  }
};

const getTaskById = async (id: number): Promise<Task | null> => {
  try {
    const sp = PnPSetup.getSP();
    const task = await sp.web.lists
      .getByTitle("Tasks")
      .items.select(
        "Id",
        "Title",
        "Status",
        "TaskType",
        "Role",
        "DueDate",
        "Created",
        "Comments",
        "ChangeRequestId",
        "ChangeRequest/Id",
        "ChangeRequest/Title",
        "ChangeRequest/DocumentNumber",
        "AssignedTo/Id",
        "AssignedTo/Title",
        "AssignedTo/EMail",
        "Author/Id",
        "Author/EMail",
        "Requestor/Id",
        "Requestor/Title",
        "Requestor/EMail",
      )
      .expand("AssignedTo", "Author", "Requestor")
      .filter(
        `Id eq ${id} and (Status eq 'Pending' or Status eq 'In Progress')`,
      )();
    return task.length > 0 ? (task[0] as Task) : null;
  } catch (error) {
    console.error(`Error fetching task with ID ${id}:`, error);
    throw error;
  }
};

const updateTask = async (
  taskId: number,
  data: Record<string, unknown>,
): Promise<void> => {
  try {
    const sp = PnPSetup.getSP();
    await sp.web.lists.getByTitle("Tasks").items.getById(taskId).update(data);
  } catch (error) {
    console.error("Error updating task:", error);
    throw error;
  }
};

const uploadAttachments = async (
  itemId: number,
  files: File[],
  onProgress?: (current: number, total: number, fileName: string) => void,
): Promise<void> => {
  const sp = PnPSetup.getSP();
  if (!itemId || itemId <= 0) {
    throw new Error("Invalid list item id for attachment upload.");
  }

  if (!files || files.length === 0) {
    return;
  }

  const list = sp.web.lists.getByTitle("Change Requests");
  const listInfo = await list.select("EnableAttachments")();
  if (!listInfo.EnableAttachments) {
    await list.update({ EnableAttachments: true });
  }

  const item = list.items.getById(itemId);
  for (let i = 0; i < files.length; i++) {
    const file = files[i];

    if (onProgress) {
      onProgress(i + 1, files.length, file.name);
    }

    const fileBuffer = await file.arrayBuffer();
    await item.attachmentFiles.add(file.name, fileBuffer);
  }
};

const getCurrentUser = async (): Promise<SharePointPerson> => {
  try {
    const sp = PnPSetup.getSP();
    const user = await sp.web.currentUser();
    console.log("Current user info:", user);
    return user;
  } catch (error) {
    console.error("Error fetching current user:", error);
    throw error;
  }
};
const searchUsers = async (searchText: string): Promise<SharePointPerson[]> => {
  try {
    const sp = PnPSetup.getSP();
    const search = searchText.toLowerCase();

    const users = await sp.web.siteUsers
      .select("Id", "Title", "EMail")
      .top(500)();

    const filtered = (users as Array<{ Title?: string; EMail?: string }>).filter(
      (u) =>
        u.Title?.toLowerCase().includes(search) ||
        u.EMail?.toLowerCase().includes(search),
    );

    console.log("Search results:", filtered);
    return filtered as SharePointPerson[];
  } catch (error) {
    console.error("Error searching users:", error);
    throw error;
  }
};

const getAttachments = async (
  itemId: number,
): Promise<{ FileName: string; ServerRelativeUrl: string }[]> => {
  try {
    const sp = PnPSetup.getSP();
    const attachments = await sp.web.lists
      .getByTitle("Change Requests")
      .items.getById(itemId)
      .attachmentFiles();
    return attachments;
  } catch (error) {
    console.error("Error fetching attachments:", error);
    return [];
  }
};

const getParticipants = async (
  changeRequestId: number,
): Promise<Participant[]> => {
  try {
    const sp = PnPSetup.getSP();
    const participants = await sp.web.lists
      .getByTitle("CR Participants")
      .items.select(
        "Id",
        "Title",
        "ChangeRequestId",
        "Role",
        "Status",
        "DueDate",
        "StartDate",
        "CompletedDate",
        "Instructions",
        "Notes",
        "Person/Id",
        "Person/Title",
        "Person/EMail",
      )
      .expand("Person")
      .filter(`ChangeRequestId eq ${changeRequestId} and Status ne 'Removed'`)();

    return participants as Participant[];
  } catch (error) {
    console.error("Error fetching participants:", error);
    throw error;
  }
};

const createParticipant = async (
  changeRequestId: number,
  contributorIds: number[],
  reviewerIds: number[],
): Promise<void> => {
  const sp = PnPSetup.getSP();
  const list = sp.web.lists.getByTitle("CR Participants");

  const all = [
    ...contributorIds.map((id) => ({ PersonId: id, Role: "Contributor" })),
    ...reviewerIds.map((id) => ({ PersonId: id, Role: "Reviewer" })),
  ];

  await Promise.all(
    all.map((p) =>
      list.items.add({
        ChangeRequestId: changeRequestId,
        PersonId: p.PersonId,
        Role: p.Role,
        Status: "Not Started",
      }),
    ),
  );
};

const addParticipant = async (
  changeRequestId: number,
  personId: number,
  role: "Reviewer" | "Contributor",
): Promise<void> => {
  const sp = PnPSetup.getSP();
  await sp.web.lists.getByTitle("CR Participants").items.add({
    ChangeRequestId: changeRequestId,
    PersonId: personId,
    Role: role,
    Status: "Not Started",
  });
};

const updateParticipant = async (
  id: number,
  data: Partial<Participant>,
): Promise<void> => {
  try {
    const sp = PnPSetup.getSP();
    await sp.web.lists
      .getByTitle("CR Participants")
      .items.getById(id)
      .update(data);
  } catch (error) {
    console.error("Error updating participant:", error);
    throw error;
  }
};

const deleteParticipant = async (
  id: number,
  changeRequestId: number,
  personId: number,
  reason: string,
  removedByName: string,
): Promise<void> => {
  try {
    const sp = PnPSetup.getSP();

    const tasks = await sp.web.lists
      .getByTitle("Tasks")
      .items.select("Id", "Status", "AssignedTo/Id")
      .expand("AssignedTo")
      .filter(
        `ChangeRequestId eq ${changeRequestId} and TaskType eq 'Participant Task'`,
      )();

    const activeTask = tasks.find(
      (t) =>
        t.AssignedTo?.Id === personId &&
        t.Status !== "Complete" &&
        t.Status !== "Cancelled",
    );

    if (activeTask) {
      await sp.web.lists
        .getByTitle("Tasks")
        .items.getById(activeTask.Id)
        .update({ Status: "Removed", Comments: `Removed by ${removedByName}: ${reason}` });
    }

    await sp.web.lists.getByTitle("CR Participants").items.getById(id).update({
      Status: "Removed",
    });
  } catch (error) {
    console.error("Error deleting participant:", error);
    throw error;
  }
};

const cancelParticipantTask = async (
  participantId: number,
  changeRequestId: number,
  personId: number,
  reason: string,
): Promise<void> => {
  try {
    const sp = PnPSetup.getSP();

    const tasks = await sp.web.lists
      .getByTitle("Tasks")
      .items.select("Id", "Status", "AssignedTo/Id")
      .expand("AssignedTo")
      .filter(
        `ChangeRequestId eq ${changeRequestId} and TaskType eq 'Participant Task'`,
      )();

    const activeTask = tasks.find(
      (t) =>
        t.AssignedTo?.Id === personId &&
        t.Status !== "Complete" &&
        t.Status !== "Cancelled",
    );

    if (activeTask) {
      await sp.web.lists
        .getByTitle("Tasks")
        .items.getById(activeTask.Id)
        .update({ Status: "Cancelled", Comments: reason });
    }

    await sp.web.lists
      .getByTitle("CR Participants")
      .items.getById(participantId)
      .update({
        Status: "Not Started",
        StartDate: null,
        CompletedDate: null,
      });
  } catch (error) {
    console.error("Error cancelling participant task:", error);
    throw error;
  }
};

const getParticipantsByChangeRequestId = async (changeRequestId: number) => {
  const sp = PnPSetup.getSP();
  const rows = await sp.web.lists
    .getByTitle("CR Participants")
    .items.select(
      "Id",
      "PersonId",
      "Role",
      "Status",
      "DueDate",
      "StartDate",
      "CompletedDate",
      "Instructions",
      "Notes",
    )
    .filter(`ChangeRequestId eq ${changeRequestId} and Status ne 'Removed'`)
    .top(50)();

  const enriched = await Promise.all(
    rows.map(async (row: any) => {
      const user = await sp.web
        .getUserById(row.PersonId)
        .select("Id", "Title")();
      return {
        Id: row.Id,
        PersonId: row.PersonId,
        Role: row.Role,
        Status: row.Status ?? "Not Started",
        DueDate: row.DueDate,
        StartDate: row.StartDate,
        CompletedDate: row.CompletedDate,
        Instructions: row.Instructions,
        Notes: row.Notes,
        Person: { Id: row.PersonId, Title: user.Title, EMail: "" },
      };
    }),
  );

  return enriched;
};

const getParticipantByTaskContext = async (
  changeRequestId: number,
  userId: number,
): Promise<{ Id: number; Notes?: string; Role?: string; Instructions?: string } | null> => {
  const sp = PnPSetup.getSP();

  const results = await sp.web.lists
    .getByTitle("CR Participants")
    .items.select("Id", "Notes", "Role", "Instructions")
    .filter(`ChangeRequestId eq ${changeRequestId} and PersonId eq ${userId} and Status ne 'Removed'`)
    .orderBy("Id", false)
    .top(1)();

  return results[0] ?? null;
};


const getParticipantTaskByContext = async (
  changeRequestId: number,
  userId: number,
  role: "Contributor" | "Reviewer",
): Promise<{ Id: number; Comments?: string } | null> => {
  const sp = PnPSetup.getSP();

  // Both contributors and reviewers use "Participant Task" as TaskType
  const tasks = await sp.web.lists
    .getByTitle("Tasks")
    .items.select("Id", "Comments", "TaskType", "AssignedTo/Id")
    .expand("AssignedTo")
    .filter(
      `ChangeRequestId eq ${changeRequestId} and AssignedTo/Id eq ${userId} and TaskType eq 'Participant Task'`,
    )
    .top(5)();

  return tasks.length > 0
    ? (tasks[0] as { Id: number; Comments?: string })
    : null;
};

const getDraftDocumentFolderByChangeRequestId = async (
  changeRequestId: number,
): Promise<string | null> => {
  try {
    const sp = PnPSetup.getSP();
    const results = await sp.web.lists
      .getByTitle("Draft Documents")
      .items.select(
        "Id",
        "Title",
        "FileRef",
        "FileLeafRef",
        "ChangeRequest/Id",
        "ChangeRequest/Title",
      )
      .expand("ChangeRequest")
      .filter(`ChangeRequest/Id eq ${changeRequestId}`)();

    return results[0]?.FileRef ?? null;
  } catch (error) {
    console.error("Error fetching draft document folder:", error);
    return null;
  }
};

const getMinorChangesByDocument = async (
  documentId: number,
): Promise<MinorChange[]> => {
  try {
    const sp = PnPSetup.getSP();
    const allItems = await sp.web.lists
      .getByTitle("Minor Changes Register")
      .items.select(
        "Id",
        "Title",
        "ScopeOfChange",
        "Status",
        "Created",
        "Notes",
        "ImplementedInCR",
        "TargetDocumentId",
        "RequestedBy/Id",
        "RequestedBy/Title",
      )
      .expand("RequestedBy")
      .orderBy("Created", false)();

    return (allItems as MinorChange[]).filter(
      (item) => item.TargetDocumentId === documentId,
    );
  } catch (error) {
    console.error("Error fetching minor changes:", error);
    return [];
  }
};

const updateMinorChange = async (
  id: number,
  data: Record<string, unknown>,
): Promise<void> => {
  try {
    const sp = PnPSetup.getSP();
    await sp.web.lists
      .getByTitle("Minor Changes Register")
      .items.getById(id)
      .update(data);
  } catch (error) {
    console.error("Error updating minor change:", error);
  }
};

const getDraftFolderFiles = async (
  folderUrl: string,
): Promise<
  { Name: string; ServerRelativeUrl: string; TimeLastModified: string }[]
> => {
  try {
    const sp = PnPSetup.getSP();
    const normalizedFolderUrl = await ensureServerRelativePath(folderUrl);
    const files = await sp.web
      .getFolderByServerRelativePath(normalizedFolderUrl)
      .files.select("Name", "ServerRelativeUrl", "TimeLastModified")
      .orderBy("TimeLastModified", false)();
    return files;
  } catch (error) {
    console.error("Error fetching draft folder files:", error);
    return [];
  }
};

const uploadFilesToDraftFolder = async (
  folderUrl: string,
  files: File[],
): Promise<void> => {
  if (!folderUrl) {
    throw new Error("Draft folder URL is required.");
  }

  if (!files || files.length === 0) {
    return;
  }

  try {
    const sp = PnPSetup.getSP();
    const normalizedFolderUrl = await ensureServerRelativePath(folderUrl);
    const folder = sp.web.getFolderByServerRelativePath(normalizedFolderUrl);

    for (let i = 0; i < files.length; i++) {
      const file = files[i];
      const buffer = await file.arrayBuffer();
      await folder.files.addUsingPath(file.name, buffer, { Overwrite: true });
    }
  } catch (error) {
    console.error("Error uploading files to draft folder:", error);
    throw error;
  }
};

const getSystemConfigValue = async (key: string): Promise<boolean> => {
  try {
    const sp = PnPSetup.getSP();
    const items = await sp.web.lists
      .getByTitle("System Configuration")
      .items.select("Title", "Active")
      .filter(`Title eq '${key}'`)
      .top(1)();
    return items.length > 0 ? !!items[0].Active : false;
  } catch (error) {
    console.error(`Error fetching system config '${key}':`, error);
    return false;
  }
};

/**
 * Fetches a single global system role (e.g. "Compliance Authority",
 * "Document Controller", "Managing Director") from the System Roles list.
 * Returns the assigned person, or null if not found.
 */
const getSystemRole = async (
  roleTitle: string,
): Promise<SharePointPerson | null> => {
  try {
    const sp = PnPSetup.getSP();
    const items = await sp.web.lists
      .getByTitle("System Roles")
      .items.select("Title", "Person/Id", "Person/Title", "Person/EMail")
      .expand("Person")
      .filter(`Title eq '${roleTitle}'`)
      .top(1)();
    return items.length > 0 ? (items[0].Person as SharePointPerson) : null;
  } catch (error) {
    console.error(`Error fetching system role '${roleTitle}':`, error);
    return null;
  }
};

const getAuditTasksByCRId = async (
  changeRequestId: number,
): Promise<(Task & { CompletedDate?: string })[]> => {
  try {
    const sp = PnPSetup.getSP();
    const tasks = await sp.web.lists
      .getByTitle("Tasks")
      .items.select(
        "Id",
        "Title",
        "TaskType",
        "Status",
        "Comments",
        "Created",
        "Modified", // ← replaces CompletedDate
        "ChangeRequestId",
        "ParentTaskId",
        "AssignedTo/Id",
        "AssignedTo/Title",
        "AssignedTo/EMail",
      )
      .expand("AssignedTo")
      .filter(`ChangeRequestId eq ${changeRequestId}`)
      .orderBy("Created", true)
      .top(200)();

    // Map Modified → CompletedDate so the tab component needs no changes
    return tasks.map((t: any) => ({
      ...t,
      CompletedDate: t.Modified,
    }));
  } catch (error) {
    console.error("Error fetching audit tasks:", error);
    throw error;
  }
};

export default {
  getDepartments,
  createChangeRequest,
  getDocuments,
  getDocumentById,
  searchDocumentContent,
  uploadAttachments,
  getChangeRequests,
  getChangeRequestById,
  updateChangeRequest,
  getCurrentUser,
  searchUsers,
  getTasks,
  getPendingReadAcknowledgements,
  acknowledgeDocument,
  getTaskById,
  updateTask,
  getAttachments,
  getParticipants,
  createParticipant,
  addParticipant,
  updateParticipant,
  deleteParticipant,
  cancelParticipantTask,
  getParticipantsByChangeRequestId,
  getParticipantByTaskContext,
  getParticipantTaskByContext,
  getDraftDocumentFolderByChangeRequestId,
  getMinorChangesByDocument,
  updateMinorChange,
  getDraftFolderFiles,
  uploadFilesToDraftFolder,
  ensureServerRelativePath,
  getSystemConfigValue,
  getSystemRole,
  getAuditTasksByCRId,
  getAudienceGroups,
};
