// src/webparts/taskHub/components/tabs/AuditTrailTab.tsx
import React, { useEffect, useState } from "react";
import Box from "@mui/material/Box";
import Typography from "@mui/material/Typography";
import CircularProgress from "@mui/material/CircularProgress";
import Tooltip from "@mui/material/Tooltip";
import { IChangeRequest } from "../../../../shared/types/ChangeRequest";
import { Task } from "../../../../shared/types/Task";
import SharePointService from "../../../../shared/services/SharePointService";


// ─── Props ────────────────────────────────────────────────────────────────────
import { getAvatarInitials, getAvatarColor } from "../../../../shared/utils/avatarUtils";

interface AuditTrailTabProps {
  cr: IChangeRequest;
}

// ─── Task type → human-readable role label ────────────────────────────────────

const TASK_TYPE_LABELS: Record<string, string> = {
  "CA Review":                     "Change Authority Review",
  "Change Authority Review":       "Change Authority Review",
  "Change Authority Approval":     "Change Authority Approval",
  "Compliance Authority Review":   "Compliance Authority Review",
  "Document Controller Review":    "Document Controller Review",
  "Document Change Process":       "Document Change Process",
  "Author Review":                 "Author Review",
  "Document Review":               "Document Review",
  "Participant Task":              "Participant",
  "Publish Document":              "Publish Document",
  "Publishing Rejection Review":   "Publishing Rejection Review",
  "CR Completion":                 "CR Completion",
  "CR Info Required":              "Info Requested",
  "Final Approval":                "Final Approval",
};

// ─── Status → colour ──────────────────────────────────────────────────────────

const STATUS_CONFIG: Record<string, { bg: string; color: string; dot: string }> = {
  Approved:                      { bg: "#DFF6DD", color: "#107C10", dot: "#107C10" },
  Complete:                      { bg: "#DFF6DD", color: "#107C10", dot: "#107C10" },
  Rejected:                      { bg: "#FDE7E9", color: "#A4262C", dot: "#D13438" },
  "Marked as Minor Change":      { bg: "#FFF4CE", color: "#835B00", dot: "#FFB900" },
  "Marked for Document Obsoletion": { bg: "#FDE7E9", color: "#A4262C", dot: "#D13438" },
  Reassigned:                    { bg: "#EFF6FC", color: "#0078D4", dot: "#0078D4" },
  Cancelled:                     { bg: "#F3F2F1", color: "#A19F9D", dot: "#C8C6C4" },
  Removed:                       { bg: "#FDE7E9", color: "#A4262C", dot: "#D13438" },
  "Needs more info":             { bg: "#FFF4CE", color: "#835B00", dot: "#FFB900" },
  Pending:                       { bg: "#F3F2F1", color: "#605E5C", dot: "#C8C6C4" },
  "In Progress":                 { bg: "#EFF6FC", color: "#0078D4", dot: "#0078D4" },
  "On Hold":                     { bg: "#FFF4CE", color: "#835B00", dot: "#FFB900" },
};

// ─── Helpers ──────────────────────────────────────────────────────────────────

const formatDateTime = (date: Date | string | undefined): string => {
  if (!date) return "—";
  const d = new Date(date);
  return d.toLocaleDateString("en-AU", {
    day: "numeric",
    month: "short",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit",
    hour12: true,
  });
};

const formatDateOnly = (date: Date | string | undefined): string => {
  if (!date) return "—";
  const d = new Date(date);
  return d.toLocaleDateString("en-AU", {
    day: "numeric",
    month: "short",
    year: "numeric",
  });
};

// Is this task completed/resolved (i.e. not still active)?
const isResolved = (status: string): boolean =>
  ["Approved", "Complete", "Rejected", "Marked as Minor Change",
  "Marked for Document Obsoletion", "Reassigned", "Cancelled", "Removed"].includes(status);

// ─── Avatar ───────────────────────────────────────────────────────────────────

const Avatar = ({ name, size = 28 }: { name: string; size?: number }) => (
  <Box
    sx={{
      width: size,
      height: size,
      borderRadius: "50%",
      backgroundColor: getAvatarColor(name),
      display: "flex",
      alignItems: "center",
      justifyContent: "center",
      flexShrink: 0,
    }}
  >
    <Typography sx={{ fontSize: size * 0.36, fontWeight: 700, color: "#fff", lineHeight: 1 }}>
   {getAvatarInitials(name)}
    </Typography>
  </Box>
);

// ─── Status chip ──────────────────────────────────────────────────────────────

const StatusChip = ({ status }: { status: string }) => {
  const cfg = STATUS_CONFIG[status] ?? { bg: "#F3F2F1", color: "#605E5C", dot: "#C8C6C4" };
  return (
    <Box
      sx={{
        display: "inline-flex",
        alignItems: "center",
        gap: 0.5,
        px: 1,
        py: 0.25,
        borderRadius: "10px",
        backgroundColor: cfg.bg,
      }}
    >
      <Box sx={{ width: 6, height: 6, borderRadius: "50%", backgroundColor: cfg.dot, flexShrink: 0 }} />
      <Typography sx={{ fontSize: 11, fontWeight: 600, color: cfg.color, lineHeight: 1.4 }}>
        {status}
      </Typography>
    </Box>
  );
};

// ─── Timeline dot (left rail) ─────────────────────────────────────────────────

const TimelineDot = ({ resolved, isFirst }: { resolved: boolean; isFirst: boolean }) => (
  <Box
    sx={{
      width: 10,
      height: 10,
      borderRadius: "50%",
      backgroundColor: isFirst ? "#0078D4" : resolved ? "#107C10" : "#C8C6C4",
      border: `2px solid ${isFirst ? "#BFE0FF" : resolved ? "#C8E6C9" : "#E1DFDD"}`,
      flexShrink: 0,
      mt: "5px",
    }}
  />
);

// ─── Single audit entry ───────────────────────────────────────────────────────

interface AuditEntryProps {
  task: Task & { CompletedDate?: string };
  isFirst: boolean;
  isLast: boolean;
}

const AuditEntry = ({ task, isFirst, isLast }: AuditEntryProps) => {
  const resolved = isResolved(task.Status);
  const isNegative = ["Cancelled", "Rejected", "Removed", "Marked for Document Obsoletion"].includes(task.Status);
  const assigneeName = task.AssignedTo?.Title ?? "Unassigned";
  const roleLabel = TASK_TYPE_LABELS[task.TaskType] ?? task.TaskType;

  // Timestamp to show: completed date if resolved, otherwise created date
  const primaryDate = resolved && task.CompletedDate
    ? task.CompletedDate
    : task.Created;
  const primaryLabel = resolved && task.CompletedDate ? "Completed" : "Assigned";

  return (
    <Box sx={{ display: "flex", gap: 1.5 }}>
      {/* Rail */}
      <Box sx={{ display: "flex", flexDirection: "column", alignItems: "center", pt: 0.5 }}>
        <TimelineDot resolved={resolved} isFirst={isFirst} />
        {!isLast && (
          <Box
            sx={{
              width: 2,
              flex: 1,
              minHeight: 20,
              backgroundColor: resolved ? "#C8E6C9" : "#E1DFDD",
              my: 0.5,
              borderRadius: 1,
            }}
          />
        )}
      </Box>

      {/* Card */}
      <Box
        sx={{
          flex: 1,
          pb: isLast ? 0 : 2,
          borderRadius: "6px",
        }}
      >
        {/* Top row: role label + status chip */}
        <Box sx={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 1, flexWrap: "wrap" }}>
          <Typography sx={{ fontSize: 12, fontWeight: 700, color: "#323130", letterSpacing: 0.2 }}>
            {roleLabel}
          </Typography>
          <StatusChip status={task.Status} />
        </Box>

        {/* Assignee */}
        <Box sx={{ display: "flex", alignItems: "center", gap: 1, mt: 0.75 }}>
          <Avatar name={assigneeName} size={22} />
          <Typography sx={{ fontSize: 12, color: "#323130", fontWeight: 500 }}>
            {assigneeName}
          </Typography>
        </Box>

        {/* Timestamps */}
        <Box sx={{ display: "flex", alignItems: "center", gap: 2, mt: 0.5, flexWrap: "wrap" }}>
          <Tooltip title={`Assigned: ${formatDateTime(task.Created)}`} arrow>
            <Typography sx={{ fontSize: 11, color: "#A19F9D", cursor: "default" }}>
              Assigned {formatDateOnly(task.Created)}
            </Typography>
          </Tooltip>
          {resolved && task.CompletedDate && (
            <Tooltip title={`Completed: ${formatDateTime(task.CompletedDate)}`} arrow>
              <Typography sx={{ fontSize: 11, color: resolved ? "#107C10" : "#A19F9D", fontWeight: 500, cursor: "default" }}>
                → Completed {formatDateOnly(task.CompletedDate)}
              </Typography>
            </Tooltip>
          )}
          {!resolved && (
            <Typography sx={{ fontSize: 11, color: "#0078D4", fontStyle: "italic" }}>
              In progress
            </Typography>
          )}
        </Box>

        {/* Comments */}
        {task.Comments && task.Comments.trim() && (
          <Box
            sx={{
              mt: 1,
              px: 1.5,
              py: 1,
              backgroundColor: isNegative ? "#FDE7E9" : "#F8F7F6",
              border: `1px solid ${isNegative ? "#F1B0B7" : "#EDEBE9"}`,
              borderLeft: `3px solid ${isNegative ? "#D13438" : "#0078D4"}`,
              borderRadius: "0 4px 4px 0",
            }}
          >
            <Typography sx={{ fontSize: 12, color: isNegative ? "#A4262C" : "#323130", lineHeight: 1.6, fontStyle: "italic" }}>
              "{task.Comments}"
            </Typography>
          </Box>
        )}
      </Box>
    </Box>
  );
};

// ─── Submission entry (synthetic first event) ─────────────────────────────────

const SubmissionEntry = ({ cr, isLast }: { cr: IChangeRequest; isLast: boolean }) => (
  <Box sx={{ display: "flex", gap: 1.5 }}>
    <Box sx={{ display: "flex", flexDirection: "column", alignItems: "center", pt: 0.5 }}>
      <Box
        sx={{
          width: 10,
          height: 10,
          borderRadius: "50%",
          backgroundColor: "#0F4C81",
          border: "2px solid #C7D9ED",
          flexShrink: 0,
          mt: "5px",
        }}
      />
      {!isLast && (
        <Box sx={{ width: 2, flex: 1, minHeight: 20, backgroundColor: "#C8E6C9", my: 0.5, borderRadius: 1 }} />
      )}
    </Box>
    <Box sx={{ flex: 1, pb: isLast ? 0 : 2 }}>
      <Box sx={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 1 }}>
        <Typography sx={{ fontSize: 12, fontWeight: 700, color: "#323130" }}>
          Change Request Submitted
        </Typography>
        <Box sx={{ display: "inline-flex", alignItems: "center", gap: 0.5, px: 1, py: 0.25, borderRadius: "10px", backgroundColor: "#EFF6FC" }}>
          <Box sx={{ width: 6, height: 6, borderRadius: "50%", backgroundColor: "#0078D4", flexShrink: 0 }} />
          <Typography sx={{ fontSize: 11, fontWeight: 600, color: "#0078D4" }}>Submitted</Typography>
        </Box>
      </Box>
      {cr.Requestor && (
        <Box sx={{ display: "flex", alignItems: "center", gap: 1, mt: 0.75 }}>
          <Avatar name={cr.Requestor.Title ?? "Unknown"} size={22} />
          <Typography sx={{ fontSize: 12, color: "#323130", fontWeight: 500 }}>
            {cr.Requestor.Title}
          </Typography>
        </Box>
      )}
      <Tooltip title={formatDateTime(cr.Created)} arrow>
        <Typography sx={{ fontSize: 11, color: "#A19F9D", mt: 0.5, cursor: "default" }}>
          {formatDateOnly(cr.Created)}
        </Typography>
      </Tooltip>
    </Box>
  </Box>
);

interface ParticipantChildGroupProps {
  children: (Task & { CompletedDate?: string })[];
}

const ParticipantChildGroup = ({ children }: ParticipantChildGroupProps): React.ReactElement => {
  const [expanded, setExpanded] = React.useState(false);
  const visible = expanded ? children : children.slice(0, 2);
  const remaining = children.length - visible.length;

  return (
    <Box sx={{ ml: 4, pl: 1.5, borderLeft: "2px solid #E1DFDD", mb: 2, display: "flex", flexDirection: "column", gap: 1 }}>
      {visible.map((task) => {
        const cfg = STATUS_CONFIG[task.Status] ?? { bg: "#F3F2F1", color: "#605E5C", dot: "#C8C6C4" };
        const assigneeName = task.AssignedTo?.Title ?? "Unassigned";
        const roleLabel = TASK_TYPE_LABELS[task.TaskType] ?? task.TaskType;
        return (
          <Box key={task.Id} sx={{ backgroundColor: "#FAFAFA", borderRadius: "6px", p: 1.25 }}>
            <Box sx={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 1 }}>
              <Typography sx={{ fontSize: 11, fontWeight: 700, color: "#323130" }}>{roleLabel}</Typography>
              <Box sx={{ display: "inline-flex", alignItems: "center", gap: 0.5, px: 1, py: 0.2, borderRadius: "10px", backgroundColor: cfg.bg }}>
                <Box sx={{ width: 5, height: 5, borderRadius: "50%", backgroundColor: cfg.dot }} />
                <Typography sx={{ fontSize: 10, fontWeight: 600, color: cfg.color }}>{task.Status}</Typography>
              </Box>
            </Box>
            <Box sx={{ display: "flex", alignItems: "center", gap: 0.75, mt: 0.5 }}>
              <Avatar name={assigneeName} size={18} />
              <Typography sx={{ fontSize: 11, color: "#323130" }}>{assigneeName}</Typography>
              <Typography sx={{ fontSize: 10, color: "#A19F9D", ml: "auto" }}>{formatDateOnly(task.Created)}</Typography>
            </Box>
            {task.Comments && task.Comments.trim() && (
              <Box sx={{
                mt: 0.75, px: 1, py: 0.75,
                backgroundColor: ["Cancelled", "Rejected", "Removed", "Marked for Document Obsoletion"].includes(task.Status) ? "#FDE7E9" : "#F3F2F1",
                borderLeft: `3px solid ${["Cancelled", "Rejected", "Removed", "Marked for Document Obsoletion"].includes(task.Status) ? "#D13438" : "#0078D4"}`,
                borderRadius: "0 4px 4px 0",
              }}>
                <Typography sx={{ fontSize: 11, color: ["Cancelled", "Rejected", "Removed", "Marked for Document Obsoletion"].includes(task.Status) ? "#A4262C" : "#323130", fontStyle: "italic" }}>
                  "{task.Comments}"
                </Typography>
              </Box>
            )}
          </Box>
        );
      })}
      {remaining > 0 && (
        <Box
          component="button"
          onClick={() => setExpanded(true)}
          sx={{ fontSize: 11, color: "#605E5C", background: "none", border: "none", textAlign: "left", cursor: "pointer", py: 0.5 }}
        >
          + {remaining} more participant task{remaining === 1 ? "" : "s"}
        </Box>
      )}
    </Box>
  );
};

// ─── Main Component ───────────────────────────────────────────────────────────

const AuditTrailTab = ({ cr }: AuditTrailTabProps): React.ReactElement => {
  const [tasks, setTasks] = useState<(Task & { CompletedDate?: string })[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    if (!cr?.ID) return;
    setLoading(true);
    setError(null);

    SharePointService.getAuditTasksByCRId(cr.ID)
      .then((results) => setTasks(results))
      .catch(() => setError("Could not load audit trail."))
      .finally(() => setLoading(false));
  }, [cr?.ID]);

  if (loading) {
    return (
      <Box display="flex" justifyContent="center" alignItems="center" mt={6}>
        <CircularProgress size={24} sx={{ color: "#0078D4" }} />
      </Box>
    );
  }

  if (error) {
    return (
      <Box p={3}>
        <Typography fontSize={13} color="error">{error}</Typography>
      </Box>
    );
  }

  const topLevelTasks = tasks.filter((t) => t.TaskType !== "Participant Task");
  const childTasksByParentId = tasks
    .filter((t) => t.TaskType === "Participant Task" && t.ParentTaskId)
    .reduce((acc, t) => {
      const key = t.ParentTaskId as number;
      if (!acc[key]) acc[key] = [];
      acc[key].push(t);
      return acc;
    }, {} as Record<number, (Task & { CompletedDate?: string })[]>);

  const orphanedChildTasks = tasks.filter(
    (t) => t.TaskType === "Participant Task" && (!t.ParentTaskId || !topLevelTasks.some((p) => p.Id === t.ParentTaskId))
  );
  const allTopLevel = [...topLevelTasks, ...orphanedChildTasks].sort(
    (a, b) => new Date(a.Created).getTime() - new Date(b.Created).getTime()
  );

  const totalEntries = allTopLevel.length + Object.values(childTasksByParentId).reduce((sum, arr) => sum + arr.length, 0) + 1;

  return (
    <Box p={3} display="flex" flexDirection="column" gap={0}>
      {/* Header summary */}
      <Box
        sx={{
          display: "flex",
          alignItems: "center",
          gap: 1.5,
          px: 2,
          py: 1.25,
          backgroundColor: "#F3F2F1",
          border: "1px solid #EDEBE9",
          borderRadius: "6px",
          mb: 2.5,
        }}
      >
        <Typography sx={{ fontSize: 13, color: "#323130" }}>
          <strong>{totalEntries}</strong>
          <Typography component="span" sx={{ fontSize: 13, color: "#605E5C", ml: 0.5 }}>
            {totalEntries === 1 ? "event recorded" : "events recorded"} for this change request
          </Typography>
        </Typography>
      </Box>

      {/* Timeline */}
      <Box sx={{ pl: 0.5 }}>
        <SubmissionEntry cr={cr} isLast={allTopLevel.length === 0} />
        {allTopLevel.map((task, index) => {
          const children = childTasksByParentId[task.Id] ?? [];
          return (
            <React.Fragment key={task.Id}>
              <AuditEntry
                task={task}
                isFirst={false}
                isLast={index === allTopLevel.length - 1 && children.length === 0}
              />
              {children.length > 0 && (
                <ParticipantChildGroup>{children}</ParticipantChildGroup>
              )}
            </React.Fragment>
          );
        })}
      </Box>
    </Box>
  );
};

export default AuditTrailTab;