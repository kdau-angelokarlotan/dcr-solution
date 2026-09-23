import React, { useState, useEffect } from "react";
import Box from "@mui/material/Box";
import Typography from "@mui/material/Typography";
import Button from "@mui/material/Button";
import TextField from "@mui/material/TextField";
import Divider from "@mui/material/Divider";
import CheckCircleOutlineIcon from "@mui/icons-material/CheckCircleOutline";
import CancelOutlinedIcon from "@mui/icons-material/CancelOutlined";
import AttachFileIcon from "@mui/icons-material/AttachFile";
import { Task } from "../../../../shared/types/Task";
import { IChangeRequest } from "../../../../shared/types/ChangeRequest";
import SharePointService from "../../../../shared/services/SharePointService";
import { emitParticipantRefetch } from "../../../../shared/hooks/useParticipants";

// ─── Props ────────────────────────────────────────────────────────────────────

interface ParticipantTaskProps {
  task: Task;
  cr: IChangeRequest | null;
  onTaskComplete: () => void;
}

// ─── Main Component ───────────────────────────────────────────────────────────

const ParticipantTask = ({
  task,
  cr,
  onTaskComplete,
}: ParticipantTaskProps): React.ReactElement => {
  const [participantNotes, setParticipantNotes] = useState("");
  const [submitting, setSubmitting] = useState(false);

  const [participant, setParticipant] = useState<{
    Id: number;
    Notes: string;
    Role: string;
    Instructions: string;
  } | null>(null);
  const participantRole = participant?.Role;
  const isReviewer = participantRole === "Reviewer";
  const positiveLabel = isReviewer ? "Approve" : "Mark Complete";
  const draftUrl = cr?.DraftDocumentUrl ?? null;
  useEffect(() => {
    if (!task.AssignedTo?.Id) return; // ← Add this guard

    SharePointService.getParticipantByTaskContext(
      task.ChangeRequestId,
      task.AssignedTo.Id,
    )
      .then((resolvedParticipant) => {
        if (resolvedParticipant) {
          const resolvedRole =
            (resolvedParticipant as { Role?: string }).Role ?? "Contributor";
          const resolvedInstructions =
            (resolvedParticipant as { Instructions?: string }).Instructions ??
            "";
          setParticipant({
            Id: resolvedParticipant.Id,
            Notes: resolvedParticipant.Notes ?? "",
            Role: resolvedRole,
            Instructions: resolvedInstructions,
          });
          setParticipantNotes(resolvedParticipant.Notes ?? "");
        }
      })
      .catch((err) => {
        console.error("Failed to fetch participant for task:", err);
      });
  }, [task.ChangeRequestId, task.AssignedTo?.Id]);

  const handleAction = async (action: "positive" | "reject"): Promise<void> => {
    setSubmitting(true);
    try {
      if (!participant) {
        throw new Error("Participant context is not loaded yet.");
      }

      const currentIsReviewer = participant.Role === "Reviewer";

      await SharePointService.updateParticipant(participant.Id, {
        Status: action === "positive" ? "Complete" : "Rejected",
        CompletedDate:
          action === "positive" ? new Date().toISOString() : undefined,
        Notes: participantNotes.trim() || undefined,
      });
      emitParticipantRefetch(task.ChangeRequestId).catch(console.error);

      await SharePointService.updateTask(task.Id, {
        Status:
          action === "positive"
            ? currentIsReviewer
              ? "Approved"
              : "Complete"
            : "Rejected",
        Comments: participantNotes.trim() || undefined,
      });
      onTaskComplete();
    } catch (err) {
      console.error("Failed to update task:", err);
    } finally {
      setSubmitting(false);
    }
  };

  return (
    <Box sx={{ display: "flex", flexDirection: "column", gap: 2.5 }}>
      {/* ── Document Link ── */}
      {draftUrl ? (
        <Box
          component="a"
          href={`${draftUrl}${draftUrl.includes("?") ? "&" : "?"}web=1`}
          target="_blank"
          rel="noopener noreferrer"
          sx={{
            display: "flex",
            alignItems: "center",
            gap: 1.5,
            p: 1.5,
            backgroundColor: "#FFFFFF",
            border: "1px solid #D0E7F9",
            borderRadius: 1.5,
            textDecoration: "none",
            transition: "background-color 0.15s ease, border-color 0.15s ease",
            "&:hover": {
              backgroundColor: "#F5FAFF",
              borderColor: "#B6DBF7",
            },
          }}
        >
          <Box
            sx={{
              width: 32,
              height: 32,
              backgroundColor: "#EAF4FD",
              borderRadius: 1,
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              flexShrink: 0,
            }}
          >
            <AttachFileIcon sx={{ fontSize: 16, color: "#0078D4" }} />
          </Box>
          <Box sx={{ flex: 1, minWidth: 0 }}>
            <Typography
              sx={{
                fontSize: 13,
                fontWeight: 600,
                color: "#0078D4",
                lineHeight: 1.3,
              }}
            >
              Open draft document
            </Typography>
            <Typography sx={{ fontSize: 11, color: "#605E5C" }}>
              Click to open in SharePoint
            </Typography>
          </Box>
          <Box sx={{ fontSize: 11, color: "#0078D4", fontWeight: 600 }}>↗</Box>
        </Box>
      ) : (
        <Box
          sx={{
            p: 1.5,
            backgroundColor: "#F8F8F8",
            border: "1px dashed #D2D0CE",
            borderRadius: 1.5,
          }}
        >
          <Typography
            sx={{ fontSize: 13, color: "#A19F9D", fontStyle: "italic" }}
          >
            No document available yet
          </Typography>
        </Box>
      )}

      <Divider />

      {/* ── Instructions ── */}
      {participant?.Instructions && (
        <Box
          sx={{
            p: 1.5,
            backgroundColor: "#EFF6FC",
            borderRadius: 1,
            border: "1px solid #C7E0F4",
          }}
        >
          <Typography
            sx={{
              fontSize: 10,
              fontWeight: 700,
              color: "#005A9E",
              letterSpacing: 0.5,
              textTransform: "uppercase",
              mb: 0.5,
            }}
          >
            Instructions
          </Typography>
          <Typography sx={{ fontSize: 12, color: "#004578", whiteSpace: "pre-wrap" }}>
            {participant.Instructions}
          </Typography>
        </Box>
      )}

      {/* ── Participant Notes ── */}
      <Box>
        <Typography
          sx={{ fontSize: 12, fontWeight: 600, color: "#323130", mb: 1 }}
        >
          Your notes
        </Typography>
        <TextField
          multiline
          rows={4}
          fullWidth
          placeholder="Add any notes for the author..."
          value={participantNotes}
          onChange={(e) => setParticipantNotes(e.target.value)}
          disabled={submitting}
          sx={{
            "& .MuiOutlinedInput-root": {
              fontSize: 13,
              "& fieldset": { borderColor: "#EDEBE9" },
              "&:hover fieldset": { borderColor: "#0078D4" },
              "&.Mui-focused fieldset": { borderColor: "#0078D4" },
            },
          }}
        />
      </Box>

      {/* ── Actions ── */}
      <Box sx={{ display: "flex", gap: 1.5 }}>
        <Button
          variant="contained"
          disableElevation
          disabled={submitting || !participant}
          onClick={() => handleAction("positive")}
          startIcon={<CheckCircleOutlineIcon sx={{ fontSize: 16 }} />}
          sx={{ textTransform: "none", fontWeight: 600, fontSize: 13 }}
        >
          {submitting ? "Saving..." : positiveLabel}
        </Button>
        <Button
          variant="outlined"
          color="error"
          disabled={submitting || !participant}
          onClick={() => handleAction("reject")}
          startIcon={<CancelOutlinedIcon sx={{ fontSize: 16 }} />}
          sx={{ textTransform: "none", fontWeight: 600, fontSize: 13 }}
        >
          Reject
        </Button>
      </Box>
    </Box>
  );
};

export default ParticipantTask;
