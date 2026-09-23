  import React, { useState, useEffect, useCallback } from "react";
  import Box from "@mui/material/Box";
  import Typography from "@mui/material/Typography";
  import Button from "@mui/material/Button";
  import Dialog from "@mui/material/Dialog";
  import DialogTitle from "@mui/material/DialogTitle";
  import DialogContent from "@mui/material/DialogContent";
  import DialogActions from "@mui/material/DialogActions";
  import IconButton from "@mui/material/IconButton";
  import TextField from "@mui/material/TextField";
  import Checkbox from "@mui/material/Checkbox";
  import Radio from "@mui/material/Radio";
  import CircularProgress from "@mui/material/CircularProgress";
  import CloseIcon from "@mui/icons-material/Close";
  import PeopleIcon from "@mui/icons-material/People";
  import WarningAmberIcon from "@mui/icons-material/WarningAmber";
  import CheckCircleOutlineIcon from "@mui/icons-material/CheckCircleOutline";
  import EditNoteIcon from "@mui/icons-material/EditNote";
  import OpenInNewIcon from "@mui/icons-material/OpenInNew";
  import FolderOpenIcon from "@mui/icons-material/FolderOpen";
  import UploadFileIcon from "@mui/icons-material/UploadFile";
  import RefreshIcon from "@mui/icons-material/Refresh";
  import InsertDriveFileOutlinedIcon from "@mui/icons-material/InsertDriveFileOutlined";
  import { IChangeRequest } from "../../../../shared/types/ChangeRequest";
  import { SharePointPerson } from "../../../../shared/types/SharePointPerson";
  import { Task } from "../../../../shared/types/Task";
  import { MinorChange } from "../../../../shared/types/MinorChange";
  import { useParticipants } from "../../../../shared/hooks/useParticipants";
  import SharePointService from "../../../../shared/services/SharePointService";
  import ParticipantsTable from "../ParticipantsTable";

  interface DocumentChangeProcessTaskProps {
    task: Task;
    cr: IChangeRequest;
    currentUser: SharePointPerson;
    onTaskComplete: () => void;
    onCRUpdate?: () => void;
  }

  const getInitials = (name: string): string =>
    name
      .split(" ")
      .map((n) => n[0])
      .slice(0, 2)
      .join("")
      .toUpperCase();

  const formatDate = (date: Date | string | undefined): string => {
    if (!date) return "—";
    const d = typeof date === "string" ? new Date(date) : date;
    return d.toLocaleDateString("en-AU", {
      day: "numeric",
      month: "short",
      year: "numeric",
    });
  };

  const DetailField = ({
    label,
    value,
  }: {
    label: string;
    value: React.ReactNode;
  }): React.ReactElement => (
    <Box
      sx={{ display: "grid", gridTemplateColumns: "120px 1fr", py: 0.75, gap: 1 }}
    >
      <Typography sx={{ fontSize: 12, fontWeight: 600, color: "#605E5C" }}>
        {label}
      </Typography>
      <Typography sx={{ fontSize: 13, color: "#323130", lineHeight: 1.5 }}>
        {value || <span style={{ color: "#A19F9D" }}>—</span>}
      </Typography>
    </Box>
  );

  // ─── URL helpers (regex-only, no new URL — SP URLs may have unencoded spaces) ──

  /** Strip origin from a full URL to get a server-relative path. */
  const toServerRelativePath = (url: string): string => {
    const m = url.trim().match(/^https?:\/\/[^/]+(\/.*)/i);
    return m ? m[1] : url.startsWith("/") ? url : `/${url}`;
  };

  /** Build a SharePoint library AllItems.aspx link that opens the folder. */
  const toFolderAllItemsUrl = (folderUrl: string): string => {
    const origin = window.location.origin;
    const serverRelative = toServerRelativePath(folderUrl);
    // Library root is the first 3 segments for /sites/ paths: /sites/{site}/{library}
    const segments = serverRelative.split("/").filter(Boolean);
    const libraryRoot = `/${segments.slice(0, 3).join("/")}`;
    return `${origin}${libraryRoot}/Forms/AllItems.aspx?id=${encodeURIComponent(serverRelative)}`;
  };

  /** Convert a server-relative path to an absolute SharePoint URL. */
  const toAbsoluteSharePointUrl = (serverRelativePath: string): string =>
    `${window.location.origin}${serverRelativePath}`;

  const EXT_COLORS: Record<string, string> = {
    docx: "#2B579A",
    doc: "#2B579A",
    xlsx: "#217346",
    xls: "#217346",
    pptx: "#D24726",
    ppt: "#D24726",
    pdf: "#D32F2F",
  };

  const getExtColor = (name: string): string => {
    const ext = name.split(".").pop()?.toLowerCase() ?? "";
    return EXT_COLORS[ext] ?? "#605E5C";
  };

  // ─── Main Component ───────────────────────────────────────────────────────────

  const DocumentChangeProcessTask = ({
    task,
    cr,
    currentUser,
    onTaskComplete,
    onCRUpdate,
  }: DocumentChangeProcessTaskProps): React.ReactElement => {
    const { contributors, reviewers, loading, refetch } = useParticipants(cr.ID);
    const [manageOpen, setManageOpen] = useState(false);
    const [rejectOpen, setRejectOpen] = useState(false);
    const [rejectReason, setRejectReason] = useState("");
    const [submitting, setSubmitting] = useState(false);
    const [refreshing, setRefreshing] = useState(false);

    const [minorChanges, setMinorChanges] = useState<MinorChange[]>([]);
    const [minorChangesLoading, setMinorChangesLoading] = useState(false);
    const [savingMinorId, setSavingMinorId] = useState<number | null>(null);
    const [selectedMinorChange, setSelectedMinorChange] =
      useState<MinorChange | null>(null);

    // Draft documents state
    const [draftFolderPath, setDraftFolderPath] = useState<string | null>(null);
    const [draftFiles, setDraftFiles] = useState<
      { Name: string; ServerRelativeUrl: string; TimeLastModified: string }[]
    >([]);
    const [draftFilesLoading, setDraftFilesLoading] = useState(false);
    const [uploading, setUploading] = useState(false);
    const [settingActive, setSettingActive] = useState(false);
    const [tasks, setTasks] = useState<Task[]>([]);

    const isAuthor = cr.Author0?.Id === currentUser.Id;
    const isCa = cr.ChangeAuthority?.Id === currentUser.Id;
    const isSelfCa = isAuthor && isCa;
    const isExistingDocument = !cr.NewDocument && !!cr.TargetDocumentId;
    const authorComplete = isSelfCa
      ? true
      : tasks.some(
          (t) =>
            t.TaskType === "Document Change Process" &&
            t.Status === "Author Approved",
        );

    const allParticipants = [...contributors, ...reviewers];
    const total = allParticipants.length;
    const complete = allParticipants.filter(
      (p) => p.Status === "Complete",
    ).length;
    const inProgress = allParticipants.filter(
      (p) => p.Status === "In Progress",
    ).length;
    const progressPercent = total > 0 ? Math.round((complete / total) * 100) : 0;
    const allParticipantsComplete =
      total > 0 && allParticipants.every((p) => p.Status === "Complete");

    const contributorsComplete = contributors.filter(
      (p) => p.Status === "Complete",
    ).length;
    const reviewersComplete = reviewers.filter(
      (p) => p.Status === "Complete",
    ).length;

    const pendingMinorChanges = minorChanges.filter(
      (mc) => mc.Status !== "Cancelled",
    );
    const allMinorChangesActioned =
      pendingMinorChanges.length === 0 ||
      pendingMinorChanges.every((mc) => mc.Status === "Implemented");

    const hasActiveDocument = !!cr.DraftDocumentUrl;
    const canSubmit =
      allParticipantsComplete && allMinorChangesActioned && hasActiveDocument;

    useEffect(() => {
      if (!isCa) return;

      SharePointService.getAuditTasksByCRId(cr.ID)
        .then((items) => setTasks(items))
        .catch(console.error);
    }, [cr.ID, isCa]);

    // Resolve draft folder: try library lookup first, fall back to cr.DraftFolderUrl
    // ensureServerRelativePath in the service handles site-relative → full server-relative
    useEffect(() => {
      const resolve = async (): Promise<void> => {
        try {
          const path =
            await SharePointService.getDraftDocumentFolderByChangeRequestId(
              cr.ID,
            );
          if (path) {
            setDraftFolderPath(path);
          } else if (cr.DraftFolderUrl) {
            setDraftFolderPath(
              await SharePointService.ensureServerRelativePath(cr.DraftFolderUrl),
            );
          } else {
            setDraftFolderPath(null);
          }
        } catch {
          if (cr.DraftFolderUrl) {
            try {
              setDraftFolderPath(
                await SharePointService.ensureServerRelativePath(
                  cr.DraftFolderUrl,
                ),
              );
            } catch {
              setDraftFolderPath(null);
            }
          } else {
            setDraftFolderPath(null);
          }
        }
      };
      resolve().catch(console.error);
    }, [cr.ID, cr.DraftFolderUrl]);

    // Fetch draft files once we have the folder path
    const fetchDraftFiles = useCallback(async () => {
      if (!draftFolderPath) return;
      setDraftFilesLoading(true);
      try {
        const files =
          await SharePointService.getDraftFolderFiles(draftFolderPath);
        setDraftFiles(files);
      } catch (err) {
        console.error("Failed to fetch draft files:", err);
      } finally {
        setDraftFilesLoading(false);
      }
    }, [draftFolderPath]);

    const handleRefreshAll = async (): Promise<void> => {
      setRefreshing(true);
      try {
        const refreshPromises: Promise<unknown>[] = [fetchDraftFiles(), refetch()];

        if (isExistingDocument && cr.TargetDocumentId) {
          refreshPromises.push(
            SharePointService.getMinorChangesByDocument(cr.TargetDocumentId).then(
              setMinorChanges,
            ),
          );
        }

        if (isCa) {
          refreshPromises.push(SharePointService.getAuditTasksByCRId(cr.ID).then(setTasks));
        }

        await Promise.all(refreshPromises);
      } catch (err) {
        console.error("Failed to refresh all data:", err);
      } finally {
        setRefreshing(false);
      }
    };

    useEffect(() => {
      fetchDraftFiles().catch(console.error);
    }, [fetchDraftFiles]);

    const handleUpload = async (
      e: React.ChangeEvent<HTMLInputElement>,
    ): Promise<void> => {
      const input = e.target;
      const fileList = input.files;
      if (!fileList || fileList.length === 0 || !draftFolderPath) return;
      setUploading(true);
      try {
        await SharePointService.uploadFilesToDraftFolder(
          draftFolderPath,
          Array.from(fileList),
        );
        await fetchDraftFiles();
      } catch (err) {
        console.error("Upload failed:", err);
      } finally {
        setUploading(false);
        input.value = "";
      }
    };

    const handleSetActiveDocument = async (
      serverRelativeUrl: string,
    ): Promise<void> => {
      setSettingActive(true);
      try {
        // Extract filename from the server relative URL
        const fileName = serverRelativeUrl.split("/").pop() ?? "";
        // Strip extension for a clean document name
        const documentName = fileName.replace(/\.[^/.]+$/, "");

        await SharePointService.updateChangeRequest(cr.ID, {
          DraftDocumentUrl: toAbsoluteSharePointUrl(serverRelativeUrl),
          DraftDocumentName: documentName,
        });
        onCRUpdate?.();
      } catch (err) {
        console.error("Failed to set active document:", err);
      } finally {
        setSettingActive(false);
      }
    };
    
    useEffect(() => {
      if (!isExistingDocument || !cr.TargetDocumentId) return;
      setMinorChangesLoading(true);
      SharePointService.getMinorChangesByDocument(cr.TargetDocumentId)
        .then(setMinorChanges)
        .catch(console.error)
        .finally(() => setMinorChangesLoading(false));
    }, [cr.TargetDocumentId, isExistingDocument]);

    const handleToggleMinorChange = async (
      item: MinorChange,
      e?: React.MouseEvent,
    ): Promise<void> => {
      if (e) e.stopPropagation();
      const isChecked = item.Status === "Implemented";
      setSavingMinorId(item.Id);
      try {
        const newStatus = isChecked ? "Pending" : "Implemented";
        await SharePointService.updateMinorChange(item.Id, {
          Status: newStatus,
          DateImplemented: isChecked ? null : new Date().toISOString(),
        });
        setMinorChanges((prev) =>
          prev.map((mc) =>
            mc.Id === item.Id
              ? { ...mc, Status: newStatus as "Pending" | "Implemented" }
              : mc,
          ),
        );
        if (selectedMinorChange?.Id === item.Id) {
          setSelectedMinorChange((prev) =>
            prev
              ? { ...prev, Status: newStatus as "Pending" | "Implemented" }
              : null,
          );
        }
      } catch (err) {
        console.error("Failed to update minor change:", err);
      } finally {
        setSavingMinorId(null);
      }
    };

  const handleSubmitToPublishing = async (): Promise<void> => {
    setSubmitting(true);
    try {
      if (isCa) {
        // CA clicks "Submit to publishing" - close their task and advance the CR
        await SharePointService.updateTask(task.Id, { Status: "Complete" });
        await SharePointService.updateChangeRequest(cr.ID, {
          Status: "Document Review",
        });
      } else {
        // Author clicks "Release candidate ready" - only mark their task as Author Approved
        // Do NOT advance the CR - CA must still submit separately
        await SharePointService.updateTask(task.Id, { Status: "Author Approved" });
      }
      onTaskComplete();
    } catch (err) {
      console.error("Failed to submit to publishing:", err);
    } finally {
      setSubmitting(false);
    }
  };

    const handleReject = async (): Promise<void> => {
      setSubmitting(true);
      try {
        await SharePointService.updateTask(task.Id, { Status: "Rejected" });
        await SharePointService.updateChangeRequest(cr.ID, {
          Status: "Rejected",
          RejectionReason: rejectReason,
        });
        setRejectOpen(false);
        onTaskComplete();
      } catch (err) {
        console.error("Failed to reject:", err);
      } finally {
        setSubmitting(false);
      }
    };

    // ── Build blocking reasons for the single warning ──
    const blockingReasons: string[] = [];
    if (!allParticipantsComplete)
      blockingReasons.push("participants to complete");
    if (!allMinorChangesActioned)
      blockingReasons.push("minor changes to be actioned");
    if (!hasActiveDocument)
      blockingReasons.push("an active document to be selected");
    if (isCa && !authorComplete)
      blockingReasons.push("author to complete their section");

    return (
      <Box sx={{ display: "flex", flexDirection: "column", gap: 2, p: 2.5 }}>
        <Box sx={{ display: "flex", justifyContent: "flex-end" }}>
          <IconButton
            onClick={handleRefreshAll}
            disabled={refreshing}
            sx={{ color: "#605E5C" }}
          >
            {refreshing ? (
              <CircularProgress size={16} />
            ) : (
              <RefreshIcon sx={{ fontSize: 18 }} />
            )}
          </IconButton>
        </Box>

        {/* ── Draft Documents Card ── */}
        <Box
          sx={{
            border: "1px solid #EDEBE9",
            borderRadius: "10px",
            backgroundColor: "#fff",
            overflow: "hidden",
          }}
        >
          {/* Header */}
          <Box
            sx={{
              px: 2,
              py: 1.25,
              borderBottom: "1px solid #EDEBE9",
              backgroundColor: "#FAFAFA",
              display: "flex",
              alignItems: "center",
              justifyContent: "space-between",
            }}
          >
            <Box sx={{ display: "flex", alignItems: "center", gap: 1 }}>
              <FolderOpenIcon sx={{ fontSize: 15, color: "#605E5C" }} />
              <Typography
                sx={{ fontSize: 12, fontWeight: 600, color: "#323130" }}
              >
                Draft Documents
              </Typography>
            </Box>
            {draftFolderPath && (
              <Box
                component="a"
                href={toFolderAllItemsUrl(draftFolderPath)}
                target="_blank"
                rel="noopener noreferrer"
                sx={{
                  display: "flex",
                  alignItems: "center",
                  gap: 0.5,
                  color: "#0078D4",
                  fontSize: 12,
                  fontWeight: 500,
                  textDecoration: "none",
                  "&:hover": { textDecoration: "underline" },
                }}
              >
                <OpenInNewIcon sx={{ fontSize: 14 }} />
                Open folder
              </Box>
            )}
          </Box>

          {/* File list */}
          <Box sx={{ py: 0.5 }}>
            {draftFilesLoading ? (
              <Box display="flex" justifyContent="center" py={2}>
                <CircularProgress size={20} />
              </Box>
            ) : draftFiles.length === 0 ? (
              <Typography
                sx={{
                  fontSize: 12,
                  color: "#A19F9D",
                  fontStyle: "italic",
                  px: 2,
                  py: 1.5,
                }}
              >
                No files uploaded yet
              </Typography>
            ) : (
              draftFiles.map((file) => {
                const fileUrl = toAbsoluteSharePointUrl(file.ServerRelativeUrl);
                const isActive =
                  cr.DraftDocumentUrl === fileUrl ||
                  cr.DraftDocumentUrl === file.ServerRelativeUrl;

                return (
                  <Box
                    key={file.ServerRelativeUrl}
                    sx={{
                      display: "flex",
                      alignItems: "center",
                      gap: 1,
                      px: 2,
                      py: 0.75,
                      borderBottom: "1px solid #F3F2F1",
                      "&:last-child": { borderBottom: "none" },
                      "&:hover": { backgroundColor: "#F9F9F9" },
                    }}
                  >
                    <Radio
                      checked={isActive}
                      disabled={settingActive}
                      onChange={() =>
                        handleSetActiveDocument(file.ServerRelativeUrl)
                      }
                      size="small"
                      sx={{
                        p: 0,
                        color: "#C8C6C4",
                        "&.Mui-checked": { color: "#0078D4" },
                      }}
                    />
                    <InsertDriveFileOutlinedIcon
                      sx={{ fontSize: 16, color: getExtColor(file.Name) }}
                    />
                    <Box
                      component="a"
                      href={`${fileUrl}?web=1`}
                      target="_blank"
                      rel="noopener noreferrer"
                      sx={{
                        flex: 1,
                        fontSize: 13,
                        color: "#323130",
                        textDecoration: "none",
                        overflow: "hidden",
                        textOverflow: "ellipsis",
                        whiteSpace: "nowrap",
                        "&:hover": {
                          color: "#0078D4",
                          textDecoration: "underline",
                        },
                      }}
                    >
                      {file.Name}
                    </Box>
                    {isActive && (
                      <Typography
                        sx={{
                          fontSize: 10,
                          fontWeight: 600,
                          color: "#107C10",
                          backgroundColor: "#DFF6DD",
                          px: 0.75,
                          py: 0.15,
                          borderRadius: "10px",
                          flexShrink: 0,
                        }}
                      >
                        Active
                      </Typography>
                    )}
                  </Box>
                );
              })
            )}
          </Box>

          {/* Upload button */}
          <Box sx={{ px: 2, py: 1.5, borderTop: "1px solid #EDEBE9" }}>
            <Button
              component="label"
              variant="outlined"
              size="small"
              disabled={uploading || !draftFolderPath}
              startIcon={
                uploading ? (
                  <CircularProgress size={14} color="inherit" />
                ) : (
                  <UploadFileIcon sx={{ fontSize: 16 }} />
                )
              }
              sx={{
                textTransform: "none",
                fontSize: 12,
                fontWeight: 500,
                borderRadius: "6px",
                borderColor: "#EDEBE9",
                color: "#323130",
                "&:hover": {
                  borderColor: "#0078D4",
                  backgroundColor: "#F3F2F1",
                },
              }}
            >
              {uploading ? "Uploading..." : "Upload files"}
              <input type="file" hidden multiple onChange={handleUpload} />
            </Button>
          </Box>
        </Box>

        {/* ── Participants Card (with manage button inside) ── */}
        <Box
          sx={{
            border: "1px solid #EDEBE9",
            borderRadius: "10px",
            backgroundColor: "#fff",
            overflow: "hidden",
          }}
        >
          {/* Header with manage link */}
          <Box
            sx={{
              px: 2,
              py: 1.25,
              borderBottom: "1px solid #EDEBE9",
              backgroundColor: "#FAFAFA",
              display: "flex",
              alignItems: "center",
              justifyContent: "space-between",
            }}
          >
            <Typography
              sx={{
                fontSize: 10,
                fontWeight: 700,
                color: "#A19F9D",
                textTransform: "uppercase",
                letterSpacing: 0.7,
              }}
            >
              Participants
            </Typography>
            <Box
              component="button"
              onClick={() => setManageOpen(true)}
              sx={{
                display: "flex",
                alignItems: "center",
                gap: 0.5,
                border: "none",
                background: "none",
                cursor: "pointer",
                color: "#0078D4",
                fontSize: 12,
                fontWeight: 500,
                p: 0,
                "&:hover": { textDecoration: "underline" },
              }}
            >
              <PeopleIcon sx={{ fontSize: 14 }} />
              Manage
            </Box>
          </Box>

          <Box sx={{ p: 2 }}>
            {/* Stat boxes */}
            <Box sx={{ display: "flex", gap: 1.5 }}>
              {[
                {
                  label: "Contributors",
                  count: contributors.length,
                  complete: contributorsComplete,
                },
                {
                  label: "Reviewers",
                  count: reviewers.length,
                  complete: reviewersComplete,
                },
              ].map((s) => (
                <Box
                  key={s.label}
                  sx={{
                    flex: 1,
                    backgroundColor: "#FAFAFA",
                    borderRadius: "8px",
                    p: 1.5,
                  }}
                >
                  <Typography
                    sx={{
                      fontSize: 10,
                      fontWeight: 700,
                      color: "#A19F9D",
                      textTransform: "uppercase",
                      letterSpacing: 0.6,
                      mb: 0.75,
                    }}
                  >
                    {s.label}
                  </Typography>
                  <Typography
                    sx={{
                      fontSize: 22,
                      fontWeight: 600,
                      color: "#323130",
                      lineHeight: 1,
                    }}
                  >
                    {s.count}
                  </Typography>
                  <Typography sx={{ fontSize: 11, color: "#605E5C", mt: 0.5 }}>
                    {s.complete} complete
                  </Typography>
                </Box>
              ))}
            </Box>

            {/* Progress bar */}
            {total > 0 && (
              <Box sx={{ mt: 2 }}>
                <Box
                  sx={{
                    display: "flex",
                    justifyContent: "space-between",
                    mb: 0.75,
                  }}
                >
                  <Typography sx={{ fontSize: 12, color: "#605E5C" }}>
                    Overall progress
                  </Typography>
                  <Typography
                    sx={{ fontSize: 12, fontWeight: 600, color: "#323130" }}
                  >
                    {complete} of {total} complete
                  </Typography>
                </Box>
                <Box
                  sx={{
                    height: 5,
                    backgroundColor: "#F3F2F1",
                    borderRadius: "99px",
                    overflow: "hidden",
                  }}
                >
                  <Box
                    sx={{
                      height: "100%",
                      width: `${progressPercent}%`,
                      backgroundColor:
                        progressPercent === 100 ? "#107C10" : "#0078D4",
                      borderRadius: "99px",
                      transition: "width 0.3s ease",
                    }}
                  />
                </Box>
                {inProgress > 0 && (
                  <Typography sx={{ fontSize: 11, color: "#0078D4", mt: 0.75 }}>
                    {inProgress} in progress
                  </Typography>
                )}
              </Box>
            )}

            {total === 0 && !loading && (
              <Typography
                sx={{
                  fontSize: 12,
                  color: "#A19F9D",
                  mt: 1.5,
                  fontStyle: "italic",
                }}
              >
                No participants assigned yet
              </Typography>
            )}
          </Box>
        </Box>

        {/* ── Minor Changes (existing documents only, no warning — release candidate handles it) ── */}
        {isExistingDocument && pendingMinorChanges.length > 0 && (
          <Box
            sx={{
              border: "1px solid #EDEBE9",
              borderRadius: "10px",
              overflow: "hidden",
              backgroundColor: "#fff",
            }}
          >
            {/* Header */}
            <Box
              sx={{
                px: 2,
                py: 1.25,
                borderBottom: "1px solid #EDEBE9",
                backgroundColor: "#FAFAFA",
                display: "flex",
                alignItems: "center",
                justifyContent: "space-between",
              }}
            >
              <Box sx={{ display: "flex", alignItems: "center", gap: 1 }}>
                <EditNoteIcon sx={{ fontSize: 15, color: "#605E5C" }} />
                <Typography
                  sx={{ fontSize: 12, fontWeight: 600, color: "#323130" }}
                >
                  Minor Changes
                </Typography>
              </Box>
              <Box sx={{ display: "flex", alignItems: "center", gap: 1.5 }}>
                <Typography sx={{ fontSize: 11, color: "#A19F9D" }}>
                  Click to view details
                </Typography>
                <Box
                  sx={{
                    display: "inline-flex",
                    px: 1,
                    py: 0.15,
                    borderRadius: "10px",
                    fontSize: 11,
                    fontWeight: 600,
                    backgroundColor: allMinorChangesActioned
                      ? "#DFF6DD"
                      : "#FFF4CE",
                    color: allMinorChangesActioned ? "#107C10" : "#835B00",
                  }}
                >
                  {
                    pendingMinorChanges.filter(
                      (mc) => mc.Status === "Implemented",
                    ).length
                  }{" "}
                  / {pendingMinorChanges.length}
                </Box>
              </Box>
            </Box>

            {/* Checklist — no warning banner, just the items */}
            <Box sx={{ py: 0.5 }}>
              {minorChangesLoading ? (
                <Box display="flex" justifyContent="center" py={2}>
                  <CircularProgress size={20} />
                </Box>
              ) : (
                pendingMinorChanges.map((mc) => {
                  const isChecked = mc.Status === "Implemented";
                  const isSaving = savingMinorId === mc.Id;

                  return (
                    <Box
                      key={mc.Id}
                      sx={{
                        display: "flex",
                        alignItems: "flex-start",
                        gap: 1,
                        py: 1,
                        px: 2,
                        borderBottom: "1px solid #F3F2F1",
                        transition: "background 0.12s",
                        "&:last-child": { borderBottom: "none" },
                        "&:hover": {
                          backgroundColor: "#F9F9F9",
                          "& .view-detail-hint": { opacity: 1 },
                        },
                      }}
                    >
                      {/* Checkbox */}
                      <Box
                        sx={{ pt: 0.1, flexShrink: 0 }}
                        onClick={(e) =>
                          !isSaving && handleToggleMinorChange(mc, e)
                        }
                      >
                        {isSaving ? (
                          <CircularProgress size={18} sx={{ ml: 0.25 }} />
                        ) : (
                          <Checkbox
                            checked={isChecked}
                            size="small"
                            sx={{
                              p: 0,
                              color: "#C8C6C4",
                              "&.Mui-checked": { color: "#107C10" },
                            }}
                          />
                        )}
                      </Box>

                      {/* Content */}
                      <Box
                        onClick={() => setSelectedMinorChange(mc)}
                        sx={{
                          flex: 1,
                          minWidth: 0,
                          cursor: "pointer",
                          display: "flex",
                          alignItems: "flex-start",
                          justifyContent: "space-between",
                          gap: 1,
                        }}
                      >
                        <Box sx={{ flex: 1, minWidth: 0 }}>
                          <Typography
                            sx={{
                              fontSize: 13,
                              fontWeight: 500,
                              color: isChecked ? "#A19F9D" : "#323130",
                              textDecoration: isChecked ? "line-through" : "none",
                              lineHeight: 1.4,
                            }}
                          >
                            {mc.Title}
                          </Typography>
                          {mc.ScopeOfChange && mc.ScopeOfChange !== mc.Title && (
                            <Typography
                              sx={{
                                fontSize: 12,
                                color: isChecked ? "#C8C6C4" : "#605E5C",
                                lineHeight: 1.4,
                                mt: 0.25,
                                overflow: "hidden",
                                textOverflow: "ellipsis",
                                display: "-webkit-box",
                                WebkitLineClamp: 2,
                                WebkitBoxOrient: "vertical",
                              }}
                            >
                              {mc.ScopeOfChange}
                            </Typography>
                          )}
                        </Box>

                        <Box
                          className="view-detail-hint"
                          sx={{
                            opacity: 0,
                            transition: "opacity 0.15s",
                            display: "flex",
                            alignItems: "center",
                            gap: 0.25,
                            flexShrink: 0,
                            pt: 0.25,
                          }}
                        >
                          <OpenInNewIcon
                            sx={{ fontSize: 13, color: "#0078D4" }}
                          />
                          <Typography
                            sx={{
                              fontSize: 11,
                              color: "#0078D4",
                              fontWeight: 500,
                              whiteSpace: "nowrap",
                            }}
                          >
                            View
                          </Typography>
                        </Box>
                      </Box>
                    </Box>
                  );
                })
              )}
            </Box>
          </Box>
        )}

        {/* ── Release Candidate (single action card with one warning) ── */}
        <Box
          sx={{
            border: "1px solid #EDEBE9",
            borderRadius: "10px",
            overflow: "hidden",
          }}
        >
          <Box
            sx={{
              px: 2,
              py: 1.25,
              borderBottom: "1px solid #EDEBE9",
              backgroundColor: "#FAFAFA",
              display: "flex",
              alignItems: "center",
              gap: 1,
            }}
          >
            <CheckCircleOutlineIcon sx={{ fontSize: 15, color: "#605E5C" }} />
            <Typography sx={{ fontSize: 12, fontWeight: 600, color: "#323130" }}>
              Release candidate
            </Typography>
          </Box>
          <Box sx={{ p: 2, display: "flex", flexDirection: "column", gap: 1.5 }}>
            {/* CA identity (only when CA is viewing) */}
            {isCa && (
              <Box sx={{ display: "flex", alignItems: "center", gap: 1 }}>
                <Box
                  sx={{
                    width: 26,
                    height: 26,
                    borderRadius: "50%",
                    backgroundColor: "#0078D4",
                    display: "flex",
                    alignItems: "center",
                    justifyContent: "center",
                    fontSize: 9,
                    fontWeight: 700,
                    color: "#fff",
                    flexShrink: 0,
                  }}
                >
                  {getInitials(cr.ChangeAuthority?.Title ?? "CA")}
                </Box>
                <Box>
                  <Typography
                    sx={{ fontSize: 12, fontWeight: 600, color: "#323130" }}
                  >
                    {cr.ChangeAuthority?.Title}
                  </Typography>
                  <Typography sx={{ fontSize: 11, color: "#A19F9D" }}>
                    Change Authority
                  </Typography>
                </Box>
              </Box>
            )}

            {/* Single status message */}
            {!canSubmit ? (
              <Box
                sx={{
                  display: "flex",
                  alignItems: "flex-start",
                  gap: 1,
                  p: 1.25,
                  backgroundColor: "#FFF4CE",
                  borderRadius: "8px",
                  border: "1px solid #FFB900",
                }}
              >
                <WarningAmberIcon
                  sx={{
                    fontSize: 14,
                    color: "#835B00",
                    mt: 0.15,
                    flexShrink: 0,
                  }}
                />
                <Typography
                  sx={{ fontSize: 12, color: "#835B00", lineHeight: 1.5 }}
                >
                  Waiting for {blockingReasons.join(" and ")}.
                </Typography>
              </Box>
            ) : (
              <Box
                sx={{
                  display: "flex",
                  alignItems: "center",
                  gap: 1,
                  p: 1.25,
                  backgroundColor: "#DFF6DD",
                  borderRadius: "8px",
                  border: "1px solid #107C10",
                }}
              >
                <CheckCircleOutlineIcon
                  sx={{ fontSize: 14, color: "#107C10", flexShrink: 0 }}
                />
                <Typography
                  sx={{ fontSize: 12, color: "#107C10", fontWeight: 600 }}
                >
                  Completed
                </Typography>
              </Box>
            )}

            {/* Actions */}
            <Box sx={{ display: "flex", gap: 1 }}>
              <Button
                variant="contained"
                disableElevation
                disabled={!canSubmit || submitting || (isCa && !authorComplete)}
                onClick={handleSubmitToPublishing}
                sx={{
                  flex: 1,
                  textTransform: "none",
                  fontWeight: 500,
                  fontSize: 13,
                  borderRadius: "6px",
                  backgroundColor: "#0078D4",
                  "&:disabled": {
                    backgroundColor: "#F3F2F1",
                    color: "#A19F9D",
                  },
                }}
              >
                {submitting
                  ? "Submitting..."
                  : isCa
                    ? "Submit to publishing"
                    : "Release candidate ready"}
              </Button>
              {isCa && !authorComplete && !submitting && (
                <Typography sx={{ fontSize: 12, color: "#605E5C", mt: 0.75 }}>
                  Waiting for the author to confirm before submitting to publishing.
                </Typography>
              )}
              {isCa && (
                <Button
                  variant="outlined"
                  disabled={submitting}
                  onClick={() => setRejectOpen(true)}
                  sx={{
                    textTransform: "none",
                    fontWeight: 500,
                    fontSize: 13,
                    borderRadius: "6px",
                    borderColor: "#F09595",
                    color: "#A32D2D",
                    "&:hover": {
                      backgroundColor: "#FCEBEB",
                      borderColor: "#E24B4A",
                    },
                  }}
                >
                  Reject
                </Button>
              )}
            </Box>
          </Box>
        </Box>

        {/* ── Minor Change Detail Modal ── */}
        <Dialog
          open={!!selectedMinorChange}
          onClose={() => setSelectedMinorChange(null)}
          maxWidth="sm"
          fullWidth
        >
          {selectedMinorChange &&
            (() => {
              const isChecked = selectedMinorChange.Status === "Implemented";
              const isSaving = savingMinorId === selectedMinorChange.Id;

              return (
                <>
                  <DialogTitle
                    sx={{
                      display: "flex",
                      alignItems: "flex-start",
                      justifyContent: "space-between",
                      gap: 2,
                    }}
                  >
                    <Box sx={{ display: "flex", alignItems: "center", gap: 1.5 }}>
                      <EditNoteIcon sx={{ fontSize: 20, color: "#835B00" }} />
                      <Box>
                        <Typography
                          sx={{
                            fontSize: 11,
                            fontWeight: 700,
                            color: "#A19F9D",
                            textTransform: "uppercase",
                            letterSpacing: 0.5,
                          }}
                        >
                          Minor Change
                        </Typography>
                        <Typography
                          sx={{
                            fontSize: 15,
                            fontWeight: 700,
                            color: "#323130",
                            lineHeight: 1.4,
                          }}
                        >
                          {selectedMinorChange.Title}
                        </Typography>
                      </Box>
                    </Box>
                    <IconButton
                      size="small"
                      onClick={() => setSelectedMinorChange(null)}
                    >
                      <CloseIcon sx={{ fontSize: 18 }} />
                    </IconButton>
                  </DialogTitle>

                  <DialogContent>
                    <Box
                      sx={{
                        display: "flex",
                        alignItems: "center",
                        gap: 1,
                        p: 1.25,
                        mb: 2.5,
                        borderRadius: "8px",
                        backgroundColor: isChecked ? "#DFF6DD" : "#FFF4CE",
                        border: `1px solid ${isChecked ? "#C8E6C9" : "#FFE082"}`,
                      }}
                    >
                      {isChecked ? (
                        <CheckCircleOutlineIcon
                          sx={{ fontSize: 16, color: "#107C10" }}
                        />
                      ) : (
                        <WarningAmberIcon
                          sx={{ fontSize: 16, color: "#835B00" }}
                        />
                      )}
                      <Typography
                        sx={{
                          fontSize: 13,
                          fontWeight: 600,
                          color: isChecked ? "#107C10" : "#835B00",
                        }}
                      >
                        {isChecked
                          ? "This change has been implemented."
                          : "This change has not been implemented yet."}
                      </Typography>
                    </Box>

                    <Box
                      sx={{
                        border: "1px solid #EDEBE9",
                        borderRadius: "8px",
                        p: 2,
                      }}
                    >
                      <DetailField
                        label="Scope of Change"
                        value={selectedMinorChange.ScopeOfChange}
                      />
                      <DetailField
                        label="Notes"
                        value={
                          selectedMinorChange.Notes ? (
                            <Typography
                              component="span"
                              sx={{
                                fontSize: 13,
                                color: "#323130",
                                fontStyle: "italic",
                              }}
                            >
                              {selectedMinorChange.Notes}
                            </Typography>
                          ) : null
                        }
                      />
                      <DetailField
                        label="Requested By"
                        value={selectedMinorChange.RequestedBy?.Title}
                      />
                      <DetailField
                        label="Registered"
                        value={formatDate(selectedMinorChange.Created)}
                      />
                      {selectedMinorChange.ChangeRequestId && (
                        <DetailField
                          label="Source CR"
                          value={`CR-${selectedMinorChange.ChangeRequestId}`}
                        />
                      )}
                    </Box>
                  </DialogContent>

                  <DialogActions
                    sx={{
                      px: 3,
                      pb: 2.5,
                      pt: 1.5,
                      display: "flex",
                      justifyContent: "space-between",
                    }}
                  >
                    <Button
                      variant="outlined"
                      color="inherit"
                      onClick={() => setSelectedMinorChange(null)}
                      sx={{ textTransform: "none" }}
                    >
                      Close
                    </Button>
                    <Button
                      variant="contained"
                      disableElevation
                      disabled={isSaving}
                      onClick={() => handleToggleMinorChange(selectedMinorChange)}
                      startIcon={
                        isSaving ? (
                          <CircularProgress size={14} color="inherit" />
                        ) : isChecked ? (
                          <CloseIcon sx={{ fontSize: 16 }} />
                        ) : (
                          <CheckCircleOutlineIcon sx={{ fontSize: 16 }} />
                        )
                      }
                      sx={{
                        textTransform: "none",
                        backgroundColor: isChecked ? "#605E5C" : "#107C10",
                        "&:hover": {
                          backgroundColor: isChecked ? "#4C555D" : "#0B6A0B",
                        },
                      }}
                    >
                      {isSaving
                        ? "Saving..."
                        : isChecked
                          ? "Mark as not implemented"
                          : "Mark as implemented"}
                    </Button>
                  </DialogActions>
                </>
              );
            })()}
        </Dialog>

        {/* ── Manage Modal ── */}
        <Dialog
          open={manageOpen}
          onClose={() => {
            setManageOpen(false);
            refetch().catch(console.error);
          }}
          maxWidth="lg"
          fullWidth
          PaperProps={{ sx: { borderRadius: 2, height: "80vh" } }}
        >
          <DialogTitle
            sx={{
              display: "flex",
              alignItems: "center",
              justifyContent: "space-between",
              fontSize: 15,
              fontWeight: 700,
            }}
          >
            Manage participants
            <IconButton
              size="small"
              onClick={() => {
                setManageOpen(false);
                refetch().catch(console.error);
              }}
            >
              <CloseIcon sx={{ fontSize: 18 }} />
            </IconButton>
          </DialogTitle>
          <DialogContent dividers sx={{ p: 3 }}>
            <ParticipantsTable
              changeRequestId={cr.ID}
              contributors={contributors}
              reviewers={reviewers}
              loading={loading}
              canAdd={isCa}
              canStart={isAuthor}
              canRemove={isCa || isAuthor}
              currentUserName={currentUser.Title}
              onRefetch={refetch}
            />
          </DialogContent>
        </Dialog>

        {/* ── Reject Modal ── */}
        <Dialog
          open={rejectOpen}
          onClose={() => setRejectOpen(false)}
          maxWidth="xs"
          fullWidth
          PaperProps={{ sx: { borderRadius: 2 } }}
        >
          <DialogTitle
            sx={{
              display: "flex",
              alignItems: "center",
              justifyContent: "space-between",
              fontSize: 14,
              fontWeight: 700,
            }}
          >
            Reject document changes
            <IconButton size="small" onClick={() => setRejectOpen(false)}>
              <CloseIcon sx={{ fontSize: 16 }} />
            </IconButton>
          </DialogTitle>
          <DialogContent sx={{ pt: 1 }}>
            <Typography sx={{ fontSize: 13, color: "#605E5C", mb: 2 }}>
              Provide a reason for rejection. This will be visible to the author.
            </Typography>
            <TextField
              label="Rejection reason"
              multiline
              rows={3}
              fullWidth
              size="small"
              value={rejectReason}
              onChange={(e) => setRejectReason(e.target.value)}
            />
          </DialogContent>
          <DialogActions sx={{ px: 3, pb: 2, gap: 1 }}>
            <Button
              size="small"
              onClick={() => setRejectOpen(false)}
              sx={{ textTransform: "none", color: "#605E5C" }}
            >
              Cancel
            </Button>
            <Button
              size="small"
              variant="contained"
              color="error"
              disableElevation
              disabled={!rejectReason.trim() || submitting}
              onClick={handleReject}
              sx={{ textTransform: "none" }}
            >
              {submitting ? "Rejecting..." : "Confirm reject"}
            </Button>
          </DialogActions>
        </Dialog>
      </Box>
    );
  };

  export default DocumentChangeProcessTask;
