import React, { useState, useEffect } from "react";
import Box from "@mui/material/Box";
import Typography from "@mui/material/Typography";
import CircularProgress from "@mui/material/CircularProgress";
import Divider from "@mui/material/Divider";
import Button from "@mui/material/Button";
import Tabs from "@mui/material/Tabs";
import Tab from "@mui/material/Tab";
import AttachFileIcon from "@mui/icons-material/AttachFile";
import ExpandMoreIcon from "@mui/icons-material/ExpandMore";
import LockOutlinedIcon from "@mui/icons-material/LockOutlined";
import { IChangeRequest } from "../../../shared/types/ChangeRequest";
import { Document } from "../../../shared/types/Document";
import { SharePointPerson } from "../../../shared/types/SharePointPerson";
import SharePointService from "../../../shared/services/SharePointService";
import { PeoplePicker } from "../../dcrForm/components/PeoplePicker";
import { InlineField } from "./InlineField";
import { useLookupData } from "../../../shared/hooks/useLookupData";
import { useDepartments } from "../../../shared/hooks/useDepartments";
import { useParticipants } from "../../../shared/hooks/useParticipants";
import MinorChangesTab from "./tabs/MinorChangesTab";
import ProgressTab from "./tabs/ProgressTab";
import { REVIEW_PERIOD_OPTIONS } from "../../../shared/constants";
import AuditTrailTab from "./tabs/AuditTrailTab";

// ─── Props ────────────────────────────────────────────────────────────────────

interface TaskDetailProps {
  cr: IChangeRequest | null;
  crLoading: boolean;
  onCRUpdate: () => void;
  currentUser?: SharePointPerson;
}

// ─── Status Badge ─────────────────────────────────────────────────────────────

const STATUS_STYLES: Record<
  string,
  { bg: string; color: string; dot: string }
> = {
  Submitted: { bg: "#EFF6FC", color: "#0078D4", dot: "#0078D4" },
  "CA Review": { bg: "#FFF4CE", color: "#835B00", dot: "#FFB900" },
  Approved: { bg: "#DFF6DD", color: "#107C10", dot: "#107C10" },
  "Document Creation": { bg: "#F0E6FF", color: "#5C2D91", dot: "#8764B8" },
  "Document Review": { bg: "#FFF4CE", color: "#835B00", dot: "#FFB900" },
  Published: { bg: "#DFF6DD", color: "#107C10", dot: "#107C10" },
  Rejected: { bg: "#FDE7E9", color: "#A4262C", dot: "#D13438" },
};

const StatusBadge = ({ status }: { status?: string }) => {
  if (!status)
    return (
      <Typography variant="body2" color="text.disabled">
        —
      </Typography>
    );
  const s = STATUS_STYLES[status] ?? {
    bg: "#F3F2F1",
    color: "#605E5C",
    dot: "#A19F9D",
  };
  return (
    <Box
      component="span"
      sx={{
        display: "inline-flex",
        alignItems: "center",
        gap: "6px",
        backgroundColor: s.bg,
        color: s.color,
        fontSize: 12,
        fontWeight: 600,
        px: "10px",
        py: "3px",
        borderRadius: "20px",
        letterSpacing: 0.3,
      }}
    >
      <Box
        component="span"
        sx={{
          width: 6,
          height: 6,
          borderRadius: "50%",
          backgroundColor: s.dot,
        }}
      />
      {status}
    </Box>
  );
};

// ─── Avatar ───────────────────────────────────────────────────────────────────

const AVATAR_COLORS = [
  "#0078D4",
  "#107C10",
  "#5C2D91",
  "#D83B01",
  "#008575",
  "#C239B3",
];
const getColor = (name: string) =>
  AVATAR_COLORS[name.charCodeAt(0) % AVATAR_COLORS.length];
const getInitials = (name: string) =>
  name
    .split(" ")
    .map((n) => n[0])
    .slice(0, 2)
    .join("")
    .toUpperCase();

const Avatar = ({ name, size = 26 }: { name: string; size?: number }) => (
  <Box
    sx={{
      width: size,
      height: size,
      borderRadius: "50%",
      backgroundColor: getColor(name),
      display: "flex",
      alignItems: "center",
      justifyContent: "center",
      fontSize: size * 0.36,
      fontWeight: 600,
      color: "#fff",
      flexShrink: 0,
    }}
  >
    {getInitials(name)}
  </Box>
);

// ─── Person Field ─────────────────────────────────────────────────────────────

interface PersonFieldProps {
  label: string;
  person: SharePointPerson | undefined | null;
  onSave: (person: SharePointPerson | undefined) => Promise<void>;
  required?: boolean;
  saving?: boolean;
  canEdit?: boolean;
}

const PersonField = ({
  label,
  person,
  onSave,
  required = false,
  saving = false,
  canEdit = false,
}: PersonFieldProps) => {
  const [editing, setEditing] = useState(false);
  const [hovered, setHovered] = useState(false);
  const isRequired = required && !person;

  const handleChange = async (selected: SharePointPerson | undefined) => {
    setEditing(false);
    await onSave(selected);
  };

  return (
    <Box
      sx={{
        display: "grid",
        gridTemplateColumns: "150px 1fr",
        alignItems: "start",
        py: "5px",
        gap: 1,
      }}
    >
      <Box sx={{ display: "flex", alignItems: "center", gap: 0.5, pt: "6px" }}>
        <Typography
          variant="caption"
          sx={{ color: "#605E5C", fontWeight: 500, fontSize: 13 }}
        >
          {label}
        </Typography>
        {required && (
          <Typography
            component="span"
            sx={{ color: "#D13438", fontSize: 11, lineHeight: 1 }}
          >
            *
          </Typography>
        )}
      </Box>
      {editing && canEdit ? (
        <Box sx={{ py: "2px" }}>
          <PeoplePicker
            label=""
            value={person ?? undefined}
            onChange={handleChange}
          />
        </Box>
      ) : (
        <Box
          onMouseEnter={() => setHovered(true)}
          onMouseLeave={() => setHovered(false)}
          onClick={() => canEdit && setEditing(true)}
          sx={{
            display: "flex",
            alignItems: "center",
            gap: 1,
            px: "8px",
            py: "4px",
            borderRadius: "4px",
            cursor: canEdit ? "pointer" : "default",
            transition: "background 0.12s",
            backgroundColor: isRequired
              ? "#FDE7E9"
              : hovered && canEdit
                ? "#F3F2F1"
                : "transparent",
            minHeight: 32,
            position: "relative",
          }}
        >
          {person ? (
            <>
              <Avatar name={person.Title} size={22} />
              <Typography variant="body2" color="text.primary">
                {person.Title}
              </Typography>
            </>
          ) : (
            <Typography
              variant="body2"
              color={isRequired ? "error" : "text.disabled"}
            >
              None
            </Typography>
          )}
          {saving && (
            <CircularProgress
              size={12}
              sx={{ position: "absolute", right: 8, color: "#0078D4" }}
            />
          )}
        </Box>
      )}
    </Box>
  );
};

// ─── Multi Person Field ───────────────────────────────────────────────────────

interface MultiPersonFieldProps {
  label: string;
  people: SharePointPerson[] | undefined;
  onSave: (people: SharePointPerson[]) => Promise<void>;
  saving?: boolean;
  canEdit?: boolean;
  excludeIds?: number[];
}

const MultiPersonField = ({
  label,
  people = [],
  onSave,
  saving = false,
  canEdit = false,
  excludeIds = [],
}: MultiPersonFieldProps) => {
  const [adding, setAdding] = useState(false);
  const [hovered, setHovered] = useState(false);

  const handleAdd = async (person: SharePointPerson | undefined) => {
    setAdding(false);
    if (!person || people.some((p) => p.Id === person.Id)) return;
    await onSave([...people, person]);
  };

  const handleRemove = async (id: number) => {
    await onSave(people.filter((p) => p.Id !== id));
  };

  return (
    <Box
      sx={{
        display: "grid",
        gridTemplateColumns: "150px 1fr",
        alignItems: "start",
        py: "5px",
        gap: 1,
      }}
    >
      <Typography
        variant="caption"
        sx={{ color: "#605E5C", fontWeight: 500, fontSize: 13, pt: "6px" }}
      >
        {label}
      </Typography>
      <Box sx={{ px: "8px", py: "4px", minHeight: 32 }}>
        <Box
          sx={{
            display: "flex",
            flexWrap: "wrap",
            gap: 0.75,
            mb: people.length > 0 ? 1 : 0,
          }}
        >
          {people.map((person) => (
            <Box
              key={person.Id}
              sx={{
                display: "flex",
                alignItems: "center",
                gap: 0.75,
                backgroundColor: "#F3F2F1",
                borderRadius: "4px",
                pl: 0.75,
                pr: 0.5,
                py: 0.25,
              }}
            >
              <Avatar name={person.Title} size={18} />
              <Typography
                variant="caption"
                sx={{ fontSize: 12, color: "#323130" }}
              >
                {person.Title}
              </Typography>
              {canEdit && (
                <Box
                  component="button"
                  onClick={() => handleRemove(person.Id)}
                  sx={{
                    display: "flex",
                    alignItems: "center",
                    justifyContent: "center",
                    width: 16,
                    height: 16,
                    borderRadius: "50%",
                    border: "none",
                    background: "none",
                    cursor: "pointer",
                    color: "#A19F9D",
                    fontSize: 14,
                    lineHeight: 1,
                    p: 0,
                    "&:hover": { backgroundColor: "#EDEBE9", color: "#323130" },
                  }}
                >
                  ×
                </Box>
              )}
            </Box>
          ))}
        </Box>
        {canEdit &&
          (adding ? (
            <Box sx={{ maxWidth: 280 }}>
              <PeoplePicker
                label=""
                value={undefined}
                onChange={handleAdd}
                excludeIds={excludeIds}
              />
            </Box>
          ) : (
            <Box
              component="button"
              onClick={() => setAdding(true)}
              onMouseEnter={() => setHovered(true)}
              onMouseLeave={() => setHovered(false)}
              sx={{
                display: "inline-flex",
                alignItems: "center",
                gap: 0.5,
                border: "none",
                background: "none",
                cursor: "pointer",
                color: hovered ? "#0078D4" : "#A19F9D",
                fontSize: 12,
                fontWeight: 500,
                p: 0,
                transition: "color 0.15s",
              }}
            >
              <Typography
                sx={{ fontSize: 16, lineHeight: 1, color: "inherit" }}
              >
                +
              </Typography>
              Add person
            </Box>
          ))}
        {saving && (
          <CircularProgress size={12} sx={{ ml: 1, color: "#0078D4" }} />
        )}
      </Box>
    </Box>
  );
};

// ─── Multi Select Field ───────────────────────────────────────────────────────

interface LookupItem {
  Id: number;
  Title: string;
}

interface MultiSelectFieldProps {
  label: string;
  selectedItems: LookupItem[];
  options: LookupItem[];
  onSave: (ids: number[]) => Promise<void>;
  saving?: boolean;
  canEdit?: boolean;
}

const MultiSelectField = ({
  label,
  selectedItems = [],
  options,
  onSave,
  saving = false,
  canEdit = false,
}: MultiSelectFieldProps) => {
  const [adding, setAdding] = useState(false);
  const [hovered, setHovered] = useState(false);

  const handleAdd = async (id: number) => {
    setAdding(false);
    if (selectedItems.some((i) => i.Id === id)) return;
    await onSave([...selectedItems.map((i) => i.Id), id]);
  };

  const handleRemove = async (id: number) => {
    await onSave(selectedItems.filter((i) => i.Id !== id).map((i) => i.Id));
  };

  const availableOptions = options.filter(
    (o) => !selectedItems.some((s) => s.Id === o.Id),
  );

  return (
    <Box
      sx={{
        display: "grid",
        gridTemplateColumns: "150px 1fr",
        alignItems: "start",
        py: "5px",
        gap: 1,
      }}
    >
      <Typography
        variant="caption"
        sx={{ color: "#605E5C", fontWeight: 500, fontSize: 13, pt: "6px" }}
      >
        {label}
      </Typography>
      <Box sx={{ px: "8px", py: "4px", minHeight: 32 }}>
        <Box
          sx={{
            display: "flex",
            flexWrap: "wrap",
            gap: 0.75,
            mb: selectedItems.length > 0 ? 1 : 0,
          }}
        >
          {selectedItems.map((item) => (
            <Box
              key={item.Id}
              sx={{
                display: "flex",
                alignItems: "center",
                gap: 0.5,
                backgroundColor: "#F3F2F1",
                borderRadius: "4px",
                pl: 1,
                pr: 0.5,
                py: 0.25,
              }}
            >
              <Typography
                variant="caption"
                sx={{ fontSize: 12, color: "#323130" }}
              >
                {item.Title}
              </Typography>
              {canEdit && (
                <Box
                  component="button"
                  onClick={() => handleRemove(item.Id)}
                  sx={{
                    display: "flex",
                    alignItems: "center",
                    justifyContent: "center",
                    width: 16,
                    height: 16,
                    borderRadius: "50%",
                    border: "none",
                    background: "none",
                    cursor: "pointer",
                    color: "#A19F9D",
                    fontSize: 14,
                    lineHeight: 1,
                    p: 0,
                    "&:hover": { backgroundColor: "#EDEBE9", color: "#323130" },
                  }}
                >
                  ×
                </Box>
              )}
            </Box>
          ))}
        </Box>
        {canEdit &&
          (adding ? (
            <Box
              sx={{
                display: "flex",
                flexDirection: "column",
                maxWidth: 220,
                backgroundColor: "#fff",
                border: "1px solid #EDEBE9",
                borderRadius: "4px",
                boxShadow: "0 4px 12px rgba(0,0,0,0.1)",
                overflow: "hidden",
              }}
            >
              {availableOptions.length === 0 ? (
                <Typography sx={{ fontSize: 12, color: "#A19F9D", p: 1 }}>
                  No more options
                </Typography>
              ) : (
                availableOptions.map((opt) => (
                  <Box
                    key={opt.Id}
                    onMouseDown={() => handleAdd(opt.Id)}
                    sx={{
                      px: 1.5,
                      py: 1,
                      fontSize: 13,
                      color: "#323130",
                      cursor: "pointer",
                      "&:hover": { backgroundColor: "#F3F2F1" },
                    }}
                  >
                    {opt.Title}
                  </Box>
                ))
              )}
            </Box>
          ) : (
            <Box
              component="button"
              onClick={() => setAdding(true)}
              onMouseEnter={() => setHovered(true)}
              onMouseLeave={() => setHovered(false)}
              sx={{
                display: "inline-flex",
                alignItems: "center",
                gap: 0.5,
                border: "none",
                background: "none",
                cursor: "pointer",
                color: hovered ? "#0078D4" : "#A19F9D",
                fontSize: 12,
                fontWeight: 500,
                p: 0,
                transition: "color 0.15s",
              }}
            >
              <Typography
                sx={{ fontSize: 16, lineHeight: 1, color: "inherit" }}
              >
                +
              </Typography>
              Add
            </Box>
          ))}
        {saving && (
          <CircularProgress size={12} sx={{ ml: 1, color: "#0078D4" }} />
        )}
      </Box>
    </Box>
  );
};

// ─── Section ──────────────────────────────────────────────────────────────────

const Section = ({
  title,
  children,
  defaultOpen = true,
}: {
  title: string;
  children: React.ReactNode;
  defaultOpen?: boolean;
}) => {
  const [open, setOpen] = useState(defaultOpen);
  return (
    <Box>
      <Box
        component="button"
        onClick={() => setOpen(!open)}
        sx={{
          display: "flex",
          alignItems: "center",
          gap: 1,
          background: "none",
          border: "none",
          cursor: "pointer",
          py: 1.25,
          px: 0,
          width: "100%",
          textAlign: "left",
          "&:hover .section-title": { color: "#0078D4" },
        }}
      >
        <ExpandMoreIcon
          sx={{
            fontSize: 18,
            color: "#605E5C",
            transform: open ? "rotate(0deg)" : "rotate(-90deg)",
            transition: "transform 0.2s ease",
          }}
        />
        <Typography
          className="section-title"
          sx={{
            fontSize: 12,
            fontWeight: 700,
            color: "#323130",
            letterSpacing: 0.6,
            textTransform: "uppercase",
            transition: "color 0.15s",
          }}
        >
          {title}
        </Typography>
      </Box>
      <Divider sx={{ mb: 1 }} />
      {open && <Box pb={1.5}>{children}</Box>}
    </Box>
  );
};

// ─── Main Component ───────────────────────────────────────────────────────────

export const TaskDetail = ({
  cr,
  crLoading,
  onCRUpdate,
  currentUser,
}: TaskDetailProps) => {
  const [tab, setTab] = useState(0);
  const [savingField, setSavingField] = useState<string | null>(null);
  const [publishedDoc, setPublishedDoc] = useState<Document | null>(null);
  const [docLoading, setDocLoading] = useState(false);
  const [attachments, setAttachments] = useState<
    { FileName: string; ServerRelativeUrl: string }[]
  >([]);
  const [attachmentsLoading, setAttachmentsLoading] = useState(false);
  const {
    contributors,
    reviewers,
    refetch: refetchParticipants,
  } = useParticipants(cr?.ID);
  const { documentTypes, categories, audienceGroups, businessFunctions } =
    useLookupData();
  const { departments } = useDepartments();

  // Only the CA can edit
  const canEdit =
    !!currentUser &&
    !!cr?.ChangeAuthority &&
    currentUser.Id === cr.ChangeAuthority.Id;
  console.log("canEdit check:", {
    currentUserId: currentUser?.Id,
    changeAuthorityId: cr?.ChangeAuthority?.Id,
    canEdit,
  });
  const documentTypeMap = documentTypes.reduce<Record<string, number>>(
    (acc, t) => ({ ...acc, [t.Title]: t.Id }) as Record<string, number>,
    {},
  );
  const audienceMap = audienceGroups.reduce<Record<string, number>>(
    (acc, a) => ({ ...acc, [a.Title]: a.Id }) as Record<string, number>,
    {},
  );
  const coreFunctionalityMap = departments.reduce<Record<string, number>>(
    (acc, d) => ({ ...acc, [d.Title]: d.Id }) as Record<string, number>,
    {},
  );

  useEffect(() => {
    if (!cr?.TargetDocumentId) return;
    setDocLoading(true);
    setPublishedDoc(null);
    SharePointService.getDocumentById(cr.TargetDocumentId)
      .then(setPublishedDoc)
      .catch(console.error)
      .finally(() => setDocLoading(false));
  }, [cr?.TargetDocumentId]);

  useEffect(() => {
    setTab(0);
  }, [cr?.ID]);

  useEffect(() => {
    if (!cr?.ID) return;
    setAttachmentsLoading(true);
    setAttachments([]);
    SharePointService.getAttachments(cr.ID)
      .then(setAttachments)
      .catch(console.error)
      .finally(() => setAttachmentsLoading(false));
  }, [cr?.ID]);

  const handleFieldSave = async (spFieldName: string, value: string) => {
    if (!cr) return;
    setSavingField(spFieldName);
    try {
      await SharePointService.updateChangeRequest(cr.ID, {
        [spFieldName]: value !== "" ? value : null,
      });
      onCRUpdate();
    } catch (err) {
      console.error(`Failed to save ${spFieldName}:`, err);
    } finally {
      setSavingField(null);
    }
  };

  const handleLookupSave = async (spFieldName: string, id: number | null) => {
    if (!cr) return;
    setSavingField(spFieldName);
    try {
      await SharePointService.updateChangeRequest(cr.ID, { [spFieldName]: id });
      onCRUpdate();
    } catch (err) {
      console.error(`Failed to save ${spFieldName}:`, err);
    } finally {
      setSavingField(null);
    }
  };

  const handleMultiLookupSave = async (spFieldName: string, ids: number[]) => {
    if (!cr) return;
    setSavingField(spFieldName);
    try {
      await SharePointService.updateChangeRequest(cr.ID, {
        [spFieldName]: ids,
      });
      onCRUpdate();
    } catch (err) {
      console.error(`Failed to save ${spFieldName}:`, err);
    } finally {
      setSavingField(null);
    }
  };

  const handlePersonSave = async (
    spFieldName: string,
    person: SharePointPerson | undefined,
  ) => {
    if (!cr) return;
    setSavingField(spFieldName);
    try {
      await SharePointService.updateChangeRequest(cr.ID, {
        [spFieldName]: person?.Id ?? null,
      });
      onCRUpdate();
    } catch (err) {
      console.error(`Failed to save ${spFieldName}:`, err);
    } finally {
      setSavingField(null);
    }
  };

  const handleAddParticipant = async (
    person: SharePointPerson,
    role: "Reviewer" | "Contributor",
  ): Promise<void> => {
    if (!cr) return;
    await SharePointService.addParticipant(cr.ID, person.Id, role);
    refetchParticipants().catch(console.error);
  };

  const handleRemoveParticipant = async (
    personId: number,
    role: "Reviewer" | "Contributor",
  ): Promise<void> => {
    const row = [...contributors, ...reviewers].find(
      (p) => p.Person?.Id === personId && p.Role === role,
    );
    if (!row || !cr) return;
    await SharePointService.deleteParticipant(
      row.Id,
      cr.ID,
      personId,
      "Removed from task detail",
      currentUser?.Title ?? "Unknown",
    );
    refetchParticipants().catch(console.error);
  };

  const contributorPeople = contributors
    .map((participant) => participant.Person)
    .filter((person): person is SharePointPerson => !!person?.Id);

  const reviewerPeople = reviewers
    .map((participant) => participant.Person)
    .filter((person): person is SharePointPerson => !!person?.Id);

  if (crLoading)
    return (
      <Box
        display="flex"
        alignItems="center"
        justifyContent="center"
        height="100%"
      >
        <CircularProgress size={28} />
      </Box>
    );

  if (!cr)
    return (
      <Box
        display="flex"
        alignItems="center"
        justifyContent="center"
        height="100%"
        flexDirection="column"
        gap={1}
      >
        <Typography variant="h6" color="text.secondary">
          No change request found
        </Typography>
        <Typography variant="body2" color="text.disabled">
          CR details will appear here
        </Typography>
      </Box>
    );

  return (
    <Box
      display="flex"
      flexDirection="column"
      height="100%"
      sx={{ backgroundColor: "#FAFAFA" }}
    >
      {/* ── Header ── */}
      <Box
        sx={{
          backgroundColor: "#fff",
          borderBottom: "1px solid #EDEBE9",
          px: 3,
          py: 2.5,
        }}
      >
        <Box
          display="flex"
          alignItems="center"
          justifyContent="space-between"
          mb={0.75}
        >
          <Typography
            sx={{
              fontSize: 11,
              fontWeight: 700,
              color: "#A19F9D",
              letterSpacing: 0.8,
            }}
          >
            {cr.ChangeRequestNumber ?? `CR-${cr.ID}`}
          </Typography>
          <Box sx={{ display: "flex", alignItems: "center", gap: 1 }}>
            {!canEdit && (
              <Box
                sx={{
                  display: "flex",
                  alignItems: "center",
                  gap: 0.5,
                  color: "#A19F9D",
                }}
              >
                <LockOutlinedIcon sx={{ fontSize: 13 }} />
                <Typography sx={{ fontSize: 11, color: "#A19F9D" }}>
                  View only
                </Typography>
              </Box>
            )}
            <StatusBadge status={cr.Status} />
          </Box>
        </Box>
        <Typography
          sx={{
            fontSize: 16,
            fontWeight: 700,
            color: "#323130",
            lineHeight: 1.4,
            mb: 0.5,
          }}
        >
          {cr.Title}
        </Typography>
        {cr.ScopeofChange && (
          <Typography
            variant="body2"
            color="text.secondary"
            sx={{ lineHeight: 1.6, fontSize: 13 }}
          >
            {cr.ScopeofChange}
          </Typography>
        )}
      </Box>

      {/* ── Tabs ── */}
      <Tabs
        value={tab}
        onChange={(_, v) => setTab(v)}
        variant="scrollable"
        scrollButtons="auto"
        sx={{
          backgroundColor: "#fff",
          borderBottom: "1px solid #EDEBE9",
          px: 2,
          minHeight: 40,
          "& .MuiTab-root": { minHeight: 40, fontSize: 13, py: 0 },
        }}
      >
        <Tab label="Details" />
        <Tab
          label={`Attachments${attachments.length > 0 ? ` (${attachments.length})` : ""}`}
        />

        <Tab label="Published Document" />
        {!cr.NewDocument && <Tab label={`Minor Changes`} />}
        <Tab label="Progress" value={4} />
        <Tab label="Audit Trail" value={5} />
      </Tabs>

      {/* ── Content ── */}
      <Box flex={1} overflow="auto">
        {tab === 0 && (
          <Box
            sx={{
              maxWidth: 760,
              px: 3,
              py: 2.5,
              display: "flex",
              flexDirection: "column",
              gap: 2,
            }}
          >
            <Section title="General">
              <InlineField
                label="Title"
                value={cr.Title}
                onSave={(v) => handleFieldSave("Title", v)}
                saving={savingField === "Title"}
                disabled={!canEdit}
              />
              <InlineField
                label="Scope of Change"
                value={cr.ScopeofChange}
                onSave={(v) => handleFieldSave("ScopeofChange", v)}
                type="textarea"
                saving={savingField === "ScopeofChange"}
                disabled={!canEdit}
              />
              <InlineField
                label="Urgency"
                value={cr.Urgency}
                onSave={(v) => handleFieldSave("Urgency", v)}
                type="select"
                options={["Standard", "Urgent", "Minor"]}
                saving={savingField === "Urgency"}
                disabled={!canEdit}
                renderValue={(v) => {
                  const colors: Record<string, string> = {
                    Urgent: "#ff1744",
                    Minor: "#00c853",
                    Standard: "#757575",
                  };
                  return v ? (
                    <Typography
                      variant="body2"
                      sx={{ color: colors[v] ?? "#605E5C", fontWeight: 600 }}
                    >
                      {v}
                    </Typography>
                  ) : null;
                }}
              />
              <InlineField
                label="New Document"
                value={cr.NewDocument ? "Yes" : "No"}
                onSave={() => Promise.resolve()}
                renderValue={(v) => (
                  <Typography variant="body2" color="text.primary">
                    {v}
                  </Typography>
                )}
                disabled
              />
              <InlineField
                label="Draft Document Name"
                value={cr.DraftDocumentName}
                onSave={(v) => handleFieldSave("DraftDocumentName", v)}
                saving={savingField === "DraftDocumentName"}
                placeholder="Not set"
                disabled={!canEdit}
              />
              <InlineField
                label="Review Period"
                value={
                  cr.ReviewPeriod
                    ? REVIEW_PERIOD_OPTIONS.find(
                        (option) => option.value === cr.ReviewPeriod
                      )?.label ?? `${cr.ReviewPeriod} months`
                    : undefined
                }
                onSave={(v) => {
                  const selected = REVIEW_PERIOD_OPTIONS.find(
                    (option) => option.label === v
                  );
                  return handleLookupSave("ReviewPeriod", selected?.value ?? null);
                }}
                type="select"
                options={REVIEW_PERIOD_OPTIONS.map((option) => option.label)}
                saving={savingField === "ReviewPeriod"}
                disabled={!canEdit}
                placeholder="Not set"
              />
            </Section>

            <Section title="Document">
              <InlineField
                label="Classification"
                value={cr.Classification}
                onSave={(v) => handleFieldSave("Classification", v)}
                type="select"
                options={["Public", "Internal", "Confidential", "Restricted"]}
                saving={savingField === "Classification"}
                disabled={!canEdit}
                renderValue={(v) =>
                  v ? (
                    <Box
                      component="span"
                      sx={{
                        display: "inline-flex",
                        alignItems: "center",
                        backgroundColor: "#EFF6FC",
                        color: "#0078D4",
                        fontSize: 12,
                        fontWeight: 600,
                        px: 1.25,
                        py: 0.25,
                        borderRadius: "10px",
                      }}
                    >
                      {v}
                    </Box>
                  ) : null
                }
              />
              <InlineField
                label="Document Type"
                value={cr.DocumentType?.Title}
                onSave={(v) =>
                  handleLookupSave("DocumentTypeId", documentTypeMap[v] ?? null)
                }
                type="select"
                options={documentTypes.map((t) => t.Title)}
                saving={savingField === "DocumentTypeId"}
                disabled={!canEdit}
              />
              <MultiSelectField
                label="Category"
                selectedItems={cr.Category ?? []}
                options={categories}
                onSave={(ids) => handleMultiLookupSave("CategoryId", ids)}
                saving={savingField === "CategoryId"}
                canEdit={canEdit}
              />
              <InlineField
                label="Audience"
                value={cr.Audience?.Title}
                onSave={(v) =>
                  handleLookupSave("AudienceId", audienceMap[v] ?? null)
                }
                type="select"
                options={audienceGroups.map((a) => a.Title)}
                saving={savingField === "AudienceId"}
                disabled={!canEdit}
              />
            </Section>

            <Section title="Department">
              <InlineField
                label="Core Functionality"
                value={cr.CoreFunctionality?.Title}
                onSave={async (v) => {
                  const dept = departments.find((d) => d.Title === v);
                  const payload: Record<string, unknown> = {
                    CoreFunctionalityId: coreFunctionalityMap[v] ?? null,
                  };
                  if (dept?.ChangeAuthority?.Id)
                    payload.ChangeAuthorityId = dept.ChangeAuthority.Id;
                  setSavingField("CoreFunctionalityId");
                  try {
                    await SharePointService.updateChangeRequest(cr.ID, payload);
                    onCRUpdate();
                  } catch (err) {
                    console.error("Failed to save Core Functionality:", err);
                  } finally {
                    setSavingField(null);
                  }
                }}
                type="select"
                options={departments.map((d) => d.Title)}
                saving={savingField === "CoreFunctionalityId"}
                placeholder="Not assigned"
                disabled={!canEdit}
              />
              <MultiSelectField
                label="Business Function"
                selectedItems={cr.BusinessFunction ?? []}
                options={businessFunctions}
                onSave={(ids) =>
                  handleMultiLookupSave("BusinessFunctionId", ids)
                }
                saving={savingField === "BusinessFunctionId"}
                canEdit={canEdit}
              />
            </Section>

            <Section title="People">
              <PersonField
                label="Change Authority"
                person={cr.ChangeAuthority}
                onSave={(p) => handlePersonSave("ChangeAuthorityId", p)}
                saving={savingField === "ChangeAuthorityId"}
                canEdit={canEdit}
              />
              <PersonField
                label="Release Authority"
                person={cr.ReleaseAuthority}
                onSave={(p) => handlePersonSave("ReleaseAuthorityId", p)}
                saving={savingField === "ReleaseAuthorityId"}
                required
                canEdit={canEdit}
              />
              <PersonField
                label="Author"
                person={cr.Author0}
                onSave={(p) => handlePersonSave("Author0Id", p)}
                saving={savingField === "Author0Id"}
                canEdit={canEdit}
              />
              <InlineField
                label="Submitted By"
                value={cr.Author?.Title}
                onSave={() => Promise.resolve()}
                disabled
              />
              <MultiPersonField
                label="Contributors"
                people={contributorPeople}
                onSave={async (newPeople) => {
                  const added = newPeople.filter(
                    (p) => !contributors.some((c) => c.Person?.Id === p.Id),
                  );
                  const removed = contributors.filter(
                    (c) => !newPeople.some((p) => p.Id === c.Person?.Id),
                  );
                  await Promise.all(
                    added.map((p) => handleAddParticipant(p, "Contributor")),
                  );
                  await Promise.all(
                    removed
                      .map((r) => r.Person?.Id)
                      .filter((id): id is number => id !== undefined)
                      .map((id) => handleRemoveParticipant(id, "Contributor")),
                  );
                }}
                canEdit={canEdit}
                excludeIds={reviewerPeople.map((p) => p.Id)}
              />
              <MultiPersonField
                label="Reviewers"
                people={reviewerPeople}
                onSave={async (newPeople) => {
                  const added = newPeople.filter(
                    (p) => !reviewers.some((r) => r.Person?.Id === p.Id),
                  );
                  const removed = reviewers.filter(
                    (r) => !newPeople.some((p) => p.Id === r.Person?.Id),
                  );
                  await Promise.all(
                    added.map((p) => handleAddParticipant(p, "Reviewer")),
                  );
                  await Promise.all(
                    removed
                      .map((r) => r.Person?.Id)
                      .filter((id): id is number => id !== undefined)
                      .map((id) => handleRemoveParticipant(id, "Reviewer")),
                  );
                }}
                canEdit={canEdit}
                excludeIds={contributorPeople.map((p) => p.Id)}
              />
            </Section>

            {/* ── Compliance Section ── */}
            {(cr.ReadAcknowledgementRequired || (cr.ReadAudienceGroups?.length ?? 0) > 0) && (
              <Section title="Compliance">
                <InlineField
                  label="Read acknowledgement"
                  value={cr.ReadAcknowledgementRequired ? "Required" : "Not required"}
                  onSave={() => Promise.resolve()}
                  disabled
                  renderValue={(v) => (
                    <Box
                      component="span"
                      sx={{
                        display: "inline-flex",
                        alignItems: "center",
                        gap: 0.5,
                        backgroundColor: cr.ReadAcknowledgementRequired
                          ? "#DFF6DD"
                          : "#F3F2F1",
                        color: cr.ReadAcknowledgementRequired
                          ? "#107C10"
                          : "#605E5C",
                        fontSize: 12,
                        fontWeight: 600,
                        px: 1.25,
                        py: 0.25,
                        borderRadius: "10px",
                      }}
                    >
                      {v}
                    </Box>
                  )}
                />
                {cr.ReadAcknowledgementRequired && (cr.ReadAudienceGroups?.length ?? 0) > 0 && (
                  <InlineField
                    label="Audience"
                    value={(cr.ReadAudienceGroups ?? []).map((a) => a.Title).join(", ")}
                    onSave={() => Promise.resolve()}
                    disabled
                  />
                )}
                {cr.ReadAcknowledgementRequired && cr.ReadDueDate && (
                  <InlineField
                    label="Due date"
                    value={new Date(cr.ReadDueDate).toLocaleDateString("en-AU", {
                      day: "numeric",
                      month: "short",
                      year: "numeric",
                    })}
                    onSave={() => Promise.resolve()}
                    disabled
                  />
                )}
              </Section>
            )}
          </Box>
        )}

        {tab === 1 && (
          <Box p={3}>
            {attachmentsLoading ? (
              <Box display="flex" justifyContent="center" mt={4}>
                <CircularProgress size={24} />
              </Box>
            ) : attachments.length === 0 ? (
              <Box
                sx={{
                  border: "1px dashed #D2D0CE",
                  borderRadius: 2,
                  p: 4,
                  textAlign: "center",
                }}
              >
                <AttachFileIcon
                  sx={{ color: "#C8C6C4", fontSize: 32, mb: 1 }}
                />
                <Typography color="text.secondary" fontSize={13}>
                  No attachments
                </Typography>
              </Box>
            ) : (
              <Box display="flex" flexDirection="column" gap={1}>
                {attachments.map((file) => (
                  <Box
                    key={file.FileName}
                    sx={{
                      display: "flex",
                      alignItems: "center",
                      gap: 2,
                      py: 1.5,
                      px: 2,
                      backgroundColor: "#fff",
                      border: "1px solid #EDEBE9",
                      borderRadius: 1,
                      "&:hover": { backgroundColor: "#F3F2F1" },
                    }}
                  >
                    <AttachFileIcon
                      fontSize="small"
                      sx={{ color: "#605E5C" }}
                    />
                    <Typography
                      variant="body2"
                      flex={1}
                      sx={{ color: "#323130", fontWeight: 500 }}
                    >
                      {file.FileName}
                    </Typography>
                    <Box display="flex" gap={1}>
                      <Button
                        size="small"
                        variant="outlined"
                        href={file.ServerRelativeUrl}
                        target="_blank"
                        rel="noopener noreferrer"
                      >
                        Open
                      </Button>
                      <Button
                        size="small"
                        variant="contained"
                        disableElevation
                        href={`${file.ServerRelativeUrl}?download=1`}
                        download={file.FileName}
                      >
                        Download
                      </Button>
                    </Box>
                  </Box>
                ))}
              </Box>
            )}
          </Box>
        )}

        {tab === 2 && (
          <Box p={3}>
            {docLoading && (
              <Box display="flex" justifyContent="center" mt={4}>
                <CircularProgress size={24} />
              </Box>
            )}
            {!docLoading && !publishedDoc && (
              <Box
                sx={{
                  border: "1px dashed #D2D0CE",
                  borderRadius: 2,
                  p: 4,
                  textAlign: "center",
                }}
              >
                <AttachFileIcon
                  sx={{ color: "#C8C6C4", fontSize: 32, mb: 1 }}
                />
                <Typography color="text.secondary" fontSize={13}>
                  No published document linked to this change request.
                </Typography>
              </Box>
            )}
            {!docLoading && publishedDoc?.FileRef && (
              <Box
                sx={{
                  display: "flex",
                  alignItems: "center",
                  gap: 2,
                  py: 1.5,
                  px: 2,
                  backgroundColor: "#fff",
                  border: "1px solid #EDEBE9",
                  borderRadius: 1,
                  "&:hover": { backgroundColor: "#F3F2F1" },
                }}
              >
                <AttachFileIcon fontSize="small" sx={{ color: "#605E5C" }} />
                <Typography
                  variant="body2"
                  flex={1}
                  sx={{ color: "#323130", fontWeight: 500 }}
                >
                  {publishedDoc.FileLeafRef ??
                    publishedDoc.DocumentTitle ??
                    "Published Document"}
                </Typography>
                <Box display="flex" gap={1}>
                  <Button
                    size="small"
                    variant="outlined"
                    href={publishedDoc.FileRef}
                    target="_blank"
                    rel="noopener noreferrer"
                  >
                    Open
                  </Button>
                  <Button
                    size="small"
                    variant="contained"
                    disableElevation
                    href={`${publishedDoc.FileRef}?download=1`}
                    download={publishedDoc.FileLeafRef}
                  >
                    Download
                  </Button>
                </Box>
              </Box>
            )}
          </Box>
        )}
        {!cr.NewDocument && tab === 3 && (
          <MinorChangesTab documentId={cr.TargetDocumentId} />
        )}
        {tab === 4 && <ProgressTab cr={cr} />}
        {tab === 5 && <AuditTrailTab cr={cr} />}
      </Box>
    </Box>
  );
};
