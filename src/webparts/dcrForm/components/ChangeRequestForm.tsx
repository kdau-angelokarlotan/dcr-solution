// src/webparts/submitChangeRequest/components/ChangeRequestForm.tsx
import React, { useEffect, useState } from "react";
import {
  Box,
  Button,
  Typography,
  Collapse,
  Alert,
  Fade,
  CircularProgress,
} from "@mui/material";
import ExpandMoreIcon from "@mui/icons-material/ExpandMore";
import ExpandLessIcon from "@mui/icons-material/ExpandLess";
import { SharePointPerson } from "../../../shared/types/SharePointPerson";
import { IChangeRequest } from "../../../shared/types/ChangeRequest";
import { useDocuments } from "../../../shared/hooks/useDocuments";
import { useDepartments } from "../../../shared/hooks/useDepartments";
import SharePointService from "../../../shared/services/SharePointService";
import { InitialForm } from "./InitialForm";
import { AdditionalForm } from "./AdditionalForm";
import { BRANDING } from "../../../shared/theme/theme";
import { getSitePrefix } from "../../../shared/utils/getSitePrefix";

export interface ChangeRequestFormData {
  title: string;
  scopeOfChange: string;
  versionNumber?: string;
  documentNumber?: string;
  newDocument: boolean;
  externalDocument: boolean;
  departmentId: number | undefined;
  changeAuthority: SharePointPerson | undefined;
  documentId: number | undefined;
  attachments: File[];
  documentTypeId: number | undefined;
  documentCategoryIds: number[];
  classification: "Public" | "Internal" | "Confidential" | "Restricted" | "";
  audienceId: number | undefined;
  businessFunctionIds: number[];
  urgency: "Standard" | "Urgent" | "";
  releaseAuthority: SharePointPerson | undefined;
  author: SharePointPerson | undefined;
  reviewerIds: number[];
  contributorIds: number[];
  draftDocumentName: string;
}

interface ChangeRequestFormProps {
  preselectedDocumentId?: number;
}

const EMPTY_FORM: ChangeRequestFormData = {
  title: "",
  scopeOfChange: "",
  versionNumber: undefined,
  documentNumber: undefined,
  departmentId: undefined,
  newDocument: true,
  externalDocument: false,
  changeAuthority: undefined,
  documentId: undefined,
  attachments: [],
  documentTypeId: undefined,
  documentCategoryIds: [],
  classification: "",
  audienceId: undefined,
  businessFunctionIds: [],
  urgency: "Standard",
  releaseAuthority: undefined,
  author: undefined,
  reviewerIds: [],
  contributorIds: [],
  draftDocumentName: "",
};

type SubmissionStatus = "form" | "waiting" | "found" | "timeout";

// ─── Success Screen ───────────────────────────────────────────────────────────

const SuccessScreen: React.FC<{ cr: IChangeRequest; onSubmitAnother: () => void }> = ({ cr, onSubmitAnother }) => (
  <Box sx={{ minHeight: "100%", backgroundColor: "white", display: "flex", flexDirection: "column" }}>
    <Box sx={{ backgroundColor: BRANDING.primary, padding: "20px 24px" }}>
      <Typography sx={{ fontSize: "18px", fontWeight: 500, color: "white", mb: "4px" }}>
        Submit Change Request
      </Typography>
      <Typography sx={{ fontSize: "12px", color: "rgba(255,255,255,0.75)" }}>
        Request a new or updated document
      </Typography>
    </Box>

    <Box sx={{ padding: "24px", maxWidth: 640 }}>
      <Box sx={{ border: "0.5px solid #E2E8F0", borderRadius: "8px", overflow: "hidden", mb: 3 }}>
        <Box sx={{ padding: "16px 18px" }}>
          <Box sx={{ display: "flex", justifyContent: "space-between", alignItems: "center", mb: 1.25 }}>
            <Typography sx={{ fontSize: "11px", color: "#94A3B8", letterSpacing: "0.5px", fontWeight: 600 }}>
              {cr.ChangeRequestNumber}
            </Typography>
            <Typography
              sx={{ fontSize: "11px", fontWeight: 500, color: "#B5850A", backgroundColor: "#FEF3E2", padding: "3px 9px", borderRadius: "3px" }}
            >
              Awaiting approval
            </Typography>
          </Box>
          <Typography sx={{ fontSize: "15px", fontWeight: 500, color: "#1E293B", mb: 1.5 }}>
            {cr.Title}
          </Typography>
          <Box sx={{ borderTop: "0.5px solid #E2E8F0", pt: 1.25, display: "grid", gridTemplateColumns: "1fr 1fr", gap: "8px 16px" }}>
            <Box>
              <Typography sx={{ fontSize: "10px", color: "#94A3B8", textTransform: "uppercase", letterSpacing: "0.5px" }}>
                Change authority
              </Typography>
              <Typography sx={{ fontSize: "12px", color: "#1E293B", mt: "2px" }}>
                {cr.ChangeAuthority?.Title || "—"}
              </Typography>
            </Box>
            <Box>
              <Typography sx={{ fontSize: "10px", color: "#94A3B8", textTransform: "uppercase", letterSpacing: "0.5px" }}>
                Next step
              </Typography>
              <Typography sx={{ fontSize: "12px", color: "#1E293B", mt: "2px" }}>
                Awaiting Change Authority approval
              </Typography>
            </Box>
          </Box>
        </Box>
      </Box>

      <Box display="flex" gap={1.5}>
        <Button
          variant="contained"
          onClick={onSubmitAnother}
          sx={{
            padding: "6px 16px", fontSize: "13px", fontWeight: 500,
            backgroundColor: BRANDING.primary, borderRadius: "2px",
            textTransform: "none", boxShadow: "none",
            "&:hover": { backgroundColor: BRANDING.primaryDark, boxShadow: "none" },
          }}
        >
          Submit another request
        </Button>
        <Button
          onClick={() =>
            window.open(`${getSitePrefix()}/SitePages/Document-Portal.aspx`, "_blank")
          }
          sx={{ padding: "6px 16px", fontSize: "13px", fontWeight: 500, color: BRANDING.primary, textTransform: "none", "&:hover": { backgroundColor: "#F3F2F1" } }}
        >
          Go to Document Portal
        </Button>
      </Box>
    </Box>
  </Box>
);

const SubmittingScreen: React.FC<{ status: "waiting" | "timeout"; onBackToInbox: () => void }> = ({ status, onBackToInbox }) => (
  <Box sx={{ minHeight: "100%", backgroundColor: "white", display: "flex", flexDirection: "column" }}>
    <Box sx={{ backgroundColor: BRANDING.primary, padding: "20px 24px" }}>
      <Typography sx={{ fontSize: "18px", fontWeight: 500, color: "white", mb: "4px" }}>
        Submit Change Request
      </Typography>
      <Typography sx={{ fontSize: "12px", color: "rgba(255,255,255,0.75)" }}>
        Request a new or updated document
      </Typography>
    </Box>

    <Box sx={{ padding: "24px", maxWidth: 640 }}>
      <Box sx={{ display: "flex", alignItems: "flex-start", gap: 1.5, border: "0.5px solid #E2E8F0", borderRadius: "8px", padding: "14px 18px" }}>
        {status === "waiting" ? (
          <CircularProgress size={16} thickness={5} sx={{ color: BRANDING.primary, mt: "2px" }} />
        ) : (
          <Box sx={{ width: 16, height: 16, mt: "2px" }} />
        )}
        <Box>
          <Typography sx={{ fontSize: "13px", fontWeight: 500, color: "#1E293B", mb: 0.5 }}>
            {status === "waiting" ? "Submitting your change request…" : "Still processing"}
          </Typography>
          <Typography sx={{ fontSize: "12px", color: "#64748B", lineHeight: 1.5 }}>
            {status === "waiting"
              ? "Assigning a reference number. This usually takes a few seconds."
              : "Your request was received and is still being processed. You'll get an email with your reference number shortly."}
          </Typography>
        </Box>
      </Box>

      {status === "timeout" && (
        <Box mt={2}>
          <Button
            variant="contained"
            onClick={onBackToInbox}
            sx={{
              padding: "6px 16px", fontSize: "13px", fontWeight: 500,
              backgroundColor: BRANDING.primary, borderRadius: "2px",
              textTransform: "none", boxShadow: "none",
              "&:hover": { backgroundColor: BRANDING.primaryDark, boxShadow: "none" },
            }}
          >
            Go to Document Portal
          </Button>
        </Box>
      )}
    </Box>
  </Box>
);

// ─── Main Component ───────────────────────────────────────────────────────────

const ChangeRequestForm: React.FC<ChangeRequestFormProps> = ({
  preselectedDocumentId,
}) => {
  const [additionalDetailsOpen, setAdditionalDetailsOpen] = useState(false);
  const [submitError, setSubmitError] = useState<string | null>(null);
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [status, setStatus] = useState<SubmissionStatus>("form");
  const [createdCR, setCreatedCR] = useState<IChangeRequest | null>(null);
  const pollingRef = React.useRef<ReturnType<typeof setInterval> | null>(null);

  useEffect(() => {
    return () => {
      if (pollingRef.current) clearInterval(pollingRef.current);
    };
  }, []);

  const stopPolling = (): void => {
    if (pollingRef.current) {
      clearInterval(pollingRef.current);
      pollingRef.current = null;
    }
  };

  const documentTypeOptions = [
    { Id: 1, Title: "Policy" },
    { Id: 2, Title: "Procedure" },
    { Id: 3, Title: "Work Instruction" },
    { Id: 4, Title: "Form" },
  ];

  const mapDocumentTypeToId = (
    documentType?: string | null,
  ): number | undefined => {
    if (!documentType) return undefined;
    const match = documentTypeOptions.find(
      (type) => type.Title.toLowerCase() === documentType.toLowerCase(),
    );
    return match?.Id ?? undefined;
  };

  const [formData, setFormData] = useState<ChangeRequestFormData>(EMPTY_FORM);

  const isExistingDocumentSelected =
    !formData.newDocument && !!formData.documentId;
  const { documents } = useDocuments();
  const { departments } = useDepartments();

  // Auto-populate from existing document
  useEffect(() => {
    const autoPopulateFromDocument = async (): Promise<void> => {
      if (formData.newDocument || !formData.documentId) return;
      if (departments.length === 0) return;

      console.log("autoPopulate firing", {
        documentId: formData.documentId,
        departmentsLoaded: departments.length,
      });

      try {
        const selectedDoc = await SharePointService.getDocumentById(
          formData.documentId,
        );
        console.log("selectedDoc", selectedDoc);
        console.log("CoreFunctionality", selectedDoc?.CoreFunctionality);
        if (!selectedDoc) return;

        // Live lookup — get current Change Authority from dept config, not from the document
        const matchedDept = departments.find(
          (d) => d.Id === selectedDoc?.CoreFunctionality?.Id,
        );
        console.log("matchedDept", matchedDept);
        console.log("changeAuthority", matchedDept?.ChangeAuthority);

        setFormData((prev) => ({
          ...prev,
          departmentId: selectedDoc.CoreFunctionality?.Id,
          changeAuthority: matchedDept?.ChangeAuthority ?? undefined, // ← was selectedDoc.ChangeAuthority
          versionNumber: selectedDoc.VersionNumber ?? undefined,
          documentNumber: selectedDoc.DocumentNumber ?? undefined,
          businessFunctionIds:
            selectedDoc.BusinessFunction?.map((bf) => bf.Id) || [],
          documentCategoryIds: selectedDoc.Category?.map((dc) => dc.Id) || [],
          documentTypeId: mapDocumentTypeToId(selectedDoc.DocumentType?.Title),
          classification: selectedDoc.Classification || "",
          audienceId: selectedDoc.Audience?.Id,
          releaseAuthority: selectedDoc.ReleaseAuthority ?? undefined,
          author: selectedDoc.Author0 ?? undefined,
          draftDocumentName: selectedDoc.DocumentTitle || "",
        }));
      } catch (error) {
        console.error("Error auto-populating from document:", error);
      }
    };
    autoPopulateFromDocument().catch(console.error);
  }, [formData.documentId, formData.newDocument, departments]);

  // Reset fields when switching to New Document
  useEffect(() => {
    if (!formData.newDocument) return;
    setFormData((prev) => ({
      ...prev,
      documentId: undefined,
      versionNumber: undefined,
      documentNumber: undefined,
      departmentId: undefined,
      changeAuthority: undefined,
      externalDocument: false,
      businessFunctionIds: [],
      documentCategoryIds: [],
      documentTypeId: undefined,
      classification: "",
      audienceId: undefined,
      releaseAuthority: undefined,
      author: undefined,
      draftDocumentName: "",
    }));
  }, [formData.newDocument]);

  useEffect(() => {
    if (!preselectedDocumentId) return;
    setFormData((prev) => ({
      ...prev,
      newDocument: false,
      documentId: preselectedDocumentId,
    }));
  }, [preselectedDocumentId]);

  const handleFieldChange = (
    field: keyof ChangeRequestFormData,
    value: ChangeRequestFormData[keyof ChangeRequestFormData],
  ): void => {
    setFormData((prev) => ({ ...prev, [field]: value }));
  };

  const isFormValid = (): boolean => {
    return !!(
      formData.title &&
      formData.scopeOfChange &&
      formData.departmentId &&
      (formData.newDocument || formData.documentId)
    );
  };

  const handleSubmitAnother = (): void => {
    stopPolling();
    setStatus("form");
    setCreatedCR(null);
    setFormData({ ...EMPTY_FORM, newDocument: false });
    setAdditionalDetailsOpen(false);
    setSubmitError(null);
  };

  const handleSubmit = async (e: React.FormEvent): Promise<void> => {
    e.preventDefault();
    setSubmitError(null);

    if (!isFormValid()) {
      setSubmitError("Please fill in all required fields before submitting.");
      return;
    }

    setIsSubmitting(true);

    try {
      const payload = {
        Title: formData.title,
        ScopeofChange: formData.scopeOfChange,
        VersionNumber: formData.versionNumber || undefined,
        DocumentNumber: formData.documentNumber || undefined,
        NewDocument: formData.newDocument,
        ExternalDocument: formData.externalDocument,
        CoreFunctionalityId: formData.departmentId,
        ChangeAuthorityId: formData.changeAuthority?.Id,
        Urgency: formData.urgency || "Standard",
        Status: "Submitted",
        TargetDocumentId: formData.documentId || undefined,
        Classification: formData.classification || undefined,
        AudienceId: formData.audienceId || undefined,
        DraftDocumentName: formData.draftDocumentName || undefined,
        ReleaseAuthorityId: formData.releaseAuthority?.Id || undefined,
        Author0Id: formData.author?.Id || undefined,
        ReviewersId: formData.reviewerIds.length > 0 ? formData.reviewerIds : undefined,
        ContributorsId: formData.contributorIds.length > 0 ? formData.contributorIds : undefined,
        BusinessFunctionId: formData.businessFunctionIds.length > 0 ? formData.businessFunctionIds : undefined,
        CategoryId: formData.documentCategoryIds.length > 0 ? formData.documentCategoryIds : undefined,
        isCrComplete: !!(formData.releaseAuthority && formData.author),
      };

      const result = await SharePointService.createChangeRequest(payload);
      const resolvedId = (result as { Id?: number }).Id;

      if (!resolvedId)
        throw new Error("Unable to resolve the created Change Request ID.");

      if (formData.reviewerIds.length > 0 || formData.contributorIds.length > 0) {
        await SharePointService.createParticipant(resolvedId, formData.contributorIds, formData.reviewerIds);
      }

      if (formData.attachments.length > 0) {
        await SharePointService.uploadAttachments(resolvedId, formData.attachments, (current, total, fileName) => {
          console.log(`Uploading file ${current} of ${total}: ${fileName}`);
        });
      }

      setIsSubmitting(false);
      setStatus("waiting");

      let attempts = 0;
      const maxAttempts = 15;

      pollingRef.current = setInterval(() => {
        attempts++;
        SharePointService.getChangeRequestById(resolvedId)
          .then((freshCR) => {
            if (freshCR?.ChangeRequestNumber) {
              stopPolling();
              setCreatedCR(freshCR);
              setStatus("found");
            } else if (attempts >= maxAttempts) {
              stopPolling();
              setStatus("timeout");
            }
          })
          .catch(() => {
            if (attempts >= maxAttempts) {
              stopPolling();
              setStatus("timeout");
            }
          });
      }, 4000);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      console.error("Error:", error);
      setSubmitError(message);
      setIsSubmitting(false);
    }
  };

  // ── Success screen ──
  if (status === "waiting" || status === "timeout") {
    return (
      <Fade in timeout={200} appear>
        <Box sx={{ minHeight: "100%", backgroundColor: "white" }}>
          <SubmittingScreen status={status} onBackToInbox={handleSubmitAnother} />
        </Box>
      </Fade>
    );
  }

  if (status === "found" && createdCR) {
    return (
      <Fade in timeout={200} appear>
        <Box sx={{ minHeight: "100%", backgroundColor: "white" }}>
          <SuccessScreen cr={createdCR} onSubmitAnother={handleSubmitAnother} />
        </Box>
      </Fade>
    );
  }

  return (
    <Box sx={{ minHeight: "100%", backgroundColor: "white" }}>
      {/* Header */}
      <Box sx={{ backgroundColor: BRANDING.primary, padding: "20px 24px" }}>
        <Typography
          sx={{
            fontSize: "18px",
            fontWeight: 500,
            color: "white",
            marginBottom: "4px",
          }}
        >
          Submit Change Request
        </Typography>
        <Typography sx={{ fontSize: "12px", color: "rgba(255,255,255,0.75)" }}>
          Request a new or updated document
        </Typography>
      </Box>

      {/* Form */}
      <Box component="form" onSubmit={handleSubmit} sx={{ padding: "24px" }}>
        {submitError && (
          <Alert
            severity="error"
            sx={{ mb: 3 }}
            onClose={() => setSubmitError(null)}
          >
            {submitError}
          </Alert>
        )}

        <InitialForm
          data={formData}
          onChange={handleFieldChange}
          documents={documents}
          departments={departments}
          isExistingDocumentSelected={isExistingDocumentSelected}
        />

        {/* Additional Details */}
        <Box
          sx={{
            border: "1px solid #E2E8F0",
            borderRadius: "8px",
            marginBottom: "24px",
            overflow: "hidden",
          }}
        >
          <Box
            onClick={() => setAdditionalDetailsOpen(!additionalDetailsOpen)}
            sx={{
              padding: "14px 20px",
              backgroundColor: "#F8FAFC",
              display: "flex",
              justifyContent: "space-between",
              alignItems: "center",
              cursor: "pointer",
              borderBottom: additionalDetailsOpen
                ? "1px solid #E2E8F0"
                : "none",
              "&:hover": { backgroundColor: "#F1F5F9" },
            }}
          >
            <Box sx={{ display: "flex", alignItems: "center", gap: "10px" }}>
              <Typography
                sx={{
                  fontSize: "14px",
                  fontWeight: 500,
                  color: additionalDetailsOpen ? BRANDING.primary : "#475569",
                }}
              >
                Additional details
              </Typography>
              <Box
                component="span"
                sx={{
                  fontSize: "11px",
                  padding: "3px 8px",
                  backgroundColor: additionalDetailsOpen
                    ? "#E6F1FB"
                    : "#E2E8F0",
                  color: additionalDetailsOpen ? BRANDING.primary : "#64748B",
                  borderRadius: "4px",
                }}
              >
                Optional
              </Box>
            </Box>
            <Box sx={{ display: "flex", alignItems: "center", gap: "8px" }}>
              {!additionalDetailsOpen && (
                <Typography sx={{ fontSize: "12px", color: "#94A3B8" }}>
                  Classification, team, settings
                </Typography>
              )}
              {additionalDetailsOpen ? (
                <ExpandLessIcon
                  sx={{ color: BRANDING.primary, fontSize: 20 }}
                />
              ) : (
                <ExpandMoreIcon
                  sx={{ color: BRANDING.primary, fontSize: 20 }}
                />
              )}
            </Box>
          </Box>

          <Collapse in={additionalDetailsOpen}>
            <Box sx={{ padding: "20px" }}>
              <AdditionalForm
                data={formData}
                onChange={handleFieldChange}
                documents={documents}
                isExistingDocumentSelected={isExistingDocumentSelected}
              />
            </Box>
          </Collapse>
        </Box>

        <Button
          type="submit"
          variant="contained"
          fullWidth
          disabled={!isFormValid() || isSubmitting}
          sx={{
            padding: "14px",
            fontSize: "14px",
            fontWeight: 500,
            backgroundColor: BRANDING.primary,
            borderRadius: "6px",
            textTransform: "none",
            "&:hover": { backgroundColor: BRANDING.primaryDark },
            "&.Mui-disabled": { backgroundColor: "#CBD5E1", color: "white" },
          }}
        >
          {isSubmitting ? "Submitting..." : "Submit Change Request"}
        </Button>
      </Box>
    </Box>
  );
};

export default ChangeRequestForm;
