// src/webparts/documentPortal/components/DocumentDetail.tsx
import React, { useState } from "react";
import {
  Box,
  Typography,
  Button,
  Chip,
  CircularProgress,
} from "@mui/material";
import ArrowBackIcon from "@mui/icons-material/ArrowBack";
import DownloadIcon from "@mui/icons-material/Download";
import OpenInNewIcon from "@mui/icons-material/OpenInNew";
import { Document } from "../../../shared/types/Document";
import { BRANDING, getDocTypeColors, getClassificationColors } from "../../../shared/theme/theme";
import { getSitePrefix } from "../../../shared/utils/getSitePrefix";

interface DocumentDetailProps {
  document: Document;
  onBack: () => void;
}

const DocumentDetail: React.FC<DocumentDetailProps> = ({
  document,
  onBack,
}) => {
  const [pdfLoading, setPdfLoading] = useState(true);

  const sitePrefix = getSitePrefix();

  // Viewer always uses PDF
  const viewerUrl = document.PublishedFileUrl
    ? `${sitePrefix}/${document.PublishedFileUrl}`
    : document.FileRef
      ? `${window.location.origin}${document.FileRef}`
      : null;

  // Download respects DownloadFormat
  const downloadUrl = document.DownloadFormat === "Original" && document.SourceFileUrl
    ? `${sitePrefix}/${document.SourceFileUrl}`
    : viewerUrl;

  const downloadLabel = document.DownloadFormat === "Original"
    ? "Download Word"
    : "Download PDF";

  const typeColors = getDocTypeColors(document.DocumentType?.Title);
  const classificationColors = getClassificationColors(document.Classification);

  const formatDate = (date?: Date | string): string => {
    if (!date) return "—";
    const d = typeof date === "string" ? new Date(date) : date;
    return d.toLocaleDateString("en-AU", {
      day: "2-digit",
      month: "long",
      year: "numeric",
    });
  };

  const handleDownload = (): void => {
    if (downloadUrl) {
      window.open(downloadUrl, "_blank");
    }
  };

  const handleOpenInTab = (): void => {
    if (viewerUrl) {
      window.open(viewerUrl, "_blank");
    }
  };

  const handleStartChangeRequest = (): void => {
    window.location.href = `${sitePrefix}/SitePages/Submit-Change-Request.aspx?documentId=${document.Id}`;
  };

  // Label styles for metadata
  const labelSx = {
    fontSize: "10px",
    color: BRANDING.primary,
    textTransform: "uppercase" as const,
    letterSpacing: "0.5px",
    fontWeight: 600,
    marginBottom: "6px",
  };

  const valueSx = {
    fontSize: "13px",
    color: "#1E293B",
    margin: 0,
  };

  return (
    <Box
      sx={{
        height: "100%",
        display: "flex",
        flexDirection: "column",
        backgroundColor: "white",
      }}
    >
      {/* Flat Blue Header with Breadcrumb & Actions */}
      <Box
        sx={{
          backgroundColor: BRANDING.primary,
          padding: "12px 24px",
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
        }}
      >
        {/* Left side: Back + Breadcrumb */}
        <Box display="flex" alignItems="center" gap={1.5}>
          <Button
            onClick={onBack}
            startIcon={<ArrowBackIcon sx={{ fontSize: "16px !important" }} />}
            sx={{
              padding: "6px 12px",
              borderRadius: "6px",
              fontSize: "12px",
              backgroundColor: "rgba(255,255,255,0.15)",
              color: "white",
              textTransform: "none",
              fontWeight: 400,
              "&:hover": { backgroundColor: "rgba(255,255,255,0.25)" },
            }}
          >
            Back
          </Button>

          <Box
            sx={{
              width: "1px",
              height: "20px",
              backgroundColor: "rgba(255,255,255,0.3)",
            }}
          />

          {/* Breadcrumb */}
          <Box
            display="flex"
            alignItems="center"
            gap={0.75}
            sx={{ fontSize: "12px" }}
          >
            <Typography
              component="span"
              sx={{ fontSize: "12px", color: "rgba(255,255,255,0.8)" }}
            >
              Documents
            </Typography>
            <Typography
              component="span"
              sx={{ fontSize: "12px", color: "rgba(255,255,255,0.5)" }}
            >
              /
            </Typography>
            {document.DocumentType?.Title && (
              <>
                <Typography
                  component="span"
                  sx={{ fontSize: "12px", color: "rgba(255,255,255,0.8)" }}
                >
                  {document.DocumentType.Title}
                </Typography>
                <Typography
                  component="span"
                  sx={{ fontSize: "12px", color: "rgba(255,255,255,0.5)" }}
                >
                  /
                </Typography>
              </>
            )}
            <Typography
              component="span"
              sx={{
                fontSize: "12px",
                color: "white",
                fontWeight: 500,
                maxWidth: "300px",
                overflow: "hidden",
                textOverflow: "ellipsis",
                whiteSpace: "nowrap",
              }}
            >
              {document.DocumentTitle}
            </Typography>
          </Box>
        </Box>

        {/* Right side: Actions */}
        <Box display="flex" gap={1}>
          <Button
            onClick={handleDownload}
            startIcon={<DownloadIcon sx={{ fontSize: "14px !important" }} />}
            sx={{
              padding: "6px 14px",
              fontSize: "12px",
              backgroundColor: "#fff",
              color: BRANDING.primary,
              borderRadius: "6px",
              textTransform: "none",
              fontWeight: 500,
              "&:hover": { backgroundColor: "#EFF6FC" },
            }}
          >
            {downloadLabel}
          </Button>
          <Button
            onClick={handleOpenInTab}
            startIcon={<OpenInNewIcon sx={{ fontSize: "14px !important" }} />}
            sx={{
              padding: "6px 14px",
              fontSize: "12px",
              backgroundColor: "#fff",
              color: BRANDING.primary,
              borderRadius: "6px",
              textTransform: "none",
              fontWeight: 400,
              "&:hover": { backgroundColor: "#EFF6FC" },
            }}
          >
            Open in tab
          </Button>
          <Button
            onClick={handleStartChangeRequest}
            sx={{
              padding: "6px 14px",
              fontSize: "12px",
              backgroundColor: "#0D7D5F",
              color: "white",
              borderRadius: "6px",
              textTransform: "none",
              fontWeight: 400,
              "&:hover": { backgroundColor: "#0B6A4E" },
            }}
          >
            Start Change Request
          </Button>
        </Box>
      </Box>

      {/* Content: Info Panel + PDF Viewer */}
      <Box sx={{ flex: 1, display: "flex", overflow: "hidden" }}>
        {/* Info Panel */}
        <Box
          sx={{
            width: 300,
            borderRight: "1px solid #E2E8F0",
            backgroundColor: "#F8FAFC",
            padding: "24px",
            overflow: "auto",
          }}
        >
          {/* Document Title */}
          <Box sx={{ marginBottom: "24px" }}>
            {document.DocumentType?.Title && (
              <Box
                component="span"
                sx={{
                  display: "inline-block",
                  fontSize: "11px",
                  fontWeight: 500,
                  padding: "4px 10px",
                  borderRadius: "4px",
                  backgroundColor: typeColors.bg,
                  color: typeColors.text,
                  marginBottom: "12px",
                }}
              >
                {document.DocumentType.Title}
              </Box>
            )}
            <Typography
              sx={{
                fontSize: "16px",
                fontWeight: 600,
                color: "#1E293B",
                lineHeight: 1.4,
              }}
            >
              {document.DocumentTitle}
            </Typography>
          </Box>

          {/* Metadata Fields */}
          <Box display="flex" flexDirection="column" gap={2.5}>
            {/* Document Number + Version */}
            <Box sx={{ display: "flex", gap: 2 }}>
              {document.DocumentNumber && (
                <Box sx={{ flex: 1 }}>
                  <Typography sx={labelSx}>Document number</Typography>
                  <Typography sx={valueSx}>{document.DocumentNumber}</Typography>
                </Box>
              )}
              {document.VersionNumber && (
                <Box sx={{ flex: 1 }}>
                  <Typography sx={labelSx}>Version</Typography>
                  <Typography sx={valueSx}>{document.VersionNumber}</Typography>
                </Box>
              )}
            </Box>

            {/* Category */}
            {document.Category && document.Category.length > 0 && (
              <Box>
                <Typography sx={labelSx}>Category</Typography>
                <Typography sx={valueSx}>
                  {document.Category.map((c) => c.Title).join(", ")}
                </Typography>
              </Box>
            )}

            {/* Department */}
            {document.CoreFunctionality?.Title && (
              <Box>
                <Typography sx={labelSx}>Department</Typography>
                <Typography sx={valueSx}>
                  {document.CoreFunctionality.Title}
                </Typography>
              </Box>
            )}

            {/* Functions */}
            {document.BusinessFunction && document.BusinessFunction.length > 0 && (
              <Box>
                <Typography sx={labelSx}>Functions</Typography>
                <Box display="flex" gap={0.75} flexWrap="wrap">
                  {document.BusinessFunction.map((func) => (
                    <Chip
                      key={func.Id}
                      label={func.Title}
                      size="small"
                      sx={{
                        fontSize: "11px",
                        height: "26px",
                        backgroundColor: "white",
                        border: "1px solid #E2E8F0",
                        color: "#475569",
                        fontWeight: 400,
                      }}
                    />
                  ))}
                </Box>
              </Box>
            )}

            {/* Classification */}
            {document.Classification && (
              <Box>
                <Typography sx={labelSx}>Classification</Typography>
                <Chip
                  label={document.Classification}
                  size="small"
                  sx={{
                    fontSize: "12px",
                    height: "26px",
                    fontWeight: 500,
                    backgroundColor: classificationColors.bg,
                    color: classificationColors.text,
                  }}
                />
              </Box>
            )}

            {/* Last Updated + Released */}
            <Box sx={{ display: "flex", gap: 2 }}>
              {document.Modified && (
                <Box sx={{ flex: 1 }}>
                  <Typography sx={labelSx}>Last updated</Typography>
                  <Typography sx={valueSx}>{formatDate(document.Modified)}</Typography>
                </Box>
              )}
              {document.PublishedDate && (
                <Box sx={{ flex: 1 }}>
                  <Typography sx={labelSx}>Released</Typography>
                  <Typography sx={valueSx}>{formatDate(document.PublishedDate)}</Typography>
                </Box>
              )}
            </Box>

            {/* File */}
            {document.FileLeafRef && (
              <Box>
                <Typography sx={labelSx}>File</Typography>
                <Typography
                  sx={{
                    fontSize: "12px",
                    color: "#64748B",
                    wordBreak: "break-all",
                  }}
                >
                  {document.FileLeafRef}
                </Typography>
              </Box>
            )}
          </Box>
        </Box>

        {/* PDF Viewer */}
        <Box
          sx={{
            flex: 1,
            backgroundColor: "#F1F5F9",
            position: "relative",
          }}
        >
          {/* Loading state */}
          {pdfLoading && viewerUrl && (
            <Box
              sx={{
                position: "absolute",
                inset: 0,
                display: "flex",
                flexDirection: "column",
                alignItems: "center",
                justifyContent: "center",
                backgroundColor: "#F1F5F9",
                zIndex: 5,
                gap: 2,
              }}
            >
              <CircularProgress size={32} sx={{ color: BRANDING.primary }} />
              <Typography sx={{ fontSize: "13px", color: "#64748B" }}>
                Loading document...
              </Typography>
            </Box>
          )}

          {viewerUrl ? (
            <iframe
              src={viewerUrl}
              width="100%"
              height="100%"
              title="Document preview"
              style={{ border: "none" }}
              onLoad={() => setPdfLoading(false)}
            />
          ) : (
            <Box
              sx={{
                display: "flex",
                flexDirection: "column",
                alignItems: "center",
                justifyContent: "center",
                height: "100%",
                gap: 2,
              }}
            >
              <Box
                sx={{
                  width: 80,
                  height: 100,
                  backgroundColor: "white",
                  border: "1px solid #E2E8F0",
                  borderRadius: "4px",
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "center",
                  boxShadow: "0 2px 8px rgba(0,0,0,0.04)",
                }}
              >
                <Typography sx={{ fontSize: "28px", color: "#CBD5E1" }}>
                  ☰
                </Typography>
              </Box>
              <Typography sx={{ fontSize: "13px", color: "#64748B" }}>
                Document preview not available
              </Typography>
            </Box>
          )}
        </Box>
      </Box>
    </Box>
  );
};

export default DocumentDetail;