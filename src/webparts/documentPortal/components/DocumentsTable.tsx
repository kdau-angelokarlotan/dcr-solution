// src/webparts/documentPortal/components/DocumentsTable.tsx
import React, { useState, useMemo } from "react";
import {
  Box,
  Table,
  TableBody,
  TableCell,
  TableContainer,
  TableHead,
  TableRow,
  TableSortLabel,
  Typography,
  IconButton,
  Tooltip,
} from "@mui/material";
import OpenInNewIcon from "@mui/icons-material/OpenInNew";
import { Document } from "../../../shared/types/Document";
import { BRANDING, getDocTypeColors } from "../../../shared/theme/theme";
import { getSitePrefix } from "../../../shared/utils/getSitePrefix";
interface DocumentsTableProps {
  documents: Document[];
  onRowClick: (doc: Document) => void;
}

type SortKey = "DocumentTitle" | "DocumentType" | "PublishedDate" | "Modified";

const formatDate = (date?: string | Date): string => {
  if (!date) return "—";
  const d = typeof date === "string" ? new Date(date) : date;
  return d.toLocaleDateString("en-AU", {
    day: "2-digit",
    month: "short",
    year: "numeric",
  });
};

const formatRelativeTime = (date?: string | Date): string => {
  if (!date) return "—";
  const d = typeof date === "string" ? new Date(date) : date;
  const now = new Date();
  const diffMs = now.getTime() - d.getTime();
  const diffDays = Math.floor(diffMs / (1000 * 60 * 60 * 24));

  if (diffDays === 0) return "Today";
  if (diffDays === 1) return "Yesterday";
  if (diffDays < 7) return `${diffDays} days ago`;
  if (diffDays < 30) return `${Math.floor(diffDays / 7)} week${Math.floor(diffDays / 7) > 1 ? "s" : ""} ago`;
  return formatDate(d);
};

const MAX_VISIBLE_TAGS = 2;

const TagList: React.FC<{ items: string[] }> = ({ items }) => {
  if (items.length === 0) {
    return <Typography sx={{ fontSize: "12px", color: "#94A3B8" }}>—</Typography>;
  }
  const visible = items.slice(0, MAX_VISIBLE_TAGS);
  const hiddenCount = items.length - MAX_VISIBLE_TAGS;
  return (
    <Box sx={{ display: "flex", gap: 0.5, flexWrap: "wrap" }}>
      {visible.map((item) => (
        <Box
          key={item}
          component="span"
          sx={{
            fontSize: "10px",
            padding: "2px 7px",
            borderRadius: "3px",
            border: "1px solid #E2E8F0",
            color: "#64748B",
            backgroundColor: "#F8FAFC",
            whiteSpace: "nowrap",
          }}
        >
          {item}
        </Box>
      ))}
      {hiddenCount > 0 && (
        <Box
          component="span"
          sx={{
            fontSize: "10px",
            padding: "2px 7px",
            borderRadius: "3px",
            border: "1px solid #E2E8F0",
            color: "#94A3B8",
            backgroundColor: "#fff",
          }}
        >
          +{hiddenCount}
        </Box>
      )}
    </Box>
  );
};

const DocumentsTable: React.FC<DocumentsTableProps> = ({
  documents,
  onRowClick,
}) => {
  const sitePrefix = getSitePrefix();
  const [orderBy, setOrderBy] = useState<SortKey>("Modified");
  const [order, setOrder] = useState<"asc" | "desc">("desc");

  const handleSort = (column: SortKey): void => {
    if (orderBy === column) {
      setOrder(order === "asc" ? "desc" : "asc");
    } else {
      setOrderBy(column);
      setOrder("asc");
    }
  };

  const sortedDocuments = useMemo(() => {
    return [...documents].sort((a, b) => {
      let aVal: string;
      let bVal: string;

      switch (orderBy) {
        case "DocumentType":
          aVal = a.DocumentType?.Title ?? "";
          bVal = b.DocumentType?.Title ?? "";
          break;
        case "PublishedDate":
          aVal = a.PublishedDate ? new Date(a.PublishedDate).toISOString() : "";
          bVal = b.PublishedDate ? new Date(b.PublishedDate).toISOString() : "";
          break;
        case "Modified":
          aVal = a.Modified ? new Date(a.Modified).toISOString() : "";
          bVal = b.Modified ? new Date(b.Modified).toISOString() : "";
          break;
        default:
          aVal = a.DocumentTitle ?? "";
          bVal = b.DocumentTitle ?? "";
      }

      return order === "asc" ? aVal.localeCompare(bVal) : bVal.localeCompare(aVal);
    });
  }, [documents, orderBy, order]);

  // Header cell styles
  const headerCellSx = {
    backgroundColor: "#F8FAFC",
    color: BRANDING.primary,
    fontWeight: 600,
    fontSize: "10px",
    textTransform: "uppercase" as const,
    letterSpacing: "0.5px",
    padding: "10px 20px",
    borderBottom: "2px solid rgba(15, 76, 129, 0.12)",
    whiteSpace: "nowrap" as const,
  };

  const sortLabelSx = {
    color: `${BRANDING.primary} !important`,
    "&.Mui-active": {
      color: `${BRANDING.primary} !important`,
    },
    "& .MuiTableSortLabel-icon": {
      color: `${BRANDING.primary} !important`,
      opacity: 0.5,
    },
  };

  // Empty state
  if (sortedDocuments.length === 0) {
    return (
      <Box
        display="flex"
        flexDirection="column"
        alignItems="center"
        justifyContent="center"
        py={8}
        flex={1}
      >
        <Typography sx={{ color: "#64748B", fontSize: "14px", mb: 1 }}>
          No documents found
        </Typography>
        <Typography sx={{ color: "#94A3B8", fontSize: "12px" }}>
          Try adjusting your search or filters
        </Typography>
      </Box>
    );
  }

  return (
    <TableContainer sx={{ flex: 1 }}>
      <Table stickyHeader>
        <TableHead>
          <TableRow>
            <TableCell sx={{ ...headerCellSx, width: "90px" }}>Doc #</TableCell>
            <TableCell sx={{ ...headerCellSx, width: "35%" }}>
              <TableSortLabel
                active={orderBy === "DocumentTitle"}
                direction={orderBy === "DocumentTitle" ? order : "asc"}
                onClick={() => handleSort("DocumentTitle")}
                sx={sortLabelSx}
              >
                Document name
              </TableSortLabel>
            </TableCell>
            <TableCell sx={{ ...headerCellSx, width: "90px" }}>
              <TableSortLabel
                active={orderBy === "DocumentType"}
                direction={orderBy === "DocumentType" ? order : "asc"}
                onClick={() => handleSort("DocumentType")}
                sx={sortLabelSx}
              >
                Type
              </TableSortLabel>
            </TableCell>
            <TableCell sx={{ ...headerCellSx, width: "20%" }}>Function</TableCell>
            <TableCell sx={{ ...headerCellSx, width: "18%" }}>Category</TableCell>
            <TableCell sx={{ ...headerCellSx, width: "100px" }}>
              <TableSortLabel
                active={orderBy === "Modified"}
                direction={orderBy === "Modified" ? order : "desc"}
                onClick={() => handleSort("Modified")}
                sx={sortLabelSx}
              >
                Updated
              </TableSortLabel>
            </TableCell>
            <TableCell sx={{ ...headerCellSx, width: "100px" }}>
              <TableSortLabel
                active={orderBy === "PublishedDate"}
                direction={orderBy === "PublishedDate" ? order : "asc"}
                onClick={() => handleSort("PublishedDate")}
                sx={sortLabelSx}
              >
                Released
              </TableSortLabel>
            </TableCell>
            <TableCell sx={{ ...headerCellSx, width: "50px", textAlign: "center" }} />
          </TableRow>
        </TableHead>

        <TableBody>
          {sortedDocuments.map((doc) => {
            const typeColors = getDocTypeColors(doc.DocumentType?.Title);
            const functions = doc.BusinessFunction?.map((f) => f.Title) ?? [];
            const categories = doc.Category?.map((c) => c.Title) ?? [];

            return (
              <TableRow
                key={doc.Id}
                onClick={() => onRowClick(doc)}
                sx={{
                  cursor: "pointer",
                  backgroundColor: "white",
                  "&:hover": {
                    backgroundColor: "#F8FAFC",
                  },
                }}
              >
                <TableCell sx={{ padding: "14px 20px", color: "#64748B", fontSize: "12px", fontWeight: 500 }}>
                  {doc.DocumentNumber || "—"}
                </TableCell>

                {/* Document Name */}
                <TableCell
                  sx={{
                    padding: "14px 20px",
                    color: "#1E293B",
                    fontWeight: 500,
                    fontSize: "13px",
                  }}
                >
                  {doc.DocumentTitle}
                </TableCell>

                {/* Type Badge - Muted colors */}
                <TableCell sx={{ padding: "14px 20px" }}>
                  {doc.DocumentType?.Title ? (
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
                        textAlign: "center",
                      }}
                    >
                      {doc.DocumentType.Title}
                    </Box>
                  ) : (
                    <Typography sx={{ fontSize: "12px", color: "#94A3B8" }}>
                      —
                    </Typography>
                  )}
                </TableCell>

                {/* Function */}
                <TableCell sx={{ padding: "14px 20px" }}>
                  <TagList items={functions} />
                </TableCell>

                {/* Category */}
                <TableCell sx={{ padding: "14px 20px" }}>
                  <TagList items={categories} />
                </TableCell>

                {/* Updated */}
                <TableCell
                  sx={{
                    padding: "14px 20px",
                    color: "#64748B",
                    fontSize: "12px",
                  }}
                >
                  {formatRelativeTime(doc.Modified)}
                </TableCell>

                {/* Release Date */}
                <TableCell
                  sx={{
                    padding: "14px 20px",
                    color: "#64748B",
                    fontSize: "12px",
                  }}
                >
                  {formatDate(doc.PublishedDate)}
                </TableCell>

                {/* Action */}
                <TableCell sx={{ padding: "14px 20px", textAlign: "center" }}>
                  {(doc.PublishedFileUrl || doc.FileRef) && (
                    <Tooltip title="Open in new tab" arrow>
                      <IconButton
                        size="small"
                        onClick={(e) => {
                          e.stopPropagation();
                          const openUrl = doc.PublishedFileUrl
                            ? `${sitePrefix}/${doc.PublishedFileUrl}`
                            : `${window.location.origin}${doc.FileRef}`;
                          window.open(openUrl, "_blank");
                        }}
                        sx={{
                          color: "#94A3B8",
                          padding: "4px",
                          "&:hover": {
                            color: BRANDING.primary,
                            backgroundColor: "rgba(15, 76, 129, 0.08)",
                          },
                        }}
                      >
                        <OpenInNewIcon sx={{ fontSize: 18 }} />
                      </IconButton>
                    </Tooltip>
                  )}
                </TableCell>
              </TableRow>
            );
          })}
        </TableBody>
      </Table>
    </TableContainer>
  );
};

export default DocumentsTable;