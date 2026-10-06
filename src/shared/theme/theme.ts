// src/shared/theme/theme.ts
import { alpha, createTheme } from '@mui/material/styles';

/*
TENANT CONFIG — edit only this section
Tenant-editable objects: BRANDING, TEXT, BORDER, SURFACE, STATUS, AVATAR_PALETTE,
CR_STATUS_CONFIG, FONT_SIZE, SHADOW, RADIUS, CHIP_PALETTE
*/
export const BRANDING = {
  primary: '#0F4C81',
  primaryLight: '#1565A8',
  primaryDark: '#0A3A63',
  primaryBg: '#E6F1FB',
  logoUrl: '',
  companyName: 'Contoso',
} as const;

export const TEXT = {
  primary: '#1E293B',
  secondary: '#64748B',
  muted: '#94A3B8',
  inverse: '#FFFFFF',
} as const;

export const BORDER = {
  default: '#E2E8F0',
  strong: '#CBD5E1',
  subtle: '#F1F5F9',
} as const;

export const SURFACE = {
  page: '#F8FAFC',
  paper: '#FFFFFF',
  subtle: '#F1F5F9',
  hover: '#F8FAFC',
} as const;

export const STATUS = {
  success: {
    main: '#107C10',
    bg: '#DFF6DD',
    text: '#0B6A0B',
    border: alpha('#107C10', 0.28),
  },
  warning: {
    main: '#FFB900',
    bg: '#FFF4CE',
    text: '#835B00',
    border: alpha('#835B00', 0.22),
  },
  error: {
    main: '#D13438',
    bg: '#FDE7E9',
    text: '#A4262C',
    border: alpha('#A4262C', 0.24),
  },
  info: {
    main: BRANDING.primary,
    bg: BRANDING.primaryBg,
    text: BRANDING.primaryDark,
    border: alpha(BRANDING.primary, 0.22),
  },
} as const;

export const AVATAR_PALETTE = [
  '#0F6CBD',
  '#107C10',
  '#5C2D91',
  '#C50F1F',
  '#986F0B',
  '#0099BC',
] as const;

type CrStatusStyle = {
  bg: string;
  color: string;
  dot: string;
};

const CR_STATUS_FALLBACK: CrStatusStyle = {
  bg: SURFACE.subtle,
  color: TEXT.secondary,
  dot: TEXT.muted,
};

export const CR_STATUS_CONFIG: Record<string, CrStatusStyle> = {
  Submitted: {
    bg: STATUS.info.bg,
    color: STATUS.info.text,
    dot: STATUS.info.main,
  },
  'CA Review': {
    bg: STATUS.warning.bg,
    color: STATUS.warning.text,
    dot: STATUS.warning.main,
  },
  Approved: {
    bg: STATUS.success.bg,
    color: STATUS.success.text,
    dot: STATUS.success.main,
  },
  'Document Creation': {
    bg: STATUS.info.bg,
    color: STATUS.info.text,
    dot: STATUS.info.main,
  },
  'Document Review': {
    bg: STATUS.warning.bg,
    color: STATUS.warning.text,
    dot: STATUS.warning.main,
  },
  'Ready for Publishing': {
    bg: STATUS.success.bg,
    color: STATUS.success.text,
    dot: STATUS.success.main,
  },
  Published: {
    bg: STATUS.success.bg,
    color: STATUS.success.text,
    dot: STATUS.success.main,
  },
  Rejected: {
    bg: STATUS.error.bg,
    color: STATUS.error.text,
    dot: STATUS.error.main,
  },
};

const CR_STATUS_LOOKUP: Record<string, CrStatusStyle> = Object.entries(CR_STATUS_CONFIG).reduce(
  (acc, [key, value]) => {
    acc[key.toLowerCase()] = value;
    return acc;
  },
  {} as Record<string, CrStatusStyle>
);

export const getCrStatusStyle = (status?: string): CrStatusStyle => {
  if (!status) return CR_STATUS_FALLBACK;
  return CR_STATUS_LOOKUP[status.trim().toLowerCase()] ?? CR_STATUS_FALLBACK;
};

export const FONT_SIZE = {
  xs: '10px',
  sm: '12px',
  md: '13px',
  base: '14px',
  lg: '16px',
  xl: '18px',
} as const;

export const SHADOW = {
  card: `0 1px 2px ${alpha(TEXT.primary, 0.08)}`,
  dropdown: `0 8px 24px ${alpha(TEXT.primary, 0.14)}`,
  modal: `0 16px 32px ${alpha(TEXT.primary, 0.18)}`,
} as const;

export const RADIUS = {
  chip: '4px',
  button: '6px',
  buttonFlat: '2px',
  card: '8px',
  modal: '10px',
} as const;

export const GREY = {
  50: SURFACE.page,
  100: SURFACE.subtle,
  200: BORDER.default,
  300: BORDER.strong,
  400: TEXT.muted,
  500: TEXT.secondary,
  600: '#475569',
  700: '#334155',
  800: TEXT.primary,
  900: '#0F172A',
} as const;

export type ChipColor = { bg: string; text: string };

export const CHIP_PALETTE: Record<string, ChipColor> = {
  blue: { bg: BRANDING.primaryBg, text: BRANDING.primary },
  green: { bg: '#E6F7F2', text: '#0D7D5F' },
  purple: { bg: '#F3E8FF', text: '#7C3AED' },
  amber: { bg: '#FEF3E2', text: '#92650A' },
  pink: { bg: '#FCE8EE', text: '#9E3A5A' },
  cyan: { bg: '#E0F7FA', text: '#0891B2' },
  red: { bg: '#FEE2E2', text: '#B91C1C' },
  redDark: { bg: '#FEE4E2', text: '#7F1D1D' },
  grey: { bg: SURFACE.subtle, text: TEXT.secondary },
};

// ── Domain → palette key maps ──────────────────────────────────────────────────
// Each of these just names which palette color a given value uses. To retheme
// the whole app, edit CHIP_PALETTE above — these maps rarely need to change.

const DOC_TYPE_KEYS: Record<string, keyof typeof CHIP_PALETTE> = {
  policy: 'blue',
  procedure: 'green',
  form: 'pink',
  certificate: 'amber',
  guide: 'grey',
  manual: 'purple',
  checklist: 'cyan',
  template: 'red',
  'work instruction': 'purple',
};

const CLASSIFICATION_KEYS: Record<string, keyof typeof CHIP_PALETTE> = {
  public: 'green',
  internal: 'amber',
  confidential: 'red',
  restricted: 'redDark',
};

const TASK_TYPE_KEYS: Record<string, keyof typeof CHIP_PALETTE> = {
  'ca review': 'blue',
  'change authority review': 'blue',
  'change authority approval': 'green',
  'compliance authority review': 'purple',
  'document controller review': 'purple',
  'final approval': 'green',
  'author review': 'cyan',
  'document change process': 'amber',
  'participant task': 'pink',
  'cr completion': 'grey',
  'cr info required': 'grey',
  'publish document': 'cyan',
  'publishing rejection review': 'red',
  'document review': 'blue',
};

// ── Public getters — use these everywhere instead of hardcoded color objects ──

export const getDocTypeColors = (type?: string): ChipColor => {
  if (!type) return CHIP_PALETTE.grey;
  const key = DOC_TYPE_KEYS[type.toLowerCase().trim()];
  return key ? CHIP_PALETTE[key] : CHIP_PALETTE.grey;
};

export const getClassificationColors = (classification?: string): ChipColor => {
  if (!classification) return CHIP_PALETTE.grey;
  const key = CLASSIFICATION_KEYS[classification.toLowerCase()];
  return key ? CHIP_PALETTE[key] : CHIP_PALETTE.grey;
};

export const getTaskTypeColors = (taskType?: string): ChipColor => {
  if (!taskType) return CHIP_PALETTE.grey;
  const key = TASK_TYPE_KEYS[taskType.toLowerCase().trim()];
  return key ? CHIP_PALETTE[key] : CHIP_PALETTE.grey;
};

// Backwards-compatible solid-color getters (only if something still needs a
// single hex rather than a bg/text pair)
export const getDocTypeColor = (docType?: string): string =>
  getDocTypeColors(docType).text;

export const getClassificationColor = (classification?: string): string =>
  getClassificationColors(classification).text;

// ══════════════════════════════════════════════════════════════════════════════
// Theme
// ══════════════════════════════════════════════════════════════════════════════
const theme = createTheme({
  palette: {
    primary: {
      main: BRANDING.primary,
      light: BRANDING.primaryLight,
      dark: BRANDING.primaryDark,
      contrastText: TEXT.inverse,
    },
    secondary: {
      main: TEXT.secondary,
      light: TEXT.muted,
      dark: GREY[600],
    },
    success: {
      main: STATUS.success.main,
      light: STATUS.success.bg,
      dark: STATUS.success.text,
      contrastText: TEXT.inverse,
    },
    warning: {
      main: STATUS.warning.main,
      light: STATUS.warning.bg,
      dark: STATUS.warning.text,
      contrastText: TEXT.primary,
    },
    error: {
      main: STATUS.error.main,
      light: STATUS.error.bg,
      dark: STATUS.error.text,
      contrastText: TEXT.inverse,
    },
    info: {
      main: STATUS.info.main,
      light: STATUS.info.bg,
      dark: STATUS.info.text,
      contrastText: TEXT.inverse,
    },
    background: {
      default: SURFACE.page,
      paper: SURFACE.paper,
    },
    text: {
      primary: TEXT.primary,
      secondary: TEXT.secondary,
    },
    divider: BORDER.default,
    grey: {
      50: GREY[50],
      100: GREY[100],
      200: GREY[200],
      300: GREY[300],
      400: GREY[400],
      500: GREY[500],
      600: GREY[600],
      700: GREY[700],
      800: GREY[800],
      900: GREY[900],
    },
  },

  typography: {
    fontFamily: '"Segoe UI", -apple-system, BlinkMacSystemFont, Roboto, "Helvetica Neue", Arial, sans-serif',
    fontSize: Number.parseInt(FONT_SIZE.base, 10),

    // Panel/page-level titles (e.g. webpart headers)
    h6: { fontSize: FONT_SIZE.xl, fontWeight: 500, lineHeight: 1.4, color: TEXT.primary },

    // Card/document titles (e.g. "Book policy" in DocumentDetail)
    subtitle1: { fontSize: FONT_SIZE.lg, fontWeight: 600, lineHeight: 1.4, color: TEXT.primary },

    // Section headers (e.g. "OVERVIEW", "DEPARTMENT")
    overline: {
      fontSize: FONT_SIZE.xs,
      fontWeight: 600,
      textTransform: 'uppercase',
      letterSpacing: '0.5px',
      color: TEXT.muted,
    },

    // Primary readable content — table cells, form values
    body2: { fontSize: FONT_SIZE.md, fontWeight: 400, color: TEXT.primary },

    // Secondary/meta text — dates, counts, muted descriptions
    caption: { fontSize: FONT_SIZE.sm, fontWeight: 400, color: TEXT.secondary },

    button: { fontSize: FONT_SIZE.md, fontWeight: 500, textTransform: 'none' },
  },

  shape: {
    borderRadius: Number.parseInt(RADIUS.card, 10),
  },

  components: {
    MuiButton: {
      styleOverrides: {
        root: {
          textTransform: 'none',
          fontWeight: 500,
          borderRadius: RADIUS.button,
          boxShadow: 'none',
          '&:hover': { boxShadow: 'none' },
        },
      },
    },
    MuiChip: {
      styleOverrides: {
        root: {
          borderRadius: RADIUS.chip,
          fontWeight: 500,
        },
      },
    },
    MuiDialog: {
      styleOverrides: {
        paper: {
          borderRadius: RADIUS.modal,
          boxShadow: SHADOW.modal,
        },
      },
    },
    MuiTableCell: {
      styleOverrides: {
        root: {
          borderBottom: `1px solid ${BORDER.subtle}`,
        },
      },
    },
  },
});

export default theme;