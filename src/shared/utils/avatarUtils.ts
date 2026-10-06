// src/shared/utils/avatarUtils.ts
// Single source of truth for avatar colours — used by both PeoplePicker and MultiPeoplePicker
import { AVATAR_PALETTE } from '../theme/theme';

export const getAvatarColor = (name: string): string =>
  AVATAR_PALETTE[name.charCodeAt(0) % AVATAR_PALETTE.length];

export const getAvatarInitials = (name: string): string =>
  name
    .split(" ")
    .map((n) => n[0])
    .slice(0, 2)
    .join("")
    .toUpperCase();