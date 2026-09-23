export function getSitePrefix(): string {
  return `${window.location.origin}${window.location.pathname.split("/SitePages/")[0]}`;
}