export function isMobileOutlook(platform=globalThis.Office?.context?.platform) {
  return platform==='Android' || platform==='iOS';
}
