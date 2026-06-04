import type { AppMetadata } from "./types";

export const APP_DISPLAY_NAME = "时代亿信智能体（文档生成）";
export const APP_FALLBACK_VERSION = "0.0.0";

export function formatAppTitle(displayName: string, version: string): string {
  return `${displayName} v${version}`;
}

export function createAppMetadata(version: string | undefined): AppMetadata {
  const normalizedVersion = version?.trim() || APP_FALLBACK_VERSION;
  return {
    displayName: APP_DISPLAY_NAME,
    version: normalizedVersion,
    title: formatAppTitle(APP_DISPLAY_NAME, normalizedVersion)
  };
}

export const FALLBACK_APP_METADATA = createAppMetadata(APP_FALLBACK_VERSION);
