import type { SheetDataResponse } from "@/types/onedrive";

/**
 * Simple in-memory cache for parsed spreadsheet data.
 * Avoids re-downloading and re-parsing the .xlsx on every request.
 *
 * TTL: 30 seconds — stale data is at most 30s old.
 * On Vercel serverless, each instance gets its own cache.
 */

const TTL = 30_000;

let cache: { data: SheetDataResponse; ts: number } | null = null;

export function getCachedData(): SheetDataResponse | null {
  if (cache && Date.now() - cache.ts < TTL) return cache.data;
  return null;
}

export function setCachedData(data: SheetDataResponse): void {
  cache = { data, ts: Date.now() };
}

export function getCacheAge(): number | null {
  if (!cache) return null;
  return Date.now() - cache.ts;
}
