import { TokenCache } from "@/types/onedrive";

/**
 * Server-side in-memory token store.
 * In production, replace this with a database or Redis-backed store.
 * For the initial integration test, in-memory is sufficient.
 */

let tokenCache: TokenCache | null = null;

export function getStoredToken(): TokenCache | null {
  if (tokenCache && tokenCache.expiresAt > Date.now()) {
    return tokenCache;
  }
  return null;
}

export function storeToken(token: TokenCache): void {
  tokenCache = token;
}

export function clearToken(): void {
  tokenCache = null;
}

export function isAuthenticated(): boolean {
  return getStoredToken() !== null;
}
