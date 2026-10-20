import { getGraphToken } from "../auth/auth";

const graphRoot = "https://graph.microsoft.com/v1.0";
const graphOrigin = "https://graph.microsoft.com";

export class GraphError extends Error {
  constructor(
    public readonly status: number,
    public readonly responseBody: string,
  ) {
    super(`Microsoft Graph returned HTTP ${status}`);
  }
}

export async function graphRequest<T>(
  pathOrUrl: string,
  init: RequestInit = {},
): Promise<T> {
  const token = await getGraphToken();
  const url = buildGraphUrl(pathOrUrl);
  const headers = new Headers(init.headers);

  headers.set("Authorization", `Bearer ${token}`);
  headers.set("Accept", "application/json");

  if (init.body && !headers.has("Content-Type")) {
    headers.set("Content-Type", "application/json");
  }

  const response = await fetch(url, {
    ...init,
    headers,
  });

  if (!response.ok) {
    throw new GraphError(response.status, await response.text());
  }

  if (response.status === 204) {
    return undefined as T;
  }

  return response.json() as Promise<T>;
}

export function buildGraphUrl(pathOrUrl: string): string {
  if (pathOrUrl.startsWith("/")) {
    return `${graphRoot}${pathOrUrl}`;
  }

  let url: URL;

  try {
    url = new URL(pathOrUrl);
  } catch {
    throw new Error("Microsoft Graph URL is invalid.");
  }

  if (
    url.origin !== graphOrigin ||
    !url.pathname.startsWith("/v1.0/") ||
    url.username ||
    url.password ||
    url.hash
  ) {
    throw new Error("Microsoft Graph URL is outside the allowed API boundary.");
  }

  return url.toString();
}

const transientStatuses = new Set([429, 503, 504]);
const maximumRetryDelayMs = 5000;

function isTransientGraphError(error: unknown): boolean {
  return error instanceof GraphError && transientStatuses.has(error.status);
}

// Retry a Graph call a bounded number of times when Graph reports throttling or
// a transient outage. Mailbox-wide scans and bulk moves issue many sequential
// requests and are the operations most likely to be throttled; a single 429
// must not be allowed to abort them.
export async function withGraphRetry<T>(
  operation: () => Promise<T>,
  maximumRetries = 3,
): Promise<T> {
  for (let attempt = 0; ; attempt += 1) {
    try {
      return await operation();
    } catch (error) {
      if (attempt >= maximumRetries || !isTransientGraphError(error)) {
        throw error;
      }

      const backoffMs = Math.min(2 ** attempt * 250, maximumRetryDelayMs);
      await new Promise((resolve) => setTimeout(resolve, backoffMs));
    }
  }
}
