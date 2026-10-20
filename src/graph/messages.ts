import { graphRequest, withGraphRetry } from "./client";

interface Message {
  id: string;
  subject?: string;
  parentFolderId?: string;
}

export interface MessageCollection {
  value: Message[];
  "@odata.nextLink"?: string;
}

export type InboxScanStatus = "complete" | "truncated" | "failed";

export interface InboxScanResult {
  messageIds: string[];
  // complete  - confident every matching Inbox message was found
  // truncated - only the most recent messages were scanned; older case mail may
  //             remain. Deterministic on a large Inbox, so retrying is pointless.
  // failed    - discovery errored (e.g. throttling). Retrying may well succeed.
  status: InboxScanStatus;
  scannedCount: number;
}

export async function moveMessage(
  messageId: string,
  destinationFolderId: string,
): Promise<Message> {
  return withGraphRetry(() =>
    graphRequest<Message>(
      `/me/messages/${encodeURIComponent(messageId)}/move`,
      {
        method: "POST",
        body: JSON.stringify({
          destinationId: destinationFolderId,
        }),
      },
    ),
  );
}

export type MessagePageFetcher = (url: string) => Promise<MessageCollection>;

const graphMessagePageFetcher: MessagePageFetcher = (url) =>
  withGraphRetry(() => graphRequest<MessageCollection>(url));

function matchesCase(
  message: Message,
  needle: string,
  excludeId?: string,
): boolean {
  return (
    message.id !== excludeId &&
    !!message.subject &&
    message.subject.includes(needle)
  );
}

// Walk the newest Inbox messages directly. This is the authoritative matcher:
// it compares the raw subject with an exact substring, so a 16-digit main case
// number also catches 19-digit collaboration subjects, and it sees messages that
// the search index has not picked up yet (case mail is often seconds old).
// Bounded by maxScan so Track stays fast on a large mailbox.
async function scanRecentInbox(
  needle: string,
  excludeId: string | undefined,
  maxScan: number,
  fetchPage: MessagePageFetcher,
): Promise<{ messageIds: string[]; reachedEnd: boolean; failed: boolean }> {
  let url: string | undefined =
    "/me/mailFolders/inbox/messages" +
    "?$select=id,subject&$top=50&$orderby=receivedDateTime desc";
  const messageIds: string[] = [];
  let scanned = 0;

  while (url && scanned < maxScan) {
    let page: MessageCollection;

    try {
      page = await fetchPage(url);
    } catch {
      return { messageIds, reachedEnd: false, failed: true };
    }

    for (const message of page.value ?? []) {
      scanned += 1;
      if (matchesCase(message, needle, excludeId)) {
        messageIds.push(message.id);
      }
    }

    url = page["@odata.nextLink"];
  }

  return { messageIds, reachedEnd: !url, failed: false };
}

// Ask Exchange to find the case number anywhere in the Inbox. This removes the
// depth limit of the recent scan: case mail buried thousands of messages deep is
// still found in a single request. Every hit is re-checked against the same
// subject-substring predicate, so search can only ever ADD recall - it cannot
// introduce a false positive from a body match or a fuzzy term expansion.
// Note: Graph rejects $orderby together with $search on messages.
async function searchInbox(
  needle: string,
  excludeId: string | undefined,
  maxPages: number,
  fetchPage: MessagePageFetcher,
): Promise<{ messageIds: string[]; ok: boolean }> {
  let url: string | undefined =
    "/me/mailFolders/inbox/messages" +
    `?$select=id,subject&$top=50&$search=%22${encodeURIComponent(needle)}%22`;
  const messageIds: string[] = [];

  for (let pageCount = 0; url && pageCount < maxPages; pageCount += 1) {
    let page: MessageCollection;

    try {
      page = await fetchPage(url);
    } catch {
      return { messageIds, ok: false };
    }

    for (const message of page.value ?? []) {
      if (matchesCase(message, needle, excludeId)) {
        messageIds.push(message.id);
      }
    }

    url = page["@odata.nextLink"];
  }

  return { messageIds, ok: true };
}

// Find Inbox messages whose subject contains the given case number, so Track can
// sweep case mail that arrived before the Exchange rule existed (a rule only
// routes future arrivals).
//
// Two independent passes are combined: a direct scan of the newest messages and
// a server-side search of the whole Inbox. Neither alone is sufficient - the
// direct scan is depth-limited, and the search index can lag brand-new mail - so
// their results are merged. Matches found before a failure are always kept
// rather than discarded; silently losing them is what let assignment mail pile
// up unrouted.
export async function findInboxMessagesBySubject(
  needle: string,
  excludeId?: string,
  maxScan = 250,
  fetchPage: MessagePageFetcher = graphMessagePageFetcher,
): Promise<InboxScanResult> {
  const recent = await scanRecentInbox(needle, excludeId, maxScan, fetchPage);
  const searched = await searchInbox(needle, excludeId, 5, fetchPage);

  const messageIds = [
    ...new Set([...recent.messageIds, ...searched.messageIds]),
  ];

  let status: InboxScanStatus;

  if (searched.ok) {
    // The search covered the whole Inbox, so the recent scan's cap is not a
    // coverage gap and must not be reported as one.
    status = "complete";
  } else if (recent.failed) {
    status = "failed";
  } else {
    status = recent.reachedEnd ? "complete" : "truncated";
  }

  return { messageIds, status, scannedCount: maxScan };
}
