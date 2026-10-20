import {
  deleteCaseFolder,
  ensureActiveCaseFolder,
  getCaseFolderState,
  moveCaseFolder,
  type CaseFolderState,
  type CaseLocation,
  type MailFolder,
} from "../graph/folders";
import {
  moveMessage,
  findInboxMessagesBySubject,
  type InboxScanResult,
  type InboxScanStatus,
} from "../graph/messages";
import {
  deleteCaseRule,
  ensureCaseRule,
  findCaseRule,
  setCaseRuleEnabled,
  updateCaseRuleTarget,
  type MessageRule,
} from "../graph/rules";

export interface CaseMailboxOperations {
  ensureActiveFolder(trackingId: string): Promise<MailFolder>;
  getFolderState(trackingId: string): Promise<CaseFolderState>;
  moveFolder(
    folderId: string,
    destination: Exclude<CaseLocation, "untracked">,
  ): Promise<MailFolder>;
  // Moves the folder, and every message in it, to Deleted Items.
  deleteFolder(folder: MailFolder): Promise<void>;
  moveMessage(messageId: string, folderId: string): Promise<{ id: string }>;
  findInboxMessages(
    trackingId: string,
    excludeId: string,
  ): Promise<InboxScanResult>;
  ensureRule(trackingId: string, folderId: string): Promise<MessageRule>;
  findRule(trackingId: string): Promise<MessageRule | null>;
  setRuleEnabled(ruleId: string, enabled: boolean): Promise<MessageRule>;
  updateRuleTarget(
    ruleId: string,
    trackingId: string,
    folderId: string,
  ): Promise<MessageRule>;
  deleteRule(ruleId: string): Promise<void>;
}

export interface TrackCaseResult {
  trackingId: string;
  folderId: string;
  ruleId: string;
  movedMessageId: string;
  sweptMessageCount: number;
  // Messages that matched the case but could not be moved.
  unsweptMessageCount: number;
  sweepStatus: InboxScanStatus;
  // How many recent messages the direct scan was allowed to walk, so a
  // "truncated" result can say what was actually covered.
  sweepScannedCount: number;
}

export interface CaseActionResult {
  trackingId: string;
  location: Exclude<CaseLocation, "untracked">;
  folderId: string;
}

export interface DeleteCaseResult {
  trackingId: string;
  folderName: string;
}

export interface CaseStatus {
  location: CaseLocation;
  routing:
    | "enabled"
    | "disabled"
    | "missing"
    | "mistargeted"
    | "not-applicable";
}

export const graphCaseMailboxOperations: CaseMailboxOperations = {
  ensureActiveFolder: ensureActiveCaseFolder,
  getFolderState: getCaseFolderState,
  moveFolder: moveCaseFolder,
  deleteFolder: deleteCaseFolder,
  moveMessage,
  ensureRule: ensureCaseRule,
  findRule: findCaseRule,
  setRuleEnabled: setCaseRuleEnabled,
  updateRuleTarget: updateCaseRuleTarget,
  deleteRule: deleteCaseRule,
  findInboxMessages: (trackingId, excludeId) =>
    findInboxMessagesBySubject(trackingId, excludeId),
};

export async function trackCase(
  trackingId: string,
  messageId: string,
  operations: CaseMailboxOperations = graphCaseMailboxOperations,
): Promise<TrackCaseResult> {
  const state = await operations.getFolderState(trackingId);

  if (state.location === "archived") {
    throw new Error(
      `Case ${trackingId} is archived. Reopen it before tracking new messages.`,
    );
  }

  const folder =
    state.folder ?? (await operations.ensureActiveFolder(trackingId));
  const rule = await operations.ensureRule(trackingId, folder.id);
  const movedMessage = await operations.moveMessage(messageId, folder.id);

  try {
    await operations.setRuleEnabled(rule.id, true);
  } catch (error) {
    throw new Error(
      `The message was moved, but routing could not be enabled for ${trackingId}. Use Repair routing before continuing.`,
      { cause: error },
    );
  }

  // Sweep any case email already sitting in the Inbox into the case folder. The
  // Exchange rule only routes future arrivals, so without this, messages that
  // arrived before Track would be left behind. A sweep problem must not fail the
  // Track itself, since the folder, move, and rule already succeeded - but it
  // must never be hidden either. Silently reporting zero swept messages is
  // indistinguishable from a clean sweep, which previously let assignment and
  // notification mail pile up unrouted in the Inbox without any warning.
  let sweptMessageCount = 0;
  let unsweptMessageCount = 0;
  let sweepStatus: InboxScanStatus = "complete";
  let sweepScannedCount = 0;

  try {
    const scan = await operations.findInboxMessages(
      trackingId,
      movedMessage.id,
    );
    sweepStatus = scan.status;
    sweepScannedCount = scan.scannedCount;

    for (const pendingId of scan.messageIds) {
      try {
        await operations.moveMessage(pendingId, folder.id);
        sweptMessageCount += 1;
      } catch {
        unsweptMessageCount += 1;
      }
    }
  } catch {
    sweepStatus = "failed";
  }

  return {
    trackingId,
    folderId: folder.id,
    ruleId: rule.id,
    movedMessageId: movedMessage.id,
    sweptMessageCount,
    unsweptMessageCount,
    sweepStatus,
    sweepScannedCount,
  };
}

function requireActiveCaseFolder(
  trackingId: string,
  state: CaseFolderState,
): MailFolder {
  if (state.location === "untracked" || !state.folder) {
    throw new Error(`Case ${trackingId} is not tracked.`);
  }

  if (state.location !== "active") {
    throw new Error(`Case ${trackingId} is archived, not active.`);
  }

  return state.folder;
}

// Archive and delete both retire an active case: routing is stopped before the
// folder moves so no new mail lands in it mid-move, and the managed rule is
// removed afterwards so it can never route mail into the new location.
async function retireActiveCase<T>(
  trackingId: string,
  operations: CaseMailboxOperations,
  moveFolder: (folder: MailFolder) => Promise<T>,
  verb: { gerund: string; past: string },
): Promise<{ folder: MailFolder; moved: T }> {
  const state = await operations.getFolderState(trackingId);
  const folder = requireActiveCaseFolder(trackingId, state);
  const rule = await operations.findRule(trackingId);
  const wasEnabled = rule?.isEnabled ?? false;

  if (rule?.isEnabled) {
    await operations.setRuleEnabled(rule.id, false);
  }

  let moved: T;

  try {
    moved = await moveFolder(folder);
  } catch (error) {
    if (rule && wasEnabled) {
      try {
        await operations.setRuleEnabled(rule.id, true);
      } catch (rollbackError) {
        throw new AggregateError(
          [error, rollbackError],
          `${verb.gerund} ${trackingId} failed and its routing rule could not be re-enabled. Use Repair routing.`,
        );
      }
    }

    throw error;
  }

  if (rule) {
    try {
      await operations.deleteRule(rule.id);
    } catch (error) {
      throw new Error(
        `Case ${trackingId} was ${verb.past}, but its disabled OLHelper rule could not be removed.`,
        { cause: error },
      );
    }
  }

  return { folder, moved };
}

export async function archiveCase(
  trackingId: string,
  operations: CaseMailboxOperations = graphCaseMailboxOperations,
): Promise<CaseActionResult> {
  const { moved } = await retireActiveCase(
    trackingId,
    operations,
    (folder) => operations.moveFolder(folder.id, "archived"),
    { gerund: "Archiving", past: "archived" },
  );

  return {
    trackingId,
    location: "archived",
    folderId: moved.id,
  };
}

// Deletes an active case instead of archiving it: the case folder and every
// message in it go to Deleted Items (recoverable) and the managed rule is removed.
// Mail already in the Inbox, Sent Items, or any other folder is left alone.
export async function deleteCase(
  trackingId: string,
  operations: CaseMailboxOperations = graphCaseMailboxOperations,
): Promise<DeleteCaseResult> {
  const { folder } = await retireActiveCase(
    trackingId,
    operations,
    (activeFolder) => operations.deleteFolder(activeFolder),
    { gerund: "Deleting", past: "moved to Deleted Items" },
  );

  return {
    trackingId,
    folderName: folder.displayName,
  };
}

export async function reopenCase(
  trackingId: string,
  operations: CaseMailboxOperations = graphCaseMailboxOperations,
): Promise<CaseActionResult> {
  const state = await operations.getFolderState(trackingId);

  if (state.location !== "archived" || !state.folder) {
    throw new Error(`Case ${trackingId} is not archived.`);
  }

  const staleRule = await operations.findRule(trackingId);

  if (staleRule) {
    await operations.deleteRule(staleRule.id);
  }

  const folder = await operations.moveFolder(state.folder.id, "active");
  const rule = await operations.ensureRule(trackingId, folder.id);

  try {
    await operations.setRuleEnabled(rule.id, true);
  } catch (error) {
    throw new Error(
      `Case ${trackingId} was reopened, but routing could not be enabled. Use Repair routing.`,
      { cause: error },
    );
  }

  return {
    trackingId,
    location: "active",
    folderId: folder.id,
  };
}

export async function repairCaseRouting(
  trackingId: string,
  operations: CaseMailboxOperations = graphCaseMailboxOperations,
): Promise<CaseActionResult> {
  const state = await operations.getFolderState(trackingId);

  if (state.location !== "active" || !state.folder) {
    throw new Error(`Case ${trackingId} is not active.`);
  }

  const existingRule = await operations.findRule(trackingId);
  const rule = !existingRule
    ? await operations.ensureRule(trackingId, state.folder.id)
    : existingRule.actions?.moveToFolder !== state.folder.id
      ? await operations.updateRuleTarget(
          existingRule.id,
          trackingId,
          state.folder.id,
        )
      : existingRule;
  await operations.setRuleEnabled(rule.id, true);

  return {
    trackingId,
    location: "active",
    folderId: state.folder.id,
  };
}

export async function getCaseStatus(
  trackingId: string,
  operations: CaseMailboxOperations = graphCaseMailboxOperations,
): Promise<CaseStatus> {
  const state = await operations.getFolderState(trackingId);

  if (state.location !== "active") {
    return {
      location: state.location,
      routing: "not-applicable",
    };
  }

  const rule = await operations.findRule(trackingId);

  return {
    location: "active",
    routing: !rule
      ? "missing"
      : rule.actions?.moveToFolder !== state.folder?.id
        ? "mistargeted"
        : rule.isEnabled
          ? "enabled"
          : "disabled",
  };
}
