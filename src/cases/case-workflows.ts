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
  findInboxMessagesByTrackingId,
  moveMessage,
  type InboxMessageSearchResult,
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
  ): Promise<InboxMessageSearchResult>;
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
  sweep: {
    scannedMessageCount: number;
    matchedMessageCount: number;
    movedMessageCount: number;
    failedMessageCount: number;
    scanComplete: boolean;
    discoveryFailed: boolean;
  };
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
    findInboxMessagesByTrackingId(trackingId, excludeId),
};

export async function trackCase(
  trackingId: string,
  messageId: string,
  operations: CaseMailboxOperations = graphCaseMailboxOperations,
): Promise<TrackCaseResult> {
  const state = await operations.getFolderState(trackingId);

  if (state.location !== "untracked") {
    throw new Error(
      state.location === "archived"
        ? `Case ${trackingId} is archived. Reopen it before tracking new messages.`
        : `Case ${trackingId} is already active. Check its current routing status instead of tracking it again.`,
    );
  }

  const folder = await operations.ensureActiveFolder(trackingId);
  const rule = await operations.ensureRule(trackingId, folder.id);
  const sweep: TrackCaseResult["sweep"] = {
    scannedMessageCount: 0,
    matchedMessageCount: 0,
    movedMessageCount: 0,
    failedMessageCount: 0,
    scanComplete: false,
    discoveryFailed: false,
  };
  let pendingMessageIds: string[] = [];

  try {
    const search = await operations.findInboxMessages(
      trackingId,
      messageId,
    );
    pendingMessageIds = search.messageIds;
    sweep.scannedMessageCount = search.scannedMessageCount;
    sweep.matchedMessageCount = pendingMessageIds.length;
    sweep.scanComplete = search.scanComplete;
  } catch {
    sweep.discoveryFailed = true;
  }

  try {
    await operations.setRuleEnabled(rule.id, true);
  } catch (error) {
    throw new Error(
      `Routing could not be enabled for ${trackingId}. No messages were moved.`,
      { cause: error },
    );
  }

  for (const pendingId of pendingMessageIds) {
    try {
      await operations.moveMessage(pendingId, folder.id);
      sweep.movedMessageCount += 1;
    } catch {
      sweep.failedMessageCount += 1;
    }
  }

  let movedMessage: { id: string };

  try {
    movedMessage = await operations.moveMessage(messageId, folder.id);
  } catch (error) {
    if (!rule.isEnabled) {
      try {
        await operations.setRuleEnabled(rule.id, false);
      } catch (rollbackError) {
        throw new AggregateError(
          [error, rollbackError],
          `The selected message could not be moved and routing for ${trackingId} could not be disabled. Use Repair routing.`,
        );
      }
    }

    throw new Error(
      `Routing was prepared, but the selected message could not be moved for ${trackingId}.`,
      { cause: error },
    );
  }

  return {
    trackingId,
    folderId: folder.id,
    ruleId: rule.id,
    movedMessageId: movedMessage.id,
    sweep,
  };
}

// Archive and delete both retire an active case. Routing is stopped before the
// folder moves so no new mail lands in it mid-move, and the managed rule is
// removed afterwards so it can never route mail into the folder's new location.
async function retireActiveCase<T>(
  trackingId: string,
  operations: CaseMailboxOperations,
  moveFolder: (folder: MailFolder) => Promise<T>,
  verb: { gerund: string; past: string },
): Promise<{ folder: MailFolder; moved: T }> {
  const state = await operations.getFolderState(trackingId);

  if (state.location !== "active" || !state.folder) {
    throw new Error(`Case ${trackingId} is not active.`);
  }

  const folder = state.folder;
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
// message in it move to Deleted Items (recoverable) and the managed rule is
// removed. Mail in the Inbox, Sent Items, or any other folder is left alone.
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
