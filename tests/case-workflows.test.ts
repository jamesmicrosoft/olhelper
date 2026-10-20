import assert from "node:assert/strict";
import test from "node:test";
import {
  archiveCase,
  deleteCase,
  getCaseStatus,
  reopenCase,
  repairCaseRouting,
  trackCase,
  type CaseMailboxOperations,
} from "../src/cases/case-workflows";
import type {
  CaseFolderState,
  CaseLocation,
  MailFolder,
} from "../src/graph/folders";
import type { InboxScanResult } from "../src/graph/messages";
import type { MessageRule } from "../src/graph/rules";

function createOperations(
  state: CaseFolderState,
  existingRule: MessageRule | null = null,
  inboxScan: InboxScanResult | Error = {
    messageIds: [],
    status: "complete",
    scannedCount: 250,
  },
  unmovableMessageIds: string[] = [],
  deleteFolderError: Error | null = null,
): CaseMailboxOperations & { calls: string[] } {
  const calls: string[] = [];
  const folder = state.folder ?? {
    id: "active-folder",
    displayName: "CASE-1",
  };
  const rule = existingRule ?? {
    id: "rule-1",
    displayName: "OLHelper | TrackingID#CASE-1",
    sequence: 1,
    isEnabled: false,
    actions: { moveToFolder: folder.id },
  };

  return {
    calls,
    async ensureActiveFolder() {
      calls.push("ensure-folder");
      return folder;
    },
    async getFolderState() {
      calls.push("get-state");
      return state;
    },
    async moveFolder(
      _folderId: string,
      destination: Exclude<CaseLocation, "untracked">,
    ) {
      calls.push(`move-folder:${destination}`);
      return { ...folder, id: `${destination}-folder` };
    },
    async deleteFolder() {
      calls.push("delete-folder");

      if (deleteFolderError) {
        throw deleteFolderError;
      }
    },
    async moveMessage(messageId: string) {
      calls.push(`move-message:${messageId}`);

      if (unmovableMessageIds.includes(messageId)) {
        throw new Error("Message could not be moved.");
      }

      return { id: "moved-message" };
    },
    async findInboxMessages() {
      calls.push("find-inbox-messages");

      if (inboxScan instanceof Error) {
        throw inboxScan;
      }

      return inboxScan;
    },
    async ensureRule() {
      calls.push("ensure-rule");
      return rule;
    },
    async findRule() {
      calls.push("find-rule");
      return existingRule;
    },
    async setRuleEnabled(_ruleId: string, enabled: boolean) {
      calls.push(`set-rule:${enabled}`);
      return { ...rule, isEnabled: enabled };
    },
    async updateRuleTarget() {
      calls.push("update-rule-target");
      return {
        ...rule,
        isEnabled: false,
        actions: { moveToFolder: folder.id },
      };
    },
    async deleteRule() {
      calls.push("delete-rule");
    },
  };
}

test("tracks a case by moving the message before enabling routing", async () => {
  const operations = createOperations({
    location: "untracked",
    folder: null,
  });

  const result = await trackCase("CASE-1", "message-1", operations);

  assert.equal(result.movedMessageId, "moved-message");
  assert.equal(result.sweptMessageCount, 0);
  assert.equal(result.unsweptMessageCount, 0);
  assert.equal(result.sweepStatus, "complete");
  assert.deepEqual(operations.calls, [
    "get-state",
    "ensure-folder",
    "ensure-rule",
    "move-message:message-1",
    "set-rule:true",
    "find-inbox-messages",
  ]);
});

test("sweeps pre-existing Inbox mail for the case into the case folder", async () => {
  const operations = createOperations(
    { location: "untracked", folder: null },
    null,
    {
      messageIds: ["assignment-1", "assignment-2"],
      status: "complete",
      scannedCount: 250,
    },
  );

  const result = await trackCase("CASE-1", "message-1", operations);

  assert.equal(result.sweptMessageCount, 2);
  assert.equal(result.unsweptMessageCount, 0);
  assert.equal(result.sweepStatus, "complete");
  assert.deepEqual(operations.calls, [
    "get-state",
    "ensure-folder",
    "ensure-rule",
    "move-message:message-1",
    "set-rule:true",
    "find-inbox-messages",
    "move-message:assignment-1",
    "move-message:assignment-2",
  ]);
});

test("reports a failed sweep instead of silently reporting nothing", async () => {
  const operations = createOperations(
    { location: "untracked", folder: null },
    null,
    new Error("Microsoft Graph returned HTTP 429"),
  );

  const result = await trackCase("CASE-1", "message-1", operations);

  assert.equal(result.sweptMessageCount, 0);
  assert.equal(
    result.sweepStatus,
    "failed",
    "a failed sweep must be distinguishable from a clean sweep",
  );
});

test("keeps matches found before a truncated scan and flags it truncated", async () => {
  const operations = createOperations(
    { location: "untracked", folder: null },
    null,
    { messageIds: ["assignment-1"], status: "truncated", scannedCount: 250 },
  );

  const result = await trackCase("CASE-1", "message-1", operations);

  assert.equal(result.sweptMessageCount, 1);
  assert.equal(result.sweepStatus, "truncated");
});

test("counts matching messages that could not be moved", async () => {
  const operations = createOperations(
    { location: "untracked", folder: null },
    null,
    {
      messageIds: ["assignment-1", "assignment-2"],
      status: "complete",
      scannedCount: 250,
    },
    ["assignment-2"],
  );

  const result = await trackCase("CASE-1", "message-1", operations);

  assert.equal(result.sweptMessageCount, 1);
  assert.equal(result.unsweptMessageCount, 1);
  assert.equal(result.sweepStatus, "complete");
});

test("does not track a message into an archived case", async () => {
  const operations = createOperations({
    location: "archived",
    folder: { id: "archived-folder", displayName: "CASE-1" },
  });

  await assert.rejects(
    trackCase("CASE-1", "message-1", operations),
    /is archived/,
  );
  assert.deepEqual(operations.calls, ["get-state"]);
});

test("archives an active case after disabling its rule", async () => {
  const existingRule: MessageRule = {
    id: "rule-1",
    displayName: "OLHelper | TrackingID#CASE-1",
    sequence: 1,
    isEnabled: true,
    actions: { moveToFolder: "active-folder" },
  };
  const operations = createOperations(
    {
      location: "active",
      folder: { id: "active-folder", displayName: "CASE-1" },
    },
    existingRule,
  );

  const result = await archiveCase("CASE-1", operations);

  assert.equal(result.location, "archived");
  assert.deepEqual(operations.calls, [
    "get-state",
    "find-rule",
    "set-rule:false",
    "move-folder:archived",
    "delete-rule",
  ]);
});

test("reopens an archived case and recreates routing", async () => {
  const operations = createOperations({
    location: "archived",
    folder: { id: "archived-folder", displayName: "CASE-1" },
  });

  const result = await reopenCase("CASE-1", operations);

  assert.equal(result.location, "active");
  assert.deepEqual(operations.calls, [
    "get-state",
    "find-rule",
    "move-folder:active",
    "ensure-rule",
    "set-rule:true",
  ]);
});

test("repairs routing only for an active case", async () => {
  const operations = createOperations({
    location: "active",
    folder: { id: "active-folder", displayName: "CASE-1" },
  });

  await repairCaseRouting("CASE-1", operations);

  assert.deepEqual(operations.calls, [
    "get-state",
    "find-rule",
    "ensure-rule",
    "set-rule:true",
  ]);
});

test("reports an active case with disabled routing", async () => {
  const operations = createOperations(
    {
      location: "active",
      folder: { id: "active-folder", displayName: "CASE-1" },
    },
    {
      id: "rule-1",
      displayName: "OLHelper | TrackingID#CASE-1",
      sequence: 1,
      isEnabled: false,
      actions: { moveToFolder: "active-folder" },
    },
  );

  const status = await getCaseStatus("CASE-1", operations);

  assert.deepEqual(status, {
    location: "active",
    routing: "disabled",
  });
  assert.deepEqual(operations.calls, ["get-state", "find-rule"]);
});

test("reports and repairs a rule that targets the wrong folder", async () => {
  const operations = createOperations(
    {
      location: "active",
      folder: { id: "active-folder", displayName: "CASE-1" },
    },
    {
      id: "rule-1",
      displayName: "OLHelper | TrackingID#CASE-1",
      sequence: 1,
      isEnabled: true,
      actions: { moveToFolder: "old-folder" },
    },
  );

  assert.deepEqual(await getCaseStatus("CASE-1", operations), {
    location: "active",
    routing: "mistargeted",
  });

  operations.calls.length = 0;
  await repairCaseRouting("CASE-1", operations);

  assert.deepEqual(operations.calls, [
    "get-state",
    "find-rule",
    "update-rule-target",
    "set-rule:true",
  ]);
});

test("archiving an untracked case says it is not tracked", async () => {
  const operations = createOperations({ location: "untracked", folder: null });

  await assert.rejects(archiveCase("CASE-1", operations), /is not tracked/);
  assert.deepEqual(operations.calls, ["get-state"]);
});

test("deletes an active case folder after disabling its rule, then removes the rule", async () => {
  const existingRule: MessageRule = {
    id: "rule-1",
    displayName: "OLHelper | TrackingID#CASE-1",
    sequence: 1,
    isEnabled: true,
    actions: { moveToFolder: "active-folder" },
  };
  const operations = createOperations(
    {
      location: "active",
      folder: { id: "active-folder", displayName: "CASE-1 renamed" },
    },
    existingRule,
  );

  const result = await deleteCase("CASE-1", operations);

  assert.deepEqual(result, {
    trackingId: "CASE-1",
    folderName: "CASE-1 renamed",
  });
  assert.deepEqual(operations.calls, [
    "get-state",
    "find-rule",
    "set-rule:false",
    "delete-folder",
    "delete-rule",
  ]);
});

test("deleting a case without a rule still deletes the folder", async () => {
  const operations = createOperations({
    location: "active",
    folder: { id: "active-folder", displayName: "CASE-1" },
  });

  await deleteCase("CASE-1", operations);

  assert.deepEqual(operations.calls, [
    "get-state",
    "find-rule",
    "delete-folder",
  ]);
});

test("re-enables routing when deleting the case folder fails", async () => {
  const existingRule: MessageRule = {
    id: "rule-1",
    displayName: "OLHelper | TrackingID#CASE-1",
    sequence: 1,
    isEnabled: true,
    actions: { moveToFolder: "active-folder" },
  };
  const operations = createOperations(
    {
      location: "active",
      folder: { id: "active-folder", displayName: "CASE-1" },
    },
    existingRule,
    undefined,
    [],
    new Error("Folder could not be moved."),
  );

  await assert.rejects(
    deleteCase("CASE-1", operations),
    /Folder could not be moved/,
  );
  assert.deepEqual(operations.calls, [
    "get-state",
    "find-rule",
    "set-rule:false",
    "delete-folder",
    "set-rule:true",
  ]);
});

test("does not delete a case that is untracked or archived", async () => {
  const untracked = createOperations({ location: "untracked", folder: null });

  await assert.rejects(deleteCase("CASE-1", untracked), /is not tracked/);
  assert.deepEqual(untracked.calls, ["get-state"]);

  const archived = createOperations({
    location: "archived",
    folder: { id: "archived-folder", displayName: "CASE-1" },
  });

  await assert.rejects(deleteCase("CASE-1", archived), /is archived/);
  assert.deepEqual(archived.calls, ["get-state"]);
});