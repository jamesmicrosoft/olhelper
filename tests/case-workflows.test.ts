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
import type { MessageRule } from "../src/graph/rules";

function createOperations(
  state: CaseFolderState,
  existingRule: MessageRule | null = null,
  failures: { moveFolder?: Error; deleteFolder?: Error } = {},
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

      if (failures.moveFolder) {
        throw failures.moveFolder;
      }

      return { ...folder, id: `${destination}-folder` };
    },
    async deleteFolder() {
      calls.push("delete-folder");

      if (failures.deleteFolder) {
        throw failures.deleteFolder;
      }
    },
    async moveMessage(messageId) {
      calls.push(`move-message:${messageId}`);
      return { id: "moved-message" };
    },
    async findInboxMessages() {
      calls.push("find-inbox");
      return {
        messageIds: [],
        scannedMessageCount: 0,
        scanComplete: true,
      };
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

test("finishes Inbox work before moving the selected message", async () => {
  const operations = createOperations({
    location: "untracked",
    folder: null,
  });

  const result = await trackCase("CASE-1", "message-1", operations);

  assert.equal(result.movedMessageId, "moved-message");
  assert.deepEqual(result.sweep, {
    scannedMessageCount: 0,
    matchedMessageCount: 0,
    movedMessageCount: 0,
    failedMessageCount: 0,
    scanComplete: true,
    discoveryFailed: false,
  });
  assert.deepEqual(operations.calls, [
    "get-state",
    "ensure-folder",
    "ensure-rule",
    "find-inbox",
    "set-rule:true",
    "move-message:message-1",
  ]);
});

test("reports a partial Inbox sweep without failing Track", async () => {
  const operations = createOperations({
    location: "untracked",
    folder: null,
  });
  operations.findInboxMessages = async () => {
    operations.calls.push("find-inbox");
    return {
      messageIds: ["pending-1", "pending-2"],
      scannedMessageCount: 250,
      scanComplete: false,
    };
  };
  operations.moveMessage = async (messageId) => {
    operations.calls.push(`move-message:${messageId}`);
    if (messageId === "pending-2") {
      throw new Error("move failed");
    }
    return { id: `moved-${messageId}` };
  };

  const result = await trackCase("CASE-1", "selected", operations);

  assert.deepEqual(result.sweep, {
    scannedMessageCount: 250,
    matchedMessageCount: 2,
    movedMessageCount: 1,
    failedMessageCount: 1,
    scanComplete: false,
    discoveryFailed: false,
  });
  assert.deepEqual(operations.calls, [
    "get-state",
    "ensure-folder",
    "ensure-rule",
    "find-inbox",
    "set-rule:true",
    "move-message:pending-1",
    "move-message:pending-2",
    "move-message:selected",
  ]);
});

test("reports Inbox discovery failure without failing Track", async () => {
  const operations = createOperations({
    location: "untracked",
    folder: null,
  });
  operations.findInboxMessages = async () => {
    operations.calls.push("find-inbox");
    throw new Error("discovery failed");
  };

  const result = await trackCase("CASE-1", "selected", operations);

  assert.deepEqual(result.sweep, {
    scannedMessageCount: 0,
    matchedMessageCount: 0,
    movedMessageCount: 0,
    failedMessageCount: 0,
    scanComplete: false,
    discoveryFailed: true,
  });
  assert.deepEqual(operations.calls, [
    "get-state",
    "ensure-folder",
    "ensure-rule",
    "find-inbox",
    "set-rule:true",
    "move-message:selected",
  ]);
});

test("disables a newly enabled rule when the selected move fails", async () => {
  const operations = createOperations({
    location: "untracked",
    folder: null,
  });
  operations.moveMessage = async (messageId) => {
    operations.calls.push(`move-message:${messageId}`);
    throw new Error("move failed");
  };

  await assert.rejects(
    trackCase("CASE-1", "selected", operations),
    /selected message could not be moved/,
  );
  assert.deepEqual(operations.calls, [
    "get-state",
    "ensure-folder",
    "ensure-rule",
    "find-inbox",
    "set-rule:true",
    "move-message:selected",
    "set-rule:false",
  ]);
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

test("does not recreate routing for an active case", async () => {
  const operations = createOperations({
    location: "active",
    folder: { id: "active-folder", displayName: "CASE-1" },
  });

  await assert.rejects(
    trackCase("CASE-1", "message-1", operations),
    /already active/,
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

test("re-enables routing when archiving the case folder fails", async () => {
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
    { moveFolder: new Error("Folder could not be moved.") },
  );

  await assert.rejects(
    archiveCase("CASE-1", operations),
    /Folder could not be moved/,
  );
  assert.deepEqual(operations.calls, [
    "get-state",
    "find-rule",
    "set-rule:false",
    "move-folder:archived",
    "set-rule:true",
  ]);
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

test("deletes the folder of an active case whose rule is already missing", async () => {
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
    { deleteFolder: new Error("Folder could not be moved.") },
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
  for (const state of [
    { location: "untracked", folder: null },
    {
      location: "archived",
      folder: { id: "archived-folder", displayName: "CASE-1" },
    },
  ] as const) {
    const operations = createOperations(state);

    await assert.rejects(
      deleteCase("CASE-1", operations),
      /is not active/,
    );
    assert.deepEqual(operations.calls, ["get-state"]);
  }
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
