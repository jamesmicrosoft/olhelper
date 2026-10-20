import assert from "node:assert/strict";
import test from "node:test";
import { buildGraphUrl, GraphError } from "../src/graph/client";

test("builds a Microsoft Graph v1 URL from an API path", () => {
  assert.equal(
    buildGraphUrl(
      "/me/mailFolders" +
        "?$select=id,displayName&$filter=displayName%20eq%20'Support%20Cases'",
    ),
    "https://graph.microsoft.com/v1.0/me/mailFolders" +
      "?$select=id,displayName&$filter=displayName%20eq%20%27Support%20Cases%27",
  );
});

test("allows Graph pagination URLs", () => {
  const slashUrl =
    "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messageRules" +
    "?$select=id,displayName,sequence,isEnabled,isReadOnly,hasError&$skip=10";
  const canonicalUrl =
    "https://graph.microsoft.com/v1.0/me/mailFolders('Inbox')/messages" +
    "?$select=id,subject&$top=50&$orderby=receivedDateTime%20desc&$skiptoken=opaque";
  const childFolderUrl =
    "https://graph.microsoft.com/v1.0/me/mailFolders('folder-1')/childFolders" +
    "?$select=id,displayName&$filter=startswith(displayName%2C%20'1234567890123456')" +
    "&$skiptoken=opaque";

  assert.equal(buildGraphUrl(slashUrl), slashUrl);
  assert.equal(buildGraphUrl(canonicalUrl), canonicalUrl);
  assert.equal(
    buildGraphUrl(childFolderUrl),
    childFolderUrl.replace(
      "'1234567890123456'",
      "%271234567890123456%27",
    ),
  );
});

test("allows moving a case folder to the well-known Deleted Items folder", () => {
  assert.doesNotThrow(() =>
    buildGraphUrl("/me/mailFolders/folder-1/move", {
      method: "POST",
      body: JSON.stringify({ destinationId: "deleteditems" }),
    }),
  );
  assert.throws(() =>
    buildGraphUrl("/me/mailFolders/folder-1/permanentDelete", {
      method: "POST",
    }),
  );
  assert.throws(() =>
    buildGraphUrl("/me/mailFolders/folder-1", { method: "DELETE" }),
  );
});

test("rejects URLs outside the Graph v1 boundary", () => {
  assert.throws(() => buildGraphUrl("https://example.com/steal"));
  assert.throws(() => buildGraphUrl("https://graph.microsoft.com/beta/me"));
  assert.throws(() =>
    buildGraphUrl("https://graph.microsoft.com.evil.example/v1.0/me"),
  );
});

test("rejects Graph APIs outside OLHelper's operation allowlist", () => {
  assert.throws(() => buildGraphUrl("/me/drive/root"));
  assert.throws(() => buildGraphUrl("/users/another-user/messages"));
  assert.throws(() =>
    buildGraphUrl("https://graph.microsoft.com/v1.0/me/messages"),
  );
});

test("allows only the folder, message move, and rule endpoint shapes", () => {
  assert.doesNotThrow(() =>
    buildGraphUrl(
      "/me/mailFolders" +
        "?$select=id,displayName&$filter=displayName%20eq%20'Support%20Cases'",
    ),
  );
  assert.doesNotThrow(() =>
    buildGraphUrl(
      "/me/mailFolders/folder-1/childFolders" +
        "?$select=id,displayName&$filter=startswith(displayName%2C%20'1234567890123456')",
    ),
  );
  assert.doesNotThrow(() =>
    buildGraphUrl("/me/mailFolders/folder-1/move", {
      method: "POST",
      body: JSON.stringify({ destinationId: "folder-2" }),
    }),
  );
  assert.doesNotThrow(() =>
    buildGraphUrl("/me/messages/message-1/move", {
      method: "POST",
      body: JSON.stringify({ destinationId: "folder-2" }),
    }),
  );
  assert.doesNotThrow(() =>
    buildGraphUrl(
      "/me/mailFolders/inbox/messages" +
        "?$select=id,subject&$top=50&$orderby=receivedDateTime%20desc",
    ),
  );
  assert.doesNotThrow(() =>
    buildGraphUrl(
      "/me/mailFolders/inbox/messageRules" +
        "?$select=id,displayName,sequence,isEnabled,isReadOnly,hasError",
    ),
  );
  assert.doesNotThrow(() =>
    buildGraphUrl(
      "/me/mailFolders/inbox/messageRules/rule-1" +
        "?$select=id,displayName,sequence,isEnabled,isReadOnly,hasError,conditions,exceptions,actions",
    ),
  );
});

test("rejects unsupported methods, queries, headers, and request options", () => {
  assert.throws(() =>
    buildGraphUrl(
      "/me/mailFolders" +
        "?$select=id,displayName&$filter=startswith(displayName%2C%20'1234567890123456')",
    ),
  );
  assert.throws(() =>
    buildGraphUrl(
      "/me/mailFolders/inbox/messages?$select=id,subject,body&$top=50&$orderby=receivedDateTime%20desc",
    ),
  );
  assert.throws(() =>
    buildGraphUrl(
      "/me/mailFolders/inbox/messages?$select=id,subject&$top=500&$orderby=receivedDateTime%20desc",
    ),
  );
  assert.throws(() =>
    buildGraphUrl("/me/mailFolders/inbox/messages", { method: "POST" }),
  );
  assert.throws(() =>
    buildGraphUrl("/me/mailFolders/inbox/messageRules/rule-1", {
      method: "DELETE",
      body: "{}",
    }),
  );
  assert.throws(() =>
    buildGraphUrl(
      "/me/mailFolders/inbox/messageRules" +
        "?$select=id,displayName,sequence,isEnabled,isReadOnly,hasError",
      { headers: { "X-HTTP-Method-Override": "DELETE" } },
    ),
  );
  assert.throws(() =>
    buildGraphUrl(
      "/me/mailFolders/inbox/messageRules" +
        "?$select=id,displayName,sequence,isEnabled,isReadOnly,hasError",
      { credentials: "include" },
    ),
  );
});

test("validates exact folder and rule mutation bodies", () => {
  assert.doesNotThrow(() =>
    buildGraphUrl("/me/mailFolders", {
      method: "POST",
      body: JSON.stringify({
        displayName: "Support Cases",
        isHidden: false,
      }),
    }),
  );
  assert.doesNotThrow(() =>
    buildGraphUrl("/me/mailFolders/folder-1/childFolders", {
      method: "POST",
      body: JSON.stringify({
        displayName: "1234567890123456",
        isHidden: false,
      }),
    }),
  );
  assert.doesNotThrow(() =>
    buildGraphUrl("/me/mailFolders/inbox/messageRules", {
      method: "POST",
      body: JSON.stringify({
        displayName: "OLHelper | TrackingID#1234567890123456",
        sequence: 1,
        isEnabled: false,
        conditions: {
          subjectContains: ["TrackingID#1234567890123456"],
        },
        actions: {
          moveToFolder: "folder-1",
          stopProcessingRules: false,
        },
      }),
    }),
  );

  assert.throws(() =>
    buildGraphUrl("/me/mailFolders/folder-1/childFolders", {
      method: "POST",
      body: JSON.stringify({
        displayName: "Unrelated mailbox folder",
        isHidden: false,
      }),
    }),
  );
  assert.throws(() =>
    buildGraphUrl("/me/mailFolders/inbox/messageRules", {
      method: "POST",
      body: JSON.stringify({
        displayName: "OLHelper | TrackingID#1234567890123456",
        sequence: 1,
        isEnabled: true,
        conditions: {
          subjectContains: ["Invoice"],
        },
        actions: {
          forwardTo: ["attacker@example.com"],
        },
      }),
    }),
  );
});

test("Graph errors retain only the HTTP status", () => {
  const error = new GraphError(403);

  assert.equal(error.message, "Microsoft Graph returned HTTP 403");
  assert.deepEqual(Object.keys(error), ["status"]);
});
