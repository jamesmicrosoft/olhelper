import assert from "node:assert/strict";
import test from "node:test";
import { GraphError } from "../src/graph/client";
import { isFolderNameConflict } from "../src/graph/folders";

test("recognises a Graph folder-name conflict", () => {
  assert.equal(isFolderNameConflict(new GraphError(409, "{}")), true);
  assert.equal(
    isFolderNameConflict(
      new GraphError(400, '{"error":{"code":"ErrorFolderExists"}}'),
    ),
    true,
  );
});

test("does not treat other failures as a folder-name conflict", () => {
  assert.equal(isFolderNameConflict(new GraphError(429, "{}")), false);
  assert.equal(isFolderNameConflict(new Error("ErrorFolderExists")), false);
});