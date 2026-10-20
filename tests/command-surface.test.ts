import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import test from "node:test";

type ManifestAction = {
  id: string;
  type: string;
  pinnable?: boolean;
};

type ManifestRuntime = {
  code: { page: string };
  actions?: ManifestAction[];
};

type ManifestControl = {
  id: string;
  type: string;
  actionId?: string;
  items?: ManifestControl[];
};

const manifest = JSON.parse(
  readFileSync("appPackage/manifest.json", "utf8"),
) as {
  authorization: {
    permissions: {
      resourceSpecific: Array<{ name: string; type: string }>;
    };
  };
  extensions: Array<{
    requirements: {
      capabilities: Array<{ name: string; minVersion: string }>;
      scopes: string[];
    };
    runtimes: ManifestRuntime[];
    ribbons: Array<{
      contexts: string[];
      tabs: Array<{
        groups: Array<{ controls: ManifestControl[] }>;
      }>;
    }>;
  }>;
};
const commands = readFileSync("src/commands/commands.ts", "utf8");
const taskpane = readFileSync("src/taskpane/taskpane.ts", "utf8");
const extension = manifest.extensions[0];
const actions = extension.runtimes.flatMap((runtime) => runtime.actions ?? []);
const ribbons = extension.ribbons;
const controls = ribbons[0].tabs[0].groups[0].controls;

test("exposes a direct status command and task-pane case action menu", () => {
  assert.equal(ribbons.length, 1);
  assert.deepEqual(ribbons[0].contexts, ["mailRead"]);
  assert.equal(
    actions.filter(
      ({ id, type }) =>
        id === "checkCaseStatusCommand" && type === "executeFunction",
    ).length,
    1,
  );

  const menu = controls.find(({ id }) => id === "OlHelper.CaseActions");
  assert.equal(menu?.type, "menu");
  assert.deepEqual(
    menu?.items?.map(({ actionId }) => actionId),
    ["trackCase", "archiveCase", "deleteCase", "reopenCase", "repairRouting"],
  );

  assert.deepEqual(
    extension.runtimes
      .map(({ code }) => code.page)
      .filter((page) => page.includes("taskpane.html?action=")),
    [
      "https://localhost:3000/taskpane.html?action=track",
      "https://localhost:3000/taskpane.html?action=archive",
      "https://localhost:3000/taskpane.html?action=delete",
      "https://localhost:3000/taskpane.html?action=reopen",
      "https://localhost:3000/taskpane.html?action=repair",
    ],
  );
});

test("pins only the generic task pane entry point", () => {
  assert.deepEqual(
    actions.filter(({ pinnable }) => pinnable).map(({ id }) => id),
    ["openOlHelper"],
  );
  assert.deepEqual(
    actions
      .filter(({ type }) => type === "openPage")
      .map(({ id, pinnable }) => ({ id, pinnable })),
    [
      { id: "openOlHelper", pinnable: true },
      { id: "trackCase", pinnable: false },
      { id: "archiveCase", pinnable: false },
      { id: "deleteCase", pinnable: false },
      { id: "reopenCase", pinnable: false },
      { id: "repairRouting", pinnable: false },
    ],
  );
});

test("declares the Outlook and Office.js permissions used by the command surface", () => {
  assert.deepEqual(extension.requirements.scopes, ["mail"]);
  assert.deepEqual(extension.requirements.capabilities, [
    { name: "Mailbox", minVersion: "1.5" },
  ]);
  assert.deepEqual(manifest.authorization.permissions.resourceSpecific, [
    { name: "MailboxItem.Read.User", type: "Delegated" },
  ]);
});

test("completes the function command and reports status through Outlook", () => {
  assert.match(
    commands,
    /Office\.actions\.associate\(\s*"checkCaseStatusCommand"/,
  );
  assert.match(commands, /ProgressIndicator/);
  assert.match(commands, /InformationalMessage/);
  assert.match(commands, /ErrorMessage/);
  assert.match(
    commands,
    /finally\s*\{\s*releaseGraphAuthenticationContext\(\);\s*event\.completed\(\)/,
  );
});

test("task-pane action links only accept known lifecycle actions", () => {
  assert.match(taskpane, /new URLSearchParams\(window\.location\.search\)/);

  for (const action of ["track", "archive", "delete", "reopen", "repair"]) {
    assert.match(taskpane, new RegExp(`action === "${action}"`));
  }
});
