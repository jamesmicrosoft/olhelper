import assert from "node:assert/strict";
import test from "node:test";
import {
  findInboxMessagesBySubject,
  type MessageCollection,
} from "../src/graph/messages";

const caseNumber = "2600000000000001";

function page(
  messages: Array<{ id: string; subject?: string }>,
  nextLink?: string,
): MessageCollection {
  return nextLink
    ? { value: messages, "@odata.nextLink": nextLink }
    : { value: messages };
}

const nextPage =
  "https://graph.microsoft.com/v1.0/me/mailFolders('Inbox')/messages?$skip=50";

function isSearch(url: string): boolean {
  return url.includes("$search");
}

test("matches real assignment subjects that carry no TrackingID# token", async () => {
  const scan = await findInboxMessagesBySubject(
    caseNumber,
    undefined,
    250,
    async () =>
      page([
        {
          id: "vdm",
          subject: `VDM has assigned SR ${caseNumber} to jamesshin`,
        },
        {
          id: "assignment-change",
          subject: `Case assignment change: CrmGlobal-DFM-MSaaS assigned Severity C case ${caseNumber} to James Shin`,
        },
        { id: "unrelated", subject: "Stock Award Vest Confirmation" },
      ]),
  );

  assert.deepEqual(scan.messageIds, ["vdm", "assignment-change"]);
  assert.equal(scan.status, "complete");
});

// Regression: a 4,755-message Inbox always hits the recent-scan cap, which used
// to be reported as an incomplete sweep on EVERY Track. A warning that always
// fires gets ignored, which is just the old silent failure wearing a hat.
test("a capped recent scan is not reported as incomplete when search succeeds", async () => {
  const scan = await findInboxMessagesBySubject(
    caseNumber,
    undefined,
    100,
    async (url) =>
      isSearch(url)
        ? page([{ id: "deep", subject: `SR ${caseNumber} follow-up` }])
        : page(
            Array.from({ length: 50 }, (_, i) => ({
              id: `noise-${i}`,
              subject: "unrelated",
            })),
            nextPage,
          ),
  );

  assert.equal(
    scan.status,
    "complete",
    "search covered the whole Inbox, so the recent-scan cap is not a coverage gap",
  );
  assert.deepEqual(scan.messageIds, ["deep"]);
});

test("search finds case mail deeper than the recent-scan cap", async () => {
  const scan = await findInboxMessagesBySubject(
    caseNumber,
    undefined,
    50,
    async (url) =>
      isSearch(url)
        ? page([
            { id: "old-jit", subject: `DfM JIT approved (Case #${caseNumber})` },
          ])
        : page(
            [{ id: "recent", subject: `VDM has assigned SR ${caseNumber}` }],
            nextPage,
          ),
  );

  assert.deepEqual(
    [...scan.messageIds].sort(),
    ["old-jit", "recent"],
    "results from both passes must be merged",
  );
  assert.equal(scan.status, "complete");
});

test("a message found by both passes is only swept once", async () => {
  const scan = await findInboxMessagesBySubject(
    caseNumber,
    undefined,
    250,
    async () => page([{ id: "dupe", subject: `SR ${caseNumber}` }]),
  );

  assert.deepEqual(scan.messageIds, ["dupe"]);
});

test("search results are re-checked, so a body-only hit is never swept", async () => {
  const scan = await findInboxMessagesBySubject(
    caseNumber,
    undefined,
    250,
    async (url) =>
      isSearch(url)
        ? page([
            { id: "body-hit", subject: "Weekly digest with no case number" },
          ])
        : page([]),
  );

  assert.deepEqual(
    scan.messageIds,
    [],
    "search may only add recall, never loosen subject matching",
  );
});

test("reports truncated when search fails and the cap was hit", async () => {
  const scan = await findInboxMessagesBySubject(
    caseNumber,
    undefined,
    50,
    async (url) => {
      if (isSearch(url)) {
        throw new Error("Microsoft Graph returned HTTP 429");
      }

      return page([{ id: "recent", subject: `SR ${caseNumber}` }], nextPage);
    },
  );

  assert.deepEqual(scan.messageIds, ["recent"]);
  assert.equal(scan.status, "truncated");
});

test("reports failed when both passes fail", async () => {
  const scan = await findInboxMessagesBySubject(
    caseNumber,
    undefined,
    250,
    async () => {
      throw new Error("Microsoft Graph returned HTTP 429");
    },
  );

  assert.equal(scan.status, "failed");
});

test("keeps matches from earlier pages when a later page fails", async () => {
  let directCall = 0;

  const scan = await findInboxMessagesBySubject(
    caseNumber,
    undefined,
    250,
    async (url) => {
      if (isSearch(url)) {
        throw new Error("Microsoft Graph returned HTTP 429");
      }

      directCall += 1;

      if (directCall === 1) {
        return page(
          [{ id: "vdm", subject: `VDM has assigned SR ${caseNumber}` }],
          nextPage,
        );
      }

      throw new Error("Microsoft Graph returned HTTP 429");
    },
  );

  assert.deepEqual(
    scan.messageIds,
    ["vdm"],
    "a throttled later page must not discard matches already found",
  );
  assert.equal(scan.status, "failed");
});

test("excludes the already-moved selected message and untitled mail", async () => {
  const scan = await findInboxMessagesBySubject(
    caseNumber,
    "selected",
    250,
    async () =>
      page([
        {
          id: "selected",
          subject: `failure in vmss - TrackingID#${caseNumber}`,
        },
        { id: "no-subject" },
        { id: "other", subject: `SR ${caseNumber} handover` },
      ]),
  );

  assert.deepEqual(scan.messageIds, ["other"]);
});

test("matches a collaboration subject from the main case number", async () => {
  const scan = await findInboxMessagesBySubject(
    caseNumber,
    undefined,
    250,
    async (url) =>
      isSearch(url)
        ? page([])
        : page([
            { id: "collab", subject: `Collab case ${caseNumber}001 update` },
          ]),
  );

  assert.deepEqual(
    scan.messageIds,
    ["collab"],
    "the direct scan must keep substring matching so collab numbers still route",
  );
});
