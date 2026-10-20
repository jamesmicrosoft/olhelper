import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import test from "node:test";

const css = readFileSync("src/taskpane/taskpane.css", "utf8");
const html = readFileSync("src/taskpane/taskpane.html", "utf8");
const script = readFileSync("src/taskpane/taskpane.ts", "utf8");
const sessionEndedCss = readFileSync(
  "src/session-ended/session-ended.css",
  "utf8",
);
const sessionEndedHtml = readFileSync(
  "src/session-ended/session-ended.html",
  "utf8",
);
const sessionEndedScript = readFileSync(
  "src/session-ended/session-ended.ts",
  "utf8",
);

function getColor(block: string, property: string): string {
  const match = new RegExp(`${property}:\\s*(#[0-9a-f]{6})`, "i").exec(block);
  assert.ok(match, `${property} must define a six-digit hex color`);
  return match[1];
}

function contrastRatio(first: string, second: string): number {
  const luminance = (hex: string): number => {
    const channels = [1, 3, 5].map((start) => {
      const channel = Number.parseInt(hex.slice(start, start + 2), 16) / 255;
      return channel <= 0.04045
        ? channel / 12.92
        : ((channel + 0.055) / 1.055) ** 2.4;
    });

    return (
      channels[0] * 0.2126 + channels[1] * 0.7152 + channels[2] * 0.0722
    );
  };

  const lighter = Math.max(luminance(first), luminance(second));
  const darker = Math.min(luminance(first), luminance(second));
  return (lighter + 0.05) / (darker + 0.05);
}

test("provides light, dark, forced-color, and reduced-motion styles", () => {
  assert.match(css, /color-scheme:\s*light dark/);
  assert.match(css, /@media \(prefers-color-scheme:\s*dark\)/);
  assert.match(css, /@media \(forced-colors:\s*active\)/);
  assert.match(css, /@media \(prefers-reduced-motion:\s*reduce\)/);
  assert.match(css, /button:focus-visible/);
});

test("fallback foreground and control colors meet WCAG AA text contrast", () => {
  const light = /:root\s*\{([\s\S]*?)\}/.exec(css)?.[1];
  const dark =
    /:root:not\(\[data-office-theme\]\)\s*\{([\s\S]*?)\}/.exec(css)?.[1];
  assert.ok(light);
  assert.ok(dark);

  const themes = [light, dark];
  const pairs = [
    ["--body-foreground", "--body-background"],
    ["--muted-foreground", "--body-background"],
    ["--primary-foreground", "--primary-background"],
    ["--warning-foreground", "--warning-background"],
    ["--disabled-foreground", "--disabled-background"],
  ];

  for (const theme of themes) {
    for (const [foreground, background] of pairs) {
      assert.ok(
        contrastRatio(
          getColor(theme, foreground),
          getColor(theme, background),
        ) >= 4.5,
        `${foreground} must have at least 4.5:1 contrast on ${background}`,
      );
    }
  }
});

test("exposes status and confirmation changes to assistive technology", () => {
  assert.match(html, /role="status"/);
  assert.match(html, /aria-live="polite"/);
  assert.match(html, /aria-atomic="true"/);
  assert.match(html, /role="alertdialog"/);
  assert.match(html, /aria-labelledby="confirmation-message"/);
  assert.match(html, /aria-describedby="confirmation-details"/);
  assert.match(html, /id="session-countdown" aria-live="off" hidden/);
});

test("uses and monitors the Outlook Office theme when supported", () => {
  assert.match(script, /isSetSupported\("Mailbox", "1\.14"\)/);
  assert.match(script, /Office\.EventType\.OfficeThemeChanged/);
  assert.match(script, /applyOfficeTheme\(event\.officeTheme\)/);
});

test("fails closed when a pinned pane changes selected messages", () => {
  assert.match(script, /Office\.EventType\.ItemChanged/);
  assert.match(script, /assertSelectedMessage\(selectedMessage\)/);
  assert.match(script, /disableAllActions\(\)/);
  assert.match(script, /window\.location\.reload\(\)/);
});

test("fails closed outside the Outlook host", () => {
  assert.match(script, /info\.host !== Office\.HostType\.Outlook/);
  assert.match(script, /OLHelper can run only in Microsoft Outlook/);
});

test("keeps mailbox actions disabled until Office initialization completes", () => {
  for (const id of [
    "check-status",
    "track-case",
    "archive-case",
    "delete-case",
    "reopen-case",
    "repair-case",
    "end-session",
  ]) {
    assert.match(
      html,
      new RegExp(`<button[^>]*id="${id}"[^>]*disabled`),
    );
  }
});

test("requires checked case state before enabling Track", () => {
  assert.match(script, /let currentCaseStatus: CaseStatus \| null = null/);
  assert.match(
    script,
    /getAvailableCaseActions\(currentCaseStatus\)\.includes\(action\)/,
  );
  assert.match(script, /getButton\("track-case"\)\.hidden = location !== "untracked"/);
  assert.match(script, /restoreActionAvailability\(\)/);
  assert.match(script, /clearCaseStatus\("Case status could not be determined\."\)/);
});

test("requires a described confirmation before deleting a case", () => {
  assert.match(script, /delete: \{\s*confirmation:/);
  assert.match(script, /move the case folder and every email in it to Deleted Items/);
  assert.match(script, /await deleteCase\(trackingId\)/);
});

test("does not display raw token or broker telemetry errors", () => {
  assert.match(script, /getSafeErrorMessage/);
});

test("provides an explicit secure-session termination control", () => {
  assert.match(html, /id="end-session"/);
  assert.match(script, /addEventListener\("click", endSecureSession\)/);
  assert.match(script, /secureOperationActive = true/);
  assert.match(script, /mutationOperationActive = isMutation/);
  assert.match(
    script,
    /if \(secureOperationActive\)\s*\{\s*return;\s*\}/,
  );
  assert.match(script, /getButton\("end-session"\)\.disabled = true/);
  assert.match(script, /if \(mutationOperationActive\)/);
  assert.match(script, /Session ending in \$\{secondsRemaining\} seconds/);
});

test("provides an accessible, theme-aware session-ended page", () => {
  assert.match(sessionEndedHtml, /id="close-pane"/);
  assert.match(sessionEndedHtml, /role="status"/);
  assert.match(sessionEndedCss, /color-scheme:\s*light dark/);
  assert.match(sessionEndedCss, /@media \(prefers-color-scheme:\s*dark\)/);
  assert.match(sessionEndedCss, /@media \(forced-colors:\s*active\)/);
  assert.match(sessionEndedCss, /button:focus-visible/);
  assert.match(sessionEndedScript, /Office\.onReady/);
  assert.match(sessionEndedScript, /Office\.EventType\.OfficeThemeChanged/);
  assert.match(sessionEndedScript, /Office\.context\.ui\.closeContainer\(\)/);
});
