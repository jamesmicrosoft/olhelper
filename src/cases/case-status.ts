import type { CaseStatus } from "./case-workflows";

export type CaseAction =
  | "track"
  | "archive"
  | "delete"
  | "reopen"
  | "repair";

export function describeCaseStatus(
  trackingId: string,
  state: CaseStatus,
): string {
  if (state.location === "untracked") {
    return `Case ${trackingId} is not currently tracked.`;
  }

  if (state.location === "archived") {
    return `Case ${trackingId} is archived and persistent routing is stopped.`;
  }

  return state.routing === "enabled"
    ? `Case ${trackingId} is active and routing is enabled.`
    : `Case ${trackingId} is active, but routing needs repair.`;
}

export function getAvailableCaseActions(
  state: CaseStatus | null,
): readonly CaseAction[] {
  if (!state) {
    return [];
  }

  if (state.location === "untracked") {
    return ["track"];
  }

  if (state.location === "archived") {
    return ["reopen"];
  }

  return state.routing === "enabled"
    ? ["archive", "delete"]
    : ["archive", "delete", "repair"];
}
