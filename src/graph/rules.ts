import { graphRequest } from "./client";

export interface MessageRule {
  id: string;
  displayName: string;
  sequence: number;
  isEnabled: boolean;
  actions?: {
    moveToFolder?: string;
  };
}

interface RuleCollection {
  value: MessageRule[];
  "@odata.nextLink"?: string;
}

export function managedRuleName(trackingId: string): string {
  return `OLHelper | TrackingID#${trackingId}`;
}

export function isManagedRuleForTrackingId(
  displayName: string,
  trackingId: string,
): boolean {
  return (
    displayName.localeCompare(
      managedRuleName(trackingId),
      undefined,
      { sensitivity: "accent" },
    ) === 0
  );
}

async function listAllRules(): Promise<MessageRule[]> {
  let url: string | undefined =
    "/me/mailFolders/inbox/messageRules";
  const rules: MessageRule[] = [];

  while (url) {
    const page: RuleCollection = await graphRequest<RuleCollection>(url);
    rules.push(...page.value);
    url = page["@odata.nextLink"];
  }

  return rules;
}

export async function findCaseRule(
  trackingId: string,
): Promise<MessageRule | null> {
  const displayName = managedRuleName(trackingId);
  const existing = (await listAllRules()).filter(
    (rule) =>
      isManagedRuleForTrackingId(rule.displayName, trackingId),
  );

  if (existing.length > 1) {
    throw new Error(`Multiple OLHelper rules exist for ${trackingId}.`);
  }

  return existing[0] ?? null;
}

export async function ensureCaseRule(
  trackingId: string,
  folderId: string,
): Promise<MessageRule> {
  const rules = await listAllRules();
  const displayName = managedRuleName(trackingId);
  const existing = rules.filter(
    (rule) =>
      isManagedRuleForTrackingId(rule.displayName, trackingId),
  );

  if (existing.length > 1) {
    throw new Error(`Multiple OLHelper rules exist for ${trackingId}.`);
  }

  if (existing.length === 1) {
    const rule = existing[0];

    if (rule.actions?.moveToFolder !== folderId) {
      throw new Error(
        `An OLHelper rule for ${trackingId} targets another folder.`,
      );
    }

    return rule;
  }

  const sequence =
    Math.max(0, ...rules.map((rule) => rule.sequence ?? 0)) + 1;

  return graphRequest<MessageRule>(
    "/me/mailFolders/inbox/messageRules",
    {
      method: "POST",
      body: JSON.stringify({
        displayName,
        sequence,
        isEnabled: false,
        conditions: {
          subjectContains: [`${trackingId}`],
        },
        actions: {
          moveToFolder: folderId,
          stopProcessingRules: false,
        },
      }),
    },
  );
}

export async function setCaseRuleEnabled(
  ruleId: string,
  isEnabled: boolean,
): Promise<MessageRule> {
  return graphRequest<MessageRule>(
    `/me/mailFolders/inbox/messageRules/${encodeURIComponent(ruleId)}`,
    {
      method: "PATCH",
      body: JSON.stringify({ isEnabled }),
    },
  );
}

export async function updateCaseRuleTarget(
  ruleId: string,
  trackingId: string,
  folderId: string,
): Promise<MessageRule> {
  return graphRequest<MessageRule>(
    `/me/mailFolders/inbox/messageRules/${encodeURIComponent(ruleId)}`,
    {
      method: "PATCH",
      body: JSON.stringify({
        isEnabled: false,
        conditions: {
          subjectContains: [`${trackingId}`],
        },
        actions: {
          moveToFolder: folderId,
          stopProcessingRules: false,
        },
      }),
    },
  );
}

export async function deleteCaseRule(ruleId: string): Promise<void> {
  await graphRequest<void>(
    `/me/mailFolders/inbox/messageRules/${encodeURIComponent(ruleId)}`,
    { method: "DELETE" },
  );
}