import "./taskpane.css";
import {
  archiveCase,
  deleteCase,
  getCaseStatus,
  reopenCase,
  repairCaseRouting,
  type CaseStatus,
} from "../cases/case-workflows";
import {
  describeCaseStatus,
  getAvailableCaseActions,
  type CaseAction,
} from "../cases/case-status";
import { trackSelectedCase } from "../cases/track-case";
import {
  caseFolderName,
  extractTrackingId,
} from "../cases/tracking-id";
import {
  assertSelectedMessage,
  getSelectedMessage,
  type SelectedMessage,
} from "../outlook/selected-message";
import { getSafeErrorMessage } from "../security/safe-error";

interface ActionDefinition {
  confirmation: string;
  details: string;
  progress: string;
  run: () => Promise<void>;
}

let confirmationTrigger: HTMLButtonElement | null = null;
let sessionEndTimer: number | undefined;
let secureOperationActive = false;
let mutationOperationActive = false;
let currentCaseStatus: CaseStatus | null = null;

Office.onReady(async (info) => {
  const endSessionButton = getButton("end-session");
  endSessionButton.disabled = false;
  endSessionButton.addEventListener("click", endSecureSession);

  if (info.host !== Office.HostType.Outlook) {
    setStatus("OLHelper can run only in Microsoft Outlook.");
    disableAllActions();
    return;
  }

  initializeOfficeTheme();

  if (
    !Office.context.requirements.isSetSupported(
      "NestedAppAuth",
      "1.1",
    )
  ) {
    setStatus(
      "This Outlook client does not support the authentication required by OLHelper.",
    );
    disableAllActions();
    return;
  }

  const item = Office.context.mailbox.item;

  if (!item) {
    setStatus("Open a received message to use OLHelper.");
    disableAllActions();
    return;
  }

  let selectedMessage: SelectedMessage;

  try {
    selectedMessage = getSelectedMessage();
    await registerItemChangedProtection();
  } catch (error) {
    setStatus(getSafeErrorMessage(error));
    disableAllActions();
    return;
  }

  const subject = selectedMessage.subject;
  const extractedTrackingId = extractTrackingId(subject);

  document.getElementById("subject")!.textContent = subject;
  document.getElementById("tracking-id")!.textContent =
    extractedTrackingId
      ? `Tracking ID: ${extractedTrackingId}`
      : "No Tracking ID detected";

  if (!extractedTrackingId) {
    disableAllActions();
    return;
  }

  const trackingId = caseFolderName(extractedTrackingId);
  const mailbox =
    Office.context.mailbox.userProfile.emailAddress ||
    "the signed-in mailbox";
  const checkStatusButton = getButton("check-status");
  const cancelButton = getButton("cancel-action");
  const confirmButton = getButton("confirm-action");
  let pendingAction: CaseAction | null = null;

  const definitions: Record<CaseAction, ActionDefinition> = {
    track: {
      confirmation: `Track ${trackingId} in ${mailbox}?`,
      details:
        "OLHelper will create or reuse the active case folder, move the selected message, scan up to the 250 most recent Inbox messages for the complete TrackingID token, move matching messages, and enable persistent Inbox routing.",
      progress: "Creating case routing and moving the message...",
      run: async () => {
        const result = await trackSelectedCase(selectedMessage);
        const { sweep } = result;
        let sweptNote = "";

        if (sweep.discoveryFailed) {
          sweptNote =
            " The case is active, but prior Inbox messages could not be checked.";
        } else {
          sweptNote =
            ` Checked ${sweep.scannedMessageCount} Inbox message(s) and moved ` +
            `${sweep.movedMessageCount} matching message(s).`;

          if (sweep.failedMessageCount > 0) {
            sweptNote +=
              ` ${sweep.failedMessageCount} matching message(s) could not be moved.`;
          }

          if (!sweep.scanComplete) {
            sweptNote +=
              " The scan limit was reached, so older Inbox messages were not checked.";
          }
        }

        setStatus(
          `Case ${result.trackingId} is active and the selected message was moved.${sweptNote}`,
        );
        showActionsForState({ location: "active", routing: "enabled" });
      },
    },
    archive: {
      confirmation: `Archive ${trackingId} in ${mailbox}?`,
      details:
        "OLHelper will stop persistent routing, move the entire case folder from Active to Archived, and remove its managed Inbox rule.",
      progress: "Stopping routing and archiving the case...",
      run: async () => {
        await archiveCase(trackingId);
        setStatus(
          `Case ${trackingId} is archived. New matching messages will remain in the Inbox.`,
        );
        showActionsForState({
          location: "archived",
          routing: "not-applicable",
        });
      },
    },
    delete: {
      confirmation: `Delete ${trackingId} and all of its emails in ${mailbox}?`,
      details:
        "OLHelper will stop persistent routing, remove its managed Inbox rule, and move the case folder and every email in it to Deleted Items. The case will not be archived. You can restore the folder from Deleted Items until it is emptied.",
      progress: "Stopping routing and deleting the case folder...",
      run: async () => {
        await deleteCase(trackingId);
        setStatus(
          `Case ${trackingId} was deleted. Its folder and emails are in Deleted Items, and routing was removed.`,
        );
        showActionsForState({
          location: "untracked",
          routing: "not-applicable",
        });
      },
    },
    reopen: {
      confirmation: `Reopen ${trackingId} in ${mailbox}?`,
      details:
        "OLHelper will move the case folder back to Active and restore persistent Inbox routing.",
      progress: "Reopening the case and restoring routing...",
      run: async () => {
        await reopenCase(trackingId);
        setStatus(`Case ${trackingId} is active and routing is enabled.`);
        showActionsForState({ location: "active", routing: "enabled" });
      },
    },
    repair: {
      confirmation: `Repair routing for ${trackingId}?`,
      details:
        "OLHelper will verify that the active case folder has one managed Inbox rule targeting the correct folder and enable it.",
      progress: "Checking and repairing case routing...",
      run: async () => {
        await repairCaseRouting(trackingId);
        setStatus(`Routing for case ${trackingId} is enabled.`);
        showActionsForState({ location: "active", routing: "enabled" });
      },
    },
  };

  for (const action of Object.keys(definitions) as CaseAction[]) {
    const actionButton = getButton(`${action}-case`);

    actionButton.addEventListener("click", () => {
      if (!getAvailableCaseActions(currentCaseStatus).includes(action)) {
        setStatus(
          "Check case status before selecting a mailbox-changing action.",
        );
        checkStatusButton.focus();
        return;
      }

      pendingAction = action;
      confirmationTrigger = actionButton;
      const definition = definitions[action];

      document.getElementById("confirmation-message")!.textContent =
        definition.confirmation;
      document.getElementById("confirmation-details")!.textContent =
        definition.details;
      document.getElementById("confirmation")!.hidden = false;
      setActionButtonsDisabled(true);
      confirmButton.focus();
      setStatus("Review and confirm the mailbox changes.");
    });
  }

  cancelButton.addEventListener("click", () => {
    pendingAction = null;
    closeConfirmation();
    setStatus("No mailbox changes were made.");
  });

  confirmButton.addEventListener("click", async () => {
    if (!pendingAction) {
      return;
    }

    const action = pendingAction;
    const definition = definitions[action];

    if (!getAvailableCaseActions(currentCaseStatus).includes(action)) {
      pendingAction = null;
      closeConfirmation();
      setStatus(
        "The action is not valid for the last checked case status. Check status again.",
      );
      return;
    }

    try {
      assertSelectedMessage(selectedMessage);
      beginSecureOperation(true);
      confirmButton.disabled = true;
      cancelButton.disabled = true;
      setStatus(definition.progress);
      await definition.run();
      pendingAction = null;
      closeConfirmation();
      scheduleSecureSessionEnd();
    } catch (error) {
      const errorMessage =
        `Unable to ${action} case: ${getSafeErrorMessage(error)}`;
      pendingAction = null;
      closeConfirmation();

      try {
        const state = await getCaseStatus(trackingId);
        showActionsForState(state);
      } catch {
        clearCaseStatus("Case status could not be refreshed.");
      }

      setStatus(errorMessage);
      scheduleSecureSessionEnd();
    }
  });

  const refreshCaseStatus = async (): Promise<CaseStatus | null> => {
    beginSecureOperation(false);
    clearCaseStatus("Checking case status...");
    checkStatusButton.disabled = true;
    setStatus("Checking case status...");

    try {
      const state = await getCaseStatus(trackingId);
      showActionsForState(state);
      setStatus(describeCaseStatus(trackingId, state));
      return state;
    } catch (error) {
      clearCaseStatus("Case status could not be determined.");
      setStatus(`Unable to check case status: ${getSafeErrorMessage(error)}`);
      return null;
    } finally {
      finishReadOnlyOperation();
    }
  };

  checkStatusButton.addEventListener("click", () => {
    void refreshCaseStatus();
  });

  const requestedAction = getRequestedAction();

  restoreActionAvailability();

  if (requestedAction) {
    const state = await refreshCaseStatus();

    if (state && getAvailableCaseActions(state).includes(requestedAction)) {
      getButton(`${requestedAction}-case`).click();
    }
  }
});

function registerItemChangedProtection(): Promise<void> {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.addHandlerAsync(
      Office.EventType.ItemChanged,
      () => {
        if (mutationOperationActive) {
          disableAllActions();
          return;
        }

        document.getElementById("confirmation")!.hidden = true;
        confirmationTrigger = null;
        disableAllActions();
        setStatus("The selected message changed. Reloading OLHelper...");
        window.location.reload();
      },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
          return;
        }

        reject(
          new Error(
            "OLHelper could not monitor message selection changes, so mailbox actions were disabled.",
          ),
        );
      },
    );
  });
}

function showActionsForState(
  state: CaseStatus,
): void {
  currentCaseStatus = state;
  const { location, routing } = state;

  document.getElementById("case-state")!.textContent =
    location === "untracked"
      ? "Status: Not tracked"
      : location === "archived"
        ? "Status: Archived"
        : `Status: Active | Routing: ${
            routing === "enabled"
              ? "Enabled"
              : routing === "mistargeted"
                ? "Wrong destination"
                : "Needs repair"
          }`;

  getButton("track-case").hidden = location !== "untracked";
  getButton("archive-case").hidden = location !== "active";
  getButton("delete-case").hidden = location !== "active";
  getButton("repair-case").hidden =
    location !== "active" || routing === "enabled";
  getButton("reopen-case").hidden = location !== "archived";
}

function clearCaseStatus(message = "Case status has not been checked."): void {
  currentCaseStatus = null;
  document.getElementById("case-state")!.textContent = message;
  getButton("track-case").hidden = false;
  getButton("archive-case").hidden = true;
  getButton("delete-case").hidden = true;
  getButton("repair-case").hidden = true;
  getButton("reopen-case").hidden = true;
}

function closeConfirmation(): void {
  document.getElementById("confirmation")!.hidden = true;
  restoreActionAvailability();
  getButton("confirm-action").disabled = false;
  getButton("cancel-action").disabled = false;

  const focusTarget =
    confirmationTrigger && !confirmationTrigger.hidden
      ? confirmationTrigger
      : getButton("check-status");
  confirmationTrigger = null;
  focusTarget.focus();
}

function setActionButtonsDisabled(disabled: boolean): void {
  for (const id of [
    "check-status",
    "track-case",
    "archive-case",
    "delete-case",
    "reopen-case",
    "repair-case",
  ]) {
    getButton(id).disabled = disabled;
  }
}

function disableAllActions(): void {
  setActionButtonsDisabled(true);
}

function restoreActionAvailability(): void {
  if (secureOperationActive) {
    disableAllActions();
    return;
  }

  const availableActions = new Set(
    getAvailableCaseActions(currentCaseStatus),
  );
  getButton("check-status").disabled = false;

  for (const action of [
    "track",
    "archive",
    "delete",
    "reopen",
    "repair",
  ] as const) {
    getButton(`${action}-case`).disabled = !availableActions.has(action);
  }
}

function scheduleSecureSessionEnd(): void {
  secureOperationActive = false;
  mutationOperationActive = false;
  disableAllActions();
  getButton("end-session").disabled = false;
  document.getElementById("confirmation")!.hidden = true;

  const status = document.getElementById("status")!;
  const currentMessage = status.textContent?.trim();
  status.textContent =
    `${currentMessage ? `${currentMessage} ` : ""}` +
    "The secure session will end automatically.";

  clearSessionEndTimer();

  const countdown = document.getElementById("session-countdown")!;
  let secondsRemaining = 10;
  countdown.hidden = false;
  countdown.textContent = `Session ending in ${secondsRemaining} seconds.`;

  sessionEndTimer = window.setInterval(() => {
    secondsRemaining -= 1;

    if (secondsRemaining <= 0) {
      endSecureSession();
      return;
    }

    countdown.textContent = `Session ending in ${secondsRemaining} seconds.`;
  }, 1_000);
}

function endSecureSession(): void {
  if (secureOperationActive) {
    return;
  }

  clearSessionEndTimer();

  disableAllActions();
  getButton("end-session").disabled = true;
  window.location.replace(
    new URL("session-ended.html", window.location.href).href,
  );
}

function beginSecureOperation(isMutation: boolean): void {
  secureOperationActive = true;
  mutationOperationActive = isMutation;
  setActionButtonsDisabled(true);
  getButton("end-session").disabled = true;
  document.getElementById("session-countdown")!.hidden = true;
  clearSessionEndTimer();
}

function finishReadOnlyOperation(): void {
  secureOperationActive = false;
  mutationOperationActive = false;
  restoreActionAvailability();
  getButton("end-session").disabled = false;
}

function clearSessionEndTimer(): void {
  if (sessionEndTimer !== undefined) {
    window.clearInterval(sessionEndTimer);
    sessionEndTimer = undefined;
  }
}

function getButton(id: string): HTMLButtonElement {
  return document.getElementById(id) as HTMLButtonElement;
}

function getRequestedAction(): CaseAction | null {
  const action = new URLSearchParams(window.location.search).get("action");

  return action === "track" ||
    action === "archive" ||
    action === "delete" ||
    action === "reopen" ||
    action === "repair"
    ? action
    : null;
}

function setStatus(message: string): void {
  document.getElementById("status")!.textContent = message;
}

function initializeOfficeTheme(): void {
  if (
    !Office.context.requirements.isSetSupported("Mailbox", "1.14")
  ) {
    return;
  }

  applyOfficeTheme(Office.context.officeTheme);
  Office.context.mailbox.addHandlerAsync(
    Office.EventType.OfficeThemeChanged,
    (event: Office.OfficeThemeChangedEventArgs) => {
      applyOfficeTheme(event.officeTheme);
    },
    (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.warn("OLHelper could not monitor Office theme changes.");
      }
    },
  );
}

function applyOfficeTheme(theme: Office.OfficeTheme): void {
  const root = document.documentElement;

  setThemeColor("--body-background", theme.bodyBackgroundColor);
  setThemeColor("--body-foreground", theme.bodyForegroundColor);
  setThemeColor("--surface-background", theme.controlBackgroundColor);
  setThemeColor("--border-color", theme.controlForegroundColor);

  root.dataset.officeTheme = isDarkColor(theme.bodyBackgroundColor)
    ? "dark"
    : "light";
}

function setThemeColor(property: string, value: string): void {
  if (/^#[0-9a-f]{6}$/i.test(value)) {
    document.documentElement.style.setProperty(property, value);
  }
}

function isDarkColor(value: string): boolean {
  const match = /^#([0-9a-f]{6})$/i.exec(value);

  if (!match) {
    return false;
  }

  const color = Number.parseInt(match[1], 16);
  const red = (color >> 16) & 0xff;
  const green = (color >> 8) & 0xff;
  const blue = color & 0xff;

  return red * 0.299 + green * 0.587 + blue * 0.114 < 128;
}
