import { GraphError, graphRequest } from "./client";

export interface MailFolder {
  id: string;
  displayName: string;
}

export type CaseLocation = "active" | "archived" | "untracked";

export interface CaseFolderState {
  location: CaseLocation;
  folder: MailFolder | null;
}

interface FolderCollection {
  value: MailFolder[];
  "@odata.nextLink"?: string;
}

const rootFolderName = "Support Cases";
const activeFolderName = "Active";
const archivedFolderName = "Archived";

function escapeODataString(value: string): string {
  return value.replaceAll("'", "''");
}

async function findFolder(
  collectionPath: string,
  displayName: string,
): Promise<MailFolder | null> {
  const filter = encodeURIComponent(
    `displayName eq '${escapeODataString(displayName)}'`,
  );

  const result = await graphRequest<FolderCollection>(
    `${collectionPath}?$select=id,displayName&$filter=${filter}`,
  );

  if (result.value.length > 1) {
    throw new Error(`Multiple folders named "${displayName}" were found.`);
  }

  return result.value[0] ?? null;
}

async function createFolder(
  collectionPath: string,
  displayName: string,
): Promise<MailFolder> {
  return graphRequest<MailFolder>(collectionPath, {
    method: "POST",
    body: JSON.stringify({
      displayName,
      isHidden: false,
    }),
  });
}

async function ensureFolder(
  collectionPath: string,
  displayName: string,
): Promise<MailFolder> {
  return (
    (await findFolder(collectionPath, displayName)) ??
    (await createFolder(collectionPath, displayName))
  );
}

// Case folders may be renamed by the user to append a customer name or topic after
// the tracking ID (for example "1234567890123456 Tim - VM resize"). Match the
// tracking ID as a prefix so those renames do not break tracking. An exact match is
// preferred; otherwise the tracking ID must be followed by a space so a longer
// numeric ID is never matched by mistake.
async function findCaseFolder(
  collectionPath: string,
  trackingId: string,
): Promise<MailFolder | null> {
  const filter = encodeURIComponent(
    `startswith(displayName, '${escapeODataString(trackingId)}')`,
  );

  const result = await graphRequest<FolderCollection>(
    `${collectionPath}?$select=id,displayName&$filter=${filter}`,
  );

  const matches = result.value.filter(
    (f) =>
      f.displayName === trackingId ||
      f.displayName.startsWith(`${trackingId} `),
  );

  const exact = matches.find((f) => f.displayName === trackingId);
  if (exact) {
    return exact;
  }

  if (matches.length > 1) {
    throw new Error(
      `Multiple case folders start with "${trackingId}". Keep only one.`,
    );
  }

  return matches[0] ?? null;
}

async function ensureCaseFolder(
  collectionPath: string,
  trackingId: string,
): Promise<MailFolder> {
  return (
    (await findCaseFolder(collectionPath, trackingId)) ??
    (await createFolder(collectionPath, trackingId))
  );
}

async function findCaseContainer(
  displayName: string,
): Promise<MailFolder | null> {
  const root = await findFolder("/me/mailFolders", rootFolderName);

  if (!root) {
    return null;
  }

  return findFolder(
    `/me/mailFolders/${encodeURIComponent(root.id)}/childFolders`,
    displayName,
  );
}

async function ensureCaseContainer(
  displayName: string,
): Promise<MailFolder> {
  const root = await ensureFolder("/me/mailFolders", rootFolderName);

  return ensureFolder(
    `/me/mailFolders/${encodeURIComponent(root.id)}/childFolders`,
    displayName,
  );
}

export async function ensureActiveCaseFolder(
  trackingId: string,
): Promise<MailFolder> {
  const active = await ensureCaseContainer(activeFolderName);

  return ensureCaseFolder(
    `/me/mailFolders/${encodeURIComponent(active.id)}/childFolders`,
    trackingId,
  );
}

export async function getCaseFolderState(
  trackingId: string,
): Promise<CaseFolderState> {
  const [active, archived] = await Promise.all([
    findCaseContainer(activeFolderName),
    findCaseContainer(archivedFolderName),
  ]);
  const [activeCase, archivedCase] = await Promise.all([
    active
      ? findCaseFolder(
          `/me/mailFolders/${encodeURIComponent(active.id)}/childFolders`,
          trackingId,
        )
      : Promise.resolve(null),
    archived
      ? findCaseFolder(
          `/me/mailFolders/${encodeURIComponent(archived.id)}/childFolders`,
          trackingId,
        )
      : Promise.resolve(null),
  ]);

  if (activeCase && archivedCase) {
    throw new Error(
      `Case ${trackingId} exists in both Active and Archived. Resolve the duplicate folders before continuing.`,
    );
  }

  if (activeCase) {
    return { location: "active", folder: activeCase };
  }

  if (archivedCase) {
    return { location: "archived", folder: archivedCase };
  }

  return { location: "untracked", folder: null };
}

export async function moveCaseFolder(
  folderId: string,
  destination: Exclude<CaseLocation, "untracked">,
): Promise<MailFolder> {
  const destinationFolder = await ensureCaseContainer(
    destination === "active" ? activeFolderName : archivedFolderName,
  );

  return graphRequest<MailFolder>(
    `/me/mailFolders/${encodeURIComponent(folderId)}/move`,
    {
      method: "POST",
      body: JSON.stringify({
        destinationId: destinationFolder.id,
      }),
    },
  );
}

export function isFolderNameConflict(error: unknown): boolean {
  return (
    error instanceof GraphError &&
    (error.status === 409 || error.responseBody.includes("ErrorFolderExists"))
  );
}

// Moves the case folder, with all of its messages, to Deleted Items. This is a
// move rather than a permanent delete, so the user can still recover it.
export async function deleteCaseFolder(folder: MailFolder): Promise<void> {
  const folderPath = `/me/mailFolders/${encodeURIComponent(folder.id)}`;
  const moveToDeletedItems = () =>
    graphRequest<MailFolder>(`${folderPath}/move`, {
      method: "POST",
      body: JSON.stringify({ destinationId: "deleteditems" }),
    });

  try {
    await moveToDeletedItems();
  } catch (error) {
    if (!isFolderNameConflict(error)) {
      throw error;
    }

    // Deleted Items already holds a folder with this name, for example from an
    // earlier delete of the same case. Rename ours so the move can succeed. The
    // new name still starts with the tracking ID, so it keeps matching as a case
    // folder if the retry fails.
    const stamp = new Date().toISOString().slice(0, 16).replace("T", " ").replace(":", "-");

    await graphRequest<MailFolder>(folderPath, {
      method: "PATCH",
      body: JSON.stringify({
        displayName: `${folder.displayName} (deleted ${stamp})`,
      }),
    });
    await moveToDeletedItems();
  }
}