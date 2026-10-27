// Attachments for a shipment request.
//
// The intake form creates one folder per request — /Shipment Requests/<date> <sender> in the
// "Shipment Attachments" library (SHIP-YYYY-NNNNNN for older requests) — and stamps its link
// into AttachmentsFolderUrl on every row of the request. We find the folder from that stored
// link instead of rebuilding its path, so this app needs no library settings of its own.
//
// As with the intake form, file bytes never pass through this server: the browser gets a
// pre-authorized Graph upload URL per file and PUTs straight to SharePoint.

import { graph } from "@/lib/graph";

// Keep in sync with the intake form (shipment-form/app/page.tsx).
export const ACCEPTED_EXTENSIONS = ["pdf", "xlsx", "docx", "jpg", "jpeg", "png"];
export const MAX_FILE_BYTES = 25 * 1024 * 1024;
export const MAX_FILES = 20;

export type Folder = { driveId: string; id: string };

/** Graph's /shares endpoint takes any SharePoint URL encoded as "u!" + unpadded base64url. */
function shareId(url: string): string {
  return "u!" + Buffer.from(url).toString("base64").replace(/=+$/, "").replace(/\//g, "_").replace(/\+/g, "-");
}

/** The intake form's parent folder for every request folder. Same setting, same default. */
function rootFolderName(): string {
  return (process.env.SHAREPOINT_ROOT_FOLDER?.trim() || "Shipment Requests").replace(/^\/+|\/+$/g, "");
}

/**
 * The request's attachment folder, from the link on its list rows.
 *
 * Older requests' folders are named after the request id; newer ones are named
 * "<date> <sender>". Either way the folder must sit directly in the requests root folder —
 * a row whose link was edited to point somewhere else in the library is refused rather than
 * letting uploads land there.
 */
export async function resolveFolder(folderUrl: string, requestId: string): Promise<Folder> {
  const item = await graph<{
    id: string;
    name?: string;
    folder?: unknown;
    parentReference?: { driveId?: string; path?: string };
  }>(`/shares/${shareId(folderUrl)}/driveItem?$select=id,name,folder,parentReference`);

  const driveId = item.parentReference?.driveId;
  if (!item.folder || !driveId) throw new Error("The attachment link on this request doesn't point at a folder");
  if (item.name === requestId) return { driveId, id: item.id };

  // parentReference.path looks like "/drives/<id>/root:/Shipment Requests" (sometimes encoded).
  let parent = item.parentReference?.path ?? "";
  try {
    parent = decodeURIComponent(parent);
  } catch {
    /* already plain */
  }
  const parentFolder = parent.split("root:")[1]?.replace(/^\/+|\/+$/g, "") ?? "";
  if (parentFolder.toLowerCase() !== rootFolderName().toLowerCase() || item.name?.startsWith("_")) {
    throw new Error(`The attachment link on ${requestId} points outside the "${rootFolderName()}" folder`);
  }
  return { driveId, id: item.id };
}

/** SharePoint rejects " * : < > ? / \ | and leading/trailing spaces & dots. */
export function safeFileName(name: string): string {
  const cleaned = name
    .replace(/[\"*:<>?\/\\|#%]/g, "-")
    .replace(/\s+/g, " ")
    .trim()
    .replace(/^\.+|\.+$/g, "")
    .slice(0, 180);
  return cleaned || "attachment";
}

export function extensionOf(name: string): string {
  const dot = name.lastIndexOf(".");
  return dot === -1 ? "" : name.slice(dot + 1).toLowerCase();
}

/** A pre-authorized URL the browser can PUT one file to. A clashing name gets renamed, not overwritten. */
export async function createUploadSession(folder: Folder, name: string): Promise<string> {
  const session = await graph<{ uploadUrl: string }>(
    `/drives/${folder.driveId}/items/${folder.id}:/${encodeURIComponent(name)}:/createUploadSession`,
    {
      method: "POST",
      body: JSON.stringify({ item: { "@microsoft.graph.conflictBehavior": "rename" } }),
    }
  );
  return session.uploadUrl;
}

/** Files (not sub-folders) currently in the folder — the source of truth for AttachmentCount. */
export async function countFiles(folder: Folder): Promise<number> {
  let count = 0;
  let url = `/drives/${folder.driveId}/items/${folder.id}/children?$select=id,file&$top=200`;
  for (let page = 0; page < 50 && url; page++) {
    const j = await graph<{ value: { file?: unknown }[]; "@odata.nextLink"?: string }>(url);
    count += (j.value ?? []).filter((child) => child.file).length;
    const next = j["@odata.nextLink"];
    url = next ? next.replace("https://graph.microsoft.com/v1.0", "") : "";
  }
  return count;
}
