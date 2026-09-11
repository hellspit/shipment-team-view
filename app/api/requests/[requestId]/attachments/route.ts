// Add attachments to a shipment request from the shipping-team page.
//
// POST { files: [{ name, size }] } → { uploads: [{ name, uploadUrl }] }
//
// The browser then PUTs each file to its uploadUrl directly and calls ./sync, which recounts
// the folder and updates AttachmentCount.

import { rowsFor, REQUEST_ID_RE } from "@/lib/list";
import {
  ACCEPTED_EXTENSIONS, MAX_FILE_BYTES, MAX_FILES,
  createUploadSession, extensionOf, resolveFolder, safeFileName,
} from "@/lib/attachments";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

type IncomingFile = { name?: unknown; size?: unknown };

export async function POST(request: Request, ctx: { params: Promise<{ requestId: string }> }) {
  const { requestId } = await ctx.params;
  if (!REQUEST_ID_RE.test(requestId)) return new Response("Invalid request id", { status: 400 });

  let body: { files?: IncomingFile[] };
  try {
    body = await request.json();
  } catch {
    return new Response("Bad JSON", { status: 400 });
  }

  const files = Array.isArray(body.files) ? body.files : [];
  if (!files.length) return new Response("No files", { status: 400 });
  if (files.length > MAX_FILES) return new Response(`Too many files (max ${MAX_FILES})`, { status: 400 });
  for (const f of files) {
    if (typeof f?.name !== "string" || !f.name.trim())
      return new Response("Every file needs a name", { status: 400 });
    if (!ACCEPTED_EXTENSIONS.includes(extensionOf(f.name)))
      return new Response(`"${f.name}" isn't an accepted type (PDF, XLSX, DOCX, JPG, PNG)`, { status: 400 });
    if (typeof f?.size !== "number" || f.size < 0 || f.size > MAX_FILE_BYTES)
      return new Response(`"${f.name}" is over 25 MB`, { status: 400 });
  }

  try {
    const rows = await rowsFor(requestId);
    if (!rows.length) return new Response("Request not found", { status: 404 });

    const folderUrl = String(rows[0].fields.AttachmentsFolderUrl ?? "").trim();
    if (!folderUrl) return new Response("This request has no attachment folder", { status: 409 });

    const folder = await resolveFolder(folderUrl, requestId);
    const uploads = await Promise.all(
      files.map(async (f) => {
        const name = safeFileName(String(f.name));
        return { name, uploadUrl: await createUploadSession(folder, name) };
      })
    );

    console.log(JSON.stringify({ event: "attachments_started", requestId, files: uploads.length }));
    return Response.json({ uploads });
  } catch (e) {
    const msg = e instanceof Error ? e.message : "Could not prepare the upload";
    console.error(JSON.stringify({ event: "attachments_start_failed", requestId, msg }));
    return new Response(msg, { status: 502 });
  }
}
