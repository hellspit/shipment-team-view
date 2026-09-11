// Recount a request's attachment folder and write the number to AttachmentCount on every row.
//
// Counting the folder, rather than adding the number just uploaded, keeps the column right
// even when an upload partly failed or two people added files at the same time.

import { graph } from "@/lib/graph";
import { listBase, rowsFor, REQUEST_ID_RE } from "@/lib/list";
import { countFiles, resolveFolder } from "@/lib/attachments";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function POST(_request: Request, ctx: { params: Promise<{ requestId: string }> }) {
  const { requestId } = await ctx.params;
  if (!REQUEST_ID_RE.test(requestId)) return new Response("Invalid request id", { status: 400 });

  try {
    const rows = await rowsFor(requestId);
    if (!rows.length) return new Response("Request not found", { status: 404 });

    const folderUrl = String(rows[0].fields.AttachmentsFolderUrl ?? "").trim();
    if (!folderUrl) return new Response("This request has no attachment folder", { status: 409 });

    const folder = await resolveFolder(folderUrl, requestId);
    const count = await countFiles(folder);

    // Only touch rows whose count is actually out of date.
    const base = await listBase();
    const stale = rows.filter((r) => Number(r.fields.AttachmentCount ?? -1) !== count);
    for (const r of stale) {
      await graph(`${base}/items/${r.id}/fields`, {
        method: "PATCH",
        body: JSON.stringify({ AttachmentCount: count }),
      });
    }

    console.log(JSON.stringify({ event: "attachments_synced", requestId, count, updatedRows: stale.length }));
    return Response.json({ count, updatedRows: stale.length });
  } catch (e) {
    const msg = e instanceof Error ? e.message : "Could not update the attachment count";
    console.error(JSON.stringify({ event: "attachments_sync_failed", requestId, msg }));
    return new Response(msg, { status: 502 });
  }
}
