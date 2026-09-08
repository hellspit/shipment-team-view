// One shipment request: read it for the detail form, and save the shipping team's edits.
//
// GET   → the request's rows plus the field catalogue, so the form knows what to render
//         and which cells are still blank.
// PATCH → write the filled-in values back.
//
// Request-level columns are stored on EVERY row of the request, so a change to one is
// written to all of them; line-item columns are written only to their own row.
//
// Only cells that are currently empty may be written. That rule is enforced here, not just
// in the UI — a stale browser tab can't overwrite something another person filled in first.

import { graph, type GraphError } from "@/lib/graph";
import { fetchRows, getSchema, listBase, isEmpty, type Row } from "@/lib/list";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

const REQUEST_ID_RE = /^SHIP-\d{4}-\d{6}$/;
const MAX_BODY_BYTES = 200_000;

async function rowsFor(requestId: string): Promise<Row[]> {
  const rows = (await fetchRows()).filter((r) => String(r.fields?.Title ?? "").trim() === requestId);
  rows.sort((a, b) => Number(a.fields.LineNumber ?? 0) - Number(b.fields.LineNumber ?? 0));
  return rows;
}

export async function GET(_request: Request, ctx: { params: Promise<{ requestId: string }> }) {
  const { requestId } = await ctx.params;
  if (!REQUEST_ID_RE.test(requestId)) return new Response("Invalid request id", { status: 400 });

  try {
    const [rows, schema] = await Promise.all([rowsFor(requestId), getSchema()]);
    if (!rows.length) return new Response("Request not found", { status: 404 });

    return Response.json({
      requestId,
      sections: schema.sections,
      rows: rows.map((r) => ({
        id: r.id,
        lineNumber: Number(r.fields.LineNumber ?? 0),
        fields: r.fields,
      })),
    });
  } catch (e) {
    const msg = e instanceof Error ? e.message : "Could not load the request";
    console.error(JSON.stringify({ event: "request_load_failed", requestId, msg }));
    return new Response(msg, { status: 502 });
  }
}

type Edits = {
  // column -> value, applied to every row of the request
  request?: Record<string, unknown>;
  // row id -> { column -> value }
  items?: Record<string, Record<string, unknown>>;
};

export async function PATCH(request: Request, ctx: { params: Promise<{ requestId: string }> }) {
  const { requestId } = await ctx.params;
  if (!REQUEST_ID_RE.test(requestId)) return new Response("Invalid request id", { status: 400 });

  const raw = await request.text();
  if (raw.length > MAX_BODY_BYTES) return new Response("Payload too large", { status: 413 });

  let edits: Edits;
  try {
    edits = JSON.parse(raw);
  } catch {
    return new Response("Bad JSON", { status: 400 });
  }

  try {
    const [rows, schema] = await Promise.all([rowsFor(requestId), getSchema()]);
    if (!rows.length) return new Response("Request not found", { status: 404 });

    const byId = new Map(rows.map((r) => [r.id, r]));
    // rowId -> the columns to write on it
    const writes = new Map<string, Record<string, unknown>>();
    const rejected: string[] = [];

    const stage = (rowId: string, column: string, value: unknown) => {
      const current = writes.get(rowId) ?? {};
      current[column] = value;
      writes.set(rowId, current);
    };

    // Request-level edits go onto every row of the request.
    for (const [column, value] of Object.entries(edits.request ?? {})) {
      const field = schema.byName.get(column);
      if (!field || field.locked || field.scope !== "request") {
        rejected.push(`${column}: not editable`);
        continue;
      }
      const clean = coerce(value, field.type);
      if (clean === undefined) {
        rejected.push(`${column}: invalid value`);
        continue;
      }
      if (field.type === "choice" && clean !== "" && !field.choices?.includes(String(clean))) {
        rejected.push(`${column}: "${clean}" is not one of the allowed choices`);
        continue;
      }
      const occupied = rows.find((r) => !isEmpty(r.fields[column]));
      if (occupied) {
        rejected.push(`${column}: already has a value (${String(occupied.fields[column]).slice(0, 40)})`);
        continue;
      }
      for (const r of rows) stage(r.id, column, clean);
    }

    // Line-item edits go onto their own row only.
    for (const [rowId, columns] of Object.entries(edits.items ?? {})) {
      const row = byId.get(rowId);
      if (!row) {
        rejected.push(`row ${rowId}: not part of this request`);
        continue;
      }
      for (const [column, value] of Object.entries(columns)) {
        const field = schema.byName.get(column);
        if (!field || field.locked || field.scope !== "item") {
          rejected.push(`${column}: not editable`);
          continue;
        }
        const clean = coerce(value, field.type);
        if (clean === undefined) {
          rejected.push(`${column}: invalid value`);
          continue;
        }
        if (field.type === "choice" && clean !== "" && !field.choices?.includes(String(clean))) {
          rejected.push(`${column}: "${clean}" is not one of the allowed choices`);
          continue;
        }
        if (!isEmpty(row.fields[column])) {
          rejected.push(`${column} on line ${row.fields.LineNumber}: already has a value`);
          continue;
        }
        stage(row.id, column, clean);
      }
    }

    if (!writes.size) {
      return Response.json({ ok: rejected.length === 0, updatedRows: 0, written: 0, rejected });
    }

    const base = await listBase();
    let written = 0;
    for (const [rowId, fields] of writes) {
      await withRetry(() =>
        graph(`${base}/items/${rowId}/fields`, { method: "PATCH", body: JSON.stringify(fields) })
      );
      written += Object.keys(fields).length;
    }

    console.log(JSON.stringify({ event: "request_updated", requestId, rows: writes.size, written }));
    return Response.json({ ok: true, updatedRows: writes.size, written, rejected });
  } catch (e) {
    const msg = e instanceof Error ? e.message : "Could not save";
    console.error(JSON.stringify({ event: "request_save_failed", requestId, msg }));
    return new Response(msg, { status: 502 });
  }
}

/** Shape a submitted value for its column, or undefined when it can't be used. */
function coerce(value: unknown, type: string): unknown {
  if (value === null || value === undefined) return undefined;
  if (type === "boolean") return Boolean(value);

  const s = String(value).trim();
  if (s === "") return undefined; // nothing typed — not an edit

  switch (type) {
    case "number": {
      const n = Number(s);
      return Number.isFinite(n) ? n : undefined;
    }
    case "date":
      return /^\d{4}-\d{2}-\d{2}$/.test(s) ? s : undefined;
    case "datetime": {
      const d = new Date(s);
      return Number.isNaN(d.getTime()) ? undefined : d.toISOString();
    }
    default:
      return s;
  }
}

/** SharePoint throttles with 429/503; back off and retry a couple of times. */
async function withRetry<T>(fn: () => Promise<T>): Promise<T> {
  for (let attempt = 0; ; attempt++) {
    try {
      return await fn();
    } catch (e) {
      const status = (e as GraphError).status;
      if (attempt >= 2 || (status !== 429 && status !== 503)) throw e;
      await new Promise((r) => setTimeout(r, 1000 * (attempt + 1)));
    }
  }
}
