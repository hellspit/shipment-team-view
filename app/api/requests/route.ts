// Every shipment request, one entry per Shipping Request ID, for the dashboard table.

import { fetchRows, groupRows, getSchema, TABLE_COLUMNS } from "@/lib/list";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function GET() {
  try {
    const [rows, schema] = await Promise.all([fetchRows(), getSchema()]);
    const requests = groupRows(rows).map((r) => {
      const summary: Record<string, unknown> = {};
      for (const name of TABLE_COLUMNS) summary[name] = r.fields[name] ?? "";
      return {
        requestId: r.requestId,
        lineCount: r.lineCount,
        // How much the shipping team still has to fill in.
        blanks: countBlanks(r.fields, schema),
        fields: summary,
      };
    });
    return Response.json({ requests, columns: TABLE_COLUMNS });
  } catch (e) {
    const msg = e instanceof Error ? e.message : "Could not load requests";
    console.error(JSON.stringify({ event: "requests_load_failed", msg }));
    return new Response(msg, { status: 502 });
  }
}

function countBlanks(
  fields: Record<string, unknown>,
  schema: Awaited<ReturnType<typeof getSchema>>
): number {
  let n = 0;
  for (const [name, field] of schema.byName) {
    if (field.locked || field.scope !== "request") continue;
    const v = fields[name];
    if (v === undefined || v === null || (typeof v === "string" && v.trim() === "")) n++;
  }
  return n;
}
