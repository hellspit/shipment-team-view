// Reads the "Shipment Requests" list and describes its columns to the UI.
//
// The list is FLAT: one row per line item, with the request-level fields repeated on every
// row of a request and LineNumber / LineCount tying them together. This module hides that
// shape from the UI by grouping rows back into requests, and by knowing which columns are
// request-level (shared by every row) and which are per-line-item.
//
// The field catalogue is built from the LIVE list schema rather than hard-coded, so adding a
// column in SharePoint makes it appear here without a code change. Only the section layout
// below is ours.

import { graph } from "@/lib/graph";

// ── Which columns belong to which section, and at which scope ──
// Any column not listed here still shows up, under "Other".
export type Scope = "request" | "item";

export const SECTIONS: { title: string; scope: Scope; columns: string[] }[] = [
  { title: "Request", scope: "request", columns: ["Title", "SubmittedAt", "LineCount", "AttachmentCount", "AttachmentsFolderUrl"] },
  { title: "Requestor", scope: "request", columns: ["RequestorName", "RequesterEmail", "Phone", "ProgramName", "BusinessJustification"] },
  { title: "Shipment", scope: "request", columns: ["ShipmentType", "Direction", "Purpose", "Priority", "RequestedShipDate", "RequiredDeliveryDate"] },
  { title: "References", scope: "request", columns: ["OrderNumber", "PONumber", "QuoteNumber"] },
  { title: "Ship-from", scope: "request", columns: ["ShipFromName", "ShipFromContact", "ShipFromAddress1", "ShipFromAddress2", "ShipFromCity", "ShipFromState", "ShipFromPostalCode", "ShipFromCountry", "ShipFromPhone", "ShipFromEmail"] },
  { title: "Ship-to", scope: "request", columns: ["ShipToName", "ShipToContact", "ShipToAddress1", "ShipToSuite", "ShipToAddress2", "ShipToCity", "ShipToState", "ShipToPostalCode", "ShipToCountry", "ShipToPhone", "ShipToEmail"] },
  { title: "Customer", scope: "request", columns: ["CustomerName", "CustomerNumber", "CustomerContact", "CustomerEmail"] },
  { title: "Vendor", scope: "request", columns: ["VendorName", "VendorNumber", "VendorContact", "VendorEmail"] },
  { title: "Package", scope: "request", columns: ["NumBoxes", "TotalWeight", "Dimensions", "PackageType"] },
  { title: "Shipping requirements", scope: "request", columns: ["ShippingSpeed", "Carrier", "BillingMethod", "ShippingService", "FreightClass", "ShippingCost"] },
  { title: "Special instructions", scope: "request", columns: ["PackingInstructions", "DeliveryInstructions", "CustomsNotes", "HandlingInstructions", "ReceivingRestrictions"] },
  { title: "Approval", scope: "request", columns: ["ApprovalStatus", "ApprovedBy", "ApprovalDate"] },
  { title: "Tracking", scope: "request", columns: ["ShipmentStatus", "AssignedShippingCoordinator", "TrackingNumber", "TrackingLink", "ShipDate", "ExpectedDeliveryDate", "DeliveryDate", "SLAStatus"] },
  { title: "Line item", scope: "item", columns: ["PartNumber", "ItemDescription", "Manufacturer", "Quantity", "UnitValue", "Currency", "ItemWeight", "CountryOfOrigin", "SerialNumber", "LotNumber", "ExportClassification", "CommodityCode", "AssetTag"] },
];

/** Written by the intake form or by SharePoint itself — never editable here. */
export const LOCKED = new Set([
  "Title", "SubmissionId", "SubmittedAt", "LineNumber", "LineCount",
  "AttachmentCount", "AttachmentsFolderUrl",
]);

/** Columns the dashboard table shows, in order. */
export const TABLE_COLUMNS = [
  "Title", "SubmittedAt", "RequestorName", "ShipmentType", "Priority",
  "ApprovalStatus", "ShipmentStatus", "Carrier", "RequestedShipDate",
];

export type FieldType = "text" | "multiline" | "choice" | "number" | "date" | "datetime" | "boolean";

export type Field = {
  name: string;
  label: string;
  type: FieldType;
  choices?: string[];
  maxLength?: number;
  scope: Scope;
  locked: boolean;
};

export type Section = { title: string; scope: Scope; fields: Field[] };

type GraphColumn = {
  name: string;
  displayName: string;
  readOnly?: boolean;
  hidden?: boolean;
  text?: { allowMultipleLines?: boolean; maxLength?: number };
  choice?: { choices?: string[] };
  number?: unknown;
  currency?: unknown;
  dateTime?: { format?: string };
  boolean?: unknown;
};

function typeOf(col: GraphColumn): FieldType | null {
  if (col.choice) return "choice";
  if (col.number || col.currency) return "number";
  if (col.dateTime) return col.dateTime.format === "dateOnly" ? "date" : "datetime";
  if (col.boolean) return "boolean";
  if (col.text) return col.text.allowMultipleLines ? "multiline" : "text";
  return null;
}

// ── list id resolution (mirrors the intake form) ──
let listCache: { id: string; exp: number } | null = null;

export function siteId(): string {
  const id = process.env.SHAREPOINT_SITE_ID?.trim();
  if (!id) throw new Error("SHAREPOINT_SITE_ID not configured");
  return id;
}

export async function getListId(): Promise<string> {
  const explicit = process.env.SHAREPOINT_LIST_ID?.trim();
  if (explicit) return explicit;
  if (listCache && Date.now() < listCache.exp) return listCache.id;

  const name = process.env.SHAREPOINT_LIST_NAME?.trim() || "Shipment Requests";
  const j = await graph<{ value: { id: string; name: string; displayName: string }[] }>(
    `/sites/${siteId()}/lists?$select=id,name,displayName`
  );
  const hit = j.value.find(
    (l) => l.displayName?.toLowerCase() === name.toLowerCase() || l.name?.toLowerCase() === name.toLowerCase()
  );
  if (!hit) throw new Error(`SharePoint list "${name}" not found`);
  listCache = { id: hit.id, exp: Date.now() + 3_600_000 };
  return hit.id;
}

export async function listBase(): Promise<string> {
  return `/sites/${siteId()}/lists/${await getListId()}`;
}

// ── the field catalogue, cached ──
let schemaCache: { sections: Section[]; byName: Map<string, Field>; exp: number } | null = null;

export async function getSchema(): Promise<{ sections: Section[]; byName: Map<string, Field> }> {
  if (schemaCache && Date.now() < schemaCache.exp) return schemaCache;

  const base = await listBase();
  const j = await graph<{ value: GraphColumn[] }>(`${base}/columns`);
  const usable = new Map<string, GraphColumn>();
  for (const col of j.value ?? []) {
    if (col.readOnly || col.hidden) continue;
    if (col.name === "ContentType" || col.name === "Attachments") continue;
    if (!typeOf(col)) continue;
    usable.set(col.name, col);
  }

  const toField = (col: GraphColumn, scope: Scope): Field => ({
    name: col.name,
    label: col.displayName || col.name,
    type: typeOf(col)!,
    choices: col.choice?.choices,
    maxLength: col.text?.maxLength,
    scope,
    locked: LOCKED.has(col.name),
  });

  const placed = new Set<string>();
  const sections: Section[] = [];
  for (const s of SECTIONS) {
    const fields: Field[] = [];
    for (const name of s.columns) {
      const col = usable.get(name);
      if (!col) continue; // column was removed from the list — skip rather than break
      fields.push(toField(col, s.scope));
      placed.add(name);
    }
    if (fields.length) sections.push({ title: s.title, scope: s.scope, fields });
  }

  // Anything added to the list but not to SECTIONS still shows, at request scope.
  const leftovers = [...usable.values()].filter((c) => !placed.has(c.name) && c.name !== "SubmissionId");
  if (leftovers.length) {
    sections.push({ title: "Other", scope: "request", fields: leftovers.map((c) => toField(c, "request")) });
  }

  const byName = new Map<string, Field>();
  sections.forEach((s) => s.fields.forEach((f) => byName.set(f.name, f)));

  schemaCache = { sections, byName, exp: Date.now() + 600_000 };
  return schemaCache;
}

// ── reading rows ──
export type Row = { id: string; fields: Record<string, unknown> };

/** Every row in the list, newest request first. */
export async function fetchRows(): Promise<Row[]> {
  const base = await listBase();
  const rows: Row[] = [];
  let url = `${base}/items?$expand=fields&$top=200`;
  // Follow paging so a list bigger than one page still loads completely.
  for (let page = 0; page < 25 && url; page++) {
    const j = await graph<{ value: Row[]; "@odata.nextLink"?: string }>(url);
    rows.push(...(j.value ?? []));
    const next = j["@odata.nextLink"];
    url = next ? next.replace("https://graph.microsoft.com/v1.0", "") : "";
  }
  return rows;
}

export type RequestSummary = {
  requestId: string;
  rowIds: string[];
  lineCount: number;
  fields: Record<string, unknown>;
};

/** Collapse the flat rows into one entry per Shipping Request ID. */
export function groupRows(rows: Row[]): RequestSummary[] {
  const byRequest = new Map<string, Row[]>();
  for (const r of rows) {
    const title = String(r.fields?.Title ?? "").trim();
    if (!title) continue;
    byRequest.set(title, (byRequest.get(title) ?? []).concat(r));
  }
  const out: RequestSummary[] = [];
  for (const [requestId, group] of byRequest) {
    group.sort((a, b) => Number(a.fields.LineNumber ?? 0) - Number(b.fields.LineNumber ?? 0));
    out.push({
      requestId,
      rowIds: group.map((g) => g.id),
      lineCount: group.length,
      fields: group[0].fields,
    });
  }
  // Newest first, by request id (they're sequential) — falls back to submit time.
  out.sort((a, b) => b.requestId.localeCompare(a.requestId));
  return out;
}

/** True when a column holds nothing — the rule for "this cell is still fillable". */
export function isEmpty(value: unknown): boolean {
  if (value === undefined || value === null) return true;
  if (typeof value === "string") return value.trim() === "";
  return false;
}

/** Minted by the intake form; the only request ids the API routes accept. */
export const REQUEST_ID_RE = /^SHIP-\d{4}-\d{6}$/;

/** Every row of one request, in line order. */
export async function rowsFor(requestId: string): Promise<Row[]> {
  const rows = (await fetchRows()).filter((r) => String(r.fields?.Title ?? "").trim() === requestId);
  rows.sort((a, b) => Number(a.fields.LineNumber ?? 0) - Number(b.fields.LineNumber ?? 0));
  return rows;
}
