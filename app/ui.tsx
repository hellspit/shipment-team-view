// Shared styling and value rendering for the shipping-team dashboard.

import React from "react";

/** Friendlier headers than the raw SharePoint column names. */
export const LABELS: Record<string, string> = {
  Title: "Request ID",
  SubmittedAt: "Submitted",
  RequestorName: "Requestor",
  ShipmentType: "Type",
  ApprovalStatus: "Approval",
  ShipmentStatus: "Status",
  RequestedShipDate: "Ship by",
};

const PALETTE = {
  green: { background: "#dff6dd", color: "#0b6a0b" },
  amber: { background: "#fff4ce", color: "#855a00" },
  red: { background: "#fde7e9", color: "#a4262c" },
  blue: { background: "#eff6fc", color: "#0f6cbd" },
  grey: { background: "#f3f2f1", color: "#605e5c" },
  indigo: { background: "#eeedfb", color: "#3a31b0" },
  orange: { background: "#fed9cc", color: "#8a3707" },
} as const;

type Tone = keyof typeof PALETTE;

/** Same colour language as the SharePoint column formatting. */
function toneFor(column: string, value: string): Tone | null {
  if (!value) return null;
  if (column === "ApprovalStatus")
    return value === "Approved" ? "green"
      : value === "Auto-Approved" ? "blue"
      : value === "Rejected" ? "red"
      : value === "Pending" ? "amber" : "grey";
  if (column === "ShipmentStatus")
    return value === "Delivered" || value === "Delivery Confirmed" ? "green"
      : value === "Delayed" ? "red"
      : value === "In Transit" || value === "Out for Delivery" || value === "Carrier Pickup" ? "amber"
      : value === "Draft" || value === "Cancelled" ? "grey" : "indigo";
  if (column === "Priority")
    return value === "Critical" ? "red"
      : value === "High" ? "orange"
      : value === "Medium" ? "blue" : "grey";
  if (column === "SLAStatus")
    return value === "On Time" ? "green" : value === "At Risk" ? "amber" : value === "Late" ? "red" : "grey";
  return null;
}

export function pill(column: string, value: string): React.ReactElement | null {
  const tone = toneFor(column, value);
  if (!tone) return null;
  return <span style={{ ...S.pill, ...PALETTE[tone] }}>{value}</span>;
}

/** Dates from SharePoint arrive as ISO; show them the way a person reads them. */
export function fmtValue(column: string, value: unknown): string {
  const s = String(value ?? "");
  if (!s) return "—";
  if (/^\d{4}-\d{2}-\d{2}T/.test(s)) {
    const d = new Date(s);
    const withTime = column === "SubmittedAt" || column === "ApprovalDate";
    return d.toLocaleString("en-US", {
      month: "short", day: "2-digit", year: "numeric",
      ...(withTime ? { hour: "2-digit", minute: "2-digit" } : {}),
      timeZone: "America/New_York",
    });
  }
  return s;
}

export const S: Record<string, React.CSSProperties> = {
  page: { minHeight: "100vh", background: "#f4f4f8", padding: "32px 16px", fontFamily: "system-ui, sans-serif", color: "#1d1a2b" },
  wrap: { maxWidth: 1180, margin: "0 auto", display: "flex", flexDirection: "column", gap: 16 },
  header: { display: "flex", justifyContent: "space-between", alignItems: "flex-end", gap: 16, flexWrap: "wrap" },
  h1: { margin: 0, fontSize: 26, letterSpacing: "-.02em" },
  sub: { margin: "4px 0 0", color: "#6a6780", fontSize: 14.5 },
  count: { color: "#6a6780", fontSize: 13.5, fontWeight: 600 },
  toolbar: { display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center" },
  input: { padding: "9px 11px", borderRadius: 9, border: "1px solid #e0dfea", fontSize: 14, fontFamily: "inherit", background: "#fff", color: "#1d1a2b", boxSizing: "border-box", width: "100%" },
  textarea: { padding: "9px 11px", borderRadius: 9, border: "1px solid #e0dfea", fontSize: 14, fontFamily: "inherit", background: "#fff", color: "#1d1a2b", boxSizing: "border-box", width: "100%", minHeight: 64, resize: "vertical" },
  tableWrap: { background: "#fff", borderRadius: 14, boxShadow: "0 8px 30px rgba(0,0,0,.06)", overflowX: "auto" },
  table: { width: "100%", borderCollapse: "collapse", fontSize: 13.5, whiteSpace: "nowrap" },
  th: { textAlign: "left", padding: "12px 14px", borderBottom: "1px solid #eeedf4", color: "#6a6780", fontSize: 11.5, fontWeight: 700, textTransform: "uppercase", letterSpacing: ".04em" },
  tr: { borderBottom: "1px solid #f4f3f8" },
  td: { padding: "11px 14px", verticalAlign: "middle" },
  link: { color: "#3a31b0", fontWeight: 700, textDecoration: "none" },
  pill: { display: "inline-block", padding: "2px 10px", borderRadius: 12, fontSize: 12, fontWeight: 600 },
  blankTag: { display: "inline-block", padding: "2px 10px", borderRadius: 12, fontSize: 12, fontWeight: 600, background: "#fff4ce", color: "#855a00" },
  doneTag: { display: "inline-block", padding: "2px 10px", borderRadius: 12, fontSize: 12, fontWeight: 600, background: "#dff6dd", color: "#0b6a0b" },
  note: { color: "#6a6780", fontSize: 14 },
  error: { color: "#dc2626", fontWeight: 600, fontSize: 14 },
  card: { background: "#fff", borderRadius: 14, padding: 22, boxShadow: "0 8px 30px rgba(0,0,0,.06)", display: "flex", flexDirection: "column", gap: 14 },
  h2: { margin: 0, fontSize: 16.5, borderTop: "1px solid #eeedf4", paddingTop: 14 },
  grid2: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(230px, 1fr))", gap: 12 },
  label: { fontSize: 11.5, fontWeight: 700, color: "#6a6780", textTransform: "uppercase", letterSpacing: ".03em" },
  readonly: { padding: "9px 11px", borderRadius: 9, border: "1px solid #f0eff6", background: "#faf9fd", fontSize: 14, color: "#4a4760", minHeight: 38, boxSizing: "border-box", wordBreak: "break-word" },
  ghostBtn: { padding: "9px 14px", borderRadius: 9, border: "1px solid #e0dfea", background: "#fff", color: "#4a4760", fontWeight: 600, fontSize: 13.5, cursor: "pointer" },
  primaryBtn: { padding: "11px 22px", borderRadius: 10, border: "none", background: "#4338CA", color: "#fff", fontWeight: 700, fontSize: 15, cursor: "pointer" },
  itemCard: { border: "1px solid #e5e3f0", borderRadius: 12, padding: 16, background: "#faf9fd", display: "flex", flexDirection: "column", gap: 12 },
  itemNum: { fontWeight: 700, fontSize: 13.5, color: "#4338CA" },
  bar: { position: "sticky", bottom: 0, background: "#fff", borderTop: "1px solid #eeedf4", padding: "14px 22px", display: "flex", alignItems: "center", gap: 14, borderRadius: "0 0 14px 14px", flexWrap: "wrap" },
};
