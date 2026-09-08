"use client";

import { useEffect, useMemo, useState } from "react";
import Link from "next/link";
import { S, pill, fmtValue, LABELS } from "@/app/ui";

type RequestRow = {
  requestId: string;
  lineCount: number;
  blanks: number;
  fields: Record<string, string | number>;
};

const FILTERS = [
  { key: "ShipmentStatus", label: "Shipment status" },
  { key: "ApprovalStatus", label: "Approval" },
  { key: "Priority", label: "Priority" },
  { key: "Carrier", label: "Carrier" },
] as const;

export default function Dashboard() {
  const [requests, setRequests] = useState<RequestRow[]>([]);
  const [columns, setColumns] = useState<string[]>([]);
  const [state, setState] = useState<"loading" | "ready" | "error">("loading");
  const [error, setError] = useState("");
  const [q, setQ] = useState("");
  const [picked, setPicked] = useState<Record<string, string>>({});

  useEffect(() => {
    (async () => {
      try {
        const res = await fetch("/api/requests");
        if (!res.ok) throw new Error(await res.text());
        const data = await res.json();
        setRequests(data.requests);
        setColumns(data.columns);
        setState("ready");
      } catch (e) {
        setError(e instanceof Error ? e.message : "Could not load");
        setState("error");
      }
    })();
  }, []);

  // Filter dropdowns are built from the values actually present, so they never offer a
  // choice that would return nothing.
  const options = useMemo(() => {
    const out: Record<string, string[]> = {};
    for (const f of FILTERS) {
      const values = new Set<string>();
      requests.forEach((r) => {
        const v = String(r.fields[f.key] ?? "").trim();
        if (v) values.add(v);
      });
      out[f.key] = [...values].sort();
    }
    return out;
  }, [requests]);

  const shown = useMemo(() => {
    const needle = q.trim().toLowerCase();
    return requests.filter((r) => {
      for (const [key, value] of Object.entries(picked)) {
        if (value && String(r.fields[key] ?? "") !== value) return false;
      }
      if (!needle) return true;
      return Object.values(r.fields).some((v) => String(v).toLowerCase().includes(needle));
    });
  }, [requests, q, picked]);

  return (
    <main style={S.page}>
      <div style={S.wrap}>
        <header style={S.header}>
          <div>
            <h1 style={S.h1}>Shipment requests</h1>
            <p style={S.sub}>
              Select a request to open it and fill in the shipping details.
            </p>
          </div>
          <span style={S.count}>
            {state === "ready" ? `${shown.length} of ${requests.length}` : ""}
          </span>
        </header>

        <div style={S.toolbar}>
          <input
            style={{ ...S.input, flex: "2 1 240px" }}
            placeholder="Search id, requestor, customer, carrier…"
            value={q}
            onChange={(e) => setQ(e.target.value)}
          />
          {FILTERS.map((f) => (
            <select
              key={f.key}
              style={{ ...S.input, flex: "1 1 150px" }}
              value={picked[f.key] ?? ""}
              onChange={(e) => setPicked({ ...picked, [f.key]: e.target.value })}
            >
              <option value="">{f.label}: all</option>
              {options[f.key]?.map((v) => (
                <option key={v} value={v}>{v}</option>
              ))}
            </select>
          ))}
          {(q || Object.values(picked).some(Boolean)) && (
            <button style={S.ghostBtn} onClick={() => { setQ(""); setPicked({}); }}>
              Clear
            </button>
          )}
        </div>

        {state === "loading" && <p style={S.note}>Loading requests…</p>}
        {state === "error" && <p style={S.error}>{error}</p>}

        {state === "ready" && (
          <div style={S.tableWrap}>
            <table style={S.table}>
              <thead>
                <tr>
                  {columns.map((c) => (
                    <th key={c} style={S.th}>{LABELS[c] ?? c}</th>
                  ))}
                  <th style={S.th}>Items</th>
                  <th style={S.th}>To fill</th>
                </tr>
              </thead>
              <tbody>
                {shown.map((r) => (
                  <tr key={r.requestId} style={S.tr}>
                    {columns.map((c) => (
                      <td key={c} style={S.td}>
                        {c === "Title" ? (
                          <Link href={`/request/${r.requestId}`} style={S.link}>
                            {String(r.fields[c] ?? "")}
                          </Link>
                        ) : pill(c, String(r.fields[c] ?? "")) ?? fmtValue(c, r.fields[c])}
                      </td>
                    ))}
                    <td style={S.td}>{r.lineCount}</td>
                    <td style={S.td}>
                      <span style={r.blanks ? S.blankTag : S.doneTag}>
                        {r.blanks ? `${r.blanks} blank` : "complete"}
                      </span>
                    </td>
                  </tr>
                ))}
                {!shown.length && (
                  <tr>
                    <td style={{ ...S.td, textAlign: "center", color: "#6a6780" }} colSpan={columns.length + 2}>
                      Nothing matches those filters.
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </main>
  );
}
