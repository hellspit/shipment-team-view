"use client";

import { use, useEffect, useMemo, useState } from "react";
import Link from "next/link";
import { S, pill, fmtValue } from "@/app/ui";
import AddAttachment from "./AddAttachment";

type FieldType = "text" | "multiline" | "choice" | "number" | "date" | "datetime" | "boolean";
type Field = {
  name: string;
  label: string;
  type: FieldType;
  choices?: string[];
  maxLength?: number;
  scope: "request" | "item";
  locked: boolean;
};
type Section = { title: string; scope: "request" | "item"; fields: Field[] };
type RowData = { id: string; lineNumber: number; fields: Record<string, unknown> };

const isBlank = (v: unknown) =>
  v === undefined || v === null || (typeof v === "string" && v.trim() === "");

/** A cell this page can write to: not system-owned, and still empty. */
const isFillable = (field: Field, value: unknown) => !field.locked && isBlank(value);

/** SharePoint returns dates as ISO; the date inputs want their own formats. */
function forInput(value: unknown, type: FieldType): string {
  const s = String(value ?? "");
  if (!s) return "";
  if (type === "date") return s.slice(0, 10);
  if (type === "datetime") return s.slice(0, 16);
  return s;
}

export default function RequestPage({ params }: { params: Promise<{ requestId: string }> }) {
  const { requestId } = use(params);

  const [sections, setSections] = useState<Section[]>([]);
  const [rows, setRows] = useState<RowData[]>([]);
  const [state, setState] = useState<"loading" | "ready" | "error">("loading");
  const [error, setError] = useState("");

  // Only what the user typed, so we never resubmit values that already exist.
  const [reqEdits, setReqEdits] = useState<Record<string, string>>({});
  const [itemEdits, setItemEdits] = useState<Record<string, Record<string, string>>>({});

  const [saving, setSaving] = useState(false);
  const [message, setMessage] = useState("");
  const [messageTone, setMessageTone] = useState<"ok" | "bad">("ok");

  // "Only blanks" hides everything the requestor already filled, leaving just the cells
  // this page can actually write to. The choice is remembered per browser.
  //
  // Read lazily rather than in an effect: the toggle is only rendered once the request has
  // loaded client-side, so it never appears in the server-rendered HTML and can't mismatch
  // on hydration. localStorage can throw outright in private mode, hence the try.
  const [onlyBlanks, setOnlyBlanks] = useState(() => {
    if (typeof window === "undefined") return false;
    try {
      return localStorage.getItem("shipment.onlyBlanks") === "1";
    } catch {
      return false;
    }
  });

  const toggleOnlyBlanks = (next: boolean) => {
    setOnlyBlanks(next);
    try {
      localStorage.setItem("shipment.onlyBlanks", next ? "1" : "0");
    } catch {
      /* not worth failing the page over */
    }
  };

  // Never flips back to "loading" on a refresh — after a save the page stays put and the
  // saved cells simply turn read-only.
  const load = async () => {
    try {
      const res = await fetch(`/api/requests/${requestId}`);
      if (!res.ok) throw new Error(await res.text());
      const data = await res.json();
      setSections(data.sections);
      setRows(data.rows);
      setReqEdits({});
      setItemEdits({});
      setState("ready");
    } catch (e) {
      setError(e instanceof Error ? e.message : "Could not load");
      setState("error");
    }
  };

  useEffect(() => {
    (async () => {
      await load();
    })();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [requestId]);

  // Request-level values live on every row; the first row is representative.
  const head = useMemo(() => rows[0]?.fields ?? {}, [rows]);

  const blanksLeft = useMemo(() => {
    let n = 0;
    for (const s of sections) {
      if (s.scope === "request") {
        n += s.fields.filter((f) => !f.locked && isBlank(head[f.name])).length;
      } else {
        for (const r of rows) n += s.fields.filter((f) => !f.locked && isBlank(r.fields[f.name])).length;
      }
    }
    return n;
  }, [sections, rows, head]);

  const editCount =
    Object.values(reqEdits).filter((v) => v !== "").length +
    Object.values(itemEdits).reduce((sum, cols) => sum + Object.values(cols).filter((v) => v !== "").length, 0);

  const save = async () => {
    setSaving(true);
    setMessage("");
    try {
      const request: Record<string, string> = {};
      Object.entries(reqEdits).forEach(([k, v]) => { if (v !== "") request[k] = v; });
      const items: Record<string, Record<string, string>> = {};
      Object.entries(itemEdits).forEach(([rowId, cols]) => {
        const kept: Record<string, string> = {};
        Object.entries(cols).forEach(([k, v]) => { if (v !== "") kept[k] = v; });
        if (Object.keys(kept).length) items[rowId] = kept;
      });

      const res = await fetch(`/api/requests/${requestId}`, {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ request, items }),
      });
      if (!res.ok) throw new Error(await res.text());
      const data = (await res.json()) as { written: number; updatedRows: number; rejected: string[] };

      if (data.rejected?.length) {
        setMessageTone("bad");
        setMessage(`Saved ${data.written} field(s). Skipped: ${data.rejected.join("; ")}`);
      } else {
        setMessageTone("ok");
        setMessage(`Saved ${data.written} field(s) across ${data.updatedRows} row(s).`);
      }
      await load(); // re-read so saved cells become read-only
    } catch (e) {
      setMessageTone("bad");
      setMessage(e instanceof Error ? e.message : "Could not save");
    } finally {
      setSaving(false);
    }
  };

  if (state === "loading") return <main style={S.page}><div style={S.wrap}><p style={S.note}>Loading {requestId}…</p></div></main>;
  if (state === "error") return <main style={S.page}><div style={S.wrap}><p style={S.error}>{error}</p><Link href="/" style={S.link}>← Back</Link></div></main>;

  const folder = String(head.AttachmentsFolderUrl ?? "");

  return (
    <main style={S.page}>
      <div style={{ ...S.wrap, maxWidth: 900 }}>
        <header style={S.header}>
          <div>
            <Link href="/" style={{ ...S.link, fontSize: 13.5 }}>← All requests</Link>
            <h1 style={{ ...S.h1, marginTop: 6 }}>{requestId}</h1>
            <p style={S.sub}>
              {String(head.RequestorName ?? "—")} · {rows.length} line item{rows.length === 1 ? "" : "s"} ·{" "}
              {blanksLeft ? `${blanksLeft} field${blanksLeft === 1 ? "" : "s"} still blank` : "nothing left to fill"}
            </p>
          </div>
          <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
            {pill("Priority", String(head.Priority ?? ""))}
            {pill("ApprovalStatus", String(head.ApprovalStatus ?? ""))}
            {pill("ShipmentStatus", String(head.ShipmentStatus ?? ""))}
          </div>
        </header>

        <div style={S.toolbar}>
          <div style={S.segmented}>
            <button type="button" onClick={() => toggleOnlyBlanks(false)}
              style={{ ...S.segBtn, ...(onlyBlanks ? null : S.segBtnOn) }}>
              All fields
            </button>
            <button type="button" onClick={() => toggleOnlyBlanks(true)}
              style={{ ...S.segBtn, ...(onlyBlanks ? S.segBtnOn : null) }}>
              Only blanks{blanksLeft ? ` (${blanksLeft})` : ""}
            </button>
          </div>
          {folder && (
            <>
              <a href={folder} target="_blank" rel="noopener noreferrer" style={S.link}>Open attachments</a>
              <AddAttachment requestId={requestId} onUploaded={load} />
            </>
          )}
        </div>

        <p style={S.note}>
          {onlyBlanks
            ? "Showing only the cells still waiting to be filled."
            : "Grey cells were filled by the requestor and can’t be changed here. Empty cells are yours to fill."}
        </p>

        <div style={S.card}>
          {onlyBlanks && blanksLeft === 0 && (
            <p style={{ ...S.note, padding: "18px 0", textAlign: "center" }}>
              Nothing left to fill on this request.
            </p>
          )}

          {sections.filter((s) => s.scope === "request").map((section) => {
            const fields = onlyBlanks
              ? section.fields.filter((f) => isFillable(f, head[f.name]))
              : section.fields;
            if (!fields.length) return null;
            return (
              <section key={section.title}>
                <h2 style={S.h2}>{section.title}</h2>
                <div style={{ ...S.grid2, marginTop: 12 }}>
                  {fields.map((f) => (
                    <FieldBox
                      key={f.name}
                      field={f}
                      value={head[f.name]}
                      draft={reqEdits[f.name] ?? ""}
                      onChange={(v) => setReqEdits({ ...reqEdits, [f.name]: v })}
                    />
                  ))}
                </div>
              </section>
            );
          })}

          {sections.filter((s) => s.scope === "item").map((section) => {
            // A line only appears if it still has something to fill.
            const visibleRows = onlyBlanks
              ? rows.filter((r) => section.fields.some((f) => isFillable(f, r.fields[f.name])))
              : rows;
            if (!visibleRows.length) return null;
            return (
              <section key={section.title}>
                <h2 style={S.h2}>{section.title}s</h2>
                <div style={{ display: "flex", flexDirection: "column", gap: 12, marginTop: 12 }}>
                  {visibleRows.map((row) => {
                    const fields = onlyBlanks
                      ? section.fields.filter((f) => isFillable(f, row.fields[f.name]))
                      : section.fields;
                    return (
                      <div key={row.id} style={S.itemCard}>
                        <span style={S.itemNum}>
                          Line {row.lineNumber} · {String(row.fields.ItemDescription ?? "")}
                        </span>
                        <div style={S.grid2}>
                          {fields.map((f) => (
                            <FieldBox
                              key={f.name}
                              field={f}
                              value={row.fields[f.name]}
                              draft={itemEdits[row.id]?.[f.name] ?? ""}
                              onChange={(v) =>
                                setItemEdits({
                                  ...itemEdits,
                                  [row.id]: { ...(itemEdits[row.id] ?? {}), [f.name]: v },
                                })
                              }
                            />
                          ))}
                        </div>
                      </div>
                    );
                  })}
                </div>
              </section>
            );
          })}

          <div style={S.bar}>
            <button style={{ ...S.primaryBtn, opacity: saving || !editCount ? 0.55 : 1 }}
              onClick={save} disabled={saving || !editCount}>
              {saving ? "Saving…" : editCount ? `Save ${editCount} field${editCount === 1 ? "" : "s"}` : "Nothing to save"}
            </button>
            {message && (
              <span style={{ color: messageTone === "ok" ? "#16a34a" : "#dc2626", fontWeight: 600, fontSize: 13.5 }}>
                {message}
              </span>
            )}
          </div>
        </div>
      </div>
    </main>
  );
}

/** One cell: read-only when the requestor already filled it, an input when it's blank. */
function FieldBox({
  field, value, draft, onChange,
}: {
  field: Field;
  value: unknown;
  draft: string;
  onChange: (v: string) => void;
}) {
  const filled = !isBlank(value);
  const editable = !field.locked && !filled;

  return (
    <label style={{ display: "flex", flexDirection: "column", gap: 5 }}>
      <span style={S.label}>{field.label}</span>

      {!editable ? (
        <div style={S.readonly}>
          {pill(field.name, String(value ?? "")) ?? (filled ? fmtValue(field.name, value) : "—")}
        </div>
      ) : field.type === "choice" ? (
        <select style={S.input} value={draft} onChange={(e) => onChange(e.target.value)}>
          <option value="">— select —</option>
          {field.choices?.map((c) => <option key={c} value={c}>{c}</option>)}
        </select>
      ) : field.type === "multiline" ? (
        <textarea style={S.textarea} value={draft} maxLength={field.maxLength}
          onChange={(e) => onChange(e.target.value)} />
      ) : field.type === "boolean" ? (
        <select style={S.input} value={draft} onChange={(e) => onChange(e.target.value)}>
          <option value="">— select —</option>
          <option value="true">Yes</option>
          <option value="false">No</option>
        </select>
      ) : (
        <input
          style={S.input}
          type={field.type === "number" ? "number" : field.type === "date" ? "date" : field.type === "datetime" ? "datetime-local" : "text"}
          value={draft}
          maxLength={field.type === "text" ? field.maxLength : undefined}
          placeholder={forInput(value, field.type) || ""}
          onChange={(e) => onChange(e.target.value)}
        />
      )}
    </label>
  );
}
