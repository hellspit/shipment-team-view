"use client";

// "+ Add attachment" for the request page. Files go straight from the browser to the
// request's SharePoint folder; see app/api/requests/[requestId]/attachments.

import { useRef, useState } from "react";
import { S } from "@/app/ui";

// Keep in sync with lib/attachments.ts.
const ACCEPTED = ".pdf,.xlsx,.docx,.jpg,.jpeg,.png";
const MAX_FILE_BYTES = 25 * 1024 * 1024;

// Graph upload sessions take chunks that are a multiple of 320 KiB; 5 MiB is 16 of them.
const UPLOAD_CHUNK = 5 * 1024 * 1024;

/** PUT one file to a pre-authorized upload URL, in chunks. */
async function uploadToSharePoint(file: File, uploadUrl: string): Promise<void> {
  // A zero-byte file has no range to send; create it with a single empty PUT.
  if (file.size === 0) {
    const res = await fetch(uploadUrl, { method: "PUT", body: new Blob([]) });
    if (!res.ok) throw new Error(`Upload failed for "${file.name}"`);
    return;
  }
  for (let start = 0; start < file.size; start += UPLOAD_CHUNK) {
    const end = Math.min(start + UPLOAD_CHUNK, file.size);
    const res = await fetch(uploadUrl, {
      method: "PUT",
      headers: {
        "Content-Length": String(end - start),
        "Content-Range": `bytes ${start}-${end - 1}/${file.size}`,
      },
      body: file.slice(start, end),
    });
    // 202 = chunk accepted, keep going. 200/201 = final chunk, file created.
    if (res.status === 200 || res.status === 201) return;
    if (!res.ok) throw new Error(`Upload failed for "${file.name}"`);
  }
}

export default function AddAttachment({
  requestId,
  onUploaded,
}: {
  requestId: string;
  onUploaded: () => Promise<void> | void;
}) {
  const input = useRef<HTMLInputElement>(null);
  const [busy, setBusy] = useState(false);
  const [note, setNote] = useState("");
  const [tone, setTone] = useState<"busy" | "ok" | "bad">("busy");

  const sync = async (): Promise<number> => {
    const res = await fetch(`/api/requests/${requestId}/attachments/sync`, { method: "POST" });
    if (!res.ok) throw new Error(await res.text());
    return ((await res.json()) as { count: number }).count;
  };

  const upload = async (list: FileList | null) => {
    const files = Array.from(list ?? []);
    if (!files.length) return;

    const tooBig = files.find((f) => f.size > MAX_FILE_BYTES);
    if (tooBig) {
      setTone("bad");
      setNote(`"${tooBig.name}" is over 25 MB.`);
      return;
    }

    setBusy(true);
    setTone("busy");
    setNote("Preparing upload…");
    try {
      const res = await fetch(`/api/requests/${requestId}/attachments`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ files: files.map((f) => ({ name: f.name, size: f.size })) }),
      });
      if (!res.ok) throw new Error(await res.text());
      const { uploads } = (await res.json()) as { uploads: { name: string; uploadUrl: string }[] };

      for (let i = 0; i < files.length; i++) {
        setNote(`Uploading ${i + 1} of ${files.length}: ${files[i].name}`);
        await uploadToSharePoint(files[i], uploads[i].uploadUrl);
      }

      setNote("Updating the attachment count…");
      const count = await sync();
      await onUploaded();
      setTone("ok");
      setNote(
        `Added ${files.length} file${files.length === 1 ? "" : "s"} · ` +
        `${count} attachment${count === 1 ? "" : "s"} on this request`
      );
    } catch (e) {
      setTone("bad");
      setNote(e instanceof Error ? e.message : "Upload failed");
      // Some files may have landed before the failure — recount so the list stays truthful.
      sync().then(() => onUploaded()).catch(() => {});
    } finally {
      setBusy(false);
    }
  };

  return (
    <>
      <button
        type="button"
        style={{ ...S.ghostBtn, color: "#3a31b0", borderColor: "#c9c5e6", opacity: busy ? 0.6 : 1 }}
        disabled={busy}
        onClick={() => input.current?.click()}
      >
        {busy ? "Uploading…" : "+ Add attachment"}
      </button>
      <input
        ref={input}
        type="file"
        multiple
        accept={ACCEPTED}
        style={{ display: "none" }}
        onChange={(e) => {
          upload(e.target.files);
          e.target.value = ""; // let the same file be picked again
        }}
      />
      {note && (
        <span
          style={{
            fontSize: 13.5,
            fontWeight: 600,
            color: tone === "ok" ? "#16a34a" : tone === "bad" ? "#dc2626" : "#6a6780",
          }}
        >
          {note}
        </span>
      )}
    </>
  );
}
