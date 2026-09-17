"use client";

import { useState, useCallback } from "react";
import * as XLSX from "xlsx";

// ── Types ──

interface Recipient {
  name: string;
  type: "EP" | "BD";
  sheet: string;
  email: string;
}

interface EmailPayload {
  to: string[];
  cc: string[];
  subject: string;
  body_html: string;
  name: string;
  type: string;
  sheet: string;
  row_count: number;
}

interface DraftInfo {
  message_id: string;
  name: string;
  status: "draft" | "sent" | "failed";
}

// ── Constants ──

const API_BASE =
  process.env.NEXT_PUBLIC_API_URL ||
  "https://revenue-warriors-ep-bd-report-production.up.railway.app";

const RECIPIENTS: Recipient[] = [
  { name: "Aditya Achalkar", type: "EP", sheet: "EP - Aditya Achalkar", email: "aditya.achalkar@practus.com" },
  { name: "Ameya Waingankar", type: "EP", sheet: "EP - Ameya Waingankar", email: "ameya.waingankar@practus.com" },
  { name: "Foram Maru", type: "EP", sheet: "EP - Foram Maru", email: "foram.maru@practus.com" },
  { name: "Shashank Silhare", type: "EP", sheet: "EP - Shashank Silhare", email: "shashank.silhare@practus.com" },
  { name: "Yashwin Pamecha", type: "EP", sheet: "EP - Yashwin Pamecha", email: "yashwin.pamecha@practus.com" },
  { name: "Varun Shankarnarayan", type: "EP", sheet: "EP - Varun Shankarnarayan", email: "varun.shankarnarayan@practus.com" },
  { name: "Sharan Prakash", type: "EP", sheet: "EP - Sharan Prakash", email: "sharan.prakash@practus.com" },
  { name: "Priyanka Baram", type: "EP", sheet: "EP - Priyanka Baram", email: "priyanka.baram@practus.com" },
  { name: "Rajaram Ganesan", type: "EP", sheet: "EP - Rajaram Ganesan", email: "rajaram.ganesan@practus.com" },
  { name: "Ravikanth Rao", type: "EP", sheet: "EP - Ravikanth Rao", email: "ravikanth.rao@practus.com" },
  { name: "Bhavik Desai", type: "BD", sheet: "BD - Bhavik Desai", email: "bhavik.desai@practus.com" },
  { name: "Deepak Narayanan", type: "BD", sheet: "BD - Deepak Narayanan", email: "deepak.narayanan@practus.com" },
  { name: "Priyanka Baram", type: "BD", sheet: "BD - Priyanka Baram", email: "priyanka.baram@practus.com" },
  { name: "Rajaram Ganesan", type: "BD", sheet: "BD - Rajaram Ganesan", email: "rajaram.ganesan@practus.com" },
  { name: "Ravikanth Rao", type: "BD", sheet: "BD - Ravikanth Rao", email: "ravikanth.rao@practus.com" },
  { name: "Sreejit Nair", type: "BD", sheet: "BD - Sreejit Nair", email: "sreejit.nair@practus.com" },
  { name: "Srinivasan Venkataraman", type: "BD", sheet: "BD - Srinivasan Venkataraman", email: "srinivasan.venkataraman@practus.com" },
];

// ── API helpers ──

async function apiPost(path: string, body: object) {
  const r = await fetch(`${API_BASE}${path}`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
  if (!r.ok) throw new Error(`API error ${r.status}`);
  return r.json();
}

async function apiUpload(path: string, file: File, params: Record<string, string> = {}) {
  const form = new FormData();
  form.append("file", file);
  const qs = new URLSearchParams(params).toString();
  const r = await fetch(`${API_BASE}${path}${qs ? "?" + qs : ""}`, {
    method: "POST",
    body: form,
  });
  if (!r.ok) throw new Error(`API error ${r.status}`);
  return r.json();
}

// ── Component ──

export default function RevenueWarriorsPage() {
  const [step, setStep] = useState(1);
  const [file, setFile] = useState<File | null>(null);
  const [week, setWeek] = useState("");
  const [cc, setCc] = useState("venkat@practus.com");
  const [emails, setEmails] = useState<EmailPayload[]>([]);
  const [drafts, setDrafts] = useState<DraftInfo[]>([]);
  const [progress, setProgress] = useState({ current: 0, total: 0, label: "" });
  const [expandedEmail, setExpandedEmail] = useState<number | null>(null);
  const [confirmed, setConfirmed] = useState(false);
  const [done, setDone] = useState(false);

  // ── Step 1: Upload ──

  const handleFile = useCallback(async (f: File) => {
    setFile(f);

    // Option A: Parse server-side via /parse endpoint
    try {
      const result = await apiUpload("/rahul/revenue-warriors/parse", f, { cc, week });
      setEmails(result.emails);
      setWeek(result.week);
      setStep(2);
      return;
    } catch {
      // Fallback: parse client-side with SheetJS
    }

    // Option B: Client-side parsing
    const data = await f.arrayBuffer();
    const wb = XLSX.read(new Uint8Array(data), { type: "array" });
    const weeklySheet = wb.SheetNames.find((s) => s.startsWith("Revenue_FY27"));
    const detectedWeek = weeklySheet ? "Week of " + weeklySheet.replace("Revenue_FY27 ", "") : "Current Week";
    if (!week) setWeek(detectedWeek);

    const built: EmailPayload[] = [];
    for (const r of RECIPIENTS) {
      const ws = wb.Sheets[r.sheet];
      if (!ws) continue;
      const json: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
      if (json.length < 3) continue;
      built.push({
        to: [r.email],
        cc: cc ? cc.split(",").map((e) => e.trim()).filter(Boolean) : [],
        subject: `Revenue Warriors — Your ${r.type === "EP" ? "Engagement Partner" : "BD Owner"} Report (${detectedWeek})`,
        body_html: `<p>Parsed ${json.length} rows from ${r.sheet}</p>`,
        name: r.name,
        type: r.type,
        sheet: r.sheet,
        row_count: json.length - 3,
      });
    }
    setEmails(built);
    setStep(2);
  }, [cc, week]);

  const onDrop = useCallback(
    (e: React.DragEvent) => {
      e.preventDefault();
      const f = e.dataTransfer.files[0];
      if (f) handleFile(f);
    },
    [handleFile]
  );

  // ── Step 3: Create Drafts ──

  const createDrafts = async () => {
    setStep(3);
    const newDrafts: DraftInfo[] = [];
    setProgress({ current: 0, total: emails.length, label: "Starting..." });

    for (let i = 0; i < emails.length; i++) {
      const em = emails[i];
      setProgress({ current: i + 1, total: emails.length, label: em.name });
      try {
        const result = await apiPost("/rahul/revenue-warriors/draft", {
          to: em.to,
          cc: em.cc,
          subject: em.subject,
          body_html: em.body_html,
        });
        newDrafts.push({ message_id: result.message_id, name: em.name, status: "draft" });
      } catch {
        newDrafts.push({ message_id: "", name: em.name, status: "failed" });
      }
      setDrafts([...newDrafts]);
    }
    setProgress({ current: emails.length, total: emails.length, label: "Done" });
  };

  // ── Step 4: Send All ──

  const sendAll = async () => {
    setDone(false);
    const ids = drafts.filter((d) => d.status === "draft").map((d) => d.message_id);
    try {
      const result = await apiPost("/rahul/revenue-warriors/send-all", { message_ids: ids });
      setDrafts((prev) =>
        prev.map((d) => {
          const r = result.results.find((x: any) => x.message_id === d.message_id);
          return r ? { ...d, status: r.status } : d;
        })
      );
    } catch {
      // Fallback: one by one
      for (const draft of drafts) {
        if (draft.status !== "draft") continue;
        try {
          await apiPost("/rahul/revenue-warriors/send", { message_id: draft.message_id });
          setDrafts((prev) => prev.map((d) => (d.message_id === draft.message_id ? { ...d, status: "sent" } : d)));
        } catch {
          setDrafts((prev) => prev.map((d) => (d.message_id === draft.message_id ? { ...d, status: "failed" } : d)));
        }
      }
    }
    setDone(true);
  };

  // ── Render ──

  const epCount = emails.filter((e) => e.type === "EP").length;
  const bdCount = emails.filter((e) => e.type === "BD").length;
  const draftOk = drafts.filter((d) => d.status === "draft" || d.status === "sent").length;
  const sentOk = drafts.filter((d) => d.status === "sent").length;

  return (
    <div className="max-w-4xl mx-auto p-6">
      {/* Header */}
      <div className="flex items-center gap-4 mb-6">
        <span className="text-3xl">⚔️</span>
        <div>
          <h1 className="text-xl font-bold text-gray-900">Revenue Warriors — Dispatch</h1>
          <p className="text-sm text-gray-500">Upload → Preview → Draft → Send</p>
        </div>
      </div>

      {/* Steps bar */}
      <div className="flex gap-2 mb-8">
        {[1, 2, 3, 4].map((s) => (
          <div
            key={s}
            className={`flex items-center gap-2 px-4 py-2 rounded-lg text-sm font-medium transition-colors ${
              s === step ? "bg-blue-100 text-blue-700" : s < step ? "bg-green-100 text-green-700" : "bg-gray-100 text-gray-400"
            }`}
          >
            <span className={`w-5 h-5 rounded-full text-xs flex items-center justify-center font-bold ${
              s === step ? "bg-blue-600 text-white" : s < step ? "bg-green-600 text-white" : "bg-gray-300 text-gray-500"
            }`}>
              {s < step ? "✓" : s}
            </span>
            {["Upload", "Preview", "Drafts", "Send"][s - 1]}
          </div>
        ))}
      </div>

      {/* Step 1: Upload */}
      {step === 1 && (
        <div className="space-y-4">
          <div
            className="border-2 border-dashed border-gray-300 rounded-xl p-16 text-center cursor-pointer hover:border-blue-400 hover:bg-blue-50 transition-colors"
            onDrop={onDrop}
            onDragOver={(e) => e.preventDefault()}
            onClick={() => document.getElementById("fileInput")?.click()}
          >
            <h3 className="text-lg font-semibold text-gray-700">Drop Revenue Warriors workbook here</h3>
            <p className="text-sm text-gray-500 mt-1">or click to browse — reads locally in your browser</p>
            <input id="fileInput" type="file" accept=".xlsx" className="hidden" onChange={(e) => e.target.files?.[0] && handleFile(e.target.files[0])} />
          </div>
          {file && <p className="text-sm text-gray-600">📊 {file.name} ({(file.size / 1024).toFixed(0)} KB)</p>}
          <div className="grid grid-cols-2 gap-4">
            <div>
              <label className="text-xs font-semibold text-gray-500 block mb-1">Week reference</label>
              <input className="w-full px-3 py-2 border rounded-lg text-sm" value={week} onChange={(e) => setWeek(e.target.value)} placeholder="Auto-detected from file" />
            </div>
            <div>
              <label className="text-xs font-semibold text-gray-500 block mb-1">CC emails</label>
              <input className="w-full px-3 py-2 border rounded-lg text-sm" value={cc} onChange={(e) => setCc(e.target.value)} />
            </div>
          </div>
        </div>
      )}

      {/* Step 2: Preview */}
      {step === 2 && (
        <div className="space-y-4">
          <div className="flex gap-6 mb-4">
            <div className="text-center"><div className="text-3xl font-bold text-blue-600">{epCount}</div><div className="text-xs text-gray-500">EP emails</div></div>
            <div className="text-center"><div className="text-3xl font-bold text-pink-600">{bdCount}</div><div className="text-xs text-gray-500">BD emails</div></div>
            <div className="text-center"><div className="text-3xl font-bold text-gray-800">{emails.length}</div><div className="text-xs text-gray-500">Total</div></div>
          </div>

          <div className="bg-blue-50 border border-blue-200 text-blue-800 text-sm rounded-lg p-3">
            Each email contains the data table from that person's sheet. Expand any to verify.
          </div>

          {emails.map((em, i) => (
            <div key={i} className="border rounded-lg overflow-hidden">
              <div className="flex justify-between items-center px-4 py-3 bg-gray-50 cursor-pointer" onClick={() => setExpandedEmail(expandedEmail === i ? null : i)}>
                <div className="flex items-center gap-2">
                  <span className={`text-xs font-bold px-2 py-0.5 rounded-full ${em.type === "EP" ? "bg-blue-100 text-blue-700" : "bg-pink-100 text-pink-700"}`}>{em.type}</span>
                  <span className="font-medium text-sm">{em.name}</span>
                  <span className="text-xs text-gray-400">{em.to[0]}</span>
                  <span className="text-xs text-gray-400">({em.row_count} rows)</span>
                </div>
                <span className="text-xs text-blue-600">{expandedEmail === i ? "▲" : "▼"}</span>
              </div>
              {expandedEmail === i && (
                <div className="p-4 text-sm border-t max-h-48 overflow-y-auto" dangerouslySetInnerHTML={{ __html: em.body_html }} />
              )}
            </div>
          ))}

          <div className="flex gap-3 mt-6">
            <button className="px-4 py-2 border rounded-lg text-sm" onClick={() => setStep(1)}>← Back</button>
            <button className="px-6 py-2 bg-blue-600 text-white rounded-lg text-sm font-semibold hover:bg-blue-700" onClick={createDrafts}>
              Create Drafts in Outlook →
            </button>
          </div>
        </div>
      )}

      {/* Step 3: Drafts */}
      {step === 3 && (
        <div className="space-y-4">
          <div className="bg-white border rounded-lg p-6">
            <h2 className="font-semibold mb-3">Creating drafts…</h2>
            <div className="h-2 bg-gray-200 rounded-full overflow-hidden mb-2">
              <div className="h-full bg-blue-600 rounded-full transition-all" style={{ width: `${(progress.current / Math.max(progress.total, 1)) * 100}%` }} />
            </div>
            <p className="text-sm text-gray-500">{progress.current}/{progress.total} — {progress.label}</p>
            <div className="mt-4 space-y-1 max-h-60 overflow-y-auto font-mono text-xs">
              {drafts.map((d, i) => (
                <div key={i} className={d.status === "failed" ? "text-red-600" : "text-green-600"}>
                  {d.status === "failed" ? "✗" : "✓"} {d.name}
                </div>
              ))}
            </div>
          </div>
          {progress.current === progress.total && progress.total > 0 && (
            <div className="flex gap-3">
              <button className="px-4 py-2 border rounded-lg text-sm" onClick={() => setStep(2)}>← Back</button>
              <button className="px-6 py-2 bg-green-600 text-white rounded-lg text-sm font-semibold hover:bg-green-700" onClick={() => setStep(4)}>
                Review done — Send All →
              </button>
            </div>
          )}
        </div>
      )}

      {/* Step 4: Send */}
      {step === 4 && (
        <div className="space-y-4">
          <div className="bg-amber-50 border border-amber-200 text-amber-800 text-sm rounded-lg p-3">
            ⚠️ This will send <strong>{draftOk}</strong> emails from Isabelle's account. Review the drafts in Outlook first.
          </div>
          {!done && (
            <>
              <label className="flex items-center gap-3 cursor-pointer text-sm">
                <input type="checkbox" className="w-4 h-4 accent-blue-600" checked={confirmed} onChange={(e) => setConfirmed(e.target.checked)} />
                I've reviewed the drafts and Venkat has approved
              </label>
              <div className="flex gap-3">
                <button className="px-4 py-2 border rounded-lg text-sm" onClick={() => setStep(3)}>← Back</button>
                <button className="px-6 py-2 bg-green-600 text-white rounded-lg text-sm font-semibold hover:bg-green-700 disabled:opacity-40" disabled={!confirmed} onClick={sendAll}>
                  🚀 Send All Emails
                </button>
              </div>
            </>
          )}
          {done && (
            <div className="bg-green-50 border border-green-200 text-green-800 text-sm rounded-lg p-4">
              ✅ {sentOk}/{drafts.length} emails sent from Isabelle's account.
            </div>
          )}
        </div>
      )}
    </div>
  );
}
