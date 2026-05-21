import { useState } from "react";
import { useReport } from "../context/ReportContext";
import { useSocket } from "../context/SocketContext";
import { useAIOptions } from "../context/AIOptionsContext";
import { DEFAULT_STYLE_CONFIG } from "../types";
import type { StyleConfig } from "../types";
import ProjectDetailsForm, { DEFAULT_PROJECT_DETAILS } from "./ProjectDetailsForm";
import type { ProjectDetails } from "./ProjectDetailsForm";
import api from "../api/axios";

const AI_OPTIONS = [
  { key: "summary_per_sheet", label: "Summary per sheet", desc: "Group summaries by sheet" },
  { key: "jargons", label: "Explain jargons", desc: "AI explains technical terms" },
  { key: "risks_slide", label: "Detect risks", desc: "Flag risks and RAG issues" },
  { key: "recommend_improvements", label: "Suggest improvements", desc: "AI suggests fixes for risks" },
  { key: "answer_style", label: "Concise answers", desc: "Short bullet-point responses" },
  { key: "insert_summary", label: "Insert summaries into PPT", desc: "Add AI summaries to slides" },
  { key: "validate_ai", label: "Validate AI outputs", desc: "Double-check facts against slides" },
];

type Tab = "actions" | "details" | "style" | "ai";

const Sidebar = () => {
  const { report, clearReport } = useReport();
  const { refineStatus, refineProgress, refineDone, refineError, requestRefineComments,
          restyleStatus, restyleDone, restyleError, requestRestyle,
          slideSummaries, summaryDone, summaryError, requestGenerateSummaries } = useSocket();
  const { options, setOptions } = useAIOptions();
  const [downloading, setDownloading] = useState(false);
  const [tab, setTab] = useState<Tab>("actions");
  const [style, setStyle] = useState<StyleConfig>(() => {
    try { const s = sessionStorage.getItem("ppt_style"); return s ? JSON.parse(s) : { ...DEFAULT_STYLE_CONFIG }; }
    catch { return { ...DEFAULT_STYLE_CONFIG }; }
  });
  const [details, setDetails] = useState<ProjectDetails>(() => {
    try { const s = sessionStorage.getItem("project_details"); return s ? JSON.parse(s) : { ...DEFAULT_PROJECT_DETAILS }; }
    catch { return { ...DEFAULT_PROJECT_DETAILS }; }
  });
  const [regenerating, setRegenerating] = useState(false);
  const [regenSuccess, setRegenSuccess] = useState(false);
  const [regenError, setRegenError] = useState<string | null>(null);

  const pptName = report?.new_fixed_ppt;
  const isRefining = !!refineStatus && !refineDone && !refineError;
  const isRestyling = restyleStatus === "processing";
  const isSummarizing = slideSummaries.length > 0 && !summaryDone && !summaryError;

  const handleDownload = async () => {
    if (!pptName) return;
    setDownloading(true);
    try {
      const res = await api.get(`/download-report/${pptName}`, { responseType: "blob" });
      const url = window.URL.createObjectURL(new Blob([res.data]));
      const a = document.createElement("a");
      a.href = url;
      a.download = report?.friendly_name || pptName;
      document.body.appendChild(a); a.click(); a.remove(); window.URL.revokeObjectURL(url);
    } catch { alert("Download failed"); }
    finally { setDownloading(false); }
  };

  const handleRefine = () => { if (pptName && !isRefining) requestRefineComments(pptName); };
  const handleSummarize = () => { if (pptName && !isSummarizing && !summaryDone) requestGenerateSummaries(pptName); };

  const handleRestyle = () => {
    if (!pptName || isRestyling) return;
    sessionStorage.setItem("ppt_style", JSON.stringify(style));
    requestRestyle(pptName, style);
  };

  const updateDetails = (d: ProjectDetails) => {
    setDetails(d);
    sessionStorage.setItem("project_details", JSON.stringify(d));
    setRegenSuccess(false);
    setRegenError(null);
  };

  const handleRegenerate = async () => {
    if (!pptName || regenerating) return;
    setRegenerating(true);
    setRegenSuccess(false);
    setRegenError(null);

    try {
      const payload = {
        ppt_name: pptName,
        project_title: details.project_title,
        author: details.author,
        report_date: details.report_date,
        date_format: details.date_format,
        roles: details.roles,
      };
      await api.post("/update-project-details", payload);
      setRegenSuccess(true);
      setTimeout(() => setRegenSuccess(false), 4000);
    } catch (e: any) {
      setRegenError(e?.response?.data?.message || e?.message || "Update failed");
    } finally {
      setRegenerating(false);
    }
  };

  const toggle = (key: string) => setOptions({ ...options, [key]: !options[key] });
  const refinePercent = refineProgress ? Math.round((refineProgress.processed / refineProgress.total) * 100) : 0;

  const updateStyle = (key: keyof StyleConfig, value: string | number | boolean) => {
    setStyle(prev => ({ ...prev, [key]: value }));
  };

  return (
    <div style={{ width: 280, minWidth: 280, height: "100%", borderRight: "1px solid var(--border)", background: "var(--bg-secondary)", display: "flex", flexDirection: "column" }}>
      {/* Brand */}
      <div style={{ padding: "14px 16px", borderBottom: "1px solid var(--border)", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
        <span style={{ fontSize: 15, fontWeight: 800, letterSpacing: "-0.04em", color: "var(--accent)" }}>ReportAI</span>
        <button onClick={clearReport} title="New report" style={{ width: 26, height: 26, borderRadius: 7, border: "1px solid var(--border)", background: "transparent", cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", color: "var(--text-muted)", fontSize: 13 }}>+</button>
      </div>

      {/* Tabs */}
      <div style={{ display: "flex", borderBottom: "1px solid var(--border)" }}>
        {(["actions", "details", "style", "ai"] as Tab[]).map(t => (
          <button key={t} onClick={() => setTab(t)} style={{
            flex: 1, padding: "9px 0", fontSize: 10, fontWeight: 600, fontFamily: "var(--font-sans)",
            border: "none", cursor: "pointer", textTransform: "capitalize", transition: "all 0.15s",
            background: tab === t ? "#fff" : "transparent",
            color: tab === t ? "var(--accent)" : "var(--text-muted)",
            borderBottom: tab === t ? "2px solid var(--accent)" : "2px solid transparent",
            boxShadow: tab === t ? "var(--shadow-sm)" : "none",
          }}>{t}</button>
        ))}
      </div>

      {/* Content */}
      <div style={{ flex: 1, overflowY: "auto", padding: 12 }}>

        {/* === ACTIONS TAB === */}
        {tab === "actions" && (
          <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
            <ActionBtn label={downloading ? "Downloading..." : "Download PPT"} color="var(--cyan)" disabled={!pptName || downloading} onClick={handleDownload} icon="↓" />
            <ActionBtn label={isRefining ? "Refining..." : "Refine Comments"} color="var(--accent)" disabled={!pptName || isRefining} onClick={handleRefine} icon="✎" />
            <ActionBtn label={isSummarizing ? "Summarizing..." : summaryDone ? "Summaries Done ✓" : "Generate Summaries"} color="var(--cyan)" disabled={!pptName || isSummarizing || summaryDone} onClick={handleSummarize} icon="≡" />

            {isRefining && refineProgress && (
              <div className="animate-fade" style={{ padding: "8px 10px", borderRadius: 8, background: "var(--accent-dim)" }}>
                <div style={{ display: "flex", justifyContent: "space-between", fontSize: 10, color: "var(--accent)", fontWeight: 600, marginBottom: 5 }}>
                  <span>{refineProgress.batch_info || "Processing..."}</span>
                  <span>{refinePercent}%</span>
                </div>
                <div style={{ width: "100%", height: 3, borderRadius: 99, background: "rgba(167,139,250,0.15)" }}>
                  <div style={{ width: `${refinePercent}%`, height: "100%", borderRadius: 99, background: "var(--accent)", transition: "width 0.3s" }} />
                </div>
              </div>
            )}
            {refineDone && <StatusBadge color="var(--success)" bg="var(--success-dim)" text={`✓ Refined ${refineDone.refined_count}/${refineDone.total_found}`} />}
            {refineError && <StatusBadge color="var(--error)" bg="var(--error-dim)" text={refineError} />}
            {restyleDone && <StatusBadge color="var(--success)" bg="var(--success-dim)" text="✓ PPT restyled" />}
            {restyleError && <StatusBadge color="var(--error)" bg="var(--error-dim)" text={restyleError} />}
          </div>
        )}

        {/* === DETAILS TAB === */}
        {tab === "details" && (
          <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
            <div style={{ fontSize: 10, color: "var(--text-muted)", lineHeight: 1.5, marginBottom: 2 }}>
              Edit project details and apply changes to the generated PPT.
            </div>

            <ProjectDetailsForm details={details} onChange={updateDetails} compact />

            {/* Regenerate / Apply button */}
            <button
              onClick={handleRegenerate}
              disabled={!pptName || regenerating}
              style={{
                width: "100%", padding: "10px 14px", borderRadius: 10, marginTop: 4,
                border: "none",
                background: (pptName && !regenerating) ? "var(--accent)" : "var(--bg-hover)",
                color: (pptName && !regenerating) ? "#fff" : "var(--text-muted)",
                fontSize: 12, fontWeight: 700, cursor: (pptName && !regenerating) ? "pointer" : "default",
                fontFamily: "var(--font-sans)", transition: "all 0.2s",
              }}
            >
              {regenerating ? "Applying Changes..." : "Apply to PPT"}
            </button>

            {regenSuccess && <StatusBadge color="var(--success)" bg="var(--success-dim)" text="✓ Project details updated in PPT" />}
            {regenError && <StatusBadge color="var(--error)" bg="var(--error-dim)" text={regenError} />}
          </div>
        )}

        {/* === STYLE TAB === */}
        {tab === "style" && (
          <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
            <SectionLabel text="Header Row" />
            <ColorField label="Background" value={style.hdr_bg_color} onChange={v => updateStyle("hdr_bg_color", v)} />
            <ColorField label="Text Color" value={style.hdr_text_color} onChange={v => updateStyle("hdr_text_color", v)} />
            <SelectField label="Font" value={style.hdr_font} options={["Calibri","Tahoma","Arial","Verdana","Georgia"]} onChange={v => updateStyle("hdr_font", v)} />
            <NumField label="Font Size" value={style.hdr_size} min={8} max={18} onChange={v => updateStyle("hdr_size", v)} />

            <SectionLabel text="Data Cells" />
            <SelectField label="Font" value={style.data_font} options={["Calibri","Tahoma","Arial","Verdana","Georgia"]} onChange={v => updateStyle("data_font", v)} />
            <NumField label="Font Size" value={style.data_size} min={8} max={16} onChange={v => updateStyle("data_size", v)} />

            <SectionLabel text="Section Titles" />
            <ColorField label="Color" value={style.title_color} onChange={v => updateStyle("title_color", v)} />
            <SelectField label="Font" value={style.title_font} options={["Tahoma","Calibri","Arial","Verdana","Georgia"]} onChange={v => updateStyle("title_font", v)} />
            <NumField label="Font Size" value={style.title_size} min={14} max={28} onChange={v => updateStyle("title_size", v)} />

            <button onClick={handleRestyle} disabled={!pptName || isRestyling} style={{
              width: "100%", padding: "10px 14px", borderRadius: 10, marginTop: 8,
              border: "none", background: (pptName && !isRestyling) ? "var(--accent)" : "var(--bg-hover)",
              color: (pptName && !isRestyling) ? "#fff" : "var(--text-muted)",
              fontSize: 12, fontWeight: 700, cursor: (pptName && !isRestyling) ? "pointer" : "default",
              fontFamily: "var(--font-sans)", transition: "all 0.2s",
            }}>
              {isRestyling ? "Applying..." : "Apply Style"}
            </button>

            <button onClick={() => { setStyle({ ...DEFAULT_STYLE_CONFIG }); }} style={{
              width: "100%", padding: "8px 14px", borderRadius: 10,
              border: "1px solid var(--border)", background: "transparent",
              color: "var(--text-muted)", fontSize: 11, fontWeight: 500, cursor: "pointer",
              fontFamily: "var(--font-sans)",
            }}>Reset to Defaults</button>
          </div>
        )}

        {/* === AI TAB === */}
        {tab === "ai" && (
          <div style={{ display: "flex", flexDirection: "column", gap: 2 }}>
            <div style={{ fontSize: 10, color: "var(--text-muted)", marginBottom: 8, lineHeight: 1.5 }}>
              These settings affect how AI responds in chat. Toggle them before asking questions.
            </div>
            {AI_OPTIONS.map(opt => {
              const on = !!options[opt.key];
              return (
                <div key={opt.key} onClick={() => toggle(opt.key)} style={{
                  display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8,
                  padding: "8px 8px", borderRadius: 8, cursor: "pointer", transition: "background 0.15s",
                }}
                onMouseEnter={e => { (e.currentTarget as HTMLElement).style.background = "var(--bg-hover)"; }}
                onMouseLeave={e => { (e.currentTarget as HTMLElement).style.background = "transparent"; }}
                >
                  <div>
                    <div style={{ fontSize: 12, color: on ? "var(--text-primary)" : "var(--text-muted)", fontWeight: 500, transition: "color 0.15s" }}>{opt.label}</div>
                    <div style={{ fontSize: 10, color: "var(--text-muted)", marginTop: 1 }}>{opt.desc}</div>
                  </div>
                  <Toggle on={on} />
                </div>
              );
            })}
          </div>
        )}
      </div>

      {/* Status */}
      <div style={{ padding: 12 }}>
        <div style={{ padding: 10, borderRadius: 8, background: "var(--bg-tertiary)", border: "1px solid var(--border)" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 5, marginBottom: 3 }}>
            <span style={{ width: 5, height: 5, borderRadius: 99, background: "var(--success)" }} className="animate-pulse-soft" />
            <span style={{ fontSize: 10, fontWeight: 500, color: "var(--text-muted)" }}>Connected</span>
          </div>
          <span style={{ fontSize: 9, color: "var(--text-muted)", fontFamily: "var(--font-mono)" }}>
            {pptName ? pptName.slice(0, 24) + "..." : "No report loaded"}
          </span>
        </div>
      </div>
    </div>
  );
};

/* ── Sub-components ── */

const Toggle = ({ on }: { on: boolean }) => (
  <div style={{
    width: 30, height: 16, borderRadius: 99, position: "relative", flexShrink: 0,
    background: on ? "var(--accent)" : "var(--bg-hover)", transition: "background 0.2s",
    border: on ? "none" : "1px solid var(--border)",
  }}>
    <div style={{
      width: 12, height: 12, borderRadius: 99, background: "#fff",
      position: "absolute", top: 2, left: on ? 16 : 2, transition: "left 0.2s",
      boxShadow: "0 1px 2px rgba(0,0,0,0.3)",
    }} />
  </div>
);

const ActionBtn = ({ label, color, disabled, onClick, icon }: { label: string; color: string; disabled: boolean; onClick: () => void; icon: string }) => (
  <button onClick={onClick} disabled={disabled} style={{
    width: "100%", padding: "10px 14px", borderRadius: 10,
    border: "1px solid var(--border)", background: "#fff",
    cursor: disabled ? "default" : "pointer", display: "flex", alignItems: "center", gap: 10,
    opacity: disabled ? 0.4 : 1, transition: "all 0.2s",
    color: "var(--text-primary)", fontSize: 12, fontWeight: 600, fontFamily: "var(--font-sans)",
    boxShadow: disabled ? "none" : "var(--shadow-sm)",
  }}
  onMouseEnter={e => { if(!disabled) { e.currentTarget.style.borderColor = color; e.currentTarget.style.transform = "translateY(-1px)"; e.currentTarget.style.boxShadow = "var(--shadow-md)"; } }}
  onMouseLeave={e => { if(!disabled) { e.currentTarget.style.borderColor = "var(--border)"; e.currentTarget.style.transform = "translateY(0)"; e.currentTarget.style.boxShadow = "var(--shadow-sm)"; } }}
  >
    <span style={{ 
      width: 24, height: 24, borderRadius: 6, background: `${color}12`, 
      display: "flex", alignItems: "center", justifyContent: "center",
      color, fontSize: 14 
    }}>{icon}</span>
    {label}
  </button>
);

const StatusBadge = ({ color, bg, text }: { color: string; bg: string; text: string }) => (
  <div className="animate-slide-up" style={{ padding: "7px 10px", borderRadius: 7, background: bg, fontSize: 11, color, fontWeight: 500 }}>{text}</div>
);

const SectionLabel = ({ text }: { text: string }) => (
  <div style={{ fontSize: 10, fontWeight: 700, color: "var(--text-muted)", letterSpacing: "0.06em", textTransform: "uppercase", fontFamily: "var(--font-mono)", marginTop: 4 }}>{text}</div>
);

const ColorField = ({ label, value, onChange }: { label: string; value: string; onChange: (v: string) => void }) => (
  <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8 }}>
    <span style={{ fontSize: 11, color: "var(--text-secondary)" }}>{label}</span>
    <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
      <input type="color" value={`#${value}`} onChange={e => onChange(e.target.value.replace("#", ""))}
        style={{ width: 24, height: 24, border: "1px solid var(--border)", borderRadius: 6, cursor: "pointer", background: "none", padding: 0 }} />
      <input type="text" value={value} onChange={e => onChange(e.target.value.replace("#", ""))} maxLength={6}
        style={{ width: 64, padding: "4px 6px", borderRadius: 6, border: "1px solid var(--border)", background: "var(--bg-tertiary)",
                 color: "var(--text-primary)", fontSize: 10, fontFamily: "var(--font-mono)", outline: "none" }} />
    </div>
  </div>
);

const SelectField = ({ label, value, options, onChange }: { label: string; value: string; options: string[]; onChange: (v: string) => void }) => (
  <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8 }}>
    <span style={{ fontSize: 11, color: "var(--text-secondary)" }}>{label}</span>
    <select value={value} onChange={e => onChange(e.target.value)} style={{
      padding: "4px 6px", borderRadius: 6, border: "1px solid var(--border)", background: "var(--bg-tertiary)",
      color: "var(--text-primary)", fontSize: 11, outline: "none", cursor: "pointer",
    }}>
      {options.map(o => <option key={o} value={o}>{o}</option>)}
    </select>
  </div>
);

const NumField = ({ label, value, min, max, onChange }: { label: string; value: number; min: number; max: number; onChange: (v: number) => void }) => (
  <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8 }}>
    <span style={{ fontSize: 11, color: "var(--text-secondary)" }}>{label}</span>
    <input type="number" value={value} min={min} max={max} onChange={e => onChange(Number(e.target.value))} style={{
      width: 50, padding: "4px 6px", borderRadius: 6, border: "1px solid var(--border)", background: "var(--bg-tertiary)",
      color: "var(--text-primary)", fontSize: 11, fontFamily: "var(--font-mono)", outline: "none", textAlign: "center",
    }} />
  </div>
);

export default Sidebar;
