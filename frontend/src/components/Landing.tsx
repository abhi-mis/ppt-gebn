import { useState } from "react";
import MultiFileUpload from "./MultiFileUpload";
import ProjectDetailsForm, { DEFAULT_PROJECT_DETAILS } from "./ProjectDetailsForm";
import type { ProjectDetails } from "./ProjectDetailsForm";
import api from "../api/axios";
import { useReport } from "../context/ReportContext";

type Step = "idle" | "uploading" | "extracting" | "generating" | "fixing" | "done" | "error";

const STEPS: { key: Step; label: string }[] = [
  { key: "uploading", label: "Uploading files" },
  { key: "extracting", label: "Extracting Excel data" },
  { key: "generating", label: "Building PPT slides" },
  { key: "fixing", label: "Fixing layout & fonts" },
  { key: "done", label: "Complete" },
];

const Landing = () => {
  const [excelFiles, setExcelFiles] = useState<File[]>([]);
  const [step, setStep] = useState<Step>("idle");
  const [error, setError] = useState<string | null>(null);
  const [autoSummarize, setAutoSummarize] = useState(false);
  const [details, setDetails] = useState<ProjectDetails>(() => {
    try { const s = sessionStorage.getItem("project_details"); return s ? JSON.parse(s) : { ...DEFAULT_PROJECT_DETAILS }; }
    catch { return { ...DEFAULT_PROJECT_DETAILS }; }
  });
  const [showDetails, setShowDetails] = useState(false);
  const { setReport } = useReport();

  const updateDetails = (d: ProjectDetails) => {
    setDetails(d);
    sessionStorage.setItem("project_details", JSON.stringify(d));
  };

  const submit = async () => {
    if (excelFiles.length === 0) return;
    setError(null);

    const formData = new FormData();
    for (const file of excelFiles) {
      formData.append("excel_files", file);
    }
    formData.append("auto_summarize", autoSummarize ? "true" : "false");

    // Project details
    if (details.project_title) formData.append("project_title", details.project_title);
    if (details.author) formData.append("author", details.author);
    if (details.report_date) formData.append("report_date", details.report_date);
    formData.append("date_format", details.date_format);
    if (Object.keys(details.roles).length > 0) {
      formData.append("roles", JSON.stringify(details.roles));
    }

    try {
      setStep("uploading");
      const stepTimer = setInterval(() => {
        setStep(prev => {
          if (prev === "uploading") return "extracting";
          if (prev === "extracting") return "generating";
          if (prev === "generating") return "fixing";
          return prev;
        });
      }, 2200);

      const res = await api.post("/upload-report", formData, {
        headers: { "Content-Type": "multipart/form-data" },
      });

      clearInterval(stepTimer);
      setStep("done");

      setTimeout(() => setReport(res.data), 600);
    } catch (e: any) {
      setStep("error");
      setError(e?.response?.data?.message || e?.message || "Upload failed");
    }
  };

  const isProcessing = step !== "idle" && step !== "error";
  const activeIdx = STEPS.findIndex(s => s.key === step);
  const canSubmit = excelFiles.length > 0 && !isProcessing;

  const filledFieldCount = [
    details.project_title,
    details.author,
    Object.keys(details.roles).length > 0,
  ].filter(Boolean).length;

  return (
    <div style={{
      minHeight: "100vh", display: "flex", flexDirection: "column",
      alignItems: "center", justifyContent: "flex-start",
      padding: "48px 24px", position: "relative",
    }}>
      {/* Subtle gradient orbs */}
      <div style={{ position: "absolute", top: "-20%", left: "-10%", width: 500, height: 500, borderRadius: "50%", background: "radial-gradient(circle, rgba(167,139,250,0.06) 0%, transparent 70%)", pointerEvents: "none" }} />
      <div style={{ position: "absolute", bottom: "-20%", right: "-10%", width: 600, height: 600, borderRadius: "50%", background: "radial-gradient(circle, rgba(34,211,238,0.04) 0%, transparent 70%)", pointerEvents: "none" }} />

      <div style={{ position: "relative", zIndex: 1, width: "100%", maxWidth: 580, display: "flex", flexDirection: "column", alignItems: "center", gap: 24 }}>

        {/* Header */}
        <div style={{ textAlign: "center" }}>
          <div style={{
            display: "inline-flex", alignItems: "center", gap: 8,
            padding: "6px 14px", borderRadius: 99,
            background: "#fff", border: "1px solid var(--border)",
            boxShadow: "var(--shadow-sm)",
            marginBottom: 20,
          }}>
            <span style={{ width: 6, height: 6, borderRadius: 99, background: "var(--success)" }} className="animate-pulse-soft" />
            <span style={{ fontSize: 10, fontWeight: 600, color: "var(--text-muted)", letterSpacing: "0.1em", textTransform: "uppercase", fontFamily: "var(--font-mono)" }}>
              Engine Ready
            </span>
          </div>

          <h1 style={{
            fontSize: "clamp(1.8rem, 4vw, 2.8rem)", fontWeight: 800,
            letterSpacing: "-0.04em", lineHeight: 1.15, marginBottom: 10,
          }}>
            Generate your<br />
            <span style={{ color: "var(--accent)" }}>project report</span>
          </h1>
          <p style={{ fontSize: 15, color: "var(--text-secondary)", lineHeight: 1.6, maxWidth: 400, margin: "0 auto" }}>
            Upload your Excel data files to create a polished status report.
          </p>
        </div>

        {/* Upload Card */}
        <div style={{
          width: "100%", padding: 24, borderRadius: 20,
          background: "var(--bg-secondary)", border: "1px solid var(--border)",
          boxShadow: "var(--shadow-md)",
        }}>
          {/* File Upload */}
          <div style={{ marginBottom: 16 }}>
            <MultiFileUpload
              label="Excel Data"
              accept=".xls,.xlsx"
              files={excelFiles}
              onFilesChange={setExcelFiles}
            />
          </div>

          {/* Project Details Collapsible */}
          <div style={{
            marginBottom: 16, borderRadius: 10,
            border: "1px solid var(--border)", overflow: "hidden",
            transition: "all 0.2s",
          }}>
            {/* Toggle header */}
            <button
              onClick={() => setShowDetails(!showDetails)}
              style={{
                width: "100%", display: "flex", alignItems: "center", justifyContent: "space-between",
                padding: "10px 12px", border: "none", cursor: "pointer",
                background: showDetails ? "rgba(167,139,250,0.04)" : "var(--bg-tertiary)",
                transition: "all 0.15s",
              }}
            >
              <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                <span style={{ fontSize: 13, color: "var(--accent)" }}>⚙</span>
                <span style={{ fontSize: 12, fontWeight: 600, color: "var(--text-primary)" }}>
                  Project Details
                </span>
                {filledFieldCount > 0 && !showDetails && (
                  <span style={{
                    padding: "1px 6px", borderRadius: 99, fontSize: 9, fontWeight: 700,
                    background: "var(--accent)", color: "#fff",
                  }}>
                    {filledFieldCount} filled
                  </span>
                )}
              </div>
              <span style={{
                fontSize: 10, color: "var(--text-muted)",
                transform: showDetails ? "rotate(180deg)" : "rotate(0deg)",
                transition: "transform 0.2s",
              }}>
                ▼
              </span>
            </button>

            {/* Collapsible form */}
            {showDetails && (
              <div style={{ padding: "12px 12px 14px", borderTop: "1px solid var(--border)" }}>
                <ProjectDetailsForm details={details} onChange={updateDetails} />
              </div>
            )}
          </div>

          {/* Auto-summarize toggle */}
          <div
            onClick={() => setAutoSummarize(prev => !prev)}
            style={{
              display: "flex", alignItems: "center", justifyContent: "space-between",
              padding: 14,
              borderRadius: 12,
              marginBottom: 16,
              background: autoSummarize ? "var(--accent-dim)" : "#fff",
              border: `1px solid ${autoSummarize ? "var(--accent-border)" : "var(--border)"}`,
              cursor: "pointer", 
              transition: "all 0.2s cubic-bezier(0.4, 0, 0.2, 1)", 
              boxShadow: autoSummarize ? "none" : "var(--shadow-sm)",
            }}
          >
            <div>
              <div style={{ fontSize: 12, fontWeight: 600, color: autoSummarize ? "var(--text-primary)" : "var(--text-muted)", transition: "color 0.15s" }}>
                Auto-summarize slides
              </div>
              <div style={{ fontSize: 10, color: "var(--text-muted)", marginTop: 2 }}>
                AI generates per-slide summaries (uses tokens)
              </div>
            </div>
            <div style={{
              width: 34, height: 18, borderRadius: 99, position: "relative", flexShrink: 0,
              background: autoSummarize ? "var(--accent)" : "var(--bg-hover)", transition: "background 0.2s",
              border: autoSummarize ? "none" : "1px solid var(--border)",
            }}>
              <div style={{
                width: 14, height: 14, borderRadius: 99, background: "#fff",
                position: "absolute", top: 2, left: autoSummarize ? 18 : 2, transition: "left 0.2s",
                boxShadow: "0 1px 2px rgba(0,0,0,0.3)",
              }} />
            </div>
          </div>

          {/* Submit */}
          <button
            onClick={submit}
            disabled={!canSubmit}
            style={{
              width: "100%", padding: "14px 24px", borderRadius: 12,
              border: "none", cursor: canSubmit ? "pointer" : "default",
              background: canSubmit ? "linear-gradient(135deg, var(--accent) 0%, #4f46e5 100%)" : "var(--bg-hover)",
              color: canSubmit ? "#fff" : "var(--text-muted)",
              fontSize: 14, fontWeight: 700, fontFamily: "var(--font-sans)",
              transition: "all 0.3s", opacity: isProcessing ? 0.7 : 1,
              letterSpacing: "-0.01em",
              boxShadow: canSubmit ? "0 4px 12px rgba(99, 102, 241, 0.25)" : "none",
            }}
            onMouseEnter={e => { if(canSubmit) e.currentTarget.style.transform = "translateY(-1px)"; }}
            onMouseLeave={e => { if(canSubmit) e.currentTarget.style.transform = "translateY(0)"; }}
          >
            {isProcessing
              ? `Processing ${excelFiles.length} file${excelFiles.length > 1 ? "s" : ""}...`
              : `Generate Report${excelFiles.length > 1 ? ` (${excelFiles.length} Excel files)` : ""}`}
          </button>

          {/* Error */}
          {error && (
            <div className="animate-slide-up" style={{
              marginTop: 12, padding: "10px 14px", borderRadius: 10,
              background: "var(--error-dim)", border: "1px solid rgba(248,113,113,0.2)",
              fontSize: 13, color: "var(--error)", fontWeight: 500,
            }}>
              {error}
            </div>
          )}
        </div>

        {/* Pipeline Status */}
        {isProcessing && (
          <div className="animate-slide-up" style={{
            width: "100%", padding: 20, borderRadius: 14,
            background: "var(--bg-secondary)", border: "1px solid var(--border)",
          }}>
            <div style={{ fontSize: 11, fontWeight: 600, color: "var(--text-muted)", letterSpacing: "0.05em", textTransform: "uppercase", marginBottom: 14, fontFamily: "var(--font-mono)" }}>
              Pipeline
            </div>
            <div style={{ display: "flex", flexDirection: "column", gap: 0 }}>
              {STEPS.map((s, i) => {
                const isActive = s.key === step;
                const isDone = i < activeIdx;
                const isPending = i > activeIdx;

                return (
                  <div key={s.key} style={{ display: "flex", alignItems: "center", gap: 12, padding: "8px 0" }}>
                    <div style={{ display: "flex", flexDirection: "column", alignItems: "center", width: 20 }}>
                      <div style={{
                        width: isDone ? 18 : isActive ? 18 : 14,
                        height: isDone ? 18 : isActive ? 18 : 14,
                        borderRadius: 99,
                        background: isDone ? "var(--success)" : isActive ? "var(--accent)" : "var(--bg-hover)",
                        border: isPending ? "2px solid var(--border)" : "none",
                        display: "flex", alignItems: "center", justifyContent: "center",
                        transition: "all 0.3s",
                      }}>
                        {isDone && <span style={{ fontSize: 10, color: "#fff", fontWeight: 800 }}>✓</span>}
                        {isActive && (
                          <div style={{ width: 6, height: 6, borderRadius: 99, background: "#fff", animation: "pulse-soft 1s infinite" }} />
                        )}
                      </div>
                    </div>

                    <span style={{
                      fontSize: 13, fontWeight: isActive ? 600 : 400,
                      color: isDone ? "var(--success)" : isActive ? "var(--text-primary)" : "var(--text-muted)",
                      transition: "all 0.3s",
                    }}>
                      {s.label}
                      {isActive && <span style={{ marginLeft: 6, fontSize: 11, opacity: 0.5 }}>...</span>}
                    </span>
                  </div>
                );
              })}
            </div>
          </div>
        )}
      </div>
    </div>
  );
};

export default Landing;
