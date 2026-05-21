import { useState } from "react";

export interface ProjectDetails {
  project_title: string;
  author: string;
  report_date: string;
  date_format: "UK" | "US";
  roles: Record<string, string>; // { "role label": "person name" }
}

export const DEFAULT_PROJECT_DETAILS: ProjectDetails = {
  project_title: "",
  author: "",
  report_date: new Date().toISOString().split("T")[0],
  date_format: "UK",
  roles: {},
};

// Default role labels from the PPT template
const DEFAULT_ROLE_LABELS = [
  "SWB - Project Manager",
  "SWB - Integration Architect",
  "SWB - Lead Architect",
  "Prolifics Senior Technical Architect",
  "Prolifics Technical Lead",
  "Prolifics Account Executive",
  "Prolifics Technology Manager & CSL",
  "Prolifics Delivery Partner",
];

interface Props {
  details: ProjectDetails;
  onChange: (d: ProjectDetails) => void;
  compact?: boolean; // true = sidebar mode (narrower)
}

const ProjectDetailsForm = ({ details, onChange, compact = false }: Props) => {
  const [newRoleLabel, setNewRoleLabel] = useState("");
  const [newRoleName, setNewRoleName] = useState("");

  const update = <K extends keyof ProjectDetails>(key: K, val: ProjectDetails[K]) => {
    onChange({ ...details, [key]: val });
  };

  const updateRole = (roleLabel: string, name: string) => {
    update("roles", { ...details.roles, [roleLabel]: name });
  };

  const removeRole = (roleLabel: string) => {
    const next = { ...details.roles };
    delete next[roleLabel];
    update("roles", next);
  };

  const addCustomRole = () => {
    const label = newRoleLabel.trim();
    if (!label) return;
    update("roles", { ...details.roles, [label]: newRoleName.trim() });
    setNewRoleLabel("");
    setNewRoleName("");
  };

  const prefillDefaults = () => {
    const merged = { ...details.roles };
    for (const label of DEFAULT_ROLE_LABELS) {
      if (!(label in merged)) merged[label] = "";
    }
    update("roles", merged);
  };

  const roleEntries = Object.entries(details.roles);
  const hasAllDefaults = DEFAULT_ROLE_LABELS.every(l => l in details.roles);

  const inputStyle: React.CSSProperties = {
    width: "100%", padding: compact ? "6px 8px" : "8px 10px",
    borderRadius: 8, border: "1px solid var(--border)",
    background: "#fff", color: "var(--text-primary)",
    fontSize: compact ? 11 : 12, fontFamily: "var(--font-sans)",
    outline: "none", transition: "all 0.2s",
    boxShadow: "var(--shadow-sm)",
  };

  const labelStyle: React.CSSProperties = {
    fontSize: compact ? 9 : 10, fontWeight: 600,
    color: "var(--text-muted)", letterSpacing: "0.04em",
    textTransform: "uppercase", marginBottom: 3, display: "block",
    fontFamily: "var(--font-mono)",
  };

  const sectionGap = compact ? 10 : 14;

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: sectionGap }}>

      {/* Project Title */}
      <div>
        <label style={labelStyle}>Project Title</label>
        <input
          style={inputStyle}
          placeholder="e.g. SWB - Van Delivery Integration"
          value={details.project_title}
          onChange={e => update("project_title", e.target.value)}
        />
      </div>

      {/* Author + Date row */}
      <div style={{ display: "flex", gap: 8 }}>
        <div style={{ flex: 1 }}>
          <label style={labelStyle}>Author</label>
          <input
            style={inputStyle}
            placeholder="Author name"
            value={details.author}
            onChange={e => update("author", e.target.value)}
          />
        </div>
        <div style={{ flex: 1 }}>
          <label style={labelStyle}>Report Date</label>
          <input
            type="date"
            style={{ ...inputStyle, cursor: "pointer" }}
            value={details.report_date}
            onChange={e => update("report_date", e.target.value)}
          />
        </div>
      </div>

      {/* Date Format */}
      <div>
        <label style={labelStyle}>Date Format</label>
        <div style={{ display: "flex", gap: 6 }}>
          {(["UK", "US"] as const).map(fmt => (
            <button
              key={fmt}
              onClick={() => update("date_format", fmt)}
              style={{
                flex: 1, padding: compact ? "5px 0" : "6px 0",
                borderRadius: 7, fontSize: compact ? 10 : 11, fontWeight: 600,
                fontFamily: "var(--font-mono)", cursor: "pointer",
                border: details.date_format === fmt ? "1.5px solid var(--accent)" : "1px solid var(--border)",
                background: details.date_format === fmt ? "rgba(167,139,250,0.08)" : "var(--bg-tertiary)",
                color: details.date_format === fmt ? "var(--accent)" : "var(--text-muted)",
                transition: "all 0.15s",
              }}
            >
              {fmt} {fmt === "UK" ? "(DD/MM/YYYY)" : "(MM/DD/YYYY)"}
            </button>
          ))}
        </div>
      </div>

      {/* Distribution / Roles */}
      <div>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 6 }}>
          <label style={{ ...labelStyle, marginBottom: 0 }}>Distribution Roles</label>
          {!hasAllDefaults && (
            <button
              onClick={prefillDefaults}
              style={{
                padding: "2px 8px", borderRadius: 5, fontSize: 9, fontWeight: 600,
                border: "1px solid var(--border)", background: "var(--bg-tertiary)",
                color: "var(--accent)", cursor: "pointer", fontFamily: "var(--font-mono)",
              }}
            >
              + Prefill defaults
            </button>
          )}
        </div>

        {/* Existing roles */}
        <div style={{ display: "flex", flexDirection: "column", gap: 5, maxHeight: compact ? 200 : 280, overflowY: "auto" }}>
          {roleEntries.map(([label, name]) => (
            <div key={label} style={{
              display: "flex", alignItems: "center", gap: 6,
              padding: "6px 10px", borderRadius: 10,
              background: "#fff", border: "1px solid var(--border)",
              boxShadow: "0 1px 3px rgba(0,0,0,0.02)",
            }}>
              {/* Role label (read-only) */}
              <div style={{
                flex: "0 0 auto", maxWidth: compact ? "45%" : "50%",
                fontSize: compact ? 9 : 10, fontWeight: 600,
                color: "var(--text-muted)", overflow: "hidden",
                textOverflow: "ellipsis", whiteSpace: "nowrap",
              }} title={label}>
                {label}
              </div>

              {/* Person name (editable) */}
              <input
                style={{
                  flex: 1, minWidth: 0, padding: "3px 6px",
                  borderRadius: 5, border: "1px solid var(--border)",
                  background: "var(--bg-secondary)", color: "var(--text-primary)",
                  fontSize: compact ? 10 : 11, outline: "none",
                }}
                placeholder="Name..."
                value={name}
                onChange={e => updateRole(label, e.target.value)}
              />

              {/* Remove */}
              <button
                onClick={() => removeRole(label)}
                style={{
                  width: 18, height: 18, borderRadius: 5, border: "none",
                  background: "transparent", cursor: "pointer", flexShrink: 0,
                  color: "var(--text-muted)", fontSize: 12,
                  display: "flex", alignItems: "center", justifyContent: "center",
                }}
                onMouseEnter={e => { e.currentTarget.style.color = "#f87171"; e.currentTarget.style.background = "rgba(248,113,113,0.1)"; }}
                onMouseLeave={e => { e.currentTarget.style.color = "var(--text-muted)"; e.currentTarget.style.background = "transparent"; }}
                title={`Remove ${label}`}
              >
                ✕
              </button>
            </div>
          ))}
        </div>

        {/* Add custom role */}
        <div style={{
          display: "flex", gap: 5, marginTop: 6,
          padding: "6px 8px", borderRadius: 8,
          background: "var(--bg-secondary)", border: "1px dashed var(--border)",
        }}>
          
          <input
            style={{
              flex: 1, minWidth: 0, padding: "3px 6px",
              borderRadius: 5, border: "1px solid var(--border)",
              background: "var(--bg-tertiary)", color: "var(--text-primary)",
              fontSize: compact ? 9 : 10, outline: "none",
            }}
            placeholder="Name..."
            value={newRoleName}
            onChange={e => setNewRoleName(e.target.value)}
            onKeyDown={e => e.key === "Enter" && addCustomRole()}
          />
          <input
            style={{
              flex: 1, minWidth: 0, padding: "3px 6px",
              borderRadius: 5, border: "1px solid var(--border)",
              background: "var(--bg-tertiary)", color: "var(--text-primary)",
              fontSize: compact ? 9 : 10, outline: "none",
            }}
            placeholder="Role title..."
            value={newRoleLabel}
            onChange={e => setNewRoleLabel(e.target.value)}
            onKeyDown={e => e.key === "Enter" && addCustomRole()}
          />
          <button
            onClick={addCustomRole}
            disabled={!newRoleLabel.trim()}
            style={{
              padding: "3px 8px", borderRadius: 5, border: "none",
              background: newRoleLabel.trim() ? "var(--accent)" : "var(--bg-hover)",
              color: newRoleLabel.trim() ? "#fff" : "var(--text-muted)",
              fontSize: 10, fontWeight: 700, cursor: newRoleLabel.trim() ? "pointer" : "default",
              flexShrink: 0, transition: "all 0.15s",
            }}
          >
            +
          </button>
        </div>
      </div>
    </div>
  );
};

export default ProjectDetailsForm;
