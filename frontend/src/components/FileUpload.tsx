import { useRef, useState } from "react";

interface Props {
  label: string;
  accept: string;
  file: File | null;
  onFile: (f: File) => void;
  icon: "xlsx" | "pptx";
}

const FileUpload = ({ label, accept, file, onFile, icon }: Props) => {
  const ref = useRef<HTMLInputElement>(null);
  const [over, setOver] = useState(false);

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault(); setOver(false);
    const f = e.dataTransfer.files[0];
    if (f) onFile(f);
  };

  const color = icon === "xlsx" ? "#4ade80" : "#f97316";

  return (
    <div
      onClick={() => ref.current?.click()}
      onDragOver={(e) => { e.preventDefault(); setOver(true); }}
      onDragLeave={() => setOver(false)}
      onDrop={handleDrop}
      style={{
        flex: 1,
        padding: "20px 16px",
        borderRadius: 12,
        border: `1.5px dashed ${over ? color : file ? color : "var(--border)"}`,
        background: over ? "rgba(255,255,255,0.02)" : file ? `${color}08` : "var(--bg-secondary)",
        cursor: "pointer",
        transition: "all 0.2s",
        display: "flex",
        flexDirection: "column",
        alignItems: "center",
        gap: 8,
        position: "relative",
      }}
    >
      <input ref={ref} type="file" accept={accept} onChange={e => { const f = e.target.files?.[0]; if (f) onFile(f); }} style={{ display: "none" }} />

      {/* Icon */}
      <div style={{
        width: 40, height: 40, borderRadius: 10,
        background: `${color}15`, display: "flex", alignItems: "center", justifyContent: "center",
      }}>
        <span style={{ fontSize: 11, fontWeight: 700, color, fontFamily: "var(--font-mono)", letterSpacing: "0.05em" }}>
          {icon === "xlsx" ? ".XLS" : ".PPT"}
        </span>
      </div>

      <span style={{ fontSize: 13, fontWeight: 600, color: "var(--text-primary)" }}>{label}</span>
      <span style={{ fontSize: 11, color: "var(--text-muted)" }}>
        {file ? file.name : "Click or drag to upload"}
      </span>

      {file && (
        <div style={{
          position: "absolute", top: 8, right: 8, width: 18, height: 18,
          borderRadius: 99, background: color,
          display: "flex", alignItems: "center", justifyContent: "center",
          fontSize: 11, fontWeight: 700, color: "#000",
        }}>✓</div>
      )}
    </div>
  );
};

export default FileUpload;
