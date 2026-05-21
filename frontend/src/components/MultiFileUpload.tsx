import { useRef, useState } from "react";

interface Props {
  label: string;
  accept: string;
  files: File[];
  onFilesChange: (files: File[]) => void;
}

const MultiFileUpload = ({ label, accept, files, onFilesChange }: Props) => {
  const ref = useRef<HTMLInputElement>(null);
  const [over, setOver] = useState(false);

  const color = "#107c10"; // Microsoft Excel Green

  const addFiles = (newFiles: FileList | File[]) => {
    const incoming = Array.from(newFiles);
    // Deduplicate by name+size
    const existing = new Set(files.map(f => `${f.name}__${f.size}`));
    const unique = incoming.filter(f => !existing.has(`${f.name}__${f.size}`));
    if (unique.length > 0) {
      onFilesChange([...files, ...unique]);
    }
  };

  const removeFile = (index: number) => {
    onFilesChange(files.filter((_, i) => i !== index));
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setOver(false);
    if (e.dataTransfer.files.length > 0) {
      addFiles(e.dataTransfer.files);
    }
  };

  const handleInputChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      addFiles(e.target.files);
    }
    // Reset input so the same file can be re-added if removed
    e.target.value = "";
  };

  const formatSize = (bytes: number) => {
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
  };

  return (
    <div style={{ flex: 1, display: "flex", flexDirection: "column", gap: 8 }}>
      {/* Drop zone */}
      <div
        onClick={() => ref.current?.click()}
        onDragOver={(e) => { e.preventDefault(); setOver(true); }}
        onDragLeave={() => setOver(false)}
        onDrop={handleDrop}
        style={{
          padding: "20px 16px",
          borderRadius: 16,
          border: `2px dashed ${over ? color : files.length > 0 ? `${color}40` : "var(--border)"}`,
          background: over ? `${color}08` : files.length > 0 ? `${color}04` : "#fff",
          cursor: "pointer",
          transition: "all 0.2s cubic-bezier(0.4, 0, 0.2, 1)",
          display: "flex",
          flexDirection: "column",
          alignItems: "center",
          gap: 10,
          position: "relative",
        }}
      >
        <input
          ref={ref}
          type="file"
          accept={accept}
          multiple
          onChange={handleInputChange}
          style={{ display: "none" }}
        />

        {/* Icon */}
        <div style={{
          width: 40, height: 40, borderRadius: 10,
          background: `${color}15`, display: "flex", alignItems: "center", justifyContent: "center",
        }}>
          <span style={{
            fontSize: 11, fontWeight: 700, color, fontFamily: "var(--font-mono)", letterSpacing: "0.05em",
          }}>
            .XLS
          </span>
        </div>

        <span style={{ fontSize: 13, fontWeight: 600, color: "var(--text-primary)" }}>{label}</span>
        <span style={{ fontSize: 11, color: "var(--text-muted)" }}>
          {files.length === 0
            ? "Click or drag to upload (multiple)"
            : `${files.length} file${files.length > 1 ? "s" : ""} selected — click to add more`}
        </span>

        {files.length > 0 && (
          <div style={{
            position: "absolute", top: 8, right: 8,
            padding: "2px 8px", borderRadius: 99,
            background: color, fontSize: 11, fontWeight: 700, color: "#fff",
          }}>
            {files.length}
          </div>
        )}
      </div>

      {/* File list */}
      {files.length > 0 && (
        <div style={{
          display: "flex", flexDirection: "column", gap: 4,
          maxHeight: 160, overflowY: "auto",
          padding: "0 2px",
        }}>
          {files.map((file, idx) => (
            <div
              key={`${file.name}__${file.size}__${idx}`}
              style={{
                display: "flex", alignItems: "center", gap: 8,
                padding: "8px 12px", borderRadius: 10,
                background: "#fff", border: "1px solid var(--border)",
                boxShadow: "var(--shadow-sm)",
                transition: "all 0.2s",
              }}
              onMouseEnter={e => { e.currentTarget.style.borderColor = color; e.currentTarget.style.transform = "translateX(2px)"; }}
              onMouseLeave={e => { e.currentTarget.style.borderColor = "var(--border)"; e.currentTarget.style.transform = "translateX(0)"; }}
            >
              {/* File icon */}
              <div style={{
                width: 24, height: 24, borderRadius: 6, flexShrink: 0,
                background: `${color}18`, display: "flex", alignItems: "center", justifyContent: "center",
              }}>
                <span style={{ fontSize: 8, fontWeight: 700, color, fontFamily: "var(--font-mono)" }}>.XLS</span>
              </div>

              {/* Name + size */}
              <div style={{ flex: 1, minWidth: 0 }}>
                <div style={{
                  fontSize: 11, fontWeight: 500, color: "var(--text-primary)",
                  overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap",
                }}>
                  {file.name}
                </div>
                <div style={{ fontSize: 9, color: "var(--text-muted)" }}>
                  {formatSize(file.size)}
                </div>
              </div>

              {/* Remove button */}
              <button
                onClick={(e) => { e.stopPropagation(); removeFile(idx); }}
                style={{
                  width: 20, height: 20, borderRadius: 6, border: "none",
                  background: "transparent", cursor: "pointer",
                  display: "flex", alignItems: "center", justifyContent: "center",
                  color: "var(--text-muted)", fontSize: 14, fontWeight: 400,
                  transition: "all 0.15s", flexShrink: 0,
                }}
                onMouseEnter={e => {
                  (e.currentTarget as HTMLElement).style.background = "rgba(248,113,113,0.15)";
                  (e.currentTarget as HTMLElement).style.color = "#f87171";
                }}
                onMouseLeave={e => {
                  (e.currentTarget as HTMLElement).style.background = "transparent";
                  (e.currentTarget as HTMLElement).style.color = "var(--text-muted)";
                }}
                title={`Remove ${file.name}`}
              >
                ✕
              </button>
            </div>
          ))}
        </div>
      )}
    </div>
  );
};

export default MultiFileUpload;
