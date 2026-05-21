import { useEffect, useRef } from "react";
import { useSocket } from "../context/SocketContext";
import type { SlideSummary } from "../types";

interface Props {
  onSlideClick: (s: SlideSummary) => void;
  selectedSlide: number | null;
}

const ActivityFeed = ({ onSlideClick, selectedSlide }: Props) => {
  const {
    slideSummaries,
    summaryDone,
    summaryError,
    refineStatus,
    refineProgress,
    refineDone,
    refineLogs,
  } = useSocket();
  const endRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    endRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [slideSummaries.length, refineStatus, refineDone, refineLogs.length]);

  const total =
    slideSummaries.length > 0
      ? parseInt(
          slideSummaries[slideSummaries.length - 1].progress.split("/")[1],
        ) || slideSummaries.length
      : 0;

  return (
    <div
      style={{
        width: 320,
        minWidth: 320,
        height: "100%",
        borderRight: "1px solid var(--border)",
        background: "var(--bg-primary)",
        display: "flex",
        flexDirection: "column",
      }}
    >
      {/* Header */}
      <div
        style={{
          padding: "13px 16px",
          borderBottom: "1px solid var(--border)",
          display: "flex",
          alignItems: "center",
          justifyContent: "space-between",
        }}
      >
        <div style={{ display: "flex", alignItems: "center", gap: 7 }}>
          <svg
            width="14"
            height="14"
            viewBox="0 0 24 24"
            fill="none"
            stroke="var(--accent)"
            strokeWidth="2"
            strokeLinecap="round"
          >
            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
            <polyline points="14,2 14,8 20,8" />
          </svg>
          <span
            style={{ fontSize: 13, fontWeight: 700, letterSpacing: "-0.02em" }}
          >
            Activity
          </span>
        </div>
        {slideSummaries.length > 0 && (
          <span
            style={{
              fontSize: 10,
              fontWeight: 700,
              color: "#fff",
              background: "var(--accent)",
              padding: "2px 7px",
              borderRadius: 99,
            }}
          >
            {slideSummaries.length}
            {total ? `/${total}` : ""}
          </span>
        )}
      </div>

      {/* Feed */}
      <div
        style={{
          flex: 1,
          overflowY: "auto",
          padding: 10,
          display: "flex",
          flexDirection: "column",
          gap: 5,
        }}
      >
        {slideSummaries.length === 0 &&
          !summaryError &&
          refineLogs.length === 0 && (
            <div
              style={{
                display: "flex",
                flexDirection: "column",
                alignItems: "center",
                justifyContent: "center",
                height: "100%",
                gap: 8,
                opacity: 0.3,
              }}
            >
              <svg
                width="28"
                height="28"
                viewBox="0 0 24 24"
                fill="none"
                stroke="var(--text-muted)"
                strokeWidth="1.5"
                strokeLinecap="round"
              >
                <rect x="2" y="3" width="20" height="14" rx="2" />
                <path d="M8 21h8M12 17v4" />
              </svg>
              <span style={{ fontSize: 11, color: "var(--text-muted)" }}>
                Waiting for activity...
              </span>
            </div>
          )}

        {/* Slide Summaries */}
        {slideSummaries.map((item, i) => {
          const isSelected = selectedSlide === item.slide_number;
          return (
            <div
              key={`${item.job_id}-${item.slide_number}-${i}`}
              className="animate-slide-up"
              onClick={() => onSlideClick(item)}
              style={{
                padding: 14,
                borderRadius: 12,
                border: `1px solid ${isSelected ? "var(--accent)" : "var(--border)"}`,
                background: isSelected
                  ? "var(--accent-dim)"
                  : "#fff",
                cursor: "pointer",
                transition: "all 0.2s cubic-bezier(0.4, 0, 0.2, 1)",
                boxShadow: isSelected ? "none" : "var(--shadow-sm)",
                animationDelay: `${i * 30}ms`,
              }}
              onMouseEnter={(e) => {
                if (!isSelected)
                  (e.currentTarget as HTMLElement).style.background =
                    "var(--bg-tertiary)";
              }}
              onMouseLeave={(e) => {
                if (!isSelected)
                  (e.currentTarget as HTMLElement).style.background =
                    "var(--bg-secondary)";
              }}
            >
              <div
                style={{
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "space-between",
                  marginBottom: 5,
                }}
              >
                <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                  <span
                    style={{
                      width: 22,
                      height: 22,
                      borderRadius: 5,
                      background: isSelected
                        ? "var(--accent)"
                        : "var(--bg-hover)",
                      display: "flex",
                      alignItems: "center",
                      justifyContent: "center",
                      fontSize: 10,
                      fontWeight: 700,
                      fontFamily: "var(--font-mono)",
                      color: isSelected ? "#fff" : "var(--text-secondary)",
                      transition: "all 0.15s",
                    }}
                  >
                    {item.slide_number}
                  </span>
                  <span style={{ fontSize: 11, fontWeight: 600 }}>
                    Slide {item.slide_number}
                  </span>
                </div>
                <span
                  style={{
                    fontSize: 9,
                    color: "var(--text-muted)",
                    fontFamily: "var(--font-mono)",
                  }}
                >
                  {item.progress}
                </span>
              </div>
              <p
                style={{
                  fontSize: 11,
                  color: "var(--text-secondary)",
                  lineHeight: 1.45,
                  margin: 0,
                  display: "-webkit-box",
                  WebkitLineClamp: 2,
                  WebkitBoxOrient: "vertical" as const,
                  overflow: "hidden",
                }}
              >
                {item.summary}
              </p>
            </div>
          );
        })}

        {/* Loading */}
        {slideSummaries.length > 0 && !summaryDone && !summaryError && (
          <div
            className="animate-fade"
            style={{
              padding: "8px 10px",
              borderRadius: 7,
              background: "var(--bg-secondary)",
              border: "1px solid var(--border-subtle)",
              display: "flex",
              alignItems: "center",
              gap: 6,
            }}
          >
            <div
              style={{
                width: 5,
                height: 5,
                borderRadius: 99,
                background: "var(--accent)",
              }}
              className="animate-pulse-soft"
            />
            <span style={{ fontSize: 10, color: "var(--text-muted)" }}>
              Processing next slide...
            </span>
          </div>
        )}

        {summaryDone && (
          <StatusPill
            color="var(--success)"
            bg="var(--success-dim)"
            text={`✓ All ${slideSummaries.length} slides processed`}
          />
        )}

        {/* Refine Logs */}
        {refineLogs.length > 0 && (
          <>
            <div
              style={{
                fontSize: 10,
                fontWeight: 700,
                color: "var(--text-muted)",
                letterSpacing: "0.06em",
                textTransform: "uppercase",
                fontFamily: "var(--font-mono)",
                marginTop: 8,
                marginBottom: 4,
              }}
            >
              Refine Log ({refineLogs.length})
            </div>
            {refineLogs.map((log, i) => (
              <div
                key={i}
                className="animate-slide-up"
                style={{
                  padding: "8px 10px",
                  borderRadius: 7,
                  fontSize: 10,
                  lineHeight: 1.5,
                  background:
                    log.status === "refined"
                      ? "rgba(167,139,250,0.06)"
                      : "var(--bg-secondary)",
                  border: `1px solid ${log.status === "refined" ? "var(--accent-border)" : "var(--border-subtle)"}`,
                }}
              >
                <div
                  style={{
                    display: "flex",
                    alignItems: "center",
                    justifyContent: "space-between",
                    marginBottom: 3,
                  }}
                >
                  <span
                    style={{ fontWeight: 600, color: "var(--text-secondary)" }}
                  >
                    Slide {log.slide} · {log.column}
                  </span>
                  <span
                    style={{
                      fontSize: 9,
                      fontWeight: 600,
                      padding: "1px 6px",
                      borderRadius: 4,
                      background:
                        log.status === "refined"
                          ? "var(--accent-dim)"
                          : log.status === "unchanged"
                            ? "var(--bg-hover)"
                            : "var(--error-dim)",
                      color:
                        log.status === "refined"
                          ? "var(--accent)"
                          : log.status === "unchanged"
                            ? "var(--text-muted)"
                            : "var(--error)",
                    }}
                  >
                    {log.status}
                  </span>
                </div>
                {log.status === "refined" && (
                  <>
                    <div
                      style={{ color: "var(--text-muted)", marginBottom: 2 }}
                    >
                      <span style={{ fontWeight: 600 }}>Before:</span>{" "}
                      {log.original}
                    </div>
                    <div style={{ color: "var(--accent)" }}>
                      <span style={{ fontWeight: 600 }}>After:</span>{" "}
                      {log.refined}
                    </div>
                  </>
                )}
                {log.status === "skipped" && log.reason && (
                  <div style={{ color: "var(--text-muted)" }}>{log.reason}</div>
                )}
              </div>
            ))}
          </>
        )}

        {/* Refine progress in feed */}
        {refineStatus && !refineDone && (
          <div
            className="animate-slide-up"
            style={{
              padding: "8px 10px",
              borderRadius: 7,
              background: "var(--accent-dim)",
              border: "1px solid var(--accent-border)",
            }}
          >
            <div style={{ display: "flex", alignItems: "center", gap: 5 }}>
              <div
                style={{
                  width: 12,
                  height: 12,
                  border: "2px solid var(--accent)",
                  borderTopColor: "transparent",
                  borderRadius: 99,
                  animation: "spin 0.8s linear infinite",
                }}
              />
              <span
                style={{
                  fontSize: 11,
                  color: "var(--accent)",
                  fontWeight: 500,
                }}
              >
                Refining
                {refineProgress
                  ? ` (${refineProgress.processed}/${refineProgress.total})`
                  : "..."}
              </span>
            </div>
          </div>
        )}

        {refineDone && (
          <StatusPill
            color="var(--success)"
            bg="var(--success-dim)"
            text={`✓ ${refineDone.message}`}
          />
        )}
        {summaryError && (
          <StatusPill
            color="var(--error)"
            bg="var(--error-dim)"
            text={`Error: ${summaryError}`}
          />
        )}

        <div ref={endRef} />
      </div>
    </div>
  );
};

const StatusPill = ({
  color,
  bg,
  text,
}: {
  color: string;
  bg: string;
  text: string;
}) => (
  <div
    className="animate-slide-up"
    style={{
      padding: "7px 10px",
      borderRadius: 7,
      background: bg,
      border: `1px solid ${color}22`,
      display: "flex",
      alignItems: "center",
      gap: 5,
      fontSize: 11,
      color,
      fontWeight: 500,
    }}
  >
    <span>{text}</span>
  </div>
);

export default ActivityFeed;
