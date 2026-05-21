import { useState, useRef, useEffect } from "react";
import { useSocket } from "../context/SocketContext";
import { useAIOptions } from "../context/AIOptionsContext";
import type { ChatMessage } from "../types";

interface Props {
  contextSlide: number | null;
  onClearContext: () => void;
}

const ChatPanel = ({ contextSlide, onClearContext }: Props) => {
  const { chatMessages, sendMessage, isAiTyping } = useSocket();
  const { options } = useAIOptions();
  const [text, setText] = useState("");
  const textareaRef = useRef<HTMLTextAreaElement>(null);
  const bottomRef = useRef<HTMLDivElement>(null);

  useEffect(() => { bottomRef.current?.scrollIntoView({ behavior: "smooth" }); }, [chatMessages.length, isAiTyping]);
  useEffect(() => {
    const el = textareaRef.current;
    if (el) { el.style.height = "auto"; el.style.height = `${Math.min(el.scrollHeight, 140)}px`; }
  }, [text]);

  const send = () => {
    const msg = text.trim();
    if (!msg || isAiTyping) return;
    // Pass AI options with the message so backend knows which behaviors to apply
    sendMessage(msg, contextSlide ?? undefined, options);
    setText("");
  };

  const onKey = (e: React.KeyboardEvent) => {
    if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); send(); }
  };

  const hasMessages = chatMessages.length > 0;
  const canSend = !!text.trim() && !isAiTyping;

  // Count active AI options for display
  const activeOptions = Object.entries(options).filter(([, v]) => v).map(([k]) => k);

  return (
    <div style={{ flex: 1, display: "flex", flexDirection: "column", height: "100%", minHeight: 0 }}>
      {/* Header */}
      <div style={{ padding: "11px 18px", borderBottom: "1px solid var(--border)", display: "flex", alignItems: "center", justifyContent: "space-between", background: "var(--bg-secondary)", flexShrink: 0 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 7 }}>
          <span style={{ width: 6, height: 6, borderRadius: 99, background: "var(--accent)", boxShadow: "0 0 6px rgba(167,139,250,0.3)" }} />
          <span style={{ fontSize: 13, fontWeight: 700, letterSpacing: "-0.02em" }}>Chat</span>
          {activeOptions.length > 0 && (
            <span style={{ fontSize: 9, padding: "2px 7px", borderRadius: 99, background: "var(--accent-dim)", color: "var(--accent)", fontWeight: 600, border: "1px solid var(--accent-border)" }}>
              {activeOptions.length} AI mode{activeOptions.length > 1 ? "s" : ""} active
            </span>
          )}
        </div>
        {contextSlide && (
          <div className="animate-fade" style={{ display: "flex", alignItems: "center", gap: 5, padding: "2px 9px", borderRadius: 99, background: "var(--accent-dim)", border: "1px solid var(--accent-border)", fontSize: 10, color: "var(--accent)", fontWeight: 600 }}>
            <span>Slide #{contextSlide}</span>
            <button onClick={onClearContext} style={{ background: "none", border: "none", color: "var(--accent)", cursor: "pointer", fontSize: 13, lineHeight: 1, padding: 0 }}>×</button>
          </div>
        )}
      </div>

      {/* Messages */}
      <div style={{ flex: 1, overflowY: "auto", padding: 18, minHeight: 0 }}>
        <div style={{ display: "flex", flexDirection: "column", ...(hasMessages ? { justifyContent: "flex-start" } : { justifyContent: "center", alignItems: "center", height: "100%" }) }}>
          {!hasMessages && (
            <div style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 10, maxWidth: 340, textAlign: "center" }}>
              <div style={{ width: 48, height: 48, borderRadius: 99, background: "var(--accent-dim)", border: "1px solid var(--accent-border)", display: "flex", alignItems: "center", justifyContent: "center" }}>
                <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="var(--accent)" strokeWidth="1.5" strokeLinecap="round"><path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z" /></svg>
              </div>
              <h3 style={{ fontSize: 14, fontWeight: 700, letterSpacing: "-0.02em", marginBottom: 2 }}>
                {contextSlide ? `Slide #${contextSlide} selected` : "Ask anything"}
              </h3>
              <p style={{ fontSize: 12, color: "var(--text-secondary)", lineHeight: 1.5 }}>
                {contextSlide ? "Type a question about this slide." : "Select a slide or ask about your report. AI settings from the sidebar affect responses."}
              </p>
              {!contextSlide && (
                <div style={{ display: "flex", gap: 5, marginTop: 2, flexWrap: "wrap", justifyContent: "center" }}>
                  {["Summarize all slides", "Key risks?", "Compare milestones"].map(h => (
                    <button key={h} onClick={() => { setText(h); setTimeout(() => textareaRef.current?.focus(), 0); }} style={{
                      fontSize: 10, padding: "4px 10px", borderRadius: 99, border: "1px solid var(--border)", background: "transparent",
                      color: "var(--text-muted)", cursor: "pointer", fontFamily: "var(--font-sans)",
                    }}>{h}</button>
                  ))}
                </div>
              )}
            </div>
          )}

          {hasMessages && (
            <div style={{ display: "flex", flexDirection: "column", gap: 12, width: "100%", maxWidth: 660, margin: "0 auto" }}>
              {chatMessages.map(msg => <Bubble key={msg.id} msg={msg} />)}
              {isAiTyping && (
                <div className="animate-fade" style={{ display: "flex", alignItems: "flex-start", gap: 7, alignSelf: "flex-start" }}>
                  <Avatar role="ai" />
                  <div style={{ padding: "9px 12px", borderRadius: "11px 11px 11px 3px", background: "var(--bg-tertiary)", border: "1px solid var(--border)", display: "flex", gap: 3, alignItems: "center" }}>
                    {[0, 1, 2].map(i => (<div key={i} style={{ width: 4, height: 4, borderRadius: 99, background: "var(--text-muted)", animation: `pulse-soft 1.2s ease-in-out ${i * 0.15}s infinite` }} />))}
                  </div>
                </div>
              )}
              <div ref={bottomRef} />
            </div>
          )}
        </div>
      </div>

      {/* Input */}
      <div style={{ borderTop: "1px solid var(--border)", padding: 12, background: "var(--bg-secondary)", flexShrink: 0 }}>
        <div style={{ display: "flex", alignItems: "flex-end", gap: 7, maxWidth: 660, margin: "0 auto", width: "100%" }}>
          <div style={{ flex: 1, borderRadius: 11, background: "var(--bg-tertiary)", border: "1px solid var(--border)", overflow: "hidden", transition: "border-color 0.2s" }}
            onFocus={e => { (e.currentTarget as HTMLElement).style.borderColor = "var(--accent-border)"; }}
            onBlur={e => { (e.currentTarget as HTMLElement).style.borderColor = "var(--border)"; }}
          >
            {contextSlide && (
              <div style={{ padding: "4px 11px", borderBottom: "1px solid var(--border-subtle)", fontSize: 9, color: "var(--accent)", fontWeight: 600, display: "flex", alignItems: "center", gap: 3 }}>
                Slide #{contextSlide}
              </div>
            )}
            <textarea ref={textareaRef} value={text} onChange={e => setText(e.target.value)} onKeyDown={onKey} rows={1} disabled={isAiTyping}
              placeholder={isAiTyping ? "AI is thinking..." : contextSlide ? `Ask about Slide #${contextSlide}...` : "Ask anything about your report..."}
              style={{ width: "100%", resize: "none", overflow: "hidden", background: "transparent", border: "none", outline: "none",
                padding: "9px 13px", color: "var(--text-primary)", fontSize: 13, fontFamily: "var(--font-sans)", lineHeight: 1.5, maxHeight: 140, opacity: isAiTyping ? 0.5 : 1 }} />
          </div>
          <button onClick={send} disabled={!canSend} style={{
            width: 38, height: 38, borderRadius: 9, border: "none", flexShrink: 0,
            background: canSend ? "var(--accent)" : "var(--bg-tertiary)", cursor: canSend ? "pointer" : "default",
            display: "flex", alignItems: "center", justifyContent: "center", transition: "all 0.2s",
          }}>
            <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke={canSend ? "#fff" : "var(--text-muted)"} strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
              <line x1="22" y1="2" x2="11" y2="13"/><polygon points="22 2 15 22 11 13 2 9 22 2"/>
            </svg>
          </button>
        </div>
      </div>
    </div>
  );
};

const Bubble = ({ msg }: { msg: ChatMessage }) => {
  const isUser = msg.role === "user";
  return (
    <div className="animate-slide-up" style={{ display: "flex", alignItems: "flex-start", gap: 7, alignSelf: isUser ? "flex-end" : "flex-start", flexDirection: isUser ? "row-reverse" : "row", maxWidth: "82%" }}>
      <Avatar role={msg.role} />
      <div style={{ padding: "9px 13px", borderRadius: isUser ? "11px 11px 3px 11px" : "11px 11px 11px 3px",
        background: isUser ? "var(--accent-dim)" : "var(--bg-tertiary)", border: `1px solid ${isUser ? "var(--accent-border)" : "var(--border)"}` }}>
        {msg.slideNumber && (
          <div style={{ display: "inline-flex", alignItems: "center", gap: 2, padding: "1px 6px", borderRadius: 4, marginBottom: 4,
            background: "var(--cyan-dim)", border: "1px solid rgba(34,211,238,0.15)", fontSize: 9, color: "var(--cyan)", fontWeight: 600 }}>Slide #{msg.slideNumber}</div>
        )}
        <p style={{ fontSize: 12, color: isUser ? "var(--text-primary)" : "var(--text-secondary)", lineHeight: 1.55, margin: 0, whiteSpace: "pre-wrap", wordBreak: "break-word" }}>{msg.content}</p>
        <span style={{ fontSize: 9, color: "var(--text-muted)", marginTop: 3, display: "block", textAlign: isUser ? "right" : "left" }}>
          {new Date(msg.timestamp).toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" })}
        </span>
      </div>
    </div>
  );
};

const Avatar = ({ role }: { role: "user" | "ai" }) => (
  <div style={{ width: 24, height: 24, borderRadius: 6, flexShrink: 0, background: role === "user" ? "var(--accent-dim)" : "var(--bg-hover)",
    border: `1px solid ${role === "user" ? "var(--accent-border)" : "var(--border)"}`, display: "flex", alignItems: "center", justifyContent: "center" }}>
    {role === "user"
      ? <svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="var(--accent)" strokeWidth="2" strokeLinecap="round"><path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/><circle cx="12" cy="7" r="4"/></svg>
      : <svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="var(--cyan)" strokeWidth="2" strokeLinecap="round"><path d="M12 2L2 7l10 5 10-5-10-5z"/><path d="M2 17l10 5 10-5"/></svg>}
  </div>
);

export default ChatPanel;
