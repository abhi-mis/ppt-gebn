import { useState } from "react";
import Sidebar from "./Sidebar";
import ActivityFeed from "./ActivityFeed";
import ChatPanel from "./ChatPanel";
import type { SlideSummary } from "../types";

const Workspace = () => {
  const [contextSlide, setContextSlide] = useState<number | null>(null);

  const handleSlideClick = (s: SlideSummary) => {
    setContextSlide(s.slide_number);
  };

  return (
    <div style={{
      display: "flex", height: "100vh", maxHeight: "100vh",
      background: "var(--bg-primary)", color: "var(--text-primary)",
      overflow: "hidden",
    }}>
      <Sidebar />
      <ActivityFeed onSlideClick={handleSlideClick} selectedSlide={contextSlide} />
      <ChatPanel contextSlide={contextSlide} onClearContext={() => setContextSlide(null)} />
    </div>
  );
};

export default Workspace;
