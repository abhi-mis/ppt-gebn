import {
  createContext, useContext, useEffect, useState, useCallback, type ReactNode,
} from "react";
import { socket } from "../api/socket";
import type { ChatMessage, SlideSummary, SlideImage, RefineProgress, RefineLogEntry, StyleConfig } from "../types";

interface SocketContextType {
  slideSummaries: SlideSummary[];
  summaryDone: boolean;
  summaryError: string | null;
  chatMessages: ChatMessage[];
  sendMessage: (message: string, slideNumber?: number, options?: Record<string, boolean>) => void;
  isAiTyping: boolean;
  slideImages: SlideImage[];
  slideImagesLoading: boolean;
  slideImagesDone: boolean;
  slideImagesError: string | null;
  requestSlideImages: (pptName: string) => void;
  refineStatus: string | null;
  refineProgress: RefineProgress | null;
  refineDone: { refined_count: number; total_found: number; message: string } | null;
  refineError: string | null;
  refineLogs: RefineLogEntry[];
  requestRefineComments: (pptName: string) => void;
  restyleStatus: string | null;
  restyleDone: boolean;
  restyleError: string | null;
  requestRestyle: (pptName: string, styleConfig: StyleConfig) => void;
  requestGenerateSummaries: (pptName: string) => void;
}

const SocketContext = createContext<SocketContextType | undefined>(undefined);

export const SocketProvider = ({ children }: { children: ReactNode }) => {
  const [slideSummaries, setSlideSummaries] = useState<SlideSummary[]>([]);
  const [summaryDone, setSummaryDone] = useState(false);
  const [summaryError, setSummaryError] = useState<string | null>(null);
  const [chatMessages, setChatMessages] = useState<ChatMessage[]>([]);
  const [isAiTyping, setIsAiTyping] = useState(false);
  const [slideImages, setSlideImages] = useState<SlideImage[]>([]);
  const [slideImagesLoading, setSlideImagesLoading] = useState(false);
  const [slideImagesDone, setSlideImagesDone] = useState(false);
  const [slideImagesError, setSlideImagesError] = useState<string | null>(null);
  const [refineStatus, setRefineStatus] = useState<string | null>(null);
  const [refineProgress, setRefineProgress] = useState<RefineProgress | null>(null);
  const [refineDone, setRefineDone] = useState<{ refined_count: number; total_found: number; message: string } | null>(null);
  const [refineError, setRefineError] = useState<string | null>(null);
  const [refineLogs, setRefineLogs] = useState<RefineLogEntry[]>([]);
  const [restyleStatus, setRestyleStatus] = useState<string | null>(null);
  const [restyleDone, setRestyleDone] = useState(false);
  const [restyleError, setRestyleError] = useState<string | null>(null);

  const getPptName = (): string | undefined => {
    try { const d = sessionStorage.getItem("report_data"); return d ? JSON.parse(d).new_fixed_ppt : undefined; }
    catch { return undefined; }
  };

  const sendMessage = useCallback((message: string, slideNumber?: number, options?: Record<string, boolean>) => {
    const text = message.trim();
    if (!text) return;
    setChatMessages(p => [...p, { id: `u-${Date.now()}`, role: "user", content: text, slideNumber, timestamp: Date.now() }]);
    setIsAiTyping(true);
    socket.emit("message", { message: text, slideNumber, ppt_name: getPptName(), options: options || {} });
  }, []);

  const requestSlideImages = useCallback((pptName: string) => {
    setSlideImages([]); setSlideImagesLoading(true); setSlideImagesDone(false); setSlideImagesError(null);
    socket.emit("generate_slide_images", { ppt_name: pptName });
  }, []);

  const requestRefineComments = useCallback((pptName: string) => {
    setRefineStatus("started"); setRefineProgress(null); setRefineDone(null); setRefineError(null); setRefineLogs([]);
    socket.emit("refine_comments", { ppt_name: pptName });
  }, []);

  const requestRestyle = useCallback((pptName: string, styleConfig: StyleConfig) => {
    setRestyleStatus("processing"); setRestyleDone(false); setRestyleError(null);
    socket.emit("restyle_ppt", { ppt_name: pptName, style_config: styleConfig });
  }, []);

  const requestGenerateSummaries = useCallback((pptName: string) => {
    setSlideSummaries([]); setSummaryDone(false); setSummaryError(null);
    socket.emit("generate_summaries", { ppt_name: pptName });
  }, []);

  useEffect(() => {
    socket.on("connect", () => console.log("Socket connected:", socket.id));
    socket.on("slide_summary", (d: SlideSummary) => setSlideSummaries(p => [...p, d]));
    socket.on("summary_complete", () => setSummaryDone(true));
    socket.on("summary_error", (d) => setSummaryError(d?.error || "Unknown error"));
    socket.on("message_response", (d) => {
      setIsAiTyping(false);
      const content = d?.message || d?.content || "";
      if (content) setChatMessages(p => [...p, { id: `ai-${Date.now()}`, role: "ai", content, slideNumber: d.slideNumber, timestamp: Date.now() }]);
    });
    socket.on("slide_image", (d: SlideImage) => {
      setSlideImages(p => { if (p.some(s => s.slide_number === d.slide_number)) return p; return [...p, d].sort((a, b) => a.slide_number - b.slide_number); });
    });
    socket.on("slide_images_complete", () => { setSlideImagesLoading(false); setSlideImagesDone(true); });
    socket.on("slide_images_error", (d) => { setSlideImagesLoading(false); setSlideImagesError(d?.error || "Error"); });
    socket.on("refine_status", (d) => setRefineStatus(d?.status || d?.message || null));
    socket.on("refine_progress", (d: RefineProgress) => setRefineProgress(d));
    socket.on("refine_complete", (d) => { setRefineStatus(null); setRefineDone(d); });
    socket.on("refine_error", (d) => { setRefineStatus(null); setRefineError(d?.error || "Refinement failed"); });
    socket.on("refine_log", (d: RefineLogEntry) => setRefineLogs(p => [...p, d]));
    socket.on("restyle_status", (d) => setRestyleStatus(d?.status || null));
    socket.on("restyle_complete", () => { setRestyleStatus(null); setRestyleDone(true); });
    socket.on("restyle_error", (d) => { setRestyleStatus(null); setRestyleError(d?.error || "Restyle failed"); });

    return () => {
      ["connect","slide_summary","summary_complete","summary_error","message_response",
       "slide_image","slide_images_complete","slide_images_error",
       "refine_status","refine_progress","refine_complete","refine_error","refine_log",
       "restyle_status","restyle_complete","restyle_error"
      ].forEach(e => socket.off(e));
    };
  }, []);

  return (
    <SocketContext.Provider value={{
      slideSummaries, summaryDone, summaryError, chatMessages, sendMessage, isAiTyping,
      slideImages, slideImagesLoading, slideImagesDone, slideImagesError, requestSlideImages,
      refineStatus, refineProgress, refineDone, refineError, refineLogs, requestRefineComments,
      restyleStatus, restyleDone, restyleError, requestRestyle,
      requestGenerateSummaries,
    }}>
      {children}
    </SocketContext.Provider>
  );
};

export const useSocket = () => {
  const ctx = useContext(SocketContext);
  if (!ctx) throw new Error("useSocket must be within SocketProvider");
  return ctx;
};
