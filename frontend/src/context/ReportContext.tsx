import { createContext, useContext, useState, type ReactNode } from "react";
import type { ReportData } from "../types";

interface ReportContextType {
  report: ReportData | null;
  setReport: (data: ReportData) => void;
  clearReport: () => void;
}

const ReportContext = createContext<ReportContextType>({ report: null, setReport: () => {}, clearReport: () => {} });
const KEY = "report_data";

export const ReportProvider = ({ children }: { children: ReactNode }) => {
  const [report, setReportState] = useState<ReportData | null>(() => {
    try { const s = sessionStorage.getItem(KEY); return s ? JSON.parse(s) : null; } catch { return null; }
  });
  const setReport = (data: ReportData) => { setReportState(data); sessionStorage.setItem(KEY, JSON.stringify(data)); };
  const clearReport = () => { setReportState(null); sessionStorage.removeItem(KEY); };
  return <ReportContext.Provider value={{ report, setReport, clearReport }}>{children}</ReportContext.Provider>;
};

export const useReport = () => useContext(ReportContext);
