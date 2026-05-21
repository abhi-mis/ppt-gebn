export interface ReportData {
  extractedJson: string[];
  new_fixed_ppt: string;
  friendly_name?: string;
  pptDownloadUrl: string;
}

export interface ChatMessage {
  id: string;
  role: "user" | "ai";
  content: string;
  slideNumber?: number;
  timestamp: number;
}

export interface SlideSummary {
  job_id: string;
  slide_number: number;
  summary: string;
  progress: string;
}

export interface SlideImage {
  slide_number: number;
  total_slides: number;
  image: string;
  mime_type: string;
  ppt_name: string;
}

export interface RefineProgress {
  total: number;
  processed: number;
  batch_info?: string;
  ppt_name?: string;
}

export interface RefineLogEntry {
  index: number;
  total: number;
  slide: number;
  column: string;
  status: "refined" | "unchanged" | "skipped";
  reason?: string;
  original: string;
  refined: string;
}

export interface StyleConfig {
  hdr_bg_color: string;
  hdr_text_color: string;
  hdr_font: string;
  hdr_size: number;
  hdr_bold: boolean;
  data_font: string;
  data_size: number;
  title_font: string;
  title_size: number;
  title_color: string;
  title_bold: boolean;
}

export const DEFAULT_STYLE_CONFIG: StyleConfig = {
  hdr_bg_color: "b8cce4",
  hdr_text_color: "000000",
  hdr_font: "Calibri",
  hdr_size: 11,
  hdr_bold: true,
  data_font: "Calibri",
  data_size: 11,
  title_font: "Tahoma",
  title_size: 20,
  title_color: "14366B",
  title_bold: true,
};
