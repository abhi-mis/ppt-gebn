import { createContext, useContext, useState, type ReactNode } from "react";

type Opts = Record<string, boolean>;
interface AIOptionsCtx { options: Opts; setOptions: (d: Opts) => void; }

const Ctx = createContext<AIOptionsCtx>({ options: {}, setOptions: () => {} });
const KEY = "ai_options";

export const AIOptionsProvider = ({ children }: { children: ReactNode }) => {
  const [options, setOpts] = useState<Opts>(() => {
    try { const s = sessionStorage.getItem(KEY); return s ? JSON.parse(s) : {}; }
    catch { return {}; }
  });
  const setOptions = (d: Opts) => { setOpts(d); sessionStorage.setItem(KEY, JSON.stringify(d)); };
  return <Ctx.Provider value={{ options, setOptions }}>{children}</Ctx.Provider>;
};

export const useAIOptions = () => useContext(Ctx);
