import { useReport } from "./context/ReportContext";
import Landing from "./components/Landing";
import Workspace from "./components/Workspace";

const App = () => {
  const { report } = useReport();
  return report ? <Workspace /> : <Landing />;
};

export default App;
