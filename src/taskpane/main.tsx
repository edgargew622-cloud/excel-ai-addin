import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import Taskpane from "./Taskpane";
import "./styles.css";

// Рендерим только после Office.onReady: до этого Excel.run недоступен.
Office.onReady((info) => {
  const root = createRoot(document.getElementById("root")!);

  if (info.host !== Office.HostType.Excel) {
    root.render(<p style={{ padding: 16 }}>Эта надстройка работает только в Excel.</p>);
    return;
  }

  root.render(
    <StrictMode>
      <Taskpane />
    </StrictMode>
  );
});
