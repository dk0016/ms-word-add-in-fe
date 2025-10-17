import React from "react";
import { createRoot } from "react-dom/client";
import App from "./App";

const container = document.getElementById("root");
const root = createRoot(container);

// Initialize Office Add-in
Office.onReady((info) => {
  if (info.host === "Word") {
    root.render(<App />);
  }
});

// ✅ Conditionally load Office.js only when inside Office environment
if (isRunningInOffice()) {
  // Publish the loading promise on window so other modules can await it.
  window.__officeReady = loadOfficeJs()
    .then(() => {
      console.log("✅ Office.js loaded successfully");
    })
    .catch((e) => console.warn("⚠️ Office.js failed to load:", e));
} else {
  console.log("🖥️ Running locally — Office.js not loaded");
}
