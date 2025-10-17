import React from "react";
import { createRoot } from "react-dom/client";
import App from "./App";

const container = document.getElementById("root");
const root = createRoot(container);

// Helper to detect if running inside MS Word/Excel/Office
function isRunningInOffice() {
  return (
    typeof window !== "undefined" &&
    (window.Office || window.location.href.includes("taskpane.html"))
  );
}

// Initialize Office Add-in
Office.onReady((info) => {
  if (info.host === "Word") {
    root.render(<App />);
  }
});

// If not running in Office environment, render the app directly
if (!isRunningInOffice()) {
  console.log("🖥️ Running locally — Office.js not loaded");
  root.render(<App />);
}
