import React from "react";
import { createRoot } from "react-dom/client";
import App from "./App";

const container = document.getElementById("root");
const root = createRoot(container);

async function loadOfficeJs() {
  if (typeof window !== "undefined" && !window.Office) {
    return new Promise((resolve) => {
      const script = document.createElement("script");
      script.src = "https://appsforoffice.microsoft.com/lib/1/hosted/office.js";
      script.onload = () => resolve();
      script.onerror = () => resolve();
      document.head.appendChild(script);
    });
  }
  return Promise.resolve();
}

function isRunningInOffice() {
  return (
    typeof window !== "undefined" &&
    (window.Office || window.location.href.includes("taskpane.html"))
  );
}

async function init() {
  if (isRunningInOffice()) {
    await loadOfficeJs();

    if (window.Office && typeof window.Office.onReady === "function") {
      window.Office.onReady((info) => {
        console.log("✅ Office ready:", info);
        root.render(<App />);
      });
    } else {
      console.log("⚠️ Office.js failed to initialize, rendering anyway");
      root.render(<App />);
    }
  } else {
    console.log("🖥️ Running locally — Office.js not loaded");
    root.render(<App />);
  }
}

init();
