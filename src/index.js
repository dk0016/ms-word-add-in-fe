import React from "react";
import { createRoot } from "react-dom/client";
import App from "./App";
const container = document.getElementById("root");

// ✅ Function to dynamically load Office.js only when needed
async function loadOfficeJs() {
  if (typeof window !== "undefined" && !window.Office) {
    return new Promise((resolve) => {
      const script = document.createElement("script");
      script.src = "https://appsforoffice.microsoft.com/lib/1/hosted/office.js";
      script.onload = () => {
        try {
          if (window.Office && typeof window.Office.onReady === "function") {
            const ready = window.Office.onReady();
            if (ready && typeof ready.then === "function") {
              ready.then(() => resolve()).catch(() => resolve());
            } else {
              window.Office.onReady(() => resolve());
            }
          } else {
            resolve();
          }
        } catch (err) {
          console.warn("Office onReady failed:", err);
          resolve();
        }
      };
      script.onerror = (e) => {
        console.warn("Office.js failed to load", e);
        resolve(); // Still resolve to prevent app hang
      };
      document.head.appendChild(script);
    });
  }
  return Promise.resolve();
}

// ✅ Helper to detect if running inside MS Word/Excel/Office
function isRunningInOffice() {
  return (
    typeof window !== "undefined" &&
    (window.Office || window.location.href.includes("taskpane.html"))
  );
}
const root = createRoot(container);
root.render(<App />);

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
