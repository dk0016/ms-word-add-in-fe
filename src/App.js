// import React, { useEffect, useState } from "react";
// import { insertHtmlToWord } from "./wordHelper";

// export default function App() {
//   const [files, setFiles] = useState([]);
//   const [loading, setLoading] = useState(false);
//   const [selected, setSelected] = useState(null);

//   useEffect(() => {
//     async function load() {
//       try {
//         const res = await fetch(
//           "https://ms-word-add-in-backend.vercel.app/files"
//         );
//         const js = await res.json();
//         setFiles(js);
//       } catch (e) {
//         console.error(e);
//       }
//     }
//     load();
//   }, []);

//   const onClick = async (f) => {
//     setSelected(f.name);
//     setLoading(true);
//     try {
//       const res = await fetch(
//         `https://ms-word-add-in-backend.vercel.app/file?name=${encodeURIComponent(
//           f.name
//         )}`
//       );
//       const js = await res.json(); // { html }
//       await insertHtmlToWord(js.html);
//     } catch (e) {
//       console.error(e);
//       alert("Error loading file: " + e.message);
//     } finally {
//       setLoading(false);
//     }
//   };

//   return (
//     <div style={{ padding: 12, fontFamily: "Segoe UI, sans-serif" }}>
//       <h3>📁 .docx Files</h3>
//       {files.length === 0 && (
//         <div>No files found. Put .docx files into server/files/</div>
//       )}
//       <ul style={{ listStyle: "none", padding: 0 }}>
//         {files.map((f) => (
//           <li
//             key={f.name}
//             onClick={() => onClick(f)}
//             style={{
//               padding: "8px",
//               margin: "6px 0",
//               borderRadius: 6,
//               cursor: "pointer",
//               background: selected === f.name ? "#e6f2ff" : "#f7f7f7",
//             }}
//           >
//             {f.name}
//           </li>
//         ))}
//       </ul>
//       {loading && <div>Loading file into Word...</div>}
//     </div>
//   );
// }

import React, { useEffect, useState } from "react";
import { insertHtmlToWord, clearWordBody } from "./wordHelper";

export default function App() {
  const [files, setFiles] = useState([]);
  const [loading, setLoading] = useState(false);
  const [selected, setSelected] = useState(null);

  const onClick = async (f) => {
    setSelected(f.name);
    setLoading(true);

    try {
      // Immediately clear Word panel
      await clearWordBody();
      await insertHtmlToWord("<p>Loading document...</p>");
    } catch (e) {
      console.error("Error clearing Word panel", e);
    }

    try {
      const res = await fetch(
        `https://ms-word-add-in-backend.vercel.app/file?name=${encodeURIComponent(
          f.name
        )}`
      );
      const js = await res.json(); // { html }

      await insertHtmlToWord(js.html);
    } catch (e) {
      console.error(e);
      alert("Error loading file: " + e.message);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    async function load() {
      try {
        const res = await fetch(
          "https://ms-word-add-in-backend.vercel.app/files"
        );
        const js = await res.json();
        setFiles(js);
      } catch (e) {
        console.error(e);
      }
    }
    load();
  }, []);

  return (
    <div style={{ padding: 12, fontFamily: "Segoe UI, sans-serif" }}>
      <h3>📁 .docx Files</h3>
      {files.length === 0 && (
        <div>No files found. Put .docx files into server/files/</div>
      )}
      <ul style={{ listStyle: "none", padding: 0 }}>
        {files.map((f) => (
          <li
            key={f.name}
            onClick={() => onClick(f)}
            style={{
              padding: "8px",
              margin: "6px 0",
              borderRadius: 6,
              cursor: "pointer",
              background: selected === f.name ? "#e6f2ff" : "#f7f7f7",
            }}
          >
            {f.name}
          </li>
        ))}
      </ul>
      {loading && <div>Loading file into Word...</div>}
    </div>
  );
}
