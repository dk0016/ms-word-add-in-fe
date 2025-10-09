// export async function insertHtmlToWord(html) {
//   try {
//     await Office.onReady();
//     await Word.run(async (context) => {
//       const body = context.document.body;
//       body.clear();
//       // Insert HTML at start
//       body.insertHtml(html, Word.InsertLocation.start);
//       await context.sync();
//     });
//   } catch (e) {
//     console.error('Office insertion error', e);
//     throw e;
//   }
// }

export async function insertHtmlToWord(html) {
  try {
    await Office.onReady();
    await Word.run(async (context) => {
      const body = context.document.body;

      // Clear existing content
      body.clear();

      // Optional: remove leftover formatting (numbering, bullets, etc.)
      const paras = body.paragraphs;
      paras.load("items");
      await context.sync();

      paras.items.forEach((p) => {
        p.clear(); // clear formatting for each paragraph
      });

      // Insert new HTML at start
      body.insertHtml(html, Word.InsertLocation.start);

      await context.sync();
    });
  } catch (e) {
    console.error("Office insertion error", e);
    throw e;
  }
}
