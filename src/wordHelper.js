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

export async function clearWordBody() {
  try {
    await Office.onReady();
    await Word.run(async (context) => {
      const body = context.document.body;
      body.clear();

      const paras = body.paragraphs;
      paras.load("items");
      await context.sync();

      paras.items.forEach((p) => p.clear());
      await context.sync();
    });
  } catch (e) {
    console.error("Office clear error", e);
    throw e;
  }
}

export async function insertHtmlToWord(html) {
  try {
    await Office.onReady();
    await Word.run(async (context) => {
      const body = context.document.body;

      body.clear(); // extra safety
      body.insertHtml(html, Word.InsertLocation.start);

      await context.sync();
    });
  } catch (e) {
    console.error("Office insertion error", e);
    throw e;
  }
}
