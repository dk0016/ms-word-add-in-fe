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
  await Office.onReady();
  await Word.run(async (context) => {
    const body = context.document.body;

    // Load all paragraphs
    const paras = body.paragraphs;
    paras.load("items");
    await context.sync();

    // Delete each paragraph individually
    paras.items.forEach((p) => p.delete());

    await context.sync();
  });
}

export async function insertHtmlToWord(html) {
  await Office.onReady();
  await Word.run(async (context) => {
    const body = context.document.body;

    // Ensure the body is empty
    const paras = body.paragraphs;
    paras.load("items");
    await context.sync();

    paras.items.forEach((p) => p.delete());
    await context.sync();

    // Wrap HTML in a div to prevent style inheritance
    const wrappedHtml = `<div style="margin:0; padding:0; list-style:none;">${html}</div>`;
    body.insertHtml(wrappedHtml, Word.InsertLocation.start);

    await context.sync();
  });
}
