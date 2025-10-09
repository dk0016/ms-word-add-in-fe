export async function insertHtmlToWord(html) {
  try {
    await Office.onReady();
    await Word.run(async (context) => {
      const body = context.document.body;
      body.clear();
      // Insert HTML at start
      body.insertHtml(html, Word.InsertLocation.start);
      await context.sync();
    });
  } catch (e) {
    console.error("Office insertion error", e);
    throw e;
  }
}
