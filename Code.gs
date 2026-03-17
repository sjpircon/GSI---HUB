// ============================================================
// CODE.GS — Web App Entry Point
// Spreadsheet Integration Hub
// ============================================================

function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) ? e.parameter.page : 'home';
  const template = HtmlService.createTemplateFromFile('Index');
  template.page = page;
  return template.evaluate()
    .setTitle('Spreadsheet Integration Hub')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── Include helper for CSS/JS partials ──────────────────────
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
