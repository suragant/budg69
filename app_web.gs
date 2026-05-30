// app_web.gs
// Web entry points and HTML includes.

function doGet(e) {
  try {
    const page = (e && e.parameter && e.parameter.page)
      ? String(e.parameter.page).toLowerCase().trim()
      : '';
    const templateName = (page === 'support') ? 'SupportIndex' : 'Index';
    return HtmlService.createTemplateFromFile(templateName)
      .evaluate()
      .setTitle('ระบบสารสนเทศในการบริหารงบประมาณ')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    handleError('doGet', err, { page: e?.parameter?.page });
    try {
      return HtmlService.createTemplateFromFile('Index').evaluate();
    } catch (e) {}
    return HtmlService.createHtmlOutput('Error: ' + (err.message || err.toString()));
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
