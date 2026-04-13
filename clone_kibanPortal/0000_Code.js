const APP_TITLE = 'SOL推進基盤GrPortal';

function doGet(e) {
  const p = e?.parameter || {};
  const action = (p.action || '').toLowerCase();

  if (action) {
    const out = handleApiGet_(e) || unknownActionError_(action);
    return json_(out);
  }

  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle(APP_TITLE)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  return json_(handleApiPost_(e));
}

function handleApiGet_(e) {
  const p = e.parameter || {};
  const action = (p.action || '').toLowerCase();

  switch (action) {
    case 'case':
      return apiGetCaseByKey_(p.key || '');
    default:
      return unknownActionError_(action);
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON
  );
}

function handleApiPost_(e) {
  const body = parsePostBody_(e);

  switch (body.action) {
    case 'updateSv':
      updateSv_(body);
      return { ok: true };
    default:
      return unknownActionError_(body.action || '');
  }
}

function parsePostBody_(e) {
  const raw = e?.postData?.contents || '{}';
  try {
    return JSON.parse(raw);
  } catch (error) {
    return {};
  }
}

function unknownActionError_(action) {
  return { ok: false, error: `unknown action: ${action}` };
}

/** dead-simple health check */
function api_ping() {
  return {
    ok: true,
    ts: new Date().toISOString(),
    viewSsid: WEBVIEW.VIEW_SSID,
    sheetSV: WEBVIEW.SHEET_SV,
    sheetCL: WEBVIEW.SHEET_CL,
    config: {
      SV: VIEW_CONFIG.SV,
      CL: VIEW_CONFIG.CL,
    },
  };
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
