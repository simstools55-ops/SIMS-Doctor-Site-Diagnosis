function sdsdExtractDriveFileId_(text) {
  const s = String(text || '').trim();

  let m = s.match(/\/file\/d\/([A-Za-z0-9_-]+)/);
  if (m) return m[1];

  m = s.match(/[?&]id=([A-Za-z0-9_-]+)/);
  if (m) return m[1];

  if (/^[A-Za-z0-9_-]{20,}$/.test(s)) return s;

  return '';
}

function sdsdImportEvidencePackageZip() {
  sdsdProductEnsureSheets_();

  const ui = SpreadsheetApp.getUi();
  const prompt = ui.prompt(
    `SIMS Doctor Site Diagnosis ${SDSD_VERSION}`,
    'Google Drive上のCollector Evidence ZIPのURL、またはファイルIDを入力してください。',
    ui.ButtonSet.OK_CANCEL
  );
  if (prompt.getSelectedButton() !== ui.Button.OK) return;

  const fileId = sdsdExtractDriveFileId_(prompt.getResponseText());
  if (!fileId) {
    ui.alert('DriveファイルURLまたはファイルIDを認識できませんでした。');
    return;
  }

  const file = DriveApp.getFileById(fileId);
  const name = file.getName();
  if (!/\.zip$/i.test(name)) {
    throw new Error(`ZIPファイルではありません: ${name}`);
  }

  const blobs = Utilities.unzip(file.getBlob());
  const fileMap = {};
  blobs.forEach(blob => {
    const n = String(blob.getName() || '').split('/').pop();
    if (n) fileMap[n] = blob;
  });

  const required = [
    {file:'page_summary.csv', sheet:SDSD_CONFIG.sheets.evidencePageSummary},
    {file:'page_weekly.csv', sheet:SDSD_CONFIG.sheets.evidencePageWeekly},
    {file:'page_query_top.csv', sheet:SDSD_CONFIG.sheets.evidencePageQuery}
  ];

  const missing = required.filter(x => !fileMap[x.file]).map(x => x.file);
  if (missing.length) {
    throw new Error(`Evidence ZIPに必要ファイルがありません: ${missing.join(', ')}`);
  }

  const report = [];
  required.forEach(x => {
    const text = fileMap[x.file].getDataAsString('UTF-8').replace(/^\uFEFF/, '');
    const values = Utilities.parseCsv(text);

    if (!values.length) {
      throw new Error(`${x.file} が空です。`);
    }

    const sh = SpreadsheetApp.getActive().getSheetByName(x.sheet);
    sh.clearContents();

    const width = Math.max.apply(null, values.map(r => r.length));
    const normalized = values.map(r => {
      const row = r.slice();
      while (row.length < width) row.push('');
      return row;
    });

    sh.getRange(1,1,normalized.length,width).setValues(normalized);
    report.push({
      file:x.file,
      sheet:x.sheet,
      dataRows:Math.max(normalized.length - 1, 0),
      columns:width
    });
  });

  PropertiesService.getDocumentProperties().setProperty(
    'SDSD_LAST_EVIDENCE_FILE_ID',
    fileId
  );

  const diag = sdsdEvidenceImportIntegrity_();

  ui.alert(
    `Evidence Package取込完了\n\n` +
    `ZIP: ${name}\n` +
    `page_summary: ${report[0].dataRows}行\n` +
    `page_weekly: ${report[1].dataRows}行\n` +
    `page_query_top: ${report[2].dataRows}行\n\n` +
    `Query URL数: ${diag.queryUrlCount}\n` +
    `Query行数: ${diag.queryRows}\n\n` +
    `次に「3. サイト診断を実行」を実行してください。`
  );
}

function sdsdEvidenceImportIntegrity_() {
  const queryMap = sdsdBuildQueryEvidenceMap_();
  const queryRows = sdsdReadObjects_(SDSD_CONFIG.sheets.evidencePageQuery);

  return {
    queryRows: queryRows.length,
    queryUrlCount: Object.keys(queryMap).length
  };
}
