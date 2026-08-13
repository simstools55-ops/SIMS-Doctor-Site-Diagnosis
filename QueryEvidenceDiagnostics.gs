function sdsdDiagnoseQueryEvidenceInput() {
  const ss = SpreadsheetApp.getActive();
  const sourceName = SDSD_CONFIG.sheets.evidencePageQuery;
  const sh = ss.getSheetByName(sourceName);
  if (!sh) throw new Error(`Query Evidenceシートが見つかりません: ${sourceName}`);

  const values = sh.getDataRange().getValues();
  const outName = 'Query Evidence Diagnostics';
  let out = ss.getSheetByName(outName);
  if (!out) out = ss.insertSheet(outName);
  out.clear();

  const report = [];
  report.push(['Check','Value']);

  if (!values.length) {
    report.push(['Source Sheet', sourceName]);
    report.push(['Raw Rows', 0]);
    out.getRange(1,1,report.length,2).setValues(report);
    ss.setActiveSheet(out);
    SpreadsheetApp.getUi().alert('Query Evidence Diagnostics完了\n元データ行数: 0');
    return;
  }

  const rawHeaders = values[0].map(v => String(v || ''));
  const normalizedHeaders = rawHeaders.map(h =>
    String(h || '').replace(/^\uFEFF/, '').trim().toLowerCase()
  );

  report.push(['Source Sheet', sourceName]);
  report.push(['Raw Rows', Math.max(values.length - 1, 0)]);
  report.push(['Raw Headers', rawHeaders.join(' | ')]);
  report.push(['Normalized Headers', normalizedHeaders.join(' | ')]);

  const objects = sdsdReadObjects_(sourceName);
  report.push(['Objects Read', objects.length]);

  let pageValues = 0;
  let queryValues = 0;
  let parsedRows = 0;
  const normalizedUrls = {};
  const queryMap = {};

  objects.forEach(r => {
    const page = sdsdObjectValue_(r, ['page','key','url','URL','記事URL']);
    const query = sdsdObjectValue_(r, ['query','クエリ','検索クエリ']);

    if (String(page || '').trim()) pageValues++;
    if (String(query || '').trim()) queryValues++;

    const url = sdsdNormalizeUrl_(page);
    if (url && String(query || '').trim()) {
      parsedRows++;
      normalizedUrls[url] = true;
      if (!queryMap[url]) queryMap[url] = 0;
      queryMap[url]++;
    }
  });

  report.push(['Rows with Page', pageValues]);
  report.push(['Rows with Query', queryValues]);
  report.push(['Parsed Page+Query Rows', parsedRows]);
  report.push(['Normalized URL Count', Object.keys(normalizedUrls).length]);

  const selected = ss.getSheetByName(SDSD_CONFIG.sheets.selectedCases);
  let selectedCount = 0;
  let matched = 0;
  let unmatched = [];

  if (selected) {
    const svals = selected.getDataRange().getValues();
    if (svals.length > 1) {
      const sheaders = svals[0].map(String);
      const idx = {};
      sheaders.forEach((h,i) => idx[h] = i);
      const urlIdx = idx['URL'];

      if (urlIdx != null) {
        svals.slice(1).forEach(r => {
          const u = sdsdNormalizeUrl_(r[urlIdx]);
          if (!u) return;
          selectedCount++;
          if (queryMap[u] > 0) matched++;
          else unmatched.push(u);
        });
      }
    }
  }

  report.push(['Selected Cases', selectedCount]);
  report.push(['Selected Cases Matched to Query Evidence', matched]);
  report.push(['Selected Cases Unmatched', unmatched.length]);

  out.getRange(1,1,report.length,2).setValues(report);

  const start = report.length + 3;
  out.getRange(start,1,1,3).setValues([['Unmatched URL','Normalized URL','Query Rows']]);
  if (unmatched.length) {
    const rows = unmatched.map(u => [u,u,queryMap[u] || 0]);
    out.getRange(start+1,1,rows.length,3).setValues(rows);
  }

  const sampleStart = start + Math.max(unmatched.length,1) + 3;
  out.getRange(sampleStart,1,1,4).setValues([['Sample URL','Query Count','First Query','First Impressions']]);

  const built = sdsdBuildQueryEvidenceMap_();
  const samples = Object.keys(built).slice(0,10).map(u => {
    const arr = built[u] || [];
    return [
      u,
      arr.length,
      arr[0] ? arr[0].query : '',
      arr[0] ? arr[0].impressions : ''
    ];
  });
  if (samples.length) {
    out.getRange(sampleStart+1,1,samples.length,4).setValues(samples);
  }

  out.setFrozenRows(1);
  out.autoResizeColumns(1,4);
  ss.setActiveSheet(out);

  SpreadsheetApp.getUi().alert(
    `Query Evidence Diagnostics完了\n\n` +
    `元データ行数: ${Math.max(values.length-1,0)}\n` +
    `Page認識: ${pageValues}\n` +
    `Query認識: ${queryValues}\n` +
    `URL正規化後: ${Object.keys(normalizedUrls).length}\n` +
    `選定18件一致: ${matched}/${selectedCount}`
  );
}
