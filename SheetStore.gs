function sdsdEnsureSheets_() {
  const ss = SpreadsheetApp.getActive();
  const names = [
    SDSD_CONFIG.sheets.evidencePageSummary,
    SDSD_CONFIG.sheets.evidencePageWeekly,
    SDSD_CONFIG.sheets.evidencePageQuery,
    SDSD_CONFIG.sheets.sbmHistory,
    SDSD_CONFIG.sheets.candidates,
    SDSD_CONFIG.sheets.selectedCases,
    SDSD_CONFIG.sheets.articleMaster
  ];
  names.forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (name !== SDSD_CONFIG.sheets.candidates && name !== SDSD_CONFIG.sheets.selectedCases) {
      try { sh.hideSheet(); } catch (e) {}
    }
  });
  try { sdsdHideInternalSheets_(); } catch (e) {}
}

function sdsdReadObjects_(sheetName) {
  const sh = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sh) return [];
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(String);
  return values.slice(1).filter(r => r.some(v => v !== '')).map(r => {
    const o = {};
    headers.forEach((h, i) => o[h] = r[i]);
    return o;
  });
}

function sdsdWriteCandidates_(rows) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SDSD_CONFIG.sheets.candidates);
  sh.clear();

  const userHeaders = [
    '優先順位','記事タイトル','記事URL','優先度','選定理由','対応方針'
  ];
  const technicalHeaders = [
    'Rank','Normalized URL','TVS','Demand','Opportunity','Urgency','Asset Value',
    'Ownership','Recent Treatment Guard','Weekly Trend','Evidence Confidence',
    'Treatment Risk','External Factor','Priority Candidate','Reason'
  ];
  const headers = userHeaders.concat(technicalHeaders);

  const articleMap = sdsdArticleTitleMap_();
  const displayRows = rows.slice().sort((a,b) => {
    const pa = sdsdPriorityRank_(a.priority);
    const pb = sdsdPriorityRank_(b.priority);
    if (pa !== pb) return pa - pb;
    return Number(b.tvs || 0) - Number(a.tvs || 0);
  });

  sh.getRange(1,1,1,headers.length).setValues([headers]);

  if (displayRows.length) {
    const values = displayRows.map((r,i) => [
      i+1,
      sdsdDisplayTitle_(r.url, articleMap),
      r.url,
      sdsdPriorityJa_(r.priority),
      sdsdReasonJa_(r.reason),
      sdsdActionJa_(r),

      i+1,r.url,r.tvs,r.demand,r.opportunity,r.urgency,r.asset,
      r.ownership,r.guard,r.weeklyTrend,r.evidenceConfidence,
      r.treatmentRisk,r.externalFactor,r.priority,r.reason
    ]);
    sh.getRange(2,1,values.length,headers.length).setValues(values);
  }

  sh.setFrozenRows(1);
  sh.getRange(1,1,1,userHeaders.length).setFontWeight('bold');
  sh.setColumnWidth(1, 90);
  sh.setColumnWidth(2, 360);
  sh.setColumnWidth(3, 320);
  sh.setColumnWidth(4, 110);
  sh.setColumnWidth(5, 460);
  sh.setColumnWidth(6, 170);
  sh.getRange(1,1,Math.max(sh.getLastRow(),1),userHeaders.length).setWrap(true);
  sh.getRange(1,4).setNote(
    '最優先: Doctor精密診断の優先度が特に高い\n' +
    '優先: Doctor精密診断を推奨\n' +
    '要確認: 追加情報を確認して判断\n' +
    '日常改善: SBMで対応\n' +
    '保護: 今は大きく触らない\n' +
    '経過観察: 推移を見る'
  );
  sdsdHideTechnicalColumns_(sh, userHeaders.length + 1, headers.length);
}

