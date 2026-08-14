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
    '順位','記事URL','診断優先度','選定理由','現在の状態'
  ];
  const technicalHeaders = [
    'Rank','Normalized URL','TVS','Demand','Opportunity','Urgency','Asset Value',
    'Ownership','Recent Treatment Guard','Weekly Trend','Evidence Confidence',
    'Treatment Risk','External Factor','Priority Candidate','Reason'
  ];
  const headers = userHeaders.concat(technicalHeaders);

  sh.getRange(1,1,1,headers.length).setValues([headers]);

  if (rows.length) {
    const values = rows.map((r,i) => {
      const priorityJa = r.priority === 'A1_CANDIDATE' ? '最優先'
        : r.priority === 'A2_CANDIDATE' ? '優先'
        : r.priority === 'WAIT' ? '経過観察'
        : r.priority === 'PROTECTED' ? '保護'
        : r.priority === 'SBM' ? '日常改善'
        : r.priority === 'DOCTOR_REVIEW' || r.priority === 'REVIEW' ? '要確認'
        : String(r.priority || '');

      const statusJa = r.guard !== 'PASS' ? '最近処置済み'
        : r.ownership !== 'DOCTOR_OWNED' ? 'Doctor対象外'
        : r.priority === 'PROTECTED' ? '回復・成長中'
        : r.priority === 'WAIT' ? 'モニター中'
        : r.priority === 'A1_CANDIDATE' || r.priority === 'A2_CANDIDATE' ? '診断候補'
        : '確認済み';

      return [
        i+1,
        r.url,
        priorityJa,
        r.reason,
        statusJa,

        i+1,r.url,r.tvs,r.demand,r.opportunity,r.urgency,r.asset,
        r.ownership,r.guard,r.weeklyTrend,r.evidenceConfidence,
        r.treatmentRisk,r.externalFactor,r.priority,r.reason
      ];
    });
    sh.getRange(2,1,values.length,headers.length).setValues(values);
  }

  sh.setFrozenRows(1);
  sh.getRange(1,1,1,userHeaders.length).setFontWeight('bold');
  sh.autoResizeColumns(1,userHeaders.length);
  sh.setColumnWidth(2, 360);
  sh.setColumnWidth(4, 420);
  sh.setColumnWidth(5, 150);
  sdsdHideTechnicalColumns_(sh, userHeaders.length + 1, headers.length);
}

