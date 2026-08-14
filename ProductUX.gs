const SDSD_INTERNAL_SHEET_NAMES_ = Object.freeze([
  '_SDSD_PAGE_SUMMARY',
  '_SDSD_PAGE_WEEKLY',
  '_SDSD_PAGE_QUERY_TOP',
  '_SDSD_SBM_HISTORY',
  '_SDSD_ARTICLE_MASTER',
  'Weekly Trend Validation',
  'Priority Validation',
  'Query Evidence Diagnostics',
  'Site Diagnosis Candidates',
  'Selected Treatment Cases',
  'シート1',
  'Sheet1'
]);


function sdsdProductEnsureSheets_() {
  const ss = SpreadsheetApp.getActive();

  const legacyPairs = [
    ['Site Diagnosis Candidates', SDSD_CONFIG.sheets.candidates],
    ['Selected Treatment Cases', SDSD_CONFIG.sheets.selectedCases]
  ];
  legacyPairs.forEach(pair => {
    const legacy = ss.getSheetByName(pair[0]);
    const current = ss.getSheetByName(pair[1]);

    if (legacy && !current) {
      try {
        legacy.setName(pair[1]);
      } catch (e) {}
      return;
    }

    if (legacy && current) {
      try {
        if (ss.getActiveSheet().getSheetId() === legacy.getSheetId()) {
          ss.setActiveSheet(current);
        }
        legacy.hideSheet();
      } catch (e) {}
    }
  });

  const names = [
    SDSD_CONFIG.sheets.evidencePageSummary,
    SDSD_CONFIG.sheets.evidencePageWeekly,
    SDSD_CONFIG.sheets.evidencePageQuery,
    SDSD_CONFIG.sheets.sbmHistory,
    SDSD_CONFIG.sheets.summary,
    SDSD_CONFIG.sheets.candidates,
    SDSD_CONFIG.sheets.selectedCases,
    SDSD_CONFIG.sheets.articleMaster
  ];

  names.forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
  });

  sdsdHideUnusedDefaultSheet_();
  sdsdHideInternalSheets_();
  if (!ss.getSheetByName(SDSD_CONFIG.sheets.opportunities)) ss.insertSheet(SDSD_CONFIG.sheets.opportunities);
  if (!ss.getSheetByName(SDSD_CONFIG.sheets.opportunityCases)) ss.insertSheet(SDSD_CONFIG.sheets.opportunityCases);
  if (!ss.getSheetByName(SDSD_CONFIG.sheets.siteWideResult)) ss.insertSheet(SDSD_CONFIG.sheets.siteWideResult);
  if (!ss.getSheetByName(SDSD_CONFIG.sheets.treatmentPlan)) ss.insertSheet(SDSD_CONFIG.sheets.treatmentPlan);
  if (!ss.getSheetByName(SDSD_CONFIG.sheets.siteWideResultImport)) ss.insertSheet(SDSD_CONFIG.sheets.siteWideResultImport);
}

function sdsdHideUnusedDefaultSheet_() {
  const ss = SpreadsheetApp.getActive();

  ['シート1', 'Sheet1'].forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;

    const values = sh.getDataRange().getDisplayValues();
    const hasContent = values.some(row => row.some(v => String(v).trim() !== ''));
    if (hasContent) return;

    try {
      const preferred =
        ss.getSheetByName(SDSD_CONFIG.sheets.summary) ||
        ss.getSheetByName(SDSD_CONFIG.sheets.candidates) ||
        ss.getSheetByName(SDSD_CONFIG.sheets.selectedCases);

      if (preferred && ss.getActiveSheet().getSheetId() === sh.getSheetId()) {
        ss.setActiveSheet(preferred);
        SpreadsheetApp.flush();
      }

      sh.hideSheet();
    } catch (e) {}
  });
}

function sdsdHideTechnicalColumns_(sheet, firstTechnicalCol, totalCols) {
  if (!sheet || totalCols < firstTechnicalCol) return;
  try {
    sheet.hideColumns(firstTechnicalCol, totalCols - firstTechnicalCol + 1);
  } catch (e) {}
}


function sdsdArticleTitleMap_() {
  try {
    const map = sdsdBuildArticleMasterMap_();
    return map || {};
  } catch (e) {
    return {};
  }
}

function sdsdDisplayTitle_(url, articleMap) {
  const normalized = sdsdNormalizeUrl_(url);
  const master = articleMap && articleMap[normalized] ? articleMap[normalized] : null;
  return master && master.title ? String(master.title) : '（タイトル未取得）';
}

function sdsdPriorityJa_(priority) {
  const p = String(priority || '');
  if (p === 'A1_CANDIDATE') return '最優先';
  if (p === 'A2_CANDIDATE') return '優先';
  if (p === 'DOCTOR_REVIEW' || p === 'REVIEW' || p === 'B_CANDIDATE' || p === 'CANDIDATE') return '要確認';
  if (p === 'SBM') return '日常改善';
  if (p === 'PROTECTED') return '保護';
  if (p === 'WAIT') return '経過観察';
  return p || '要確認';
}

function sdsdPriorityRank_(priority) {
  const p = String(priority || '');
  if (p === 'A1_CANDIDATE') return 10;
  if (p === 'A2_CANDIDATE') return 20;
  if (p === 'DOCTOR_REVIEW') return 30;
  if (p === 'REVIEW' || p === 'B_CANDIDATE' || p === 'CANDIDATE') return 40;
  if (p === 'SBM') return 50;
  if (p === 'PROTECTED') return 60;
  if (p === 'WAIT') return 70;
  return 80;
}

function sdsdActionJa_(row) {
  const p = String(row.priority || '');
  if (p === 'A1_CANDIDATE' || p === 'A2_CANDIDATE') return 'Doctorで精密診断';
  if (p === 'DOCTOR_REVIEW' || p === 'REVIEW' || p === 'B_CANDIDATE' || p === 'CANDIDATE') return '追加確認';
  if (p === 'SBM') return 'SBMで日常改善';
  if (p === 'PROTECTED') return '今は大きく触らない';
  if (p === 'WAIT') return '経過観察';
  return '確認';
}

function sdsdReasonJa_(reason) {
  let s = String(reason || '');
  const replacements = [
    ['週次:SEVERE_DECLINE', '直近で大きく悪化'],
    ['週次:TRAFFIC_DECLINE', '直近の検索流入が低下'],
    ['週次:RANKING_DECLINE', '直近の検索順位が低下'],
    ['週次:VOLATILE', '直近の推移が不安定'],
    ['週次:GROWTH', '回復・成長傾向'],
    ['週次:STABLE', '直近の推移は安定'],
    ['外部要因:PLATFORM_OR_OS_CHANGE', 'プラットフォーム・OS変更の影響可能性'],
    ['PLATFORM_OR_OS_CHANGE', 'プラットフォーム・OS変更の影響可能性'],
    ['高Risk', '変更リスクが高い'],
    ['追加Evidence確認', '追加データの確認が必要'],
    ['TRAFFIC_DECLINE', '検索流入が低下'],
    ['RANKING_DECLINE', '検索順位が低下'],
    ['SEVERE_DECLINE', '直近で大きく悪化'],
    ['VOLATILE', '直近の推移が不安定'],
    ['GROWTH', '回復・成長傾向'],
    ['STABLE', '直近の推移は安定'],
    ['主病変がCTR/即効性改善', 'CTR（クリック率）の改善余地が大きい'],
    ['DOCTOR_OWNED', 'Doctor精密診断向き'],
    ['SBM_OWNED', 'SBMの日常改善向き']
  ];
  replacements.forEach(pair => {
    s = s.split(pair[0]).join(pair[1]);
  });
  s = s.replace(/\s*\/\s*/g, '／');
  s = s.replace(/\|/g, '・');
  return s;
}

function sdsdSiteMeaning_(row) {
  const trend = String(row.weeklyTrend || '');
  const ext = String(row.externalFactor || '');
  if (trend === 'SEVERE_DECLINE') return '大幅な悪化が出ている記事群の代表';
  if (trend === 'TRAFFIC_DECLINE') return '検索流入低下が出ている記事群の代表';
  if (trend === 'RANKING_DECLINE') return '検索順位低下が出ている記事群の代表';
  if (trend === 'VOLATILE') return '検索推移が不安定な記事群の代表';
  if (ext.indexOf('PLATFORM_OR_OS_CHANGE') >= 0) return '外部要因の影響を確認する代表';
  return '優先的に原因確認する代表記事';
}

function sdsdWriteSiteSummary_(rows, result) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SDSD_CONFIG.sheets.summary);
  if (!sh) sh = ss.insertSheet(SDSD_CONFIG.sheets.summary);
  sh.clear();

  const total = result.total || rows.length;
  const a1 = rows.filter(r => r.priority === 'A1_CANDIDATE').length;
  const a2 = rows.filter(r => r.priority === 'A2_CANDIDATE').length;
  const priorityCount = a1 + a2;
  const review = rows.filter(r =>
    r.priority === 'DOCTOR_REVIEW' || r.priority === 'REVIEW' ||
    r.priority === 'B_CANDIDATE' || r.priority === 'CANDIDATE'
  ).length;
  const sbm = rows.filter(r => r.priority === 'SBM').length;
  const protectedCount = rows.filter(r => r.priority === 'PROTECTED').length;
  const wait = rows.filter(r => r.priority === 'WAIT').length;

  const severe = rows.filter(r => r.weeklyTrend === 'SEVERE_DECLINE').length;
  const traffic = rows.filter(r => r.weeklyTrend === 'TRAFFIC_DECLINE').length;
  const ranking = rows.filter(r => r.weeklyTrend === 'RANKING_DECLINE').length;
  const volatile = rows.filter(r => r.weeklyTrend === 'VOLATILE').length;
  const growth = rows.filter(r => r.weeklyTrend === 'GROWTH').length;
  const external = rows.filter(r => String(r.externalFactor || '') !== '').length;

  let selectedCount = 0;
  let selectedUrls = {};
  try {
    const selected = ss.getSheetByName(SDSD_CONFIG.sheets.selectedCases);
    if (selected && selected.getLastRow() >= 2) {
      const vals = selected.getDataRange().getValues();
      const hdr = vals[0].map(String);
      const urlCol = hdr.indexOf('URL') >= 0 ? hdr.indexOf('URL') : hdr.indexOf('記事URL');
      if (urlCol >= 0) {
        vals.slice(1).forEach(r => {
          const u = String(r[urlCol] || '').trim();
          if (u) selectedUrls[u] = true;
        });
        selectedCount = Object.keys(selectedUrls).length;
      }
    }
  } catch (e) {}

  const priorityRows = rows.filter(r =>
    r.priority === 'A1_CANDIDATE' || r.priority === 'A2_CANDIDATE'
  );
  const notSelected = priorityRows.filter(r => !selectedUrls[String(r.url || '').trim()]);
  const notSelectedCount = Math.max(priorityCount - selectedCount, 0);

  let notSelectedReason = '';
  if (notSelectedCount > 0) {
    const guardCount = notSelected.filter(r =>
      String(r.guard || '').toUpperCase() !== '' &&
      String(r.guard || '').toUpperCase() !== 'PASS'
    ).length;
    const riskCount = notSelected.filter(r =>
      String(r.treatmentRisk || '').toUpperCase() === 'HIGH'
    ).length;

    if (guardCount > 0) {
      notSelectedReason = `最終確認・保護条件により今回は見送り ${guardCount}記事`;
    } else if (riskCount > 0) {
      notSelectedReason = `処置リスクが高いため今回は見送り ${riskCount}記事`;
    } else {
      notSelectedReason = `Treatment Batchの選定上限・優先順位により今回は見送り ${notSelectedCount}記事`;
    }
  }

  let overall = 'サイト全体を一律に修正する状態ではありません。';
  if (priorityCount > 0) {
    overall += ` 影響の大きい${priorityCount}記事を優先してDoctorで原因を確認します。`;
  }
  if (growth + protectedCount > 0) {
    overall += ' 回復・成長中の記事は保護し、不要な修正を避けます。';
  }

  const values = [
    ['サイト診断サマリー', ''],
    ['診断対象', `${total}記事`],
    ['Doctor精密診断の優先候補', `${priorityCount}記事（最優先 ${a1} / 優先 ${a2}）`],
    ['今回Doctorへ送る記事', selectedCount ? `${selectedCount}記事` : 'Treatment Batch作成前'],
    ['今回送らない優先候補', selectedCount ? `${notSelectedCount}記事` : 'Treatment Batch作成前'],
    ['今回送らない理由', selectedCount ? (notSelectedReason || 'なし') : 'Treatment Batch作成前'],
    ['追加確認が必要', `${review}記事`],
    ['SBMの日常改善向き', `${sbm}記事`],
    ['回復・成長中などの保護対象', `${protectedCount}記事`],
    ['最近処置済み・経過観察', `${wait}記事`],
    ['', ''],
    ['サイト全体の所見', overall],
    ['', ''],
    ['サイト全体の主な症状', ''],
    ['直近で大きく悪化', `${severe}記事`],
    ['検索流入低下', `${traffic}記事`],
    ['検索順位低下', `${ranking}記事`],
    ['推移が不安定', `${volatile}記事`],
    ['回復・成長傾向', `${growth}記事`],
    ['OS・サービス変更など外部要因の可能性', `${external}記事`],
    ['', ''],
    ['この診断の読み方', 'サイト全体を一律に修正するのではなく、優先候補をDoctorで精密診断し、共通原因と処置の必要性を確認します。'],
    ['注意', 'この段階では改善率を断定しません。実際の処置はDoctorの精密診断後に決めます。'],
    ['', ''],
    ['優先度の見方', ''],
    ['最優先', 'Doctorで詳しく診断する優先度が特に高い'],
    ['優先', 'Doctorでの精密診断を推奨'],
    ['要確認', '追加情報を確認してから判断'],
    ['日常改善', 'DoctorではなくSBMの日常改善で対応'],
    ['保護', '回復・成長中などのため今は大きく触らない'],
    ['経過観察', '最近の処置後などのため、しばらく推移を見る'],
    ['', ''],
    ['Doctor Case Package', '未生成']
  ];

  sh.getRange(1,1,values.length,2).setValues(values);
  sh.getRange('A1:B1').merge();
  sh.getRange('A1').setValue('サイト診断サマリー');
  sh.getRange('A1').setFontWeight('bold').setFontSize(14);
  sh.getRange('A12').setFontWeight('bold');
  sh.getRange('A14').setFontWeight('bold');
  sh.getRange('A26').setFontWeight('bold');
  sh.getRange(1,1,values.length,1).setFontWeight('bold');
  sh.setColumnWidth(1, 250);
  sh.setColumnWidth(2, 650);
  sh.getRange(1,1,values.length,2).setWrap(true);
  sh.setFrozenRows(1);

  try { ss.setActiveSheet(sh); } catch (e) {}
}

function sdsdUpdateSummaryAfterBatch_(batch, guard) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SDSD_CONFIG.sheets.summary);
  if (!sh) return;
  const values = sh.getRange(1,1,sh.getLastRow(),2).getValues();
  for (let i=0; i<values.length; i++) {
    if (String(values[i][0]) === '今回Doctorへ送る記事') {
      sh.getRange(i+1,2).setValue(
        `${batch.selectedCount}記事` +
        (guard.blocked ? `（最終確認で保留 ${guard.blocked}）` : '')
      );
      return;
    }
  }
}

function sdsdUpdateSummaryAfterPackage_(caseCount, fileUrl) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SDSD_CONFIG.sheets.summary);
  if (!sh) return;
  const values = sh.getRange(1,1,sh.getLastRow(),2).getValues();
  for (let i=0; i<values.length; i++) {
    if (String(values[i][0]) === 'Doctor Case Package') {
      sh.getRange(i+1,2).setValue(`${caseCount}件生成済み`);
      if (fileUrl) sh.getRange(i+1,2).setNote(fileUrl);
      return;
    }
  }
}


function sdsdHeaderIndexMap_(headers) {
  const map = {};
  headers.forEach((h,i) => map[String(h)] = i);
  return map;
}

function sdsdCandidateRowsFromSheet_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SDSD_CONFIG.sheets.candidates);
  if (!sh || sh.getLastRow() < 2) return [];

  const values = sh.getDataRange().getValues();
  const headers = values[0].map(String);
  const idx = sdsdHeaderIndexMap_(headers);

  if (idx['Normalized URL'] == null || idx['Priority Candidate'] == null) {
    return [];
  }

  return values.slice(1)
    .filter(r => String(r[idx['Normalized URL']] || ''))
    .map(r => ({
      url: String(r[idx['Normalized URL']] || ''),
      tvs: Number(r[idx['TVS']] || 0),
      demand: Number(r[idx['Demand']] || 0),
      opportunity: Number(r[idx['Opportunity']] || 0),
      urgency: Number(r[idx['Urgency']] || 0),
      asset: Number(r[idx['Asset Value']] || 0),
      ownership: String(r[idx['Ownership']] || ''),
      guard: String(r[idx['Recent Treatment Guard']] || ''),
      weeklyTrend: String(r[idx['Weekly Trend']] || ''),
      evidenceConfidence: String(r[idx['Evidence Confidence']] || ''),
      treatmentRisk: String(r[idx['Treatment Risk']] || ''),
      externalFactor: String(r[idx['External Factor']] || ''),
      priority: String(r[idx['Priority Candidate']] || ''),
      reason: String(r[idx['Reason']] || '')
    }));
}

function sdsdRefreshSiteSummaryFromCandidates_() {
  const rows = sdsdCandidateRowsFromSheet_();
  if (!rows.length) return false;

  const result = {
    total: rows.length,
    priorityCandidates: rows.filter(r =>
      r.priority === 'A1_CANDIDATE' || r.priority === 'A2_CANDIDATE'
    ).length,
    wait: rows.filter(r => r.priority === 'WAIT').length,
    protected: rows.filter(r => r.priority === 'PROTECTED').length,
    sbm: rows.filter(r => r.priority === 'SBM').length,
    review: rows.filter(r =>
      r.priority === 'REVIEW' || r.priority === 'DOCTOR_REVIEW' ||
      r.priority === 'B_CANDIDATE' || r.priority === 'CANDIDATE'
    ).length
  };

  sdsdWriteSiteSummary_(rows, result);
  return true;
}

function sdsdRefreshCandidatesView_() {
  const rows = sdsdCandidateRowsFromSheet_();
  if (!rows.length) return false;
  sdsdWriteCandidates_(rows);
  return true;
}

function sdsdRefreshSelectedCasesView_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SDSD_CONFIG.sheets.selectedCases);
  if (!sh || sh.getLastRow() < 2) return;

  const values = sh.getDataRange().getValues();
  const headers = values[0].map(String);
  const idx = sdsdHeaderIndexMap_(headers);

  if (idx['URL'] == null || idx['Referral JSON'] == null) return;

  const technicalHeaders = [
    'Batch Order','Site Priority','URL','TVS','Weekly Trend','Evidence Confidence',
    'Treatment Risk','External Factor','Ownership','Recent Treatment Guard',
    'Top Queries','Selection Reason','Referral Status','Referral JSON',
    'ArticleID','Article Title','Main Query','Article Fetch Status',
    'Case Package Status','Article Cache Key','Query Evidence Count'
  ].filter(h => idx[h] != null);

  const articleMap = sdsdArticleTitleMap_();
  const rows = values.slice(1).filter(r => String(r[idx['URL']] || ''));
  if (!rows.length) return;

  const userHeaders = [
    'No.','記事タイトル','記事URL','優先度','選定理由','サイト全体での意味'
  ];
  const newHeaders = userHeaders.concat(technicalHeaders);

  const newValues = rows.map((r,i) => {
    const url = String(r[idx['URL']] || '');
    const sitePriority = String(r[idx['Site Priority']] || '');
    const priorityJa = sitePriority === 'A1' ? '最優先' : sitePriority === 'A2' ? '優先' : sitePriority || '要確認';
    const reason = String(r[idx['Selection Reason']] || '');
    const displayRow = {
      priority: sitePriority === 'A1' ? 'A1_CANDIDATE' : sitePriority === 'A2' ? 'A2_CANDIDATE' : sitePriority,
      weeklyTrend: String(r[idx['Weekly Trend']] || ''),
      externalFactor: String(r[idx['External Factor']] || ''),
      reason: reason
    };

    const titleFromSheet = idx['Article Title'] != null ? String(r[idx['Article Title']] || '') : '';
    const visible = [
      i+1,
      titleFromSheet || sdsdDisplayTitle_(url, articleMap),
      url,
      priorityJa,
      sdsdReasonJa_(reason),
      sdsdSiteMeaning_(displayRow)
    ];
    const tech = technicalHeaders.map(h => r[idx[h]]);
    return visible.concat(tech);
  });

  sh.clear();
  sh.getRange(1,1,1,newHeaders.length).setValues([newHeaders]);
  sh.getRange(2,1,newValues.length,newHeaders.length).setValues(newValues);
  sh.setFrozenRows(1);
  sh.getRange(1,1,1,userHeaders.length).setFontWeight('bold');
  sh.setColumnWidth(1,70);
  sh.setColumnWidth(2,360);
  sh.setColumnWidth(3,320);
  sh.setColumnWidth(4,110);
  sh.setColumnWidth(5,460);
  sh.setColumnWidth(6,260);
  sh.getRange(1,1,Math.max(sh.getLastRow(),1),userHeaders.length).setWrap(true);
  sdsdHideTechnicalColumns_(sh, userHeaders.length + 1, newHeaders.length);
}

function sdsdArticleMasterCoverageForSelected_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SDSD_CONFIG.sheets.selectedCases);
  if (!sh || sh.getLastRow() < 2) return {selected:0, matched:0, withArticleId:0};

  const values = sh.getDataRange().getValues();
  const headers = values[0].map(String);
  const idx = sdsdHeaderIndexMap_(headers);
  if (idx['URL'] == null) return {selected:0, matched:0, withArticleId:0};

  const articleMap = sdsdBuildArticleMasterMap_();
  let selected = 0;
  let matched = 0;
  let withArticleId = 0;

  values.slice(1).forEach(r => {
    const raw = String(r[idx['URL']] || '');
    if (!raw) return;
    selected++;
    const master = articleMap[sdsdNormalizeUrl_(raw)] || null;
    if (master) {
      matched++;
      if (String(master.articleId || '')) withArticleId++;
    }
  });

  return {selected:selected, matched:matched, withArticleId:withArticleId};
}

function sdsdProgress_(step, total, message) {
  try {
    SpreadsheetApp.getActive().toast(
      `Step ${step}/${total}  ${message}`,
      'SIMS Doctor Site Diagnosis',
      8
    );
    SpreadsheetApp.flush();
  } catch (e) {}
}

function sdsdHideInternalSheets_() {
  const ss = SpreadsheetApp.getActive();
  const current = ss.getActiveSheet();

  if (current && SDSD_INTERNAL_SHEET_NAMES_.indexOf(current.getName()) >= 0) {
    const visible =
      ss.getSheetByName(SDSD_CONFIG.sheets.summary) ||
      ss.getSheetByName(SDSD_CONFIG.sheets.candidates) ||
      ss.getSheetByName(SDSD_CONFIG.sheets.selectedCases);
    if (visible && !visible.isSheetHidden()) {
      ss.setActiveSheet(visible);
    }
  }

  SDSD_INTERNAL_SHEET_NAMES_.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    try { sh.hideSheet(); } catch (e) {}
  });
}

function sdsdHideInternalSheets() {
  sdsdHideInternalSheets_();
  SpreadsheetApp.getUi().alert(
    '内部処理用シートを非表示にしました。\n通常利用では「診断候補」と「選定案件」だけ確認すれば大丈夫です。'
  );
}

function sdsdShowInternalSheets() {
  const ss = SpreadsheetApp.getActive();
  let count = 0;
  SDSD_INTERNAL_SHEET_NAMES_.forEach(name => {
    if (name === 'シート1' || name === 'Sheet1') return;
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    try {
      sh.showSheet();
      count++;
    } catch (e) {}
  });
  SpreadsheetApp.getUi().alert(
    `保守確認用として内部シートを表示しました。\n表示: ${count}シート\n\n確認後は「内部シートを隠す」を実行してください。`
  );
}

function sdsdRunProductDiagnosis() {
  sdsdHideInternalSheets_();
  try {
    const result = sdsdRunAnalysis({silent:true});
    sdsdHideInternalSheets_();

    SpreadsheetApp.getUi().alert(
      `サイト診断が完了しました。\n\n` +
      `対象記事: ${result.total}件\n` +
      `Doctor精密診断の優先候補: ${result.priorityCandidates}件\n` +
      `回復・成長中のため保護: ${result.protected}件\n` +
      `SBMの日常改善対象: ${result.sbm}件\n\n` +
      `「サイト診断サマリー」でブログ全体の状態を確認してください。\n` +
      `詳しい記事一覧は「4. 診断候補を見る」で確認できます。`
    );
  } catch (e) {
    SpreadsheetApp.getUi().alert(
      `サイト診断を完了できませんでした。\n\n${e.message || e}`
    );
    throw e;
  }
}

function sdsdCreateProductTreatmentBatch() {
  try {
    sdsdProgress_(1, 2, '治療候補を選定しています');
    const batch = sdsdBuildTreatmentBatch({silent:true, noActivate:true});

    sdsdProgress_(2, 2, '最近の処置履歴を最終確認しています');
    const guard = sdsdRunFinalGuard({silent:true});
    sdsdUpdateSummaryAfterBatch_(batch, guard);
    sdsdHideInternalSheets_();

    SpreadsheetApp.getUi().alert(
      `Treatment Batchを作成しました。\n\n` +
      `診断対象: ${batch.articleCount}記事\n` +
      `Doctor精密診断の適格候補: ${batch.eligibleCount}件\n` +
      `今回Doctorへ送る記事: ${batch.selectedCount}件\n` +
      `最終確認で保留: ${guard.blocked}件\n\n` +
      `次に「6. 選定案件を見る」で内容を確認してください。`
    );
  } catch (e) {
    SpreadsheetApp.getUi().alert(
      `Treatment Batchを作成できませんでした。\n\n${e.message || e}`
    );
    throw e;
  }
}

function sdsdCreateProductCasePackage() {
  try {
    sdsdRefreshSelectedCasesView_();

    const coverage = sdsdArticleMasterCoverageForSelected_();
    if (coverage.selected === 0) {
      SpreadsheetApp.getUi().alert(
        'Doctor Case Packageを生成できません。\n\n今回の診断対象がありません。'
      );
      return;
    }

    if (coverage.withArticleId < coverage.selected) {
      SpreadsheetApp.getUi().alert(
        `Doctor Case Packageを生成する前に、今回のサイトの「記事管理」データが必要です。\n\n` +
        `今回の診断対象: ${coverage.selected}件\n` +
        `Article MasterでURL一致: ${coverage.matched}件\n` +
        `ArticleIDまで確認できた記事: ${coverage.withArticleId}件\n\n` +
        `「データ準備 → Article Masterの取込案内」から、` +
        `今回診断しているブログのSBM「記事管理」CSVを取り込んでください。\n\n` +
        `取り込んだ後は1～5をやり直さず、もう一度「7. Doctor Case Packageを生成」を実行できます。`
      );
      return;
    }

    sdsdProgress_(1, 3, 'Doctor Case Packageを準備しています');
    const enrichment = sdsdEnrichSelectedCases({
      silent: true,
      maxPerRun: 3
    });

    sdsdRefreshSelectedCasesView_();

    if (enrichment.review > 0) {
      SpreadsheetApp.getUi().alert(
        `Doctor Case Packageの準備中に要確認記事が見つかりました。\n\n` +
        `準備完了: ${enrichment.ready}/${enrichment.total}件\n` +
        `要確認: ${enrichment.review}件\n` +
        `未処理: ${enrichment.pending}件\n\n` +
        `「6. 選定案件を見る」で要確認案件を確認してください。`
      );
      return;
    }

    if (!enrichment.complete) {
      SpreadsheetApp.getUi().alert(
        `Doctor Case Packageを準備しています。\n\n` +
        `準備完了: ${enrichment.ready}/${enrichment.total}件\n` +
        `未処理: ${enrichment.pending}件\n\n` +
        `今回の処理結果は保存しました。\n` +
        `もう一度「7. Doctor Case Packageを生成」を実行すると、未処理の記事から続けます。`
      );
      return;
    }

    sdsdProgress_(2, 3, '全記事の準備完了。ZIPを生成しています');
    const exported = sdsdExportDoctorCasePackageZip({silent:true});

    sdsdUpdateSummaryAfterPackage_(exported.caseCount, exported.fileUrl);
    sdsdRefreshSelectedCasesView_();
    sdsdProgress_(3, 3, 'Doctor Case Packageを保存しました');

    SpreadsheetApp.getUi().alert(
      `Doctor Case Packageの生成が完了しました。\n\n` +
      `案件数: ${exported.caseCount}件\n` +
      `本文再取得: ${exported.refetched}件\n\n` +
      `Google DriveにZIPを保存しました。\n${exported.fileUrl}\n\n` +
      `このZIPをSIMS Doctorへ渡してください。`
    );
  } catch (e) {
    SpreadsheetApp.getUi().alert(
      `Doctor Case Packageを生成できませんでした。\n\n${e.message || e}`
    );
    throw e;
  }
}
