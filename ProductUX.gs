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

  const values = [
    ['サイト診断サマリー', ''],
    ['診断対象', `${total}記事`],
    ['Doctor精密診断の優先候補', `${a1 + a2}記事（最優先 ${a1} / 優先 ${a2}）`],
    ['追加確認が必要', `${review}記事`],
    ['SBMの日常改善向き', `${sbm}記事`],
    ['回復・成長中のため保護', `${protectedCount}記事`],
    ['最近処置済み・経過観察', `${wait}記事`],
    ['', ''],
    ['サイト全体の主な症状', ''],
    ['大幅な悪化', `${severe}記事`],
    ['検索流入低下', `${traffic}記事`],
    ['検索順位低下', `${ranking}記事`],
    ['推移が不安定', `${volatile}記事`],
    ['回復・成長傾向', `${growth}記事`],
    ['外部要因の影響可能性', `${external}記事`],
    ['', ''],
    ['今回の読み方', 'サイト全体を一律に修正するのではなく、優先候補をDoctorで精密診断し、共通原因と処置の必要性を確認します。'],
    ['注意', 'この段階では「修正すれば何％改善する」とは断定しません。修正効果はDoctor診断後に判断します。'],
    ['', ''],
    ['優先度の見方', ''],
    ['最優先', 'Doctorで詳しく診断する優先度が特に高い'],
    ['優先', 'Doctorでの精密診断を推奨'],
    ['要確認', '追加情報を確認してから判断'],
    ['日常改善', 'DoctorではなくSBMの日常改善で対応'],
    ['保護', '回復・成長中などのため今は大きく触らない'],
    ['経過観察', '最近の処置後などのため、しばらく推移を見る'],
    ['', ''],
    ['今回Doctorへ送る記事', 'Treatment Batch作成前'],
    ['Doctor Case Package', '未生成']
  ];

  sh.getRange(1,1,values.length,2).setValues(values);
  sh.getRange('A1:B1').merge();
  sh.getRange('A1').setValue('サイト診断サマリー');
  sh.getRange('A1').setFontWeight('bold').setFontSize(14);
  sh.getRange('A9').setFontWeight('bold');
  sh.getRange('A20').setFontWeight('bold');
  sh.getRange(1,1,values.length,1).setFontWeight('bold');
  sh.setColumnWidth(1, 230);
  sh.setColumnWidth(2, 620);
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
    sdsdProgress_(1, 3, '記事情報と検索データを確認しています');
    const enrichment = sdsdEnrichSelectedCases({silent:true});

    if (enrichment.failed > 0) {
      SpreadsheetApp.getUi().alert(
        `Doctor Case Packageを生成できません。\n\n` +
        `準備完了: ${enrichment.ready}件\n` +
        `要確認: ${enrichment.failed}件\n\n` +
        `「6. 選定案件を見る」で要確認案件を確認してください。`
      );
      return;
    }

    sdsdProgress_(2, 3, '記事本文をCase Packageへ添付しています');
    const exported = sdsdExportDoctorCasePackageZip({silent:true});

    sdsdUpdateSummaryAfterPackage_(exported.caseCount, exported.fileUrl);
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
      `Doctor Case Packageを生成できませんでした。\n\n${e.message || e}\n\n` +
      `記事管理データとCase Package関連ファイルを確認してください。`
    );
    throw e;
  }
}
