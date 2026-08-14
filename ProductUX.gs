const SDSD_INTERNAL_SHEET_NAMES_ = Object.freeze([
  '_SDSD_PAGE_SUMMARY',
  '_SDSD_PAGE_WEEKLY',
  '_SDSD_PAGE_QUERY_TOP',
  '_SDSD_SBM_HISTORY',
  '_SDSD_ARTICLE_MASTER',
  'Weekly Trend Validation',
  'Priority Validation',
  'Query Evidence Diagnostics'
]);

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
    '内部処理用シートを非表示にしました。\\n通常利用では「診断候補」と「選定案件」だけ確認すれば大丈夫です。'
  );
}

function sdsdShowInternalSheets() {
  const ss = SpreadsheetApp.getActive();
  let count = 0;
  SDSD_INTERNAL_SHEET_NAMES_.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    try {
      sh.showSheet();
      count++;
    } catch (e) {}
  });
  SpreadsheetApp.getUi().alert(
    `保守確認用として内部シートを表示しました。\\n表示: ${count}シート\\n\\n確認後は「内部シートを隠す」を実行してください。`
  );
}

function sdsdRunProductDiagnosis() {
  sdsdHideInternalSheets_();
  const result = sdsdRunAnalysis({silent:true});
  sdsdHideInternalSheets_();

  SpreadsheetApp.getUi().alert(
    `サイト診断が完了しました。\\n\\n` +
    `対象記事: ${result.total}件\\n` +
    `精密診断の優先候補: ${result.priorityCandidates}件\\n` +
    `最近処置済み・モニター中: ${result.wait}件\\n` +
    `回復・成長中のため保護: ${result.protected}件\\n` +
    `SBMの日常改善対象: ${result.sbm}件\\n\\n` +
    `次に「4. 診断候補を見る」で内容を確認し、\\n` +
    `問題なければ「5. Treatment Batchを作成」を実行してください。`
  );
}

function sdsdCreateProductTreatmentBatch() {
  sdsdProgress_(1, 2, '治療候補を選定しています');
  const batch = sdsdBuildTreatmentBatch({silent:true, noActivate:true});

  sdsdProgress_(2, 2, '最近の処置履歴を最終確認しています');
  const guard = sdsdRunFinalGuard({silent:true});
  sdsdHideInternalSheets_();

  SpreadsheetApp.getUi().alert(
    `Treatment Batchを作成しました。\\n\\n` +
    `診断対象: ${batch.articleCount}記事\\n` +
    `適格候補: ${batch.eligibleCount}件\\n` +
    `今回選定: ${batch.selectedCount}件\\n` +
    `最終確認で保留: ${guard.blocked}件\\n\\n` +
    `次に「6. 選定案件を見る」で内容を確認してください。`
  );
}

function sdsdCreateProductCasePackage() {
  sdsdProgress_(1, 3, '記事情報と検索データを確認しています');
  const enrichment = sdsdEnrichSelectedCases({silent:true});

  if (enrichment.failed > 0) {
    SpreadsheetApp.getUi().alert(
      `Doctor Case Packageを生成できません。\\n\\n` +
      `準備完了: ${enrichment.ready}件\\n` +
      `要確認: ${enrichment.failed}件\\n\\n` +
      `「6. 選定案件を見る」で要確認案件を確認してください。`
    );
    return;
  }

  sdsdProgress_(2, 3, '記事本文をCase Packageへ添付しています');
  const exported = sdsdExportDoctorCasePackageZip({silent:true});

  sdsdProgress_(3, 3, 'Doctor Case Packageを保存しました');
  SpreadsheetApp.getUi().alert(
    `Doctor Case Packageの生成が完了しました。\\n\\n` +
    `案件数: ${exported.caseCount}件\\n` +
    `本文再取得: ${exported.refetched}件\\n\\n` +
    `Google DriveにZIPを保存しました。\\n${exported.fileUrl}\\n\\n` +
    `このZIPをSIMS Doctorへ渡してください。`
  );
}
