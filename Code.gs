function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('SIMS Doctor Site Diagnosis')
    .addItem('1. 初期化', 'sdsdInitialize')
    .addItem('2. Evidence Package ZIPを読み込む', 'sdsdImportEvidencePackageZip')
    .addItem('3. SBM改善履歴を取り込む', 'sdsdImportHistoryHelp')
    .addItem('4. サイト分析を実行', 'sdsdRunAnalysis')
    .addItem('5. 診断候補を開く', 'sdsdOpenCandidates')
    .addItem('6. 週次トレンドを検証', 'sdsdValidateWeeklyTrends')
    .addItem('7. 最終優先度を検証', 'sdsdValidateFinalPriorities')
    .addSeparator()
    .addItem('8. 治療バッチを作成', 'sdsdBuildTreatmentBatch')
    .addItem('9. Final Guardを実行', 'sdsdRunFinalGuard')
    .addItem('10. 選定案件を開く', 'sdsdOpenSelectedCases')
    .addSeparator()
    .addItem('11. Article Master取込案内', 'sdsdArticleMasterImportHelp')
    .addItem('12. Case Enrichmentを実行', 'sdsdEnrichSelectedCases')
    .addItem('13. Doctor Case Package ZIPを生成', 'sdsdExportDoctorCasePackageZip')
    .addSeparator()
    .addItem('保守: Query Evidenceを診断', 'sdsdDiagnoseQueryEvidenceInput')
    .addToUi();
}

function sdsdInitialize() {
  sdsdEnsureSheets_();
  SpreadsheetApp.getUi().alert(
    `SIMS Doctor Site Diagnosis ${SDSD_VERSION}\n初期化しました。`
  );
}


function sdsdImportHistoryHelp() {
  SpreadsheetApp.getUi().alert(
    'SBMの「改善履歴」CSVを _SDSD_SBM_HISTORY シートへ貼り付けてください。'
  );
}

function sdsdOpenCandidates() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SDSD_CONFIG.sheets.candidates);
  if (sh) SpreadsheetApp.getActive().setActiveSheet(sh);
}
