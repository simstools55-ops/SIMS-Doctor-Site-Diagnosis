function onOpen() {
  const ui = SpreadsheetApp.getUi();

  const dataMenu = ui.createMenu('データ準備')
    .addItem('SBM改善履歴の取込案内', 'sdsdImportHistoryHelp')
    .addItem('Article Masterの取込案内', 'sdsdArticleMasterImportHelp');

  const maintenanceMenu = ui.createMenu('保守・診断')
    .addItem('Query Evidenceを診断', 'sdsdDiagnoseQueryEvidenceInput')
    .addItem('週次トレンド検証表を作成', 'sdsdValidateWeeklyTrends')
    .addItem('優先度検証表を作成', 'sdsdValidateFinalPriorities')
    .addSeparator()
    .addItem('内部シートを表示', 'sdsdShowInternalSheets')
    .addItem('内部シートを隠す', 'sdsdHideInternalSheets');

  ui.createMenu('SIMS Doctor Site Diagnosis')
    .addItem('1. 初期設定', 'sdsdInitialize')
    .addItem('2. Evidence Packageを読み込む', 'sdsdImportEvidencePackageZip')
    .addItem('3. サイト診断を実行', 'sdsdRunProductDiagnosis')
    .addItem('サイト診断サマリーを見る', 'sdsdOpenSiteSummary')
    .addItem('4. 診断候補を見る', 'sdsdOpenCandidates')
    .addSeparator()
    .addItem('5. Treatment Batchを作成', 'sdsdCreateProductTreatmentBatch')
    .addItem('6. 選定案件を見る', 'sdsdOpenSelectedCases')
    .addItem('7. Doctor Case Packageを生成', 'sdsdCreateProductCasePackage')
    .addSeparator()
    .addItem('8. サイト横断の改善機会を診断', 'sdsdRunSiteOpportunityDiagnosis')
    .addItem('9. サイト横断診断案件を作成', 'sdsdBuildSiteOpportunityCases')
    .addItem('10. サイト横断診断案件を見る', 'sdsdOpenSiteOpportunityCases')
    .addSeparator()
    .addSubMenu(dataMenu)
    .addSubMenu(maintenanceMenu)
    .addToUi();

  try { sdsdHideInternalSheets_(); } catch (e) {}
  try { sdsdRefreshCandidatesView_(); } catch (e) {}
  try { sdsdRefreshSelectedCasesView_(); } catch (e) {}
}

function sdsdInitialize() {
  sdsdProductEnsureSheets_();
  sdsdHideInternalSheets_();
  SpreadsheetApp.getUi().alert(
    `SIMS Doctor Site Diagnosis ${SDSD_VERSION}\n初期設定が完了しました。\n\n次に「2. Evidence Packageを読み込む」を実行してください。`
  );
}


function sdsdImportHistoryHelp() {
  SpreadsheetApp.getUi().alert(
    'SBMの「改善履歴」CSVを _SDSD_SBM_HISTORY シートへ貼り付けてください。'
  );
}

function sdsdOpenSiteSummary() {
  try { sdsdRefreshSiteSummaryFromCandidates_(); } catch (e) {}
  const sh = SpreadsheetApp.getActive().getSheetByName(SDSD_CONFIG.sheets.summary);
  if (sh) SpreadsheetApp.getActive().setActiveSheet(sh);
}

function sdsdOpenCandidates() {
  try {
    sdsdRefreshCandidatesView_();
  } catch (e) {
    SpreadsheetApp.getUi().alert(
      `診断候補の表示を更新できませんでした。\n\n${e.message || e}`
    );
  }
  const sh = SpreadsheetApp.getActive().getSheetByName(SDSD_CONFIG.sheets.candidates);
  if (sh) SpreadsheetApp.getActive().setActiveSheet(sh);
}
