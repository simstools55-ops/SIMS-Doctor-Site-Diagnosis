function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('SIMS Doctor Site Diagnosis')
    .addItem('1. Initialize', 'sdsdInitialize')
    .addItem('2. Import Evidence CSVs', 'sdsdImportEvidenceHelp')
    .addItem('3. Import SBM History CSV', 'sdsdImportHistoryHelp')
    .addItem('4. Run Site Analysis', 'sdsdRunAnalysis')
    .addItem('5. Open Candidates', 'sdsdOpenCandidates')
    .addItem('6. Validate Weekly Trends', 'sdsdValidateWeeklyTrends')
    .addItem('7. Validate Final Priorities', 'sdsdValidateFinalPriorities')
    .addToUi();
}

function sdsdInitialize() {
  sdsdEnsureSheets_();
  SpreadsheetApp.getUi().alert(
    `SIMS Doctor Site Diagnosis ${SDSD_VERSION}\n初期化しました。`
  );
}

function sdsdImportEvidenceHelp() {
  SpreadsheetApp.getUi().alert(
    'Sprint 1では、Evidence ZIP内の page_summary.csv / page_weekly.csv / page_query_top.csv を\n' +
    '対応する非表示シートへ貼り付けてから分析します。\n\n' +
    '次SprintでZIP直接読込へ置き換える予定です。'
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
