function sdsdRunFinalGuard() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SDSD_CONFIG.sheets.selectedCases);
  if (!sh) throw new Error('先に Treatment Batch を生成してください。');

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return;

  const headers = values[0].map(String);
  const idx = {};
  headers.forEach((h,i) => idx[h] = i);

  const historyMap = sdsdBuildHistoryMap_();
  let blocked = 0;

  for (let i=1; i<values.length; i++) {
    const url = String(values[i][idx['URL']] || '');
    if (!url) continue;

    const guard = sdsdRecentTreatmentGuard_(url, historyMap);
    if (guard.status === 'WAIT') {
      sh.getRange(i+1, idx['Recent Treatment Guard']+1).setValue('WAIT');
      sh.getRange(i+1, idx['Referral Status']+1).setValue('BLOCKED_BY_FINAL_GUARD');
      blocked++;
    }
  }

  SpreadsheetApp.getUi().alert(
    `Final Guard完了\n再確認対象: ${values.length-1}件\nブロック: ${blocked}件`
  );
}
