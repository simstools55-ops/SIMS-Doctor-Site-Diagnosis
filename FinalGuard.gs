function sdsdRunFinalGuard(options) {
  options = options || {};
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SDSD_CONFIG.sheets.selectedCases);
  if (!sh) throw new Error('先に Treatment Batch を生成してください。');

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return {checked:0, blocked:0};

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

  // 利用者向け「状態」列も同期する。
  const userStatusIdx = headers.indexOf('状態');
  if (userStatusIdx >= 0 && values.length > 1) {
    const refreshed = sh.getDataRange().getValues();
    const refIdx = headers.indexOf('Referral Status');
    const statuses = refreshed.slice(1).map(r => {
      const ref = String(r[refIdx] || '');
      return [ref.indexOf('BLOCK') >= 0 ? '保留' : 'Doctor診断待ち'];
    });
    if (statuses.length) sh.getRange(2, userStatusIdx + 1, statuses.length, 1).setValues(statuses);
  }

  const result = {checked: values.length - 1, blocked: blocked};
  if (!options.silent) {
    SpreadsheetApp.getUi().alert(
      `最終確認が完了しました。\n再確認対象: ${result.checked}件\n保留: ${result.blocked}件`
    );
  }
  return result;
}
