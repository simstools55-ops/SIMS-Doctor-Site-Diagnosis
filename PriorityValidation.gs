function sdsdValidateFinalPriorities() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SDSD_CONFIG.sheets.candidates);
  if (!sh) throw new Error('先に Run Site Analysis を実行してください。');

  const values = sh.getDataRange().getValues();
  if (values.length < 2) throw new Error('候補データがありません。');

  const headers = values[0].map(String);
  const idx = {};
  headers.forEach((h,i) => idx[h] = i);

  const priorityCol = idx['Priority Candidate'];
  if (priorityCol == null) throw new Error('Priority Candidate列が見つかりません。');

  const rows = values.slice(1).filter(r => r[1]);
  const groups = {};
  rows.forEach(r => {
    const p = String(r[priorityCol] || 'EMPTY');
    (groups[p] ||= []).push(r);
  });

  const order = ['A1_CANDIDATE','A2_CANDIDATE','B_CANDIDATE','DOCTOR_REVIEW',
                 'REVIEW','WAIT','SBM','PROTECTED','CANDIDATE','EMPTY'];

  const outName = 'Priority Validation';
  const ss = SpreadsheetApp.getActive();
  let out = ss.getSheetByName(outName);
  if (!out) out = ss.insertSheet(outName);
  out.clear();

  const summaryHeaders = ['Priority','Count','Share','Meaning','Gate Check'];
  out.getRange(1,1,1,summaryHeaders.length).setValues([summaryHeaders]);

  const meanings = {
    A1_CANDIDATE:'Doctor最優先候補',
    A2_CANDIDATE:'Doctor次点候補',
    B_CANDIDATE:'Doctor低優先候補',
    DOCTOR_REVIEW:'Doctor追加確認',
    REVIEW:'Evidence/Ownership追加確認',
    WAIT:'最近処置済み・モニター中',
    SBM:'SBMの日常改善へ',
    PROTECTED:'回復・成長中のため保護',
    CANDIDATE:'低優先候補',
    EMPTY:'未分類'
  };

  const summary = [];
  order.forEach(p => {
    const count = (groups[p] || []).length;
    if (!count) return;
    let gate = 'PASS';
    if (p === 'EMPTY') gate = 'FAIL';
    summary.push([p,count,count/rows.length,meanings[p] || '',gate]);
  });
  if (summary.length) out.getRange(2,1,summary.length,summaryHeaders.length).setValues(summary);

  const start = summary.length + 4;
  const detailHeaders = ['Priority','Rank','URL','TVS','Ownership','Guard','Weekly Trend',
                         'Evidence Confidence','Treatment Risk','External Factor','Reason'];
  out.getRange(start,1,1,detailHeaders.length).setValues([detailHeaders]);

  const sample = [];
  order.forEach(p => {
    (groups[p] || []).slice(0,5).forEach(r => sample.push([
      p,
      r[idx['Rank']],
      r[idx['Normalized URL']],
      r[idx['TVS']],
      r[idx['Ownership']],
      r[idx['Recent Treatment Guard']],
      r[idx['Weekly Trend']],
      r[idx['Evidence Confidence']],
      r[idx['Treatment Risk']],
      r[idx['External Factor']],
      r[idx['Reason']]
    ]));
  });
  if (sample.length) out.getRange(start+1,1,sample.length,detailHeaders.length).setValues(sample);

  out.setFrozenRows(1);
  out.getRange(2,3,Math.max(summary.length,1),1).setNumberFormat('0.0%');
  out.autoResizeColumns(1,detailHeaders.length);
  ss.setActiveSheet(out);

  SpreadsheetApp.getUi().alert(
    `最終Priority検証表を作成しました。\n対象記事: ${rows.length}\n` +
    `Priority Validation シートを確認してください。`
  );
}
