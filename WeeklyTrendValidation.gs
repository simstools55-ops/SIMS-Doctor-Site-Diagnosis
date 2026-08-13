function sdsdValidateWeeklyTrends() {
  const weeklyMap = sdsdBuildWeeklyTrendMap_();
  const wanted = ['GROWTH','VOLATILE','TRAFFIC_DECLINE','STABLE'];
  const selected = {};

  Object.keys(weeklyMap).forEach(url => {
    const info = sdsdClassifyWeeklyTrend_(weeklyMap[url]);
    if (wanted.indexOf(info.trend) >= 0 && !selected[info.trend]) {
      selected[info.trend] = {url:url, info:info, series:weeklyMap[url]};
    }
  });

  const ss = SpreadsheetApp.getActive();
  const name = 'Weekly Trend Validation';
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  sh.clear();

  const headers = [
    'Pattern','URL','Week Start','Week End','Clicks','Impressions','Position',
    'First-vs-Last Impression Change','Position Change','Volatility','Classification'
  ];
  sh.getRange(1,1,1,headers.length).setValues([headers]);

  const out = [];
  wanted.forEach(pattern => {
    const x = selected[pattern];
    if (!x) return;
    x.series.forEach((w,i) => out.push([
      pattern, x.url, w.weekStart, w.weekEnd, w.clicks, w.impressions, w.position,
      i===0 ? x.info.declineRatio : '',
      i===0 ? x.info.positionChange : '',
      i===0 ? x.info.volatility : '',
      i===0 ? x.info.trend : ''
    ]));
  });

  if (out.length) sh.getRange(2,1,out.length,headers.length).setValues(out);
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1,headers.length);
  ss.setActiveSheet(sh);

  SpreadsheetApp.getUi().alert(
    '4パターン回帰検証表を作成しました。\n' +
    'Weekly Trend Validation シートを確認してください。'
  );
}
