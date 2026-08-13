function sdsdBuildWeeklyTrendMap_() {
  const rows = sdsdReadObjects_(SDSD_CONFIG.sheets.evidencePageWeekly);
  const map = {};

  rows.forEach(r => {
    const raw = r.page || r.key || r.url || '';
    const url = sdsdNormalizeUrl_(raw);
    if (!url) return;

    const item = map[url] || [];
    item.push({
      weekStart: String(r.week_start || r.start_date || ''),
      weekEnd: String(r.week_end || r.end_date || ''),
      clicks: Number(r.clicks || 0),
      impressions: Number(r.impressions || 0),
      position: Number(r.position || 0)
    });
    map[url] = item;
  });

  Object.keys(map).forEach(url => {
    map[url].sort((a,b) => String(a.weekStart).localeCompare(String(b.weekStart)));
  });
  return map;
}

function sdsdClassifyWeeklyTrend_(series) {
  if (!series || series.length < 4) {
    return {
      trend:'INSUFFICIENT',
      confidence:'LOW',
      declineRatio:0,
      positionChange:0,
      volatility:0
    };
  }

  const valid = series.filter(x => x.impressions > 0);
  if (valid.length < 4) {
    return {
      trend:'INSUFFICIENT',
      confidence:'LOW',
      declineRatio:0,
      positionChange:0,
      volatility:0
    };
  }

  const n = valid.length;
  const block = Math.max(2, Math.floor(n / 3));
  const first = valid.slice(0, block);
  const last = valid.slice(n - block);

  const avg = (arr, key) => arr.reduce((a,x)=>a+Number(x[key]||0),0) / arr.length;
  const firstImp = avg(first,'impressions');
  const lastImp = avg(last,'impressions');
  const firstPos = avg(first,'position');
  const lastPos = avg(last,'position');

  const declineRatio = firstImp > 0 ? (firstImp - lastImp) / firstImp : 0;
  const positionChange = lastPos - firstPos;

  const impAvg = avg(valid,'impressions');
  const variance = valid.reduce((a,x)=>a+Math.pow((x.impressions||0)-impAvg,2),0)/valid.length;
  const volatility = impAvg > 0 ? Math.sqrt(variance)/impAvg : 0;

  let trend = 'STABLE';
  if (declineRatio >= 0.45 && positionChange >= 1.0) trend = 'SEVERE_DECLINE';
  else if (declineRatio >= 0.30) trend = 'TRAFFIC_DECLINE';
  else if (positionChange >= 1.5) trend = 'RANKING_DECLINE';
  else if (declineRatio <= -0.30) trend = 'GROWTH';
  else if (volatility >= 0.60) trend = 'VOLATILE';

  let confidence = 'HIGH';
  if (valid.length < 8) confidence = 'MEDIUM';
  if (impAvg < 25) confidence = 'LOW';

  return {trend, confidence, declineRatio, positionChange, volatility};
}
