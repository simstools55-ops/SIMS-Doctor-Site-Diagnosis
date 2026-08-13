function sdsdBuildWeeklyTrendMap_() {
  const rows = sdsdReadObjects_(SDSD_CONFIG.sheets.evidencePageWeekly);
  const grouped = {};

  rows.forEach(r => {
    const raw = r.page || r.key || r.url || '';
    const url = sdsdNormalizeUrl_(raw);
    if (!url) return;

    const weekStart = sdsdDateKey_(r.week_start || r.start_date || '');
    const weekEnd = sdsdDateKey_(r.week_end || r.end_date || '');
    if (!weekStart) return;

    const key = `${url}||${weekStart}`;
    const clicks = Number(r.clicks || 0);
    const impressions = Number(r.impressions || 0);
    const position = Number(r.position || 0);

    const g = grouped[key] || {
      url,
      weekStart,
      weekEnd,
      clicks: 0,
      impressions: 0,
      positionWeightedSum: 0,
      positionFallbackSum: 0,
      positionFallbackCount: 0
    };

    g.clicks += clicks;
    g.impressions += impressions;

    if (impressions > 0 && isFinite(position)) {
      g.positionWeightedSum += position * impressions;
    } else if (isFinite(position) && position > 0) {
      g.positionFallbackSum += position;
      g.positionFallbackCount += 1;
    }

    grouped[key] = g;
  });

  const map = {};

  Object.keys(grouped).forEach(key => {
    const g = grouped[key];

    let position = 0;
    if (g.impressions > 0) {
      position = g.positionWeightedSum / g.impressions;
    } else if (g.positionFallbackCount > 0) {
      position = g.positionFallbackSum / g.positionFallbackCount;
    }

    const arr = map[g.url] || [];
    arr.push({
      weekStart: g.weekStart,
      weekEnd: g.weekEnd,
      clicks: g.clicks,
      impressions: g.impressions,
      position
    });
    map[g.url] = arr;
  });

  Object.keys(map).forEach(url => {
    map[url].sort((a,b) => String(a.weekStart).localeCompare(String(b.weekStart)));
  });

  return map;
}

function sdsdDateKey_(value) {
  if (!value) return '';

  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone() || 'Asia/Tokyo', 'yyyy-MM-dd');
  }

  const s = String(value).trim();

  // Common YYYY-MM-DD text.
  const direct = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (direct) return `${direct[1]}-${direct[2]}-${direct[3]}`;

  // Google Sheets can expose a Date as a locale string.
  const d = new Date(value);
  if (!isNaN(d.getTime())) {
    return Utilities.formatDate(d, Session.getScriptTimeZone() || 'Asia/Tokyo', 'yyyy-MM-dd');
  }

  return s;
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

  const avg = (arr, key) =>
    arr.reduce((a,x) => a + Number(x[key] || 0), 0) / arr.length;

  const firstImp = avg(first, 'impressions');
  const lastImp = avg(last, 'impressions');
  const firstPos = avg(first, 'position');
  const lastPos = avg(last, 'position');

  const declineRatio = firstImp > 0 ? (firstImp - lastImp) / firstImp : 0;
  const positionChange = lastPos - firstPos;

  const impAvg = avg(valid, 'impressions');
  const variance =
    valid.reduce((a,x) => a + Math.pow((x.impressions || 0) - impAvg, 2), 0) / valid.length;
  const volatility = impAvg > 0 ? Math.sqrt(variance) / impAvg : 0;

  let trend = 'STABLE';

  if (declineRatio >= 0.45 && positionChange >= 1.0) {
    trend = 'SEVERE_DECLINE';
  } else if (declineRatio >= 0.30) {
    trend = 'TRAFFIC_DECLINE';
  } else if (positionChange >= 1.5) {
    trend = 'RANKING_DECLINE';
  } else if (declineRatio <= -0.30) {
    trend = 'GROWTH';
  } else if (volatility >= 0.60) {
    trend = 'VOLATILE';
  }

  let confidence = 'HIGH';
  if (valid.length < 8) confidence = 'MEDIUM';
  if (impAvg < 25) confidence = 'LOW';

  return {
    trend,
    confidence,
    declineRatio,
    positionChange,
    volatility
  };
}
