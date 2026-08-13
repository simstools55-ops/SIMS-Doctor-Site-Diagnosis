function sdsdObjectValue_(obj, names) {
  const keys = Object.keys(obj || {});
  const normalized = {};
  keys.forEach(k => {
    const nk = String(k || '').replace(/^\uFEFF/, '').trim().toLowerCase();
    normalized[nk] = obj[k];
  });

  for (let i=0; i<names.length; i++) {
    const key = String(names[i]).trim().toLowerCase();
    if (normalized[key] !== undefined && normalized[key] !== null) {
      return normalized[key];
    }
  }
  return '';
}

function sdsdBuildQueryEvidenceMap_() {
  const rows = sdsdReadObjects_(SDSD_CONFIG.sheets.evidencePageQuery);
  const map = {};

  rows.forEach(r => {
    const raw = sdsdObjectValue_(r, ['page','key','url','URL','記事URL']);
    const url = sdsdNormalizeUrl_(raw);
    if (!url) return;

    const query = String(sdsdObjectValue_(r, ['query','クエリ','検索クエリ']) || '').trim();
    if (!query) return;

    const arr = map[url] || [];
    arr.push({
      query: query,
      clicks: Number(sdsdObjectValue_(r, ['clicks','クリック数']) || 0),
      impressions: Number(sdsdObjectValue_(r, ['impressions','表示回数']) || 0),
      ctr: Number(sdsdObjectValue_(r, ['ctr','CTR']) || 0),
      position: Number(sdsdObjectValue_(r, ['position','掲載順位','平均掲載順位']) || 0)
    });
    map[url] = arr;
  });

  Object.keys(map).forEach(url => {
    map[url].sort((a,b) => {
      if (b.impressions !== a.impressions) return b.impressions - a.impressions;
      return b.clicks - a.clicks;
    });
  });

  return map;
}

function sdsdQueryEvidenceSourceCount_() {
  return sdsdReadObjects_(SDSD_CONFIG.sheets.evidencePageQuery).length;
}
