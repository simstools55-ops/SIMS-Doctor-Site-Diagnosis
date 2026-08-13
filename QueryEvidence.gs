function sdsdBuildQueryEvidenceMap_() {
  const rows = sdsdReadObjects_(SDSD_CONFIG.sheets.evidencePageQuery);
  const map = {};
  rows.forEach(r => {
    const raw = r.page || r.key || r.url || '';
    const url = sdsdNormalizeUrl_(raw);
    if (!url) return;
    const arr = map[url] || [];
    arr.push({
      query: String(r.query || ''),
      clicks: Number(r.clicks || 0),
      impressions: Number(r.impressions || 0),
      ctr: Number(r.ctr || 0),
      position: Number(r.position || 0)
    });
    map[url] = arr;
  });
  Object.keys(map).forEach(url => {
    map[url].sort((a,b) => b.impressions - a.impressions);
  });
  return map;
}
