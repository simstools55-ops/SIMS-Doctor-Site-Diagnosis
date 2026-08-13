function sdsdFirstNumber_() {
  for (let i = 0; i < arguments.length; i++) {
    const v = arguments[i];
    if (v !== null && v !== undefined && v !== '') {
      const n = Number(v);
      if (!isNaN(n)) return n;
    }
  }
  return 0;
}

function sdsdBuildEvidenceMap_() {
  const rows = sdsdReadObjects_(SDSD_CONFIG.sheets.evidencePageSummary);
  const map = {};
  rows.forEach(r => {
    const raw = r.key || r.page || r.url || '';
    const url = sdsdNormalizeUrl_(raw);
    if (!url) return;

    const item = map[url] || {
      url,
      clicksFull: 0,
      impressionsFull: 0,
      clicksRecent: 0,
      impressionsRecent: 0,
      positionRecent: 0,
      clicksPrevious: 0,
      impressionsPrevious: 0,
      positionPrevious: 0
    };

    item.clicksFull += sdsdFirstNumber_(r.clicks_full, r.clicks_180d, r.clicks);
    item.impressionsFull += sdsdFirstNumber_(r.impressions_full, r.impressions_180d, r.impressions);

    item.clicksRecent += sdsdFirstNumber_(r.clicks_recent28, r.clicks_recent28d);
    item.impressionsRecent += sdsdFirstNumber_(r.impressions_recent28, r.impressions_recent28d);
    item.positionRecent = sdsdFirstNumber_(r.position_recent28, r.position_recent28d, item.positionRecent);

    item.clicksPrevious += sdsdFirstNumber_(r.clicks_previous28, r.clicks_previous28d);
    item.impressionsPrevious += sdsdFirstNumber_(r.impressions_previous28, r.impressions_previous28d);
    item.positionPrevious = sdsdFirstNumber_(r.position_previous28, r.position_previous28d, item.positionPrevious);

    map[url] = item;
  });
  return map;
}
