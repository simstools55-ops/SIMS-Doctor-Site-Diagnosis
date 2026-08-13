function sdsdBuildEvidenceMap_() {
  const rows = sdsdReadObjects_(SDSD_CONFIG.sheets.evidencePageSummary);
  const universe = sdsdBuildArticleUniverse_();
  const map = {};

  rows.forEach(r => {
    const raw = r.key || r.page || r.url || '';
    const url = sdsdNormalizeUrl_(raw);
    if (!url) return;

    // Gate: only score URLs that belong to the Article Universe.
    if (!universe.urls[url]) return;

    const item = map[url] || {
      url,
      clicksFull:0, impressionsFull:0,
      clicksRecent:0, impressionsRecent:0, positionRecent:0,
      clicksPrevious:0, impressionsPrevious:0, positionPrevious:0
    };

    item.clicksFull += Number((r.clicks_full ?? r.clicks_180d ?? r.clicks ?? 0));
    item.impressionsFull += Number((r.impressions_full ?? r.impressions_180d ?? r.impressions ?? 0));

    item.clicksRecent += Number((r.clicks_recent28 ?? r.clicks_recent28d ?? 0));
    item.impressionsRecent += Number((r.impressions_recent28 ?? r.impressions_recent28d ?? 0));
    item.positionRecent = Number((r.position_recent28 ?? r.position_recent28d ?? item.positionRecent ?? 0));

    item.clicksPrevious += Number((r.clicks_previous28 ?? r.clicks_previous28d ?? 0));
    item.impressionsPrevious += Number((r.impressions_previous28 ?? r.impressions_previous28d ?? 0));
    item.positionPrevious = Number((r.position_previous28 ?? r.position_previous28d ?? item.positionPrevious ?? 0));

    map[url] = item;
  });

  return {
    map,
    universeStrategy: universe.strategy,
    universeCount: universe.count
  };
}
