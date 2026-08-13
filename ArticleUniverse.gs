function sdsdLooksLikeArticleUrl_(url) {
  const s = String(url || '').trim().split('#')[0].split('?')[0];

  // Hatena Blog article URL.
  if (/^https?:\/\/[^\/]+\/entry\/\d{4}\/\d{2}\/\d{2}\/[^\/?#]+\/?$/i.test(s)) return true;

  // Common WordPress numeric permalink.
  if (/^https?:\/\/[^\/]+\/\d+\/?$/i.test(s)) return true;

  // Dated permalink fallback.
  if (/^https?:\/\/[^\/]+\/\d{4}\/\d{2}\/\d{2}\/.+/i.test(s)) return true;

  return false;
}

function sdsdBuildArticleUniverse_() {
  // Preferred source: page_query_top is already narrowed by Collector RC5.
  const queryRows = sdsdReadObjects_(SDSD_CONFIG.sheets.evidencePageQuery);
  const fromPageQuery = {};
  queryRows.forEach(r => {
    const raw = r.page || r.key || r.url || '';
    const url = sdsdNormalizeUrl_(raw);
    if (url) fromPageQuery[url] = true;
  });

  // If Collector provided a plausible active-page set, use it as the primary universe.
  const pqUrls = Object.keys(fromPageQuery);
  if (pqUrls.length >= 20) {
    return {
      strategy: 'COLLECTOR_PAGE_QUERY_UNIVERSE',
      urls: fromPageQuery,
      count: pqUrls.length
    };
  }

  // Fallback: derive article-shaped URLs from page_summary.
  const summaryRows = sdsdReadObjects_(SDSD_CONFIG.sheets.evidencePageSummary);
  const heuristic = {};
  summaryRows.forEach(r => {
    const raw = r.key || r.page || r.url || '';
    if (!sdsdLooksLikeArticleUrl_(raw)) return;
    const url = sdsdNormalizeUrl_(raw);
    if (url) heuristic[url] = true;
  });

  return {
    strategy: 'ARTICLE_URL_HEURISTIC',
    urls: heuristic,
    count: Object.keys(heuristic).length
  };
}
