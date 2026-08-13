function sdsdEvidenceConfidence_(x, weeklyInfo) {
  if (weeklyInfo && weeklyInfo.confidence === 'LOW') return 'LOW';
  if ((x.impressionsRecent || 0) < 50 && (x.impressionsFull || 0) < 200) return 'LOW';
  if ((x.impressionsRecent || 0) < 200) return 'MEDIUM';
  return weeklyInfo && weeklyInfo.confidence === 'MEDIUM' ? 'MEDIUM' : 'HIGH';
}

function sdsdTreatmentRisk_(x, weeklyInfo) {
  const highTraffic = (x.impressionsRecent || 0) >= 20000;
  const highRank = (x.positionRecent || 99) <= 5;
  const volatile = weeklyInfo && weeklyInfo.trend === 'VOLATILE';
  const growth = weeklyInfo && weeklyInfo.trend === 'GROWTH';

  if (growth) return 'HIGH';
  if (highTraffic && highRank) return 'HIGH';
  if (volatile || highTraffic) return 'MEDIUM';
  return 'LOW';
}

function sdsdExternalFactorFlag_(x, queryMap) {
  const queries = queryMap[x.url] || [];
  const text = queries.map(q => q.query).join(' ').toLowerCase();

  const flags = [];
  if (/chatgpt|openai/.test(text)) flags.push('FAST_CHANGING_SERVICE');
  if (/windows 11|windows11|iphone|ios|apple watch|line|instagram|インスタ|x /.test(text)) {
    flags.push('PLATFORM_OR_OS_CHANGE');
  }

  return flags;
}
