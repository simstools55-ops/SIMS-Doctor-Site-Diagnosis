function sdsdRunAnalysis() {
  sdsdEnsureSheets_();

  const evidence = sdsdBuildEvidenceMap_();
  const evidenceMap = evidence.map;
  const historyMap = sdsdBuildHistoryMap_();
  const weeklyMap = sdsdBuildWeeklyTrendMap_();
  const queryMap = sdsdBuildQueryEvidenceMap_();
  const items = Object.keys(evidenceMap).map(k => evidenceMap[k]);

  let scored = sdsdScoreAll_(items);

  scored = scored.map(x => {
    const own = sdsdOwnership_(x);
    const guard = sdsdRecentTreatmentGuard_(x.url, historyMap);
    const weekly = sdsdClassifyWeeklyTrend_(weeklyMap[x.url] || []);
    const evidenceConfidence = sdsdEvidenceConfidence_(x, weekly);
    const treatmentRisk = sdsdTreatmentRisk_(x, weekly);
    const externalFlags = sdsdExternalFactorFlag_(x, queryMap);

    let priority = 'CANDIDATE';

    if (guard.status === 'WAIT') {
      priority = 'WAIT';
    } else if (own.ownership === 'SBM_OWNED') {
      priority = 'SBM';
    } else if (weekly.trend === 'GROWTH') {
      priority = 'PROTECTED';
    } else if (evidenceConfidence === 'LOW') {
      priority = 'REVIEW';
    } else if (own.ownership === 'REVIEW') {
      if (weekly.trend === 'SEVERE_DECLINE' || weekly.trend === 'TRAFFIC_DECLINE' || weekly.trend === 'RANKING_DECLINE') {
        priority = 'DOCTOR_REVIEW';
      } else {
        priority = 'REVIEW';
      }
    } else if (own.ownership === 'DOCTOR_OWNED') {
      if (treatmentRisk === 'HIGH' && weekly.trend !== 'SEVERE_DECLINE') {
        priority = 'DOCTOR_REVIEW';
      } else if (x.tvs >= 70) {
        priority = 'A1_CANDIDATE';
      } else if (x.tvs >= 60) {
        priority = 'A2_CANDIDATE';
      } else if (x.tvs >= 50) {
        priority = 'B_CANDIDATE';
      } else {
        priority = 'CANDIDATE';
      }
    } else {
      priority = 'REVIEW';
    }

    const reasonParts = [own.reason, guard.reason];
    if (weekly.trend && weekly.trend !== 'STABLE') reasonParts.push(`週次:${weekly.trend}`);
    if (externalFlags.length) reasonParts.push(`外部要因:${externalFlags.join('|')}`);
    if (treatmentRisk === 'HIGH') reasonParts.push('高Risk');

    return Object.assign({}, x, {
      ownership: own.ownership,
      guard: guard.status,
      weeklyTrend: weekly.trend,
      evidenceConfidence,
      treatmentRisk,
      externalFactor: externalFlags.join('|') || '',
      priority,
      reason: reasonParts.filter(Boolean).join(' / ')
    });
  });

  scored.sort((a,b) => b.tvs - a.tvs);
  sdsdWriteCandidates_(scored);

  SpreadsheetApp.getUi().alert(
    `Site Analysis完了\n対象記事: ${scored.length}\n` +
    `Article Universe: ${evidence.universeCount} (${evidence.universeStrategy})\n` +
    `Sprint 2: 週次Trend / Confidence / Risk / 外部要因を反映しました。\n` +
    `上位候補を「${SDSD_CONFIG.sheets.candidates}」へ出力しました。`
  );
}
