function sdsdRunAnalysis() {
  sdsdEnsureSheets_();

  const evidenceMap = sdsdBuildEvidenceMap_();
  const historyMap = sdsdBuildHistoryMap_();
  const items = Object.keys(evidenceMap).map(k => evidenceMap[k]);

  let scored = sdsdScoreAll_(items);

  scored = scored.map(x => {
    const own = sdsdOwnership_(x);
    const guard = sdsdRecentTreatmentGuard_(x.url, historyMap);

    let priority = 'CANDIDATE';
    if (guard.status === 'WAIT') priority = 'WAIT';
    else if (own.ownership === 'SBM_OWNED') priority = 'SBM';
    else if (x.tvs >= 70) priority = 'A1_CANDIDATE';
    else if (x.tvs >= 60) priority = 'A2_CANDIDATE';
    else if (x.tvs >= 50) priority = 'B_CANDIDATE';

    return Object.assign({}, x, {
      ownership: own.ownership,
      guard: guard.status,
      priority,
      reason: [own.reason, guard.reason].filter(Boolean).join(' / ')
    });
  });

  scored.sort((a,b) => b.tvs - a.tvs);
  sdsdWriteCandidates_(scored);

  SpreadsheetApp.getUi().alert(
    `Site Analysis完了\n対象URL: ${scored.length}\n` +
    `上位候補を「${SDSD_CONFIG.sheets.candidates}」へ出力しました。`
  );
}
