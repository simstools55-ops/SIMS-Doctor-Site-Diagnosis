function sdsdRunAnalysis(options) {
  options = options || {};
  sdsdProgress_(1, 5, 'Evidence Packageのデータを確認しています');
  sdsdProductEnsureSheets_();

  const evidence = sdsdBuildEvidenceMap_();
  const evidenceMap = evidence.map;

  sdsdProgress_(2, 5, '改善履歴と記事状態を確認しています');
  const historyMap = sdsdBuildHistoryMap_();

  sdsdProgress_(3, 5, '週次推移と検索クエリを分析しています');
  const weeklyMap = sdsdBuildWeeklyTrendMap_();
  const queryMap = sdsdBuildQueryEvidenceMap_();
  const items = Object.keys(evidenceMap).map(k => evidenceMap[k]);

  sdsdProgress_(4, 5, '診断候補の優先度を評価しています');
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
  sdsdProgress_(5, 5, '診断結果を整理しています');
  sdsdWriteCandidates_(scored);

  const result = {
    total: scored.length,
    priorityCandidates: scored.filter(x =>
      x.priority === 'A1_CANDIDATE' || x.priority === 'A2_CANDIDATE'
    ).length,
    wait: scored.filter(x => x.priority === 'WAIT').length,
    protected: scored.filter(x => x.priority === 'PROTECTED').length,
    sbm: scored.filter(x => x.priority === 'SBM').length,
    review: scored.filter(x =>
      x.priority === 'REVIEW' || x.priority === 'DOCTOR_REVIEW'
    ).length,
    universeCount: evidence.universeCount,
    universeStrategy: evidence.universeStrategy
  };

  sdsdWriteSiteSummary_(scored, result);

  if (!options.silent) {
    SpreadsheetApp.getUi().alert(
      `サイト診断が完了しました。\n\n` +
      `対象記事: ${result.total}件\n` +
      `精密診断の優先候補: ${result.priorityCandidates}件\n` +
      `最近処置済み・モニター中: ${result.wait}件\n` +
      `回復・成長中のため保護: ${result.protected}件\n\n` +
      `「${SDSD_CONFIG.sheets.candidates}」で結果を確認してください。`
    );
  }
  return result;
}
