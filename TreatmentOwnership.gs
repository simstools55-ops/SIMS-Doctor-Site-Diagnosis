function sdsdOwnership_(x) {
  const impDecline = x.impressionsPrevious > 0
    ? (x.impressionsPrevious - x.impressionsRecent) / x.impressionsPrevious
    : 0;
  const posWorse = x.positionRecent - x.positionPrevious;
  const ctr = x.impressionsRecent ? x.clicksRecent / x.impressionsRecent : 0;
  const expected = sdsdExpectedCtr_(x.positionRecent);
  const pureCtrGap = expected > 0 && ctr < expected * 0.5 && impDecline < 0.15 && posWorse < 0.75;

  if (pureCtrGap) return {
    ownership:'SBM_OWNED',
    reason:'主病変がCTR/即効性改善'
  };

  if (impDecline >= 0.25 || posWorse >= 1.5) return {
    ownership:'DOCTOR_OWNED',
    reason:'流入低下または順位悪化が強い'
  };

  return {
    ownership:'REVIEW',
    reason:'追加Evidence確認'
  };
}
