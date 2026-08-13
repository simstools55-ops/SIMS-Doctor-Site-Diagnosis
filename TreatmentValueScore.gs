function sdsdPercentileRanks_(items, getter) {
  const sorted = items.map(getter).filter(v => isFinite(v)).sort((a,b)=>a-b);
  return function(value) {
    if (!sorted.length) return 0;
    let count = 0;
    while (count < sorted.length && sorted[count] <= value) count++;
    return count / sorted.length;
  };
}

function sdsdExpectedCtr_(position) {
  if (!position) return 0;
  if (position <= 3) return 0.10;
  if (position <= 5) return 0.06;
  if (position <= 7) return 0.04;
  if (position <= 10) return 0.025;
  if (position <= 15) return 0.012;
  if (position <= 20) return 0.007;
  return 0.004;
}

function sdsdScoreAll_(items) {
  const demandRank = sdsdPercentileRanks_(items, x => x.impressionsRecent);
  const assetRank = sdsdPercentileRanks_(items, x => x.impressionsFull);

  return items.map(x => {
    const demand = Math.round(demandRank(x.impressionsRecent) * 30 * 10) / 10;
    const asset = Math.round(assetRank(x.impressionsFull) * 15 * 10) / 10;

    let opportunity = 0;
    if (x.positionRecent >= 4 && x.positionRecent <= 10) opportunity += 15;
    else if (x.positionRecent > 10 && x.positionRecent <= 15) opportunity += 12;
    else if (x.positionRecent > 15 && x.positionRecent <= 20) opportunity += 8;

    if (x.impressionsRecent >= 50) {
      const ctr = x.impressionsRecent ? x.clicksRecent / x.impressionsRecent : 0;
      const expected = sdsdExpectedCtr_(x.positionRecent);
      const ratio = expected ? ctr / expected : 1;
      if (ratio <= 0.25) opportunity += 15;
      else if (ratio <= 0.5) opportunity += 11;
      else if (ratio <= 0.75) opportunity += 7;
      else if (ratio < 1) opportunity += 3;
    }

    let urgency = 0;
    if (x.impressionsPrevious > 0) {
      const decline = (x.impressionsPrevious - x.impressionsRecent) / x.impressionsPrevious;
      if (decline >= 0.5) urgency += 12;
      else if (decline >= 0.35) urgency += 10;
      else if (decline >= 0.2) urgency += 7;
      else if (decline >= 0.1) urgency += 4;
    }
    const posWorse = x.positionRecent - x.positionPrevious;
    if (posWorse >= 3) urgency += 8;
    else if (posWorse >= 2) urgency += 6;
    else if (posWorse >= 1) urgency += 4;
    else if (posWorse >= 0.5) urgency += 2;

    if (x.clicksPrevious > 0) {
      const cDecline = (x.clicksPrevious - x.clicksRecent) / x.clicksPrevious;
      if (cDecline >= 0.5) urgency += 5;
      else if (cDecline >= 0.25) urgency += 3;
    }

    urgency = Math.min(25, urgency);
    opportunity = Math.min(30, opportunity);

    return Object.assign({}, x, {
      demand, opportunity, urgency, asset,
      tvs: Math.round((demand + opportunity + urgency + asset) * 10) / 10
    });
  });
}
