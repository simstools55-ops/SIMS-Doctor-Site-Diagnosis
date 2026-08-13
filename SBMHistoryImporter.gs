function sdsdBuildHistoryMap_() {
  const rows = sdsdReadObjects_(SDSD_CONFIG.sheets.sbmHistory);
  const map = {};
  rows.forEach(r => {
    const rawUrl = r['記事URL'] || r['URL'] || r['url'] || '';
    const url = sdsdNormalizeUrl_(rawUrl);
    if (!url) return;

    // SBM「改善履歴」シートの現行ヘッダーを最優先で読む。
    // 旧名称も後方互換として残す。
    const status = String(r['状態'] || r['作業状態'] || r['現在状態'] || r['判定'] || '');
    const dateText = r['改善実施日'] || r['改善日'] || r['処置日'] || r['日付'] || '';
    const route = String(r['改善経路'] || r['経路'] || '');

    const d = dateText ? new Date(dateText) : null;
    const prev = map[url];

    if (!prev || (d && (!prev.date || d > prev.date))) {
      map[url] = { date:d, status, route };
    }
  });
  return map;
}

function sdsdRecentTreatmentGuard_(url, historyMap) {
  const h = historyMap[url];
  if (!h) return {status:'PASS', reason:''};

  if (String(h.status).indexOf('モニター中') >= 0) {
    return {status:'WAIT', reason:'SBMモニター中'};
  }

  if (h.date && !isNaN(h.date.getTime())) {
    const days = (Date.now() - h.date.getTime()) / 86400000;
    if (days >= 0 && days <= SDSD_CONFIG.guardDays) {
      return {status:'WAIT', reason:`直近${Math.floor(days)}日前に改善済み`};
    }
  }
  return {status:'PASS', reason:''};
}
