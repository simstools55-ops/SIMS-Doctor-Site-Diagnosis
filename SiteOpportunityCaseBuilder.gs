/**
 * SIMS Doctor Site Diagnosis v0.5.0 Sprint 2
 * Convert Site Improvement Plan findings into durable diagnosis cases.
 *
 * This does NOT export Doctor ZIPs yet.
 */

function sdsdOpportunityDigest_(text) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(text || ''),
    Utilities.Charset.UTF_8
  );
  return bytes.slice(0,6).map(b => {
    const v = b < 0 ? b + 256 : b;
    return ('0' + v.toString(16)).slice(-2);
  }).join('').toUpperCase();
}

function sdsdBuildOpportunityCaseId_(item) {
  const signature = [
    item.type || '',
    item.diagnosisTheme || item.parentTheme || item.theme || '',
    (item.targets || []).slice().sort().join('|')
  ].join('||');

  return 'SITE-OPP-' + sdsdOpportunityDigest_(signature);
}

function sdsdOpportunityRoute_(item) {
  const type = String(item.type || '');

  if (type === 'カニバリ疑い') {
    return {
      primary: 'Doctor',
      next: 'Doctor精密診断 → 必要ならMerge / Writer',
      purpose: '複数記事の検索意図と役割を比較し、維持・役割分担・統合の要否を判断'
    };
  }

  if (type === 'コンテンツギャップ') {
    return {
      primary: 'Doctor',
      next: 'Doctor確認 → Writer / Creatorを振り分け',
      purpose: '既存記事で補うべきか、新記事として分離すべきかを判断'
    };
  }

  if (type === '新規記事機会') {
    return {
      primary: 'Doctor',
      next: 'Doctor確認 → Creator',
      purpose: '既存記事との重複・カニバリを避けたうえで新記事化の妥当性を判断'
    };
  }

  return {
    primary: 'Doctor',
    next: 'Doctorで確認',
    purpose: 'サイト全体への影響と処置方針を確認'
  };
}

function sdsdWriteOpportunityCases_(items) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SDSD_CONFIG.sheets.opportunityCases);
  if (!sh) sh = ss.insertSheet(SDSD_CONFIG.sheets.opportunityCases);
  sh.clear();

  const titleMap = sdsdArticleTitleMap_();

  const headers = [
    'No.','案件ID','改善テーマ','診断テーマ',
    '対象記事数','対象記事','関連クエリ数','主な検索テーマ',
    '確信度','サイト全体への期待効果',
    '診断で確認すること','次の担当','状態'
  ];

  const values = items.map((item,i) => {
    const route = sdsdOpportunityRoute_(item);
    const caseId = sdsdBuildOpportunityCaseId_(item);

    return [
      i + 1,
      caseId,
      item.type,
      item.diagnosisTheme || item.parentTheme || item.theme,
      (item.targets || []).length,
      (item.targets || []).map(u =>
        sdsdDisplayTitle_(u, titleMap) + '\n' + u
      ).join('\n\n'),
      (item.queries || []).length,
      (item.queries || []).slice(0,5).join(' / ') +
        ((item.queries || []).length > 5
          ? ` ほか${item.queries.length - 5}件`
          : ''),
      item.confidence || '中',
      item.siteImpact || 'サイト全体への影響をDoctorで確認します。',
      route.purpose,
      route.next,
      'Doctor診断待ち'
    ];
  });

  sh.getRange(1,1,1,headers.length).setValues([headers]);
  if (values.length) {
    sh.getRange(2,1,values.length,headers.length).setValues(values);
  }

  sh.setFrozenRows(1);
  sh.getRange(1,1,1,headers.length).setFontWeight('bold');
  sh.getRange(
    1,1,Math.max(values.length + 1,1),headers.length
  ).setWrap(true);

  const widths = [
    60,170,140,260,90,430,100,360,90,420,420,260,130
  ];
  widths.forEach((w,i) => sh.setColumnWidth(i+1,w));

  return {
    total: items.length,
    cannibal: items.filter(x => x.type === 'カニバリ疑い').length,
    newArticle: items.filter(x => x.type === '新規記事機会').length,
    gap: items.filter(x => x.type === 'コンテンツギャップ').length
  };
}

function sdsdBuildSiteOpportunityCases() {
  try {
    if (!sdsdQueryEvidenceSourceCount_()) {
      SpreadsheetApp.getUi().alert(
        'サイト横断診断案件を作成できません。\n\n' +
        '先にEvidence Packageを読み込んでください。'
      );
      return;
    }

    const items = sdsdDetectSiteOpportunities_();
    if (!items.length) {
      SpreadsheetApp.getUi().alert(
        'Doctorへ渡す横断診断案件はありません。'
      );
      return;
    }

    const result = sdsdWriteOpportunityCases_(items);
    const sh = SpreadsheetApp.getActive()
      .getSheetByName(SDSD_CONFIG.sheets.opportunityCases);
    if (sh) SpreadsheetApp.getActive().setActiveSheet(sh);

    SpreadsheetApp.getUi().alert(
      `サイト横断診断案件を作成しました。\n\n` +
      `合計: ${result.total}案件\n` +
      `カニバリ疑い: ${result.cannibal}案件\n` +
      `新規記事機会: ${result.newArticle}案件\n` +
      `コンテンツギャップ: ${result.gap}案件\n\n` +
      `各案件に固定の案件IDを付けました。\n` +
      `この段階では記事の統合・新規作成・リライトは行いません。`
    );
  } catch(e) {
    SpreadsheetApp.getUi().alert(
      `サイト横断診断案件を作成できませんでした。\n\n${e.message || e}`
    );
    throw e;
  }
}

function sdsdOpenSiteOpportunityCases() {
  const sh = SpreadsheetApp.getActive()
    .getSheetByName(SDSD_CONFIG.sheets.opportunityCases);
  if (sh) SpreadsheetApp.getActive().setActiveSheet(sh);
}
