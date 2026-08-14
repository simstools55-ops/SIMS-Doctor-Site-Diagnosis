/**
 * SIMS Doctor Site Diagnosis v0.5.0 Sprint 1.1
 * Query-wide opportunity diagnosis with intent clustering.
 *
 * Scope:
 * - Cannibalization suspicion
 * - New article opportunity
 * - Content gap
 *
 * Candidate detection only. No automatic treatment.
 */

function sdsdNormalizeQuery_(q) {
  return String(q || '')
    .toLowerCase()
    .replace(/[　\s]+/g, ' ')
    .replace(/[｜|／/・,，。!！?？:：;；()[\]【】「」『』]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function sdsdQueryTokens_(q) {
  return sdsdNormalizeQuery_(q)
    .split(' ')
    .map(x => x.trim())
    .filter(x => x && x.length >= 2)
    .filter(x => !/^(202[0-9]|20[0-9]{2}|最新版|最新|完全版|方法|やり方)$/.test(x));
}

function sdsdQuerySimilarity_(a, b) {
  const na = sdsdNormalizeQuery_(a);
  const nb = sdsdNormalizeQuery_(b);
  if (!na || !nb) return 0;
  if (na === nb) return 1;
  if (na.indexOf(nb) >= 0 || nb.indexOf(na) >= 0) return 0.85;

  const ta = sdsdQueryTokens_(na);
  const tb = sdsdQueryTokens_(nb);
  if (!ta.length || !tb.length) return 0;

  const sa = {};
  ta.forEach(x => sa[x] = true);
  let common = 0;
  tb.forEach(x => { if (sa[x]) common++; });

  const union = {};
  ta.concat(tb).forEach(x => union[x] = true);
  return common / Math.max(Object.keys(union).length, 1);
}

function sdsdBuildQueryUrlIndex_() {
  const rows = sdsdReadObjects_(SDSD_CONFIG.sheets.evidencePageQuery);
  const index = {};

  rows.forEach(r => {
    const url = sdsdNormalizeUrl_(
      sdsdObjectValue_(r, ['page','key','url','URL','記事URL'])
    );
    const query = String(
      sdsdObjectValue_(r, ['query','クエリ','検索クエリ']) || ''
    ).trim();
    const nq = sdsdNormalizeQuery_(query);
    if (!url || !nq) return;

    (index[nq] = index[nq] || []).push({
      query: query,
      url: url,
      clicks: Number(sdsdObjectValue_(r, ['clicks','クリック数']) || 0),
      impressions: Number(sdsdObjectValue_(r, ['impressions','表示回数']) || 0),
      position: Number(
        sdsdObjectValue_(r, ['position','掲載順位','平均掲載順位']) || 0
      )
    });
  });

  return index;
}

function sdsdArticleTitleText_(articleMap, url) {
  const master = articleMap[sdsdNormalizeUrl_(url)] || null;
  if (!master) return '';
  if (typeof master === 'string') return master;
  return String(master.title || master.articleTitle || '');
}

function sdsdTitleQueryCoverage_(title, query) {
  const nt = sdsdNormalizeQuery_(title);
  const nq = sdsdNormalizeQuery_(query);
  if (!nt || !nq) return 0;

  if (nt.indexOf(nq) >= 0) return 1;

  const terms = sdsdQueryTokens_(nq);
  if (!terms.length) {
    return nt.indexOf(nq) >= 0 ? 1 : 0;
  }

  let hit = 0;
  terms.forEach(t => {
    if (nt.indexOf(t) >= 0) hit++;
  });
  return hit / terms.length;
}

function sdsdDetectRawSiteOpportunities_() {
  const index = sdsdBuildQueryUrlIndex_();
  const articleMap = sdsdArticleTitleMap_();
  const out = [];

  Object.keys(index).forEach(nq => {
    const byUrl = {};

    index[nq].forEach(x => {
      const a = byUrl[x.url] || {
        url: x.url,
        query: x.query,
        clicks: 0,
        impressions: 0,
        weightedPosition: 0
      };
      a.clicks += x.clicks;
      a.impressions += x.impressions;
      a.weightedPosition += x.position * Math.max(x.impressions, 1);
      byUrl[x.url] = a;
    });

    const urls = Object.keys(byUrl)
      .map(u => {
        const a = byUrl[u];
        a.position = a.impressions
          ? a.weightedPosition / a.impressions
          : 0;
        return a;
      })
      .sort((a,b) => b.impressions - a.impressions);

    const totalImp = urls.reduce((s,x) => s + x.impressions, 0);
    if (totalImp < 10) return;

    // Cannibalization candidate.
    if (urls.length >= 2) {
      const first = urls[0];
      const second = urls[1];
      const secondShare = totalImp ? second.impressions / totalImp : 0;

      if (second.impressions >= 5 && secondShare >= 0.20) {
        out.push({
          type: 'カニバリ疑い',
          priority: totalImp >= 100 ? 1 : 2,
          theme: first.query,
          queries: [first.query],
          targets: [first.url, second.url],
          totalImpressions: totalImp,
          evidence:
            `同じ検索クエリで複数記事に表示が分散` +
            `（合計${Math.round(totalImp)}表示、` +
            `2番目の記事の比率${Math.round(secondShare*100)}%）`,
          confidence: totalImp >= 100 ? '高' : '中',
          action: 'Doctorで検索意図と記事本文を比較',
          destination: 'Doctor'
        });
      }
    }

    // New article vs content gap.
    const lead = urls[0];
    if (lead && totalImp >= 30 && lead.position >= 10) {
      const title = sdsdArticleTitleText_(articleMap, lead.url);
      const coverage = sdsdTitleQueryCoverage_(title, lead.query);

      if (coverage < 0.5) {
        out.push({
          type: '新規記事機会',
          priority: totalImp >= 100 ? 1 : 2,
          theme: lead.query,
          queries: [lead.query],
          targets: [lead.url],
          totalImpressions: totalImp,
          evidence:
            `検索需要${Math.round(totalImp)}表示に対し、` +
            `主な表示記事の順位は${lead.position.toFixed(1)}位。` +
            `記事タイトルとの一致も弱い`,
          confidence: totalImp >= 100 ? '高' : '中',
          action: '既存記事との検索意図重複を確認後、新記事候補として評価',
          destination: 'Creator'
        });
      } else {
        out.push({
          type: 'コンテンツギャップ',
          priority: totalImp >= 100 ? 1 : 2,
          theme: lead.query,
          queries: [lead.query],
          targets: [lead.url],
          totalImpressions: totalImp,
          evidence:
            `既存記事に関連する検索需要${Math.round(totalImp)}表示があるが、` +
            `主な掲載順位は${lead.position.toFixed(1)}位`,
          confidence: totalImp >= 100 ? '高' : '中',
          action: '既存記事が検索意図へ十分回答しているか確認',
          destination: 'Writer / Doctor'
        });
      }
    }
  });

  return out;
}

function sdsdSameTargetSet_(a, b) {
  const aa = (a || []).slice().sort();
  const bb = (b || []).slice().sort();
  if (aa.length !== bb.length) return false;
  for (let i=0; i<aa.length; i++) {
    if (aa[i] !== bb[i]) return false;
  }
  return true;
}

function sdsdCanClusterOpportunity_(cluster, item) {
  if (cluster.type !== item.type) return false;

  // Cannibalization must concern the same pair/set of URLs.
  if (item.type === 'カニバリ疑い') {
    if (!sdsdSameTargetSet_(cluster.targets, item.targets)) return false;
  } else {
    // New article / content gap must have the same lead URL.
    if (!cluster.targets.length || !item.targets.length ||
        cluster.targets[0] !== item.targets[0]) return false;
  }

  return sdsdQuerySimilarity_(cluster.theme, item.theme) >= 0.45;
}

function sdsdClusterSiteOpportunities_(raw) {
  const sorted = raw.slice().sort((a,b) =>
    a.priority - b.priority ||
    b.totalImpressions - a.totalImpressions
  );

  const clusters = [];

  sorted.forEach(item => {
    let target = null;

    for (let i=0; i<clusters.length; i++) {
      if (sdsdCanClusterOpportunity_(clusters[i], item)) {
        target = clusters[i];
        break;
      }
    }

    if (!target) {
      clusters.push({
        type: item.type,
        priority: item.priority,
        theme: item.theme,
        queries: item.queries.slice(),
        targets: item.targets.slice(),
        totalImpressions: item.totalImpressions,
        confidence: item.confidence,
        action: item.action,
        destination: item.destination,
        rawCount: 1
      });
      return;
    }

    target.rawCount++;
    target.totalImpressions += item.totalImpressions;
    item.queries.forEach(q => {
      if (target.queries.indexOf(q) < 0) target.queries.push(q);
    });

    if (item.priority < target.priority) target.priority = item.priority;
    if (item.confidence === '高') target.confidence = '高';

    // Keep the shortest query as the representative theme in most cases.
    if (String(item.theme).length < String(target.theme).length) {
      target.theme = item.theme;
    }
  });

  clusters.forEach(c => {
    const queryPreview = c.queries.slice(0, 5).join(' / ');
    const more = c.queries.length > 5
      ? ` ほか${c.queries.length - 5}件`
      : '';

    if (c.type === 'カニバリ疑い') {
      c.evidence =
        `近い検索意図の${c.queries.length}クエリを1案件に統合。` +
        `対象記事${c.targets.length}本、合計${Math.round(c.totalImpressions)}表示。` +
        ` 主なクエリ: ${queryPreview}${more}`;
    } else if (c.type === '新規記事機会') {
      c.evidence =
        `近い検索意図の${c.queries.length}クエリを1テーマに統合。` +
        `合計${Math.round(c.totalImpressions)}表示。` +
        ` 主なクエリ: ${queryPreview}${more}`;
    } else {
      c.evidence =
        `既存記事に関連する${c.queries.length}クエリを1テーマに統合。` +
        `合計${Math.round(c.totalImpressions)}表示。` +
        ` 主なクエリ: ${queryPreview}${more}`;
    }
  });

  return clusters.sort((a,b) =>
    a.priority - b.priority ||
    a.type.localeCompare(b.type,'ja') ||
    b.totalImpressions - a.totalImpressions
  );
}

function sdsdDetectSiteOpportunities_() {
  return sdsdClusterSiteOpportunities_(
    sdsdDetectRawSiteOpportunities_()
  );
}

function sdsdWriteSiteOpportunities_(rows) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SDSD_CONFIG.sheets.opportunities);
  if (!sh) sh = ss.insertSheet(SDSD_CONFIG.sheets.opportunities);
  sh.clear();

  const titleMap = sdsdArticleTitleMap_();
  const headers = [
    '優先順位','改善テーマ','検索テーマ','関連クエリ数',
    '対象記事','根拠','確信度','推奨対応','担当'
  ];

  const values = rows.map((x,i) => [
    i+1,
    x.type,
    x.theme,
    x.queries.length,
    x.targets.map(u =>
      sdsdDisplayTitle_(u, titleMap) + '\n' + u
    ).join('\n\n'),
    x.evidence,
    x.confidence,
    x.action,
    x.destination
  ]);

  sh.getRange(1,1,1,headers.length).setValues([headers]);

  if (values.length) {
    sh.getRange(2,1,values.length,headers.length).setValues(values);
  }

  sh.setFrozenRows(1);
  sh.getRange(1,1,1,headers.length).setFontWeight('bold');
  sh.getRange(
    1,1,Math.max(1,values.length+1),headers.length
  ).setWrap(true);

  [70,150,240,100,430,500,90,360,120]
    .forEach((w,i) => sh.setColumnWidth(i+1,w));

  if (!values.length) {
    sh.getRange(2,1,1,headers.length).merge();
    sh.getRange(2,1).setValue(
      '現在のEvidenceから、中以上の確信度で提示できる横断改善候補は見つかりませんでした。'
    );
  }
}

function sdsdRunSiteOpportunityDiagnosis() {
  try {
    if (!sdsdQueryEvidenceSourceCount_()) {
      SpreadsheetApp.getUi().alert(
        'サイト横断診断を実行できません。\n\n' +
        'page_query_top のEvidenceがありません。'
      );
      return;
    }

    const raw = sdsdDetectRawSiteOpportunities_();
    const rows = sdsdClusterSiteOpportunities_(raw);
    sdsdWriteSiteOpportunities_(rows);

    const rawCounts = {};
    raw.forEach(x => rawCounts[x.type] = (rawCounts[x.type] || 0) + 1);

    const counts = {};
    rows.forEach(x => counts[x.type] = (counts[x.type] || 0) + 1);

    const sh = SpreadsheetApp.getActive()
      .getSheetByName(SDSD_CONFIG.sheets.opportunities);
    if (sh) SpreadsheetApp.getActive().setActiveSheet(sh);

    SpreadsheetApp.getUi().alert(
      `サイト横断診断が完了しました。\n\n` +
      `カニバリ疑い: ${counts['カニバリ疑い']||0}案件` +
      `（元候補 ${rawCounts['カニバリ疑い']||0}件）\n` +
      `新規記事機会: ${counts['新規記事機会']||0}テーマ` +
      `（元候補 ${rawCounts['新規記事機会']||0}件）\n` +
      `コンテンツギャップ: ${counts['コンテンツギャップ']||0}テーマ` +
      `（元候補 ${rawCounts['コンテンツギャップ']||0}件）\n\n` +
      `近い検索意図は1案件へまとめています。\n` +
      `これは一次候補であり、自動処置は行いません。`
    );
  } catch(e) {
    SpreadsheetApp.getUi().alert(
      `サイト横断診断を完了できませんでした。\n\n${e.message || e}`
    );
    throw e;
  }
}
