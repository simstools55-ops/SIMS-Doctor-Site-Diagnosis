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



function sdsdDiagnosisThemeKey_(item, articleMap) {
  const type = String(item.type || '');
  const theme = String(item.parentTheme || item.theme || '');
  const targets = item.targets || [];

  // For site-wide diagnosis, article titles give a stronger semantic anchor
  // than raw error-code queries alone.
  const titleTokens = [];
  targets.forEach(url => {
    const title = sdsdArticleTitleText_(articleMap, url);
    sdsdParentThemeTokens_(title).forEach(t => {
      if (titleTokens.indexOf(t) < 0) titleTokens.push(t);
    });
  });

  const queryTokens = sdsdParentThemeTokens_(theme);

  // Prefer product/service/topic tokens shared by article titles and query themes.
  const shared = queryTokens.filter(t => titleTokens.indexOf(t) >= 0);

  let base = '';
  if (shared.length) {
    base = shared.slice(0,3).join(' ');
  } else if (titleTokens.length) {
    base = titleTokens.slice(0,3).join(' ');
  } else {
    base = queryTokens.slice(0,3).join(' ');
  }

  // Preserve type boundary. A cannibalization case should not merge into a
  // Creator opportunity even if the topic is identical.
  return type + '|' + (base || sdsdNormalizeQuery_(theme));
}

function sdsdDiagnosisThemeSimilarity_(a, b) {
  const ka = String(a || '').split('|').slice(1).join('|');
  const kb = String(b || '').split('|').slice(1).join('|');
  if (!ka || !kb) return 0;
  if (ka === kb) return 1;
  if (ka.indexOf(kb) >= 0 || kb.indexOf(ka) >= 0) return 0.9;
  return sdsdQuerySimilarity_(ka, kb);
}

function sdsdCanMergeDiagnosisCases_(cluster, item, articleMap) {
  if (cluster.type !== item.type) return false;

  const itemKey = sdsdDiagnosisThemeKey_(item, articleMap);
  const themeSim = sdsdDiagnosisThemeSimilarity_(cluster.diagnosisThemeKey, itemKey);

  if (item.type === 'カニバリ疑い') {
    const overlap = sdsdTargetOverlap_(cluster.targets, item.targets);
    // Diagnosis cases may merge when topic is very close and at least one
    // article overlaps, or when the topic is nearly identical.
    return (themeSim >= 0.65 && overlap > 0) || themeSim >= 0.9;
  }

  // Creator / content-gap cases: same diagnosis topic plus either same lead
  // article or very strong semantic identity.
  const sameLead =
    cluster.targets.length && item.targets.length &&
    cluster.targets[0] === item.targets[0];

  return (themeSim >= 0.7 && sameLead) || themeSim >= 0.9;
}

function sdsdBuildDiagnosisThemeLabel_(cluster) {
  const raw = String(cluster.diagnosisThemeKey || '');
  const label = raw.split('|').slice(1).join('|').trim();
  return label || cluster.parentTheme || cluster.theme;
}

function sdsdParentThemeTokens_(q) {
  const stop = {
    'エラー':true,'エラーコード':true,'error':true,'コード':true,
    'できない':true,'見れない':true,'見られない':true,'直し方':true,
    '対処法':true,'原因':true,'方法':true,'やり方':true,'設定':true,
    '最新版':true,'最新':true,'完全版':true,'解決':true
  };

  return sdsdQueryTokens_(q).filter(t => {
    if (stop[t]) return false;
    if (/^[0-9]{5,}$/.test(t)) return false;
    if (/^[a-z]{1,5}-?[a-z0-9]{1,8}$/i.test(t) && /\d/.test(t)) return false;
    return true;
  });
}

function sdsdParentThemeKey_(q) {
  const toks = sdsdParentThemeTokens_(q);
  if (!toks.length) return sdsdNormalizeQuery_(q);
  return toks.slice(0, 3).join(' ');
}

function sdsdParentThemeSimilarity_(a, b) {
  const ka = sdsdParentThemeKey_(a);
  const kb = sdsdParentThemeKey_(b);
  if (!ka || !kb) return 0;
  if (ka === kb) return 1;
  if (ka.indexOf(kb) >= 0 || kb.indexOf(ka) >= 0) return 0.9;
  return sdsdQuerySimilarity_(ka, kb);
}

function sdsdTargetOverlap_(a, b) {
  const aa = {};
  (a || []).forEach(x => aa[x] = true);
  let common = 0;
  (b || []).forEach(x => { if (aa[x]) common++; });
  const denom = Math.max(Math.min((a||[]).length,(b||[]).length),1);
  return common / denom;
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

  if (item.type === 'カニバリ疑い') {
    // Stage 1: same target set or strong overlap + same parent theme.
    const exact = sdsdSameTargetSet_(cluster.targets, item.targets);
    const overlap = sdsdTargetOverlap_(cluster.targets, item.targets);
    const parentSim = sdsdParentThemeSimilarity_(cluster.theme, item.theme);
    return exact || (overlap >= 0.5 && parentSim >= 0.6);
  }

  // New article / content gap: same lead URL and same parent theme.
  if (!cluster.targets.length || !item.targets.length ||
      cluster.targets[0] !== item.targets[0]) return false;

  return sdsdParentThemeSimilarity_(cluster.theme, item.theme) >= 0.6;
}

function sdsdClusterSiteOpportunities_(raw) {
  const sorted = raw.slice().sort((a,b) =>
    a.priority - b.priority ||
    b.totalImpressions - a.totalImpressions
  );

  // Stage 1: query-intent / parent-theme clustering.
  const stage1 = [];

  sorted.forEach(item => {
    let target = null;

    for (let i=0; i<stage1.length; i++) {
      if (sdsdCanClusterOpportunity_(stage1[i], item)) {
        target = stage1[i];
        break;
      }
    }

    if (!target) {
      stage1.push({
        type: item.type,
        priority: item.priority,
        theme: item.theme,
        parentTheme: sdsdParentThemeKey_(item.theme),
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
    item.targets.forEach(u => {
      if (target.targets.indexOf(u) < 0) target.targets.push(u);
    });

    if (item.priority < target.priority) target.priority = item.priority;
    if (item.confidence === '高') target.confidence = '高';

    const key = sdsdParentThemeKey_(item.theme);
    if (key && key.length < target.parentTheme.length) {
      target.parentTheme = key;
    }

    if (String(item.theme).length < String(target.theme).length) {
      target.theme = item.theme;
    }
  });

  // Stage 2: site-wide diagnosis-case clustering.
  const articleMap = sdsdArticleTitleMap_();
  const stage2 = [];

  stage1.forEach(item => {
    item.diagnosisThemeKey = sdsdDiagnosisThemeKey_(item, articleMap);

    let target = null;
    for (let i=0; i<stage2.length; i++) {
      if (sdsdCanMergeDiagnosisCases_(stage2[i], item, articleMap)) {
        target = stage2[i];
        break;
      }
    }

    if (!target) {
      stage2.push({
        type: item.type,
        priority: item.priority,
        theme: item.theme,
        parentTheme: item.parentTheme,
        diagnosisThemeKey: item.diagnosisThemeKey,
        queries: item.queries.slice(),
        targets: item.targets.slice(),
        totalImpressions: item.totalImpressions,
        confidence: item.confidence,
        action: item.action,
        destination: item.destination,
        rawCount: item.rawCount,
        parentCaseCount: 1
      });
      return;
    }

    target.parentCaseCount++;
    target.rawCount += item.rawCount;
    target.totalImpressions += item.totalImpressions;

    item.queries.forEach(q => {
      if (target.queries.indexOf(q) < 0) target.queries.push(q);
    });
    item.targets.forEach(u => {
      if (target.targets.indexOf(u) < 0) target.targets.push(u);
    });

    if (item.priority < target.priority) target.priority = item.priority;
    if (item.confidence === '高') target.confidence = '高';

    if (String(item.theme).length < String(target.theme).length) {
      target.theme = item.theme;
    }
  });

  stage2.forEach(c => {
    c.diagnosisTheme = sdsdBuildDiagnosisThemeLabel_(c);

    const preview = c.queries.slice(0,5).join(' / ');
    const more = c.queries.length > 5
      ? ` ほか${c.queries.length-5}件`
      : '';

    if (c.type === 'カニバリ疑い') {
      c.evidence =
        `診断テーマ「${c.diagnosisTheme}」として` +
        `${c.parentCaseCount}個の親テーマ案件、${c.queries.length}クエリを統合。` +
        `対象記事${c.targets.length}本、合計${Math.round(c.totalImpressions)}表示。` +
        ` 主なクエリ: ${preview}${more}`;
    } else if (c.type === '新規記事機会') {
      c.evidence =
        `診断テーマ「${c.diagnosisTheme}」として` +
        `${c.parentCaseCount}個の親テーマ、${c.queries.length}クエリを統合。` +
        `合計${Math.round(c.totalImpressions)}表示。` +
        ` 主なクエリ: ${preview}${more}`;
    } else {
      c.evidence =
        `診断テーマ「${c.diagnosisTheme}」として` +
        `${c.parentCaseCount}個の親テーマ、${c.queries.length}クエリを統合。` +
        `合計${Math.round(c.totalImpressions)}表示。` +
        ` 主なクエリ: ${preview}${more}`;
    }
  });

  return stage2.sort((a,b) =>
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
    '優先順位','改善テーマ','診断テーマ','親テーマ','検索テーマ','関連クエリ数',
    '対象記事','根拠','確信度','推奨対応','担当'
  ];

  const values = rows.map((x,i) => [
    i+1,
    x.type,
    x.diagnosisTheme || x.parentTheme || x.theme,
    x.parentTheme || x.theme,
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

  [70,150,220,220,240,100,430,520,90,360,120]
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
      `検索意図 → 親テーマ → 診断テーマの3段階で、Doctorが一度に判断すべき案件へまとめています。\n` +
      `これは一次候補であり、自動処置は行いません。`
    );
  } catch(e) {
    SpreadsheetApp.getUi().alert(
      `サイト横断診断を完了できませんでした。\n\n${e.message || e}`
    );
    throw e;
  }
}
