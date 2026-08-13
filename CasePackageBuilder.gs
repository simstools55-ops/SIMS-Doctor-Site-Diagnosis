function sdsdArticleCacheKey_(articleId, url) {
  const raw = String(articleId || '') + '|' + String(url || '');
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw);
  return 'SDSD_ARTICLE_' + digest.map(b => ('0' + ((b + 256) % 256).toString(16)).slice(-2)).join('');
}

function sdsdEnrichSelectedCases() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SDSD_CONFIG.sheets.selectedCases);
  if (!sh) throw new Error('先に Treatment Batch を生成してください。');

  const values = sh.getDataRange().getValues();
  if (values.length < 2) throw new Error('Selected Treatment Cases に案件がありません。');

  const headers = values[0].map(String);
  const idx = {};
  headers.forEach((h,i) => idx[h] = i);

  const articleMap = sdsdBuildArticleMasterMap_();
  if (!Object.keys(articleMap).length) {
    throw new Error('記事管理データがありません。_SDSD_ARTICLE_MASTERへSBM「記事管理」CSVを入れてください。');
  }

  const extraHeaders = [
    'ArticleID','Article Title','Main Query','Article Fetch Status',
    'Case Package Status','Article Cache Key'
  ];
  let lastCol = headers.length;
  extraHeaders.forEach(h => {
    if (idx[h] == null) {
      lastCol++;
      sh.getRange(1,lastCol).setValue(h);
      idx[h] = lastCol - 1;
      headers.push(h);
    }
  });

  const cache = CacheService.getDocumentCache();
  let ready = 0;
  let failed = 0;

  for (let r=1; r<values.length; r++) {
    const url = String(values[r][idx['URL']] || '');
    if (!url) continue;

    const master = articleMap[sdsdNormalizeUrl_(url)] || null;
    const articleId = master ? master.articleId : '';
    const title = master ? master.title : '';
    const mainQuery = master ? master.mainQuery : '';
    const fetched = sdsdFetchArticleEvidence_(url);

    sh.getRange(r+1, idx['ArticleID']+1).setValue(articleId);
    sh.getRange(r+1, idx['Article Title']+1).setValue(title || fetched.title);
    sh.getRange(r+1, idx['Main Query']+1).setValue(mainQuery);
    sh.getRange(r+1, idx['Article Fetch Status']+1).setValue(fetched.status);

    if (!master || !articleId || fetched.status !== 'VALID') {
      sh.getRange(r+1, idx['Case Package Status']+1).setValue('NEEDS_REVIEW');
      sh.getRange(r+1, idx['Referral Status']+1).setValue('NEEDS_CASE_ENRICHMENT_REVIEW');
      failed++;
      continue;
    }

    const cacheKey = sdsdArticleCacheKey_(articleId, url);
    // Cache value limit is ~100 KB. If too large, export step will refetch it.
    let cached = false;
    try {
      if (fetched.articleHtml.length < 90000) {
        cache.put(cacheKey, fetched.articleHtml, 21600);
        cached = true;
      }
    } catch(e) {}

    const oldReferral = String(values[r][idx['Referral JSON']] || '{}');
    let referral = {};
    try { referral = JSON.parse(oldReferral); } catch(e) {}

    referral.format = 'SIMS_DOCTOR_INDIVIDUAL_CASE_PACKAGE_V1';
    referral.contract_version = '1.0';
    referral.case_identity = referral.case_identity || {};
    referral.case_identity.article_id = articleId;
    referral.case_identity.url = url;

    // IMPORTANT: article_html is deliberately NOT stored in the sheet.
    referral.article_evidence = {
      status: 'VALID',
      title: title || fetched.title,
      page_title: fetched.title,
      meta_description: fetched.metaDescription,
      canonical_url: fetched.canonicalUrl || url,
      main_query: mainQuery,
      body_storage: cached ? 'DOCUMENT_CACHE' : 'REFETCH_ON_EXPORT',
      fetched_at: new Date().toISOString()
    };

    referral.required_examinations = [
      '記事本文全文を読み、Site Referralの仮説を独立に検証する',
      '主要クエリと現在SERPの検索意図を確認する',
      '公式一次情報で現行仕様・鮮度を確認する',
      'カニバリ・内部リンク・外部環境要因を確認する',
      '処方前にRecent Treatment Guardを再確認する'
    ];
    referral.treatment_constraints = [
      '全面リライトを前提としない',
      '原因確定前にURL変更・大規模構成変更を決定しない',
      '既存の有効な独自情報・広告・アフィリエイト要素は保護対象として評価する'
    ];

    sh.getRange(r+1, idx['Referral JSON']+1).setValue(JSON.stringify(referral));
    sh.getRange(r+1, idx['Article Cache Key']+1).setValue(cacheKey);
    sh.getRange(r+1, idx['Referral Status']+1).setValue('READY_FOR_INDIVIDUAL_DOCTOR');
    sh.getRange(r+1, idx['Case Package Status']+1).setValue('READY');
    ready++;
  }

  SpreadsheetApp.getUi().alert(
    `Case Enrichment完了\nReady: ${ready}件\n要確認: ${failed}件\n\n` +
    `記事本文はセルへ保存せず、ZIP生成時に安全に添付します。`
  );
}

function sdsdExportDoctorCasePackageZip() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SDSD_CONFIG.sheets.selectedCases);
  if (!sh) throw new Error('Selected Treatment Cases がありません。');

  const values = sh.getDataRange().getValues();
  const headers = values[0].map(String);
  const idx = {};
  headers.forEach((h,i)=>idx[h]=i);

  const rows = values.slice(1).filter(r => r[idx['URL']]);
  const notReady = rows.filter(r => String(r[idx['Referral Status']] || '') !== 'READY_FOR_INDIVIDUAL_DOCTOR');
  if (notReady.length) {
    throw new Error(`READYではない案件が${notReady.length}件あります。先に Case Enrichment を完了してください。`);
  }

  const cache = CacheService.getDocumentCache();
  const blobs = [];
  const manifestCases = [];
  let refetched = 0;

  rows.forEach((r,i) => {
    const articleId = String(r[idx['ArticleID']] || `CASE-${i+1}`);
    const url = String(r[idx['URL']] || '');
    const jsonText = String(r[idx['Referral JSON']] || '{}');
    const safeId = articleId.replace(/[^A-Za-z0-9._-]/g,'_');
    const folder = `cases/${String(i+1).padStart(2,'0')}-${safeId}`;

    let parsed = {};
    try { parsed = JSON.parse(jsonText); } catch(e) {}

    let html = '';
    const cacheKey = idx['Article Cache Key'] != null
      ? String(r[idx['Article Cache Key']] || '') : '';

    if (cacheKey) {
      try { html = cache.get(cacheKey) || ''; } catch(e) {}
    }

    if (!html) {
      const fetched = sdsdFetchArticleEvidence_(url);
      if (fetched.status !== 'VALID') {
        throw new Error(`記事本文の再取得に失敗しました: ${articleId} / ${url} / ${fetched.status}`);
      }
      html = fetched.articleHtml;
      refetched++;
    }

    // ZIP case.json stays lightweight; article body is a separate file.
    if (parsed.article_evidence) {
      parsed.article_evidence.body_storage = 'PACKAGE_FILE';
      parsed.article_evidence.body_file = 'article.html';
    }

    blobs.push(Utilities.newBlob(
      JSON.stringify(parsed,null,2),
      'application/json',
      `${folder}/case.json`
    ));
    blobs.push(Utilities.newBlob(
      html,
      'text/html',
      `${folder}/article.html`
    ));

    manifestCases.push({
      order: i+1,
      article_id: articleId,
      url: url,
      priority: String(r[idx['Site Priority']] || ''),
      tvs: Number(r[idx['TVS']] || 0),
      case_file: `${folder}/case.json`,
      article_file: `${folder}/article.html`
    });
  });

  const manifest = {
    format: 'SIMS_DOCTOR_SITE_TREATMENT_BATCH_V1',
    contract_version: '1.0',
    generated_at: new Date().toISOString(),
    case_count: rows.length,
    cases: manifestCases
  };
  blobs.push(Utilities.newBlob(
    JSON.stringify(manifest,null,2),
    'application/json',
    'manifest.json'
  ));

  const zipName =
    `SIMS-Doctor-Site-Treatment-Batch-${Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone() || 'Asia/Tokyo',
      'yyyyMMdd-HHmmss'
    )}.zip`;

  const file = DriveApp.createFile(Utilities.zip(blobs, zipName));

  SpreadsheetApp.getUi().alert(
    `Doctor Case Package ZIP生成完了\n\n` +
    `案件数: ${rows.length}\n` +
    `本文再取得: ${refetched}件\n` +
    `Last error: (なし)\n\n` +
    `保存先:\n${file.getUrl()}`
  );
}
