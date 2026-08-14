/**
 * SIMS Doctor Site Diagnosis v0.5.0 Sprint 5
 * Build precision-diagnosis packages for Doctor-prioritized site-wide clusters.
 *
 * Scope:
 * - Only canonical site-wide result cases with:
 *   doctor_decision = ADDITIONAL_EVIDENCE_REQUIRED
 *   route_to = NEEDS_EVIDENCE
 * - Includes article body + page summary + weekly trend rows + query evidence.
 * - One ZIP contains all priority clusters, preserving site-wide context.
 */

function sdsdReadStoredSiteWideResult_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SDSD_CONFIG.sheets.siteWideResult);
  if (!sh || !sh.getRange('A1').getValue()) {
    throw new Error('Doctor横断診断結果が登録されていません。先に「12. Doctor横断診断結果を登録」を実行してください。');
  }
  const raw = String(sh.getRange('A1').getValue() || '');
  const obj = JSON.parse(raw);
  if (!obj || obj.format !== 'SIMS_DOCTOR_SITE_WIDE_RESULT_V1') {
    throw new Error('保存済みDoctor横断診断結果の形式が不正です。');
  }
  return obj;
}

function sdsdPrecisionClusterCases_() {
  const obj = sdsdReadStoredSiteWideResult_();
  return (obj.diagnosis_cases || [])
    .filter(c =>
      String(c.route_to || '') === 'NEEDS_EVIDENCE' &&
      String(c.doctor_decision || '') === 'ADDITIONAL_EVIDENCE_REQUIRED'
    )
    .slice(0, 5);
}

function sdsdEvidenceRowsForUrl_(sheetName, url) {
  const rows = sdsdReadObjects_(sheetName);
  const n = sdsdNormalizeUrl_(url);
  return rows.filter(r => {
    const raw = sdsdObjectValue_(r, ['page','key','url','URL','記事URL']);
    return sdsdNormalizeUrl_(raw) === n;
  });
}

function sdsdPrecisionArticleEvidence_(article) {
  const url = String(article.article_url || article.url || '');
  if (!url) throw new Error('精密診断対象記事にURLがありません。');

  const fetched = sdsdFetchArticleEvidence_(url);
  if (fetched.status !== 'VALID') {
    throw new Error(`記事本文取得失敗: ${article.article_id || ''} / ${url} / ${fetched.status}`);
  }

  const summaryRows = sdsdEvidenceRowsForUrl_(SDSD_CONFIG.sheets.evidencePageSummary, url);
  const weeklyRows = sdsdEvidenceRowsForUrl_(SDSD_CONFIG.sheets.evidencePageWeekly, url);
  const queryRows = (sdsdBuildQueryEvidenceMap_()[sdsdNormalizeUrl_(url)] || []).slice(0, 20);

  return {
    identity: {
      site_id: String(article.site_id || ''),
      article_id: String(article.article_id || ''),
      article_title: String(article.article_title || fetched.title || ''),
      article_url: url,
      main_query: String(article.main_query || '')
    },
    article_meta: {
      page_title: fetched.title,
      meta_description: fetched.metaDescription,
      canonical_url: fetched.canonicalUrl || url,
      fetched_at: new Date().toISOString()
    },
    search_console: {
      page_summary: summaryRows,
      page_weekly: weeklyRows,
      top_queries: queryRows
    },
    article_html: fetched.articleHtml
  };
}

function sdsdPrecisionReferralText_(siteResult, cases) {
  const site = siteResult.site || {};
  return [
    '# SIMS Doctor Site-wide Precision Diagnosis Referral',
    '',
    `Site ID: ${site.site_id || ''}`,
    `Site Name: ${site.site_name || ''}`,
    `Site URL: ${site.site_url || ''}`,
    `Site Diagnosis Batch ID: ${siteResult.site_diagnosis_batch_id || ''}`,
    '',
    `今回の精密診断クラスタ: ${cases.length}件`,
    '',
    '## 重要',
    '',
    '- これはSite-wide一次トリアージ後の追加Evidence精密診断です。',
    '- 各クラスタについて、本文全文とSearch Console個別データを読み、維持 / 役割分担 / Writer / Merge / Monitor を確定してください。',
    '- Creatorは今回の精密診断対象ではありません。',
    '- カニバリだから統合、という前提で判断しないでください。',
    '- 記事本文・検索意図・GSC推移が役割分担を支持する場合は、統合せず維持してください。',
    '',
    '## 返却',
    '',
    '- 元の site_diagnosis_batch_id と diagnosis_case_id を保持してください。',
    '- 各クラスタの最終 route_to は WRITER / MERGE / MONITOR / NO_ACTION / NEEDS_EVIDENCE のいずれか。',
    '- 不足Evidenceが残る場合だけ NEEDS_EVIDENCE としてください。',
    ''
  ].join('\n');
}

function sdsdExportPriorityPrecisionClusterPackage() {
  const ui = SpreadsheetApp.getUi();

  try {
    const result = sdsdReadStoredSiteWideResult_();
    const cases = sdsdPrecisionClusterCases_();

    if (!cases.length) {
      ui.alert('追加Evidence精密診断の優先クラスタはありません。');
      return;
    }

    const blobs = [];
    const manifest = {
      format: 'SIMS_DOCTOR_SITE_WIDE_PRECISION_PACKAGE_V1',
      contract_version: '1.0',
      generated_at: new Date().toISOString(),
      site_diagnosis_batch_id: String(result.site_diagnosis_batch_id || ''),
      site: result.site || {},
      source_result_format: result.format,
      precision_cluster_count: cases.length,
      clusters: []
    };

    blobs.push(Utilities.newBlob(
      sdsdPrecisionReferralText_(result, cases),
      'text/markdown',
      'DOCTOR-PRECISION-REFERRAL.md'
    ));

    cases.forEach((c, ci) => {
      const clusterFolder =
        `clusters/${String(ci+1).padStart(2,'0')}-${String(c.diagnosis_case_id || '').replace(/[^A-Za-z0-9._-]/g,'_')}`;

      const clusterManifest = {
        diagnosis_case_id: c.diagnosis_case_id,
        diagnosis_theme: c.diagnosis_theme,
        diagnosis_type: c.diagnosis_type,
        absorbed_source_case_ids: c.absorbed_source_case_ids || [],
        doctor_decision: c.doctor_decision,
        confidence: c.confidence,
        site_impact: c.site_impact,
        treatment_strategy: c.treatment_strategy,
        route_to: c.route_to,
        eventual_route: c.eventual_route,
        reason: c.reason,
        additional_evidence_needed: c.additional_evidence_needed || [],
        articles: []
      };

      (c.target_articles || []).forEach((article, ai) => {
        const ev = sdsdPrecisionArticleEvidence_(article);
        const safeArticleId = String(
          ev.identity.article_id || `ARTICLE-${ai+1}`
        ).replace(/[^A-Za-z0-9._-]/g,'_');

        const articleFolder =
          `${clusterFolder}/articles/${String(ai+1).padStart(2,'0')}-${safeArticleId}`;

        const articleJson = {
          format: 'SIMS_DOCTOR_SITE_WIDE_PRECISION_ARTICLE_V1',
          contract_version: '1.0',
          diagnosis_case_id: c.diagnosis_case_id,
          identity: ev.identity,
          article_meta: ev.article_meta,
          search_console: ev.search_console
        };

        blobs.push(Utilities.newBlob(
          JSON.stringify(articleJson, null, 2),
          'application/json',
          `${articleFolder}/evidence.json`
        ));

        blobs.push(Utilities.newBlob(
          ev.article_html,
          'text/html',
          `${articleFolder}/article.html`
        ));

        clusterManifest.articles.push({
          identity: ev.identity,
          article_meta: ev.article_meta,
          evidence_file: 'evidence.json',
          article_file: 'article.html'
        });
      });

      blobs.push(Utilities.newBlob(
        JSON.stringify(clusterManifest, null, 2),
        'application/json',
        `${clusterFolder}/cluster.json`
      ));

      manifest.clusters.push({
        diagnosis_case_id: c.diagnosis_case_id,
        diagnosis_theme: c.diagnosis_theme,
        target_article_count: (c.target_articles || []).length,
        absorbed_source_case_ids: c.absorbed_source_case_ids || []
      });
    });

    blobs.unshift(Utilities.newBlob(
      JSON.stringify(manifest, null, 2),
      'application/json',
      'manifest.json'
    ));

    const zipName =
      `SIMS-Doctor-Site-Wide-Precision-${Utilities.formatDate(
        new Date(),
        Session.getScriptTimeZone() || 'Asia/Tokyo',
        'yyyyMMdd-HHmmss'
      )}.zip`;

    const file = DriveApp.createFile(Utilities.zip(blobs, zipName));

    ui.alert(
      `サイト横断の精密診断Packageを生成しました。\n\n` +
      `優先クラスタ: ${cases.length}件\n` +
      `保存先: Google Drive\n\n` +
      `${file.getUrl()}\n\n` +
      `このZIPをSIMS Doctorへ渡してください。`
    );

    return {
      clusterCount: cases.length,
      fileId: file.getId(),
      fileUrl: file.getUrl()
    };

  } catch(e) {
    ui.alert(
      `サイト横断の精密診断Packageを生成できませんでした。\n\n` +
      `${e.message || e}`
    );
    throw e;
  }
}
