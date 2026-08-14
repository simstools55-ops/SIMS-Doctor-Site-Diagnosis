/**
 * SIMS Doctor Site Diagnosis RC5 HF3
 * Create a Merge referral from the selected MERGE row in "サイト治療計画".
 *
 * Doctor's merge direction is authoritative.
 * This module does NOT modify articles, delete URLs, or configure redirects.
 */

function sdsdMergeReferralSheetName_() {
  return 'Merge紹介状';
}

function sdsdParseArticleLabel_(text) {
  const s = String(text || '').trim();
  const parts = s.split('/').map(x => String(x || '').trim()).filter(Boolean);

  // URLs themselves contain "/", so parse ArticleID first and URL independently.
  const articleIdMatch = s.match(/\bA\d{6}\b/i);
  const urlMatch = s.match(/https?:\/\/[^\s]+/i);

  return {
    article_id: articleIdMatch ? articleIdMatch[0] : '',
    article_url: urlMatch ? urlMatch[0].replace(/[),.;、。]+$/g, '') : '',
    raw: s
  };
}

function sdsdSelectedTreatmentPlanMerge_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getActiveSheet();

  if (!sh || sh.getName() !== SDSD_CONFIG.sheets.treatmentPlan) {
    throw new Error('「サイト治療計画」シートでMERGE案件の行を選択してください。');
  }

  const row = sh.getActiveRange().getRow();
  if (row <= 1) {
    throw new Error('見出しではなく、MERGE案件の行を選択してください。');
  }

  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(1,1,1,lastCol).getDisplayValues()[0].map(String);
  const values = sh.getRange(row,1,1,lastCol).getDisplayValues()[0];
  const idx = {};
  headers.forEach((h,i) => idx[String(h || '').trim()] = i);

  function value(name) {
    return idx[name] == null ? '' : String(values[idx[name]] || '').trim();
  }

  const decision = value('Doctor判断').toUpperCase();
  if (decision !== 'MERGE') {
    throw new Error(`選択行はMERGE案件ではありません。現在: ${decision || '空欄'}`);
  }

  const survivorText = value('統合先（残す記事）');
  const absorbedText = value('統合元（吸収する記事）');

  if (!survivorText || !absorbedText) {
    throw new Error(
      '統合先または統合元が空欄です。先にPrecision Resultを再登録し、HF2のMerge計画を反映してください。'
    );
  }

  return {
    row: row,
    diagnosis_theme: value('診断テーマ'),
    doctor_decision: decision,
    next_action: value('次の処置'),
    target_articles_text: value('対象記事'),
    survivor: sdsdParseArticleLabel_(survivorText),
    absorbed: sdsdParseArticleLabel_(absorbedText),
    merge_direction: value('統合方向'),
    confidence: value('確信度'),
    reason: value('理由'),
    absorbed_source_case_ids_text: value('吸収した元案件'),
    diagnosis_case_id: value('診断案件ID'),
    parent_diagnosis_case_id: value('元診断案件ID')
  };
}

function sdsdFindStoredMergeCase_(selected) {
  // HF4:
  // The "サイト治療計画" row is already the registered Doctor decision.
  // Do not re-open or re-validate stored Doctor JSON here.
  // This removes the fragile dependency on internal storage format.
  return {
    site_result: null,
    diagnosis_case: {
      diagnosis_case_id: '',
      diagnosis_theme: String(selected.diagnosis_theme || ''),
      route_to: 'MERGE',
      confidence: String(selected.confidence || ''),
      reason: String(selected.reason || ''),
      treatment_strategy: '',
      merge_survivor: String(selected.survivor.raw || ''),
      merge_absorbed: String(selected.absorbed.raw || ''),
      merge_direction: String(selected.merge_direction || ''),
      merge_content_to_absorb: ''
    }
  };
}


function sdsdMergeReferralSiteMeta_(selected) {
  const survivorUrl = String(
    selected && selected.survivor ? selected.survivor.article_url : ''
  );
  const absorbedUrl = String(
    selected && selected.absorbed ? selected.absorbed.article_url : ''
  );

  let meta = {site_id:'', site_name:'', site_url:''};

  try {
    if (typeof sdsdSiteMetaFromArticleMaster_ === 'function') {
      const m = sdsdSiteMetaFromArticleMaster_();
      if (m) {
        meta.site_id = String(m.site_id || '');
        meta.site_name = String(m.site_name || '');
        meta.site_url = String(m.site_url || '');
      }
    }
  } catch(e) {}

  const anyUrl = survivorUrl || absorbedUrl;
  if (!meta.site_url && anyUrl) {
    const m = anyUrl.match(/^(https?:\/\/[^\/]+)/i);
    if (m) meta.site_url = m[1] + '/';
  }

  if (!meta.site_id && anyUrl) {
    try {
      if (typeof sdsdSiteIdFromUrl_ === 'function') {
        meta.site_id = String(sdsdSiteIdFromUrl_(anyUrl) || '');
      }
    } catch(e) {}
  }

  if (!meta.site_name) meta.site_name = meta.site_id;

  return meta;
}

function sdsdMergeReferralBatchId_() {
  try {
    const props = PropertiesService.getDocumentProperties();
    return String(
      props.getProperty('SDSD_SITE_WIDE_BATCH_ID') ||
      props.getProperty('SDSD_LAST_SITE_WIDE_BATCH_ID') ||
      ''
    );
  } catch(e) {
    return '';
  }
}

function sdsdBuildMergeReferralObject_(selected, stored) {
  const c = stored && stored.diagnosis_case ? stored.diagnosis_case : {};
  const site = sdsdMergeReferralSiteMeta_(selected);

  const diagnosisCaseId = String(
    c.diagnosis_case_id ||
    selected.parent_diagnosis_case_id ||
    selected.diagnosis_case_id ||
    ''
  );

  return {
    format: 'SIMS_MERGE_REFERRAL_V1',
    contract_version: '1.0',
    generated_at: new Date().toISOString(),
    source: 'SITE_TREATMENT_PLAN',
    site_diagnosis_batch_id: sdsdMergeReferralBatchId_(),
    diagnosis_case_id: diagnosisCaseId,
    site: {
      site_id: String(site.site_id || ''),
      site_name: String(site.site_name || ''),
      site_url: String(site.site_url || '')
    },
    diagnosis: {
      theme: String(selected.diagnosis_theme || ''),
      doctor_decision: 'MERGE',
      confidence: String(selected.confidence || ''),
      reason: String(selected.reason || ''),
      treatment_strategy: String(c.treatment_strategy || '')
    },
    merge_plan: {
      survivor: {
        article_id: String(selected.survivor.article_id || ''),
        article_url: String(selected.survivor.article_url || ''),
        role: 'SURVIVOR'
      },
      absorbed: {
        article_id: String(selected.absorbed.article_id || ''),
        article_url: String(selected.absorbed.article_url || ''),
        role: 'ABSORBED'
      },
      direction: String(
        selected.merge_direction ||
        c.merge_direction ||
        `${selected.absorbed.article_id} → ${selected.survivor.article_id}`
      ),
      content_to_absorb: String(c.merge_content_to_absorb || ''),
      redirect_policy: '301_REDIRECT_AFTER_SAFE_MERGE'
    },
    merge_instructions: [
      'Doctorが確定した統合方向を変更しない',
      '統合元の独自情報だけを統合先へ安全に吸収する',
      '統合先の既存の有効な内容・検索意図・URLを保護する',
      '統合作業後に統合元から統合先への301リダイレクトを設定する',
      '統合結果をSBMへ返せるよう、ArticleID・URL・実施内容を保持する'
    ],
    blocked_actions: [
      '統合方向の再判定',
      '新規記事作成',
      '統合先URLの変更',
      'Doctorが許可していない全面リライト',
      '統合確認前の統合元記事の単純削除'
    ],
    workflow: {
      current_owner: 'MERGE',
      return_to: 'SIMS_BLOG_MANAGER',
      next_after_merge: 'MONITOR'
    }
  };
}

function sdsdBuildMergeReferralMarkdown_(referral) {
  const d = referral.diagnosis || {};
  const p = referral.merge_plan || {};
  const s = p.survivor || {};
  const a = p.absorbed || {};

  return [
    '# SIMS Merge 紹介状',
    '',
    `診断テーマ: ${d.theme || ''}`,
    `Doctor判断: ${d.doctor_decision || ''}`,
    `確信度: ${d.confidence || ''}`,
    '',
    '## 統合方向',
    '',
    `残す記事: ${s.article_id || ''} / ${s.article_url || ''}`,
    `吸収する記事: ${a.article_id || ''} / ${a.article_url || ''}`,
    `方向: ${p.direction || ''}`,
    '',
    '## Doctorの理由',
    '',
    d.reason || '',
    '',
    '## 吸収する内容',
    '',
    p.content_to_absorb || '統合元にのみ存在する独自情報を確認し、必要な内容だけを吸収してください。',
    '',
    '## Mergeへの依頼',
    '',
    '- Doctorが確定した統合方向は変更しない',
    '- 統合元の独自情報のみを安全に吸収する',
    '- 統合先の既存URL・検索意図・有効な内容を保護する',
    '- 統合完了後、統合元から統合先へ301リダイレクトを設定する',
    '- 作業結果はSBMへ返却する',
    '',
    '## 禁止事項',
    '',
    '- 統合方向の再判定',
    '- 新規記事作成',
    '- 統合先URLの変更',
    '- Doctor未許可の全面リライト',
    '- 統合確認前の統合元記事の単純削除',
    ''
  ].join('\n');
}

function sdsdWriteMergeReferralSheet_(referral) {
  const ss = SpreadsheetApp.getActive();
  const name = sdsdMergeReferralSheetName_();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  sh.clear();

  const md = sdsdBuildMergeReferralMarkdown_(referral);
  const jsonText = JSON.stringify(referral, null, 2);

  const rows = [
    ['SIMS Merge 紹介状', ''],
    ['診断テーマ', referral.diagnosis.theme],
    ['Doctor判断', referral.diagnosis.doctor_decision],
    ['確信度', referral.diagnosis.confidence],
    ['統合先（残す記事）',
      `${referral.merge_plan.survivor.article_id} / ${referral.merge_plan.survivor.article_url}`],
    ['統合元（吸収する記事）',
      `${referral.merge_plan.absorbed.article_id} / ${referral.merge_plan.absorbed.article_url}`],
    ['統合方向', referral.merge_plan.direction],
    ['Doctorの理由', referral.diagnosis.reason],
    ['吸収する内容', referral.merge_plan.content_to_absorb],
    ['301方針', referral.merge_plan.redirect_policy],
    ['診断案件ID', referral.diagnosis_case_id],
    ['Site Diagnosis Batch ID', referral.site_diagnosis_batch_id],
    ['サイトID', referral.site.site_id],
    ['サイトURL', referral.site.site_url],
    ['', ''],
    ['Mergeへ渡す紹介状（Markdown）', md],
    ['', ''],
    ['Merge連携JSON', jsonText]
  ];

  sh.getRange(1,1,rows.length,2).setValues(rows);
  sh.getRange('A1:B1').setFontWeight('bold').setFontSize(14);
  sh.getRange(1,1,rows.length,2).setWrap(true);
  sh.setColumnWidth(1, 230);
  sh.setColumnWidth(2, 900);
  sh.setFrozenRows(1);

  // Make the long handoff blocks easier to copy.
  sh.setRowHeight(16, 320);
  sh.setRowHeight(18, 420);

  ss.setActiveSheet(sh);
  sh.getRange('B16').activate();

  return sh;
}

function sdsdCreateMergeReferralFromSelectedTreatment() {
  const ui = SpreadsheetApp.getUi();

  try {
    const selected = sdsdSelectedTreatmentPlanMerge_();
    const stored = sdsdFindStoredMergeCase_(selected);
    const referral = sdsdBuildMergeReferralObject_(selected, stored);

    sdsdWriteMergeReferralSheet_(referral);

    ui.alert(
      'Merge紹介状を作成しました。\n\n' +
      `診断テーマ: ${referral.diagnosis.theme}\n` +
      `統合方向: ${referral.merge_plan.direction}\n\n` +
      '「Merge紹介状」シートのMarkdownまたはJSONをSIMS Mergeへ渡してください。\n\n' +
      'この処理では記事の統合・削除・301設定は行っていません。'
    );

    return referral;

  } catch(e) {
    ui.alert(
      'Merge紹介状を作成できませんでした。\n\n' +
      String(e && e.message ? e.message : e)
    );
    throw e;
  }
}
