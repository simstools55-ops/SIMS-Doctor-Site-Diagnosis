/**
 * SIMS Doctor Site Diagnosis v0.5.0 Sprint 4
 * Formal Doctor site-wide result contract.
 */

function sdsdExtractJsonObject_(text) {
  let s = String(text || '').trim();
  if (!s) throw new Error('Doctor結果が空です。');

  const fence = s.match(/```(?:json)?\s*([\s\S]*?)```/i);
  if (fence) s = fence[1].trim();

  try {
    return JSON.parse(s);
  } catch(e) {}

  const first = s.indexOf('{');
  const last = s.lastIndexOf('}');
  if (first >= 0 && last > first) {
    return JSON.parse(s.slice(first,last+1));
  }
  throw new Error('Doctor結果からJSONを抽出できませんでした。');
}


function sdsdSiteWideSourceCaseMap_() {
  const map = {};
  try {
    sdsdSiteWidePackageRows_().forEach(x => {
      map[String(x.case_id || '')] = x;
    });
  } catch(e) {}
  return map;
}

function sdsdCanonicalArticleListFromSource_(source) {
  if (!source) return [];
  if (Array.isArray(source.target_articles)) return source.target_articles;
  return [];
}

function sdsdDoctorTriageCaseFromSource_(caseId, overrides, sourceMap) {
  overrides = overrides || {};
  const source = sourceMap[String(caseId || '')] || null;

  return {
    diagnosis_case_id: String(
      overrides.diagnosis_case_id ||
      overrides.case_id ||
      caseId ||
      ''
    ),
    diagnosis_theme: String(
      overrides.diagnosis_theme ||
      overrides.theme ||
      (source ? source.diagnosis_theme : '') ||
      ''
    ),
    diagnosis_type: String(
      overrides.diagnosis_type ||
      overrides.improvement_type ||
      (source ? source.improvement_type : '') ||
      ''
    ),
    absorbed_source_case_ids: Array.isArray(overrides.absorbed_source_case_ids)
      ? overrides.absorbed_source_case_ids.map(String) : [],
    target_articles: Array.isArray(overrides.target_articles)
      ? overrides.target_articles
      : sdsdCanonicalArticleListFromSource_(source),
    doctor_decision: String(
      overrides.doctor_decision ||
      overrides.decision ||
      overrides.routing ||
      ''
    ),
    confidence: String(
      overrides.confidence ||
      (source ? source.confidence : '') ||
      ''
    ),
    site_impact: String(
      overrides.site_impact ||
      (source ? source.expected_site_impact : '') ||
      ''
    ),
    treatment_strategy: String(
      overrides.treatment_strategy ||
      ''
    ),
    route_to: String(overrides.route_to || ''),
    eventual_route: String(overrides.eventual_route || ''),
    reason: String(
      overrides.reason ||
      overrides.rationale ||
      ''
    ),
    additional_evidence_needed: Array.isArray(overrides.additional_evidence_needed)
      ? overrides.additional_evidence_needed.map(String) : []
  };
}

function sdsdConvertDoctorTriageResult_(obj) {
  const sourceMap = sdsdSiteWideSourceCaseMap_();
  const out = [];
  const absorbed = {};

  // 1) Precision-diagnosis priority queue.
  (obj.priority_queue_for_precision_diagnosis || []).forEach(x => {
    const c = sdsdDoctorTriageCaseFromSource_(x.case_id, x, sourceMap);
    c.route_to = 'NEEDS_EVIDENCE';
    c.eventual_route = (x.next_route_candidates || []).join(' / ');
    c.doctor_decision = String(x.routing || 'ADDITIONAL_EVIDENCE_REQUIRED');
    out.push(c);

    (x.absorbed_source_case_ids || []).forEach(id => {
      absorbed[String(id)] = true;
    });
  });

  // 2) Content gaps merged with duplicate "new article" signals.
  (obj.content_gap_merged_into_existing_theme || []).forEach(x => {
    const c = sdsdDoctorTriageCaseFromSource_(x.case_id, x, sourceMap);
    c.route_to = 'NEEDS_EVIDENCE';
    c.eventual_route = 'WRITER';
    c.doctor_decision = String(
      x.routing || 'SBM_ROUTINE_QUEUE_PENDING_BODY_CHECK'
    );
    c.treatment_strategy = 'LIGHT_FIX_CANDIDATE';
    if (!c.reason) {
      c.reason = '既存記事で対応可能な可能性が高い。本文確認後にWriter軽微修正候補。';
    }
    out.push(c);

    (x.absorbed_source_case_ids || []).forEach(id => {
      absorbed[String(id)] = true;
    });
  });

  // 3) Standalone high-confidence routine queue.
  (obj.sbm_routine_queue_high_confidence_standalone || []).forEach(id => {
    if (absorbed[String(id)]) return;
    const c = sdsdDoctorTriageCaseFromSource_(id, {}, sourceMap);
    c.route_to = 'NEEDS_EVIDENCE';
    c.eventual_route = 'WRITER';
    c.doctor_decision = 'SBM_ROUTINE_QUEUE_PENDING_BODY_CHECK';
    c.treatment_strategy = 'LIGHT_FIX_CANDIDATE';
    c.reason = '高confidenceの単独コンテンツギャップ。本文確認後にWriter軽微修正候補。';
    out.push(c);
  });

  // 4) Medium-confidence / single-query findings -> monitor.
  (obj.low_priority_monitor_medium_confidence_single_query || []).forEach(id => {
    if (absorbed[String(id)]) return;
    const c = sdsdDoctorTriageCaseFromSource_(id, {}, sourceMap);
    c.route_to = 'MONITOR';
    c.doctor_decision = 'MONITOR_LOW_SIGNAL';
    c.reason = '中confidence・単一クエリの弱いシグナル。次回サイト診断まで経過観察。';
    out.push(c);
  });

  // 5) Creator candidates still require cannibalization/body check.
  (obj.creator_candidates_pending_cannibalization_check || []).forEach(id => {
    if (absorbed[String(id)]) return;
    const c = sdsdDoctorTriageCaseFromSource_(id, {}, sourceMap);
    c.route_to = 'NEEDS_EVIDENCE';
    c.eventual_route = 'CREATOR';
    c.doctor_decision = 'CREATOR_PENDING_CANNIBALIZATION_CHECK';
    c.reason = '新規記事化前に既存記事との重複・カニバリ確認が必要。';
    out.push(c);
  });

  // Deduplicate by diagnosis_case_id, keeping the first (higher-priority section).
  const seen = {};
  return out.filter(c => {
    const id = String(c.diagnosis_case_id || '');
    if (!id || seen[id]) return false;
    seen[id] = true;
    return true;
  });
}

function sdsdIsDoctorTriageShape_(obj) {
  return !!(
    obj &&
    (
      Array.isArray(obj.priority_queue_for_precision_diagnosis) ||
      Array.isArray(obj.content_gap_merged_into_existing_theme) ||
      Array.isArray(obj.sbm_routine_queue_high_confidence_standalone) ||
      Array.isArray(obj.low_priority_monitor_medium_confidence_single_query) ||
      Array.isArray(obj.creator_candidates_pending_cannibalization_check)
    )
  );
}

function sdsdValidateSiteWideResult_(obj) {
  if (!obj || obj.format !== 'SIMS_DOCTOR_SITE_WIDE_RESULT_V1') {
    throw new Error(
      '正式な横断診断結果ではありません。format は SIMS_DOCTOR_SITE_WIDE_RESULT_V1 が必要です。'
    );
  }

  if (!Array.isArray(obj.diagnosis_cases) && !sdsdIsDoctorTriageShape_(obj)) {
    throw new Error(
      'diagnosis_cases またはDoctor一次トリアージ形式の診断結果がありません。'
    );
  }

  if (Array.isArray(obj.diagnosis_cases)) {
    const allowed = {
      WRITER:true,
      MERGE:true,
      CREATOR:true,
      MONITOR:true,
      NO_ACTION:true,
      NEEDS_EVIDENCE:true
    };

    obj.diagnosis_cases.forEach((c,i) => {
      if (!c.diagnosis_case_id) {
        throw new Error(`diagnosis_cases[${i}] に diagnosis_case_id がありません。`);
      }
      if (!Array.isArray(c.absorbed_source_case_ids)) {
        throw new Error(`diagnosis_cases[${i}] に absorbed_source_case_ids がありません。`);
      }
      if (!allowed[String(c.route_to || '')]) {
        throw new Error(
          `diagnosis_cases[${i}] の route_to が不正です: ${c.route_to}`
        );
      }
    });
  }

  return true;
}

function sdsdNormalizeDoctorSiteWideResult_(obj) {
  const siteInput = obj.site || {};
  const fallbackSite = sdsdSiteMetaFromArticleMaster_();

  const rawCases = Array.isArray(obj.diagnosis_cases)
    ? obj.diagnosis_cases
    : sdsdConvertDoctorTriageResult_(obj);

  const overall =
    obj.overall_diagnosis ||
    (obj.site_wide_diagnosis
      ? obj.site_wide_diagnosis.overall_assessment
      : '') ||
    (obj.presentation ? obj.presentation.summary : '') ||
    '';

  return {
    format: obj.format,
    contract_version: String(obj.contract_version || '1.0'),
    generated_at: String(obj.generated_at || new Date().toISOString()),
    site: {
      site_id: String(
        siteInput.site_id ||
        obj.site_id ||
        fallbackSite.site_id ||
        ''
      ),
      site_name: String(
        siteInput.site_name ||
        obj.site_name ||
        fallbackSite.site_name ||
        ''
      ),
      site_url: String(
        siteInput.site_url ||
        obj.site_url ||
        fallbackSite.site_url ||
        ''
      )
    },
    site_diagnosis_batch_id: String(
      obj.site_diagnosis_batch_id || ''
    ),
    overall_diagnosis: String(overall),
    diagnosis_cases: rawCases.map(c => ({
      diagnosis_case_id: String(c.diagnosis_case_id || ''),
      diagnosis_theme: String(c.diagnosis_theme || ''),
      diagnosis_type: String(c.diagnosis_type || ''),
      absorbed_source_case_ids: (c.absorbed_source_case_ids || []).map(String),
      target_articles: Array.isArray(c.target_articles) ? c.target_articles : [],
      doctor_decision: String(c.doctor_decision || ''),
      confidence: String(c.confidence || ''),
      site_impact: String(c.site_impact || ''),
      treatment_strategy: String(c.treatment_strategy || ''),
      route_to: String(c.route_to || ''),
      eventual_route: String(c.eventual_route || ''),
      reason: String(c.reason || ''),
      additional_evidence_needed: Array.isArray(c.additional_evidence_needed)
        ? c.additional_evidence_needed.map(String) : []
    }))
  };
}

function sdsdStoreSiteWideResult_(obj) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SDSD_CONFIG.sheets.siteWideResult);
  if (!sh) sh = ss.insertSheet(SDSD_CONFIG.sheets.siteWideResult);

  sh.clear();
  sh.getRange('A1').setValue(JSON.stringify(obj, null, 2));
  sh.hideSheet();

  PropertiesService.getDocumentProperties().setProperty(
    'SDSD_LAST_SITE_WIDE_RESULT_AT',
    new Date().toISOString()
  );
}

function sdsdWriteTreatmentPlan_(obj) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SDSD_CONFIG.sheets.treatmentPlan);
  if (!sh) sh = ss.insertSheet(SDSD_CONFIG.sheets.treatmentPlan);
  sh.clear();

  const headers = [
    '優先順','診断テーマ','Doctor判断','次の処置',
    '対象記事','吸収した元案件','確信度','理由',
    '追加Evidence','状態'
  ];

  const routeJa = {
    WRITER:'Writerで既存記事を改善',
    MERGE:'Mergeで統合を検討',
    CREATOR:'Creatorで新記事を作成',
    MONITOR:'経過観察',
    NO_ACTION:'処置不要',
    NEEDS_EVIDENCE:'追加Evidenceを確認'
  };

  const priority = {
    NEEDS_EVIDENCE:1,
    MERGE:2,
    WRITER:3,
    CREATOR:4,
    MONITOR:5,
    NO_ACTION:6
  };

  const rows = obj.diagnosis_cases.slice().sort((a,b) =>
    (priority[a.route_to] || 99) - (priority[b.route_to] || 99)
  ).map((c,i) => [
    i+1,
    c.diagnosis_theme,
    c.doctor_decision,
    routeJa[c.route_to] || c.route_to,
    (c.target_articles || []).map(a => {
      if (typeof a === 'string') return a;
      return [a.article_id,a.article_title,a.article_url]
        .filter(Boolean).join(' / ');
    }).join('\n'),
    (c.absorbed_source_case_ids || []).join(' / '),
    c.confidence,
    c.reason,
    (c.additional_evidence_needed || []).join(' / '),
    c.route_to === 'NEEDS_EVIDENCE' ? '追加確認待ち' : '処置振り分け済み'
  ]);

  sh.getRange(1,1,1,headers.length).setValues([headers]);
  if (rows.length) {
    sh.getRange(2,1,rows.length,headers.length).setValues(rows);
  }

  sh.setFrozenRows(1);
  sh.getRange(1,1,1,headers.length).setFontWeight('bold');
  sh.getRange(1,1,Math.max(rows.length+1,1),headers.length).setWrap(true);
  [70,260,220,220,430,300,90,460,360,130]
    .forEach((w,i) => sh.setColumnWidth(i+1,w));

  ss.setActiveSheet(sh);
  return rows.length;
}


function sdsdEnsureSiteWideResultImportSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SDSD_CONFIG.sheets.siteWideResultImport);
  if (!sh) sh = ss.insertSheet(SDSD_CONFIG.sheets.siteWideResultImport);

  if (sh.getLastRow() === 0 || String(sh.getRange('A1').getValue() || '') !== 'Doctor横断診断結果をここに貼り付け') {
    sh.clear();
    sh.getRange('A1').setValue('Doctor横断診断結果をここに貼り付け');
    sh.getRange('A2').setValue(
      'このセル以下に、Doctor回答全文または SIMS_DOCTOR_SITE_WIDE_RESULT_V1 JSON を貼り付けてください。\n' +
      '貼り付け後、メニュー「12. Doctor横断診断結果を登録」をもう一度実行します。'
    );
    sh.getRange('A1').setFontWeight('bold').setFontSize(14);
    sh.getRange('A1:A2').setWrap(true);
    sh.setColumnWidth(1, 1000);
    sh.setRowHeight(2, 90);
  }
  return sh;
}

function sdsdReadSiteWideResultImportText_() {
  const sh = sdsdEnsureSiteWideResultImportSheet_();
  const last = Math.max(sh.getLastRow(), 2);
  const vals = sh.getRange(2,1,last-1,1).getDisplayValues()
    .map(r => String(r[0] || ''));

  const placeholder =
    'このセル以下に、Doctor回答全文または SIMS_DOCTOR_SITE_WIDE_RESULT_V1 JSON を貼り付けてください。';

  const cleaned = vals.filter((v,i) => {
    if (!String(v || '').trim()) return false;
    if (i === 0 && String(v).indexOf(placeholder) >= 0) return false;
    return true;
  });

  return cleaned.join('\n').trim();
}

function sdsdSetSiteWideRegisterStage_(stage, detail) {
  const props = PropertiesService.getDocumentProperties();
  props.setProperty('SDSD_SITE_WIDE_REGISTER_STAGE', String(stage || ''));
  props.setProperty('SDSD_SITE_WIDE_REGISTER_DETAIL', String(detail || ''));
  props.setProperty('SDSD_SITE_WIDE_REGISTER_AT', new Date().toISOString());
}

function sdsdClearSiteWideResultImport_() {
  const sh = sdsdEnsureSiteWideResultImportSheet_();
  if (sh.getLastRow() >= 2) {
    sh.getRange(2,1,Math.max(sh.getLastRow()-1,1),1).clearContent();
  }
  sh.getRange('A2').setValue(
    'このセル以下に、Doctor回答全文または SIMS_DOCTOR_SITE_WIDE_RESULT_V1 JSON を貼り付けてください。\n' +
    '貼り付け後、メニュー「12. Doctor横断診断結果を登録」をもう一度実行します。'
  );
}

function sdsdRegisterSiteWideDoctorResult() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();

  try {
    sdsdSetSiteWideRegisterStage_('START', '登録開始');

    const importSheet = sdsdEnsureSiteWideResultImportSheet_();
    const sourceText = sdsdReadSiteWideResultImportText_();

    if (!sourceText) {
      sdsdSetSiteWideRegisterStage_('WAITING_INPUT', 'Doctor結果取込シートへ入力待ち');
      ss.setActiveSheet(importSheet);
      ui.alert(
        'Doctor横断診断結果の取込準備ができました。\n\n' +
        '「Doctor結果取込」シートのA2以下へ、Doctor回答全文を貼り付けてください。\n\n' +
        '貼り付け後、もう一度\n' +
        '「12. Doctor横断診断結果を登録」\n' +
        'を実行してください。'
      );
      return;
    }

    sdsdSetSiteWideRegisterStage_('EXTRACT_JSON', `入力文字数: ${sourceText.length}`);
    const parsed = sdsdExtractJsonObject_(sourceText);

    sdsdSetSiteWideRegisterStage_('VALIDATE', parsed && parsed.format ? parsed.format : 'formatなし');
    sdsdValidateSiteWideResult_(parsed);

    sdsdSetSiteWideRegisterStage_('NORMALIZE', 'Doctor結果を正規化');
    const normalized = sdsdNormalizeDoctorSiteWideResult_(parsed);

    sdsdSetSiteWideRegisterStage_(
      'STORE_RAW_RESULT',
      `正規化案件数: ${(normalized.diagnosis_cases || []).length}`
    );
    sdsdStoreSiteWideResult_(normalized);

    sdsdSetSiteWideRegisterStage_('WRITE_TREATMENT_PLAN', 'サイト治療計画を生成');
    const count = sdsdWriteTreatmentPlan_(normalized);

    const counts = {};
    normalized.diagnosis_cases.forEach(c => {
      counts[c.route_to] = (counts[c.route_to] || 0) + 1;
    });

    sdsdSetSiteWideRegisterStage_('COMPLETE', `登録完了 ${count}件`);
    sdsdClearSiteWideResultImport_();

    ui.alert(
      `Doctor横断診断結果を登録しました。\n\n` +
      `診断案件: ${count}件\n` +
      `Writer: ${counts.WRITER || 0}件\n` +
      `Merge: ${counts.MERGE || 0}件\n` +
      `Creator: ${counts.CREATOR || 0}件\n` +
      `追加Evidence: ${counts.NEEDS_EVIDENCE || 0}件\n` +
      `経過観察/処置不要: ${(counts.MONITOR || 0) + (counts.NO_ACTION || 0)}件\n\n` +
      `「サイト治療計画」を確認してください。`
    );

  } catch(e) {
    const props = PropertiesService.getDocumentProperties();
    const stage = String(
      props.getProperty('SDSD_SITE_WIDE_REGISTER_STAGE') || 'UNKNOWN'
    );
    const message = String(e && e.message ? e.message : e);

    sdsdSetSiteWideRegisterStage_('ERROR', `${stage}: ${message}`);

    ui.alert(
      `Doctor横断診断結果を登録できませんでした。\n\n` +
      `処理段階: ${stage}\n` +
      `エラー: ${message}\n\n` +
      `Doctor結果取込シートの内容は残しています。`
    );
    throw e;
  }
}

function sdsdOpenTreatmentPlan() {
  const sh = SpreadsheetApp.getActive()
    .getSheetByName(SDSD_CONFIG.sheets.treatmentPlan);
  if (sh) SpreadsheetApp.getActive().setActiveSheet(sh);
}
