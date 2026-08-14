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

function sdsdValidateSiteWideResult_(obj) {
  if (!obj || obj.format !== 'SIMS_DOCTOR_SITE_WIDE_RESULT_V1') {
    throw new Error(
      '正式な横断診断結果ではありません。format は SIMS_DOCTOR_SITE_WIDE_RESULT_V1 が必要です。'
    );
  }

  if (!Array.isArray(obj.diagnosis_cases)) {
    throw new Error('diagnosis_cases がありません。');
  }

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

  return true;
}

function sdsdNormalizeDoctorSiteWideResult_(obj) {
  const site = obj.site || {};
  return {
    format: obj.format,
    contract_version: String(obj.contract_version || '1.0'),
    generated_at: String(obj.generated_at || new Date().toISOString()),
    site: {
      site_id: String(site.site_id || ''),
      site_name: String(site.site_name || ''),
      site_url: String(site.site_url || '')
    },
    site_diagnosis_batch_id: String(obj.site_diagnosis_batch_id || ''),
    overall_diagnosis: String(obj.overall_diagnosis || ''),
    diagnosis_cases: obj.diagnosis_cases.map(c => ({
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

function sdsdRegisterSiteWideDoctorResult() {
  const ui = SpreadsheetApp.getUi();
  const prompt = ui.prompt(
    'Doctor横断診断結果を登録',
    'Doctor回答の SIMS_DOCTOR_SITE_WIDE_RESULT_V1 JSON を貼り付けてください。',
    ui.ButtonSet.OK_CANCEL
  );
  if (prompt.getSelectedButton() !== ui.Button.OK) return;

  try {
    const parsed = sdsdExtractJsonObject_(prompt.getResponseText());
    sdsdValidateSiteWideResult_(parsed);
    const normalized = sdsdNormalizeDoctorSiteWideResult_(parsed);
    sdsdStoreSiteWideResult_(normalized);
    const count = sdsdWriteTreatmentPlan_(normalized);

    const counts = {};
    normalized.diagnosis_cases.forEach(c => {
      counts[c.route_to] = (counts[c.route_to] || 0) + 1;
    });

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
    ui.alert(`Doctor横断診断結果を登録できませんでした。\n\n${e.message || e}`);
    throw e;
  }
}

function sdsdOpenTreatmentPlan() {
  const sh = SpreadsheetApp.getActive()
    .getSheetByName(SDSD_CONFIG.sheets.treatmentPlan);
  if (sh) SpreadsheetApp.getActive().setActiveSheet(sh);
}
