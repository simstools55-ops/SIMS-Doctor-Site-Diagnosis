function sdsdBuildTreatmentBatch(options) {
  options = options || {};
  const ss = SpreadsheetApp.getActive();
  const source = ss.getSheetByName(SDSD_CONFIG.sheets.candidates);
  if (!source) throw new Error('先に Run Site Analysis を実行してください。');

  const values = source.getDataRange().getValues();
  if (values.length < 2) throw new Error('候補データがありません。');

  const headers = values[0].map(String);
  const idx = {};
  headers.forEach((h,i) => idx[h] = i);

  const rows = values.slice(1).filter(r => r[idx['Normalized URL']]);
  const articleCount = rows.length;
  const capacity = sdsdTreatmentCapacity_(articleCount);

  const eligible = rows.filter(r => {
    const p = String(r[idx['Priority Candidate']] || '');
    const guard = String(r[idx['Recent Treatment Guard']] || '');
    const ownership = String(r[idx['Ownership']] || '');
    return (
      guard === 'PASS' &&
      ownership === 'DOCTOR_OWNED' &&
      (p === 'A1_CANDIDATE' || p === 'A2_CANDIDATE')
    );
  });

  eligible.sort((a,b) => {
    const pa = String(a[idx['Priority Candidate']]) === 'A1_CANDIDATE' ? 0 : 1;
    const pb = String(b[idx['Priority Candidate']]) === 'A1_CANDIDATE' ? 0 : 1;
    if (pa !== pb) return pa - pb;
    return Number(b[idx['TVS']] || 0) - Number(a[idx['TVS']] || 0);
  });

  // standardMax is an upper bound, not a quota.
  // We never auto-fill from B/REVIEW merely to reach standardMin.
  const selected = eligible.slice(0, capacity.standardMax);

  const queryMap = sdsdBuildQueryEvidenceMap_();
  const historyMap = sdsdBuildHistoryMap_();
  const batchId = sdsdCreateBatchId_();

  let out = ss.getSheetByName(SDSD_CONFIG.sheets.selectedCases);
  if (!out) out = ss.insertSheet(SDSD_CONFIG.sheets.selectedCases);
  out.clear();

  const userHeaders = [
    'No.','記事URL','診断優先度','選定理由','状態'
  ];
  const technicalHeaders = [
    'Batch Order','Site Priority','URL','TVS','Weekly Trend','Evidence Confidence',
    'Treatment Risk','External Factor','Ownership','Recent Treatment Guard',
    'Top Queries','Selection Reason','Referral Status','Referral JSON'
  ];
  const outHeaders = userHeaders.concat(technicalHeaders);
  out.getRange(1,1,1,outHeaders.length).setValues([outHeaders]);

  const batchRows = selected.map((r,i) => {
    const url = String(r[idx['Normalized URL']] || '');
    const queries = (queryMap[url] || []).slice(0,5).map(q => q.query).filter(Boolean);
    const history = historyMap[url] || null;

    const referral = {
      format: 'SIMS_DOCTOR_INDIVIDUAL_CASE_PACKAGE_V1',
      contract_version: '1.0-draft',
      case_identity: {
        site_diagnosis_case_id: sdsdBuildSiteDiagnosisCaseId_(batchId, url),
        individual_case_id: '',
        site_id: sdsdSiteIdFromUrl_(url),
        article_id: '',
        url: url
      },
      site_referral: {
        treatment_value_score: Number(r[idx['TVS']] || 0),
        site_priority: String(r[idx['Priority Candidate']] || '').replace('_CANDIDATE',''),
        treatment_ownership: String(r[idx['Ownership']] || ''),
        treatment_risk: String(r[idx['Treatment Risk']] || ''),
        evidence_confidence: String(r[idx['Evidence Confidence']] || ''),
        weekly_trend: String(r[idx['Weekly Trend']] || ''),
        external_factors: String(r[idx['External Factor']] || '')
          .split('|').filter(Boolean),
        selection_reasons: [String(r[idx['Reason']] || '')].filter(Boolean)
      },
      search_evidence: {
        evidence_window_days: 120,
        top_queries: queries
      },
      treatment_history: history ? {
        last_treatment_date: history.date ? history.date.toISOString() : null,
        treatment_route: history.route || '',
        monitor_status: history.status || ''
      } : null,
      recent_treatment_guard: {
        status: String(r[idx['Recent Treatment Guard']] || ''),
        checked_at: new Date().toISOString()
      },
      article_evidence: {
        status: 'NOT_ATTACHED_IN_SPRINT3',
        note: 'Article body / ArticleID will be attached by the SBM/Case Packager integration.'
      },
      site_doctor_expected_route: 'INDIVIDUAL_DOCTOR',
      site_diagnosis_batch_id: batchId
    };

    const sitePriority = String(r[idx['Priority Candidate']] || '').replace('_CANDIDATE','');
    const priorityJa = sitePriority === 'A1' ? '最優先' : sitePriority === 'A2' ? '優先' : sitePriority;
    const reason = String(r[idx['Reason']] || '');

    return [
      i+1,
      url,
      priorityJa,
      reason,
      'Doctor診断待ち',

      i+1,
      sitePriority,
      url,
      Number(r[idx['TVS']] || 0),
      String(r[idx['Weekly Trend']] || ''),
      String(r[idx['Evidence Confidence']] || ''),
      String(r[idx['Treatment Risk']] || ''),
      String(r[idx['External Factor']] || ''),
      String(r[idx['Ownership']] || ''),
      String(r[idx['Recent Treatment Guard']] || ''),
      queries.join(' / '),
      reason,
      'READY_FOR_CASE_ENRICHMENT',
      JSON.stringify(referral)
    ];
  });

  if (batchRows.length) {
    out.getRange(2,1,batchRows.length,outHeaders.length).setValues(batchRows);
  }
  out.setFrozenRows(1);
  out.getRange(1,1,1,userHeaders.length).setFontWeight('bold');
  out.autoResizeColumns(1,userHeaders.length);
  out.setColumnWidth(2, 360);
  out.setColumnWidth(4, 420);
  out.setColumnWidth(5, 150);
  sdsdHideTechnicalColumns_(out, userHeaders.length + 1, outHeaders.length);
  if (!options.noActivate) ss.setActiveSheet(out);

  const status = selected.length < capacity.standardMin
    ? '標準件数未満ですが、数合わせせず終了'
    : '標準範囲内';

  const result = {
    articleCount: articleCount,
    eligibleCount: eligible.length,
    selectedCount: selected.length,
    standardMin: capacity.standardMin,
    standardMax: capacity.standardMax,
    hardMax: capacity.hardMax,
    status: status
  };

  if (!options.silent) {
    SpreadsheetApp.getUi().alert(
      `Treatment Batchを作成しました。\n\n` +
      `診断対象: ${articleCount}記事\n` +
      `標準: ${capacity.standardMin}～${capacity.standardMax}件 / 最大${capacity.hardMax}件\n` +
      `適格候補: ${eligible.length}件\n` +
      `今回選定: ${selected.length}件\n` +
      `${status}\n\n` +
      `「${SDSD_CONFIG.sheets.selectedCases}」で結果を確認してください。`
    );
  }
  return result;
}

function sdsdOpenSelectedCases() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SDSD_CONFIG.sheets.selectedCases);
  if (sh) SpreadsheetApp.getActive().setActiveSheet(sh);
}
