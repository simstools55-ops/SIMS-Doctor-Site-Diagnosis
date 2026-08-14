/**
 * SIMS Doctor Site Diagnosis v0.5.0 Sprint 3
 * Export one site-wide Doctor package preserving site context.
 */

function sdsdSiteWidePackageRows_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SDSD_CONFIG.sheets.opportunityCases);
  if (!sh || sh.getLastRow() < 2) {
    throw new Error('先に「9. サイト横断診断案件を作成」を実行してください。');
  }

  const values = sh.getDataRange().getValues();
  const headers = values[0].map(String);
  const idx = {};
  headers.forEach((h,i) => idx[h] = i);

  const required = [
    '案件ID','改善テーマ','診断テーマ','対象記事','関連クエリ数',
    '主な検索テーマ','確信度','サイト全体への期待効果',
    '診断で確認すること','次の担当','状態'
  ];
  required.forEach(h => {
    if (idx[h] == null) throw new Error(`サイト横断診断案件に必要列がありません: ${h}`);
  });

  return values.slice(1)
    .filter(r => String(r[idx['案件ID']] || '').trim())
    .map(r => ({
      case_id: String(r[idx['案件ID']] || ''),
      improvement_type: String(r[idx['改善テーマ']] || ''),
      diagnosis_theme: String(r[idx['診断テーマ']] || ''),
      target_articles_text: String(r[idx['対象記事']] || ''),
      related_query_count: Number(r[idx['関連クエリ数']] || 0),
      main_queries: String(r[idx['主な検索テーマ']] || ''),
      confidence: String(r[idx['確信度']] || ''),
      expected_site_impact: String(r[idx['サイト全体への期待効果']] || ''),
      doctor_focus: String(r[idx['診断で確認すること']] || ''),
      next_route: String(r[idx['次の担当']] || ''),
      status: String(r[idx['状態']] || '')
    }));
}

function sdsdSiteWideSummaryData_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SDSD_CONFIG.sheets.summary);
  const result = {};
  if (!sh || sh.getLastRow() < 1) return result;

  const vals = sh.getRange(1,1,sh.getLastRow(),Math.min(sh.getLastColumn(),2)).getValues();
  vals.forEach(r => {
    const key = String(r[0] || '').trim();
    const val = String(r[1] || '').trim();
    if (key) result[key] = val;
  });
  return result;
}

function sdsdBuildSiteWideDoctorPackageManifest_(rows) {
  const summary = sdsdSiteWideSummaryData_();

  const counts = {
    cannibal: rows.filter(x => x.improvement_type === 'カニバリ疑い').length,
    new_article: rows.filter(x => x.improvement_type === '新規記事機会').length,
    content_gap: rows.filter(x => x.improvement_type === 'コンテンツギャップ').length
  };

  return {
    format: 'SIMS_DOCTOR_SITE_WIDE_DIAGNOSIS_PACKAGE_V1',
    contract_version: '1.0',
    generated_at: new Date().toISOString(),
    package_scope: 'SITE_WIDE',
    diagnosis_policy: {
      purpose: 'サイト全体の改善機会を横断的に診断する',
      no_automatic_treatment: true,
      preserve_existing_successful_content: true,
      doctor_decides_final_route: true
    },
    site_summary: summary,
    case_count: rows.length,
    case_counts: counts,
    cases: rows.map(x => ({
      case_id: x.case_id,
      improvement_type: x.improvement_type,
      diagnosis_theme: x.diagnosis_theme,
      related_query_count: x.related_query_count,
      confidence: x.confidence,
      expected_site_impact: x.expected_site_impact,
      doctor_focus: x.doctor_focus,
      next_route: x.next_route,
      status: x.status
    }))
  };
}

function sdsdBuildSiteWideDoctorInstructions_(rows) {
  return [
    '# SIMS Doctor Site-wide Diagnosis Referral',
    '',
    'このパッケージは、個別記事の診断ではなくサイト全体の横断診断用です。',
    '',
    `案件数: ${rows.length}`,
    '',
    '## Doctorへの依頼',
    '',
    '- 各案件を単独で見るだけでなく、サイト全体の構造・検索意図・記事群の役割を横断して評価してください。',
    '- カニバリ疑いは、統合ありきではなく「維持 / 役割分担 / Writer修正 / Merge候補」を比較してください。',
    '- コンテンツギャップは、既存記事改善で足りるか、新記事へ分離すべきかを判断してください。',
    '- 新規記事機会は、既存記事との重複・新たなカニバリ発生リスクを確認してからCreator候補としてください。',
    '- 不明な項目は未評価とし、推測で処置を確定しないでください。',
    '',
    '## 出力してほしいもの',
    '',
    '- サイト全体の総合診断',
    '- 優先して処置すべき案件',
    '- 各案件の最終振り分け（MONITOR / Writer / Merge / Creator / 追加Evidence）',
    '- 共通原因がある場合は、案件横断でまとめて説明',
    '- 大規模改修が不要な場合は、その旨を明記',
    ''
  ].join('\n');
}

function sdsdExportSiteWideDoctorPackage() {
  try {
    const rows = sdsdSiteWidePackageRows_();
    if (!rows.length) {
      SpreadsheetApp.getUi().alert('サイト横断Doctor Packageへ入れる案件がありません。');
      return;
    }

    const manifest = sdsdBuildSiteWideDoctorPackageManifest_(rows);
    const blobs = [];

    blobs.push(
      Utilities.newBlob(
        JSON.stringify(manifest, null, 2),
        'application/json',
        'manifest.json'
      )
    );

    blobs.push(
      Utilities.newBlob(
        sdsdBuildSiteWideDoctorInstructions_(rows),
        'text/markdown',
        'DOCTOR-REFERRAL.md'
      )
    );

    rows.forEach((x,i) => {
      const folder = `cases/${String(i+1).padStart(3,'0')}-${x.case_id}`;
      blobs.push(
        Utilities.newBlob(
          JSON.stringify({
            format: 'SIMS_DOCTOR_SITE_WIDE_CASE_V1',
            contract_version: '1.0',
            case_id: x.case_id,
            improvement_type: x.improvement_type,
            diagnosis_theme: x.diagnosis_theme,
            target_articles_text: x.target_articles_text,
            related_query_count: x.related_query_count,
            main_queries: x.main_queries,
            confidence: x.confidence,
            expected_site_impact: x.expected_site_impact,
            doctor_focus: x.doctor_focus,
            next_route: x.next_route,
            status: x.status
          }, null, 2),
          'application/json',
          `${folder}/case.json`
        )
      );
    });

    const zipName =
      `SIMS-Doctor-Site-Wide-Diagnosis-${Utilities.formatDate(
        new Date(),
        Session.getScriptTimeZone() || 'Asia/Tokyo',
        'yyyyMMdd-HHmmss'
      )}.zip`;

    const file = DriveApp.createFile(Utilities.zip(blobs, zipName));

    SpreadsheetApp.getUi().alert(
      `サイト横断Doctor Packageを生成しました。\n\n` +
      `案件数: ${rows.length}件\n` +
      `保存先: Google Drive\n\n` +
      `${file.getUrl()}\n\n` +
      `このZIPをSIMS Doctorへ渡してください。`
    );

    return {
      caseCount: rows.length,
      fileId: file.getId(),
      fileName: file.getName(),
      fileUrl: file.getUrl()
    };

  } catch(e) {
    SpreadsheetApp.getUi().alert(
      `サイト横断Doctor Packageを生成できませんでした。\n\n${e.message || e}`
    );
    throw e;
  }
}
