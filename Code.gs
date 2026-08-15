// ============================================================================
// Source: SiteDiagnosisConfig.gs
// ============================================================================
const SDSD_VERSION = '0.4.0-RC9.1';

const SDSD_CONFIG = Object.freeze({
  sheets: {
    home: 'Diagnosis Home',
    evidencePageSummary: '_SDSD_PAGE_SUMMARY',
    evidencePageWeekly: '_SDSD_PAGE_WEEKLY',
    evidencePageQuery: '_SDSD_PAGE_QUERY_TOP',
    sbmHistory: '_SDSD_SBM_HISTORY',
    summary: 'サイト診断サマリー',
    candidates: '診断候補',
    selectedCases: '今回の診断対象',
    articleMaster: '_SDSD_ARTICLE_MASTER',
    opportunities: 'サイト改善プラン',
    opportunityCases: 'サイト横断診断案件',
    siteWideResult: '_SDSD_SITE_WIDE_RESULT',
    treatmentPlan: 'サイト治療計画',
    siteWideResultImport: 'Doctor結果取込'
  },
  score: {
    demandMax: 30,
    opportunityMax: 30,
    urgencyMax: 25,
    assetMax: 15
  },
  guardDays: 35
});

// ============================================================================
// Source: Code.gs
// ============================================================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  const wholeSiteMenu = ui.createMenu('1. サイト全体を診断する')
    .addItem('Evidence Packageを読み込む', 'sdsdImportEvidencePackageZip')
    .addItem('サイト診断を実行', 'sdsdRunProductDiagnosis')
    .addItem('診断結果を見る', 'sdsdOpenSiteSummary')
    .addItem('診断候補を見る', 'sdsdOpenCandidates')
    .addSeparator()
    .addItem('横断診断：改善機会を診断', 'sdsdRunSiteOpportunityDiagnosis')
    .addItem('横断診断：Doctor案件を作成', 'sdsdBuildSiteOpportunityCases')
    .addItem('横断診断：Doctor Packageを生成', 'sdsdExportSiteWideDoctorPackage')
    .addItem('横断診断：Doctor結果を登録', 'sdsdRegisterSiteWideDoctorResult');

  const treatmentMenu = ui.createMenu('2. 見つかった問題を処置する')
    .addItem('個別記事：Doctorへ送る記事を選ぶ', 'sdsdCreateProductTreatmentBatch')
    .addItem('個別記事：選定記事を見る', 'sdsdOpenSelectedCases')
    .addItem('個別記事：Doctor精密診断Packageを生成', 'sdsdCreateProductCasePackage')
    .addSeparator()
    .addItem('横断診断：サイト治療計画を見る', 'sdsdOpenTreatmentPlan')
    .addItem('横断診断：追加Evidence Packageを生成', 'sdsdExportPriorityPrecisionClusterPackage')
    .addItem('横断診断：選択中のMerge紹介状を作成', 'sdsdCreateMergeReferralFromSelectedTreatment');

  const otherMenu = ui.createMenu('その他・管理')
    .addItem('現在の診断状況を確認', 'sdsdShowCurrentSessionStatus')
    .addItem('現在の診断を終了', 'sdsdEndCurrentDiagnosisSession')
    .addSeparator()
    .addItem('初期設定を実行', 'sdsdInitialize')
    .addItem('Article Masterの取込案内', 'sdsdArticleMasterImportHelp')
    .addItem('SBM改善履歴の取込案内', 'sdsdImportHistoryHelp')
    .addSeparator()
    .addItem('内部シートを表示', 'sdsdShowInternalSheets')
    .addItem('内部シートを隠す', 'sdsdHideInternalSheets');

  ui.createMenu('SIMS Doctor Site Diagnosis')
    .addItem('Homeを開く', 'sdsdOpenHome')
    .addSeparator()
    .addSubMenu(wholeSiteMenu)
    .addSubMenu(treatmentMenu)
    .addSeparator()
    .addSubMenu(otherMenu)
    .addToUi();

  try {
    sdsdRenderHome_();
    const home=SpreadsheetApp.getActive().getSheetByName(SDSD_CONFIG.sheets.home);
    if(home) SpreadsheetApp.getActive().setActiveSheet(home);
  } catch(e) {}
}


function sdsdOpenHome(){
  sdsdRenderHome_();
  const sh=SpreadsheetApp.getActive().getSheetByName(SDSD_CONFIG.sheets.home);
  if(sh) SpreadsheetApp.getActive().setActiveSheet(sh);
}

function sdsdHomeReadStoredSiteWideResult_(){
  try{
    const sh=SpreadsheetApp.getActive().getSheetByName(SDSD_CONFIG.sheets.siteWideResult);
    if(!sh||!sh.getLastRow())return null;
    const raw=String(sh.getRange('A1').getValue()||'').trim();
    return raw?JSON.parse(raw):null;
  }catch(e){return null;}
}

function sdsdHomeDiagnosisMetrics_(){
  let rows=[];
  try{rows=sdsdCandidateRowsFromSheet_()||[];}catch(e){rows=[];}

  const m={
    total:rows.length,a1:0,a2:0,review:0,sbm:0,protected:0,wait:0,
    severe:0,traffic:0,ranking:0,volatile:0,growth:0,stable:0,
    crossTotal:0,cannibal:0,newArticle:0,gap:0
  };
  rows.forEach(r=>{
    const p=String(r.priority||'');
    if(p==='A1_CANDIDATE')m.a1++;
    else if(p==='A2_CANDIDATE')m.a2++;
    else if(/DOCTOR_REVIEW|REVIEW|B_CANDIDATE|CANDIDATE/.test(p))m.review++;
    else if(p==='SBM')m.sbm++;
    else if(p==='PROTECTED')m.protected++;
    else if(p==='WAIT')m.wait++;

    const t=String(r.weeklyTrend||'');
    if(t==='SEVERE_DECLINE')m.severe++;
    else if(t==='TRAFFIC_DECLINE')m.traffic++;
    else if(t==='RANKING_DECLINE')m.ranking++;
    else if(t==='VOLATILE')m.volatile++;
    else if(t==='GROWTH')m.growth++;
    else if(t==='STABLE')m.stable++;
  });

  try{
    const opp=sdsdDetectSiteOpportunities_()||[];
    m.crossTotal=opp.length;
    opp.forEach(x=>{
      const type=String(x.type||'');
      if(type.indexOf('カニバリ')>=0)m.cannibal++;
      else if(type.indexOf('新規記事')>=0)m.newArticle++;
      else if(type.indexOf('コンテンツギャップ')>=0)m.gap++;
    });
  }catch(e){}

  return m;
}

function sdsdHomeOverallStatus_(session,m,work,stored){
  if(!session.active)return {label:'未診断',tone:'GRAY',note:'Evidence Packageを読み込むと診断を開始できます。'};
  if(stored){
    const cases=Array.isArray(stored.diagnosis_cases)?stored.diagnosis_cases:[];
    const urgent=cases.filter(c=>/MERGE|WRITER|CREATOR/.test(String(c.route_to||''))).length;
    const needs=cases.filter(c=>String(c.route_to||'')==='NEEDS_EVIDENCE').length;
    if(urgent>0)return {label:'治療方針あり',tone:'YELLOW',note:`Doctor診断済み。処置候補${urgent}件、詳しい確認が必要${needs}件です。`};
    if(needs>0)return {label:'詳しい確認が必要',tone:'YELLOW',note:`Doctor診断済み。詳しい確認が必要な案件が${needs}件あります。`};
    return {label:'大きな緊急所見なし',tone:'GREEN',note:'Doctor横断診断では緊急の治療案件は確認されていません。'};
  }
  if(m.severe>0||m.a1>0)return {label:'優先確認あり',tone:'RED',note:`優先度の高い診断候補${m.a1}件、または大きな悪化${m.severe}件を確認しています。`};
  if(m.crossTotal>0)return {label:'横断確認を推奨',tone:'YELLOW',note:`関連する複数記事・検索機会を${m.crossTotal}件検出しています。`};
  if(m.total>0)return {label:'個別確認あり',tone:'YELLOW',note:`詳しく確認する候補が${m.total}件あります。`};
  return {label:'診断待ち',tone:'BLUE',note:'Evidenceは読み込み済みです。サイト診断を実行してください。'};
}

function sdsdHomeGuide_(session,m,work,stored){
  const props=PropertiesService.getDocumentProperties();
  const stage=String(props.getProperty('SDSD_SITE_WIDE_REGISTER_STAGE')||'');

  if(!session.active){
    return {
      title:'Evidence Packageを読み込んでください',
      reason:'まだ診断対象サイトのデータが読み込まれていません。',
      path:'1. サイト全体を診断する → Evidence Packageを読み込む',
      tone:'BLUE'
    };
  }

  if(stage==='WAITING_INPUT'){
    return {
      title:'Doctorの横断診断結果を登録してください',
      reason:'Doctorへ依頼する段階は完了しています。返ってきた診断結果をDiagnosisへ戻す段階です。',
      path:'1. サイト全体を診断する → 横断診断：Doctor結果を登録',
      tone:'YELLOW'
    };
  }

  if(stored){
    if(work.additionalEvidence>0){
      return {
        title:`詳しい診断が必要な案件が${work.additionalEvidence}件あります`,
        reason:'Doctorが、現在の情報だけでは治療方法を確定できないと判断しました。記事本文などを追加して詳しく診断します。',
        path:'2. 見つかった問題を処置する → 横断診断：追加Evidence Packageを生成',
        tone:'YELLOW'
      };
    }
    if(work.actionableTreatment>0){
      return {
        title:`治療へ進める案件が${work.actionableTreatment}件あります`,
        reason:'Doctorの診断で治療方針が決まりました。Writer / Merge / Creatorへ引き渡します。',
        path:'2. 見つかった問題を処置する → サイト治療計画を見る',
        tone:'BLUE'
      };
    }
    return {
      title:'Diagnosisで行う作業は完了しています',
      reason:'現在、追加診断や治療へ送る案件はありません。処置済み記事はSBMで経過観察します。',
      path:'必要なら「その他・管理 → 現在の診断を終了」で次のサイトへ進めます',
      tone:'GREEN'
    };
  }

  if(work.opportunityCases>0){
    return {
      title:'Doctorによるサイト横断診断を進めてください',
      reason:`Diagnosisが複数記事をまとめて確認すべき案件を${work.opportunityCases}件作成済みです。`,
      path:'1. サイト全体を診断する → 横断診断：Doctor Packageを生成',
      tone:'BLUE'
    };
  }

  if(m.crossTotal>0){
    const reasons=[];
    if(m.cannibal)reasons.push(`カニバリ疑い${m.cannibal}件`);
    if(m.newArticle)reasons.push(`新規記事機会${m.newArticle}件`);
    if(m.gap)reasons.push(`コンテンツギャップ${m.gap}件`);
    return {
      title:'Doctorによるサイト横断診断が必要です',
      reason:`Diagnosisが、1記事だけでは判断しにくい関連案件を検出しました（${reasons.join('、')||m.crossTotal+'件'}）。利用者が要否を判断する必要はありません。`,
      path:'1. サイト全体を診断する → 横断診断：Doctor案件を作成',
      tone:'YELLOW'
    };
  }

  if(m.total>0){
    return {
      title:`個別に詳しく診る記事が${m.total}件あります`,
      reason:'他記事との関係をまとめて調べる横断診断は不要と判定しました。優先記事をDoctorで詳しく診断します。',
      path:'2. 見つかった問題を処置する → 個別記事：Doctorへ送る記事を選ぶ',
      tone:'BLUE'
    };
  }

  return {
    title:'サイト診断を実行してください',
    reason:'Evidence Packageは読み込み済みですが、サイト全体の診断結果がまだありません。',
    path:'1. サイト全体を診断する → サイト診断を実行',
    tone:'BLUE'
  };
}

function sdsdHomeComment_(m,stored){
  if(stored&&stored.overall_diagnosis)return String(stored.overall_diagnosis);
  if(!m.total)return 'サイト診断を実行すると、ここにサイト全体の所見を表示します。';
  const parts=[];
  if(m.a1||m.a2)parts.push(`優先的に詳しく確認する記事が${m.a1+m.a2}件あります`);
  if(m.severe||m.traffic||m.ranking)parts.push(`悪化シグナルが${m.severe+m.traffic+m.ranking}件あります`);
  if(m.crossTotal)parts.push(`記事同士の関係を確認すべき横断案件が${m.crossTotal}件あります`);
  if(!parts.length)parts.push('緊急性の高いシグナルは多くありません');
  return parts.join('。')+'。Diagnosisが次の診断経路を自動判定します。';
}

function sdsdToneColor_(tone){
  const map={
    BLUE:{bg:'#E8F0FE',fg:'#174EA6',accent:'#185ABC'},
    GREEN:{bg:'#E6F4EA',fg:'#137333',accent:'#188038'},
    YELLOW:{bg:'#FEF7E0',fg:'#9A6700',accent:'#F9AB00'},
    RED:{bg:'#FCE8E6',fg:'#B3261E',accent:'#D93025'},
    GRAY:{bg:'#F1F3F4',fg:'#5F6368',accent:'#80868B'}
  };
  return map[tone]||map.BLUE;
}

function sdsdRenderHome_(){
  sdsdProductEnsureSheets_();
  const ss=SpreadsheetApp.getActive();
  let sh=ss.getSheetByName(SDSD_CONFIG.sheets.home);
  if(!sh)sh=ss.insertSheet(SDSD_CONFIG.sheets.home,0);
  sh.showSheet();
  sh.clear();

  const session=sdsdGetCurrentSession_();
  const work=sdsdSessionWorkSummary_();
  const metrics=sdsdHomeDiagnosisMetrics_();
  const stored=sdsdHomeReadStoredSiteWideResult_();
  const overall=sdsdHomeOverallStatus_(session,metrics,work,stored);
  const guide=sdsdHomeGuide_(session,metrics,work,stored);
  const siteLabel=session.active?(session.siteName||session.siteId||session.host||'判定できません'):'未読込';
  const overallColor=sdsdToneColor_(overall.tone);
  const guideColor=sdsdToneColor_(guide.tone);

  sh.getRange('A1:H1').merge().setValue('SIMS Doctor | Site Diagnosis');
  sh.getRange('A2:H2').merge().setValue('この画面を見れば、サイトの状態と次に行う作業が分かります。');

  sh.getRange('A4:B8').setValues([
    ['対象サイト',siteLabel],
    ['Evidence',session.evidenceFileName||'未読込'],
    ['総合状態',overall.label],
    ['診断候補',metrics.total+'件'],
    ['要確認作業',work.pendingTotal+'件']
  ]);

  sh.getRange('D4:H4').merge().setValue('サイト全体の分析結果');
  sh.getRange('D5:E9').setValues([
    ['優先確認',metrics.a1+metrics.a2+'件'],
    ['大きな悪化',metrics.severe+'件'],
    ['カニバリ疑い',metrics.cannibal+'件'],
    ['新規記事機会',metrics.newArticle+'件'],
    ['コンテンツギャップ',metrics.gap+'件']
  ]);
  sh.getRange('F5:H9').merge().setValue(sdsdHomeComment_(metrics,stored));

  sh.getRange('A11:H11').merge().setValue('Diagnosisの判断');
  sh.getRange('A12:H13').merge().setValue(overall.note);

  sh.getRange('A15:H15').merge().setValue('次に行うこと');
  sh.getRange('A16:H17').merge().setValue(guide.title);
  sh.getRange('A18:H19').merge().setValue('理由：'+guide.reason);
  sh.getRange('A20:H20').merge().setValue('操作：'+guide.path);

  sh.getRange('A22:H22').merge().setValue('基本の使い方');
  sh.getRange('A23:H23').merge().setValue('Collectorで最新データを収集 → Diagnosisへ読み込む → このHomeの「次に行うこと」に従う');

  sh.getRange('A25:H25').merge().setValue('状態の色');
  sh.getRange('A26:H26').merge().setValue('青：通常の操作　　緑：完了　　黄：確認・追加診断　　赤：優先確認　　灰：未診断・現在不要');

  sh.getRange('A1:H1').setBackground('#185ABC').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(18);
  sh.getRange('A2:H2').setBackground('#D2E3FC').setFontColor('#174EA6').setFontSize(11);

  sh.getRange('A4:A8').setBackground('#F1F3F4').setFontWeight('bold').setFontColor('#5F6368');
  sh.getRange('B4:B8').setBackground('#FFFFFF');
  sh.getRange('B6').setBackground(overallColor.bg).setFontColor(overallColor.fg).setFontWeight('bold');

  sh.getRange('D4:H4').setBackground('#D2E3FC').setFontColor('#174EA6').setFontWeight('bold');
  sh.getRange('D5:D9').setBackground('#F8F9FA').setFontWeight('bold').setFontColor('#5F6368');
  sh.getRange('E5:E9').setHorizontalAlignment('center').setFontWeight('bold');
  sh.getRange('F5:H9').setBackground('#F8F9FA');

  ['A11:H11','A15:H15','A22:H22','A25:H25'].forEach(r=>sh.getRange(r).setBackground('#D2E3FC').setFontColor('#174EA6').setFontWeight('bold'));
  sh.getRange('A12:H13').setBackground(overallColor.bg).setFontColor(overallColor.fg);
  sh.getRange('A16:H17').setBackground(guideColor.bg).setFontColor(guideColor.fg).setFontWeight('bold').setFontSize(14);
  sh.getRange('A18:H20').setBackground('#FFFFFF');
  sh.getRange('A20:H20').setFontWeight('bold').setFontColor('#174EA6');

  sh.getRange('A1:H26').setWrap(true).setVerticalAlignment('middle');
  sh.getRange('A1:H26').setFontFamily('Arial');
  sh.setFrozenRows(2);
  sh.setHiddenGridlines(true);
  sh.setColumnWidth(1,150);sh.setColumnWidth(2,310);
  sh.setColumnWidth(3,20);sh.setColumnWidth(4,145);sh.setColumnWidth(5,105);
  sh.setColumnWidth(6,150);sh.setColumnWidth(7,150);sh.setColumnWidth(8,150);
  sh.setRowHeight(1,36);sh.setRowHeight(2,28);
  sh.setRowHeights(4,5,30);sh.setRowHeights(11,16,28);
  sh.getRange('A1:H26').setBorder(false,false,false,false,false,false);
  sh.getRange('A4:B8').setBorder(true,true,true,true,true,true,'#DADCE0',SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange('D4:H9').setBorder(true,true,true,true,true,true,'#DADCE0',SpreadsheetApp.BorderStyle.SOLID);
}


function sdsdInitialize() {
  sdsdProductEnsureSheets_();
  sdsdHideInternalSheets_();
  SpreadsheetApp.getUi().alert(
    `SIMS Doctor Site Diagnosis ${SDSD_VERSION}\n初期設定が完了しました。\n\n次に「1. Evidence Packageを読み込む」を実行してください。`
  );
}


function sdsdImportHistoryHelp() {
  SpreadsheetApp.getUi().alert(
    'SBMの「改善履歴」CSVを _SDSD_SBM_HISTORY シートへ貼り付けてください。'
  );
}

function sdsdOpenSiteSummary() {
  try { sdsdRefreshSiteSummaryFromCandidates_(); } catch (e) {}
  const sh = SpreadsheetApp.getActive().getSheetByName(SDSD_CONFIG.sheets.summary);
  if (sh) SpreadsheetApp.getActive().setActiveSheet(sh);
}

function sdsdOpenCandidates() {
  try {
    sdsdRefreshCandidatesView_();
  } catch (e) {
    SpreadsheetApp.getUi().alert(
      `診断候補の表示を更新できませんでした。\n\n${e.message || e}`
    );
  }
  const sh = SpreadsheetApp.getActive().getSheetByName(SDSD_CONFIG.sheets.candidates);
  if (sh) SpreadsheetApp.getActive().setActiveSheet(sh);
}

// ============================================================================
// Source: ProductUX.gs
// ============================================================================
const SDSD_INTERNAL_SHEET_NAMES_ = Object.freeze([
  '_SDSD_PAGE_SUMMARY',
  '_SDSD_PAGE_WEEKLY',
  '_SDSD_PAGE_QUERY_TOP',
  '_SDSD_SBM_HISTORY',
  '_SDSD_ARTICLE_MASTER',
  'Weekly Trend Validation',
  'Priority Validation',
  'Query Evidence Diagnostics',
  'Site Diagnosis Candidates',
  'Selected Treatment Cases',
  'シート1',
  'Sheet1'
]);


function sdsdProductEnsureSheets_() {
  const ss = SpreadsheetApp.getActive();

  const legacyPairs = [
    ['Site Diagnosis Candidates', SDSD_CONFIG.sheets.candidates],
    ['Selected Treatment Cases', SDSD_CONFIG.sheets.selectedCases]
  ];
  legacyPairs.forEach(pair => {
    const legacy = ss.getSheetByName(pair[0]);
    const current = ss.getSheetByName(pair[1]);

    if (legacy && !current) {
      try {
        legacy.setName(pair[1]);
      } catch (e) {}
      return;
    }

    if (legacy && current) {
      try {
        if (ss.getActiveSheet().getSheetId() === legacy.getSheetId()) {
          ss.setActiveSheet(current);
        }
        legacy.hideSheet();
      } catch (e) {}
    }
  });

  const names = [
    SDSD_CONFIG.sheets.evidencePageSummary,
    SDSD_CONFIG.sheets.evidencePageWeekly,
    SDSD_CONFIG.sheets.evidencePageQuery,
    SDSD_CONFIG.sheets.sbmHistory,
    SDSD_CONFIG.sheets.summary,
    SDSD_CONFIG.sheets.candidates,
    SDSD_CONFIG.sheets.selectedCases,
    SDSD_CONFIG.sheets.articleMaster
  ];

  names.forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
  });

  sdsdHideUnusedDefaultSheet_();
  sdsdHideInternalSheets_();
  if (!ss.getSheetByName(SDSD_CONFIG.sheets.opportunities)) ss.insertSheet(SDSD_CONFIG.sheets.opportunities);
  if (!ss.getSheetByName(SDSD_CONFIG.sheets.opportunityCases)) ss.insertSheet(SDSD_CONFIG.sheets.opportunityCases);
  if (!ss.getSheetByName(SDSD_CONFIG.sheets.siteWideResult)) ss.insertSheet(SDSD_CONFIG.sheets.siteWideResult);
  if (!ss.getSheetByName(SDSD_CONFIG.sheets.treatmentPlan)) ss.insertSheet(SDSD_CONFIG.sheets.treatmentPlan);
  if (!ss.getSheetByName(SDSD_CONFIG.sheets.siteWideResultImport)) ss.insertSheet(SDSD_CONFIG.sheets.siteWideResultImport);
}

function sdsdHideUnusedDefaultSheet_() {
  const ss = SpreadsheetApp.getActive();

  ['シート1', 'Sheet1'].forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;

    const values = sh.getDataRange().getDisplayValues();
    const hasContent = values.some(row => row.some(v => String(v).trim() !== ''));
    if (hasContent) return;

    try {
      const preferred =
        ss.getSheetByName(SDSD_CONFIG.sheets.summary) ||
        ss.getSheetByName(SDSD_CONFIG.sheets.candidates) ||
        ss.getSheetByName(SDSD_CONFIG.sheets.selectedCases);

      if (preferred && ss.getActiveSheet().getSheetId() === sh.getSheetId()) {
        ss.setActiveSheet(preferred);
        SpreadsheetApp.flush();
      }

      sh.hideSheet();
    } catch (e) {}
  });
}

function sdsdHideTechnicalColumns_(sheet, firstTechnicalCol, totalCols) {
  if (!sheet || totalCols < firstTechnicalCol) return;
  try {
    sheet.hideColumns(firstTechnicalCol, totalCols - firstTechnicalCol + 1);
  } catch (e) {}
}


function sdsdArticleTitleMap_() {
  try {
    const map = sdsdBuildArticleMasterMap_();
    return map || {};
  } catch (e) {
    return {};
  }
}

function sdsdDisplayTitle_(url, articleMap) {
  const normalized = sdsdNormalizeUrl_(url);
  const master = articleMap && articleMap[normalized] ? articleMap[normalized] : null;
  return master && master.title ? String(master.title) : '（タイトル未取得）';
}

function sdsdPriorityJa_(priority) {
  const p = String(priority || '');
  if (p === 'A1_CANDIDATE') return '最優先';
  if (p === 'A2_CANDIDATE') return '優先';
  if (p === 'DOCTOR_REVIEW' || p === 'REVIEW' || p === 'B_CANDIDATE' || p === 'CANDIDATE') return '要確認';
  if (p === 'SBM') return '日常改善';
  if (p === 'PROTECTED') return '保護';
  if (p === 'WAIT') return '経過観察';
  return p || '要確認';
}

function sdsdPriorityRank_(priority) {
  const p = String(priority || '');
  if (p === 'A1_CANDIDATE') return 10;
  if (p === 'A2_CANDIDATE') return 20;
  if (p === 'DOCTOR_REVIEW') return 30;
  if (p === 'REVIEW' || p === 'B_CANDIDATE' || p === 'CANDIDATE') return 40;
  if (p === 'SBM') return 50;
  if (p === 'PROTECTED') return 60;
  if (p === 'WAIT') return 70;
  return 80;
}

function sdsdActionJa_(row) {
  const p = String(row.priority || '');
  if (p === 'A1_CANDIDATE' || p === 'A2_CANDIDATE') return 'Doctorで精密診断';
  if (p === 'DOCTOR_REVIEW' || p === 'REVIEW' || p === 'B_CANDIDATE' || p === 'CANDIDATE') return '追加確認';
  if (p === 'SBM') return 'SBMで日常改善';
  if (p === 'PROTECTED') return '今は大きく触らない';
  if (p === 'WAIT') return '経過観察';
  return '確認';
}

function sdsdReasonJa_(reason) {
  let s = String(reason || '');
  const replacements = [
    ['週次:SEVERE_DECLINE', '直近で大きく悪化'],
    ['週次:TRAFFIC_DECLINE', '直近の検索流入が低下'],
    ['週次:RANKING_DECLINE', '直近の検索順位が低下'],
    ['週次:VOLATILE', '直近の推移が不安定'],
    ['週次:GROWTH', '回復・成長傾向'],
    ['週次:STABLE', '直近の推移は安定'],
    ['外部要因:PLATFORM_OR_OS_CHANGE', 'プラットフォーム・OS変更の影響可能性'],
    ['PLATFORM_OR_OS_CHANGE', 'プラットフォーム・OS変更の影響可能性'],
    ['高Risk', '変更リスクが高い'],
    ['追加Evidence確認', '追加データの確認が必要'],
    ['TRAFFIC_DECLINE', '検索流入が低下'],
    ['RANKING_DECLINE', '検索順位が低下'],
    ['SEVERE_DECLINE', '直近で大きく悪化'],
    ['VOLATILE', '直近の推移が不安定'],
    ['GROWTH', '回復・成長傾向'],
    ['STABLE', '直近の推移は安定'],
    ['主病変がCTR/即効性改善', 'CTR（クリック率）の改善余地が大きい'],
    ['DOCTOR_OWNED', 'Doctor精密診断向き'],
    ['SBM_OWNED', 'SBMの日常改善向き']
  ];
  replacements.forEach(pair => {
    s = s.split(pair[0]).join(pair[1]);
  });
  s = s.replace(/\s*\/\s*/g, '／');
  s = s.replace(/\|/g, '・');
  return s;
}

function sdsdSiteMeaning_(row) {
  const trend = String(row.weeklyTrend || '');
  const ext = String(row.externalFactor || '');
  if (trend === 'SEVERE_DECLINE') return '大幅な悪化が出ている記事群の代表';
  if (trend === 'TRAFFIC_DECLINE') return '検索流入低下が出ている記事群の代表';
  if (trend === 'RANKING_DECLINE') return '検索順位低下が出ている記事群の代表';
  if (trend === 'VOLATILE') return '検索推移が不安定な記事群の代表';
  if (ext.indexOf('PLATFORM_OR_OS_CHANGE') >= 0) return '外部要因の影響を確認する代表';
  return '優先的に原因確認する代表記事';
}

function sdsdWriteSiteSummary_(rows, result) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SDSD_CONFIG.sheets.summary);
  if (!sh) sh = ss.insertSheet(SDSD_CONFIG.sheets.summary);
  sh.clear();

  const total = result.total || rows.length;
  const a1 = rows.filter(r => r.priority === 'A1_CANDIDATE').length;
  const a2 = rows.filter(r => r.priority === 'A2_CANDIDATE').length;
  const priorityCount = a1 + a2;
  const review = rows.filter(r =>
    r.priority === 'DOCTOR_REVIEW' || r.priority === 'REVIEW' ||
    r.priority === 'B_CANDIDATE' || r.priority === 'CANDIDATE'
  ).length;
  const sbm = rows.filter(r => r.priority === 'SBM').length;
  const protectedCount = rows.filter(r => r.priority === 'PROTECTED').length;
  const wait = rows.filter(r => r.priority === 'WAIT').length;

  const severe = rows.filter(r => r.weeklyTrend === 'SEVERE_DECLINE').length;
  const traffic = rows.filter(r => r.weeklyTrend === 'TRAFFIC_DECLINE').length;
  const ranking = rows.filter(r => r.weeklyTrend === 'RANKING_DECLINE').length;
  const volatile = rows.filter(r => r.weeklyTrend === 'VOLATILE').length;
  const growth = rows.filter(r => r.weeklyTrend === 'GROWTH').length;
  const external = rows.filter(r => String(r.externalFactor || '') !== '').length;

  let selectedCount = 0;
  let selectedUrls = {};
  try {
    const selected = ss.getSheetByName(SDSD_CONFIG.sheets.selectedCases);
    if (selected && selected.getLastRow() >= 2) {
      const vals = selected.getDataRange().getValues();
      const hdr = vals[0].map(String);
      const urlCol = hdr.indexOf('URL') >= 0 ? hdr.indexOf('URL') : hdr.indexOf('記事URL');
      if (urlCol >= 0) {
        vals.slice(1).forEach(r => {
          const u = String(r[urlCol] || '').trim();
          if (u) selectedUrls[u] = true;
        });
        selectedCount = Object.keys(selectedUrls).length;
      }
    }
  } catch (e) {}

  const priorityRows = rows.filter(r =>
    r.priority === 'A1_CANDIDATE' || r.priority === 'A2_CANDIDATE'
  );
  const notSelected = priorityRows.filter(r => !selectedUrls[String(r.url || '').trim()]);
  const notSelectedCount = Math.max(priorityCount - selectedCount, 0);

  let notSelectedReason = '';
  if (notSelectedCount > 0) {
    const guardCount = notSelected.filter(r =>
      String(r.guard || '').toUpperCase() !== '' &&
      String(r.guard || '').toUpperCase() !== 'PASS'
    ).length;
    const riskCount = notSelected.filter(r =>
      String(r.treatmentRisk || '').toUpperCase() === 'HIGH'
    ).length;

    if (guardCount > 0) {
      notSelectedReason = `最終確認・保護条件により今回は見送り ${guardCount}記事`;
    } else if (riskCount > 0) {
      notSelectedReason = `処置リスクが高いため今回は見送り ${riskCount}記事`;
    } else {
      notSelectedReason = `Treatment Batchの選定上限・優先順位により今回は見送り ${notSelectedCount}記事`;
    }
  }

  let overall = 'サイト全体を一律に修正する状態ではありません。';
  if (priorityCount > 0) {
    overall += ` 影響の大きい${priorityCount}記事を優先してDoctorで原因を確認します。`;
  }
  if (growth + protectedCount > 0) {
    overall += ' 回復・成長中の記事は保護し、不要な修正を避けます。';
  }

  const values = [
    ['サイト診断サマリー', ''],
    ['診断対象', `${total}記事`],
    ['Doctor精密診断の優先候補', `${priorityCount}記事（最優先 ${a1} / 優先 ${a2}）`],
    ['今回Doctorへ送る記事', selectedCount ? `${selectedCount}記事` : 'Treatment Batch作成前'],
    ['今回送らない優先候補', selectedCount ? `${notSelectedCount}記事` : 'Treatment Batch作成前'],
    ['今回送らない理由', selectedCount ? (notSelectedReason || 'なし') : 'Treatment Batch作成前'],
    ['追加確認が必要', `${review}記事`],
    ['SBMの日常改善向き', `${sbm}記事`],
    ['回復・成長中などの保護対象', `${protectedCount}記事`],
    ['最近処置済み・経過観察', `${wait}記事`],
    ['', ''],
    ['サイト全体の所見', overall],
    ['', ''],
    ['サイト全体の主な症状', ''],
    ['直近で大きく悪化', `${severe}記事`],
    ['検索流入低下', `${traffic}記事`],
    ['検索順位低下', `${ranking}記事`],
    ['推移が不安定', `${volatile}記事`],
    ['回復・成長傾向', `${growth}記事`],
    ['OS・サービス変更など外部要因の可能性', `${external}記事`],
    ['', ''],
    ['この診断の読み方', 'サイト全体を一律に修正するのではなく、優先候補をDoctorで精密診断し、共通原因と処置の必要性を確認します。'],
    ['注意', 'この段階では改善率を断定しません。実際の処置はDoctorの精密診断後に決めます。'],
    ['', ''],
    ['優先度の見方', ''],
    ['最優先', 'Doctorで詳しく診断する優先度が特に高い'],
    ['優先', 'Doctorでの精密診断を推奨'],
    ['要確認', '追加情報を確認してから判断'],
    ['日常改善', 'DoctorではなくSBMの日常改善で対応'],
    ['保護', '回復・成長中などのため今は大きく触らない'],
    ['経過観察', '最近の処置後などのため、しばらく推移を見る'],
    ['', ''],
    ['Doctor Case Package', '未生成']
  ];

  sh.getRange(1,1,values.length,2).setValues(values);
  sh.getRange('A1:B1').merge();
  sh.getRange('A1').setValue('サイト診断サマリー');
  sh.getRange('A1').setFontWeight('bold').setFontSize(14);
  sh.getRange('A12').setFontWeight('bold');
  sh.getRange('A14').setFontWeight('bold');
  sh.getRange('A26').setFontWeight('bold');
  sh.getRange(1,1,values.length,1).setFontWeight('bold');
  sh.setColumnWidth(1, 250);
  sh.setColumnWidth(2, 650);
  sh.getRange(1,1,values.length,2).setWrap(true);
  sh.setFrozenRows(1);

  try { ss.setActiveSheet(sh); } catch (e) {}
}

function sdsdUpdateSummaryAfterBatch_(batch, guard) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SDSD_CONFIG.sheets.summary);
  if (!sh) return;
  const values = sh.getRange(1,1,sh.getLastRow(),2).getValues();
  for (let i=0; i<values.length; i++) {
    if (String(values[i][0]) === '今回Doctorへ送る記事') {
      sh.getRange(i+1,2).setValue(
        `${batch.selectedCount}記事` +
        (guard.blocked ? `（最終確認で保留 ${guard.blocked}）` : '')
      );
      return;
    }
  }
}

function sdsdUpdateSummaryAfterPackage_(caseCount, fileUrl) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SDSD_CONFIG.sheets.summary);
  if (!sh) return;
  const values = sh.getRange(1,1,sh.getLastRow(),2).getValues();
  for (let i=0; i<values.length; i++) {
    if (String(values[i][0]) === 'Doctor Case Package') {
      sh.getRange(i+1,2).setValue(`${caseCount}件生成済み`);
      if (fileUrl) sh.getRange(i+1,2).setNote(fileUrl);
      return;
    }
  }
}


function sdsdHeaderIndexMap_(headers) {
  const map = {};
  headers.forEach((h,i) => map[String(h)] = i);
  return map;
}

function sdsdCandidateRowsFromSheet_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SDSD_CONFIG.sheets.candidates);
  if (!sh || sh.getLastRow() < 2) return [];

  const values = sh.getDataRange().getValues();
  const headers = values[0].map(String);
  const idx = sdsdHeaderIndexMap_(headers);

  if (idx['Normalized URL'] == null || idx['Priority Candidate'] == null) {
    return [];
  }

  return values.slice(1)
    .filter(r => String(r[idx['Normalized URL']] || ''))
    .map(r => ({
      url: String(r[idx['Normalized URL']] || ''),
      tvs: Number(r[idx['TVS']] || 0),
      demand: Number(r[idx['Demand']] || 0),
      opportunity: Number(r[idx['Opportunity']] || 0),
      urgency: Number(r[idx['Urgency']] || 0),
      asset: Number(r[idx['Asset Value']] || 0),
      ownership: String(r[idx['Ownership']] || ''),
      guard: String(r[idx['Recent Treatment Guard']] || ''),
      weeklyTrend: String(r[idx['Weekly Trend']] || ''),
      evidenceConfidence: String(r[idx['Evidence Confidence']] || ''),
      treatmentRisk: String(r[idx['Treatment Risk']] || ''),
      externalFactor: String(r[idx['External Factor']] || ''),
      priority: String(r[idx['Priority Candidate']] || ''),
      reason: String(r[idx['Reason']] || '')
    }));
}

function sdsdRefreshSiteSummaryFromCandidates_() {
  const rows = sdsdCandidateRowsFromSheet_();
  if (!rows.length) return false;

  const result = {
    total: rows.length,
    priorityCandidates: rows.filter(r =>
      r.priority === 'A1_CANDIDATE' || r.priority === 'A2_CANDIDATE'
    ).length,
    wait: rows.filter(r => r.priority === 'WAIT').length,
    protected: rows.filter(r => r.priority === 'PROTECTED').length,
    sbm: rows.filter(r => r.priority === 'SBM').length,
    review: rows.filter(r =>
      r.priority === 'REVIEW' || r.priority === 'DOCTOR_REVIEW' ||
      r.priority === 'B_CANDIDATE' || r.priority === 'CANDIDATE'
    ).length
  };

  sdsdWriteSiteSummary_(rows, result);
  return true;
}

function sdsdRefreshCandidatesView_() {
  const rows = sdsdCandidateRowsFromSheet_();
  if (!rows.length) return false;
  sdsdWriteCandidates_(rows);
  return true;
}

function sdsdRefreshSelectedCasesView_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SDSD_CONFIG.sheets.selectedCases);
  if (!sh || sh.getLastRow() < 2) return;

  const values = sh.getDataRange().getValues();
  const headers = values[0].map(String);
  const idx = sdsdHeaderIndexMap_(headers);

  if (idx['URL'] == null || idx['Referral JSON'] == null) return;

  const technicalHeaders = [
    'Batch Order','Site Priority','URL','TVS','Weekly Trend','Evidence Confidence',
    'Treatment Risk','External Factor','Ownership','Recent Treatment Guard',
    'Top Queries','Selection Reason','Referral Status','Referral JSON',
    'ArticleID','Article Title','Main Query','Article Fetch Status',
    'Case Package Status','Article Cache Key','Query Evidence Count'
  ].filter(h => idx[h] != null);

  const articleMap = sdsdArticleTitleMap_();
  const rows = values.slice(1).filter(r => String(r[idx['URL']] || ''));
  if (!rows.length) return;

  const userHeaders = [
    'No.','記事タイトル','記事URL','優先度','選定理由','サイト全体での意味'
  ];
  const newHeaders = userHeaders.concat(technicalHeaders);

  const newValues = rows.map((r,i) => {
    const url = String(r[idx['URL']] || '');
    const sitePriority = String(r[idx['Site Priority']] || '');
    const priorityJa = sitePriority === 'A1' ? '最優先' : sitePriority === 'A2' ? '優先' : sitePriority || '要確認';
    const reason = String(r[idx['Selection Reason']] || '');
    const displayRow = {
      priority: sitePriority === 'A1' ? 'A1_CANDIDATE' : sitePriority === 'A2' ? 'A2_CANDIDATE' : sitePriority,
      weeklyTrend: String(r[idx['Weekly Trend']] || ''),
      externalFactor: String(r[idx['External Factor']] || ''),
      reason: reason
    };

    const titleFromSheet = idx['Article Title'] != null ? String(r[idx['Article Title']] || '') : '';
    const visible = [
      i+1,
      titleFromSheet || sdsdDisplayTitle_(url, articleMap),
      url,
      priorityJa,
      sdsdReasonJa_(reason),
      sdsdSiteMeaning_(displayRow)
    ];
    const tech = technicalHeaders.map(h => r[idx[h]]);
    return visible.concat(tech);
  });

  sh.clear();
  sh.getRange(1,1,1,newHeaders.length).setValues([newHeaders]);
  sh.getRange(2,1,newValues.length,newHeaders.length).setValues(newValues);
  sh.setFrozenRows(1);
  sh.getRange(1,1,1,userHeaders.length).setFontWeight('bold');
  sh.setColumnWidth(1,70);
  sh.setColumnWidth(2,360);
  sh.setColumnWidth(3,320);
  sh.setColumnWidth(4,110);
  sh.setColumnWidth(5,460);
  sh.setColumnWidth(6,260);
  sh.getRange(1,1,Math.max(sh.getLastRow(),1),userHeaders.length).setWrap(true);
  sdsdHideTechnicalColumns_(sh, userHeaders.length + 1, newHeaders.length);
}

function sdsdArticleMasterCoverageForSelected_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SDSD_CONFIG.sheets.selectedCases);
  if (!sh || sh.getLastRow() < 2) return {selected:0, matched:0, withArticleId:0};

  const values = sh.getDataRange().getValues();
  const headers = values[0].map(String);
  const idx = sdsdHeaderIndexMap_(headers);
  if (idx['URL'] == null) return {selected:0, matched:0, withArticleId:0};

  const articleMap = sdsdBuildArticleMasterMap_();
  let selected = 0;
  let matched = 0;
  let withArticleId = 0;

  values.slice(1).forEach(r => {
    const raw = String(r[idx['URL']] || '');
    if (!raw) return;
    selected++;
    const master = articleMap[sdsdNormalizeUrl_(raw)] || null;
    if (master) {
      matched++;
      if (String(master.articleId || '')) withArticleId++;
    }
  });

  return {selected:selected, matched:matched, withArticleId:withArticleId};
}

function sdsdProgress_(step, total, message) {
  try {
    SpreadsheetApp.getActive().toast(
      `Step ${step}/${total}  ${message}`,
      'SIMS Doctor Site Diagnosis',
      8
    );
    SpreadsheetApp.flush();
  } catch (e) {}
}

function sdsdHideInternalSheets_() {
  const ss = SpreadsheetApp.getActive();
  const current = ss.getActiveSheet();

  if (current && SDSD_INTERNAL_SHEET_NAMES_.indexOf(current.getName()) >= 0) {
    const visible =
      ss.getSheetByName(SDSD_CONFIG.sheets.summary) ||
      ss.getSheetByName(SDSD_CONFIG.sheets.candidates) ||
      ss.getSheetByName(SDSD_CONFIG.sheets.selectedCases);
    if (visible && !visible.isSheetHidden()) {
      ss.setActiveSheet(visible);
    }
  }

  SDSD_INTERNAL_SHEET_NAMES_.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    try { sh.hideSheet(); } catch (e) {}
  });
}

function sdsdHideInternalSheets() {
  sdsdHideInternalSheets_();
  SpreadsheetApp.getUi().alert(
    '内部処理用シートを非表示にしました。\n通常利用では「診断候補」と「選定案件」だけ確認すれば大丈夫です。'
  );
}

function sdsdShowInternalSheets() {
  const ss = SpreadsheetApp.getActive();
  let count = 0;
  SDSD_INTERNAL_SHEET_NAMES_.forEach(name => {
    if (name === 'シート1' || name === 'Sheet1') return;
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    try {
      sh.showSheet();
      count++;
    } catch (e) {}
  });
  SpreadsheetApp.getUi().alert(
    `保守確認用として内部シートを表示しました。\n表示: ${count}シート\n\n確認後は「内部シートを隠す」を実行してください。`
  );
}

function sdsdRunProductDiagnosis() {
  sdsdHideInternalSheets_();
  try {
    const result = sdsdRunAnalysis({silent:true});
    sdsdHideInternalSheets_();

    SpreadsheetApp.getUi().alert(
      `サイト診断が完了しました。\n\n` +
      `対象記事: ${result.total}件\n` +
      `Doctor精密診断の優先候補: ${result.priorityCandidates}件\n` +
      `回復・成長中のため保護: ${result.protected}件\n` +
      `SBMの日常改善対象: ${result.sbm}件\n\n` +
      `「3. 診断結果を見る」でブログ全体の状態を確認してください。\n` +
      `詳しい記事一覧は「4. 診断候補を見る」で確認できます。`
    );
  } catch (e) {
    SpreadsheetApp.getUi().alert(
      `サイト診断を完了できませんでした。\n\n${e.message || e}`
    );
    throw e;
  }
}

function sdsdCreateProductTreatmentBatch() {
  try {
    sdsdProgress_(1, 2, '治療候補を選定しています');
    const batch = sdsdBuildTreatmentBatch({silent:true, noActivate:true});

    sdsdProgress_(2, 2, '最近の処置履歴を最終確認しています');
    const guard = sdsdRunFinalGuard({silent:true});
    sdsdUpdateSummaryAfterBatch_(batch, guard);
    sdsdHideInternalSheets_();

    SpreadsheetApp.getUi().alert(
      `Treatment Batchを作成しました。\n\n` +
      `診断対象: ${batch.articleCount}記事\n` +
      `Doctor精密診断の適格候補: ${batch.eligibleCount}件\n` +
      `今回Doctorへ送る記事: ${batch.selectedCount}件\n` +
      `最終確認で保留: ${guard.blocked}件\n\n` +
      `次に「個別記事の精密診断 → 2. 選定記事を見る」で内容を確認してください。`
    );
  } catch (e) {
    SpreadsheetApp.getUi().alert(
      `Treatment Batchを作成できませんでした。\n\n${e.message || e}`
    );
    throw e;
  }
}

function sdsdCreateProductCasePackage() {
  try {
    sdsdRefreshSelectedCasesView_();

    const coverage = sdsdArticleMasterCoverageForSelected_();
    if (coverage.selected === 0) {
      SpreadsheetApp.getUi().alert(
        'Doctor Case Packageを生成できません。\n\n今回の診断対象がありません。'
      );
      return;
    }

    if (coverage.withArticleId < coverage.selected) {
      SpreadsheetApp.getUi().alert(
        `Doctor Case Packageを生成する前に、今回のサイトの「記事管理」データが必要です。\n\n` +
        `今回の診断対象: ${coverage.selected}件\n` +
        `Article MasterでURL一致: ${coverage.matched}件\n` +
        `ArticleIDまで確認できた記事: ${coverage.withArticleId}件\n\n` +
        `「初期設定・データ準備 → Article Masterの取込案内」から、` +
        `今回診断しているブログのSBM「記事管理」CSVを取り込んでください。\n\n` +
        `取り込んだ後は前の処理をやり直さず、「個別記事の精密診断 → 3. Doctor精密診断Packageを生成」をもう一度実行できます。`
      );
      return;
    }

    sdsdProgress_(1, 3, 'Doctor Case Packageを準備しています');
    const enrichment = sdsdEnrichSelectedCases({
      silent: true,
      maxPerRun: 3
    });

    sdsdRefreshSelectedCasesView_();

    if (enrichment.review > 0) {
      SpreadsheetApp.getUi().alert(
        `Doctor Case Packageの準備中に要確認記事が見つかりました。\n\n` +
        `準備完了: ${enrichment.ready}/${enrichment.total}件\n` +
        `要確認: ${enrichment.review}件\n` +
        `未処理: ${enrichment.pending}件\n\n` +
        `「個別記事の精密診断 → 2. 選定記事を見る」で要確認案件を確認してください。`
      );
      return;
    }

    if (!enrichment.complete) {
      SpreadsheetApp.getUi().alert(
        `Doctor Case Packageを準備しています。\n\n` +
        `準備完了: ${enrichment.ready}/${enrichment.total}件\n` +
        `未処理: ${enrichment.pending}件\n\n` +
        `今回の処理結果は保存しました。\n` +
        `もう一度「個別記事の精密診断 → 3. Doctor精密診断Packageを生成」を実行すると、未処理の記事から続けます。`
      );
      return;
    }

    sdsdProgress_(2, 3, '全記事の準備完了。ZIPを生成しています');
    const exported = sdsdExportDoctorCasePackageZip({silent:true});

    sdsdUpdateSummaryAfterPackage_(exported.caseCount, exported.fileUrl);
    sdsdRefreshSelectedCasesView_();
    sdsdProgress_(3, 3, 'Doctor Case Packageを保存しました');

    SpreadsheetApp.getUi().alert(
      `Doctor Case Packageの生成が完了しました。\n\n` +
      `案件数: ${exported.caseCount}件\n` +
      `本文再取得: ${exported.refetched}件\n\n` +
      `Google DriveにZIPを保存しました。\n${exported.fileUrl}\n\n` +
      `このZIPをSIMS Doctorへ渡してください。`
    );
  } catch (e) {
    SpreadsheetApp.getUi().alert(
      `Doctor Case Packageを生成できませんでした。\n\n${e.message || e}`
    );
    throw e;
  }
}

// ============================================================================
// Source: ArticleFetcher.gs
// ============================================================================
function sdsdFetchArticleEvidence_(url) {
  const res = UrlFetchApp.fetch(url, {
    muteHttpExceptions: true,
    followRedirects: true,
    headers: {
      'User-Agent': 'Mozilla/5.0 (compatible; SIMS-Doctor-Site-Diagnosis/0.3.1)'
    }
  });

  const code = res.getResponseCode();
  if (code < 200 || code >= 400) {
    return {
      status: 'FETCH_ERROR',
      httpStatus: code,
      finalUrl: url,
      title: '',
      metaDescription: '',
      canonicalUrl: '',
      articleHtml: '',
      pageHtml: '',
      error: `HTTP ${code}`
    };
  }

  const html = res.getContentText();
  const title = sdsdDecodeHtml_(sdsdFirstMatch_(html, /<title[^>]*>([\s\S]*?)<\/title>/i));
  const metaDescription = sdsdDecodeHtml_(
    sdsdFirstMatch_(html, /<meta[^>]+name=["']description["'][^>]+content=["']([\s\S]*?)["'][^>]*>/i) ||
    sdsdFirstMatch_(html, /<meta[^>]+content=["']([\s\S]*?)["'][^>]+name=["']description["'][^>]*>/i)
  );
  const canonicalUrl =
    sdsdFirstMatch_(html, /<link[^>]+rel=["']canonical["'][^>]+href=["']([^"']+)["'][^>]*>/i) ||
    sdsdFirstMatch_(html, /<link[^>]+href=["']([^"']+)["'][^>]+rel=["']canonical["'][^>]*>/i);

  let articleHtml =
    sdsdFirstMatch_(html, /(<article\b[\s\S]*?<\/article>)/i) ||
    sdsdFirstMatch_(html, /(<div[^>]+class=["'][^"']*entry-content[^"']*["'][^>]*>[\s\S]*?<\/div>)/i) ||
    '';

  if (!articleHtml) {
    // Fail visibly rather than pretending the whole page is the article body.
    return {
      status: 'BODY_NOT_FOUND',
      httpStatus: code,
      finalUrl: canonicalUrl || url,
      title,
      metaDescription,
      canonicalUrl,
      articleHtml: '',
      pageHtml: '',
      error: 'Article body container not found.'
    };
  }

  return {
    status: 'VALID',
    httpStatus: code,
    finalUrl: canonicalUrl || url,
    title,
    metaDescription,
    canonicalUrl,
    articleHtml,
    pageHtml: '',
    error: ''
  };
}

function sdsdFirstMatch_(text, re) {
  const m = String(text || '').match(re);
  return m ? String(m[1] || '').trim() : '';
}

function sdsdDecodeHtml_(text) {
  return String(text || '')
    .replace(/&amp;/g,'&')
    .replace(/&quot;/g,'"')
    .replace(/&#39;/g,"'")
    .replace(/&lt;/g,'<')
    .replace(/&gt;/g,'>')
    .replace(/&nbsp;/g,' ');
}

// ============================================================================
// Source: ArticleMaster.gs
// ============================================================================
function sdsdBuildArticleMasterMap_() {
  const rows = sdsdReadObjects_(SDSD_CONFIG.sheets.articleMaster);
  const map = {};

  rows.forEach(r => {
    const rawUrl =
      r['記事URL'] || r['URL'] || r['url'] || r['Url'] ||
      r['公開URL'] || r['ページURL'] || '';
    const url = sdsdNormalizeUrl_(rawUrl);
    if (!url) return;

    const articleId =
      r['ArticleID'] || r['Article ID'] || r['記事ID'] || r['記事Id'] ||
      r['article_id'] || '';

    const title =
      r['記事タイトル'] || r['タイトル'] || r['Title'] || r['title'] || '';

    const mainQuery =
      r['メインクエリ'] || r['主要クエリ'] || r['Main Query'] ||
      r['main_query'] || '';

    const state =
      r['状態'] || r['記事状態'] || r['Status'] || r['status'] || '';

    map[url] = {
      articleId: String(articleId || ''),
      title: String(title || ''),
      mainQuery: String(mainQuery || ''),
      state: String(state || ''),
      raw: r
    };
  });

  return map;
}

function sdsdArticleMasterImportHelp() {
  SpreadsheetApp.getUi().alert(
    'Doctor Case Packageを作るには、今回診断しているブログのSBM「記事管理」データが必要です。\n\n' +
    '1. SBMで「記事管理」シートをCSVとして保存します。\n' +
    '2. Site Diagnosisの内部シート「_SDSD_ARTICLE_MASTER」へインポートします。\n' +
    '3. Googleスプレッドシートの「ファイル → インポート → アップロード → 現在のシートを置換」を使用します。\n\n' +
    'ArticleID・記事タイトル・URL・メインクエリをCase Packageへ利用します。\n' +
    '別ブログの記事管理データではCase Packageを生成できません。'
  );
}

// ============================================================================
// Source: ArticleUniverse.gs
// ============================================================================
function sdsdLooksLikeArticleUrl_(url) {
  const s = String(url || '').trim().split('#')[0].split('?')[0];

  // Hatena Blog article URL.
  if (/^https?:\/\/[^\/]+\/entry\/\d{4}\/\d{2}\/\d{2}\/[^\/?#]+\/?$/i.test(s)) return true;

  // Common WordPress numeric permalink.
  if (/^https?:\/\/[^\/]+\/\d+\/?$/i.test(s)) return true;

  // Dated permalink fallback.
  if (/^https?:\/\/[^\/]+\/\d{4}\/\d{2}\/\d{2}\/.+/i.test(s)) return true;

  return false;
}

function sdsdBuildArticleUniverse_() {
  // Preferred source: page_query_top is already narrowed by Collector RC5.
  const queryRows = sdsdReadObjects_(SDSD_CONFIG.sheets.evidencePageQuery);
  const fromPageQuery = {};
  queryRows.forEach(r => {
    const raw = r.page || r.key || r.url || '';
    const url = sdsdNormalizeUrl_(raw);
    if (url) fromPageQuery[url] = true;
  });

  // If Collector provided a plausible active-page set, use it as the primary universe.
  const pqUrls = Object.keys(fromPageQuery);
  if (pqUrls.length >= 20) {
    return {
      strategy: 'COLLECTOR_PAGE_QUERY_UNIVERSE',
      urls: fromPageQuery,
      count: pqUrls.length
    };
  }

  // Fallback: derive article-shaped URLs from page_summary.
  const summaryRows = sdsdReadObjects_(SDSD_CONFIG.sheets.evidencePageSummary);
  const heuristic = {};
  summaryRows.forEach(r => {
    const raw = r.key || r.page || r.url || '';
    if (!sdsdLooksLikeArticleUrl_(raw)) return;
    const url = sdsdNormalizeUrl_(raw);
    if (url) heuristic[url] = true;
  });

  return {
    strategy: 'ARTICLE_URL_HEURISTIC',
    urls: heuristic,
    count: Object.keys(heuristic).length
  };
}

// ============================================================================
// Source: CaseIdentity.gs
// ============================================================================
function sdsdShortHash_(text) {
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(text || '')
  );
  return digest
    .slice(0, 5)
    .map(b => ('0' + ((b + 256) % 256).toString(16)).slice(-2))
    .join('')
    .toUpperCase();
}

function sdsdCreateBatchId_() {
  const tz = Session.getScriptTimeZone() || 'Asia/Tokyo';
  const stamp = Utilities.formatDate(new Date(), tz, 'yyyyMMdd-HHmmss');
  const nonce = Utilities.getUuid().replace(/-/g, '').slice(0, 6).toUpperCase();
  const batchId = `${stamp}-${nonce}`;
  PropertiesService.getDocumentProperties().setProperty('SDSD_ACTIVE_BATCH_ID', batchId);
  return batchId;
}

function sdsdGetActiveBatchId_() {
  return String(
    PropertiesService.getDocumentProperties().getProperty('SDSD_ACTIVE_BATCH_ID') || ''
  );
}

function sdsdBuildSiteDiagnosisCaseId_(batchId, url) {
  if (!batchId) throw new Error('Site Diagnosis Batch ID がありません。Treatment Batchを作り直してください。');
  return `SDC-${batchId}-${sdsdShortHash_(url)}`;
}

function sdsdBuildIndividualCaseId_(batchId, articleId) {
  if (!batchId) throw new Error('Site Diagnosis Batch ID がありません。Treatment Batchを作り直してください。');
  if (!articleId) throw new Error('ArticleID がないため Individual Case ID を生成できません。');
  return `CASE-${batchId}-${String(articleId).replace(/[^A-Za-z0-9._-]/g, '_')}`;
}

function sdsdBuildRequestId_(batchId, articleId) {
  if (!batchId) throw new Error('Site Diagnosis Batch ID がありません。Treatment Batchを作り直してください。');
  if (!articleId) throw new Error('ArticleID がないため Request ID を生成できません。');
  return `REQ-${batchId}-${String(articleId).replace(/[^A-Za-z0-9._-]/g, '_')}`;
}

function sdsdSiteIdFromMaster_(master) {
  if (!master || !master.raw) return '';
  const r = master.raw;
  const value =
    r['SiteID'] || r['Site ID'] || r['サイトID'] || r['site_id'] ||
    r['SiteId'] || r['siteId'] || '';
  return String(value || '').trim();
}

function sdsdSiteIdFromUrl_(url) {
  const raw = String(url || '').trim();
  if (!raw) return '';

  let host = '';
  try {
    const m = raw.match(/^https?:\/\/([^\/:?#]+)/i);
    host = m ? String(m[1]).toLowerCase() : '';
  } catch(e) {}
  if (!host) return '';

  host = host.replace(/^www\./, '');
  const parts = host.split('.').filter(Boolean);
  if (!parts.length) return '';

  // Hatena Blog subdomains map naturally to the SBM site identifier used in testing.
  if (host.endsWith('.hatenablog.com') && parts.length >= 3) {
    return parts[0];
  }

  // Custom domains: use the registrable-domain label as a deterministic fallback.
  // An explicit SiteID in SBM Article Master always takes precedence.
  if (parts.length >= 2) return parts[parts.length - 2];
  return parts[0];
}

function sdsdResolveSiteId_(master, url) {
  const explicit = sdsdSiteIdFromMaster_(master);
  if (explicit) return explicit;

  const fallback = sdsdSiteIdFromUrl_(url);
  if (!fallback) {
    throw new Error(`SiteIDを確定できません: ${url}`);
  }
  return fallback;
}

// ============================================================================
// Source: CasePackageBuilder.gs
// ============================================================================
function sdsdArticleCacheKey_(articleId, url) {
  const raw = String(articleId || '') + '|' + String(url || '');
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw);
  return 'SDSD_ARTICLE_' + digest.map(b => ('0' + ((b + 256) % 256).toString(16)).slice(-2)).join('');
}

function sdsdEnrichSelectedCases(options) {
  options = options || {};
  const maxPerRun = Number(options.maxPerRun || 3);

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SDSD_CONFIG.sheets.selectedCases);
  if (!sh) throw new Error('先に Treatment Batch を生成してください。');

  const values = sh.getDataRange().getValues();
  if (values.length < 2) throw new Error('今回の診断対象に案件がありません。');

  const headers = values[0].map(String);
  const idx = {};
  headers.forEach((h,i) => idx[h] = i);

  const articleMap = sdsdBuildArticleMasterMap_();
  const queryMap = sdsdBuildQueryEvidenceMap_();
  const querySourceCount = sdsdQueryEvidenceSourceCount_();

  if (!Object.keys(articleMap).length) {
    throw new Error('記事管理データがありません。_SDSD_ARTICLE_MASTERへSBM「記事管理」CSVを入れてください。');
  }

  const extraHeaders = [
    'ArticleID','Article Title','Main Query','Article Fetch Status',
    'Case Package Status','Article Cache Key','Query Evidence Count'
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

  const firstTechnicalCol = headers.indexOf('Batch Order') + 1;
  if (firstTechnicalCol > 0) {
    sdsdHideTechnicalColumns_(sh, firstTechnicalCol, sh.getLastColumn());
  }

  // Refresh values after any missing technical columns were added.
  const data = sh.getDataRange().getValues();
  const refreshedHeaders = data[0].map(String);
  const col = {};
  refreshedHeaders.forEach((h,i) => col[h] = i);

  const cache = CacheService.getDocumentCache();
  let alreadyReady = 0;
  let processedThisRun = 0;
  let newlyReady = 0;
  let failed = 0;
  let remaining = 0;

  const candidateRows = [];

  for (let r=1; r<data.length; r++) {
    const url = String(data[r][col['URL']] || '');
    if (!url) continue;

    const referralStatus = String(data[r][col['Referral Status']] || '');
    const packageStatus = String(data[r][col['Case Package Status']] || '');

    if (referralStatus === 'READY_FOR_INDIVIDUAL_DOCTOR' && packageStatus === 'READY') {
      alreadyReady++;
      continue;
    }

    candidateRows.push(r);
  }

  // Work on only a small number per invocation. Existing READY rows are preserved.
  const workRows = candidateRows.slice(0, maxPerRun);
  remaining = Math.max(candidateRows.length - workRows.length, 0);

  for (let wi=0; wi<workRows.length; wi++) {
    const r = workRows[wi];
    const row = data[r];
    const url = String(row[col['URL']] || '');
    if (!url) continue;

    const master = articleMap[sdsdNormalizeUrl_(url)] || null;
    const articleId = master ? master.articleId : '';
    const title = master ? master.title : '';
    const mainQuery = master ? master.mainQuery : '';
    const queryEvidence = (queryMap[sdsdNormalizeUrl_(url)] || []).slice(0,10);

    // Missing identity/search prerequisites are cheap to detect; do not fetch the page.
    if (!master || !articleId ||
        (querySourceCount > 0 && queryEvidence.length === 0)) {
      let reviewReason = 'NEEDS_REVIEW';
      if (querySourceCount > 0 && queryEvidence.length === 0) {
        reviewReason = 'QUERY_EVIDENCE_MISSING';
      }

      sh.getRange(r+1, col['ArticleID']+1).setValue(articleId);
      sh.getRange(r+1, col['Article Title']+1).setValue(title);
      sh.getRange(r+1, col['Main Query']+1).setValue(mainQuery);
      sh.getRange(r+1, col['Query Evidence Count']+1).setValue(queryEvidence.length);
      sh.getRange(r+1, col['Case Package Status']+1).setValue(reviewReason);
      sh.getRange(r+1, col['Referral Status']+1).setValue('NEEDS_CASE_ENRICHMENT_REVIEW');

      failed++;
      processedThisRun++;
      continue;
    }

    const fetched = sdsdFetchArticleEvidence_(url);

    sh.getRange(r+1, col['ArticleID']+1).setValue(articleId);
    sh.getRange(r+1, col['Article Title']+1).setValue(title || fetched.title);
    sh.getRange(r+1, col['Main Query']+1).setValue(mainQuery);
    sh.getRange(r+1, col['Article Fetch Status']+1).setValue(fetched.status);
    sh.getRange(r+1, col['Query Evidence Count']+1).setValue(queryEvidence.length);

    if (col['Top Queries'] != null) {
      sh.getRange(r+1, col['Top Queries']+1).setValue(
        queryEvidence.map(q => q.query).join(' / ')
      );
    }

    if (fetched.status !== 'VALID') {
      sh.getRange(r+1, col['Case Package Status']+1).setValue('ARTICLE_FETCH_REVIEW');
      sh.getRange(r+1, col['Referral Status']+1).setValue('NEEDS_CASE_ENRICHMENT_REVIEW');
      failed++;
      processedThisRun++;
      continue;
    }

    const cacheKey = sdsdArticleCacheKey_(articleId, url);
    let cached = false;
    try {
      if (fetched.articleHtml.length < 90000) {
        cache.put(cacheKey, fetched.articleHtml, 21600);
        cached = true;
      }
    } catch(e) {}

    const oldReferral = String(row[col['Referral JSON']] || '{}');
    let referral = {};
    try { referral = JSON.parse(oldReferral); } catch(e) {}

    referral.format = 'SIMS_DOCTOR_INDIVIDUAL_CASE_PACKAGE_V1';
    referral.contract_version = '1.0';
    referral.case_identity = referral.case_identity || {};

    const batchId = String(
      referral.site_diagnosis_batch_id || sdsdGetActiveBatchId_() || ''
    );
    const siteId = sdsdResolveSiteId_(master, url);

    referral.site_diagnosis_batch_id = batchId;
    referral.case_identity.site_diagnosis_case_id = String(
      referral.case_identity.site_diagnosis_case_id ||
      sdsdBuildSiteDiagnosisCaseId_(batchId, url)
    );
    referral.case_identity.individual_case_id = String(
      referral.case_identity.individual_case_id ||
      sdsdBuildIndividualCaseId_(batchId, articleId)
    );
    referral.case_identity.site_id = siteId;
    referral.case_identity.article_id = articleId;
    referral.case_identity.url = url;
    referral.case_identity.request_id = String(
      referral.case_identity.request_id ||
      referral.request_id ||
      sdsdBuildRequestId_(batchId, articleId)
    );

    // RC4 Identity Contract: preserve canonical identifiers unchanged.
    referral.case_id = referral.case_identity.individual_case_id;
    referral.request_id = referral.case_identity.request_id;
    referral.site_diagnosis_case_id = referral.case_identity.site_diagnosis_case_id;
    referral.site_diagnosis_batch_id = batchId;
    referral.site_id = siteId;
    referral.article_id = articleId;
    referral.article_url = url;

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

    referral.search_evidence = referral.search_evidence || {};
    referral.search_evidence.evidence_window_days = 120;
    referral.search_evidence.top_queries = queryEvidence.map(q => ({
      query: q.query,
      clicks: q.clicks,
      impressions: q.impressions,
      ctr: q.ctr,
      position: q.position
    }));
    referral.search_evidence.query_count = queryEvidence.length;

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

    sh.getRange(r+1, col['Referral JSON']+1).setValue(JSON.stringify(referral));
    sh.getRange(r+1, col['Article Cache Key']+1).setValue(cacheKey);
    sh.getRange(r+1, col['Referral Status']+1).setValue('READY_FOR_INDIVIDUAL_DOCTOR');
    sh.getRange(r+1, col['Case Package Status']+1).setValue('READY');

    newlyReady++;
    processedThisRun++;
  }

  SpreadsheetApp.flush();

  // Recalculate current totals from the sheet after saving this chunk.
  const finalValues = sh.getDataRange().getValues();
  const finalHeaders = finalValues[0].map(String);
  const finalIdx = {};
  finalHeaders.forEach((h,i) => finalIdx[h] = i);

  let total = 0;
  let readyTotal = 0;
  let reviewTotal = 0;

  finalValues.slice(1).forEach(r => {
    if (!String(r[finalIdx['URL']] || '')) return;
    total++;
    const rs = String(r[finalIdx['Referral Status']] || '');
    const ps = String(r[finalIdx['Case Package Status']] || '');

    if (rs === 'READY_FOR_INDIVIDUAL_DOCTOR' && ps === 'READY') {
      readyTotal++;
    } else if (rs === 'NEEDS_CASE_ENRICHMENT_REVIEW') {
      reviewTotal++;
    }
  });

  const pending = Math.max(total - readyTotal - reviewTotal, 0);

  if (firstTechnicalCol > 0) {
    sdsdHideTechnicalColumns_(sh, firstTechnicalCol, sh.getLastColumn());
  }

  const result = {
    total: total,
    ready: readyTotal,
    review: reviewTotal,
    pending: pending,
    processedThisRun: processedThisRun,
    newlyReady: newlyReady,
    alreadyReady: alreadyReady,
    complete: pending === 0 && reviewTotal === 0
  };

  if (!options.silent) {
    SpreadsheetApp.getUi().alert(
      `Case Package準備\n\n` +
      `準備完了: ${result.ready}/${result.total}件\n` +
      `要確認: ${result.review}件\n` +
      `未処理: ${result.pending}件`
    );
  }

  return result;
}

function sdsdExportDoctorCasePackageZip(options) {
  options = options || {};
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

    const identity = parsed.case_identity || {};
    const requestId = String(parsed.request_id || identity.request_id || '');
    if (!identity.site_diagnosis_case_id || !identity.individual_case_id || !identity.site_id || !requestId) {
      throw new Error(`Case Identityが不完全です: ${articleId}`);
    }

    manifestCases.push({
      order: i+1,
      request_id: requestId,
      site_diagnosis_case_id: String(identity.site_diagnosis_case_id),
      individual_case_id: String(identity.individual_case_id),
      site_id: String(identity.site_id),
      article_id: articleId,
      url: url,
      priority: String(r[idx['Site Priority']] || ''),
      tvs: Number(r[idx['TVS']] || 0),
      case_file: `${folder}/case.json`,
      article_file: `${folder}/article.html`
    });
  });

  const batchIds = rows.map(r => {
    try {
      const p = JSON.parse(String(r[idx['Referral JSON']] || '{}'));
      return String(p.site_diagnosis_batch_id || '');
    } catch(e) { return ''; }
  }).filter(Boolean);
  const uniqueBatchIds = [...new Set(batchIds)];
  if (uniqueBatchIds.length !== 1) {
    throw new Error(`Treatment Batch IDが一意ではありません: ${uniqueBatchIds.join(', ')}`);
  }

  const manifest = {
    format: 'SIMS_DOCTOR_SITE_TREATMENT_BATCH_V1',
    contract_version: '1.1',
    site_diagnosis_batch_id: uniqueBatchIds[0],
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

  const result = {
    caseCount: rows.length,
    refetched: refetched,
    fileUrl: file.getUrl(),
    fileId: file.getId(),
    fileName: file.getName()
  };

  if (!options.silent) {
    SpreadsheetApp.getUi().alert(
      `Doctor Case Packageの生成が完了しました。\n\n` +
      `案件数: ${result.caseCount}件\n` +
      `本文再取得: ${result.refetched}件\n\n` +
      `Google Driveに保存しました。\n${result.fileUrl}`
    );
  }
  return result;
}

// ============================================================================
// Source: CaseSelection.gs
// ============================================================================
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
  const articleMap = sdsdArticleTitleMap_();
  const batchId = sdsdCreateBatchId_();

  let out = ss.getSheetByName(SDSD_CONFIG.sheets.selectedCases);
  if (!out) out = ss.insertSheet(SDSD_CONFIG.sheets.selectedCases);
  out.clear();

  const userHeaders = [
    'No.','記事タイトル','記事URL','優先度','選定理由','サイト全体での意味'
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

    const displayRow = {
      priority: String(r[idx['Priority Candidate']] || ''),
      weeklyTrend: String(r[idx['Weekly Trend']] || ''),
      externalFactor: String(r[idx['External Factor']] || ''),
      reason: reason
    };

    return [
      i+1,
      sdsdDisplayTitle_(url, articleMap),
      url,
      priorityJa,
      sdsdReasonJa_(reason),
      sdsdSiteMeaning_(displayRow),

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
  out.setColumnWidth(1, 70);
  out.setColumnWidth(2, 360);
  out.setColumnWidth(3, 320);
  out.setColumnWidth(4, 110);
  out.setColumnWidth(5, 460);
  out.setColumnWidth(6, 260);
  out.getRange(1,1,Math.max(out.getLastRow(),1),userHeaders.length).setWrap(true);
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
  try { sdsdRefreshSelectedCasesView_(); } catch (e) {}
  const sh = SpreadsheetApp.getActive().getSheetByName(SDSD_CONFIG.sheets.selectedCases);
  if (sh) SpreadsheetApp.getActive().setActiveSheet(sh);
}

// ============================================================================
// Source: EvidenceModel.gs
// ============================================================================
function sdsdBuildEvidenceMap_() {
  const rows = sdsdReadObjects_(SDSD_CONFIG.sheets.evidencePageSummary);
  const universe = sdsdBuildArticleUniverse_();
  const map = {};

  rows.forEach(r => {
    const raw = r.key || r.page || r.url || '';
    const url = sdsdNormalizeUrl_(raw);
    if (!url) return;

    // Gate: only score URLs that belong to the Article Universe.
    if (!universe.urls[url]) return;

    const item = map[url] || {
      url,
      clicksFull:0, impressionsFull:0,
      clicksRecent:0, impressionsRecent:0, positionRecent:0,
      clicksPrevious:0, impressionsPrevious:0, positionPrevious:0
    };

    item.clicksFull += Number((r.clicks_full ?? r.clicks_180d ?? r.clicks ?? 0));
    item.impressionsFull += Number((r.impressions_full ?? r.impressions_180d ?? r.impressions ?? 0));

    item.clicksRecent += Number((r.clicks_recent28 ?? r.clicks_recent28d ?? 0));
    item.impressionsRecent += Number((r.impressions_recent28 ?? r.impressions_recent28d ?? 0));
    item.positionRecent = Number((r.position_recent28 ?? r.position_recent28d ?? item.positionRecent ?? 0));

    item.clicksPrevious += Number((r.clicks_previous28 ?? r.clicks_previous28d ?? 0));
    item.impressionsPrevious += Number((r.impressions_previous28 ?? r.impressions_previous28d ?? 0));
    item.positionPrevious = Number((r.position_previous28 ?? r.position_previous28d ?? item.positionPrevious ?? 0));

    map[url] = item;
  });

  return {
    map,
    universeStrategy: universe.strategy,
    universeCount: universe.count
  };
}

// ============================================================================
// Source: DiagnosisSessionLifecycle.gs (RC8)
// ============================================================================
const SDSD_SESSION_PROP_KEYS_ = Object.freeze([
  'SDSD_SESSION_STATUS',
  'SDSD_SESSION_SITE_ID',
  'SDSD_SESSION_SITE_NAME',
  'SDSD_SESSION_SITE_URL',
  'SDSD_SESSION_HOST',
  'SDSD_SESSION_EVIDENCE_FILE_ID',
  'SDSD_SESSION_EVIDENCE_FILE_NAME',
  'SDSD_SESSION_STARTED_AT',
  'SDSD_LAST_EVIDENCE_FILE_ID',
  'SDSD_ACTIVE_BATCH_ID',
  'SDSD_LAST_SITE_WIDE_RESULT_AT',
  'SDSD_SITE_WIDE_REGISTER_STAGE',
  'SDSD_SITE_WIDE_REGISTER_DETAIL',
  'SDSD_SITE_WIDE_REGISTER_AT'
]);

function sdsdSheetDataRowCount_(sheetName) {
  const sh = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sh) return 0;
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const values = sh.getRange(2, 1, last - 1, Math.max(sh.getLastColumn(), 1)).getDisplayValues();
  return values.filter(r => r.some(v => String(v || '').trim() !== '')).length;
}

function sdsdInferCurrentEvidenceSite_() {
  const rows = sdsdReadObjects_(SDSD_CONFIG.sheets.evidencePageSummary);
  if (!rows.length) return {siteId:'', host:'', url:''};
  const row = rows[0] || {};
  const url = String(
    row.key || row.url || row.URL || row.page || row.Page || row['ページ'] || ''
  ).trim();
  if (!url) return {siteId:'', host:'', url:''};
  const m = url.match(/^https?:\/\/([^\/:?#]+)/i);
  const host = m ? String(m[1]).toLowerCase().replace(/^www\./, '') : '';
  return {siteId:sdsdSiteIdFromUrl_(url), host:host, url:url};
}

function sdsdGetCurrentSession_() {
  const props = PropertiesService.getDocumentProperties();
  const inferred = sdsdInferCurrentEvidenceSite_();
  const evidenceRows = sdsdSheetDataRowCount_(SDSD_CONFIG.sheets.evidencePageSummary);
  let status = String(props.getProperty('SDSD_SESSION_STATUS') || '').trim();
  if (!status && evidenceRows > 0) status = 'ACTIVE_LEGACY';
  return {
    active: status === 'ACTIVE' || status === 'ACTIVE_LEGACY',
    status: status,
    siteId: String(props.getProperty('SDSD_SESSION_SITE_ID') || inferred.siteId || ''),
    siteName: String(props.getProperty('SDSD_SESSION_SITE_NAME') || ''),
    siteUrl: String(props.getProperty('SDSD_SESSION_SITE_URL') || inferred.url || ''),
    host: String(props.getProperty('SDSD_SESSION_HOST') || inferred.host || ''),
    evidenceFileId: String(props.getProperty('SDSD_SESSION_EVIDENCE_FILE_ID') || props.getProperty('SDSD_LAST_EVIDENCE_FILE_ID') || ''),
    evidenceFileName: String(props.getProperty('SDSD_SESSION_EVIDENCE_FILE_NAME') || ''),
    startedAt: String(props.getProperty('SDSD_SESSION_STARTED_AT') || ''),
    evidenceRows: evidenceRows
  };
}

function sdsdSessionWorkSummary_() {
  const selected = sdsdSheetDataRowCount_(SDSD_CONFIG.sheets.selectedCases);
  const opportunityCases = sdsdSheetDataRowCount_(SDSD_CONFIG.sheets.opportunityCases);
  const candidates = sdsdSheetDataRowCount_(SDSD_CONFIG.sheets.candidates);

  let actionableTreatment = 0;
  let additionalEvidence = 0;
  const sh = SpreadsheetApp.getActive().getSheetByName(SDSD_CONFIG.sheets.treatmentPlan);
  if (sh && sh.getLastRow() >= 2) {
    const vals = sh.getDataRange().getDisplayValues();
    const headers = vals[0].map(x => String(x || '').trim());
    const nextIdx = headers.indexOf('次の処置');
    const stateIdx = headers.indexOf('状態');
    vals.slice(1).forEach(r => {
      if (!r.some(v => String(v || '').trim() !== '')) return;
      const next = nextIdx >= 0 ? String(r[nextIdx] || '') : '';
      const state = stateIdx >= 0 ? String(r[stateIdx] || '') : '';
      if (/Writer|Merge|Creator/.test(next)) actionableTreatment++;
      if (state.indexOf('追加確認待ち') >= 0 || next.indexOf('追加Evidence') >= 0) additionalEvidence++;
    });
  }

  return {
    candidates: candidates,
    selectedCases: selected,
    opportunityCases: opportunityCases,
    actionableTreatment: actionableTreatment,
    additionalEvidence: additionalEvidence,
    pendingTotal: selected + opportunityCases + actionableTreatment + additionalEvidence
  };
}

function sdsdShowCurrentSessionStatus() {
  sdsdProductEnsureSheets_();
  const ui = SpreadsheetApp.getUi();
  const session = sdsdGetCurrentSession_();
  const work = sdsdSessionWorkSummary_();
  if (!session.active) {
    ui.alert(
      '現在、診断中のサイトはありません。\n\n' +
      '「1. サイト全体を診断する → Evidence Packageを読み込む」から診断を開始してください。'
    );
    return;
  }
  ui.alert(
    '現在の診断セッション\n\n' +
    `対象サイト: ${session.siteId || session.host || '判定できません'}\n` +
    (session.host ? `ホスト: ${session.host}\n` : '') +
    (session.evidenceFileName ? `Evidence: ${session.evidenceFileName}\n` : '') +
    (session.startedAt ? `開始: ${session.startedAt}\n` : '') +
    `Evidence記事数: ${session.evidenceRows}件\n\n` +
    `診断候補: ${work.candidates}件\n` +
    `個別精密診断対象: ${work.selectedCases}件\n` +
    `サイト横断Doctor案件: ${work.opportunityCases}件\n` +
    `Writer/Merge/Creator振り分け: ${work.actionableTreatment}件\n` +
    `追加Evidence待ち: ${work.additionalEvidence}件\n\n` +
    '別サイトを診断する場合は、先に「診断セッション → 現在の診断を終了」を実行してください。'
  );
}

function sdsdClearDiagnosisSessionData_() {
  const ss = SpreadsheetApp.getActive();
  const names = [
    SDSD_CONFIG.sheets.evidencePageSummary,
    SDSD_CONFIG.sheets.evidencePageWeekly,
    SDSD_CONFIG.sheets.evidencePageQuery,
    SDSD_CONFIG.sheets.sbmHistory,
    SDSD_CONFIG.sheets.summary,
    SDSD_CONFIG.sheets.candidates,
    SDSD_CONFIG.sheets.selectedCases,
    SDSD_CONFIG.sheets.articleMaster,
    SDSD_CONFIG.sheets.opportunities,
    SDSD_CONFIG.sheets.opportunityCases,
    SDSD_CONFIG.sheets.siteWideResult,
    SDSD_CONFIG.sheets.treatmentPlan
  ];
  names.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (sh) sh.clearContents();
  });

  const merge = ss.getSheetByName(sdsdMergeReferralSheetName_());
  if (merge) merge.clearContents();

  const importSh = ss.getSheetByName(SDSD_CONFIG.sheets.siteWideResultImport);
  if (importSh) importSh.clearContents();
  try { sdsdEnsureSiteWideResultImportSheet_(); } catch (e) {}

  const props = PropertiesService.getDocumentProperties();
  SDSD_SESSION_PROP_KEYS_.forEach(k => props.deleteProperty(k));
}

function sdsdEndCurrentDiagnosisSession() {
  sdsdProductEnsureSheets_();
  const ui = SpreadsheetApp.getUi();
  const session = sdsdGetCurrentSession_();
  const work = sdsdSessionWorkSummary_();

  if (!session.active && work.pendingTotal === 0 && work.candidates === 0) {
    ui.alert('現在、終了する診断セッションはありません。');
    return;
  }

  let warning =
    `対象サイト: ${session.siteId || session.host || '判定できません'}\n\n` +
    'この操作を行うと、現在のEvidence・診断候補・診断結果・Article Master・SBM改善履歴の作業コピーをクリアします。\n' +
    'SBM本体に登録済みの改善履歴には影響しません。\n\n';

  if (work.pendingTotal > 0) {
    warning +=
      `要確認の作業データが ${work.pendingTotal}件あります。\n` +
      `・個別精密診断対象: ${work.selectedCases}件\n` +
      `・サイト横断Doctor案件: ${work.opportunityCases}件\n` +
      `・Writer/Merge/Creator振り分け: ${work.actionableTreatment}件\n` +
      `・追加Evidence待ち: ${work.additionalEvidence}件\n\n` +
      '必要な案件をDoctor/SBMへ引き渡したことを確認してから終了してください。\n\n';
  }

  warning += '現在の診断を終了して、次のサイトを診断できる状態にしますか？';
  const answer = ui.alert('現在の診断を終了', warning, ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) return;

  sdsdClearDiagnosisSessionData_();
  ui.alert(
    '現在の診断セッションを終了しました。\n\n' +
    '作業データをクリアしました。\n' +
    '次に「1. Evidence Packageを読み込む」から別サイトの診断を開始できます。'
  );
}

function sdsdRegisterDiagnosisSession_(fileId, fileName, packageMeta) {
  const props = PropertiesService.getDocumentProperties();
  const info = sdsdInferCurrentEvidenceSite_();
  props.setProperty('SDSD_SESSION_STATUS', 'ACTIVE');
  props.setProperty('SDSD_SESSION_SITE_ID', String((packageMeta&&packageMeta.siteId)||info.siteId||''));
  props.setProperty('SDSD_SESSION_SITE_NAME', String((packageMeta&&packageMeta.siteName)||''));
  props.setProperty('SDSD_SESSION_SITE_URL', String((packageMeta&&packageMeta.siteUrl)||info.url||''));
  props.setProperty('SDSD_SESSION_HOST', String(info.host || ''));
  props.setProperty('SDSD_SESSION_EVIDENCE_FILE_ID', String(fileId || ''));
  props.setProperty('SDSD_SESSION_EVIDENCE_FILE_NAME', String(fileName || ''));
  props.setProperty('SDSD_SESSION_STARTED_AT', new Date().toISOString());
  return info;
}

function sdsdAssertNoActiveDiagnosisSessionBeforeImport_() {
  const session = sdsdGetCurrentSession_();
  if (!session.active) return;
  const work = sdsdSessionWorkSummary_();
  throw new Error(
    '現在の診断セッションが残っています。\n\n' +
    `対象サイト: ${session.siteId || session.host || '判定できません'}\n` +
    `診断候補: ${work.candidates}件 / 要確認作業: ${work.pendingTotal}件\n\n` +
    '別のEvidenceを読み込む前に、メニュー「その他・管理 → 現在の診断を終了」を実行してください。\n' +
    'これにより、別サイトのデータが混在することを防ぎます。'
  );
}

// ============================================================================
// Source: EvidencePackageImporter.gs
// ============================================================================
function sdsdExtractDriveFileId_(text) {
  const s = String(text || '').trim();

  let m = s.match(/\/file\/d\/([A-Za-z0-9_-]+)/);
  if (m) return m[1];

  m = s.match(/[?&]id=([A-Za-z0-9_-]+)/);
  if (m) return m[1];

  if (/^[A-Za-z0-9_-]{20,}$/.test(s)) return s;

  return '';
}

function sdsdImportEvidencePackageZip(){
  sdsdProductEnsureSheets_();
  const ui=SpreadsheetApp.getUi();
  try{sdsdAssertNoActiveDiagnosisSessionBeforeImport_();}
  catch(e){ui.alert(String(e&&e.message?e.message:e));return;}

  const root=DriveApp.getRootFolder();
  const html=HtmlService.createHtmlOutput(sdsdEvidencePickerHtml_({folderId:root.getId(),folderName:'マイドライブ'}))
    .setWidth(700).setHeight(590);
  ui.showModalDialog(html,'Evidence Packageを選ぶ');
}

function sdsdListEvidencePickerFolder(folderId){
  let folder;
  try{folder=folderId?DriveApp.getFolderById(folderId):DriveApp.getRootFolder();}
  catch(e){folder=DriveApp.getRootFolder();}
  const folders=[],files=[];
  let it=folder.getFolders(),n=0;
  while(it.hasNext()&&n<150){const f=it.next();folders.push({id:f.getId(),name:f.getName()});n++;}
  let fit=folder.getFiles(),m=0;
  while(fit.hasNext()&&m<300){
    const f=fit.next(),name=f.getName();
    if(/\.zip$/i.test(name)&&(/SIMS/i.test(name)||/Evidence/i.test(name)))files.push({id:f.getId(),name:name,updated:f.getLastUpdated().toISOString()});
    m++;
  }
  folders.sort((a,b)=>a.name.localeCompare(b.name,'ja'));
  files.sort((a,b)=>String(b.updated).localeCompare(String(a.updated)));
  let parent=null;
  try{const ps=folder.getParents();if(ps.hasNext()){const p=ps.next();parent={id:p.getId(),name:p.getName()||'マイドライブ'};}}catch(e){}
  return {id:folder.getId(),name:folder.getName()||'マイドライブ',parent:parent,folders:folders,files:files};
}

function sdsdInspectEvidenceFile(fileId){
  const file=DriveApp.getFileById(fileId);
  if(!/\.zip$/i.test(file.getName()))throw new Error('ZIPファイルではありません。');
  const blobs=Utilities.unzip(file.getBlob());
  let manifest=null;
  blobs.forEach(b=>{if(String(b.getName()||'').split('/').pop()==='manifest.json'){try{manifest=JSON.parse(b.getDataAsString('UTF-8'));}catch(e){}}});
  const site=manifest&&manifest.site?manifest.site:{};
  const period=manifest&&manifest.period?manifest.period:{};
  return {
    fileId:file.getId(),fileName:file.getName(),
    siteName:String(site.siteName||site.site_name||''),
    siteUrl:String(site.siteUrl||site.site_url||site.searchConsoleProperty||''),
    generatedAt:String((manifest&&manifest.generatedAt)||''),
    periodLabel:period.days?String(period.days)+'日':(period.start&&period.end?period.start+' ～ '+period.end:''),
    format:String((manifest&&manifest.format)||'')
  };
}

function sdsdImportSelectedEvidence(payload){
  const fileId=String(payload&&payload.fileId||'');
  if(!fileId)throw new Error('Evidence Packageが選択されていません。');
  sdsdAssertNoActiveDiagnosisSessionBeforeImport_();
  return sdsdImportEvidencePackageById_(fileId);
}

function sdsdImportEvidencePackageById_(fileId){
  const file=DriveApp.getFileById(fileId);
  const name=file.getName();
  if(!/\.zip$/i.test(name))throw new Error(`ZIPファイルではありません: ${name}`);

  const blobs=Utilities.unzip(file.getBlob()),fileMap={};
  blobs.forEach(blob=>{const n=String(blob.getName()||'').split('/').pop();if(n)fileMap[n]=blob;});
  const required=[
    {file:'page_summary.csv',sheet:SDSD_CONFIG.sheets.evidencePageSummary},
    {file:'page_weekly.csv',sheet:SDSD_CONFIG.sheets.evidencePageWeekly},
    {file:'page_query_top.csv',sheet:SDSD_CONFIG.sheets.evidencePageQuery}
  ];
  const missing=required.filter(x=>!fileMap[x.file]).map(x=>x.file);
  if(missing.length)throw new Error(`Evidence ZIPに必要ファイルがありません: ${missing.join(', ')}`);

  let manifest={};
  if(fileMap['manifest.json']){try{manifest=JSON.parse(fileMap['manifest.json'].getDataAsString('UTF-8'));}catch(e){}}
  const site=manifest.site||{};
  const packageMeta={
    siteName:String(site.siteName||site.site_name||''),
    siteUrl:String(site.siteUrl||site.site_url||site.searchConsoleProperty||''),
    siteId:''
  };
  if(packageMeta.siteUrl)packageMeta.siteId=sdsdSiteIdFromUrl_(packageMeta.siteUrl);

  const report=[];
  required.forEach(x=>{
    const text=fileMap[x.file].getDataAsString('UTF-8').replace(/^\uFEFF/,'');
    const values=Utilities.parseCsv(text);
    if(!values.length)throw new Error(`${x.file} が空です。`);
    const sh=SpreadsheetApp.getActive().getSheetByName(x.sheet);sh.clearContents();
    const width=Math.max.apply(null,values.map(r=>r.length));
    const normalized=values.map(r=>{const row=r.slice();while(row.length<width)row.push('');return row;});
    sh.getRange(1,1,normalized.length,width).setValues(normalized);
    report.push({file:x.file,sheet:x.sheet,dataRows:Math.max(normalized.length-1,0),columns:width});
  });

  PropertiesService.getDocumentProperties().setProperty('SDSD_LAST_EVIDENCE_FILE_ID',fileId);
  const sessionSite=sdsdRegisterDiagnosisSession_(fileId,name,packageMeta);
  const diag=sdsdEvidenceImportIntegrity_();
  sdsdRenderHome_();
  return {
    ok:true,fileName:name,
    siteName:packageMeta.siteName||sessionSite.siteId||sessionSite.host||'判定できません',
    siteUrl:packageMeta.siteUrl||sessionSite.url||'',
    rows:report[0].dataRows,queryUrlCount:diag.queryUrlCount,queryRows:diag.queryRows,
    next:'次に「サイト全体を診断する → サイト診断を実行」へ進んでください。'
  };
}

function sdsdEvidencePickerHtml_(o){
  const data=JSON.stringify(o||{}).replace(/</g,'\\u003c');
  return `<!doctype html><html><head><base target="_top"><style>
  body{font-family:Arial,"Noto Sans JP",sans-serif;margin:0;background:#f8fafd;color:#202124}.wrap{padding:20px}
  .hero{background:#185abc;color:#fff;padding:16px 18px;border-radius:10px}.hero h2{margin:0 0 5px;font-size:20px}.hero p{margin:0;font-size:13px}
  .card{background:#fff;border:1px solid #dadce0;border-radius:10px;margin-top:14px;padding:14px}.bar{display:flex;gap:8px;align-items:center}.where{flex:1;font-weight:bold;color:#174ea6}
  button{border:1px solid #dadce0;background:#fff;border-radius:6px;padding:8px 12px;cursor:pointer}button.primary{background:#1a73e8;color:#fff;border-color:#1a73e8;font-weight:bold}
  .list{height:235px;overflow:auto;border:1px solid #e0e0e0;border-radius:7px;margin-top:10px}.row{padding:9px 11px;border-bottom:1px solid #f1f3f4;cursor:pointer}.row:hover{background:#f8f9fa}.selected{background:#e8f0fe!important}
  .meta{margin-top:12px;background:#f8fafd;border-radius:7px;padding:10px;line-height:1.7;font-size:13px}.hint{color:#5f6368;font-size:12px}.err{color:#b3261e;margin-top:8px}.actions{text-align:right;margin-top:12px}
  </style></head><body><div class="wrap"><div class="hero"><h2>Evidence Packageを読み込む</h2><p>Collectorで作成したZIPを、フォルダを移動して選択します。</p></div>
  <div class="card"><div class="bar"><button id="up">↑ 上へ</button><div id="where" class="where"></div></div><div id="list" class="list"></div><div class="hint">📁 フォルダをクリックして移動し、📦 Evidence ZIPを選択してください。</div></div>
  <div id="meta" class="meta">Evidence Packageを選択すると、サイト名・URL・作成日時・収集期間を確認できます。</div><div id="err" class="err"></div>
  <div class="actions"><button onclick="google.script.host.close()">キャンセル</button> <button id="import" class="primary" disabled>このEvidenceを読み込む</button></div>
  </div><script>
  const init=${data};let current=null,selected=null;
  const esc=s=>String(s||'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
  function fail(e){document.getElementById('err').textContent=e.message||e;}
  function load(id){document.getElementById('list').innerHTML='<div class="row">読み込み中...</div>';google.script.run.withSuccessHandler(render).withFailureHandler(fail).sdsdListEvidencePickerFolder(id);}
  function render(d){current=d;document.getElementById('where').textContent=d.name;document.getElementById('up').disabled=!d.parent;const box=document.getElementById('list');box.innerHTML='';
    d.folders.forEach(f=>{const x=document.createElement('div');x.className='row';x.textContent='📁 '+f.name;x.onclick=()=>load(f.id);box.appendChild(x);});
    d.files.forEach(f=>{const x=document.createElement('div');x.className='row';x.textContent='📦 '+f.name;x.onclick=()=>choose(f,x);box.appendChild(x);});
    if(!d.folders.length&&!d.files.length)box.innerHTML='<div class="row">このフォルダにEvidence ZIPはありません。</div>';
  }
  function choose(f,el){selected=f;document.querySelectorAll('.selected').forEach(x=>x.classList.remove('selected'));el.classList.add('selected');document.getElementById('import').disabled=true;
    document.getElementById('meta').textContent='内容を確認しています...';google.script.run.withSuccessHandler(m=>{document.getElementById('meta').innerHTML='<b>'+esc(m.fileName)+'</b><br>サイト名：'+esc(m.siteName||'不明')+'<br>サイトURL：'+esc(m.siteUrl||'不明')+'<br>作成日時：'+esc(m.generatedAt||'不明')+'<br>収集期間：'+esc(m.periodLabel||'不明');document.getElementById('import').disabled=false;}).withFailureHandler(fail).sdsdInspectEvidenceFile(f.id);}
  document.getElementById('up').onclick=()=>{if(current&&current.parent)load(current.parent.id)};
  document.getElementById('import').onclick=()=>{if(!selected)return;const b=document.getElementById('import');b.disabled=true;b.textContent='読み込み中...';google.script.run.withSuccessHandler(r=>{document.getElementById('meta').innerHTML='<b style="color:#137333">読み込み完了</b><br>サイト：'+esc(r.siteName)+'<br>'+esc(r.next);setTimeout(()=>google.script.host.close(),1800);}).withFailureHandler(e=>{b.disabled=false;b.textContent='このEvidenceを読み込む';fail(e)}).sdsdImportSelectedEvidence({fileId:selected.id});};
  load(init.folderId);
  </script></body></html>`;
}


function sdsdEvidenceImportIntegrity_() {
  const queryMap = sdsdBuildQueryEvidenceMap_();
  const queryRows = sdsdReadObjects_(SDSD_CONFIG.sheets.evidencePageQuery);

  return {
    queryRows: queryRows.length,
    queryUrlCount: Object.keys(queryMap).length
  };
}

// ============================================================================
// Source: EvidenceRisk.gs
// ============================================================================
function sdsdEvidenceConfidence_(x, weeklyInfo) {
  if (weeklyInfo && weeklyInfo.confidence === 'LOW') return 'LOW';
  if ((x.impressionsRecent || 0) < 50 && (x.impressionsFull || 0) < 200) return 'LOW';
  if ((x.impressionsRecent || 0) < 200) return 'MEDIUM';
  return weeklyInfo && weeklyInfo.confidence === 'MEDIUM' ? 'MEDIUM' : 'HIGH';
}

function sdsdTreatmentRisk_(x, weeklyInfo) {
  const highTraffic = (x.impressionsRecent || 0) >= 20000;
  const highRank = (x.positionRecent || 99) <= 5;
  const volatile = weeklyInfo && weeklyInfo.trend === 'VOLATILE';
  const growth = weeklyInfo && weeklyInfo.trend === 'GROWTH';

  if (growth) return 'HIGH';
  if (highTraffic && highRank) return 'HIGH';
  if (volatile || highTraffic) return 'MEDIUM';
  return 'LOW';
}

function sdsdExternalFactorFlag_(x, queryMap) {
  const queries = queryMap[x.url] || [];
  const text = queries.map(q => q.query).join(' ').toLowerCase();

  const flags = [];
  if (/chatgpt|openai/.test(text)) flags.push('FAST_CHANGING_SERVICE');
  if (/windows 11|windows11|iphone|ios|apple watch|line|instagram|インスタ|x /.test(text)) {
    flags.push('PLATFORM_OR_OS_CHANGE');
  }

  return flags;
}

// ============================================================================
// Source: FinalGuard.gs
// ============================================================================
function sdsdRunFinalGuard(options) {
  options = options || {};
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SDSD_CONFIG.sheets.selectedCases);
  if (!sh) throw new Error('先に Treatment Batch を生成してください。');

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return {checked:0, blocked:0};

  const headers = values[0].map(String);
  const idx = {};
  headers.forEach((h,i) => idx[h] = i);

  const historyMap = sdsdBuildHistoryMap_();
  let blocked = 0;

  for (let i=1; i<values.length; i++) {
    const url = String(values[i][idx['URL']] || '');
    if (!url) continue;

    const guard = sdsdRecentTreatmentGuard_(url, historyMap);
    if (guard.status === 'WAIT') {
      sh.getRange(i+1, idx['Recent Treatment Guard']+1).setValue('WAIT');
      sh.getRange(i+1, idx['Referral Status']+1).setValue('BLOCKED_BY_FINAL_GUARD');
      blocked++;
    }
  }

  // 利用者向け「状態」列も同期する。
  const userStatusIdx = headers.indexOf('状態');
  if (userStatusIdx >= 0 && values.length > 1) {
    const refreshed = sh.getDataRange().getValues();
    const refIdx = headers.indexOf('Referral Status');
    const statuses = refreshed.slice(1).map(r => {
      const ref = String(r[refIdx] || '');
      return [ref.indexOf('BLOCK') >= 0 ? '保留' : 'Doctor診断待ち'];
    });
    if (statuses.length) sh.getRange(2, userStatusIdx + 1, statuses.length, 1).setValues(statuses);
  }

  const result = {checked: values.length - 1, blocked: blocked};
  if (!options.silent) {
    SpreadsheetApp.getUi().alert(
      `最終確認が完了しました。\n再確認対象: ${result.checked}件\n保留: ${result.blocked}件`
    );
  }
  return result;
}

// ============================================================================
// Source: PriorityValidation.gs
// ============================================================================
function sdsdValidateFinalPriorities() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SDSD_CONFIG.sheets.candidates);
  if (!sh) throw new Error('先に Run Site Analysis を実行してください。');

  const values = sh.getDataRange().getValues();
  if (values.length < 2) throw new Error('候補データがありません。');

  const headers = values[0].map(String);
  const idx = {};
  headers.forEach((h,i) => idx[h] = i);

  const priorityCol = idx['Priority Candidate'];
  if (priorityCol == null) throw new Error('Priority Candidate列が見つかりません。');

  const rows = values.slice(1).filter(r => r[1]);
  const groups = {};
  rows.forEach(r => {
    const p = String(r[priorityCol] || 'EMPTY');
    if (!groups[p]) groups[p] = [];
    groups[p].push(r);
  });

  const order = ['A1_CANDIDATE','A2_CANDIDATE','B_CANDIDATE','DOCTOR_REVIEW',
                 'REVIEW','WAIT','SBM','PROTECTED','CANDIDATE','EMPTY'];

  const outName = 'Priority Validation';
  const ss = SpreadsheetApp.getActive();
  let out = ss.getSheetByName(outName);
  if (!out) out = ss.insertSheet(outName);
  out.clear();

  const summaryHeaders = ['Priority','Count','Share','Meaning','Gate Check'];
  out.getRange(1,1,1,summaryHeaders.length).setValues([summaryHeaders]);

  const meanings = {
    A1_CANDIDATE:'Doctor最優先候補',
    A2_CANDIDATE:'Doctor次点候補',
    B_CANDIDATE:'Doctor低優先候補',
    DOCTOR_REVIEW:'Doctor追加確認',
    REVIEW:'Evidence/Ownership追加確認',
    WAIT:'最近処置済み・モニター中',
    SBM:'SBMの日常改善へ',
    PROTECTED:'回復・成長中のため保護',
    CANDIDATE:'低優先候補',
    EMPTY:'未分類'
  };

  const summary = [];
  order.forEach(p => {
    const count = (groups[p] || []).length;
    if (!count) return;
    let gate = 'PASS';
    if (p === 'EMPTY') gate = 'FAIL';
    summary.push([p,count,count/rows.length,meanings[p] || '',gate]);
  });
  if (summary.length) out.getRange(2,1,summary.length,summaryHeaders.length).setValues(summary);

  const start = summary.length + 4;
  const detailHeaders = ['Priority','Rank','URL','TVS','Ownership','Guard','Weekly Trend',
                         'Evidence Confidence','Treatment Risk','External Factor','Reason'];
  out.getRange(start,1,1,detailHeaders.length).setValues([detailHeaders]);

  const sample = [];
  order.forEach(p => {
    (groups[p] || []).slice(0,5).forEach(r => sample.push([
      p,
      r[idx['Rank']],
      r[idx['Normalized URL']],
      r[idx['TVS']],
      r[idx['Ownership']],
      r[idx['Recent Treatment Guard']],
      r[idx['Weekly Trend']],
      r[idx['Evidence Confidence']],
      r[idx['Treatment Risk']],
      r[idx['External Factor']],
      r[idx['Reason']]
    ]));
  });
  if (sample.length) out.getRange(start+1,1,sample.length,detailHeaders.length).setValues(sample);

  out.setFrozenRows(1);
  out.getRange(2,3,Math.max(summary.length,1),1).setNumberFormat('0.0%');
  out.autoResizeColumns(1,detailHeaders.length);
  ss.setActiveSheet(out);

  SpreadsheetApp.getUi().alert(
    `最終Priority検証表を作成しました。\n対象記事: ${rows.length}\n` +
    `Priority Validation シートを確認してください。`
  );
}

// ============================================================================
// Source: QueryEvidence.gs
// ============================================================================
function sdsdObjectValue_(obj, names) {
  const keys = Object.keys(obj || {});
  const normalized = {};
  keys.forEach(k => {
    const nk = String(k || '').replace(/^\uFEFF/, '').trim().toLowerCase();
    normalized[nk] = obj[k];
  });

  for (let i=0; i<names.length; i++) {
    const key = String(names[i]).trim().toLowerCase();
    if (normalized[key] !== undefined && normalized[key] !== null) {
      return normalized[key];
    }
  }
  return '';
}

function sdsdBuildQueryEvidenceMap_() {
  const rows = sdsdReadObjects_(SDSD_CONFIG.sheets.evidencePageQuery);
  const map = {};

  rows.forEach(r => {
    const raw = sdsdObjectValue_(r, ['page','key','url','URL','記事URL']);
    const url = sdsdNormalizeUrl_(raw);
    if (!url) return;

    const query = String(sdsdObjectValue_(r, ['query','クエリ','検索クエリ']) || '').trim();
    if (!query) return;

    const arr = map[url] || [];
    arr.push({
      query: query,
      clicks: Number(sdsdObjectValue_(r, ['clicks','クリック数']) || 0),
      impressions: Number(sdsdObjectValue_(r, ['impressions','表示回数']) || 0),
      ctr: Number(sdsdObjectValue_(r, ['ctr','CTR']) || 0),
      position: Number(sdsdObjectValue_(r, ['position','掲載順位','平均掲載順位']) || 0)
    });
    map[url] = arr;
  });

  Object.keys(map).forEach(url => {
    map[url].sort((a,b) => {
      if (b.impressions !== a.impressions) return b.impressions - a.impressions;
      return b.clicks - a.clicks;
    });
  });

  return map;
}

function sdsdQueryEvidenceSourceCount_() {
  return sdsdReadObjects_(SDSD_CONFIG.sheets.evidencePageQuery).length;
}

// ============================================================================
// Source: QueryEvidenceDiagnostics.gs
// ============================================================================
function sdsdDiagnoseQueryEvidenceInput() {
  const ss = SpreadsheetApp.getActive();
  const sourceName = SDSD_CONFIG.sheets.evidencePageQuery;
  const sh = ss.getSheetByName(sourceName);
  if (!sh) throw new Error(`Query Evidenceシートが見つかりません: ${sourceName}`);

  const values = sh.getDataRange().getValues();
  const outName = 'Query Evidence Diagnostics';
  let out = ss.getSheetByName(outName);
  if (!out) out = ss.insertSheet(outName);
  out.clear();

  const report = [];
  report.push(['Check','Value']);

  if (!values.length) {
    report.push(['Source Sheet', sourceName]);
    report.push(['Raw Rows', 0]);
    out.getRange(1,1,report.length,2).setValues(report);
    ss.setActiveSheet(out);
    SpreadsheetApp.getUi().alert('Query Evidence Diagnostics完了\n元データ行数: 0');
    return;
  }

  const rawHeaders = values[0].map(v => String(v || ''));
  const normalizedHeaders = rawHeaders.map(h =>
    String(h || '').replace(/^\uFEFF/, '').trim().toLowerCase()
  );

  report.push(['Source Sheet', sourceName]);
  report.push(['Raw Rows', Math.max(values.length - 1, 0)]);
  report.push(['Raw Headers', rawHeaders.join(' | ')]);
  report.push(['Normalized Headers', normalizedHeaders.join(' | ')]);

  const objects = sdsdReadObjects_(sourceName);
  report.push(['Objects Read', objects.length]);

  let pageValues = 0;
  let queryValues = 0;
  let parsedRows = 0;
  const normalizedUrls = {};
  const queryMap = {};

  objects.forEach(r => {
    const page = sdsdObjectValue_(r, ['page','key','url','URL','記事URL']);
    const query = sdsdObjectValue_(r, ['query','クエリ','検索クエリ']);

    if (String(page || '').trim()) pageValues++;
    if (String(query || '').trim()) queryValues++;

    const url = sdsdNormalizeUrl_(page);
    if (url && String(query || '').trim()) {
      parsedRows++;
      normalizedUrls[url] = true;
      if (!queryMap[url]) queryMap[url] = 0;
      queryMap[url]++;
    }
  });

  report.push(['Rows with Page', pageValues]);
  report.push(['Rows with Query', queryValues]);
  report.push(['Parsed Page+Query Rows', parsedRows]);
  report.push(['Normalized URL Count', Object.keys(normalizedUrls).length]);

  const selected = ss.getSheetByName(SDSD_CONFIG.sheets.selectedCases);
  let selectedCount = 0;
  let matched = 0;
  let unmatched = [];

  if (selected) {
    const svals = selected.getDataRange().getValues();
    if (svals.length > 1) {
      const sheaders = svals[0].map(String);
      const idx = {};
      sheaders.forEach((h,i) => idx[h] = i);
      const urlIdx = idx['URL'];

      if (urlIdx != null) {
        svals.slice(1).forEach(r => {
          const u = sdsdNormalizeUrl_(r[urlIdx]);
          if (!u) return;
          selectedCount++;
          if (queryMap[u] > 0) matched++;
          else unmatched.push(u);
        });
      }
    }
  }

  report.push(['Selected Cases', selectedCount]);
  report.push(['Selected Cases Matched to Query Evidence', matched]);
  report.push(['Selected Cases Unmatched', unmatched.length]);

  out.getRange(1,1,report.length,2).setValues(report);

  const start = report.length + 3;
  out.getRange(start,1,1,3).setValues([['Unmatched URL','Normalized URL','Query Rows']]);
  if (unmatched.length) {
    const rows = unmatched.map(u => [u,u,queryMap[u] || 0]);
    out.getRange(start+1,1,rows.length,3).setValues(rows);
  }

  const sampleStart = start + Math.max(unmatched.length,1) + 3;
  out.getRange(sampleStart,1,1,4).setValues([['Sample URL','Query Count','First Query','First Impressions']]);

  const built = sdsdBuildQueryEvidenceMap_();
  const samples = Object.keys(built).slice(0,10).map(u => {
    const arr = built[u] || [];
    return [
      u,
      arr.length,
      arr[0] ? arr[0].query : '',
      arr[0] ? arr[0].impressions : ''
    ];
  });
  if (samples.length) {
    out.getRange(sampleStart+1,1,samples.length,4).setValues(samples);
  }

  out.setFrozenRows(1);
  out.autoResizeColumns(1,4);
  ss.setActiveSheet(out);

  SpreadsheetApp.getUi().alert(
    `Query Evidence Diagnostics完了\n\n` +
    `元データ行数: ${Math.max(values.length-1,0)}\n` +
    `Page認識: ${pageValues}\n` +
    `Query認識: ${queryValues}\n` +
    `URL正規化後: ${Object.keys(normalizedUrls).length}\n` +
    `選定18件一致: ${matched}/${selectedCount}`
  );
}

// ============================================================================
// Source: SBMHistoryImporter.gs
// ============================================================================
function sdsdBuildHistoryMap_() {
  const rows = sdsdReadObjects_(SDSD_CONFIG.sheets.sbmHistory);
  const map = {};
  rows.forEach(r => {
    const rawUrl = r['記事URL'] || r['URL'] || r['url'] || '';
    const url = sdsdNormalizeUrl_(rawUrl);
    if (!url) return;

    // SBM「改善履歴」シートの現行ヘッダーを最優先で読む。
    // 旧名称も後方互換として残す。
    const status = String(r['状態'] || r['作業状態'] || r['現在状態'] || r['判定'] || '');
    const dateText = r['改善実施日'] || r['改善日'] || r['処置日'] || r['日付'] || '';
    const route = String(r['改善経路'] || r['経路'] || '');

    const d = dateText ? new Date(dateText) : null;
    const prev = map[url];

    if (!prev || (d && (!prev.date || d > prev.date))) {
      map[url] = { date:d, status, route };
    }
  });
  return map;
}

function sdsdRecentTreatmentGuard_(url, historyMap) {
  const h = historyMap[url];
  if (!h) return {status:'PASS', reason:''};

  if (String(h.status).indexOf('モニター中') >= 0) {
    return {status:'WAIT', reason:'SBMモニター中'};
  }

  if (h.date && !isNaN(h.date.getTime())) {
    const days = (Date.now() - h.date.getTime()) / 86400000;
    if (days >= 0 && days <= SDSD_CONFIG.guardDays) {
      return {status:'WAIT', reason:`直近${Math.floor(days)}日前に改善済み`};
    }
  }
  return {status:'PASS', reason:''};
}

// ============================================================================
// Source: SheetStore.gs
// ============================================================================
function sdsdEnsureSheets_() {
  const ss = SpreadsheetApp.getActive();
  const names = [
    SDSD_CONFIG.sheets.evidencePageSummary,
    SDSD_CONFIG.sheets.evidencePageWeekly,
    SDSD_CONFIG.sheets.evidencePageQuery,
    SDSD_CONFIG.sheets.sbmHistory,
    SDSD_CONFIG.sheets.candidates,
    SDSD_CONFIG.sheets.selectedCases,
    SDSD_CONFIG.sheets.articleMaster
  ];
  names.forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (name !== SDSD_CONFIG.sheets.candidates && name !== SDSD_CONFIG.sheets.selectedCases) {
      try { sh.hideSheet(); } catch (e) {}
    }
  });
  try { sdsdHideInternalSheets_(); } catch (e) {}
}

function sdsdReadObjects_(sheetName) {
  const sh = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sh) return [];
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(String);
  return values.slice(1).filter(r => r.some(v => v !== '')).map(r => {
    const o = {};
    headers.forEach((h, i) => o[h] = r[i]);
    return o;
  });
}

function sdsdWriteCandidates_(rows) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SDSD_CONFIG.sheets.candidates);
  sh.clear();

  const userHeaders = [
    '優先順位','記事タイトル','記事URL','優先度','選定理由','対応方針'
  ];
  const technicalHeaders = [
    'Rank','Normalized URL','TVS','Demand','Opportunity','Urgency','Asset Value',
    'Ownership','Recent Treatment Guard','Weekly Trend','Evidence Confidence',
    'Treatment Risk','External Factor','Priority Candidate','Reason'
  ];
  const headers = userHeaders.concat(technicalHeaders);

  const articleMap = sdsdArticleTitleMap_();
  const displayRows = rows.slice().sort((a,b) => {
    const pa = sdsdPriorityRank_(a.priority);
    const pb = sdsdPriorityRank_(b.priority);
    if (pa !== pb) return pa - pb;
    return Number(b.tvs || 0) - Number(a.tvs || 0);
  });

  sh.getRange(1,1,1,headers.length).setValues([headers]);

  if (displayRows.length) {
    const values = displayRows.map((r,i) => [
      i+1,
      sdsdDisplayTitle_(r.url, articleMap),
      r.url,
      sdsdPriorityJa_(r.priority),
      sdsdReasonJa_(r.reason),
      sdsdActionJa_(r),

      i+1,r.url,r.tvs,r.demand,r.opportunity,r.urgency,r.asset,
      r.ownership,r.guard,r.weeklyTrend,r.evidenceConfidence,
      r.treatmentRisk,r.externalFactor,r.priority,r.reason
    ]);
    sh.getRange(2,1,values.length,headers.length).setValues(values);
  }

  sh.setFrozenRows(1);
  sh.getRange(1,1,1,userHeaders.length).setFontWeight('bold');
  sh.setColumnWidth(1, 90);
  sh.setColumnWidth(2, 360);
  sh.setColumnWidth(3, 320);
  sh.setColumnWidth(4, 110);
  sh.setColumnWidth(5, 460);
  sh.setColumnWidth(6, 170);
  sh.getRange(1,1,Math.max(sh.getLastRow(),1),userHeaders.length).setWrap(true);
  sh.getRange(1,4).setNote(
    '最優先: Doctor精密診断の優先度が特に高い\n' +
    '優先: Doctor精密診断を推奨\n' +
    '要確認: 追加情報を確認して判断\n' +
    '日常改善: SBMで対応\n' +
    '保護: 今は大きく触らない\n' +
    '経過観察: 推移を見る'
  );
  sdsdHideTechnicalColumns_(sh, userHeaders.length + 1, headers.length);
}

// ============================================================================
// Source: SiteAnalyzer.gs
// ============================================================================
function sdsdRunAnalysis(options) {
  options = options || {};
  sdsdProgress_(1, 5, 'Evidence Packageのデータを確認しています');
  sdsdProductEnsureSheets_();

  const evidence = sdsdBuildEvidenceMap_();
  const evidenceMap = evidence.map;

  sdsdProgress_(2, 5, '改善履歴と記事状態を確認しています');
  const historyMap = sdsdBuildHistoryMap_();

  sdsdProgress_(3, 5, '週次推移と検索クエリを分析しています');
  const weeklyMap = sdsdBuildWeeklyTrendMap_();
  const queryMap = sdsdBuildQueryEvidenceMap_();
  const items = Object.keys(evidenceMap).map(k => evidenceMap[k]);

  sdsdProgress_(4, 5, '診断候補の優先度を評価しています');
  let scored = sdsdScoreAll_(items);

  scored = scored.map(x => {
    const own = sdsdOwnership_(x);
    const guard = sdsdRecentTreatmentGuard_(x.url, historyMap);
    const weekly = sdsdClassifyWeeklyTrend_(weeklyMap[x.url] || []);
    const evidenceConfidence = sdsdEvidenceConfidence_(x, weekly);
    const treatmentRisk = sdsdTreatmentRisk_(x, weekly);
    const externalFlags = sdsdExternalFactorFlag_(x, queryMap);

    let priority = 'CANDIDATE';

    if (guard.status === 'WAIT') {
      priority = 'WAIT';
    } else if (own.ownership === 'SBM_OWNED') {
      priority = 'SBM';
    } else if (weekly.trend === 'GROWTH') {
      priority = 'PROTECTED';
    } else if (evidenceConfidence === 'LOW') {
      priority = 'REVIEW';
    } else if (own.ownership === 'REVIEW') {
      if (weekly.trend === 'SEVERE_DECLINE' || weekly.trend === 'TRAFFIC_DECLINE' || weekly.trend === 'RANKING_DECLINE') {
        priority = 'DOCTOR_REVIEW';
      } else {
        priority = 'REVIEW';
      }
    } else if (own.ownership === 'DOCTOR_OWNED') {
      if (treatmentRisk === 'HIGH' && weekly.trend !== 'SEVERE_DECLINE') {
        priority = 'DOCTOR_REVIEW';
      } else if (x.tvs >= 70) {
        priority = 'A1_CANDIDATE';
      } else if (x.tvs >= 60) {
        priority = 'A2_CANDIDATE';
      } else if (x.tvs >= 50) {
        priority = 'B_CANDIDATE';
      } else {
        priority = 'CANDIDATE';
      }
    } else {
      priority = 'REVIEW';
    }

    const reasonParts = [own.reason, guard.reason];
    if (weekly.trend && weekly.trend !== 'STABLE') reasonParts.push(`週次:${weekly.trend}`);
    if (externalFlags.length) reasonParts.push(`外部要因:${externalFlags.join('|')}`);
    if (treatmentRisk === 'HIGH') reasonParts.push('高Risk');

    return Object.assign({}, x, {
      ownership: own.ownership,
      guard: guard.status,
      weeklyTrend: weekly.trend,
      evidenceConfidence,
      treatmentRisk,
      externalFactor: externalFlags.join('|') || '',
      priority,
      reason: reasonParts.filter(Boolean).join(' / ')
    });
  });

  scored.sort((a,b) => b.tvs - a.tvs);
  sdsdProgress_(5, 5, '診断結果を整理しています');
  sdsdWriteCandidates_(scored);

  const result = {
    total: scored.length,
    priorityCandidates: scored.filter(x =>
      x.priority === 'A1_CANDIDATE' || x.priority === 'A2_CANDIDATE'
    ).length,
    wait: scored.filter(x => x.priority === 'WAIT').length,
    protected: scored.filter(x => x.priority === 'PROTECTED').length,
    sbm: scored.filter(x => x.priority === 'SBM').length,
    review: scored.filter(x =>
      x.priority === 'REVIEW' || x.priority === 'DOCTOR_REVIEW'
    ).length,
    universeCount: evidence.universeCount,
    universeStrategy: evidence.universeStrategy
  };

  sdsdWriteSiteSummary_(scored, result);

  if (!options.silent) {
    SpreadsheetApp.getUi().alert(
      `サイト診断が完了しました。\n\n` +
      `対象記事: ${result.total}件\n` +
      `精密診断の優先候補: ${result.priorityCandidates}件\n` +
      `最近処置済み・モニター中: ${result.wait}件\n` +
      `回復・成長中のため保護: ${result.protected}件\n\n` +
      `「${SDSD_CONFIG.sheets.candidates}」で結果を確認してください。`
    );
  }
  return result;
}

// ============================================================================
// Source: SiteOpportunityCaseBuilder.gs
// ============================================================================
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

// ============================================================================
// Source: SiteOpportunityDiagnosis.gs
// ============================================================================
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


function sdsdUserDiagnosisThemeLabel_(cluster) {
  const type = String(cluster.type || '');
  const q = String(cluster.parentTheme || cluster.theme || '').trim();
  const targets = cluster.targets || [];

  // Keep familiar product/service names concise.
  if (/nhk\s*プラス/i.test(q) || /nhk\s*プラス/i.test(String(cluster.diagnosisThemeKey || ''))) {
    return type === 'カニバリ疑い' ? 'NHKプラスのエラー・視聴トラブル記事群' : 'NHKプラスの検索ニーズ';
  }
  if (/獅子舞/.test(q)) return '獅子舞のご祝儀・相場';
  if (/line.*電話|line電話/i.test(q)) return 'LINE電話の充電切れ時の挙動';
  if (/まぜてる|ませてる/.test(q)) return '「ませてる」の意味・使い方';
  if (/自転車.*塗装|塗装.*自転車/.test(q)) return '自転車塗装の料金・相場';
  if (/vrchat/i.test(q)) return 'VRChatのPS5対応・遊び方';
  if (/u-?next/i.test(q)) return 'U-NEXTのミラーリング・テレビ視聴';
  if (/元ヤン/.test(q)) return '元ヤンの特徴・見分け方';

  // Generic cleanup: remove noisy modifiers and keep a readable topic.
  let label = q
    .replace(/[【\[].*?[】\]]/g, ' ')
    .replace(/\b20\d{2}年?(最新版|最新)?\b/g, ' ')
    .replace(/完全(対処|解決)?ガイド|完全版|徹底解説|完全解説/g, ' ')
    .replace(/原因と(直し方|対処法)|原因別(解決|対処)ガイド/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  if (label.length > 34) label = label.slice(0, 34) + '…';
  return label || sdsdBuildDiagnosisThemeLabel_(cluster);
}

function sdsdSiteImpactText_(cluster) {
  const type = String(cluster.type || '');
  const n = (cluster.targets || []).length;
  const qn = (cluster.queries || []).length;

  if (type === 'カニバリ疑い') {
    return `検索意図が重なる${n}記事の役割を整理し、評価の分散や順位競合を減らせる可能性があります。Doctorが統合・役割分担・維持を判断します。`;
  }
  if (type === '新規記事機会') {
    return `既存記事で十分に受け止められていない検索需要を補い、サイトの検索入口を増やせる可能性があります。Doctor確認後にCreator候補とします。`;
  }
  if (type === 'コンテンツギャップ') {
    return `既存記事に関連する未充足ニーズを補い、${qn}件の関連検索に対する網羅性と内部導線を強化できる可能性があります。`;
  }
  return 'サイト全体への影響をDoctorで確認します。';
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
    c.diagnosisTheme = sdsdUserDiagnosisThemeLabel_(c);
    c.siteImpact = sdsdSiteImpactText_(c);

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
    '優先順位','改善テーマ','診断テーマ','サイト全体への期待効果','親テーマ','検索テーマ','関連クエリ数',
    '対象記事','根拠','確信度','推奨対応','担当'
  ];

  const values = rows.map((x,i) => [
    i+1,
    x.type,
    x.diagnosisTheme || x.parentTheme || x.theme,
    x.siteImpact || 'サイト全体への影響をDoctorで確認します。',
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

  [70,150,260,420,200,220,100,430,520,90,360,120]
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
      `検索意図 → 親テーマ → 診断テーマの3段階で整理し、サイト全体への期待効果も表示しています。\n` +
      `これは一次候補であり、自動処置は行いません。`
    );
  } catch(e) {
    SpreadsheetApp.getUi().alert(
      `サイト横断診断を完了できませんでした。\n\n${e.message || e}`
    );
    throw e;
  }
}

// ============================================================================
// Source: SiteSizePolicy.gs
// ============================================================================
function sdsdTreatmentCapacity_(articleCount) {
  if (articleCount <= 99) return {standardMin:5, standardMax:8, hardMax:10};
  if (articleCount <= 199) return {standardMin:8, standardMax:12, hardMax:15};
  if (articleCount <= 399) return {standardMin:10, standardMax:15, hardMax:18};
  if (articleCount <= 599) return {standardMin:12, standardMax:18, hardMax:20};
  return {standardMin:15, standardMax:20, hardMax:25};
}

// ============================================================================
// Source: SiteWideDoctorPackage.gs
// ============================================================================

function sdsdExtractUrlsFromText_(text) {
  const matches = String(text || '').match(/https?:\/\/[^\s]+/g) || [];
  return matches.map(x => x.replace(/[),.;、。]+$/g, ''));
}

function sdsdSiteWideBatchId_() {
  const props = PropertiesService.getDocumentProperties();
  let id = String(props.getProperty('SDSD_SITE_WIDE_BATCH_ID') || '');
  if (id) return id;

  const tz = Session.getScriptTimeZone() || 'Asia/Tokyo';
  const stamp = Utilities.formatDate(new Date(), tz, 'yyyyMMdd-HHmmss');
  id = 'SITEWIDE-' + stamp + '-' +
    Utilities.getUuid().replace(/-/g,'').slice(0,6).toUpperCase();
  props.setProperty('SDSD_SITE_WIDE_BATCH_ID', id);
  return id;
}

function sdsdSiteMetaFromArticleMaster_() {
  const articleMap = sdsdBuildArticleMasterMap_();
  const urls = Object.keys(articleMap);
  if (!urls.length) return {
    site_id: '',
    site_name: '',
    site_url: ''
  };

  const firstUrl = urls[0];
  const master = articleMap[firstUrl];
  const raw = master && master.raw ? master.raw : {};

  const siteId = sdsdResolveSiteId_(master, firstUrl);
  const siteName = String(
    raw['SiteName'] || raw['Site Name'] || raw['サイト名'] ||
    raw['ブログ名'] || raw['site_name'] || siteId || ''
  ).trim();

  let siteUrl = '';
  const explicit = String(
    raw['BlogURL'] || raw['ブログURL'] || raw['SiteURL'] ||
    raw['サイトURL'] || raw['site_url'] || ''
  ).trim();

  if (explicit) {
    siteUrl = explicit;
  } else {
    const m = firstUrl.match(/^(https?:\/\/[^\/]+)/i);
    siteUrl = m ? m[1] + '/' : '';
  }

  return {
    site_id: siteId,
    site_name: siteName,
    site_url: siteUrl
  };
}

function sdsdArticleIdentityForUrl_(url) {
  const map = sdsdBuildArticleMasterMap_();
  const n = sdsdNormalizeUrl_(url);
  const master = map[n] || null;

  return {
    site_id: master ? sdsdResolveSiteId_(master, url) : sdsdSiteIdFromUrl_(url),
    article_id: master ? String(master.articleId || '') : '',
    article_title: master ? String(master.title || '') : '',
    article_url: url,
    main_query: master ? String(master.mainQuery || '') : ''
  };
}

/**
 * SIMS Doctor Site Diagnosis v0.5.0 Sprint 3
 * Export one site-wide Doctor package preserving site context.
 */

function sdsdSiteWidePackageRows_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SDSD_CONFIG.sheets.opportunityCases);
  if (!sh || sh.getLastRow() < 2) {
    throw new Error('先に「サイト横断の診断・治療 → 2. Doctor診断案件を作成」を実行してください。');
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
    .map(r => {
      const targetText = String(r[idx['対象記事']] || '');
      const urls = sdsdExtractUrlsFromText_(targetText);
      return {
        case_id: String(r[idx['案件ID']] || ''),
        improvement_type: String(r[idx['改善テーマ']] || ''),
        diagnosis_theme: String(r[idx['診断テーマ']] || ''),
        target_articles_text: targetText,
        target_articles: urls.map(sdsdArticleIdentityForUrl_),
        related_query_count: Number(r[idx['関連クエリ数']] || 0),
        main_queries: String(r[idx['主な検索テーマ']] || ''),
        confidence: String(r[idx['確信度']] || ''),
        expected_site_impact: String(r[idx['サイト全体への期待効果']] || ''),
        doctor_focus: String(r[idx['診断で確認すること']] || ''),
        next_route: String(r[idx['次の担当']] || ''),
        status: String(r[idx['状態']] || '')
      };
    });
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
  const site = sdsdSiteMetaFromArticleMaster_();
  const batchId = sdsdSiteWideBatchId_();

  const counts = {
    cannibal: rows.filter(x => x.improvement_type === 'カニバリ疑い').length,
    new_article: rows.filter(x => x.improvement_type === '新規記事機会').length,
    content_gap: rows.filter(x => x.improvement_type === 'コンテンツギャップ').length
  };

  return {
    format: 'SIMS_DOCTOR_SITE_WIDE_DIAGNOSIS_PACKAGE_V2',
    contract_version: '2.0',
    generated_at: new Date().toISOString(),
    package_scope: 'SITE_WIDE',
    site_diagnosis_batch_id: batchId,
    site: site,
    diagnosis_policy: {
      purpose: 'サイト全体の改善機会を横断的に診断する',
      no_automatic_treatment: true,
      preserve_existing_successful_content: true,
      doctor_decides_final_route: true,
      creator_requires_doctor_approval: true
    },
    site_summary: summary,
    case_count: rows.length,
    case_counts: counts,
    cases: rows.map(x => ({
      case_id: x.case_id,
      improvement_type: x.improvement_type,
      diagnosis_theme: x.diagnosis_theme,
      target_articles: x.target_articles,
      related_query_count: x.related_query_count,
      main_queries: x.main_queries,
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
    '# SIMS Doctor Site-wide Diagnosis Referral V2',
    '',
    'このパッケージは、個別記事の診断ではなくサイト全体の横断診断用です。',
    '',
    `案件数: ${rows.length}`,
    '',
    '## Doctorへの依頼',
    '',
    '- 各案件を単独で見るだけでなく、サイト全体の構造・検索意図・記事群の役割を横断して評価してください。\n- 入力CaseはDoctor判断で統合して構いません。統合した場合は absorbed_source_case_ids を返してください。',
    '- カニバリ疑いは、統合ありきではなく「維持 / 役割分担 / Writer修正 / Merge候補」を比較してください。',
    '- コンテンツギャップは、既存記事改善で足りるか、新記事へ分離すべきかを判断してください。',
    '- 新規記事機会は、既存記事との重複・新たなカニバリ発生リスクを確認してからCreator候補としてください。',
    '- 不明な項目は未評価とし、推測で処置を確定しないでください。',
    '',
    '## 出力してほしいもの',
    '',
    '- 返却JSON format は SIMS_DOCTOR_SITE_WIDE_RESULT_V1 としてください。\n- 返却先はSBMの単体診断登録欄ではありません。利用者はSite Diagnosisの「サイト横断の診断・治療 → 5. Doctor診断結果を登録」から登録します。\n- 可能なら diagnosis_cases[] の正規形式で返してください。難しい場合は priority_queue_for_precision_diagnosis / content_gap_merged_into_existing_theme / sbm_routine_queue_high_confidence_standalone / low_priority_monitor_medium_confidence_single_query / creator_candidates_pending_cannibalization_check の一次トリアージ形式でも受け付けます。\n- サイト全体の総合診断',
    '- 優先して処置すべき案件',
    '- 各案件の最終振り分け（MONITOR / Writer / Merge / Creator / 追加Evidence）',
    '- 共通原因がある場合は、案件横断でまとめて説明',
    '- 大規模改修が不要な場合は、その旨を明記',
    '',
    '## 推奨返却JSONの最小形',
    '',
    '{',
    '  "format": "SIMS_DOCTOR_SITE_WIDE_RESULT_V1",',
    '  "contract_version": "1.0",',
    '  "site_diagnosis_batch_id": "...",',
    '  "site": {"site_id":"...","site_name":"...","site_url":"..."},',
    '  "overall_diagnosis": "...",',
    '  "diagnosis_cases": [',
    '    {',
    '      "diagnosis_case_id": "...",',
    '      "diagnosis_theme": "...",',
    '      "diagnosis_type": "...",',
    '      "absorbed_source_case_ids": [],',
    '      "target_articles": [],',
    '      "doctor_decision": "...",',
    '      "confidence": "...",',
    '      "site_impact": "...",',
    '      "treatment_strategy": "...",',
    '      "route_to": "WRITER|MERGE|CREATOR|MONITOR|NO_ACTION|NEEDS_EVIDENCE",',
    '      "eventual_route": "...",',
    '      "reason": "...",',
    '      "additional_evidence_needed": []',
    '    }',
    '  ]',
    '}',
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
            format: 'SIMS_DOCTOR_SITE_WIDE_CASE_V2',
            contract_version: '2.0',
            case_id: x.case_id,
            improvement_type: x.improvement_type,
            diagnosis_theme: x.diagnosis_theme,
            target_articles: x.target_articles,
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

// ============================================================================
// Source: SiteWideMergeReferral.gs
// ============================================================================
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

// ============================================================================
// Source: SiteWidePrecisionPackage.gs
// ============================================================================
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
    throw new Error('Doctor横断診断結果が登録されていません。先に「サイト横断の診断・治療 → 5. Doctor診断結果を登録」を実行してください。');
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

// ============================================================================
// Source: SiteWideResultContract.gs
// ============================================================================
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

function sdsdIsSiteWidePrecisionResult_(obj) {
  return !!(obj && obj.format === 'SIMS_DOCTOR_SITE_WIDE_PRECISION_RESULT_V1');
}

function sdsdPrecisionArticles_(group, cluster) {
  const src = group || {};
  const list = src.articles || src.target_articles || [];
  if (Array.isArray(list) && list.length) return list;
  return Array.isArray(cluster.target_articles) ? cluster.target_articles : [];
}


function sdsdMergePlanArticles_(mergePlan) {
  mergePlan = mergePlan || {};
  const target = mergePlan.target_article || null;
  const source = mergePlan.source_article || null;
  const list = [];

  [target, source].forEach(a => {
    if (!a) return;
    list.push({
      site_id: String(a.site_id || ''),
      article_id: String(a.article_id || ''),
      article_title: String(a.article_title || ''),
      article_url: String(a.article_url || a.url || '')
    });
  });

  return list;
}

function sdsdMergePlanText_(mergePlan) {
  mergePlan = mergePlan || {};
  const target = mergePlan.target_article || {};
  const source = mergePlan.source_article || {};

  function articleLabel(a) {
    return [
      String(a.article_id || ''),
      String(a.article_title || ''),
      String(a.article_url || a.url || '')
    ].filter(Boolean).join(' / ');
  }

  return {
    survivor: articleLabel(target),
    absorbed: articleLabel(source),
    direction: String(
      mergePlan.redirect_direction ||
      (
        articleLabel(source) && articleLabel(target)
          ? `${articleLabel(source)} → ${articleLabel(target)}`
          : ''
      )
    ),
    content_to_absorb: String(mergePlan.content_to_absorb || '')
  };
}

function sdsdConvertPrecisionResult_(obj) {
  const out = [];
  (obj.clusters || []).forEach((cluster, ci) => {
    const cr = cluster.cluster_result || cluster.result || {};
    const groups = Array.isArray(cr.sub_groups) && cr.sub_groups.length
      ? cr.sub_groups
      : [cr];

    groups.forEach((g, gi) => {
      const route = String(g.route_to || cr.route_to || cluster.route_to || '').toUpperCase();
      if (!route) return;

      const baseId = String(cluster.diagnosis_case_id || cluster.case_id || `PRECISION-${ci+1}`);
      const caseId = groups.length > 1 ? `${baseId}#${gi+1}` : baseId;
      const groupLabel = String(g.group_type || g.group_name || g.name || '');
      const diagnosisTheme = String(
        g.diagnosis_theme || g.theme || cluster.diagnosis_theme || cluster.theme || ''
      );
      const summary = String(
        g.reason || g.rationale || g.diagnosis_summary ||
        cr.diagnosis_summary || cluster.diagnosis_summary || ''
      );

      const mergePlan = g.merge_plan || cr.merge_plan || cluster.merge_plan || null;
      let targetArticles = sdsdPrecisionArticles_(g, cluster);

      // Precision MERGE results often provide article identities only inside merge_plan.
      // Preserve that authoritative direction instead of leaving target_articles blank.
      if (route === 'MERGE' && mergePlan) {
        const mergeArticles = sdsdMergePlanArticles_(mergePlan);
        if (mergeArticles.length) targetArticles = mergeArticles;
      }

      const mergeText = route === 'MERGE' && mergePlan
        ? sdsdMergePlanText_(mergePlan)
        : {survivor:'',absorbed:'',direction:'',content_to_absorb:''};

      out.push({
        diagnosis_case_id: caseId,
        parent_diagnosis_case_id: baseId,
        diagnosis_theme: groupLabel ? `${diagnosisTheme} / ${groupLabel}` : diagnosisTheme,
        diagnosis_type: String(g.group_type || cluster.diagnosis_type || cluster.improvement_type || ''),
        absorbed_source_case_ids: Array.isArray(cluster.absorbed_source_case_ids)
          ? cluster.absorbed_source_case_ids.map(String) : [],
        target_articles: targetArticles,
        doctor_decision: String(g.doctor_decision || g.decision || cr.doctor_decision || route),
        confidence: String(g.confidence || cr.confidence || cluster.confidence || ''),
        site_impact: String(g.site_impact || cr.site_impact || cluster.site_impact || ''),
        treatment_strategy: String(g.treatment_strategy || cr.treatment_strategy || ''),
        route_to: route,
        eventual_route: '',
        reason: summary,
        additional_evidence_needed: Array.isArray(g.additional_evidence_needed)
          ? g.additional_evidence_needed.map(String) : [],
        precision_group_type: groupLabel,
        merge_plan: mergePlan || null,
        merge_survivor: mergeText.survivor,
        merge_absorbed: mergeText.absorbed,
        merge_direction: mergeText.direction,
        merge_content_to_absorb: mergeText.content_to_absorb
      });
    });
  });
  return out;
}

function sdsdValidateSiteWideResult_(obj) {
  const isTriage = !!(obj && obj.format === 'SIMS_DOCTOR_SITE_WIDE_RESULT_V1');
  const isPrecision = sdsdIsSiteWidePrecisionResult_(obj);

  if (!isTriage && !isPrecision) {
    throw new Error(
      '正式な横断診断結果ではありません。format は SIMS_DOCTOR_SITE_WIDE_RESULT_V1 または SIMS_DOCTOR_SITE_WIDE_PRECISION_RESULT_V1 が必要です。'
    );
  }

  if (isPrecision) {
    if (!Array.isArray(obj.clusters) || !obj.clusters.length) {
      throw new Error('Precision Result に clusters がありません。');
    }
    const converted = sdsdConvertPrecisionResult_(obj);
    if (!converted.length) {
      throw new Error('Precision Result から最終処置ルートを取得できません。clusters[].cluster_result.sub_groups[].route_to を確認してください。');
    }
    const allowedPrecision = {WRITER:true, MERGE:true, MONITOR:true, NO_ACTION:true, NEEDS_EVIDENCE:true};
    converted.forEach((c,i) => {
      if (!allowedPrecision[c.route_to]) {
        throw new Error(`Precision Result の処置ルートが不正です: ${c.route_to} (${i+1})`);
      }
    });
    return true;
  }

  if (!Array.isArray(obj.diagnosis_cases) && !sdsdIsDoctorTriageShape_(obj)) {
    throw new Error(
      'diagnosis_cases またはDoctor一次トリアージ形式の診断結果がありません。'
    );
  }

  if (Array.isArray(obj.diagnosis_cases)) {
    const allowed = {
      WRITER:true, MERGE:true, CREATOR:true, MONITOR:true,
      NO_ACTION:true, NEEDS_EVIDENCE:true
    };
    obj.diagnosis_cases.forEach((c,i) => {
      if (!c.diagnosis_case_id) throw new Error(`diagnosis_cases[${i}] に diagnosis_case_id がありません。`);
      if (!Array.isArray(c.absorbed_source_case_ids)) throw new Error(`diagnosis_cases[${i}] に absorbed_source_case_ids がありません。`);
      if (!allowed[String(c.route_to || '')]) throw new Error(`diagnosis_cases[${i}] の route_to が不正です: ${c.route_to}`);
    });
  }
  return true;
}

function sdsdNormalizeDoctorSiteWideResult_(obj) {
  const siteInput = obj.site || {};
  const fallbackSite = sdsdSiteMetaFromArticleMaster_();

  const rawCases = sdsdIsSiteWidePrecisionResult_(obj)
    ? sdsdConvertPrecisionResult_(obj)
    : (Array.isArray(obj.diagnosis_cases)
      ? obj.diagnosis_cases
      : sdsdConvertDoctorTriageResult_(obj));

  const overall =
    obj.overall_diagnosis ||
    obj.overall_assessment ||
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
        ? c.additional_evidence_needed.map(String) : [],
      merge_plan: c.merge_plan || null,
      merge_survivor: String(c.merge_survivor || ''),
      merge_absorbed: String(c.merge_absorbed || ''),
      merge_direction: String(c.merge_direction || ''),
      merge_content_to_absorb: String(c.merge_content_to_absorb || '')
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
    '対象記事','統合先（残す記事）','統合元（吸収する記事）','統合方向',
    '吸収した元案件','確信度','理由','追加Evidence','状態'
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
    c.route_to === 'MERGE' ? String(c.merge_survivor || '') : '',
    c.route_to === 'MERGE' ? String(c.merge_absorbed || '') : '',
    c.route_to === 'MERGE' ? String(c.merge_direction || '') : '',
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
  [70,260,220,220,430,360,360,420,300,90,460,360,130]
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
      'このセル以下に、Doctor回答全文、SIMS_DOCTOR_SITE_WIDE_RESULT_V1 または SIMS_DOCTOR_SITE_WIDE_PRECISION_RESULT_V1 JSON を貼り付けてください。\n' +
      '貼り付け後、メニュー「サイト横断の診断・治療 → 5. Doctor診断結果を登録」をもう一度実行します。'
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
    'このセル以下に、Doctor回答全文、SIMS_DOCTOR_SITE_WIDE_RESULT_V1 または SIMS_DOCTOR_SITE_WIDE_PRECISION_RESULT_V1 JSON を貼り付けてください。';

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
    'このセル以下に、Doctor回答全文、SIMS_DOCTOR_SITE_WIDE_RESULT_V1 または SIMS_DOCTOR_SITE_WIDE_PRECISION_RESULT_V1 JSON を貼り付けてください。\n' +
    '貼り付け後、メニュー「サイト横断の診断・治療 → 5. Doctor診断結果を登録」をもう一度実行します。'
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
        '「サイト横断の診断・治療 → 5. Doctor診断結果を登録」\n' +
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
    try{sdsdRenderHome_();}catch(e){}

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

// ============================================================================
// Source: TreatmentOwnership.gs
// ============================================================================
function sdsdOwnership_(x) {
  const impDecline = x.impressionsPrevious > 0
    ? (x.impressionsPrevious - x.impressionsRecent) / x.impressionsPrevious
    : 0;
  const posWorse = x.positionRecent - x.positionPrevious;
  const ctr = x.impressionsRecent ? x.clicksRecent / x.impressionsRecent : 0;
  const expected = sdsdExpectedCtr_(x.positionRecent);
  const pureCtrGap = expected > 0 && ctr < expected * 0.5 && impDecline < 0.15 && posWorse < 0.75;

  if (pureCtrGap) return {
    ownership:'SBM_OWNED',
    reason:'主病変がCTR/即効性改善'
  };

  if (impDecline >= 0.25 || posWorse >= 1.5) return {
    ownership:'DOCTOR_OWNED',
    reason:'流入低下または順位悪化が強い'
  };

  return {
    ownership:'REVIEW',
    reason:'追加Evidence確認'
  };
}

// ============================================================================
// Source: TreatmentValueScore.gs
// ============================================================================
function sdsdPercentileRanks_(items, getter) {
  const sorted = items.map(getter).filter(v => isFinite(v)).sort((a,b)=>a-b);
  return function(value) {
    if (!sorted.length) return 0;
    let count = 0;
    while (count < sorted.length && sorted[count] <= value) count++;
    return count / sorted.length;
  };
}

function sdsdExpectedCtr_(position) {
  if (!position) return 0;
  if (position <= 3) return 0.10;
  if (position <= 5) return 0.06;
  if (position <= 7) return 0.04;
  if (position <= 10) return 0.025;
  if (position <= 15) return 0.012;
  if (position <= 20) return 0.007;
  return 0.004;
}

function sdsdScoreAll_(items) {
  const demandRank = sdsdPercentileRanks_(items, x => x.impressionsRecent);
  const assetRank = sdsdPercentileRanks_(items, x => x.impressionsFull);

  return items.map(x => {
    const demand = Math.round(demandRank(x.impressionsRecent) * 30 * 10) / 10;
    const asset = Math.round(assetRank(x.impressionsFull) * 15 * 10) / 10;

    let opportunity = 0;
    if (x.positionRecent >= 4 && x.positionRecent <= 10) opportunity += 15;
    else if (x.positionRecent > 10 && x.positionRecent <= 15) opportunity += 12;
    else if (x.positionRecent > 15 && x.positionRecent <= 20) opportunity += 8;

    if (x.impressionsRecent >= 50) {
      const ctr = x.impressionsRecent ? x.clicksRecent / x.impressionsRecent : 0;
      const expected = sdsdExpectedCtr_(x.positionRecent);
      const ratio = expected ? ctr / expected : 1;
      if (ratio <= 0.25) opportunity += 15;
      else if (ratio <= 0.5) opportunity += 11;
      else if (ratio <= 0.75) opportunity += 7;
      else if (ratio < 1) opportunity += 3;
    }

    let urgency = 0;
    if (x.impressionsPrevious > 0) {
      const decline = (x.impressionsPrevious - x.impressionsRecent) / x.impressionsPrevious;
      if (decline >= 0.5) urgency += 12;
      else if (decline >= 0.35) urgency += 10;
      else if (decline >= 0.2) urgency += 7;
      else if (decline >= 0.1) urgency += 4;
    }
    const posWorse = x.positionRecent - x.positionPrevious;
    if (posWorse >= 3) urgency += 8;
    else if (posWorse >= 2) urgency += 6;
    else if (posWorse >= 1) urgency += 4;
    else if (posWorse >= 0.5) urgency += 2;

    if (x.clicksPrevious > 0) {
      const cDecline = (x.clicksPrevious - x.clicksRecent) / x.clicksPrevious;
      if (cDecline >= 0.5) urgency += 5;
      else if (cDecline >= 0.25) urgency += 3;
    }

    urgency = Math.min(25, urgency);
    opportunity = Math.min(30, opportunity);

    return Object.assign({}, x, {
      demand, opportunity, urgency, asset,
      tvs: Math.round((demand + opportunity + urgency + asset) * 10) / 10
    });
  });
}

// ============================================================================
// Source: UrlNormalizer.gs
// ============================================================================
function sdsdNormalizeUrl_(url) {
  const s = String(url || '').trim().split('#')[0].split('?')[0];
  return s.replace(/\/+$/, '');
}

// ============================================================================
// Source: WeeklyTrendAnalyzer.gs
// ============================================================================
function sdsdBuildWeeklyTrendMap_() {
  const rows = sdsdReadObjects_(SDSD_CONFIG.sheets.evidencePageWeekly);
  const grouped = {};

  rows.forEach(r => {
    const raw = r.page || r.key || r.url || '';
    const url = sdsdNormalizeUrl_(raw);
    if (!url) return;

    const weekStart = sdsdDateKey_(r.week_start || r.start_date || '');
    const weekEnd = sdsdDateKey_(r.week_end || r.end_date || '');
    if (!weekStart) return;

    const key = `${url}||${weekStart}`;
    const clicks = Number(r.clicks || 0);
    const impressions = Number(r.impressions || 0);
    const position = Number(r.position || 0);

    const g = grouped[key] || {
      url,
      weekStart,
      weekEnd,
      clicks: 0,
      impressions: 0,
      positionWeightedSum: 0,
      positionFallbackSum: 0,
      positionFallbackCount: 0
    };

    g.clicks += clicks;
    g.impressions += impressions;

    if (impressions > 0 && isFinite(position)) {
      g.positionWeightedSum += position * impressions;
    } else if (isFinite(position) && position > 0) {
      g.positionFallbackSum += position;
      g.positionFallbackCount += 1;
    }

    grouped[key] = g;
  });

  const map = {};

  Object.keys(grouped).forEach(key => {
    const g = grouped[key];

    let position = 0;
    if (g.impressions > 0) {
      position = g.positionWeightedSum / g.impressions;
    } else if (g.positionFallbackCount > 0) {
      position = g.positionFallbackSum / g.positionFallbackCount;
    }

    const arr = map[g.url] || [];
    arr.push({
      weekStart: g.weekStart,
      weekEnd: g.weekEnd,
      clicks: g.clicks,
      impressions: g.impressions,
      position
    });
    map[g.url] = arr;
  });

  Object.keys(map).forEach(url => {
    map[url].sort((a,b) => String(a.weekStart).localeCompare(String(b.weekStart)));
  });

  return map;
}

function sdsdDateKey_(value) {
  if (!value) return '';

  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone() || 'Asia/Tokyo', 'yyyy-MM-dd');
  }

  const s = String(value).trim();

  // Common YYYY-MM-DD text.
  const direct = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (direct) return `${direct[1]}-${direct[2]}-${direct[3]}`;

  // Google Sheets can expose a Date as a locale string.
  const d = new Date(value);
  if (!isNaN(d.getTime())) {
    return Utilities.formatDate(d, Session.getScriptTimeZone() || 'Asia/Tokyo', 'yyyy-MM-dd');
  }

  return s;
}

function sdsdClassifyWeeklyTrend_(series) {
  if (!series || series.length < 4) {
    return {
      trend:'INSUFFICIENT',
      confidence:'LOW',
      declineRatio:0,
      positionChange:0,
      volatility:0
    };
  }

  const valid = series.filter(x => x.impressions > 0);
  if (valid.length < 4) {
    return {
      trend:'INSUFFICIENT',
      confidence:'LOW',
      declineRatio:0,
      positionChange:0,
      volatility:0
    };
  }

  const n = valid.length;
  const block = Math.max(2, Math.floor(n / 3));
  const first = valid.slice(0, block);
  const last = valid.slice(n - block);

  const avg = (arr, key) =>
    arr.reduce((a,x) => a + Number(x[key] || 0), 0) / arr.length;

  const firstImp = avg(first, 'impressions');
  const lastImp = avg(last, 'impressions');
  const firstPos = avg(first, 'position');
  const lastPos = avg(last, 'position');

  const declineRatio = firstImp > 0 ? (firstImp - lastImp) / firstImp : 0;
  const positionChange = lastPos - firstPos;

  const impAvg = avg(valid, 'impressions');
  const variance =
    valid.reduce((a,x) => a + Math.pow((x.impressions || 0) - impAvg, 2), 0) / valid.length;
  const volatility = impAvg > 0 ? Math.sqrt(variance) / impAvg : 0;

  let trend = 'STABLE';

  if (declineRatio >= 0.45 && positionChange >= 1.0) {
    trend = 'SEVERE_DECLINE';
  } else if (declineRatio >= 0.30) {
    trend = 'TRAFFIC_DECLINE';
  } else if (positionChange >= 1.5) {
    trend = 'RANKING_DECLINE';
  } else if (declineRatio <= -0.30) {
    trend = 'GROWTH';
  } else if (volatility >= 0.60) {
    trend = 'VOLATILE';
  }

  let confidence = 'HIGH';
  if (valid.length < 8) confidence = 'MEDIUM';
  if (impAvg < 25) confidence = 'LOW';

  return {
    trend,
    confidence,
    declineRatio,
    positionChange,
    volatility
  };
}

// ============================================================================
// Source: WeeklyTrendValidation.gs
// ============================================================================
function sdsdValidateWeeklyTrends() {
  const weeklyMap = sdsdBuildWeeklyTrendMap_();
  const wanted = ['GROWTH','VOLATILE','TRAFFIC_DECLINE','STABLE'];
  const selected = {};

  Object.keys(weeklyMap).forEach(url => {
    const info = sdsdClassifyWeeklyTrend_(weeklyMap[url]);
    if (wanted.indexOf(info.trend) >= 0 && !selected[info.trend]) {
      selected[info.trend] = {url:url, info:info, series:weeklyMap[url]};
    }
  });

  const ss = SpreadsheetApp.getActive();
  const name = 'Weekly Trend Validation';
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  sh.clear();

  const headers = [
    'Pattern','URL','Week Start','Week End','Clicks','Impressions','Position',
    'First-vs-Last Impression Change','Position Change','Volatility','Classification'
  ];
  sh.getRange(1,1,1,headers.length).setValues([headers]);

  const out = [];
  wanted.forEach(pattern => {
    const x = selected[pattern];
    if (!x) return;
    x.series.forEach((w,i) => out.push([
      pattern, x.url, w.weekStart, w.weekEnd, w.clicks, w.impressions, w.position,
      i===0 ? x.info.declineRatio : '',
      i===0 ? x.info.positionChange : '',
      i===0 ? x.info.volatility : '',
      i===0 ? x.info.trend : ''
    ]));
  });

  if (out.length) sh.getRange(2,1,out.length,headers.length).setValues(out);
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1,headers.length);
  ss.setActiveSheet(sh);

  SpreadsheetApp.getUi().alert(
    '4パターン回帰検証表を作成しました。\n' +
    'Weekly Trend Validation シートを確認してください。'
  );
}
