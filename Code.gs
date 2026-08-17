// ============================================================================
// Source: SiteDiagnosisConfig.gs
// ============================================================================
const SDSD_VERSION = '0.8.3';

const SDSD_CONFIG = Object.freeze({
  sheets: {
    home: 'Diagnosis Home',
    evidencePageSummary: '_SDSD_PAGE_SUMMARY',
    evidencePageWeekly: '_SDSD_PAGE_WEEKLY',
    evidencePageQuery: '_SDSD_PAGE_QUERY_TOP',
    sbmHistory: '_SDSD_SBM_HISTORY',
    summary: 'サイト診断詳細',
    candidates: '診断候補',
    selectedCases: '今回の診断対象',
    articleMaster: '_SDSD_ARTICLE_MASTER',
    opportunities: 'サイト改善プラン',
    opportunityCases: 'サイト横断診断案件',
    siteWideResult: '_SDSD_SITE_WIDE_RESULT',
    treatmentPlan: 'サイト治療計画',
    siteWideResultImport: 'Doctor結果取込',
    creatorValidation: 'Creator候補チェック',
    creatorSerpReferral: 'Creator SERP確認',
    longtailDiscovery: 'ロングテール探索'
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

  const checkMenu = ui.createMenu('確認する')
    .addItem('サイト診断詳細を見る', 'sdsdOpenSiteSummary')
    .addItem('診断候補を見る', 'sdsdOpenCandidates')
    .addItem('個別精密診断対象を見る', 'sdsdOpenSelectedCases')
    .addItem('サイト治療計画を見る', 'sdsdOpenTreatmentPlan')
    .addSeparator()
    .addItem('Creator候補チェックを見る', 'sdsdOpenCreatorValidation')
    .addItem('Creator SERP確認紹介状を見る', 'sdsdOpenCreatorSerpReferral')
    .addItem('Creator SERP Doctor回答を取り込む', 'sdsdShowCreatorSerpResultDialog')
    .addSeparator()
    .addItem('ロングテール探索を見る', 'sdsdOpenLongtailDiscovery')
    .addItem('Doctorロングテール探索結果を取り込む', 'sdsdShowLongtailDiscoveryResultDialog')
    .addItem('選択したロングテール候補をCreator案件にする', 'sdsdPromoteSelectedLongtailCandidateToCreator');

  const manualMenu = ui.createMenu('保守・復旧操作')
    .addItem('Evidence Packageを読み込む', 'sdsdImportEvidencePackageZip')
    .addItem('サイト診断を実行', 'sdsdRunProductDiagnosis')
    .addSeparator()
    .addItem('横断診断：改善機会を診断', 'sdsdRunSiteOpportunityDiagnosis')
    .addItem('横断診断：Doctor案件を作成', 'sdsdBuildSiteOpportunityCases')
    .addItem('横断診断：Doctor Packageを生成', 'sdsdExportSiteWideDoctorPackage')
    .addItem('横断診断：Doctor結果を登録', 'sdsdRegisterSiteWideDoctorResult')
    .addSeparator()
    .addItem('個別記事：精密診断Packageを生成・再開', 'sdsdProceedIndividualPrecisionDiagnosis')
    .addItem('横断診断：追加Evidence Packageを生成', 'sdsdExportPriorityPrecisionClusterPackage')
    .addItem('Creator候補：作成前チェックを実行', 'sdsdRunCreatorCandidateValidation')
    .addItem('Creator候補：選択案件のSERP確認紹介状を作成', 'sdsdCreateCreatorSerpReferralFromSelected')
    .addItem('Creator候補：Doctor SERP結果を登録', 'sdsdShowCreatorSerpResultDialog')
    .addItem('Creator候補：選択案件のロングテール探索紹介状を作成', 'sdsdCreateLongtailDiscoveryReferralFromSelected')
    .addItem('Creator候補：Doctorロングテール探索結果を登録', 'sdsdShowLongtailDiscoveryResultDialog')
    .addItem('Creator候補：選択ロングテールをCreator案件に登録', 'sdsdPromoteSelectedLongtailCandidateToCreator')
    .addItem('横断診断：選択中のMerge紹介状を作成', 'sdsdCreateMergeReferralFromSelectedTreatment')
    .addSeparator()
    .addItem('Article Master（任意）の案内', 'sdsdArticleMasterImportHelp')
    .addItem('SBM改善履歴の取込案内', 'sdsdImportHistoryHelp')
    .addSeparator()
    .addItem('内部シートを表示', 'sdsdShowInternalSheets')
    .addItem('内部シートを隠す', 'sdsdHideInternalSheets');

  const otherMenu = ui.createMenu('設定・管理')
    .addItem('現在の診断状況を確認', 'sdsdShowCurrentSessionStatus')
    .addItem('現在の診断を終了', 'sdsdEndCurrentDiagnosisSession')
    .addSeparator()
    .addItem('ZIP保存先', 'sdsdChooseOutputFolder')
    .addSeparator()
    .addItem('初期設定を実行', 'sdsdInitialize')
    .addSubMenu(manualMenu);

  ui.createMenu('SIMS Doctor Site Diagnosis')
    .addItem('Homeを開く', 'sdsdOpenHome')
    .addItem('1. Site Diagnosisを進める', 'sdsdProceedNextGuided')
    .addSeparator()
    .addSubMenu(checkMenu)
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

  // Doctor結果取込後は、一次検出件数ではなくCreatorゲート再判定後の
  // GREEN + YELLOWを「新規記事機会」として表示する。
  try{
    const stored=sdsdHomeReadStoredSiteWideResult_();
    if(stored && Array.isArray(stored.diagnosis_cases)){
      const validations=sdsdCreatorCandidateValidations_();
      if(validations.length){
        m.newArticle=validations.filter(x=>x.grade==='GREEN'||x.grade==='YELLOW').length;
      }
    }
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
      path:'メニュー最上段 → ▶ 次に進む（Diagnosisに任せる）',
      tone:'BLUE'
    };
  }

  if(stage==='WAITING_INPUT'){
    return {
      title:'Doctorの横断診断結果を登録してください',
      reason:'Doctorへ依頼する段階は完了しています。返ってきた診断結果をDiagnosisへ戻す段階です。',
      path:'メニュー最上段 → ▶ 次に進む（Diagnosisに任せる）',
      tone:'YELLOW'
    };
  }

  const individualPackage=sdsdGetIndividualDoctorPackageState_();
  if(individualPackage.status==='PACKAGE_READY_FOR_DOCTOR'){
    return {
      title:'生成済みPackageをDoctorへ渡してください',
      reason:'個別精密診断Packageは生成済みです。再生成せず、SIMS Doctorへ渡して精密診断を依頼してください。',
      path:'メニュー最上段 → 1. Site Diagnosisを進める',
      tone:'YELLOW'
    };
  }
  if(individualPackage.status==='WAITING_DOCTOR_RESULT'){
    return {
      title:'Doctor結果をSIMS-Blog-Managerへ登録してください',
      reason:'個別精密診断はDoctorへ依頼済みです。返ってきた各記事の診断結果をSBMへ登録してください。',
      path:'メニュー最上段 → 1. Site Diagnosisを進める',
      tone:'YELLOW'
    };
  }

  // v0.7.2: Once Doctor has finalized any Writer / Merge / Creator treatment,
  // hand it to SBM before continuing with still-unresolved Creator candidates.
  // This lets one confirmed Creator case complete end-to-end without forcing the
  // user to finish every remaining YELLOW candidate first.
  if(stored && work.actionableTreatment>0){
    const handoffState=sdsdGetSbmHandoffState_();
    if(!sdsdIsSbmHandoffCurrent_()){
      return {
        title:`SBMへ引き渡す確定案件が${work.actionableTreatment}件あります`,
        reason:'Doctorで処置方針が確定した案件を先にSBMへ渡します。未確定のCreator候補は引き渡し後に続けられます。',
        path:'メニュー最上段 → ▶ 次に進む（Diagnosisに任せる）',
        tone:'BLUE'
      };
    }
  }

  // Creator candidates take priority over a legacy site-wide Precision package
  // only after there is no newly-finalized SBM handoff waiting.
  if(stored){
    const creatorPending=sdsdCreatorValidationCases_().length;
    if(creatorPending>0){
      const creatorSheet=SpreadsheetApp.getActive().getSheetByName(SDSD_CONFIG.sheets.creatorValidation);
      return {
        title:creatorSheet ? '新記事候補のSERP確認を進めてください' : `新記事候補${creatorPending}件の作成前チェックを行います`,
        reason:'既存記事との重大なカニバリだけを避け、独立したロングテール検索意図はCreator候補として前進させます。旧バージョンで生成済みの精密診断Packageは保持したまま後で再開できます。',
        path:creatorSheet ? '確認する → Creator候補チェックを見る → 対象行を選択' : 'メニュー最上段 → ▶ 次に進む（Diagnosisに任せる）',
        tone:'BLUE'
      };
    }
  }

  const precisionPackage=sdsdGetSiteWidePrecisionPackageState_();
  if(precisionPackage.status==='WAITING_DOCTOR_RESULT'){
    return {
      title:'サイト横断の精密診断結果を待っています',
      reason:'追加Evidence精密診断Packageは生成済みです。Doctorから結果が返ったら「Doctor結果取込」へ貼り付けてください。',
      path:'▶ 次に進む → 「Doctor診断結果を貼り付ける」',
      tone:'YELLOW'
    };
  }

  const individualEligible=sdsdEligibleIndividualPrecisionCount_();
  if(individualEligible>0){
    return {
      title:`優先記事${individualEligible}件をDoctorで詳しく診断します`,
      reason:'Diagnosisが優先度・処置履歴・リスクを確認し、個別精密診断へ進める記事を自動判定しました。行を選択する必要はありません。',
      path:'メニュー最上段 → ▶ 次に進む（Diagnosisに任せる）',
      tone:'BLUE'
    };
  }

  if(stored){
    if(work.additionalEvidence>0){
      return {
        title:`詳しい診断が必要な案件が${work.additionalEvidence}件あります`,
        reason:'Doctorが、現在の情報だけでは治療方法を確定できないと判断しました。記事本文などを追加して詳しく診断します。',
        path:'メニュー最上段 → ▶ 次に進む（Diagnosisに任せる）',
        tone:'YELLOW'
      };
    }
    if(work.actionableTreatment>0){
      const handoffState=sdsdGetSbmHandoffState_();
      if(sdsdIsSbmHandoffCurrent_()){
        return {
          title:'SBMへの診断結果引き渡しは完了しています',
          reason:'Writer / Creator / Mergeへの紹介状作成とその後の経過観察はSIMS-Blog-Managerで続けます。',
          path:'SIMS-Blog-Manager → SIMS Doctor → 5．Site Diagnosisの処置を進める',
          tone:'GREEN'
        };
      }
      return {
        title:`SBMへ引き渡す治療案件が${work.actionableTreatment}件あります`,
        reason:'Doctorの診断で治療方針が決まりました。DiagnosisからSBMへ診断結果を渡し、SBMがWriter / Creator / Mergeへ正式に振り分けます。',
        path:'メニュー最上段 → ▶ 次に進む（Diagnosisに任せる）',
        tone:'BLUE'
      };
    }
    return {
      title:'Diagnosisで行う作業は完了しています',
      reason:'現在、追加診断や治療へ送る案件はありません。処置済み記事はSBMで経過観察します。',
      path:'必要なら「設定・管理 → 現在の診断を終了」で次のサイトへ進めます',
      tone:'GREEN'
    };
  }

  if(work.opportunityCases>0){
    return {
      title:'Doctorによるサイト横断診断を進めてください',
      reason:`Diagnosisが複数記事をまとめて確認すべき案件を${work.opportunityCases}件作成済みです。`,
      path:'メニュー最上段 → ▶ 次に進む（Diagnosisに任せる）',
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
      path:'メニュー最上段 → ▶ 次に進む（Diagnosisに任せる）',
      tone:'YELLOW'
    };
  }

  if(m.total>0){
    return {
      title:`個別に詳しく診る記事が${m.total}件あります`,
      reason:'他記事との関係をまとめて調べる横断診断は不要と判定しました。優先記事をDoctorで詳しく診断します。',
      path:'メニュー最上段 → ▶ 次に進む（Diagnosisに任せる）',
      tone:'BLUE'
    };
  }

  return {
    title:'サイト診断を実行してください',
    reason:'Evidence Packageは読み込み済みですが、サイト全体の診断結果がまだありません。',
    path:'メニュー最上段 → ▶ 次に進む（Diagnosisに任せる）',
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
  sh.getRange('A3:H3').merge().setValue('迷ったときは、メニュー最上段の「▶ 次に進む（Diagnosisに任せる）」を実行してください。');

  sh.getRange('A4:B9').setValues([
    ['対象サイト',siteLabel],
    ['Evidence',session.evidenceFileName||'未読込'],
    ['総合状態',overall.label],
    ['診断候補',metrics.total+'件'],
    ['要確認作業',work.pendingTotal+'件'],
    ['バージョン','v'+SDSD_VERSION]
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
  sh.getRange('A3:H3').setBackground('#E6F4EA').setFontColor('#137333').setFontWeight('bold').setFontSize(11);

  sh.getRange('A4:A9').setBackground('#F1F3F4').setFontWeight('bold').setFontColor('#5F6368');
  sh.getRange('B4:B9').setBackground('#FFFFFF');
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
  sh.setRowHeights(4,6,30);sh.setRowHeights(11,16,28);
  sh.getRange('A1:H26').setBorder(false,false,false,false,false,false);
  sh.getRange('A4:B9').setBorder(true,true,true,true,true,true,'#DADCE0',SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange('D4:H9').setBorder(true,true,true,true,true,true,'#DADCE0',SpreadsheetApp.BorderStyle.SOLID);
}


function sdsdInitialize() {
  sdsdProductEnsureSheets_();
  sdsdHideInternalSheets_();
  SpreadsheetApp.getUi().alert(
    `SIMS Doctor Site Diagnosis ${SDSD_VERSION}\n初期設定が完了しました。\n\n次にメニュー最上段の「▶ 次に進む（Diagnosisに任せる）」を実行してください。`
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
    ['Selected Treatment Cases', SDSD_CONFIG.sheets.selectedCases],
    ['サイト診断サマリー', SDSD_CONFIG.sheets.summary]
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

function sdsdTitleFromUrlFallback_(url){
  try{
    const u=String(url||'').replace(/[?#].*$/,'').replace(/\/+$/,'');
    const slug=decodeURIComponent(u.substring(u.lastIndexOf('/')+1)||'');
    if(!slug)return '（タイトル未取得）';
    return slug.replace(/[-_]+/g,' ').replace(/\s+/g,' ').trim();
  }catch(e){
    return '（タイトル未取得）';
  }
}

function sdsdDisplayTitleSmart_(url, articleMap){
  const fromMaster=sdsdDisplayTitle_(url,articleMap);
  if(fromMaster && fromMaster!=='（タイトル未取得）')return fromMaster;

  const cache=CacheService.getDocumentCache();
  const key='SDSD_TITLE_'+Utilities.base64EncodeWebSafe(String(url||'')).slice(0,80);
  try{
    const cached=cache.get(key);
    if(cached)return cached;
  }catch(e){}

  try{
    const fetched=sdsdFetchArticleEvidence_(url);
    const title=String((fetched&&fetched.title)||'').trim();
    if(title){
      try{cache.put(key,title,21600);}catch(e){}
      return title;
    }
  }catch(e){}

  return sdsdTitleFromUrlFallback_(url);
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

function sdsdDetailSymptom_(r){
  const t=String(r.weeklyTrend||'');
  if(t==='SEVERE_DECLINE')return '直近で大きく悪化';
  if(t==='TRAFFIC_DECLINE')return '検索流入が低下';
  if(t==='RANKING_DECLINE')return '検索順位が低下';
  if(t==='VOLATILE')return '検索推移が不安定';
  if(t==='GROWTH')return '回復・成長傾向';
  if(String(r.ownership||'')==='SBM_OWNED')return '順位に対してCTR改善余地';
  if(String(r.evidenceConfidence||'')==='LOW')return 'データ不足で判定保留';
  return '優先度スコアから要確認';
}

function sdsdDetailCause_(r){
  const parts=[];
  const t=String(r.weeklyTrend||'');
  const own=String(r.ownership||'');
  const ext=String(r.externalFactor||'');
  if(own==='SBM_OWNED')parts.push('本文全体より、タイトル・スニペット側のCTR不足をまず確認');
  if(t==='SEVERE_DECLINE'||t==='TRAFFIC_DECLINE')parts.push('表示回数・検索需要・主要クエリ順位のどこで落ちたかを切り分ける');
  if(t==='RANKING_DECLINE')parts.push('主要クエリの順位低下と検索意図・競合変化を確認');
  if(t==='VOLATILE')parts.push('短期変動の可能性があるため、週次推移を確認して過剰修正を避ける');
  if(ext)parts.push('OS・サービス変更など外部要因の影響可能性を確認');
  if(String(r.evidenceConfidence||'')==='LOW')parts.push('Evidence量が少ないため、追加データを確認してから判断');
  if(String(r.treatmentRisk||'')==='HIGH')parts.push('変更リスクが高いため、全面改修前にDoctorで原因確定が必要');
  return parts.length?parts.join('。'):'検索パフォーマンスと記事内容を精密診断して原因を特定する。';
}

function sdsdDetailNextAction_(r){
  const p=String(r.priority||'');
  const own=String(r.ownership||'');
  const t=String(r.weeklyTrend||'');
  if(p==='WAIT')return '現在は修正せず、SBMで経過観察を続ける。';
  if(p==='PROTECTED'||t==='GROWTH')return '回復中のため大きく触らない。必要なら軽微な鮮度確認だけ行う。';
  if(own==='SBM_OWNED'||p==='SBM')return 'SBMの日常改善へ回し、タイトル・meta descriptionなど入口改善を優先する。';
  if(p==='REVIEW'||p==='DOCTOR_REVIEW'||p==='B_CANDIDATE'||p==='CANDIDATE')
    return '追加Evidenceを確認し、Doctor精密診断へ進めるかをDiagnosisで再判定する。';
  if(p==='A1_CANDIDATE'||p==='A2_CANDIDATE')
    return 'Diagnosisが自動で精密診断対象にします。メニュー最上段の「▶ 次に進む（Diagnosisに任せる）」を実行してください。';
  return '診断候補画面で詳細を確認し、Homeの案内に従って次の処置へ進む。';
}

function sdsdDetailPriorityRows_(rows){
  return rows.slice().sort((a,b)=>{
    const pa=sdsdPriorityRank_(a.priority),pb=sdsdPriorityRank_(b.priority);
    if(pa!==pb)return pa-pb;
    return Number(b.tvs||0)-Number(a.tvs||0);
  });
}

function sdsdWriteSiteSummary_(rows, result) {
  const ss=SpreadsheetApp.getActive();
  let sh=ss.getSheetByName(SDSD_CONFIG.sheets.summary);
  if(!sh)sh=ss.insertSheet(SDSD_CONFIG.sheets.summary);
  sh.clear();

  const total=result.total||rows.length;
  const articleMap=sdsdArticleTitleMap_();
  const priorityRows=sdsdDetailPriorityRows_(rows);
  const top=priorityRows.filter(r =>
    ['A1_CANDIDATE','A2_CANDIDATE','DOCTOR_REVIEW','REVIEW','B_CANDIDATE','CANDIDATE']
      .indexOf(String(r.priority||''))>=0
  ).slice(0,5);

  const a1=rows.filter(r=>r.priority==='A1_CANDIDATE').length;
  const a2=rows.filter(r=>r.priority==='A2_CANDIDATE').length;
  const severe=rows.filter(r=>r.weeklyTrend==='SEVERE_DECLINE').length;
  const traffic=rows.filter(r=>r.weeklyTrend==='TRAFFIC_DECLINE').length;
  const ranking=rows.filter(r=>r.weeklyTrend==='RANKING_DECLINE').length;
  const volatile=rows.filter(r=>r.weeklyTrend==='VOLATILE').length;
  const growth=rows.filter(r=>r.weeklyTrend==='GROWTH').length;
  const external=rows.filter(r=>String(r.externalFactor||'')!=='').length;

  sh.getRange('A1:F1').merge().setValue('サイト診断詳細');
  sh.getRange('A2:F2').merge().setValue(
    'Homeの概要をさらに掘り下げ、どの記事に何が起きていて、次に何を確認・処置するかを示します。'
  );
  sh.getRange('A3:F3').merge().setValue(
    '操作はメニュー最上段の「▶ 次に進む（Diagnosisに任せる）」を基本にしてください。Diagnosisが対象記事と次の処理を自動判定します。'
  ).setBackground('#E6F4EA').setFontColor('#137333').setFontWeight('bold');

  sh.getRange('A4:B10').setValues([
    ['診断対象',`${total}記事`],
    ['最優先',`${a1}記事`],
    ['優先',`${a2}記事`],
    ['直近で大きく悪化',`${severe}記事`],
    ['検索流入低下',`${traffic}記事`],
    ['検索順位低下',`${ranking}記事`],
    ['外部要因の可能性',`${external}記事`]
  ]);

  sh.getRange('D4:F4').merge().setValue('診断所見');
  let finding='サイト全体を一律に修正する状態ではありません。';
  if(top.length){
    const named=top.slice(0,2).map(r=>{
      const title=sdsdDisplayTitleSmart_(r.url,articleMap);
      return `${title}（${r.url}）`;
    });
    finding+=` まず ${named.join('、')} を優先して原因を確認します。`;
    finding+=' 下の「優先して確認する記事」に、症状・原因候補・具体的な次の処置を示しています。';
  }else{
    finding+=' 現時点ではDoctor精密診断を急ぐ記事はありません。';
  }
  if(growth>0)finding+=` 回復・成長中の記事${growth}件は保護し、不要な修正を避けます。`;
  sh.getRange('D5:F10').merge().setValue(finding);

  sh.getRange('A12:F12').merge().setValue('優先して確認する記事');
  const headers=['優先','記事','現在の症状','確認すべき原因','次に行うこと','根拠'];
  sh.getRange(13,1,1,headers.length).setValues([headers]);

  if(top.length){
    const vals=top.map((r,i)=>{
      const title=sdsdDisplayTitleSmart_(r.url,articleMap);
      return [
        i+1,
        `${title}\n${r.url}`,
        sdsdDetailSymptom_(r),
        sdsdDetailCause_(r),
        sdsdDetailNextAction_(r),
        `優先度:${sdsdPriorityJa_(r.priority)} / TVS:${Number(r.tvs||0).toFixed(1)} / Evidence:${r.evidenceConfidence||'-'} / Risk:${r.treatmentRisk||'-'}`
      ];
    });
    sh.getRange(14,1,vals.length,headers.length).setValues(vals);
  }else{
    sh.getRange('A14:F15').merge().setValue('現在、優先して精密診断する記事はありません。Homeの「次に行うこと」を確認してください。');
  }

  const base=top.length?14+top.length+2:17;
  sh.getRange(base,1,1,6).merge().setValue('症状別の詳しい見方');
  const symptomRows=[
    ['大きな悪化',`${severe}記事`,'表示回数・クリック・順位のどこが落ちたかを切り分け、急落時期と外部要因を確認する。'],
    ['検索流入低下',`${traffic}記事`,'需要減少か順位低下かを分け、主要クエリの推移を確認する。'],
    ['検索順位低下',`${ranking}記事`,'検索意図・競合・鮮度・記事役割のズレを精密診断で確認する。'],
    ['推移が不安定',`${volatile}記事`,'短期変動なら修正を急がず、週次データで再確認する。'],
    ['回復・成長',`${growth}記事`,'原則保護。大幅な構成変更は行わない。'],
    ['外部要因候補',`${external}記事`,'OS・サービス仕様変更やアルゴリズム要因の可能性をDoctorで検証する。']
  ];
  sh.getRange(base+1,1,1,3).setValues([['症状','件数','具体的な確認・解決方針']]);
  sh.getRange(base+2,1,symptomRows.length,3).setValues(symptomRows);

  const noteRow=base+9;
  sh.getRange(noteRow,1,1,6).merge().setValue('この詳細画面の役割');
  sh.getRange(noteRow+1,1,2,6).merge().setValue(
    'ここでは「どの記事を、なぜ、どう確認するか」を示します。最終的なリライト・Merge・新記事作成は、Doctorの精密診断で原因を確定してから決めます。Homeは全体状況と次の一手、この画面は原因分析と具体的な処置候補を確認する場所です。'
  );

  // Product styling
  sh.getRange('A1:F1').setBackground('#185ABC').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(18);
  sh.getRange('A2:F2').setBackground('#D2E3FC').setFontColor('#174EA6').setFontSize(11);
  sh.getRange('A4:A10').setBackground('#F1F3F4').setFontWeight('bold').setFontColor('#5F6368');
  sh.getRange('B4:B10').setBackground('#FFFFFF').setFontWeight('bold');
  sh.getRange('D4:F4').setBackground('#D2E3FC').setFontColor('#174EA6').setFontWeight('bold');
  sh.getRange('D5:F10').setBackground('#F8F9FA');
  sh.getRange('A12:F12').setBackground('#D2E3FC').setFontColor('#174EA6').setFontWeight('bold');
  sh.getRange(13,1,1,6).setBackground('#E8F0FE').setFontColor('#174EA6').setFontWeight('bold');

  if(top.length){
    for(let i=0;i<top.length;i++){
      const row=14+i;
      const tone=i<2?'#FEF7E0':'#FFFFFF';
      sh.getRange(row,1,1,6).setBackground(tone);
      sh.getRange(row,1).setFontWeight('bold').setHorizontalAlignment('center');
      sh.getRange(row,2).setFontWeight('bold');
      if(i<2)sh.getRange(row,1,1,2).setFontColor('#9A6700');
    }
  }
  sh.getRange(base,1,1,6).setBackground('#D2E3FC').setFontColor('#174EA6').setFontWeight('bold');
  sh.getRange(base+1,1,1,3).setBackground('#E8F0FE').setFontColor('#174EA6').setFontWeight('bold');
  sh.getRange(noteRow,1,1,6).setBackground('#D2E3FC').setFontColor('#174EA6').setFontWeight('bold');
  sh.getRange(noteRow+1,1,2,6).setBackground('#F8F9FA');

  sh.setFrozenRows(2);
  sh.setHiddenGridlines(true);
  [70,360,180,430,430,300].forEach((w,i)=>sh.setColumnWidth(i+1,w));
  sh.getRange(1,1,Math.max(sh.getLastRow(),1),6).setWrap(true).setVerticalAlignment('top');
  sh.getRange(1,1,Math.max(sh.getLastRow(),1),6).setFontFamily('Arial');
  sh.getRange(1,1,Math.max(sh.getLastRow(),1),6).setBorder(false,false,false,false,false,false);
  sh.getRange('A4:B10').setBorder(true,true,true,true,true,true,'#DADCE0',SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange('D4:F10').setBorder(true,true,true,true,true,true,'#DADCE0',SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(13,1,Math.max(top.length+1,2),6).setBorder(true,true,true,true,true,true,'#DADCE0',SpreadsheetApp.BorderStyle.SOLID);

  try{ss.setActiveSheet(sh);}catch(e){}
}

function sdsdUpdateSummaryAfterBatch_(batch, guard) {
  const sh=SpreadsheetApp.getActive().getSheetByName(SDSD_CONFIG.sheets.summary);
  if(!sh)return;
  try{
    sh.getRange('A3:F3').merge().setValue(
      `Treatment Batch：Doctorへ送る ${batch.selectedCount}記事`+
      (guard&&guard.blocked?` / 最終確認で保留 ${guard.blocked}記事`:'')
    ).setBackground('#E6F4EA').setFontColor('#137333').setFontWeight('bold');
  }catch(e){}
}

function sdsdSetIndividualDoctorPackageState_(state) {
  const props = PropertiesService.getDocumentProperties();
  const obj = state || {};
  props.setProperty('SDSD_INDIVIDUAL_DOCTOR_PACKAGE_STATUS', String(obj.status || ''));
  props.setProperty('SDSD_INDIVIDUAL_DOCTOR_PACKAGE_BATCH_ID', String(obj.batchId || ''));
  props.setProperty('SDSD_INDIVIDUAL_DOCTOR_PACKAGE_FILE_URL', String(obj.fileUrl || ''));
  props.setProperty('SDSD_INDIVIDUAL_DOCTOR_PACKAGE_FILE_NAME', String(obj.fileName || ''));
  props.setProperty('SDSD_INDIVIDUAL_DOCTOR_PACKAGE_CASE_COUNT', String(obj.caseCount || 0));
  props.setProperty('SDSD_INDIVIDUAL_DOCTOR_PACKAGE_AT', new Date().toISOString());
}

function sdsdGetIndividualDoctorPackageState_() {
  const props = PropertiesService.getDocumentProperties();
  return {
    status: String(props.getProperty('SDSD_INDIVIDUAL_DOCTOR_PACKAGE_STATUS') || ''),
    batchId: String(props.getProperty('SDSD_INDIVIDUAL_DOCTOR_PACKAGE_BATCH_ID') || ''),
    fileUrl: String(props.getProperty('SDSD_INDIVIDUAL_DOCTOR_PACKAGE_FILE_URL') || ''),
    fileName: String(props.getProperty('SDSD_INDIVIDUAL_DOCTOR_PACKAGE_FILE_NAME') || ''),
    caseCount: Number(props.getProperty('SDSD_INDIVIDUAL_DOCTOR_PACKAGE_CASE_COUNT') || 0),
    at: String(props.getProperty('SDSD_INDIVIDUAL_DOCTOR_PACKAGE_AT') || '')
  };
}

function sdsdClearIndividualDoctorPackageState_() {
  const props = PropertiesService.getDocumentProperties();
  [
    'SDSD_INDIVIDUAL_DOCTOR_PACKAGE_STATUS',
    'SDSD_INDIVIDUAL_DOCTOR_PACKAGE_BATCH_ID',
    'SDSD_INDIVIDUAL_DOCTOR_PACKAGE_FILE_URL',
    'SDSD_INDIVIDUAL_DOCTOR_PACKAGE_FILE_NAME',
    'SDSD_INDIVIDUAL_DOCTOR_PACKAGE_CASE_COUNT',
    'SDSD_INDIVIDUAL_DOCTOR_PACKAGE_AT'
  ].forEach(k => props.deleteProperty(k));
}

function sdsdGetCompletedIndividualUrls_() {
  const raw = String(
    PropertiesService.getDocumentProperties().getProperty('SDSD_INDIVIDUAL_DOCTOR_COMPLETED_URLS') || ''
  );
  if (!raw) return {};
  try {
    const arr = JSON.parse(raw);
    const map = {};
    (Array.isArray(arr) ? arr : []).forEach(url => {
      const normalized = sdsdNormalizeUrl_(url);
      if (normalized) map[normalized] = true;
    });
    return map;
  } catch (e) {
    return {};
  }
}

function sdsdAddCompletedIndividualUrls_(urls) {
  const props = PropertiesService.getDocumentProperties();
  const map = sdsdGetCompletedIndividualUrls_();
  (urls || []).forEach(url => {
    const normalized = sdsdNormalizeUrl_(url);
    if (normalized) map[normalized] = true;
  });
  props.setProperty('SDSD_INDIVIDUAL_DOCTOR_COMPLETED_URLS', JSON.stringify(Object.keys(map)));
}

function sdsdCurrentSelectedCaseUrls_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SDSD_CONFIG.sheets.selectedCases);
  if (!sh || sh.getLastRow() < 2) return [];
  const values = sh.getDataRange().getValues();
  const headers = values[0].map(String);
  const idx = {};
  headers.forEach((h,i) => idx[h] = i);
  const urlIdx = idx['URL'];
  if (urlIdx == null) return [];
  return values.slice(1)
    .map(r => String(r[urlIdx] || ''))
    .filter(Boolean);
}

function sdsdMarkCurrentIndividualBatchHandedOff_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SDSD_CONFIG.sheets.selectedCases);
  const urls = sdsdCurrentSelectedCaseUrls_();
  if (urls.length) sdsdAddCompletedIndividualUrls_(urls);

  if (sh && sh.getLastRow() >= 2) {
    const values = sh.getDataRange().getValues();
    const headers = values[0].map(String);
    const idx = {};
    headers.forEach((h,i) => idx[h] = i);
    const refIdx = idx['Referral Status'];
    const pkgIdx = idx['Case Package Status'];

    for (let i=1; i<values.length; i++) {
      if (refIdx != null) values[i][refIdx] = 'HANDED_OFF_TO_SBM';
      if (pkgIdx != null) values[i][pkgIdx] = 'HANDED_OFF_TO_SBM';
    }
    sh.getRange(2, 1, values.length - 1, headers.length).setValues(values.slice(1));
  }

  sdsdClearIndividualDoctorPackageState_();
  try { sdsdRefreshSelectedCasesView_(); } catch (e) {}
  try { sdsdRenderHome_(); } catch (e) {}

  return urls.length;
}

function sdsdPendingSelectedCaseCount_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SDSD_CONFIG.sheets.selectedCases);
  if (!sh || sh.getLastRow() < 2) return 0;
  const values = sh.getDataRange().getDisplayValues();
  const headers = values[0].map(String);
  const refIdx = headers.indexOf('Referral Status');
  if (refIdx < 0) return values.length - 1;

  return values.slice(1).filter(r => {
    if (!r.some(v => String(v || '').trim() !== '')) return false;
    return String(r[refIdx] || '') !== 'HANDED_OFF_TO_SBM';
  }).length;
}


function sdsdShowIndividualPackageReady_() {
  const state=sdsdGetIndividualDoctorPackageState_();
  const esc=sdsdWorkflowEsc_;
  const fileName=String(state.fileName||'個別精密診断Package');
  const fileUrl=String(state.fileUrl||'');
  const count=Number(state.caseCount||0);
  const link=fileUrl
    ? `<a href="${esc(fileUrl)}" target="_blank" style="color:#1a73e8;font-weight:bold">Google DriveでPackageを開く</a>`
    : '<span style="color:#5f6368">保存先URLを取得できませんでした。</span>';

  const html=`<!doctype html><html><head><base target="_top"><style>
    body{font-family:Arial,"Noto Sans JP",sans-serif;margin:0;background:#f8fafd;color:#202124}.wrap{padding:20px}
    .hero{background:#185abc;color:#fff;border-radius:12px;padding:17px}.hero h2{margin:0 0 6px;font-size:20px}.hero p{margin:0;font-size:13px}
    .card{background:#fff;border:1px solid #dadce0;border-radius:10px;padding:14px;margin-top:13px}.title{font-size:17px;font-weight:bold}
    .text{font-size:13px;line-height:1.75;margin-top:8px}.file{background:#f1f3f4;border-radius:8px;padding:11px;margin-top:10px;font-size:13px}
    .actions{text-align:right;margin-top:16px}button{border:1px solid #dadce0;background:#fff;border-radius:7px;padding:9px 13px;margin-left:6px;cursor:pointer}
    .primary{background:#1a73e8;color:#fff;border-color:#1a73e8;font-weight:bold}
  </style></head><body><div class="wrap">
    <div class="hero"><h2>Doctorへ個別精密診断を依頼してください</h2><p>Packageは生成済みです。再生成する必要はありません。</p></div>
    <div class="card"><div class="title">生成済みPackage</div><div class="text">対象：${count}記事</div>
      <div class="file"><b>ファイル：</b>${esc(fileName)}<br><br>${link}</div></div>
    <div class="card"><div class="title">次にすること</div><div class="text">
      1. このPackageをSIMS Doctorへ渡して、個別精密診断を依頼してください。<br>
      2. Doctorへ渡したら「Doctorへ依頼しました」を押してください。<br>
      3. 回答が返ったら、各記事の診断結果をSBMへ登録します。
    </div></div>
    <div class="actions"><button onclick="google.script.host.close()">中断して閉じる</button>
      <button class="primary" onclick="markRequested()">Doctorへ依頼しました</button></div>
    <div id="status" style="font-size:12px;color:#5f6368;text-align:right;margin-top:8px"></div>
  </div><script>
    function markRequested(){
      document.getElementById('status').textContent='状態を保存しています...';
      google.script.run.withSuccessHandler(()=>google.script.host.close())
        .withFailureHandler(e=>document.getElementById('status').textContent=(e&&e.message)?e.message:String(e))
        .sdsdMarkIndividualDoctorRequested();
    }
  </script></body></html>`;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(690).setHeight(560),'個別精密診断Package');
}

function sdsdMarkIndividualDoctorRequested() {
  const state=sdsdGetIndividualDoctorPackageState_();
  if(!state || state.status!=='PACKAGE_READY_FOR_DOCTOR') throw new Error('Doctorへ渡す個別精密診断Packageが確認できません。');
  sdsdSetIndividualDoctorPackageState_({
    status:'WAITING_DOCTOR_RESULT',
    batchId:state.batchId,
    fileUrl:state.fileUrl,
    fileName:state.fileName,
    caseCount:state.caseCount
  });
  sdsdSetWorkflowState_({status:'WAIT_INDIVIDUAL'});
}

function sdsdShowIndividualDoctorResultWaiting_() {
  const ui = SpreadsheetApp.getUi();
  const state = sdsdGetIndividualDoctorPackageState_();
  const detail = state.fileName
    ? `\n\n生成済みPackage: ${state.fileName}`
    : '';

  const answer = ui.alert(
    'Doctor結果をSBMへ登録',
    '個別精密診断Packageは生成済みです。\n\n' +
    'まず、このPackageをSIMS Doctorへ渡して精密診断を依頼してください。\n' +
    'Doctorから回答が返ったら、各記事の診断結果JSONをSIMS-Blog-Managerへ登録してください。' +
    detail + '\n\n' +
    'Doctor結果のSBMへの登録は完了しましたか？',
    ui.ButtonSet.YES_NO
  );

  if (answer !== ui.Button.YES) {
    ui.alert(
      'Doctor結果待ちです。\n\n' +
      'SIMS Doctorの回答から各記事の診断結果JSONをSIMS-Blog-Managerへ登録してください。\n' +
      '登録が終わったら、もう一度「1. Site Diagnosisを進める」を実行してください。'
    );
    return false;
  }

  const count = sdsdMarkCurrentIndividualBatchHandedOff_();
  ui.alert(
    '個別精密診断の引き渡しを完了として記録しました。\n\n' +
    `SBMへ引き渡し済み: ${count}記事\n\n` +
    '同じ記事は、この診断セッションでは個別精密診断の対象として再選定しません。\n' +
    '続きは「1. Site Diagnosisを進める」で進められます。'
  );
  return true;
}

function sdsdUpdateSummaryAfterPackage_(caseCount, fileUrl) {
  const sh=SpreadsheetApp.getActive().getSheetByName(SDSD_CONFIG.sheets.summary);
  if(!sh)return;
  try{
    sh.getRange('A3:F3').merge().setValue(`Doctor Case Package：${caseCount}件生成済み`)
      .setBackground('#E6F4EA').setFontColor('#137333').setFontWeight('bold');
    if(fileUrl)sh.getRange('A3').setNote(fileUrl);
  }catch(e){}
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
    '内部処理用シートを非表示にしました。\n通常利用ではHome・「サイト診断詳細」・「診断候補」を確認すれば大丈夫です。'
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

    const rec=sdsdWorkflowRecommendation_();
    sdsdSetWorkflowState_({
      route:null,
      status:'ANALYSIS_COMPLETE',
      recommendation:rec.route,
      recommendationReason:rec.reason
    });

    try { sdsdRenderHome_(); } catch(e) {}
    sdsdShowWorkflowRecommendationDialog_();
    return result;
  } catch (e) {
    SpreadsheetApp.getUi().alert(
      `サイト診断を完了できませんでした。\n\n${e.message || e}`
    );
    throw e;
  }
}



function sdsdIndividualEnrichmentProgress_(){
  const sh=SpreadsheetApp.getActive().getSheetByName(SDSD_CONFIG.sheets.selectedCases);
  if(!sh || sh.getLastRow()<2)return {active:false,total:0,ready:0,review:0,pending:0};

  const values=sh.getDataRange().getDisplayValues();
  const headers=values[0].map(String);
  const urlIdx=headers.indexOf('URL');
  const refIdx=headers.indexOf('Referral Status');
  const pkgIdx=headers.indexOf('Case Package Status');
  if(urlIdx<0 || refIdx<0)return {active:false,total:0,ready:0,review:0,pending:0};

  let total=0,ready=0,review=0;
  values.slice(1).forEach(r=>{
    if(!String(r[urlIdx]||'').trim())return;
    total++;
    const rs=String(r[refIdx]||'');
    const ps=pkgIdx>=0?String(r[pkgIdx]||''):'';
    if(rs==='READY_FOR_INDIVIDUAL_DOCTOR' && ps==='READY')ready++;
    else if(rs==='NEEDS_CASE_ENRICHMENT_REVIEW')review++;
  });
  return {
    active:total>0 && (ready>0 || review>0),
    total:total,
    ready:ready,
    review:review,
    pending:Math.max(total-ready-review,0)
  };
}

function sdsdHasResumableIndividualEnrichment_(){
  const p=sdsdIndividualEnrichmentProgress_();
  return p.active && p.pending>0;
}

function sdsdProceedIndividualPrecisionDiagnosis(){
  const ui=SpreadsheetApp.getUi();
  try{
    const before=sdsdIndividualEnrichmentProgress_();
    let batch=null;

    if(sdsdHasResumableIndividualEnrichment_()){
      batch={
        selectedCount:before.total,
        resumed:true
      };
      sdsdProgress_(1,4,`個別精密診断を再開しています（${before.ready}/${before.total}件準備済み）`);
    }else{
      sdsdProgress_(1,4,'Diagnosisが優先記事を自動選定しています');
      batch=sdsdBuildTreatmentBatch({silent:true,noActivate:true});
      const guard=sdsdRunFinalGuard({silent:true});
      sdsdUpdateSummaryAfterBatch_(batch,guard);
      sdsdHideInternalSheets_();

      if(!batch.selectedCount){
        ui.alert(
          '個別精密診断へ送る記事はありません。\n\n' +
          'Diagnosisが優先度・処置履歴・リスクを確認しましたが、現在Doctorへ送る適格記事はありませんでした。\n' +
          'Homeの「次に行うこと」を確認してください。'
        );
        try{sdsdRenderHome_();}catch(e){}
        return;
      }
    }

    sdsdProgress_(2,4,'記事識別情報を自動解決しています');
    sdsdArticleMasterCoverageForSelected_();

    sdsdProgress_(3,4,'記事本文と主要クエリを準備しています');
    const enrichment=sdsdEnrichSelectedCases({silent:true,maxPerRun:3});
    try{sdsdRefreshSelectedCasesView_();}catch(e){}

    if(enrichment.review>0){
      ui.alert(
        `精密診断Packageの準備中に、追加確認が必要な記事が見つかりました。\n\n` +
        `準備完了: ${enrichment.ready}/${enrichment.total}件\n` +
        `要確認: ${enrichment.review}件\n\n` +
        '「確認する → 個別精密診断対象を見る」で内容を確認してください。'
      );
      try{sdsdRenderHome_();}catch(e){}
      return;
    }

    if(!enrichment.complete){
      ui.alert(
        `個別精密診断を分割して準備しています。\n\n` +
        `全体: ${enrichment.total}件\n` +
        `準備完了: ${enrichment.ready}件\n` +
        `残り: ${enrichment.pending}件\n\n` +
        'Apps Scriptの安全な実行時間内で一度区切りました。\n' +
        '「OK」を押したあと、もう一度「1. Site Diagnosisを進める」を開いてください。\n' +
        `次回は準備済み${enrichment.ready}件を作り直さず、残り${enrichment.pending}件から再開します。`
      );
      sdsdSetWorkflowState_({status:'INDIVIDUAL_ENRICHMENT_PARTIAL'});
      try{sdsdRenderHome_();}catch(e){}
      return;
    }

    sdsdProgress_(4,4,'Doctor精密診断Packageを生成しています');
    const exported=sdsdExportDoctorCasePackageZip({silent:true});
    sdsdSetIndividualDoctorPackageState_({
      status:'PACKAGE_READY_FOR_DOCTOR',
      batchId:sdsdGetActiveBatchId_(),
      fileUrl:exported.fileUrl,
      fileName:exported.fileName,
      caseCount:exported.caseCount
    });
    sdsdUpdateSummaryAfterPackage_(exported.caseCount,exported.fileUrl);
    sdsdSetWorkflowState_({status:'INDIVIDUAL_PACKAGE_READY'});
    try{sdsdRefreshSelectedCasesView_();}catch(e){}
    try{sdsdRenderHome_();}catch(e){}

    ui.alert(
      `Doctor精密診断Packageを生成しました。\n\n` +
      `対象: ${exported.caseCount}記事\n` +
      `ファイル: ${exported.fileName}\n\n` +
      '次に、このPackageをSIMS Doctorへ渡して精密診断を依頼してください。\n' +
      'Doctorから回答が返ったら、各記事の診断結果をSBMへ登録します。\n\n' +
      exported.fileUrl
    );
  }catch(e){
    ui.alert(`個別精密診断を進められませんでした。\n\n${e.message||e}`);
    throw e;
  }
}


function sdsdEligibleIndividualPrecisionCount_(){
  const completed=sdsdGetCompletedIndividualUrls_();
  const rows=sdsdCandidateRowsFromSheet_();
  return rows.filter(r =>
    !completed[sdsdNormalizeUrl_(r.url)] &&
    String(r.guard||'')==='PASS' &&
    String(r.ownership||'')==='DOCTOR_OWNED' &&
    (String(r.priority||'')==='A1_CANDIDATE' || String(r.priority||'')==='A2_CANDIDATE')
  ).length;
}

function sdsdPendingDoctorResultImport_() {
  try {
    const text = sdsdReadSiteWideResultImportText_();
    if (!text) return null;
    const parsed = sdsdExtractJsonObject_(text);
    if (!parsed) return null;
    if (
      parsed.format === 'SIMS_DOCTOR_SITE_WIDE_RESULT_V1' ||
      parsed.format === 'SIMS_DOCTOR_SITE_WIDE_PRECISION_RESULT_V1'
    ) {
      sdsdValidateSiteWideResult_(parsed);
      return parsed;
    }
  } catch (e) {}
  return null;
}

function sdsdSetSiteWidePrecisionPackageState_(state) {
  const props = PropertiesService.getDocumentProperties();
  const obj = state || {};
  props.setProperty('SDSD_SITE_WIDE_PRECISION_PACKAGE_STATUS', String(obj.status || ''));
  props.setProperty('SDSD_SITE_WIDE_PRECISION_PACKAGE_FILE_ID', String(obj.fileId || ''));
  props.setProperty('SDSD_SITE_WIDE_PRECISION_PACKAGE_FILE_URL', String(obj.fileUrl || ''));
  props.setProperty('SDSD_SITE_WIDE_PRECISION_PACKAGE_FILE_NAME', String(obj.fileName || ''));
  props.setProperty('SDSD_SITE_WIDE_PRECISION_PACKAGE_CLUSTER_COUNT', String(obj.clusterCount || 0));
  props.setProperty('SDSD_SITE_WIDE_PRECISION_PACKAGE_AT', new Date().toISOString());
}

function sdsdGetSiteWidePrecisionPackageState_() {
  const props = PropertiesService.getDocumentProperties();
  return {
    status: String(props.getProperty('SDSD_SITE_WIDE_PRECISION_PACKAGE_STATUS') || ''),
    fileId: String(props.getProperty('SDSD_SITE_WIDE_PRECISION_PACKAGE_FILE_ID') || ''),
    fileUrl: String(props.getProperty('SDSD_SITE_WIDE_PRECISION_PACKAGE_FILE_URL') || ''),
    fileName: String(props.getProperty('SDSD_SITE_WIDE_PRECISION_PACKAGE_FILE_NAME') || ''),
    clusterCount: Number(props.getProperty('SDSD_SITE_WIDE_PRECISION_PACKAGE_CLUSTER_COUNT') || 0),
    at: String(props.getProperty('SDSD_SITE_WIDE_PRECISION_PACKAGE_AT') || '')
  };
}

function sdsdClearSiteWidePrecisionPackageState_() {
  const props = PropertiesService.getDocumentProperties();
  [
    'SDSD_SITE_WIDE_PRECISION_PACKAGE_STATUS',
    'SDSD_SITE_WIDE_PRECISION_PACKAGE_FILE_ID',
    'SDSD_SITE_WIDE_PRECISION_PACKAGE_FILE_URL',
    'SDSD_SITE_WIDE_PRECISION_PACKAGE_FILE_NAME',
    'SDSD_SITE_WIDE_PRECISION_PACKAGE_CLUSTER_COUNT',
    'SDSD_SITE_WIDE_PRECISION_PACKAGE_AT'
  ].forEach(k => props.deleteProperty(k));
}

function sdsdShowSiteWidePrecisionResultWaiting_() {
  const state = sdsdGetSiteWidePrecisionPackageState_();
  sdsdShowGuidedActionDialog_({
    title:'Doctorの精密診断結果を待っています',
    purpose:'追加Evidenceを使って、既存記事の役割分担・Merge・Writer等をDoctorに確定してもらう工程です。',
    done:'精密診断Packageは生成済みです。再生成する必要はありません。',
    next:'Doctorから回答が返ってきたら、下のボタンから回答全文を貼り付けてください。',
    fileName:state.fileName||'', fileUrl:state.fileUrl||'',
    primaryLabel:'Doctor診断結果を貼り付ける', primaryFn:'sdsdShowSiteWideDoctorResultDialog'
  });
}

function sdsdMergePrecisionIntoStoredResult_(precisionNormalized) {
  let current = null;
  try { current = sdsdReadStoredSiteWideResult_(); } catch (e) { current = null; }

  if (!current || !Array.isArray(current.diagnosis_cases)) {
    const fallback = JSON.parse(JSON.stringify(precisionNormalized));
    fallback.format = 'SIMS_DOCTOR_SITE_WIDE_RESULT_V1';
    return fallback;
  }

  const normalizedText = v => String(v || '').trim().toLowerCase();
  const normalizedUrl = v => normalizedText(v).replace(/\/+$/, '');
  const resolvedSourceIds = {};
  const resolvedCaseIds = {};
  const resolvedUrls = {};
  const resolvedThemes = {};

  (precisionNormalized.diagnosis_cases || []).forEach(c => {
    const caseId = String(c.diagnosis_case_id || '');
    if (caseId) resolvedCaseIds[caseId] = true;

    (c.absorbed_source_case_ids || []).forEach(id => {
      const sourceId = String(id || '');
      if (sourceId) resolvedSourceIds[sourceId] = true;
    });

    const urls = [];
    if (c.article_url) urls.push(c.article_url);
    if (Array.isArray(c.articles)) {
      c.articles.forEach(a => {
        if (a && a.article_url) urls.push(a.article_url);
      });
    }
    urls.map(normalizedUrl).filter(Boolean).forEach(url => { resolvedUrls[url] = true; });

    const theme = normalizedText(c.diagnosis_theme);
    if (theme) resolvedThemes[theme] = true;
  });

  const remaining = current.diagnosis_cases.filter(c => {
    const id = String(c.diagnosis_case_id || '');
    if (resolvedCaseIds[id] || resolvedSourceIds[id]) return false;

    const absorbed = Array.isArray(c.absorbed_source_case_ids)
      ? c.absorbed_source_case_ids.map(String)
      : [];
    if (absorbed.some(id => resolvedSourceIds[id])) return false;

    // Precision packages can contain legacy/source cases whose IDs are not
    // returned by Doctor. For unresolved NEEDS_EVIDENCE cases only, safely
    // match the finalized result by exact article URL first, then exact theme.
    const route = String(c.route_to || c.doctor_decision || '').toUpperCase();
    const isNeedsEvidence =
      route === 'NEEDS_EVIDENCE' ||
      route === 'ADDITIONAL_EVIDENCE_REQUIRED';

    if (isNeedsEvidence) {
      const urls = [];
      if (c.article_url) urls.push(c.article_url);
      if (Array.isArray(c.articles)) {
        c.articles.forEach(a => {
          if (a && a.article_url) urls.push(a.article_url);
        });
      }
      const urlMatched = urls.map(normalizedUrl).filter(Boolean).some(url => resolvedUrls[url]);
      if (urlMatched) return false;

      const theme = normalizedText(c.diagnosis_theme);
      if (theme && resolvedThemes[theme]) return false;
    }

    return true;
  });

  const merged = JSON.parse(JSON.stringify(current));
  merged.format = 'SIMS_DOCTOR_SITE_WIDE_RESULT_V1';
  merged.generated_at = new Date().toISOString();
  merged.diagnosis_cases = remaining.concat(precisionNormalized.diagnosis_cases || []);
  merged.last_precision_result_at = new Date().toISOString();
  merged.last_precision_result_count = (precisionNormalized.diagnosis_cases || []).length;
  return merged;
}


function sdsdSbmHandoffRoute_(route) {
  route = String(route || '').toUpperCase();
  // SBM's Site Diagnosis bridge treats MONITOR as the non-treatment tracking route.
  // Preserve NO_ACTION in source_route while handing it to SBM as MONITOR.
  return route === 'NO_ACTION' ? 'MONITOR' : route;
}

function sdsdSbmArticleFromTarget_(article) {
  article = article || {};
  const url = String(article.article_url || article.url || '').trim();
  const identity = url ? sdsdArticleIdentityForUrl_(url) : {
    site_id:'', article_id:'', article_title:'', article_url:'', main_query:''
  };
  return {
    site_id: String(article.site_id || identity.site_id || ''),
    article_id: String(article.article_id || identity.article_id || ''),
    article_title: String(
      article.article_title || article.title || identity.article_title || ''
    ),
    article_url: url || String(identity.article_url || ''),
    role: String(article.role || '')
  };
}

function sdsdSbmMergePlanFromCase_(c) {
  if (c && c.merge_plan) return c.merge_plan;

  const articles = (c && Array.isArray(c.target_articles) ? c.target_articles : [])
    .map(sdsdSbmArticleFromTarget_)
    .filter(a => a.article_url);

  let target = null;
  let source = null;

  articles.forEach(a => {
    const role = String(a.role || '');
    if (!target && /維持|軸記事|統合先|PRIMARY|SURVIVOR/i.test(role)) target = a;
    if (!source && /Merge対象|統合元|吸収|リーク|SOURCE|ABSORB/i.test(role)) source = a;
  });

  // Only use positional fallback when exactly two articles exist.
  if ((!target || !source) && articles.length === 2) {
    target = target || articles[0];
    source = source || articles[1];
  }

  if (!target || !source || target.article_url === source.article_url) return null;

  return {
    target_article: target,
    source_article: source,
    redirect_direction: `${source.article_url} → ${target.article_url}`,
    content_to_absorb: String(c && c.merge_content_to_absorb || '')
  };
}

function sdsdBuildSbmSiteDiagnosisHandoff_() {
  const result = sdsdReadStoredSiteWideResult_();
  const storedSite = result.site || {};
  const fallbackSite = sdsdSiteMetaFromArticleMaster_();
  const site = {
    site_id: String(storedSite.site_id || fallbackSite.site_id || ''),
    site_name: String(storedSite.site_name || fallbackSite.site_name || ''),
    site_url: String(storedSite.site_url || fallbackSite.site_url || '')
  };
  const batchId = String(result.site_diagnosis_batch_id || sdsdSiteWideBatchId_() || '');
  const clusters = [];

  (result.diagnosis_cases || []).forEach((c, i) => {
    const sourceRoute = String(c.route_to || '').toUpperCase();
    if (!sourceRoute || sourceRoute === 'NEEDS_EVIDENCE') return;

    const route = sdsdSbmHandoffRoute_(sourceRoute);
    const articles = (Array.isArray(c.target_articles) ? c.target_articles : [])
      .map(sdsdSbmArticleFromTarget_)
      .filter(a => a.article_url);

    // A representative article is mandatory for the existing SBM bridge.
    if (!articles.length) return;

    const clusterResult = {
      diagnosis_theme: String(c.diagnosis_theme || ''),
      diagnosis_summary: String(c.reason || ''),
      doctor_decision: String(c.doctor_decision || sourceRoute),
      confidence: c.confidence,
      site_impact: String(c.site_impact || ''),
      treatment_strategy: String(c.treatment_strategy || ''),
      route_to: route,
      source_route_to: sourceRoute,
      articles: articles,
      allowed_scope: Array.isArray(c.allowed_scope) && c.allowed_scope.length
        ? c.allowed_scope
        : (c.treatment_plan && Array.isArray(c.treatment_plan.allowed_scope)
          ? c.treatment_plan.allowed_scope : []),
      blocked_scope: Array.isArray(c.blocked_scope) && c.blocked_scope.length
        ? c.blocked_scope
        : (c.treatment_plan && Array.isArray(c.treatment_plan.blocked_scope)
          ? c.treatment_plan.blocked_scope : []),
      internal_link_recommendations: Array.isArray(c.internal_link_recommendations)
        ? c.internal_link_recommendations : [],
      presentation: c.presentation || null,
      creator_plan: sourceRoute === 'CREATOR' ? (c.creator_plan || null) : null,
      new_article_target: sourceRoute === 'CREATOR'
        ? (c.new_article_target || (c.creator_plan && c.creator_plan.new_article_target) || null)
        : null,
      reference_articles: sourceRoute === 'CREATOR'
        ? (Array.isArray(c.reference_articles) ? c.reference_articles : articles)
        : [],
      article_identity_semantics: sourceRoute === 'CREATOR'
        ? {
            primary_target_type:'NEW_ARTICLE',
            legacy_bridge_articles_are_reference_only:true,
            legacy_bridge_identity_role:'REFERENCE_ONLY'
          }
        : null
    };

    if (route === 'WRITER' && !clusterResult.allowed_scope.length) {
      throw new Error(
        `Writer案件にDoctorのallowed_scopeがありません：${c.diagnosis_theme || c.diagnosis_case_id || (i+1)}。` +
        ' Doctor精密診断結果のworkflow_handoff.allowed_scopeまたはクラスタ別allowed_scopeを確認してください。'
      );
    }

    if (sourceRoute === 'MERGE') {
      const mergePlan = sdsdSbmMergePlanFromCase_(c);
      if (!mergePlan) {
        throw new Error(
          `Merge案件の統合方向を安全に確定できません：${c.diagnosis_theme || c.diagnosis_case_id || (i+1)}`
        );
      }
      clusterResult.merge_plan = mergePlan;
    }

    clusters.push({
      diagnosis_case_id: String(c.diagnosis_case_id || `SDSD-${i+1}`),
      site_diagnosis_batch_id: batchId,
      site_id: String(site.site_id || ''),
      diagnosis_theme: String(c.diagnosis_theme || ''),
      workflow_handoff: {
        next_action: route,
        handoff_mode: route === 'MONITOR'
          ? 'RETURN_TO_SBM_FOR_MONITORING'
          : 'RETURN_TO_SBM_FOR_REFERRAL',
        allowed_scope: clusterResult.allowed_scope,
        blocked_scope: clusterResult.blocked_scope
      },
      cluster_result: clusterResult
    });
  });

  if (!clusters.length) throw new Error('SBMへ引き渡す確定案件がありません。');

  return {
    format: 'SIMS_DOCTOR_SITE_WIDE_PRECISION_RESULT_V1',
    contract_name: 'SIMS_DOCTOR_SITE_WIDE_PRECISION_RESULT_V1',
    contract_version: '1.0',
    generated_at: new Date().toISOString(),
    site_diagnosis_batch_id: batchId,
    site_id: String(site.site_id || ''),
    site_name: String(site.site_name || ''),
    site_url: String(site.site_url || ''),
    source_system: 'SIMS_DOCTOR_SITE_DIAGNOSIS',
    target_system: 'SIMS_BLOG_MANAGER',
    clusters: clusters,
    workflow: {return_to:'SIMS_BLOG_MANAGER'}
  };
}

function sdsdSetSbmHandoffState_(status) {
  const props = PropertiesService.getDocumentProperties();
  props.setProperty('SDSD_SBM_HANDOFF_STATUS', String(status || ''));
  props.setProperty('SDSD_SBM_HANDOFF_AT', new Date().toISOString());
}

function sdsdGetSbmHandoffState_() {
  return String(
    PropertiesService.getDocumentProperties().getProperty('SDSD_SBM_HANDOFF_STATUS') || ''
  );
}

function sdsdSbmHandoffFingerprint_() {
  const obj = sdsdBuildSbmSiteDiagnosisHandoff_();
  const stable = (obj.clusters || []).map(c => {
    const cr = c.cluster_result || {};
    const cp = cr.creator_plan || {};
    const mp = cr.merge_plan || {};
    return {
      diagnosis_case_id:String(c.diagnosis_case_id || ''),
      next_action:String(c.workflow_handoff && c.workflow_handoff.next_action || ''),
      diagnosis_theme:String(c.diagnosis_theme || ''),
      source_route_to:String(cr.source_route_to || ''),
      creator_keyword:String(cp.candidate_keyword || ''),
      creator_role:String(cp.role_with_existing_articles || ''),
      creator_monitor_days:Number(cp.monitor_days || 0),
      merge_direction:String(mp.redirect_direction || ''),
      article_urls:(Array.isArray(cr.articles) ? cr.articles : []).map(a => String(a.article_url || '')).sort()
    };
  }).sort((a,b) => a.diagnosis_case_id.localeCompare(b.diagnosis_case_id));
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    JSON.stringify(stable),
    Utilities.Charset.UTF_8
  );
  return digest.map(b => ('0' + ((b + 256) % 256).toString(16)).slice(-2)).join('');
}

function sdsdIsSbmHandoffCurrent_() {
  const props = PropertiesService.getDocumentProperties();
  if (sdsdGetSbmHandoffState_() !== 'COMPLETE') return false;
  const saved = String(props.getProperty('SDSD_SBM_HANDOFF_FINGERPRINT') || '');
  // Older versions did not persist a payload fingerprint. Force one fresh handoff
  // so newly-created Creator cases in an existing diagnosis session are not skipped.
  if (!saved) return false;
  try { return saved === sdsdSbmHandoffFingerprint_(); } catch (e) { return false; }
}

function sdsdShowSbmHandoffDialog_() {
  const ui = SpreadsheetApp.getUi();
  const obj = sdsdBuildSbmSiteDiagnosisHandoff_();
  const text = JSON.stringify(obj, null, 2);
  const encoded = Utilities.base64EncodeWebSafe(text, Utilities.Charset.UTF_8);

  const html =
    '<!doctype html><html><head><base target="_top"><meta charset="UTF-8"><style>' +
    'body{font-family:Arial,"Noto Sans JP",sans-serif;padding:18px;background:#f8f9fa;color:#202124}' +
    'h2{margin:0 0 8px;font-size:18px}.flow{background:#e8f0fe;color:#174ea6;padding:9px 11px;border-radius:7px;font-weight:700;font-size:12px;margin-bottom:10px}' +
    '.note{font-size:12px;line-height:1.7;color:#5f6368;margin-bottom:10px}' +
    'textarea{box-sizing:border-box;width:100%;height:350px;padding:10px;font:12px/1.45 monospace;white-space:pre;border:1px solid #bdc1c6;border-radius:7px;background:#fff}' +
    '.actions{display:flex;gap:8px;justify-content:flex-end;margin-top:10px}button{padding:9px 15px;border:0;border-radius:6px;font-weight:700;cursor:pointer}.primary{background:#1a73e8;color:#fff}.done{background:#137333;color:#fff}.secondary{background:#e8eaed;color:#202124}.ok{font-size:12px;color:#137333;margin-top:8px}' +
    '</style></head><body><h2>SBMへ診断結果を引き渡す</h2>' +
    '<div class="flow">Site Diagnosis → Doctor → Diagnosis → SBM → Writer / Creator / Merge</div>' +
    '<div class="note">下のJSONをコピーし、SIMS-Blog-Managerの「SIMS Doctor → 5．Site Diagnosisの処置を進める」→「① DiagnosisからのDoctor診断結果を登録」へ貼り付けてください。SBMがWriter / Creator / Merge / 経過観察へ振り分けます。</div>' +
    '<textarea id="t" readonly></textarea><div class="actions"><button class="secondary" onclick="google.script.host.close()">閉じる</button><button class="primary" onclick="copyText()">SBM用JSONをコピー</button><button class="done" onclick="done()">SBMへの登録完了</button></div><div id="s" class="ok"></div>' +
    '<script>const raw="' + encoded + '";function dec(x){x=x.replace(/-/g,"+").replace(/_/g,"/");while(x.length%4)x+="=";return decodeURIComponent(escape(atob(x)))}const t=document.getElementById("t");t.value=dec(raw);function copyText(){t.select();t.setSelectionRange(0,999999);navigator.clipboard.writeText(t.value).then(()=>document.getElementById("s").textContent="コピーしました。SBMへ貼り付けてください。").catch(()=>{document.execCommand("copy");document.getElementById("s").textContent="コピーしました。SBMへ貼り付けてください。"})}function done(){const s=document.getElementById("s");s.textContent="登録完了を記録しています...";google.script.run.withSuccessHandler(()=>{s.textContent="SBMへの登録完了を記録しました。";setTimeout(()=>google.script.host.close(),500)}).withFailureHandler(err=>{const msg=(err&&err.message)?err.message:String(err||"不明なエラー");s.textContent="登録完了を記録できませんでした："+msg}).sdsdMarkSbmHandoffComplete()}</script></body></html>';

  sdsdSetSbmHandoffState_('WAITING_SBM_REGISTRATION');
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(840).setHeight(650),
    'SBMへ診断結果を引き渡す'
  );
}

function sdsdMarkSbmHandoffComplete() {
  return sdsdMarkSbmHandoffComplete_();
}

function sdsdMarkSbmHandoffComplete_() {
  const props = PropertiesService.getDocumentProperties();
  props.setProperty('SDSD_SBM_HANDOFF_FINGERPRINT', sdsdSbmHandoffFingerprint_());
  sdsdSetSbmHandoffState_('COMPLETE');
  try {
    const stored = sdsdReadStoredSiteWideResult_();
    if (stored && Array.isArray(stored.diagnosis_cases)) {
      sdsdWriteTreatmentPlan_(stored);
    }
  } catch (e) {}
  try { sdsdRenderHome_(); } catch (e) {}
  return true;
}


function sdsdProceedNextGuided(){
  sdsdOpenWorkflowController_();
}

function sdsdWorkflowProps_(){
  return PropertiesService.getDocumentProperties();
}

function sdsdGetWorkflowState_(){
  const p=sdsdWorkflowProps_();
  return {
    route:String(p.getProperty('SDSD_WORKFLOW_ROUTE')||''),
    status:String(p.getProperty('SDSD_WORKFLOW_STATUS')||''),
    recommendation:String(p.getProperty('SDSD_WORKFLOW_RECOMMENDATION')||''),
    recommendationReason:String(p.getProperty('SDSD_WORKFLOW_RECOMMENDATION_REASON')||''),
    updatedAt:String(p.getProperty('SDSD_WORKFLOW_UPDATED_AT')||'')
  };
}

function sdsdSetWorkflowState_(patch){
  const p=sdsdWorkflowProps_();
  const map={
    route:'SDSD_WORKFLOW_ROUTE',
    status:'SDSD_WORKFLOW_STATUS',
    recommendation:'SDSD_WORKFLOW_RECOMMENDATION',
    recommendationReason:'SDSD_WORKFLOW_RECOMMENDATION_REASON',
    updatedAt:'SDSD_WORKFLOW_UPDATED_AT'
  };
  Object.keys(patch||{}).forEach(k=>{
    if(!map[k])return;
    const v=patch[k];
    if(v===null||v===undefined||v==='')p.deleteProperty(map[k]);
    else p.setProperty(map[k],String(v));
  });
  p.setProperty('SDSD_WORKFLOW_UPDATED_AT',new Date().toISOString());
}

function sdsdClearWorkflowState_(){
  const p=sdsdWorkflowProps_();
  [
    'SDSD_WORKFLOW_ROUTE',
    'SDSD_WORKFLOW_STATUS',
    'SDSD_WORKFLOW_RECOMMENDATION',
    'SDSD_WORKFLOW_RECOMMENDATION_REASON',
    'SDSD_WORKFLOW_UPDATED_AT'
  ].forEach(k=>p.deleteProperty(k));
}

function sdsdWorkflowRecommendation_(){
  const m=sdsdHomeDiagnosisMetrics_();

  const existingScore =
    Number(m.a1||0)*6 +
    Number(m.a2||0)*4 +
    Number(m.severe||0)*4 +
    Number(m.traffic||0)*2 +
    Number(m.ranking||0)*2 +
    Number(m.volatile||0);

  const creatorScore =
    Number(m.newArticle||0)*5 +
    Number(m.gap||0)*3;

  let route='EXISTING';
  let reason='既存記事の悪化・優先診断候補を先に処置した方が、現在ある検索流入を守る効果が高いと判断しました。';

  if(creatorScore>existingScore){
    route='CREATOR';
    reason='新規記事機会・コンテンツギャップが既存記事の緊急処置より強く、先に新しい検索需要を取りに行く価値が高いと判断しました。';
  }else if(existingScore===0 && creatorScore>0){
    route='CREATOR';
    reason='緊急性の高い既存記事候補がなく、新規記事機会が確認されたためCreatorルートを優先します。';
  }else if(existingScore===0 && creatorScore===0){
    route='EXISTING';
    reason='明確な新規記事機会がまだ確認されていないため、既存記事側の確認を先に行います。';
  }

  return {
    route:route,
    reason:reason,
    metrics:m,
    existingScore:existingScore,
    creatorScore:creatorScore
  };
}

function sdsdWorkflowRouteLabel_(route){
  return String(route||'').toUpperCase()==='CREATOR'
    ? '新記事を見つけてCreatorへ送る'
    : '既存記事を診断・治療する';
}

function sdsdWorkflowStep_(route){
  route=String(route||'').toUpperCase();
  const session=sdsdGetCurrentSession_();
  const metrics=sdsdHomeDiagnosisMetrics_();
  const work=sdsdSessionWorkSummary_();
  const stored=sdsdHomeReadStoredSiteWideResult_();
  const props=PropertiesService.getDocumentProperties();
  const stage=String(props.getProperty('SDSD_SITE_WIDE_REGISTER_STAGE')||'');

  if(!session.active){
    return {
      code:'IMPORT_EVIDENCE',
      title:'CollectorのEvidence Packageを読み込みます',
      detail:'最初にCollectorから受け取ったZIPを選択します。',
      button:'Evidence Packageを選ぶ'
    };
  }

  if(!metrics.total){
    return {
      code:'RUN_ANALYSIS',
      title:'サイト全体を分析します',
      detail:'既存記事の状態と、新記事の検索機会を同時に分析して、どちらを先に進めるべきかDiagnosisが提案します。',
      button:'サイト全体を分析する'
    };
  }

  const pendingDoctorResult=sdsdPendingDoctorResultImport_();
  if(pendingDoctorResult){
    return {
      code:'IMPORT_SITEWIDE_RESULT',
      title:'Doctorの回答を登録します',
      detail:'貼り付け済みのDoctor回答があります。内容を検証してDiagnosisへ登録します。',
      button:'Doctor回答を登録する'
    };
  }

  if(stage==='WAITING_INPUT'){
    return {
      code:'SHOW_SITEWIDE_INPUT',
      title:'Doctorの回答待ちです',
      detail:'生成済みPackageをDoctorへ渡し、返ってきた回答全文をこのダイアログから貼り付けます。',
      button:'Doctor回答を貼り付ける'
    };
  }

  const creatorSerpState=sdsdGetCreatorSerpState_();
  if(creatorSerpState.status==='WAITING_DOCTOR_RESULT'){
    return {
      code:'SHOW_CREATOR_SERP_INPUT',
      title:'Creator候補のSERP確認結果を待っています',
      detail:'Doctorから返ったSERP確認結果を貼り付けて、新記事として進めてよいか確定します。',
      button:'SERP確認結果を貼り付ける'
    };
  }

  if(stored && work.actionableTreatment>0 && !sdsdIsSbmHandoffCurrent_()){
    return {
      code:'SBM_HANDOFF',
      title:'確定した案件をSBMへ引き渡します',
      detail:'Doctorで処置が確定した案件をSBMへ渡します。Writer・Merge・Creatorの実処置はSBM側で続けます。',
      button:'SBM引き渡しを進める'
    };
  }

  if(route==='CREATOR'){
    if(stored){
      const creatorPending=sdsdCreatorValidationCases_().length;
      if(creatorPending>0){
        const creatorSheet=SpreadsheetApp.getActive().getSheetByName(SDSD_CONFIG.sheets.creatorValidation);
        return {
          code:creatorSheet?'OPEN_CREATOR_VALIDATION':'RUN_CREATOR_VALIDATION',
          title:'新記事候補の作成前チェックを行います',
          detail:'既存記事との重大なカニバリを避けながら、独立した検索意図ならCreator候補として前進させます。',
          button:creatorSheet?'Creator候補チェックを開く':'Creator候補チェックを作る'
        };
      }
      if(work.additionalEvidence>0){
        return {
          code:'EXPORT_PRECISION',
          title:'追加Evidenceで候補を詳しく確認します',
          detail:'Doctorが追加確認を必要とした候補について、精密診断Packageを生成します。',
          button:'追加Evidence Packageを作る'
        };
      }
    }

    const precisionPackage=sdsdGetSiteWidePrecisionPackageState_();
    if(precisionPackage.status==='WAITING_DOCTOR_RESULT'){
      return {
        code:'WAIT_PRECISION',
        title:'Doctorの精密診断結果を待っています',
        detail:'Packageは生成済みです。再生成せず、Doctorの回答が返ったら貼り付けます。',
        button:'回答入力画面を開く'
      };
    }

    if(work.opportunityCases>0){
      return {
        code:'EXPORT_SITEWIDE',
        title:'新記事機会をDoctorへ渡します',
        detail:'サイト全体から見つかった新規記事機会・コンテンツギャップ等をDoctorが整理できるPackageを作ります。',
        button:'Doctor Packageを作る'
      };
    }

    if(metrics.crossTotal>0){
      return {
        code:'BUILD_OPPORTUNITIES',
        title:'新記事機会を案件化します',
        detail:'サイト全体分析で見つかった検索機会を、Doctorへ渡せる案件に整理します。',
        button:'新記事機会を整理する'
      };
    }

    return {
      code:'ROUTE_COMPLETE',
      title:'Creator優先ルートは現在完了しています',
      detail:'現在このルートで追加作業はありません。もう一方の「既存記事を診断・治療する」ルートへ切り替えられます。',
      button:'既存記事ルートへ切り替える'
    };
  }

  const precisionPackage=sdsdGetSiteWidePrecisionPackageState_();
  if(precisionPackage.status==='WAITING_DOCTOR_RESULT'){
    return {
      code:'WAIT_PRECISION',
      title:'Doctorの精密診断結果を待っています',
      detail:'Packageは生成済みです。Doctorの回答が返ったら登録してください。',
      button:'回答入力画面を開く'
    };
  }

  const individualPackage=sdsdGetIndividualDoctorPackageState_();
  if(individualPackage.status==='PACKAGE_READY_FOR_DOCTOR'){
    return {
      code:'INDIVIDUAL_PACKAGE_READY',
      title:'Doctorへ個別精密診断を依頼してください',
      detail:`${Number(individualPackage.caseCount||0)}記事の精密診断Packageは生成済みです。${String(individualPackage.fileName||'生成済みPackage')} をSIMS Doctorへ渡してください。再生成する必要はありません。`,
      button:'Packageと次の手順を確認する'
    };
  }
  if(individualPackage.status==='WAITING_DOCTOR_RESULT'){
    return {
      code:'WAIT_INDIVIDUAL',
      title:'Doctorの個別精密診断結果を待っています',
      detail:'Doctorへ依頼済みのPackageがあります。回答が返ったら、各記事の診断結果をSBMへ登録してください。',
      button:'登録状況を確認する'
    };
  }

  const individualProgress=sdsdIndividualEnrichmentProgress_();
  if(individualProgress.active && individualProgress.pending>0){
    return {
      code:'INDIVIDUAL_PRECISION',
      title:`個別精密診断を再開します（${individualProgress.ready}/${individualProgress.total}件準備済み）`,
      detail:`全${individualProgress.total}件のうち${individualProgress.ready}件は準備済みです。残り${individualProgress.pending}件から再開し、準備済みの記事は作り直しません。`,
      button:`残り${individualProgress.pending}件の準備を続ける`
    };
  }

  const individualEligible=sdsdEligibleIndividualPrecisionCount_();
  if(individualEligible>0){
    const workNow=sdsdSessionWorkSummary_();
    return {
      code:'INDIVIDUAL_PRECISION',
      title:`優先記事${individualEligible}件を詳しく診断します`,
      detail:`追加確認が必要な案件は${Number(workNow.additionalEvidence||0)}件あります。その中から優先度の高い${individualEligible}記事を先に個別精密診断します。`,
      button:'個別精密診断を進める'
    };
  }

  if(stored && work.additionalEvidence>0){
    return {
      code:'EXPORT_PRECISION',
      title:'追加Evidenceで既存記事を詳しく診断します',
      detail:'現在の情報だけでは処置を確定できない案件について、本文等を追加したPackageを作ります。',
      button:'追加Evidence Packageを作る'
    };
  }

  return {
    code:'ROUTE_COMPLETE',
    title:'既存記事の優先ルートは現在完了しています',
    detail:'現在このルートで追加作業はありません。もう一方の「新記事を見つけてCreatorへ送る」ルートへ切り替えられます。',
    button:'Creatorルートへ切り替える'
  };
}

function sdsdWorkflowEsc_(v){
  return String(v==null?'':v)
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

function sdsdShowWorkflowRecommendationDialog_(){
  const rec=sdsdWorkflowRecommendation_();
  sdsdSetWorkflowState_({
    status:'ROUTE_RECOMMENDED',
    recommendation:rec.route,
    recommendationReason:rec.reason
  });

  const m=rec.metrics;
  const recommended=sdsdWorkflowRouteLabel_(rec.route);
  const other=sdsdWorkflowRouteLabel_(rec.route==='CREATOR'?'EXISTING':'CREATOR');

  const html=`<!doctype html><html><head><base target="_top"><style>
    body{font-family:Arial,"Noto Sans JP",sans-serif;margin:0;background:#f8fafd;color:#202124}
    .wrap{padding:20px}.hero{background:#185abc;color:#fff;border-radius:12px;padding:18px}
    h2{margin:0 0 6px;font-size:21px}.sub{font-size:13px;line-height:1.6}
    .card{background:#fff;border:1px solid #dadce0;border-radius:10px;padding:14px;margin-top:14px}
    .rec{border:2px solid #1a73e8}.tag{display:inline-block;background:#e8f0fe;color:#174ea6;font-weight:bold;border-radius:999px;padding:4px 9px;font-size:12px}
    .title{font-weight:bold;font-size:17px;margin-top:8px}.text{font-size:13px;line-height:1.7;margin-top:7px}
    .metrics{display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-top:10px}.metric{background:#f1f3f4;border-radius:7px;padding:9px;font-size:12px}
    .actions{text-align:right;margin-top:16px}button{border:1px solid #dadce0;border-radius:7px;background:#fff;padding:9px 13px;margin-left:6px;cursor:pointer}
    .primary{background:#1a73e8;color:#fff;border-color:#1a73e8;font-weight:bold}.pause{float:left;color:#5f6368}
  </style></head><body><div class="wrap">
    <div class="hero"><h2>サイト全体の分析が完了しました</h2><div class="sub">Diagnosisが2つの大きな流れを比較し、先に進めるルートを提案します。</div></div>
    <div class="card rec"><span class="tag">Diagnosisの推奨</span><div class="title">${sdsdWorkflowEsc_(recommended)}</div>
      <div class="text">${sdsdWorkflowEsc_(rec.reason)}</div>
      <div class="metrics">
        <div class="metric">既存記事：A1 ${Number(m.a1||0)}件 / A2 ${Number(m.a2||0)}件 / 大幅悪化 ${Number(m.severe||0)}件</div>
        <div class="metric">新記事：新規記事機会 ${Number(m.newArticle||0)}件 / コンテンツギャップ ${Number(m.gap||0)}件</div>
      </div>
    </div>
    <div class="card"><div class="title">もう一方の流れ</div><div class="text">${sdsdWorkflowEsc_(other)}。推奨とは別にこちらを先に選ぶこともできます。</div></div>
    <div class="actions">
      <button class="pause" onclick="google.script.host.close()">中断して閉じる</button>
      <button onclick="${rec.route==='CREATOR'?'chooseExisting()':'chooseCreator()'}">もう一方を先に進める</button>
      <button class="primary" onclick="${rec.route==='CREATOR'?'chooseCreator()':'chooseExisting()'}">推奨ルートで進める</button>
    </div>
    <div id="status" style="font-size:12px;color:#5f6368;text-align:right;margin-top:8px"></div>
  </div><script>
    function chooseExisting(){
      document.getElementById('status').textContent='保存しています...';
      google.script.run.withSuccessHandler(()=>google.script.host.close())
        .withFailureHandler(e=>document.getElementById('status').textContent=(e&&e.message)?e.message:String(e))
        .sdsdSelectExistingWorkflow();
    }
    function chooseCreator(){
      document.getElementById('status').textContent='保存しています...';
      google.script.run.withSuccessHandler(()=>google.script.host.close())
        .withFailureHandler(e=>document.getElementById('status').textContent=(e&&e.message)?e.message:String(e))
        .sdsdSelectCreatorWorkflow();
    }
  </script></body></html>`;

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(690).setHeight(570),
    'Site Diagnosisを進める'
  );
}

function sdsdShowWorkflowProgressDialog_(){
  const st=sdsdGetWorkflowState_();
  const route=st.route||st.recommendation||'EXISTING';
  const step=sdsdWorkflowStep_(route);
  sdsdSetWorkflowState_({route:route,status:step.code});

  const routeLabel=sdsdWorkflowRouteLabel_(route);
  const resume=!!st.updatedAt;
  const html=`<!doctype html><html><head><base target="_top"><style>
    body{font-family:Arial,"Noto Sans JP",sans-serif;margin:0;background:#f8fafd;color:#202124}.wrap{padding:20px}
    .hero{background:#185abc;color:#fff;border-radius:12px;padding:17px}.hero h2{margin:0 0 6px;font-size:20px}.hero p{margin:0;font-size:13px;line-height:1.6}
    .route{background:#e8f0fe;color:#174ea6;border-radius:8px;padding:10px 12px;margin-top:13px;font-weight:bold}
    .card{background:#fff;border:1px solid #dadce0;border-radius:10px;margin-top:13px;padding:14px}.label{font-size:12px;color:#5f6368;font-weight:bold}
    .step{font-size:18px;font-weight:bold;margin-top:5px}.detail{font-size:13px;line-height:1.7;margin-top:7px}
    .resume{font-size:12px;color:#137333;margin-top:9px}.actions{text-align:right;margin-top:16px}
    button{border:1px solid #dadce0;background:#fff;border-radius:7px;padding:9px 13px;margin-left:6px;cursor:pointer}
    .primary{background:#1a73e8;color:#fff;border-color:#1a73e8;font-weight:bold}.pause{float:left;color:#5f6368}
  </style></head><body><div class="wrap">
    <div class="hero"><h2>Site Diagnosisを進める</h2><p>必要な入出力をこの進行画面から順番に行います。途中で閉じても、次回は現在地から再開できます。</p></div>
    <div class="route">現在の優先ルート：${sdsdWorkflowEsc_(routeLabel)}</div>
    <div class="card">
      <div class="label">現在の工程</div><div class="step">${sdsdWorkflowEsc_(step.title)}</div>
      <div class="detail">${sdsdWorkflowEsc_(step.detail)}</div>
      ${resume?'<div class="resume">前回の作業状態を保存済みです。この工程から再開します。</div>':''}
    </div>
    <div class="actions">
      <button class="pause" onclick="google.script.host.close()">中断して閉じる</button>
      <button onclick="switchRoute()">もう一方へ切り替える</button>
      <button class="primary" onclick="advance()">${sdsdWorkflowEsc_(step.button)}</button>
    </div>
    <div id="status" style="font-size:12px;color:#5f6368;text-align:right;margin-top:8px"></div>
  </div><script>
    function advance(){
      document.getElementById('status').textContent='処理しています...';
      google.script.run.withSuccessHandler(()=>google.script.host.close())
        .withFailureHandler(e=>document.getElementById('status').textContent=(e&&e.message)?e.message:String(e))
        .sdsdWorkflowAdvance();
    }
    function switchRoute(){
      document.getElementById('status').textContent='ルートを切り替えています...';
      google.script.run.withSuccessHandler(()=>google.script.host.close())
        .withFailureHandler(e=>document.getElementById('status').textContent=(e&&e.message)?e.message:String(e))
        .sdsdSwitchWorkflowRoute();
    }
  </script></body></html>`;

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(690).setHeight(535),
    'Site Diagnosisを進める'
  );
}

function sdsdOpenWorkflowController_(){
  const ui=SpreadsheetApp.getUi();
  try{
    sdsdProductEnsureSheets_();
    const session=sdsdGetCurrentSession_();
    if(!session.active){
      sdsdImportEvidencePackageZip();
      return;
    }

    const metrics=sdsdHomeDiagnosisMetrics_();
    if(!metrics.total){
      sdsdRunProductDiagnosis();
      return;
    }

    const st=sdsdGetWorkflowState_();
    if(!st.route){
      sdsdShowWorkflowRecommendationDialog_();
      return;
    }
    sdsdShowWorkflowProgressDialog_();
  }catch(e){
    ui.alert(`Site Diagnosisを進められませんでした。\n\n${e.message||e}`);
    throw e;
  }
}

function sdsdSelectExistingWorkflow(){
  sdsdSetWorkflowState_({route:'EXISTING',status:'ROUTE_SELECTED'});
  sdsdShowWorkflowProgressDialog_();
}

function sdsdSelectCreatorWorkflow(){
  sdsdSetWorkflowState_({route:'CREATOR',status:'ROUTE_SELECTED'});
  sdsdShowWorkflowProgressDialog_();
}

function sdsdSwitchWorkflowRoute(){
  const st=sdsdGetWorkflowState_();
  const next=String(st.route||'EXISTING').toUpperCase()==='CREATOR'?'EXISTING':'CREATOR';
  sdsdSetWorkflowState_({route:next,status:'ROUTE_SWITCHED'});
  sdsdShowWorkflowProgressDialog_();
}

function sdsdWorkflowAdvance(){
  const st=sdsdGetWorkflowState_();
  const route=st.route||st.recommendation||'EXISTING';
  const step=sdsdWorkflowStep_(route);
  sdsdSetWorkflowState_({route:route,status:step.code});

  switch(step.code){
    case 'IMPORT_EVIDENCE': sdsdImportEvidencePackageZip(); return;
    case 'RUN_ANALYSIS': sdsdRunProductDiagnosis(); return;
    case 'IMPORT_SITEWIDE_RESULT': sdsdRegisterSiteWideDoctorResult(); return;
    case 'SHOW_SITEWIDE_INPUT': sdsdShowSiteWideDoctorResultDialog(); return;
    case 'SHOW_CREATOR_SERP_INPUT': sdsdShowCreatorSerpResultDialog(); return;
    case 'SBM_HANDOFF': sdsdShowSbmHandoffDialog_(); return;
    case 'RUN_CREATOR_VALIDATION':
      sdsdRunCreatorCandidateValidation();
      sdsdShowWorkflowProgressDialog_(); return;
    case 'OPEN_CREATOR_VALIDATION': sdsdOpenCreatorValidation(); return;
    case 'EXPORT_PRECISION': sdsdExportPriorityPrecisionClusterPackage(); return;
    case 'WAIT_PRECISION': sdsdShowSiteWidePrecisionResultWaiting_(); return;
    case 'EXPORT_SITEWIDE': sdsdExportSiteWideDoctorPackage(); return;
    case 'BUILD_OPPORTUNITIES':
      sdsdBuildSiteOpportunityCases();
      sdsdShowWorkflowProgressDialog_(); return;
    case 'INDIVIDUAL_PACKAGE_READY': sdsdShowIndividualPackageReady_(); return;
    case 'WAIT_INDIVIDUAL': sdsdShowIndividualDoctorResultWaiting_(); return;
    case 'INDIVIDUAL_PRECISION': sdsdProceedIndividualPrecisionDiagnosis(); return;
    case 'ROUTE_COMPLETE': sdsdSwitchWorkflowRoute(); return;
    default: sdsdShowWorkflowProgressDialog_(); return;
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
      `次に「確認する → 個別精密診断対象を見る」で内容を確認してください。`
    );
  } catch (e) {
    SpreadsheetApp.getUi().alert(
      `Treatment Batchを作成できませんでした。\n\n${e.message || e}`
    );
    throw e;
  }
}

function sdsdCreateProductCasePackage() {
  // Backward-compatible entry point. RC9.3.2 normal operation no longer
  // requires a manually imported Article Master.
  return sdsdProceedIndividualPrecisionDiagnosis();
}

// ============================================================================
// Source: ArticleFetcher.gs
// ============================================================================
function sdsdFetchArticleEvidence_(url) {
  const requestedUrl = String(url || '').trim();
  const fetchOptions = {
    muteHttpExceptions: true,
    followRedirects: false,
    headers: {
      'User-Agent': 'Mozilla/5.0 (compatible; SIMS-Doctor-Site-Diagnosis/0.5.8)'
    }
  };
  const redirectCodes = {301:true, 302:true, 303:true, 307:true, 308:true};
  const maxRedirects = 5;
  const redirectChain = [];
  let currentUrl = requestedUrl;
  let initialHttpStatus = 0;
  let redirectTargetUrl = '';
  let res = null;
  let code = 0;

  for (let hop = 0; hop <= maxRedirects; hop++) {
    res = UrlFetchApp.fetch(currentUrl, fetchOptions);
    code = res.getResponseCode();
    if (hop === 0) initialHttpStatus = code;

    if (!redirectCodes[code]) break;

    const location = sdsdGetResponseHeader_(res, 'Location');
    if (!location) {
      return sdsdFetchArticleError_(requestedUrl, initialHttpStatus, currentUrl, code,
        redirectChain, redirectTargetUrl, `HTTP ${code}: Location header not found.`);
    }

    const nextUrl = sdsdResolveRedirectUrl_(currentUrl, location);
    if (!nextUrl) {
      return sdsdFetchArticleError_(requestedUrl, initialHttpStatus, currentUrl, code,
        redirectChain, redirectTargetUrl, `HTTP ${code}: Invalid redirect target.`);
    }
    if (!redirectTargetUrl) redirectTargetUrl = nextUrl;

    redirectChain.push({
      from_url: currentUrl,
      http_status: code,
      to_url: nextUrl
    });

    if (redirectChain.some((x, i) => i < redirectChain.length - 1 &&
        sdsdNormalizeUrl_(x.from_url) === sdsdNormalizeUrl_(nextUrl))) {
      return sdsdFetchArticleError_(requestedUrl, initialHttpStatus, currentUrl, code,
        redirectChain, redirectTargetUrl, 'Redirect loop detected.');
    }

    currentUrl = nextUrl;

    if (hop === maxRedirects) {
      return sdsdFetchArticleError_(requestedUrl, initialHttpStatus, currentUrl, code,
        redirectChain, redirectTargetUrl, `Redirect chain exceeded ${maxRedirects} hops.`);
    }
  }

  const finalUrl = currentUrl;
  const finalHttpStatus = code;
  const redirectDetected = redirectChain.length > 0;

  if (finalHttpStatus < 200 || finalHttpStatus >= 400) {
    return sdsdFetchArticleError_(requestedUrl, initialHttpStatus, finalUrl, finalHttpStatus,
      redirectChain, redirectTargetUrl, `HTTP ${finalHttpStatus}`);
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

  const canonicalMatchesFinalUrl = Boolean(
    canonicalUrl && sdsdNormalizeUrl_(canonicalUrl) === sdsdNormalizeUrl_(finalUrl)
  );
  const redirectCanonicalConsistent = Boolean(
    redirectDetected && canonicalUrl && canonicalMatchesFinalUrl
  );

  let articleHtml =
    sdsdFirstMatch_(html, /(<article\b[\s\S]*?<\/article>)/i) ||
    sdsdFirstMatch_(html, /(<div[^>]+class=["'][^"']*entry-content[^"']*["'][^>]*>[\s\S]*?<\/div>)/i) ||
    '';

  const common = {
    httpStatus: finalHttpStatus,
    requestedUrl,
    initialHttpStatus,
    redirectDetected,
    redirectTargetUrl,
    redirectChain,
    finalUrl,
    finalHttpStatus,
    canonicalUrl,
    canonicalMatchesFinalUrl,
    redirectCanonicalConsistent
  };

  if (!articleHtml) {
    return Object.assign({}, common, {
      status: 'BODY_NOT_FOUND',
      title,
      metaDescription,
      articleHtml: '',
      pageHtml: '',
      error: 'Article body container not found.'
    });
  }

  return Object.assign({}, common, {
    status: 'VALID',
    title,
    metaDescription,
    articleHtml,
    pageHtml: '',
    error: ''
  });
}

function sdsdFetchArticleError_(requestedUrl, initialHttpStatus, finalUrl, finalHttpStatus,
  redirectChain, redirectTargetUrl, error) {
  return {
    status: 'FETCH_ERROR',
    httpStatus: finalHttpStatus,
    requestedUrl: requestedUrl,
    initialHttpStatus: initialHttpStatus,
    redirectDetected: Array.isArray(redirectChain) && redirectChain.length > 0,
    redirectTargetUrl: redirectTargetUrl || '',
    redirectChain: redirectChain || [],
    finalUrl: finalUrl || requestedUrl,
    finalHttpStatus: finalHttpStatus,
    title: '',
    metaDescription: '',
    canonicalUrl: '',
    canonicalMatchesFinalUrl: false,
    redirectCanonicalConsistent: false,
    articleHtml: '',
    pageHtml: '',
    error: error || 'Fetch failed.'
  };
}

function sdsdGetResponseHeader_(res, name) {
  const headers = res && res.getHeaders ? res.getHeaders() : {};
  const target = String(name || '').toLowerCase();
  const keys = Object.keys(headers || {});
  for (let i = 0; i < keys.length; i++) {
    if (String(keys[i]).toLowerCase() === target) return String(headers[keys[i]] || '').trim();
  }
  return '';
}

function sdsdResolveRedirectUrl_(baseUrl, location) {
  const loc = String(location || '').trim();
  if (!loc) return '';
  if (/^https?:\/\//i.test(loc)) return loc;

  const base = String(baseUrl || '').trim();
  const originMatch = base.match(/^(https?:\/\/[^\/]+)/i);
  if (!originMatch) return '';
  const origin = originMatch[1];

  if (loc.indexOf('//') === 0) {
    const scheme = (base.match(/^(https?):/i) || [,'https'])[1];
    return scheme + ':' + loc;
  }
  if (loc.charAt(0) === '/') return origin + loc;

  const withoutQuery = base.split('#')[0].split('?')[0];
  const slash = withoutQuery.lastIndexOf('/');
  const dir = slash >= origin.length ? withoutQuery.slice(0, slash + 1) : origin + '/';
  const combined = dir + loc;

  const m = combined.match(/^(https?:\/\/[^\/]+)(\/.*)?$/i);
  if (!m) return combined;
  const parts = String(m[2] || '/').split('/');
  const stack = [];
  parts.forEach(part => {
    if (!part || part === '.') return;
    if (part === '..') stack.pop();
    else stack.push(part);
  });
  return m[1] + '/' + stack.join('/');
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
    'Article Masterは現在、精密診断Package生成の必須条件ではありません。\n\n' +
    'DiagnosisはArticleIDが無い場合、記事URLから安定した内部IDを自動生成してDoctor診断まで進めます。\n' +
    'SBMの記事管理データがすでにある場合は、本物のArticleID・記事タイトル・メインクエリを優先利用できます。\n\n' +
    '通常運用では手動CSV取込は不要です。'
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

function sdsdSurrogateArticleId_(url) {
  return `REF-${sdsdShortHash_(sdsdNormalizeUrl_(url))}`;
}

function sdsdResolveArticleIdentity_(master, url) {
  const realId = master ? String(master.articleId || '').trim() : '';
  if (realId) {
    return {
      articleId: realId,
      source: 'SBM_ARTICLE_MASTER',
      surrogate: false
    };
  }
  return {
    articleId: sdsdSurrogateArticleId_(url),
    source: 'URL_SURROGATE',
    surrogate: true
  };
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

  let values = sh.getDataRange().getValues();
  if (values.length < 2) throw new Error('今回の診断対象に案件がありません。');

  let headers = values[0].map(String);
  const idx = {};
  headers.forEach((h,i) => idx[h] = i);

  const articleMap = sdsdBuildArticleMasterMap_();
  const queryBundle = sdsdBuildQueryEvidenceBundle_();
  const queryMap = queryBundle.map;
  const querySourceCount = queryBundle.sourceCount;

  const extraHeaders = [
    'ArticleID','Article Title','Main Query','Article Fetch Status',
    'Case Package Status','Article Cache Key','Query Evidence Count'
  ];

  const missingHeaders = extraHeaders.filter(h => idx[h] == null);
  if (missingHeaders.length) {
    const startCol = headers.length + 1;
    sh.getRange(1, startCol, 1, missingHeaders.length).setValues([missingHeaders]);
    headers = headers.concat(missingHeaders);
  }

  // Read once after schema normalization; all case updates below happen in memory.
  const data = sh.getDataRange().getValues();
  const refreshedHeaders = data[0].map(String);
  const col = {};
  refreshedHeaders.forEach((h,i) => col[h] = i);

  const firstTechnicalCol = refreshedHeaders.indexOf('Batch Order') + 1;
  if (firstTechnicalCol > 0) {
    sdsdHideTechnicalColumns_(sh, firstTechnicalCol, sh.getLastColumn());
  }

  const cache = CacheService.getDocumentCache();
  let alreadyReady = 0;
  let processedThisRun = 0;
  let newlyReady = 0;
  let failed = 0;

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

  const workRows = candidateRows.slice(0, maxPerRun);

  for (let wi=0; wi<workRows.length; wi++) {
    const r = workRows[wi];
    const row = data[r];
    const url = String(row[col['URL']] || '');
    if (!url) continue;

    const master = articleMap[sdsdNormalizeUrl_(url)] || null;
    const identityInfo = sdsdResolveArticleIdentity_(master, url);
    const articleId = identityInfo.articleId;
    const queryEvidence = (queryMap[sdsdNormalizeUrl_(url)] || []).slice(0,10);

    row[col['ArticleID']] = articleId;
    row[col['Article Title']] = master ? String(master.title || '') : '';
    row[col['Main Query']] = master ? String(master.mainQuery || '') : '';
    row[col['Query Evidence Count']] = queryEvidence.length;

    if (querySourceCount > 0 && queryEvidence.length === 0) {
      row[col['Case Package Status']] = 'QUERY_EVIDENCE_MISSING';
      row[col['Referral Status']] = 'NEEDS_CASE_ENRICHMENT_REVIEW';
      failed++;
      processedThisRun++;
      continue;
    }

    const fetched = sdsdFetchArticleEvidence_(url);
    const masterTitle = (master && master.title) ? String(master.title) : '';
    const liveTitle = String(fetched.title || '').trim();
    const title = liveTitle || masterTitle || sdsdTitleFromUrlFallback_(url);
    const mainQuery = (master && master.mainQuery)
      ? String(master.mainQuery)
      : String((queryEvidence[0] && queryEvidence[0].query) || '');

    row[col['Article Title']] = title || fetched.title;
    row[col['Main Query']] = mainQuery;
    row[col['Article Fetch Status']] = fetched.status;
    row[col['Query Evidence Count']] = queryEvidence.length;

    if (col['Top Queries'] != null) {
      row[col['Top Queries']] = queryEvidence.map(q => q.query).join(' / ');
    }

    if (fetched.status !== 'VALID') {
      row[col['Case Package Status']] = 'ARTICLE_FETCH_REVIEW';
      row[col['Referral Status']] = 'NEEDS_CASE_ENRICHMENT_REVIEW';
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
    referral.case_identity.article_id_source = identityInfo.source;
    referral.case_identity.article_id_is_surrogate = identityInfo.surrogate;
    referral.case_identity.canonical_article_url = url;
    referral.case_identity.url = url;
    referral.case_identity.request_id = String(
      referral.case_identity.request_id ||
      referral.request_id ||
      sdsdBuildRequestId_(batchId, articleId)
    );

    referral.case_id = referral.case_identity.individual_case_id;
    referral.request_id = referral.case_identity.request_id;
    referral.site_diagnosis_case_id = referral.case_identity.site_diagnosis_case_id;
    referral.site_diagnosis_batch_id = batchId;
    referral.site_id = siteId;
    referral.article_id = articleId;
    referral.article_id_source = identityInfo.source;
    referral.article_id_is_surrogate = identityInfo.surrogate;
    referral.article_url = url;

    const requestedCanonical = sdsdNormalizeUrl_(url);
    const observedCanonical = String(fetched.canonicalUrl || '').trim();
    const observedCanonicalNormalized = observedCanonical ? sdsdNormalizeUrl_(observedCanonical) : '';
    const titleMismatch = Boolean(masterTitle && liveTitle && masterTitle !== liveTitle);
    const canonicalMismatch = Boolean(
      observedCanonicalNormalized &&
      requestedCanonical &&
      observedCanonicalNormalized !== requestedCanonical
    );

    referral.data_quality_flags = Array.isArray(referral.data_quality_flags)
      ? referral.data_quality_flags
      : [];
    if (titleMismatch) {
      referral.data_quality_flags.push({
        field: 'article_evidence.title',
        issue: 'Article Master title differs from the live page title.',
        master_value: masterTitle,
        live_value: liveTitle,
        resolution: 'LIVE_PAGE_TITLE_USED'
      });
    }
    if (canonicalMismatch) {
      referral.data_quality_flags.push({
        field: 'article_evidence.observed_html_canonical_url',
        issue: 'HTML canonical differs from the requested/canonical article URL used by Diagnosis.',
        requested_url: url,
        observed_html_canonical_url: observedCanonical,
        resolution: 'REQUESTED_ARTICLE_URL_RETAINED_AS_CANONICAL_IDENTITY'
      });
    }

    referral.article_evidence = {
      status: 'VALID',
      title: title,
      title_source: liveTitle ? 'LIVE_PAGE' : (masterTitle ? 'ARTICLE_MASTER_FALLBACK' : 'URL_FALLBACK'),
      page_title: liveTitle,
      article_url: url,
      canonical_url: url,
      canonical_url_source: 'DIAGNOSIS_ARTICLE_URL',
      observed_html_canonical_url: observedCanonical,
      canonical_mismatch: canonicalMismatch,
      meta_description: fetched.metaDescription,
      main_query: mainQuery,
      identity_note: identityInfo.surrogate
        ? 'SBM ArticleID未取得。記事URLを正本としてDiagnosis内部IDを使用。SBM返却時はURL照合してください。'
        : 'SBM ArticleID確認済み。',
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

    row[col['Referral JSON']] = JSON.stringify(referral);
    row[col['Article Cache Key']] = cacheKey;
    row[col['Referral Status']] = 'READY_FOR_INDIVIDUAL_DOCTOR';
    row[col['Case Package Status']] = 'READY';

    newlyReady++;
    processedThisRun++;
  }

  // One sheet write for all case rows instead of many per-cell writes.
  if (data.length > 1 && refreshedHeaders.length > 0) {
    sh.getRange(2, 1, data.length - 1, refreshedHeaders.length).setValues(data.slice(1));
  }
  SpreadsheetApp.flush();

  let total = 0;
  let readyTotal = 0;
  let reviewTotal = 0;

  data.slice(1).forEach(r => {
    if (!String(r[col['URL']] || '')) return;
    total++;
    const rs = String(r[col['Referral Status']] || '');
    const ps = String(r[col['Case Package Status']] || '');

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
  let realArticleIdCount = 0;
  let surrogateArticleIdCount = 0;

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
    const isSurrogate = Boolean(identity.article_id_is_surrogate) || /^REF-/.test(articleId);
    if (isSurrogate) surrogateArticleIdCount++;
    else realArticleIdCount++;
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
      article_id_source: String(identity.article_id_source || (isSurrogate ? 'URL_SURROGATE' : 'SBM_ARTICLE_MASTER')),
      article_id_is_surrogate: isSurrogate,
      canonical_article_url: url,
      canonical_article_url_source: 'DIAGNOSIS_ARTICLE_URL',
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

  const created = sdsdCreateZipInOutputFolder_(blobs, zipName);
  const file = created.file;

  const result = {
    caseCount: rows.length,
    refetched: refetched,
    fileUrl: file.getUrl(),
    fileId: file.getId(),
    fileName: file.getName(),
    realArticleIdCount: realArticleIdCount,
    surrogateArticleIdCount: surrogateArticleIdCount
  };

  if (!options.silent) {
    sdsdShowZipExportComplete_(
      'Doctor Case Packageの生成が完了しました。',
      [
        `案件数: ${result.caseCount}件`,
        `本文再取得: ${result.refetched}件`,
        `SBM ArticleID使用: ${result.realArticleIdCount}件 / URL自動識別: ${result.surrogateArticleIdCount}件`
      ],
      file,
      created.folderInfo,
      'このZIPをSIMS Doctorへ渡してください。個別精密診断のDoctor回答はSIMS-Blog-Managerへ登録します。',
      {purpose:'優先記事を個別に詳しく診断し、Writer等の処置方針を決めるPackageです。',primaryLabel:'次の手順を確認',primaryFn:'sdsdShowIndividualDoctorResultWaiting_'}
    );
  }
  return result;
}

// ============================================================================
// Source: OutputFolderSettings.gs
// ============================================================================
const SDSD_OUTPUT_FOLDER_PROP = 'SDSD_OUTPUT_FOLDER_ID';

function sdsdOutputFolderInfo_() {
  const props = PropertiesService.getDocumentProperties();
  const id = String(props.getProperty(SDSD_OUTPUT_FOLDER_PROP) || '').trim();
  if (id) {
    try {
      const folder = DriveApp.getFolderById(id);
      return {id:folder.getId(),name:folder.getName(),url:folder.getUrl(),isDefault:false,folder:folder};
    } catch (e) {
      props.deleteProperty(SDSD_OUTPUT_FOLDER_PROP);
    }
  }
  const root = DriveApp.getRootFolder();
  return {id:root.getId(),name:'マイドライブ',url:'https://drive.google.com/drive/my-drive',isDefault:true,folder:root};
}

function sdsdShowOutputFolder() {
  const info = sdsdOutputFolderInfo_();
  SpreadsheetApp.getUi().alert(
    '現在のZIP保存先\n\n' +
    'フォルダー: ' + info.name + '\n' +
    (info.isDefault ? '状態: 未設定のためマイドライブ直下を使用\n' : '状態: 保存先設定済み\n') +
    '\n' + info.url
  );
}

function sdsdChooseOutputFolder() {
  const current = sdsdOutputFolderInfo_();
  const html = HtmlService.createHtmlOutput(sdsdOutputFolderPickerHtml_({
    folderId:current.id,
    currentFolderId:current.id,
    currentFolderName:current.name,
    currentFolderUrl:current.url,
    currentIsDefault:current.isDefault
  })).setWidth(760).setHeight(560);
  SpreadsheetApp.getUi().showModalDialog(html, 'ZIP保存先');
}

function sdsdListOutputFolderPickerFolder(folderId) {
  const folder = folderId ? DriveApp.getFolderById(String(folderId)) : DriveApp.getRootFolder();
  const folders = [];
  const it = folder.getFolders();
  while (it.hasNext()) { const f=it.next(); folders.push({id:f.getId(),name:f.getName()}); }
  folders.sort((a,b)=>String(a.name).localeCompare(String(b.name),'ja'));
  let parent=null;
  const parents=folder.getParents();
  if (parents.hasNext()) { const p=parents.next(); parent={id:p.getId(),name:p.getName()}; }
  return {id:folder.getId(),name:folder.getName()||'マイドライブ',parent:parent,folders:folders};
}

function sdsdSetOutputFolder(payload) {
  const folderId=String(payload&&payload.folderId||'').trim();
  if(!folderId) throw new Error('保存先フォルダーが選択されていません。');

  const folder=DriveApp.getFolderById(folderId);
  const savedId=folder.getId();
  const props=PropertiesService.getDocumentProperties();
  props.setProperty(SDSD_OUTPUT_FOLDER_PROP,savedId);

  // Read back the saved value so the UI never reports success unless persistence succeeded.
  const verifiedId=String(props.getProperty(SDSD_OUTPUT_FOLDER_PROP)||'').trim();
  if(verifiedId!==savedId){
    throw new Error('保存先設定の確認に失敗しました。もう一度設定してください。');
  }

  const verified=DriveApp.getFolderById(verifiedId);
  return {
    id:verified.getId(),
    name:verified.getName()||'マイドライブ',
    url:verified.getUrl(),
    verified:true
  };
}

function sdsdOutputFolderPickerHtml_(o) {
  const data=JSON.stringify(o||{}).replace(/</g,'\\u003c');
  return `<!doctype html><html><head><base target="_top"><style>
  body{font-family:Arial,"Noto Sans JP",sans-serif;margin:0;background:#f8fafd;color:#202124}.wrap{padding:18px 20px 20px}
  .hero{background:#185abc;color:#fff;padding:14px 18px;border-radius:10px}.hero h2{margin:0 0 5px;font-size:20px}.hero p{margin:0;font-size:13px}
  .saved,.browse,.success{border-radius:8px;padding:10px 12px;margin-top:10px;font-size:13px}
  .saved{background:#e8f0fe;border:1px solid #aecbfa}.browse{background:#fef7e0;border:1px solid #fdd663}
  .success{display:none;background:#e6f4ea;border:1px solid #81c995;color:#137333;font-weight:bold}
  .card{background:#fff;border:1px solid #dadce0;border-radius:10px;margin-top:10px;padding:12px}.bar{display:flex;gap:8px;align-items:center}.where{flex:1;font-weight:bold;color:#174ea6}
  button{border:1px solid #dadce0;background:#fff;border-radius:6px;padding:8px 12px;cursor:pointer}button.primary{background:#1a73e8;color:#fff;border-color:#1a73e8;font-weight:bold}button:disabled{opacity:.55;cursor:default}
  .list{height:175px;overflow:auto;border:1px solid #e0e0e0;border-radius:7px;margin-top:9px}.row{padding:9px 11px;border-bottom:1px solid #f1f3f4;cursor:pointer}.row:hover{background:#f8f9fa}
  .hint{color:#5f6368;font-size:12px;margin-top:7px}.err{color:#b3261e;margin-top:8px;white-space:pre-wrap}.actions{text-align:right;margin-top:12px}
  .url{font-size:11px;color:#5f6368;word-break:break-all;margin-top:3px}
  </style></head><body><div class="wrap">
  <div class="hero"><h2>ZIP保存先</h2><p>Evidence / Doctor Packageを保存するGoogle Driveフォルダーを選びます。</p></div>
  <div class="saved">現在設定済みの保存先：<b id="savedName"></b><div id="savedUrl" class="url"></div></div>
  <div class="browse">いま表示中のフォルダー：<b id="browseName">読み込み中...</b></div>
  <div id="success" class="success"></div>
  <div class="card"><div class="bar"><button id="up">↑ 上へ</button><div id="where" class="where"></div></div><div id="list" class="list"></div><div class="hint">📁 フォルダーをクリックして移動し、「いま表示中のフォルダー」を保存先に設定してください。</div></div>
  <div id="err" class="err"></div><div class="actions"><button id="close">閉じる</button> <button id="set" class="primary">このフォルダーを保存先にする</button></div>
  </div><script>
  const init=${data};let current=null;
  document.getElementById('savedName').textContent=init.currentFolderName||'マイドライブ';
  document.getElementById('savedUrl').textContent=init.currentFolderUrl||'';
  document.getElementById('close').onclick=()=>google.script.host.close();

  function fail(e){
    document.getElementById('err').textContent=(e&&e.message)||String(e||'不明なエラー');
    document.getElementById('success').style.display='none';
  }
  function load(id){
    document.getElementById('err').textContent='';
    document.getElementById('list').innerHTML='<div class="row" style="cursor:default">読み込み中...</div>';
    google.script.run.withSuccessHandler(render).withFailureHandler(fail).sdsdListOutputFolderPickerFolder(id);
  }
  function render(d){
    current=d;
    document.getElementById('where').textContent=d.name||'マイドライブ';
    document.getElementById('browseName').textContent=d.name||'マイドライブ';
    document.getElementById('up').disabled=!d.parent;
    const box=document.getElementById('list');box.innerHTML='';
    d.folders.forEach(f=>{
      const x=document.createElement('div');x.className='row';x.textContent='📁 '+f.name;
      x.onclick=()=>load(f.id);box.appendChild(x);
    });
    if(!d.folders.length)box.innerHTML='<div class="row" style="cursor:default;color:#5f6368">このフォルダー内にサブフォルダーはありません。</div>';
  }
  document.getElementById('up').onclick=()=>{if(current&&current.parent)load(current.parent.id);};
  document.getElementById('set').onclick=()=>{
    if(!current){fail('フォルダー情報の読み込みが完了していません。');return;}
    const b=document.getElementById('set');
    b.disabled=true;b.textContent='設定中...';
    document.getElementById('err').textContent='';
    document.getElementById('success').style.display='none';
    google.script.run
      .withSuccessHandler(r=>{
        if(!r||!r.verified){fail('保存先設定を確認できませんでした。');b.disabled=false;b.textContent='このフォルダーを保存先にする';return;}
        document.getElementById('savedName').textContent=r.name;
        document.getElementById('savedUrl').textContent=r.url||'';
        const s=document.getElementById('success');
        s.textContent='保存先を「'+r.name+'」に設定しました。';
        s.style.display='block';
        b.textContent='設定済み';
        b.disabled=true;
      })
      .withFailureHandler(e=>{
        b.disabled=false;b.textContent='このフォルダーを保存先にする';fail(e);
      })
      .sdsdSetOutputFolder({folderId:current.id});
  };
  // Start browsing from the folder that is actually configured, not from My Drive.
  load(init.folderId);
  </script></body></html>`;
}

function sdsdCreateZipInOutputFolder_(blobs, zipName) {
  const info=sdsdOutputFolderInfo_();
  const file=info.folder.createFile(Utilities.zip(blobs,zipName));
  return {file:file,folderInfo:info};
}

function sdsdSafeSiteToken_() {
  const session = sdsdGetCurrentSession_();
  let token = String(session.siteId || session.siteName || session.host || 'site').trim();
  token = token.normalize ? token.normalize('NFKC') : token;
  token = token.replace(/^https?:\/\//i,'').replace(/^www\./i,'').replace(/[^A-Za-z0-9._-]+/g,'-');
  token = token.replace(/^-+|-+$/g,'').replace(/-+/g,'-');
  return token || 'site';
}

function sdsdShowGuidedActionDialog_(o) {
  o=o||{};
  const esc=v=>String(v==null?'':v)
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
  const title=esc(o.title||'次の操作');
  const purpose=esc(o.purpose||'').replace(/\n/g,'<br>');
  const done=esc(o.done||'').replace(/\n/g,'<br>');
  const next=esc(o.next||'').replace(/\n/g,'<br>');
  const primaryLabel=esc(o.primaryLabel||'次へ');
  const primaryFn=String(o.primaryFn||'').replace(/[^A-Za-z0-9_]/g,'');
  const secondaryLabel=esc(o.secondaryLabel||'閉じる');
  const fileUrl=esc(o.fileUrl||'');
  const fileName=esc(o.fileName||'');
  const body=`<!doctype html><html><head><base target="_top"><style>
  body{font-family:Arial,"Noto Sans JP",sans-serif;margin:0;background:#f8fafd;color:#202124}.wrap{padding:20px}
  .hero{background:#185abc;color:#fff;padding:16px 18px;border-radius:10px}.hero h2{margin:0 0 5px;font-size:20px}.hero p{margin:0;line-height:1.6;font-size:13px}
  .card{background:#fff;border:1px solid #dadce0;border-radius:10px;margin-top:14px;padding:14px}.label{font-weight:bold;color:#174ea6;margin-bottom:5px}.text{line-height:1.7;font-size:13px}
  .file{margin-top:12px;padding:10px;background:#f1f3f4;border-radius:7px;word-break:break-all}.actions{text-align:right;margin-top:16px}
  button,a.btn{display:inline-block;border:1px solid #dadce0;background:#fff;border-radius:6px;padding:9px 14px;cursor:pointer;text-decoration:none;color:#202124;margin-left:6px}
  button.primary{background:#1a73e8;color:#fff;border-color:#1a73e8;font-weight:bold}.status{font-size:12px;color:#5f6368;margin-top:8px;text-align:right}
  </style></head><body><div class="wrap"><div class="hero"><h2>${title}</h2><p>${purpose}</p></div>
  ${done?`<div class="card"><div class="label">完了したこと</div><div class="text">${done}</div></div>`:''}
  ${next?`<div class="card"><div class="label">次にすること</div><div class="text">${next}</div></div>`:''}
  ${fileName?`<div class="file"><b>ファイル：</b>${fileName}${fileUrl?`<br><a href="${fileUrl}" target="_blank">Google Driveで開く</a>`:''}</div>`:''}
  <div class="actions"><button onclick="google.script.host.close()">${secondaryLabel}</button>${primaryFn?`<button class="primary" id="primary">${primaryLabel}</button>`:''}</div><div id="status" class="status"></div>
  </div><script>${primaryFn?`document.getElementById('primary').onclick=()=>{const b=document.getElementById('primary');b.disabled=true;document.getElementById('status').textContent='開いています...';google.script.run.withSuccessHandler(()=>google.script.host.close()).withFailureHandler(e=>{b.disabled=false;document.getElementById('status').textContent=(e&&e.message)?e.message:String(e||'エラー');}).${primaryFn}();};`:''}</script></body></html>`;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(body).setWidth(620).setHeight(520), String(o.windowTitle||o.title||'次の操作'));
}

function sdsdShowSiteWideDoctorResultDialog() {
  const html='<!doctype html><html><head><base target="_top"><style>'+ 
    'body{font-family:Arial,"Noto Sans JP",sans-serif;margin:0;background:#f8fafd;color:#202124}.wrap{padding:20px}.hero{background:#185abc;color:#fff;padding:16px 18px;border-radius:10px}.hero h2{margin:0 0 6px;font-size:20px}.hero p{margin:0;font-size:13px;line-height:1.6}.card{background:#fff;border:1px solid #dadce0;border-radius:10px;margin-top:14px;padding:14px}textarea{width:100%;height:260px;box-sizing:border-box;border:1px solid #bdc1c6;border-radius:7px;padding:10px;font-family:monospace}.actions{text-align:right;margin-top:12px}button{border:1px solid #dadce0;background:#fff;border-radius:6px;padding:9px 14px;cursor:pointer;margin-left:6px}button.primary{background:#1a73e8;color:#fff;border-color:#1a73e8;font-weight:bold}.status{margin-top:8px;color:#5f6368;font-size:12px}'+
    '</style></head><body><div class="wrap"><div class="hero"><h2>Doctor診断結果を取り込む</h2><p>Doctorの回答をDiagnosisへ戻し、治療計画またはCreator候補の次工程を決めます。</p></div><div class="card"><b>Doctor回答全文を貼り付けてください</b><p style="font-size:12px;color:#5f6368">横断診断・サイト横断精密診断のどちらでも、回答全文またはJSONをそのまま貼り付けられます。</p><textarea id="t" placeholder="Doctor回答全文またはJSONを貼り付け"></textarea><div class="status" id="s"></div></div><div class="actions"><button onclick="google.script.host.close()">キャンセル</button><button class="primary" id="go">診断結果を取り込む</button></div></div><script>document.getElementById("go").onclick=()=>{const t=document.getElementById("t").value.trim(),b=document.getElementById("go"),s=document.getElementById("s");if(!t){s.textContent="Doctor回答を貼り付けてください。";return;}b.disabled=true;s.textContent="取り込んでいます...";google.script.run.withSuccessHandler(()=>{s.textContent="取り込みを開始しました。";setTimeout(()=>google.script.host.close(),700);}).withFailureHandler(e=>{b.disabled=false;s.textContent=(e&&e.message)?e.message:String(e||"エラー");}).sdsdRegisterSiteWideDoctorResultText(t);};</script></body></html>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(720).setHeight(590),'Doctor診断結果を取り込む');
}

function sdsdRegisterSiteWideDoctorResultText(text) {
  const t=String(text||'').trim();
  if(!t) throw new Error('Doctor回答が空です。');
  const sh=sdsdEnsureSiteWideResultImportSheet_();
  if(sh.getLastRow()>=2)sh.getRange(2,1,Math.max(sh.getLastRow()-1,1),1).clearContent();
  sh.getRange('A2').setValue(t);
  return sdsdRegisterSiteWideDoctorResult();
}

function sdsdShowZipExportComplete_(title, detailLines, file, folderInfo, nextText, options) {
  const lines=Array.isArray(detailLines)?detailLines:[];
  const o=options||{};
  sdsdShowGuidedActionDialog_({
    title:String(title||'ZIP Packageを生成しました。'),
    purpose:String(o.purpose||'Doctorへ渡す診断Packageの準備が完了しました。'),
    done:(lines.length?lines.join('\n')+'\n':'')+'保存先: '+folderInfo.name,
    next:String(nextText||'このZIPをSIMS Doctorへ渡してください。'),
    fileName:file.getName(),
    fileUrl:file.getUrl(),
    primaryLabel:String(o.primaryLabel||'次の手順を確認'),
    primaryFn:String(o.primaryFn||''),
    secondaryLabel:'閉じる'
  });
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

  const completedIndividualUrls = sdsdGetCompletedIndividualUrls_();
  const eligible = rows.filter(r => {
    const normalizedUrl = sdsdNormalizeUrl_(String(r[idx['Normalized URL']] || ''));
    const p = String(r[idx['Priority Candidate']] || '');
    const guard = String(r[idx['Recent Treatment Guard']] || '');
    const ownership = String(r[idx['Ownership']] || '');
    return (
      !completedIndividualUrls[normalizedUrl] &&
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
  'SDSD_INDIVIDUAL_DOCTOR_COMPLETED_URLS',
  'SDSD_INDIVIDUAL_DOCTOR_PACKAGE_AT',
  'SDSD_INDIVIDUAL_DOCTOR_PACKAGE_CASE_COUNT',
  'SDSD_INDIVIDUAL_DOCTOR_PACKAGE_FILE_NAME',
  'SDSD_INDIVIDUAL_DOCTOR_PACKAGE_FILE_URL',
  'SDSD_INDIVIDUAL_DOCTOR_PACKAGE_BATCH_ID',
  'SDSD_INDIVIDUAL_DOCTOR_PACKAGE_STATUS',
  'SDSD_SITE_WIDE_PRECISION_PACKAGE_STATUS',
  'SDSD_SITE_WIDE_PRECISION_PACKAGE_FILE_ID',
  'SDSD_SITE_WIDE_PRECISION_PACKAGE_FILE_URL',
  'SDSD_SITE_WIDE_PRECISION_PACKAGE_FILE_NAME',
  'SDSD_SITE_WIDE_PRECISION_PACKAGE_CLUSTER_COUNT',
  'SDSD_SITE_WIDE_PRECISION_PACKAGE_AT',
  'SDSD_SBM_HANDOFF_STATUS',
  'SDSD_SBM_HANDOFF_AT',
  'SDSD_SBM_HANDOFF_FINGERPRINT',
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
  const selected = sdsdPendingSelectedCaseCount_();
  const opportunityCases = sdsdSheetDataRowCount_(SDSD_CONFIG.sheets.opportunityCases);
  const candidates = sdsdSheetDataRowCount_(SDSD_CONFIG.sheets.candidates);

  let actionableTreatment = 0;
  let additionalEvidence = 0;

  // When a stored Doctor site-wide result exists, use the same NEEDS_EVIDENCE
  // predicate that the precision-package exporter uses. This keeps Home,
  // guided routing, and package selection in sync.
  let stored = null;
  try { stored = sdsdReadStoredSiteWideResult_(); } catch (e) { stored = null; }

  if (stored && Array.isArray(stored.diagnosis_cases)) {
    additionalEvidence = sdsdNeedsEvidenceCaseCount_();
    stored.diagnosis_cases.forEach(c => {
      const route = String(c.route_to || '');
      if (/^(WRITER|MERGE|CREATOR)$/.test(route)) actionableTreatment++;
    });
  } else {
    // Compatibility fallback for legacy sessions that only have the treatment-plan sheet.
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
      'メニュー最上段の「▶ 次に進む（Diagnosisに任せる）」から診断を開始してください。'
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
    '別サイトを診断する場合は、先に「設定・管理 → 現在の診断を終了」を実行してください。'
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
    '別サイトを診断するときは、メニュー最上段の「▶ 次に進む（Diagnosisに任せる）」から開始できます。'
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
    '別のEvidenceを読み込む前に、メニュー「設定・管理 → 現在の診断を終了」を実行してください。\n' +
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
    next:'次はサイト全体の分析です。通常はそのまま「サイト診断を開始」を押してください。'
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
  document.getElementById('import').onclick=()=>{if(!selected)return;const b=document.getElementById('import');b.disabled=true;b.textContent='読み込み中...';google.script.run.withSuccessHandler(r=>{document.getElementById('meta').innerHTML='<b style="color:#137333">読み込み完了</b><br>サイト：'+esc(r.siteName)+'<br>'+esc(r.next);b.textContent='サイト診断を開始';b.disabled=false;b.onclick=()=>{b.disabled=true;b.textContent='分析中...';google.script.run.withSuccessHandler(()=>google.script.host.close()).withFailureHandler(e=>{b.disabled=false;b.textContent='サイト診断を開始';fail(e)}).sdsdRunProductDiagnosis();};}).withFailureHandler(e=>{b.disabled=false;b.textContent='このEvidenceを読み込む';fail(e)}).sdsdImportSelectedEvidence({fileId:selected.id});};
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
      values[i][idx['Recent Treatment Guard']] = 'WAIT';
      values[i][idx['Referral Status']] = 'BLOCKED_BY_FINAL_GUARD';
      blocked++;
    }
  }

  const userStatusIdx = headers.indexOf('状態');
  const refIdx = headers.indexOf('Referral Status');
  if (userStatusIdx >= 0 && refIdx >= 0) {
    for (let i=1; i<values.length; i++) {
      const ref = String(values[i][refIdx] || '');
      values[i][userStatusIdx] = ref.indexOf('BLOCK') >= 0 ? '保留' : 'Doctor診断待ち';
    }
  }

  // Persist guard and user-facing status changes in one batch.
  sh.getRange(2, 1, values.length - 1, headers.length).setValues(values.slice(1));

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

function sdsdBuildQueryEvidenceBundle_() {
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

  return {map: map, sourceCount: rows.length};
}

function sdsdBuildQueryEvidenceMap_() {
  return sdsdBuildQueryEvidenceBundle_().map;
}

function sdsdQueryEvidenceSourceCount_() {
  return sdsdBuildQueryEvidenceBundle_().sourceCount;
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
      `SIMS-Doctor-Site-Wide-Diagnosis-${sdsdSafeSiteToken_()}-${Utilities.formatDate(
        new Date(),
        Session.getScriptTimeZone() || 'Asia/Tokyo',
        'yyyyMMdd-HHmmss'
      )}.zip`;

    const created = sdsdCreateZipInOutputFolder_(blobs, zipName);
    const file = created.file;

    sdsdShowZipExportComplete_(
      'サイト横断Doctor Packageを生成しました。',
      [`案件数: ${rows.length}件`],
      file,
      created.folderInfo,
      'このZIPをSIMS Doctorへ渡してください。Doctorの回答が返ったら「Doctor診断結果を取り込む」へ進みます。',
      {purpose:'サイト横断で見つかったカニバリ疑い・新記事機会・コンテンツギャップをDoctorに整理してもらうPackageです。',primaryLabel:'Doctor診断結果を取り込む',primaryFn:'sdsdShowSiteWideDoctorResultDialog'}
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
    throw new Error('Doctor診断結果が登録されていません。▶ 次に進むから「Doctor診断結果を貼り付ける」を開いてください。');
  }
  const raw = String(sh.getRange('A1').getValue() || '');
  const obj = JSON.parse(raw);
  if (!obj || obj.format !== 'SIMS_DOCTOR_SITE_WIDE_RESULT_V1') {
    throw new Error('保存済みDoctor横断診断結果の形式が不正です。');
  }
  return obj;
}

function sdsdResolvedCreatorValidationSourceIds_(obj) {
  const resolved = {};
  const cases = obj && Array.isArray(obj.diagnosis_cases) ? obj.diagnosis_cases : [];
  cases.forEach(c => {
    const route = String(c && c.route_to || '').toUpperCase();
    const type = String(c && c.diagnosis_type || '').toUpperCase();
    const decision = String(c && c.doctor_decision || '').toUpperCase();
    const isResolvedCreator = route === 'CREATOR' && (
      type === 'LONGTAIL_DISCOVERY_CREATOR' ||
      decision === 'CREATOR_APPROVED_AFTER_LONGTAIL_DISCOVERY'
    );
    if (!isResolvedCreator) return;
    (Array.isArray(c.absorbed_source_case_ids) ? c.absorbed_source_case_ids : [])
      .map(String).filter(Boolean).forEach(id => { resolved[id] = true; });
  });
  return resolved;
}

function sdsdNeedsEvidenceCases_() {
  const obj = sdsdReadStoredSiteWideResult_();
  const resolvedCreatorSources = sdsdResolvedCreatorValidationSourceIds_(obj);
  return (obj.diagnosis_cases || []).filter(c => {
    if (String(c.route_to || '') !== 'NEEDS_EVIDENCE') return false;
    const id = String(c.diagnosis_case_id || '');
    if (id && resolvedCreatorSources[id]) return false;
    if (String(c.creator_validation_status || '').toUpperCase() === 'RESOLVED_TO_LONGTAIL_CREATOR') return false;
    return true;
  });
}

function sdsdEvidencePurpose_(c) {
  c = c || {};
  const eventual = String(c.eventual_route || '').toUpperCase();
  const decision = String(c.doctor_decision || '').toUpperCase();
  const type = String(c.diagnosis_type || '').toUpperCase();

  if (eventual === 'CREATOR' ||
      decision.indexOf('CREATOR_PENDING') >= 0 ||
      type.indexOf('NEW_ARTICLE') >= 0) {
    return 'CREATOR_VALIDATION';
  }
  return 'PRECISION_DIAGNOSIS';
}

function sdsdNeedsEvidenceCaseCount_() {
  return sdsdNeedsEvidenceCases_().length;
}

function sdsdCreatorValidationCases_() {
  return sdsdNeedsEvidenceCases_().filter(c =>
    sdsdEvidencePurpose_(c) === 'CREATOR_VALIDATION'
  );
}

function sdsdCreatorValidationCaseCount_() {
  return sdsdCreatorValidationCases_().length;
}

function sdsdPrecisionClusterCases_() {
  return sdsdNeedsEvidenceCases_()
    .filter(c => sdsdEvidencePurpose_(c) === 'PRECISION_DIAGNOSIS')
    .slice(0, 5);
}

function sdsdCreatorCandidateQuery_(c) {
  const theme = String(c && c.diagnosis_theme || '').trim();
  if (theme) return theme;

  const sourceMap = sdsdSiteWideSourceCaseMap_();
  const ids = (c && c.absorbed_source_case_ids) || [];
  for (let i=0; i<ids.length; i++) {
    const src = sourceMap[String(ids[i])] || null;
    if (!src) continue;
    const q = String(src.main_queries || src.diagnosis_theme || '').trim();
    if (q) return q.split(/[\n,、]/)[0].trim();
  }
  return '';
}

function sdsdCreatorValidationContext_() {
  const articleMap = sdsdBuildArticleMasterMap_();
  const queryRows = sdsdReadObjects_(SDSD_CONFIG.sheets.evidencePageQuery).map(r => ({
    url: sdsdNormalizeUrl_(sdsdObjectValue_(r, ['page','key','url','URL','記事URL'])),
    query: String(sdsdObjectValue_(r, ['query','クエリ','検索クエリ']) || '').trim(),
    clicks: Number(sdsdObjectValue_(r, ['clicks','クリック数']) || 0),
    impressions: Number(sdsdObjectValue_(r, ['impressions','表示回数']) || 0),
    position: Number(sdsdObjectValue_(r, ['position','掲載順位','平均掲載順位']) || 0)
  })).filter(x => x.url && x.query);
  return {articleMap: articleMap, queryRows: queryRows};
}


function sdsdCreatorQueryClusterProfile_(query, ctx) {
  const rows = (ctx && ctx.queryRows) || [];
  const members = [];
  rows.forEach(r => {
    const sim = sdsdQuerySimilarity_(r.query, query);
    if (sim < 0.30) return;
    members.push({
      query: r.query,
      article_url: r.url,
      similarity: Math.round(sim * 100) / 100,
      impressions: Number(r.impressions || 0),
      clicks: Number(r.clicks || 0),
      position: Number(r.position || 0)
    });
  });
  members.sort((a,b) => b.similarity - a.similarity || b.impressions - a.impressions);

  const totalImp = members.reduce((n,x) => n + Number(x.impressions || 0), 0);
  const urls = {};
  members.forEach(x => { if (x.article_url) urls[x.article_url] = true; });
  const base = sdsdParentThemeKey_(query) || query;

  return {
    cluster_label: base,
    cluster_queries: members.slice(0, 20),
    cluster_impressions: totalImp,
    cluster_article_count: Object.keys(urls).length
  };
}

function sdsdCreatorRoleDesign_(query, closest, profile) {
  const candidateTokens = sdsdQueryTokens_(query);
  const existingText = (closest || []).map(x => `${x.article_title || ''} ${x.main_query || ''}`).join(' ');
  const existingNorm = sdsdNormalizeQuery_(existingText);
  const differentiating = candidateTokens.filter(t => existingNorm.indexOf(t) < 0);
  const focus = differentiating.length ? differentiating.join(' / ') : query;
  const avoid = (closest || []).slice(0,3).map(x => x.main_query || x.article_title).filter(Boolean);

  return {
    role_summary: `新記事は「${query}」を主目的とし、特に「${focus}」の検索意図へ答える。既存記事の一般論を広げすぎず、独立したロングテール入口として設計する。`,
    do_not_target: avoid,
    differentiating_terms: differentiating,
    cluster_label: profile.cluster_label
  };
}

function sdsdCreatorInternalLinkCandidates_(closest, profile, ctx) {
  const out = [];
  const seen = {};
  function add(url, reason) {
    url = sdsdNormalizeUrl_(url);
    if (!url || seen[url]) return;
    seen[url] = true;
    const a = (ctx.articleMap || {})[url] || {};
    out.push({
      article_url: url,
      article_id: String(a.articleId || ''),
      article_title: String(a.title || ''),
      reason: reason
    });
  }
  (closest || []).forEach(x => add(x.article_url, '検索意図が近い既存記事'));
  (profile.cluster_queries || []).forEach(x => add(x.article_url, `同一キーワードクラスター: ${x.query}`));
  return out.slice(0, 8);
}

function sdsdCreatorCandidateValidation_(c, ctx) {
  const query = sdsdCreatorCandidateQuery_(c);
  const nq = sdsdNormalizeQuery_(query);
  if (!nq) {
    return {
      diagnosis_case_id: String(c.diagnosis_case_id || ''),
      diagnosis_theme: String(c.diagnosis_theme || ''),
      candidate_query: '',
      grade: 'YELLOW',
      gate: 'PROCEED_WITH_30D_MONITOR',
      reason: '候補キーワードを特定できないため、公開前に検索意図を手動確認してください。',
      differentiating_terms: [],
      closest_articles: [],
      related_queries: [],
      serp_check: 'RECOMMENDED'
    };
  }

  const qTokens = sdsdQueryTokens_(query);
  const articles = [];
  Object.keys(ctx.articleMap || {}).forEach(url => {
    const a = ctx.articleMap[url] || {};
    const title = String(a.title || '');
    const mainQuery = String(a.mainQuery || '');
    const titleCoverage = sdsdTitleQueryCoverage_(title, query);
    const mainQuerySimilarity = mainQuery ? sdsdQuerySimilarity_(mainQuery, query) : 0;
    const titleSimilarity = title ? sdsdQuerySimilarity_(title, query) : 0;
    const intentScore = Math.max(titleCoverage, mainQuerySimilarity, titleSimilarity);
    if (intentScore >= 0.35) {
      articles.push({
        article_url: url,
        article_id: String(a.articleId || ''),
        article_title: title,
        main_query: mainQuery,
        intent_score: Math.round(intentScore * 100) / 100
      });
    }
  });
  articles.sort((a,b) => b.intent_score - a.intent_score);

  const related = [];
  (ctx.queryRows || []).forEach(r => {
    const sim = sdsdQuerySimilarity_(r.query, query);
    if (sim < 0.45) return;
    related.push({
      query: r.query,
      article_url: r.url,
      similarity: Math.round(sim * 100) / 100,
      impressions: r.impressions,
      clicks: r.clicks,
      position: r.position
    });
  });
  related.sort((a,b) =>
    b.similarity - a.similarity || b.impressions - a.impressions
  );

  const closest = articles.slice(0, 5);
  const best = closest.length ? closest[0].intent_score : 0;

  // v0.6.3: GSCで既存記事が候補KWをすでに担当している事実を、
  // タイトル推定より強いEvidenceとしてCreatorゲートへ反映する。
  const exactQueryRows = related.filter(x =>
    sdsdNormalizeQuery_(x.query) === nq && !!x.article_url
  );
  const strongNearRows = related.filter(x =>
    x.similarity >= 0.85 && !!x.article_url && sdsdNormalizeQuery_(x.query) !== nq
  );
  const exactUrls = Array.from(new Set(exactQueryRows.map(x => x.article_url)));
  const strongNearUrls = Array.from(new Set(strongNearRows.map(x => x.article_url)));

  let grade = 'GREEN';
  let gate = 'PROCEED';
  let reason = '候補キーワードおよび主要な近似キーワードを担当する既存記事が確認されず、独立したロングテール記事として試す余地があります。';

  if (exactQueryRows.length) {
    grade = 'RED';
    gate = 'BLOCK_CREATOR_REVIEW_EXISTING';
    const top = exactQueryRows[0];
    reason = `既存記事 ${top.article_url} が候補キーワード「${query}」で既に表示されています（表示回数 ${Number(top.impressions||0)}）。新記事は作らず、既存記事の強化・役割整理を先に確認してください。`;
  } else if (strongNearRows.length || best >= 0.85) {
    grade = 'YELLOW';
    gate = 'SERP_REVIEW_BEFORE_CREATOR';
    const top = strongNearRows[0] || null;
    reason = top
      ? `候補キーワードと強く近い既存クエリ「${top.query}」を ${top.article_url} が取得しています（類似度 ${top.similarity}、表示回数 ${Number(top.impressions||0)}）。追加語による独立意図があるかDoctorのSERP確認を行ってください。`
      : '既存記事に非常に近い検索意図が確認されます。独立ロングテールとして分離できるかDoctorのSERP確認を行ってください。';
  } else {
    // v0.6.5: GSCに一致/近似クエリが出ていなくても、Doctor一次診断の
    // target_articles に既存URLがある案件は「既存記事なし」とみなさない。
    // 少なくともYELLOWへ落とし、SERPで独立性を確認する。
    const targetUrls = Array.from(new Set(((c && c.target_articles) || [])
      .map(x => sdsdNormalizeUrl_(x && (x.article_url || x.url || '')))
      .filter(Boolean)));
    if (targetUrls.length) {
      grade = 'YELLOW';
      gate = 'SERP_REVIEW_BEFORE_CREATOR';
      reason = `同一テーマに紐づく既存記事 ${targetUrls.slice(0,3).join(' / ')} が確認されています。GSCの完全一致がなくても既存記事との役割重複を否定できないため、DoctorのSERP確認後にCreator可否を決めてください。`;
    } else if (best >= 0.55 || related.some(x => x.similarity >= 0.67)) {
      grade = 'YELLOW';
      gate = 'SERP_REVIEW_BEFORE_CREATOR';
      reason = '関連する既存記事・近接クエリがあります。独立した検索意図として分離できる可能性は残るため、DoctorのSERP確認後にCreator可否を決めてください。';
    }
  }

  if (grade === 'GREEN' && (best >= 0.55 || related.some(x => x.similarity >= 0.67))) {
    grade = 'YELLOW';
    gate = 'SERP_REVIEW_BEFORE_CREATOR';
    reason = '関連する既存記事・近接クエリがあります。独立した検索意図として分離できる可能性は残るため、DoctorのSERP確認後にCreator可否を決めてください。';
  }

  const baseText = closest.map(x => `${x.article_title} ${x.main_query}`).join(' ');
  const nb = sdsdNormalizeQuery_(baseText);
  const differentiating = qTokens.filter(t => nb.indexOf(t) < 0);
  const clusterProfile = sdsdCreatorQueryClusterProfile_(query, ctx);
  const roleDesign = sdsdCreatorRoleDesign_(query, closest, clusterProfile);
  const internalLinks = sdsdCreatorInternalLinkCandidates_(closest, clusterProfile, ctx);

  return {
    diagnosis_case_id: String(c.diagnosis_case_id || ''),
    diagnosis_theme: String(c.diagnosis_theme || ''),
    candidate_query: query,
    grade: grade,
    gate: gate,
    reason: reason,
    differentiating_terms: roleDesign.differentiating_terms.length ? roleDesign.differentiating_terms : differentiating,
    keyword_cluster: clusterProfile.cluster_label,
    cluster_queries: clusterProfile.cluster_queries,
    cluster_impressions: clusterProfile.cluster_impressions,
    role_summary: roleDesign.role_summary,
    do_not_target: roleDesign.do_not_target,
    internal_link_candidates: internalLinks,
    closest_articles: closest,
    related_queries: related.slice(0, 12),
    serp_check: grade === 'GREEN' ? 'OPTIONAL' : (grade === 'YELLOW' ? 'REQUIRED_BEFORE_CREATOR' : 'NOT_APPLICABLE'),
    monitor_days: grade === 'RED' ? 0 : 30
  };
}

function sdsdCreatorCandidateValidations_() {
  const ctx = sdsdCreatorValidationContext_();
  return sdsdCreatorValidationCases_().map(c =>
    sdsdCreatorCandidateValidation_(c, ctx)
  );
}

function sdsdWriteCreatorValidationSheet_(results) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SDSD_CONFIG.sheets.creatorValidation);
  if (!sh) sh = ss.insertSheet(SDSD_CONFIG.sheets.creatorValidation);
  sh.clear();

  const headers = [
    '案件ID','新記事候補','メインキーワード','判定','Creator可否',
    '判定理由','検索クラスター','差別化語','役割分担','狙わない検索意図','内部リンク候補','近い既存記事','関連GSCクエリ','SERP確認','公開後モニター'
  ];
  // Creator作業用の通常表示は、前進可能なGREEN/YELLOWに限定する。
  // REDは集計・再判定には保持するが、Creator候補一覧へは表示しない。
  const visibleResults = (results || []).filter(x => x.grade !== 'RED');
  const rows = visibleResults.map(x => [
    x.diagnosis_case_id,
    x.diagnosis_theme,
    x.candidate_query,
    x.grade,
    x.gate,
    x.reason,
    x.keyword_cluster || '',
    (x.differentiating_terms || []).join(' / '),
    x.role_summary || '',
    (x.do_not_target || []).join(' / '),
    (x.internal_link_candidates || []).map(a => `${a.article_title || '（タイトル未取得）'} / ${a.article_url} / ${a.reason || ''}`).join('\n'),
    (x.closest_articles || []).map(a =>
      `${a.article_title || '（タイトル未取得）'} / ${a.article_url} / intent=${a.intent_score}`
    ).join('\n'),
    (x.related_queries || []).map(q =>
      `${q.query} / ${q.article_url} / sim=${q.similarity} / imp=${q.impressions}`
    ).join('\n'),
    x.serp_check,
    x.monitor_days ? `${x.monitor_days}日` : ''
  ]);

  sh.getRange(1,1,1,headers.length).setValues([headers]);
  if (rows.length) sh.getRange(2,1,rows.length,headers.length).setValues(rows);
  sh.setFrozenRows(1);
  sh.getRange(1,1,1,headers.length).setFontWeight('bold');
  sh.getRange(1,1,Math.max(rows.length+1,1),headers.length).setWrap(true);
  [190,260,240,90,220,420,220,220,480,320,520,520,520,220,150].forEach((w,i) => sh.setColumnWidth(i+1,w));
  ss.setActiveSheet(sh);
  return rows.length;
}

function sdsdRunCreatorCandidateValidation() {
  const ui = SpreadsheetApp.getUi();
  try {
    const results = sdsdCreatorCandidateValidations_();
    if (!results.length) {
      ui.alert('現在、新記事作成前チェックの対象案件はありません。');
      return [];
    }
    sdsdWriteCreatorValidationSheet_(results);
    const g = results.filter(x => x.grade === 'GREEN').length;
    const y = results.filter(x => x.grade === 'YELLOW').length;
    const r = results.filter(x => x.grade === 'RED').length;
    ui.alert(
      'Creator候補の作成前チェックが完了しました。\n\n' +
      `GREEN（独立性高・作成候補）: ${g}件\n` +
      `YELLOW（SERP確認後に判断）: ${y}件\n` +
      `RED（既存担当記事あり・Creator除外）: ${r}件\n\n` +
      '「Creator候補チェック」シートにはGREEN/YELLOWのみ表示します。\n' +
      'YELLOWはDoctorのSERP確認後にCreator可否を確定してください。'
    );
    return results;
  } catch (e) {
    ui.alert('Creator候補の作成前チェックに失敗しました。\n\n' + String(e && e.message ? e.message : e));
    throw e;
  }
}


function sdsdSelectedCreatorValidation_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getActiveSheet();
  if (!sh || sh.getName() !== SDSD_CONFIG.sheets.creatorValidation) {
    throw new Error('「Creator候補チェック」シートで対象行を選択してください。');
  }
  const row = sh.getActiveRange().getRow();
  if (row < 2) throw new Error('見出しではなくCreator候補の行を選択してください。');
  const caseId = String(sh.getRange(row,1).getValue() || '').trim();
  if (!caseId) throw new Error('選択行の案件IDを取得できません。');
  const c = sdsdCreatorValidationCases_().find(x => String(x.diagnosis_case_id || '') === caseId);
  if (!c) throw new Error('選択したCreator候補を保存済みDoctor結果から確認できません。');
  const ctx = sdsdCreatorValidationContext_();
  return {caseData:c, validation:sdsdCreatorCandidateValidation_(c, ctx)};
}

function sdsdBuildCreatorSerpReferralObject_(selected) {
  const v = selected.validation || {};
  const siteResult = sdsdReadStoredSiteWideResult_();
  return {
    format: 'SIMS_DOCTOR_CREATOR_SERP_REFERRAL_V1',
    contract_version: '1.0',
    site_diagnosis_batch_id: String(siteResult.site_diagnosis_batch_id || ''),
    site: siteResult.site || {},
    diagnosis_case_id: String(v.diagnosis_case_id || ''),
    candidate_keyword: String(v.candidate_query || ''),
    diagnosis_theme: String(v.diagnosis_theme || ''),
    local_gate: String(v.grade || ''),
    local_reason: String(v.reason || ''),
    keyword_cluster: String(v.keyword_cluster || ''),
    cluster_queries: v.cluster_queries || [],
    proposed_role: String(v.role_summary || ''),
    do_not_target: v.do_not_target || [],
    closest_existing_articles: v.closest_articles || [],
    internal_link_candidates: v.internal_link_candidates || [],
    monitor_days_after_publish: Number(v.monitor_days || 30),
    doctor_task: {
      purpose: '新記事作成前のSERP独立性と競合価値を確認する',
      checks: [
        '候補キーワードを実際に検索し、上位ページの主な検索意図を確認する',
        '元の広いテーマ/近接キーワードとSERPの顔ぶれ・回答内容が分離しているか確認する',
        '既存記事で十分に回答できる意図か、新記事として独立させる価値があるか判断する',
        '新記事化する場合、既存記事と競合しない役割分担と狙わない意図を確定する',
        '内部リンク候補が妥当か確認する'
      ],
      risk_policy: '明確な同一意図・重大カニバリだけをBLOCKする。ロングテールで独立性が見込める場合は多少の不確実性を許容し、30日モニター前提でCREATORを許可する。'
    },
    required_response_format: 'SIMS_DOCTOR_CREATOR_SERP_RESULT_V1'
  };
}

function sdsdBuildCreatorSerpReferralMarkdown_(o) {
  return [
    '# SIMS Doctor Creator SERP確認紹介状',
    '',
    `案件ID: ${o.diagnosis_case_id}`,
    `候補キーワード: ${o.candidate_keyword}`,
    `キーワードクラスター: ${o.keyword_cluster}`,
    `Diagnosis暫定判定: ${o.local_gate}`,
    '',
    '## 新記事の予定役割',
    o.proposed_role || '未設定',
    '',
    '## 狙わない検索意図',
    ...(o.do_not_target || []).map(x => `- ${x}`),
    '',
    '## 近い既存記事',
    ...(o.closest_existing_articles || []).map(x => `- ${x.article_title || ''} | ${x.main_query || ''} | ${x.article_url || ''}`),
    '',
    '## 内部リンク候補',
    ...(o.internal_link_candidates || []).map(x => `- ${x.article_title || ''} | ${x.article_url || ''} | ${x.reason || ''}`),
    '',
    '## Doctorへの依頼',
    '候補キーワードの現在SERPを確認し、競合ページの検索意図・構成を比較してください。',
    '明確な同一意図の既存記事がある場合だけ新記事を止め、独立したロングテール需要がある場合はCREATORを優先してください。',
    '多少の不確実性は許容し、公開後30日をSBMで観察する前提です。',
    '',
    '## 返却JSON',
    '```json',
    JSON.stringify({
      format:'SIMS_DOCTOR_CREATOR_SERP_RESULT_V1',
      diagnosis_case_id:o.diagnosis_case_id,
      candidate_keyword:o.candidate_keyword,
      decision:'CREATOR|WRITER|BLOCK|NEEDS_EVIDENCE',
      serp_independence:'HIGH|MEDIUM|LOW',
      search_intent:'',
      competitor_summary:'',
      role_with_existing_articles:'',
      do_not_target:[],
      internal_link_recommendations:[],
      new_article_reason:'',
      monitor_days:30,
      confidence:0
    }, null, 2),
    '```'
  ].join('\n');
}

function sdsdWriteCreatorSerpReferralSheet_(o) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SDSD_CONFIG.sheets.creatorSerpReferral);
  if (!sh) sh = ss.insertSheet(SDSD_CONFIG.sheets.creatorSerpReferral);
  sh.clear();
  const md = sdsdBuildCreatorSerpReferralMarkdown_(o);
  const rows = [
    ['Creator SERP確認紹介状',''],
    ['案件ID',o.diagnosis_case_id],
    ['候補キーワード',o.candidate_keyword],
    ['キーワードクラスター',o.keyword_cluster],
    ['Diagnosis暫定判定',o.local_gate],
    ['Doctorへ渡す紹介状（Markdown）',md],
    ['Doctorへ渡す紹介状（JSON）',JSON.stringify(o,null,2)]
  ];
  sh.getRange(1,1,rows.length,2).setValues(rows);
  sh.getRange(1,1,rows.length,2).setWrap(true).setVerticalAlignment('top');
  sh.getRange(1,1,1,2).setFontWeight('bold');
  sh.setColumnWidth(1,260); sh.setColumnWidth(2,900);
  ss.setActiveSheet(sh);
}

function sdsdCreateCreatorSerpReferralFromSelected() {
  const ui = SpreadsheetApp.getUi();
  try {
    const selected = sdsdSelectedCreatorValidation_();
    const referral = sdsdBuildCreatorSerpReferralObject_(selected);
    sdsdWriteCreatorSerpReferralSheet_(referral);
    sdsdSetCreatorSerpState_('WAITING_DOCTOR_RESULT', referral.diagnosis_case_id);
    ui.alert(
      'Creator候補のSERP確認紹介状を作成しました。\n\n' +
      `候補キーワード: ${referral.candidate_keyword}\n\n` +
      '目的：実際の検索結果でライバルと検索意図を確認し、新記事として独立させる価値を最終確認します。\n' +
      '次の操作：「Creator SERP確認」シートのMarkdownまたはJSONをSIMS Doctorへ渡してください。\n' +
      'Doctor回答が返ったら、▶ 次に進む、または「確認する → Creator SERP Doctor回答を取り込む」から回答全文を貼り付けてください。'
    );
    return referral;
  } catch(e) {
    ui.alert('Creator SERP確認紹介状を作成できませんでした。\n\n' + String(e && e.message ? e.message : e));
    throw e;
  }
}


function sdsdSetCreatorSerpState_(status, caseId) {
  const props = PropertiesService.getDocumentProperties();
  props.setProperty('SDSD_CREATOR_SERP_STATUS', String(status || ''));
  props.setProperty('SDSD_CREATOR_SERP_CASE_ID', String(caseId || ''));
  props.setProperty('SDSD_CREATOR_SERP_AT', new Date().toISOString());
}

function sdsdGetCreatorSerpState_() {
  const props = PropertiesService.getDocumentProperties();
  return {
    status: String(props.getProperty('SDSD_CREATOR_SERP_STATUS') || ''),
    caseId: String(props.getProperty('SDSD_CREATOR_SERP_CASE_ID') || '')
  };
}

function sdsdClearCreatorSerpState_() {
  const props = PropertiesService.getDocumentProperties();
  ['SDSD_CREATOR_SERP_STATUS','SDSD_CREATOR_SERP_CASE_ID','SDSD_CREATOR_SERP_AT']
    .forEach(k => props.deleteProperty(k));
}

function sdsdValidateCreatorSerpResult_(o) {
  if (!o || o.format !== 'SIMS_DOCTOR_CREATOR_SERP_RESULT_V1') {
    throw new Error('Doctor結果のformatがSIMS_DOCTOR_CREATOR_SERP_RESULT_V1ではありません。');
  }
  if (!String(o.diagnosis_case_id || '').trim()) {
    throw new Error('Doctor結果にdiagnosis_case_idがありません。');
  }
  const decision = String(o.decision || '').toUpperCase();
  if (!['CREATOR','WRITER','BLOCK','NEEDS_EVIDENCE'].includes(decision)) {
    throw new Error('decisionはCREATOR / WRITER / BLOCK / NEEDS_EVIDENCEのいずれかで返してください。');
  }
  return true;
}

function sdsdCreatorValidationForCaseId_(caseId) {
  const c = sdsdCreatorValidationCases_().find(x => String(x.diagnosis_case_id || '') === String(caseId || ''));
  if (!c) return null;
  return sdsdCreatorCandidateValidation_(c, sdsdCreatorValidationContext_());
}

function sdsdApplyCreatorSerpResult_(doctorResult) {
  sdsdValidateCreatorSerpResult_(doctorResult);
  const stored = sdsdReadStoredSiteWideResult_();
  const cases = Array.isArray(stored.diagnosis_cases) ? stored.diagnosis_cases : [];
  const idx = cases.findIndex(c => String(c.diagnosis_case_id || '') === String(doctorResult.diagnosis_case_id || ''));
  if (idx < 0) throw new Error('Doctor結果の案件IDが保存済みサイト治療計画にありません。');

  const c = Object.assign({}, cases[idx]);
  if (sdsdEvidencePurpose_(c) !== 'CREATOR_VALIDATION' && String(c.eventual_route || '').toUpperCase() !== 'CREATOR') {
    throw new Error('選択案件はCreator候補ではありません。');
  }

  const local = sdsdCreatorValidationForCaseId_(c.diagnosis_case_id) || {};
  const decision = String(doctorResult.decision || '').toUpperCase();
  const monitorDays = Math.max(14, Number(doctorResult.monitor_days || local.monitor_days || 30) || 30);
  const closest = Array.isArray(local.closest_articles) ? local.closest_articles : [];
  const doctorLinks = Array.isArray(doctorResult.internal_link_recommendations)
    ? doctorResult.internal_link_recommendations : [];

  c.confidence = String(doctorResult.confidence || c.confidence || '');
  c.additional_evidence_needed = [];
  c.internal_link_recommendations = doctorLinks.length ? doctorLinks : (local.internal_link_candidates || []);

  if (decision === 'CREATOR') {
    c.route_to = 'CREATOR';
    c.eventual_route = 'CREATOR';
    c.doctor_decision = 'CREATOR_APPROVED_AFTER_SERP_CHECK';
    c.treatment_strategy = 'CREATE_NEW_ARTICLE_AND_MONITOR';
    c.reason = String(doctorResult.new_article_reason || doctorResult.role_with_existing_articles || c.reason || '');
    const creatorKeyword = String(doctorResult.candidate_keyword || local.candidate_query || c.diagnosis_theme || '');
    const creatorReferenceArticles = (Array.isArray(c.target_articles) ? c.target_articles : []).map(a => {
      if (typeof a === 'string') return {article_url:String(a), role:'REFERENCE_ONLY'};
      return Object.assign({}, a || {}, {role:'REFERENCE_ONLY'});
    });
    c.target_articles = creatorReferenceArticles; // legacy SBM identity compatibility only
    c.reference_articles = creatorReferenceArticles;
    c.new_article_target = {
      target_type:'NEW_ARTICLE',
      candidate_keyword:creatorKeyword,
      article_id:'',
      article_url:'',
      publication_status:'NOT_PUBLISHED'
    };
    c.creator_plan = {
      candidate_keyword: creatorKeyword,
      keyword_cluster: String(local.keyword_cluster || ''),
      cluster_queries: local.cluster_queries || [],
      search_intent: String(doctorResult.search_intent || ''),
      serp_independence: String(doctorResult.serp_independence || ''),
      competitor_summary: String(doctorResult.competitor_summary || ''),
      new_article_reason: String(doctorResult.new_article_reason || ''),
      role_with_existing_articles: String(doctorResult.role_with_existing_articles || local.role_summary || ''),
      do_not_target: Array.isArray(doctorResult.do_not_target) && doctorResult.do_not_target.length
        ? doctorResult.do_not_target : (local.do_not_target || []),
      internal_link_candidates: c.internal_link_recommendations,
      closest_existing_articles: closest,
      local_gate: String(local.grade || ''),
      monitor_days: monitorDays,
      post_publish_policy: '公開後はSBMで約30日を初回評価点としてモニター。データ不足ならMONITOR延長。重大なカニバリが確認された場合はDoctor/Writer/Mergeへ再診断。',
      new_article_target: c.new_article_target,
      reference_articles: creatorReferenceArticles,
      bridge_identity_policy:'REFERENCE_ARTICLE_FOR_LEGACY_SBM_IDENTITY_ONLY'
    };
  } else if (decision === 'WRITER') {
    if (!closest.length || !String(closest[0].article_url || '')) {
      c.route_to = 'NEEDS_EVIDENCE';
      c.eventual_route = 'WRITER';
      c.doctor_decision = 'WRITER_RECOMMENDED_BUT_TARGET_UNRESOLVED';
      c.treatment_strategy = 'WAIT_FOR_EXISTING_ARTICLE_TARGET';
      c.additional_evidence_needed = ['Writerで改善する既存記事URLの特定'];
      c.reason = String(doctorResult.new_article_reason || doctorResult.role_with_existing_articles || '既存記事で検索意図を吸収できるため、新記事作成よりWriterを優先。ただし対象記事確定が必要。');
    } else {
      c.route_to = 'WRITER';
      c.eventual_route = 'WRITER';
      c.doctor_decision = 'WRITER_INSTEAD_OF_CREATOR';
      c.treatment_strategy = 'EXPAND_EXISTING_ARTICLE';
      c.target_articles = [closest[0]];
      c.reason = String(doctorResult.new_article_reason || doctorResult.role_with_existing_articles || 'SERP上の検索意図が既存記事と近いため、既存記事改善を優先。');
    }
    c.creator_plan = null;
  } else if (decision === 'BLOCK') {
    c.route_to = 'NO_ACTION';
    c.eventual_route = 'NO_ACTION';
    c.doctor_decision = 'CREATOR_BLOCKED_BY_CANNIBALIZATION_RISK';
    c.treatment_strategy = 'DO_NOT_CREATE_NEW_ARTICLE';
    c.reason = String(doctorResult.new_article_reason || doctorResult.role_with_existing_articles || '既存記事との重大なカニバリリスクがあるため新記事作成を見送る。');
    c.creator_plan = null;
  } else {
    c.route_to = 'NEEDS_EVIDENCE';
    c.eventual_route = 'CREATOR';
    c.doctor_decision = 'CREATOR_NEEDS_MORE_EVIDENCE';
    c.treatment_strategy = 'WAIT_FOR_CREATOR_VALIDATION_EVIDENCE';
    c.reason = String(doctorResult.new_article_reason || doctorResult.role_with_existing_articles || c.reason || 'Creator可否の追加確認が必要。');
    c.additional_evidence_needed = Array.isArray(doctorResult.additional_evidence_needed)
      ? doctorResult.additional_evidence_needed.map(String)
      : ['SERP独立性または既存記事との役割分担に関する追加Evidence'];
    c.creator_plan = null;
  }

  cases[idx] = c;
  stored.diagnosis_cases = cases;
  sdsdStoreSiteWideResult_(stored);
  sdsdWriteTreatmentPlan_(stored);
  sdsdClearCreatorSerpState_();
  try { sdsdWriteCreatorValidationSheet_(sdsdCreatorCandidateValidations_()); } catch(e) {}
  try { sdsdRenderHome_(); } catch(e) {}
  return c;
}

function sdsdRegisterCreatorSerpResult(text) {
  const parsed = sdsdExtractJsonObject_(String(text || ''));
  const updated = sdsdApplyCreatorSerpResult_(parsed);
  return {
    diagnosis_case_id: updated.diagnosis_case_id,
    diagnosis_theme: updated.diagnosis_theme,
    route_to: updated.route_to,
    decision: updated.doctor_decision,
    creator_plan: updated.creator_plan || null
  };
}

function sdsdOpenCreatorSerpReferral() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SDSD_CONFIG.sheets.creatorSerpReferral);
  if (!sh) {
    SpreadsheetApp.getUi().alert(
      'Creator SERP確認紹介状はまだ作成されていません。\n\n' +
      '「Creator候補チェック」で対象行を選択し、\n' +
      '設定・管理 → 保守・復旧操作 → Creator候補：選択案件のSERP確認紹介状を作成\n' +
      'を実行してください。'
    );
    return;
  }
  ss.setActiveSheet(sh);
}

function sdsdShowCreatorSerpResultDialog() {
  const state = sdsdGetCreatorSerpState_();
  const caseLabel = state.caseId ? `対象案件: ${state.caseId}` : '対象案件: JSON内のdiagnosis_case_idで照合します';
  const html = '<!doctype html><html><head><base target="_top"><meta charset="UTF-8"><style>' +
    'body{font-family:Arial,"Noto Sans JP",sans-serif;padding:18px;background:#f8f9fa;color:#202124}' +
    'h2{margin:0 0 6px;font-size:18px}.flow{background:#e8f0fe;color:#174ea6;padding:9px 11px;border-radius:7px;font-weight:700;font-size:12px;margin-bottom:10px}' +
    '.why{font-size:13px;line-height:1.7;background:#fff;padding:10px;border-radius:7px;margin-bottom:10px}.case{font-size:12px;color:#5f6368;margin-bottom:8px}' +
    'textarea{box-sizing:border-box;width:100%;height:330px;padding:10px;font:12px/1.45 monospace;white-space:pre;border:1px solid #bdc1c6;border-radius:7px;background:#fff}' +
    '.actions{display:flex;gap:8px;justify-content:flex-end;margin-top:10px}button{padding:9px 15px;border:0;border-radius:6px;font-weight:700;cursor:pointer}.primary{background:#1a73e8;color:#fff}.secondary{background:#e8eaed;color:#202124}.status{font-size:12px;margin-top:8px;color:#137333;white-space:pre-wrap}' +
    '</style></head><body><h2>DoctorのCreator SERP診断結果を取り込む</h2>' +
    '<div class="flow">Site Diagnosis → Doctor（SERP確認）→ Diagnosis → SBM → Creator</div>' +
    '<div class="why"><b>何のため？</b><br>候補キーワードが既存記事と重大なカニバリを起こさず、新記事として試す価値があるかを確定します。CREATORなら、メインKW・検索意図・役割分担・内部リンク・30日モニター条件をSBMへ引き継ぎます。</div>' +
    '<div class="case">' + caseLabel.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;') + '</div>' +
    '<textarea id="t" placeholder="Doctor回答全文または SIMS_DOCTOR_CREATOR_SERP_RESULT_V1 JSON を貼り付けてください"></textarea>' +
    '<div class="actions"><button class="secondary" onclick="google.script.host.close()">閉じる</button><button class="primary" onclick="send()">診断結果を取り込む</button></div><div id="s" class="status"></div>' +
    '<script>function send(){const t=document.getElementById("t").value.trim(),s=document.getElementById("s");if(!t){s.textContent="Doctor回答を貼り付けてください。";return}s.textContent="取り込んでいます...";google.script.run.withSuccessHandler(r=>{s.textContent="登録しました。次の処置: "+r.route_to+"\\nダイアログを閉じた後、▶ 次に進むで続けられます。";setTimeout(()=>google.script.host.close(),1200)}).withFailureHandler(e=>{s.textContent="取り込めませんでした："+((e&&e.message)?e.message:String(e||"不明なエラー"))}).sdsdRegisterCreatorSerpResult(t)}</script></body></html>';
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(760).setHeight(610),
    'Creator SERP診断結果の取込'
  );
}

function sdsdRefreshCreatorValidationSheet_() {
  const results = sdsdCreatorCandidateValidations_();
  if (!results.length) return [];
  sdsdWriteCreatorValidationSheet_(results);
  return results;
}

function sdsdOpenCreatorValidation() {
  const ss = SpreadsheetApp.getActive();
  const pending = sdsdCreatorValidationCases_().length;
  if (!pending) {
    SpreadsheetApp.getUi().alert('現在、新記事作成前チェックの対象案件はありません。');
    return;
  }

  // v0.6.4: Creator候補チェックは保存済みシートをそのまま再利用しない。
  // 判定ロジックをPATCHした後も旧GREEN判定が残らないよう、開くたびに
  // 現在のDoctor結果 + GSC Evidenceから再判定してシートを書き直す。
  sdsdRefreshCreatorValidationSheet_();
  const sh = ss.getSheetByName(SDSD_CONFIG.sheets.creatorValidation);
  if (!sh) return;
  ss.setActiveSheet(sh);
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
      article_url: url,
      canonical_url: url,
      canonical_url_source: 'DIAGNOSIS_ARTICLE_URL',
      requested_url: String(fetched.requestedUrl || url),
      initial_http_status: Number(fetched.initialHttpStatus || fetched.httpStatus || 0),
      redirect_detected: Boolean(fetched.redirectDetected),
      redirect_target_url: String(fetched.redirectTargetUrl || ''),
      redirect_chain: Array.isArray(fetched.redirectChain) ? fetched.redirectChain : [],
      final_url: String(fetched.finalUrl || url),
      final_http_status: Number(fetched.finalHttpStatus || fetched.httpStatus || 0),
      observed_html_canonical_url: String(fetched.canonicalUrl || ''),
      canonical_matches_final_url: Boolean(fetched.canonicalMatchesFinalUrl),
      redirect_canonical_consistent: Boolean(fetched.redirectCanonicalConsistent),
      canonical_mismatch: Boolean(
        fetched.canonicalUrl &&
        sdsdNormalizeUrl_(fetched.canonicalUrl) !== sdsdNormalizeUrl_(fetched.finalUrl || url)
      ),
      redirect_status: sdsdRedirectEvidenceStatus_(fetched),
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


function sdsdRedirectEvidenceStatus_(fetched) {
  const initial = Number(fetched && fetched.initialHttpStatus || 0);
  const finalStatus = Number(fetched && fetched.finalHttpStatus || fetched && fetched.httpStatus || 0);
  const redirected = Boolean(fetched && fetched.redirectDetected);
  const chain = fetched && Array.isArray(fetched.redirectChain) ? fetched.redirectChain : [];
  const canonical = String(fetched && fetched.canonicalUrl || '');
  const canonicalMatches = Boolean(fetched && fetched.canonicalMatchesFinalUrl);

  if (String(fetched && fetched.status || '') === 'FETCH_ERROR') return 'REDIRECT_ERROR';
  if (redirected && (initial === 302 || initial === 303 || initial === 307)) {
    return canonical && !canonicalMatches ? 'REDIRECT_CANONICAL_CONFLICT' : 'TEMPORARY_REDIRECT_WARNING';
  }
  if (redirected && chain.length > 1) {
    return canonical && !canonicalMatches ? 'REDIRECT_CANONICAL_CONFLICT' : 'REDIRECT_CHAIN_WARNING';
  }
  if (redirected && (initial === 301 || initial === 308)) {
    if (finalStatus >= 200 && finalStatus < 300 && canonicalMatches) return 'MERGE_REDIRECT_OK';
    if (canonical && !canonicalMatches) return 'REDIRECT_CANONICAL_CONFLICT';
    return 'REDIRECT_TARGET_REVIEW';
  }
  if (!redirected && finalStatus >= 200 && finalStatus < 300) {
    if (canonical && !canonicalMatches) return 'CANONICAL_MISMATCH';
    return 'NORMAL';
  }
  return 'REDIRECT_REVIEW';
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
    '- WRITERへ送る場合は、Writerが変更してよい範囲を allowed_scope、変更禁止範囲を blocked_scope として必ず返してください。',
    '- 複数クラスタで治療範囲が共通なら、ルート直下 workflow_handoff.allowed_scope / blocked_scope を共通既定値として返しても構いません。',
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
      ui.alert(
        '現在、追加Evidence精密診断へ送る案件はありません。\n\n' +
        'Homeを更新して「次に行うこと」を確認してください。'
      );
      try { sdsdRenderHome_(); } catch (e) {}
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
      `SIMS-Doctor-Site-Wide-Precision-${sdsdSafeSiteToken_()}-${Utilities.formatDate(
        new Date(),
        Session.getScriptTimeZone() || 'Asia/Tokyo',
        'yyyyMMdd-HHmmss'
      )}.zip`;

    const created = sdsdCreateZipInOutputFolder_(blobs, zipName);
    const file = created.file;

    sdsdSetSiteWidePrecisionPackageState_({
      status:'WAITING_DOCTOR_RESULT',
      fileId:file.getId(),
      fileUrl:file.getUrl(),
      fileName:zipName,
      clusterCount:cases.length
    });

    sdsdShowZipExportComplete_(
      'サイト横断の精密診断Packageを生成しました。',
      [`優先クラスタ: ${cases.length}件`],
      file,
      created.folderInfo,
      'このZIPをSIMS Doctorへ渡してください。Doctorの回答が返ったら「Doctor診断結果を取り込む」へ進みます。',
      {purpose:'優先クラスタの本文・Search Console等を使い、既存記事の最終処置をDoctorに確定してもらうPackageです。',primaryLabel:'Doctor診断結果を取り込む',primaryFn:'sdsdShowSiteWideDoctorResultDialog'}
    );

    return {
      clusterCount: cases.length,
      fileId: file.getId(),
      fileUrl: file.getUrl(),
      fileName: zipName
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
  if (Array.isArray(cluster.articles) && cluster.articles.length) return cluster.articles;
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


function sdsdPrecisionScopeArray_(obj, key) {
  if (!obj) return [];
  const direct = obj[key];
  if (Array.isArray(direct) && direct.length) return direct.slice();

  const handoff = obj.workflow_handoff || {};
  const h = handoff[key];
  if (Array.isArray(h) && h.length) return h.slice();

  const plan = obj.treatment_plan || {};
  const p = plan[key];
  if (Array.isArray(p) && p.length) return p.slice();

  return [];
}

function sdsdPrecisionResolveScopes_(root, cluster, cr, group) {
  const allowedCandidates = [
    sdsdPrecisionScopeArray_(group, 'allowed_scope'),
    sdsdPrecisionScopeArray_(cr, 'allowed_scope'),
    sdsdPrecisionScopeArray_(cluster, 'allowed_scope'),
    sdsdPrecisionScopeArray_(root, 'allowed_scope')
  ];
  const blockedCandidates = [
    sdsdPrecisionScopeArray_(group, 'blocked_scope'),
    sdsdPrecisionScopeArray_(cr, 'blocked_scope'),
    sdsdPrecisionScopeArray_(cluster, 'blocked_scope'),
    sdsdPrecisionScopeArray_(root, 'blocked_scope')
  ];

  const allowed = allowedCandidates.find(a => a.length) || [];
  const blocked = blockedCandidates.find(a => a.length) || [];
  return {allowed_scope:allowed.slice(), blocked_scope:blocked.slice()};
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
        cr.diagnosis_summary || cluster.diagnosis_summary || cluster.reason || ''
      );

      const mergePlan = g.merge_plan || cr.merge_plan || cluster.merge_plan || null;
      let targetArticles = sdsdPrecisionArticles_(g, cluster);

      if (route === 'MERGE' && mergePlan) {
        const mergeArticles = sdsdMergePlanArticles_(mergePlan);
        if (mergeArticles.length) targetArticles = mergeArticles;
      }

      const mergeText = route === 'MERGE' && mergePlan
        ? sdsdMergePlanText_(mergePlan)
        : {survivor:'',absorbed:'',direction:'',content_to_absorb:''};

      const treatmentPlanRaw =
        g.treatment_plan || cr.treatment_plan || cluster.treatment_plan || null;
      const scopes = sdsdPrecisionResolveScopes_(obj, cluster, cr, g);
      const treatmentPlan = treatmentPlanRaw
        ? Object.assign({}, treatmentPlanRaw, {
            allowed_scope: scopes.allowed_scope.length
              ? scopes.allowed_scope
              : (Array.isArray(treatmentPlanRaw.allowed_scope) ? treatmentPlanRaw.allowed_scope : []),
            blocked_scope: scopes.blocked_scope.length
              ? scopes.blocked_scope
              : (Array.isArray(treatmentPlanRaw.blocked_scope) ? treatmentPlanRaw.blocked_scope : [])
          })
        : {
            allowed_scope: scopes.allowed_scope,
            blocked_scope: scopes.blocked_scope
          };

      out.push({
        diagnosis_case_id: caseId,
        parent_diagnosis_case_id: baseId,
        diagnosis_theme: groupLabel ? `${diagnosisTheme} / ${groupLabel}` : diagnosisTheme,
        diagnosis_type: String(
          g.group_type || cluster.diagnosis_type || cluster.improvement_type || ''
        ),
        absorbed_source_case_ids: Array.isArray(cluster.absorbed_source_case_ids)
          ? cluster.absorbed_source_case_ids.map(String) : [],
        target_articles: targetArticles,
        doctor_decision: String(
          g.doctor_decision || g.decision || cr.doctor_decision ||
          cluster.doctor_decision || route
        ),
        confidence: String(g.confidence || cr.confidence || cluster.confidence || ''),
        site_impact: String(g.site_impact || cr.site_impact || cluster.site_impact || ''),
        treatment_strategy: String(
          g.treatment_strategy || cr.treatment_strategy || cluster.treatment_strategy || ''
        ),
        route_to: route,
        eventual_route: String(
          g.eventual_route || cr.eventual_route || cluster.eventual_route || ''
        ),
        reason: summary,
        additional_evidence_needed: Array.isArray(g.additional_evidence_needed)
          ? g.additional_evidence_needed.map(String)
          : (Array.isArray(cluster.additional_evidence_needed)
            ? cluster.additional_evidence_needed.map(String) : []),
        precision_group_type: groupLabel,
        treatment_plan: treatmentPlan,
        allowed_scope: scopes.allowed_scope,
        blocked_scope: scopes.blocked_scope,
        internal_link_recommendations: Array.isArray(cluster.internal_link_recommendations)
          ? cluster.internal_link_recommendations : [],
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
      treatment_plan: c.treatment_plan || null,
      internal_link_recommendations: Array.isArray(c.internal_link_recommendations)
        ? c.internal_link_recommendations : [],
      merge_plan: c.merge_plan || null,
      merge_survivor: String(c.merge_survivor || ''),
      merge_absorbed: String(c.merge_absorbed || ''),
      merge_direction: String(c.merge_direction || ''),
      merge_content_to_absorb: String(c.merge_content_to_absorb || ''),
      creator_plan: c.creator_plan || null
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


function sdsdApplyTreatmentPlanCompletionStyle_(sh, cases) {
  if (!sh || !Array.isArray(cases) || !cases.length) return;
  if (!sdsdIsSbmHandoffCurrent_()) return;

  const statusCol = 13;
  cases.forEach((c, i) => {
    const route = String(c && c.route_to || '').toUpperCase();
    if (!route || route === 'NEEDS_EVIDENCE') return;

    const row = i + 2;
    sh.getRange(row, 1, 1, 13)
      .setBackground('#F1F3F4')
      .setFontColor('#80868B');

    const statusText = route === 'NO_ACTION'
      ? 'SBM引き渡し済み（処置不要）'
      : 'SBM引き渡し済み';

    sh.getRange(row, statusCol)
      .setValue(statusText)
      .setFontWeight('bold')
      .setFontColor('#5F6368');
  });
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
    c.route_to === 'NEEDS_EVIDENCE' && sdsdEvidencePurpose_(c) === 'CREATOR_VALIDATION'
      ? '新記事作成前チェック'
      : (routeJa[c.route_to] || c.route_to),
    c.route_to === 'CREATOR'
      ? `新規記事（未発行） / キーワード: ${String((c.new_article_target && c.new_article_target.candidate_keyword) || (c.creator_plan && c.creator_plan.candidate_keyword) || c.diagnosis_theme || '')}`
      : (c.target_articles || []).map(a => {
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
    c.route_to === 'NEEDS_EVIDENCE'
      ? (sdsdEvidencePurpose_(c) === 'CREATOR_VALIDATION' ? '新記事候補確認待ち' : '精密診断待ち')
      : (c.route_to === 'CREATOR' ? 'SBM引き渡し待ち' : '処置振り分け済み')
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

  sdsdApplyTreatmentPlanCompletionStyle_(sh, obj.diagnosis_cases || []);

  ss.setActiveSheet(sh);
  return rows.length;
}


function sdsdEnsureSiteWideResultImportSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SDSD_CONFIG.sheets.siteWideResultImport);
  if (!sh) sh = ss.insertSheet(SDSD_CONFIG.sheets.siteWideResultImport);

  const title = 'Doctor診断結果をここに貼り付け（横断診断／精密診断 共通）';
  const instruction =
    'このセル以下に、Doctor回答全文、SIMS_DOCTOR_SITE_WIDE_RESULT_V1 または SIMS_DOCTOR_SITE_WIDE_PRECISION_RESULT_V1 JSON を貼り付けてください。\n' +
    '貼り付け後、メニュー最上段の「▶ 次に進む（Diagnosisに任せる）」を実行してください。';

  const currentA1 = String(sh.getRange('A1').getValue() || '');
  const currentA2 = String(sh.getRange('A2').getValue() || '');

  // Migration-safe: update labels without clearing an already pasted Doctor result.
  sh.getRange('A1').setValue(title);
  const looksLikePlaceholder =
    !currentA2.trim() ||
    currentA2.indexOf('このセル以下に、Doctor回答全文') >= 0;

  if (looksLikePlaceholder) {
    sh.getRange('A2').setValue(instruction);
  }

  sh.getRange('A1').setFontWeight('bold').setFontSize(14);
  sh.getRange('A1:A2').setWrap(true);
  sh.setColumnWidth(1, 1000);
  if (looksLikePlaceholder) sh.setRowHeight(2, 90);
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
    '貼り付け後、メニュー最上段の「▶ 次に進む（Diagnosisに任せる）」を実行してください。'
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
      sdsdSetSiteWideRegisterStage_('WAITING_INPUT', 'Doctor診断結果の入力待ち');
      sdsdShowSiteWideDoctorResultDialog();
      return;
    }

    sdsdSetSiteWideRegisterStage_('EXTRACT_JSON', `入力文字数: ${sourceText.length}`);
    const parsed = sdsdExtractJsonObject_(sourceText);

    sdsdSetSiteWideRegisterStage_('VALIDATE', parsed && parsed.format ? parsed.format : 'formatなし');
    sdsdValidateSiteWideResult_(parsed);

    sdsdSetSiteWideRegisterStage_('NORMALIZE', 'Doctor結果を正規化');
    const normalizedInput = sdsdNormalizeDoctorSiteWideResult_(parsed);
    const isPrecision = sdsdIsSiteWidePrecisionResult_(parsed);
    const normalized = isPrecision
      ? sdsdMergePrecisionIntoStoredResult_(normalizedInput)
      : normalizedInput;

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

    if (isPrecision) {
      sdsdClearSiteWidePrecisionPackageState_();
    }

    sdsdSetSiteWideRegisterStage_('COMPLETE', `登録完了 ${count}件`);
    sdsdClearSiteWideResultImport_();
    try{sdsdRenderHome_();}catch(e){}

    const completionMessage = isPrecision
      ? (
        `Doctor精密診断結果を登録しました。\n\n` +
        `今回確定: ${(normalizedInput.diagnosis_cases || []).length}件\n` +
        `残り追加Evidence: ${counts.NEEDS_EVIDENCE || 0}件\n` +
        `Writer: ${counts.WRITER || 0}件\n` +
        `Merge: ${counts.MERGE || 0}件\n` +
        `Creator: ${counts.CREATOR || 0}件\n` +
        `経過観察/処置不要: ${(counts.MONITOR || 0) + (counts.NO_ACTION || 0)}件\n\n` +
        `続きは「1. Site Diagnosisを進める」で進められます。`
      )
      : (
        `Doctor診断結果を登録しました。\n\n` +
        `診断案件: ${count}件\n` +
        `Writer: ${counts.WRITER || 0}件\n` +
        `Merge: ${counts.MERGE || 0}件\n` +
        `Creator: ${counts.CREATOR || 0}件\n` +
        `追加Evidence: ${counts.NEEDS_EVIDENCE || 0}件\n` +
        `経過観察/処置不要: ${(counts.MONITOR || 0) + (counts.NO_ACTION || 0)}件\n\n` +
        `「サイト治療計画」を確認してください。`
      );

    ui.alert(completionMessage);

  } catch(e) {
    const props = PropertiesService.getDocumentProperties();
    const stage = String(
      props.getProperty('SDSD_SITE_WIDE_REGISTER_STAGE') || 'UNKNOWN'
    );
    const message = String(e && e.message ? e.message : e);

    sdsdSetSiteWideRegisterStage_('ERROR', `${stage}: ${message}`);

    ui.alert(
      `Doctor診断結果を登録できませんでした。\n\n` +
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


// ============================================================================
// v0.7.0 Creator Long-tail Discovery
// Existing query clusters -> unexplored long-tail discovery -> Creator handoff.
// Doctor performs live SERP discovery/validation; Diagnosis preserves site/GSC context.
// ============================================================================

function sdsdLongtailDiscoveryState_() {
  const p = PropertiesService.getDocumentProperties();
  return {
    status: String(p.getProperty('SDSD_LONGTAIL_DISCOVERY_STATUS') || ''),
    caseId: String(p.getProperty('SDSD_LONGTAIL_DISCOVERY_CASE_ID') || '')
  };
}

function sdsdSetLongtailDiscoveryState_(status, caseId) {
  const p = PropertiesService.getDocumentProperties();
  p.setProperty('SDSD_LONGTAIL_DISCOVERY_STATUS', String(status || ''));
  p.setProperty('SDSD_LONGTAIL_DISCOVERY_CASE_ID', String(caseId || ''));
  p.setProperty('SDSD_LONGTAIL_DISCOVERY_AT', new Date().toISOString());
}

function sdsdClearLongtailDiscoveryState_() {
  const p = PropertiesService.getDocumentProperties();
  ['SDSD_LONGTAIL_DISCOVERY_STATUS','SDSD_LONGTAIL_DISCOVERY_CASE_ID','SDSD_LONGTAIL_DISCOVERY_AT']
    .forEach(k => p.deleteProperty(k));
}

function sdsdBuildLongtailDiscoveryReferralObject_(selected) {
  const v = selected.validation || {};
  const siteResult = sdsdReadStoredSiteWideResult_();
  const existingQueries = (v.cluster_queries || []).map(x => String(x.query || '')).filter(Boolean);
  const existingUrls = Array.from(new Set((v.internal_link_candidates || []).map(x => x.article_url).filter(Boolean)));
  return {
    format: 'SIMS_DOCTOR_LONGTAIL_DISCOVERY_REFERRAL_V1',
    contract_version: '1.0',
    site_diagnosis_batch_id: String(siteResult.site_diagnosis_batch_id || ''),
    site: siteResult.site || {},
    source_diagnosis_case_id: String(v.diagnosis_case_id || ''),
    source_keyword: String(v.candidate_query || v.diagnosis_theme || ''),
    keyword_cluster: String(v.keyword_cluster || ''),
    local_gate: String(v.grade || ''),
    local_reason: String(v.reason || ''),
    existing_cluster_queries: v.cluster_queries || [],
    existing_query_strings: existingQueries,
    closest_existing_articles: v.closest_articles || [],
    internal_link_candidates: v.internal_link_candidates || [],
    existing_article_urls: existingUrls,
    discovery_policy: {
      objective: '既存キーワードクラスターから、既存記事がまだ主担当していない「＋1語」を中心とするロングテール検索意図を発見する',
      rules: [
        'source_keywordそのもの・語順違い・単純な年号違いだけの候補は除外する',
        '既存GSCクエリと実質同一の候補は除外する',
        '検索意図を具体化する追加語（対象、状況、悩み、行動、比較、失敗、年代、場所など）を優先する',
        '実際のSERPを確認し、元キーワードと上位ページ/回答内容が十分に分離する候補だけCREATOR候補とする',
        '明確な同一意図・重大なカニバリがある候補はBLOCKする',
        '多少の不確実性は許容し、独立性がMEDIUM以上なら30日モニター前提でCREATORを許可してよい',
        '候補は最大5件。価値のある候補がなければ0件で返してよい'
      ]
    },
    required_response_format: 'SIMS_DOCTOR_LONGTAIL_DISCOVERY_RESULT_V1'
  };
}

function sdsdBuildLongtailDiscoveryReferralMarkdown_(o) {
  return [
    '# SIMS Doctor ロングテール探索紹介状',
    '',
    `元案件ID: ${o.source_diagnosis_case_id}`,
    `元キーワード: ${o.source_keyword}`,
    `キーワードクラスター: ${o.keyword_cluster}`,
    `Diagnosis判定: ${o.local_gate}`,
    '',
    '## 目的',
    '既存記事と同じキーワードをもう1本作るのではなく、既存クラスターから「＋1語」を中心に未開拓の検索意図を探してください。',
    '候補は実SERPで独立性を確認し、カニバリを避けつつ、個人ブログとして試す価値のあるロングテールを優先します。',
    '',
    '## 既存クラスターGSCクエリ',
    ...(o.existing_cluster_queries || []).map(x => `- ${x.query || ''} | ${x.article_url || ''} | imp=${x.impressions || 0} | pos=${x.position || 0}`),
    '',
    '## 近い既存記事',
    ...(o.closest_existing_articles || []).map(x => `- ${x.article_title || ''} | ${x.main_query || ''} | ${x.article_url || ''}`),
    '',
    '## Doctorへの依頼',
    ...((o.discovery_policy && o.discovery_policy.rules) || []).map(x => `- ${x}`),
    '',
    '## 返却JSON例',
    '```json',
    JSON.stringify({
      format:'SIMS_DOCTOR_LONGTAIL_DISCOVERY_RESULT_V1',
      source_diagnosis_case_id:o.source_diagnosis_case_id,
      source_keyword:o.source_keyword,
      candidates:[{
        candidate_keyword:'',
        added_intent_term:'',
        decision:'CREATOR|BLOCK|WRITER',
        serp_independence:'HIGH|MEDIUM|LOW',
        search_intent:'',
        competitor_summary:'',
        role_with_existing_articles:'',
        do_not_target:[],
        internal_link_recommendations:[],
        new_article_reason:'',
        monitor_days:30,
        confidence:0
      }],
      overall_note:''
    }, null, 2),
    '```'
  ].join('\n');
}

function sdsdWriteLongtailDiscoveryReferralSheet_(o) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SDSD_CONFIG.sheets.longtailDiscovery);
  if (!sh) sh = ss.insertSheet(SDSD_CONFIG.sheets.longtailDiscovery);
  sh.clear();
  const md = sdsdBuildLongtailDiscoveryReferralMarkdown_(o);
  const rows = [
    ['ロングテール探索','Doctor紹介状'],
    ['元案件ID',o.source_diagnosis_case_id],
    ['元キーワード',o.source_keyword],
    ['キーワードクラスター',o.keyword_cluster],
    ['Diagnosis判定',o.local_gate],
    ['Doctorへ渡す紹介状（Markdown）',md],
    ['Doctorへ渡す紹介状（JSON）',JSON.stringify(o,null,2)]
  ];
  sh.getRange(1,1,rows.length,2).setValues(rows);
  sh.getRange(1,1,rows.length,2).setWrap(true).setVerticalAlignment('top');
  sh.getRange(1,1,1,2).setFontWeight('bold');
  sh.setColumnWidth(1,270); sh.setColumnWidth(2,900);
  ss.setActiveSheet(sh);
}

function sdsdCreateLongtailDiscoveryReferralFromSelected() {
  const ui = SpreadsheetApp.getUi();
  try {
    const selected = sdsdSelectedCreatorValidation_();
    const referral = sdsdBuildLongtailDiscoveryReferralObject_(selected);
    sdsdWriteLongtailDiscoveryReferralSheet_(referral);
    sdsdSetLongtailDiscoveryState_('WAITING_DOCTOR_RESULT', referral.source_diagnosis_case_id);
    ui.alert(
      'ロングテール探索紹介状を作成しました。\n\n' +
      `元キーワード: ${referral.source_keyword}\n\n` +
      '何のため？：既存記事と同じ意図を増やさず、既存キーワードに意味のある「＋1語」を加えた未開拓検索意図を探します。\n' +
      '次の操作：「ロングテール探索」シートのMarkdownまたはJSONをSIMS Doctorへ渡してください。\n' +
      'Doctor回答が返ったら「確認する → Doctorロングテール探索結果を取り込む」から貼り付けてください。'
    );
    return referral;
  } catch(e) {
    ui.alert('ロングテール探索紹介状を作成できませんでした。\n\n' + String(e && e.message ? e.message : e));
    throw e;
  }
}

function sdsdValidateLongtailDiscoveryResult_(o) {
  if (!o || o.format !== 'SIMS_DOCTOR_LONGTAIL_DISCOVERY_RESULT_V1') {
    throw new Error('Doctor結果のformatがSIMS_DOCTOR_LONGTAIL_DISCOVERY_RESULT_V1ではありません。');
  }
  if (!String(o.source_diagnosis_case_id || '').trim()) throw new Error('source_diagnosis_case_idがありません。');
  if (!Array.isArray(o.candidates)) throw new Error('candidates配列がありません。');
  o.candidates.forEach((c,i) => {
    if (!String(c.candidate_keyword || '').trim()) throw new Error(`candidates[${i}].candidate_keywordがありません。`);
    const d = String(c.decision || '').toUpperCase();
    if (!['CREATOR','BLOCK','WRITER'].includes(d)) throw new Error(`candidates[${i}].decisionはCREATOR / BLOCK / WRITERで返してください。`);
  });
  return true;
}

function sdsdWriteLongtailDiscoveryResultSheet_(o) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SDSD_CONFIG.sheets.longtailDiscovery);
  if (!sh) sh = ss.insertSheet(SDSD_CONFIG.sheets.longtailDiscovery);
  sh.clear();
  const headers = ['元案件ID','元キーワード','候補キーワード','追加意図語','判定','SERP独立性','検索意図','新記事作成理由','既存記事との役割分担','狙わない検索意図','内部リンク候補','確信度','モニター'];
  const rows = (o.candidates || []).map(c => [
    String(o.source_diagnosis_case_id || ''),
    String(o.source_keyword || ''),
    String(c.candidate_keyword || ''),
    String(c.added_intent_term || ''),
    String(c.decision || '').toUpperCase(),
    String(c.serp_independence || ''),
    String(c.search_intent || ''),
    String(c.new_article_reason || ''),
    String(c.role_with_existing_articles || ''),
    Array.isArray(c.do_not_target) ? c.do_not_target.join(' / ') : '',
    Array.isArray(c.internal_link_recommendations) ? c.internal_link_recommendations.map(x => `${x.title || ''} / ${x.url || x.article_url || ''} / ${x.reason || ''}`).join('\n') : '',
    Number(c.confidence || 0),
    Number(c.monitor_days || 30) + '日'
  ]);
  sh.getRange(1,1,1,headers.length).setValues([headers]);
  if (rows.length) sh.getRange(2,1,rows.length,headers.length).setValues(rows);
  sh.setFrozenRows(1);
  sh.getRange(1,1,1,headers.length).setFontWeight('bold');
  sh.getRange(1,1,Math.max(rows.length+1,1),headers.length).setWrap(true).setVerticalAlignment('top');
  [190,220,250,160,100,120,360,420,420,300,480,90,100].forEach((w,i) => sh.setColumnWidth(i+1,w));
  ss.setActiveSheet(sh);
}

function sdsdRegisterLongtailDiscoveryResult(text) {
  const o = sdsdExtractJsonObject_(String(text || ''));
  sdsdValidateLongtailDiscoveryResult_(o);
  sdsdWriteLongtailDiscoveryResultSheet_(o);
  sdsdClearLongtailDiscoveryState_();
  return {count:o.candidates.length, creator_count:o.candidates.filter(x => String(x.decision||'').toUpperCase()==='CREATOR').length};
}

function sdsdShowLongtailDiscoveryResultDialog() {
  const state = sdsdLongtailDiscoveryState_();
  const caseLabel = state.caseId ? `対象案件: ${state.caseId}` : '対象案件: JSON内のsource_diagnosis_case_idで照合します';
  const html = '<!doctype html><html><head><base target="_top"><meta charset="UTF-8"><style>'+
    'body{font-family:Arial,"Noto Sans JP",sans-serif;padding:18px;background:#f8f9fa;color:#202124}h2{margin:0 0 6px;font-size:18px}.flow{background:#e8f0fe;color:#174ea6;padding:9px 11px;border-radius:7px;font-weight:700;font-size:12px;margin-bottom:10px}.why{font-size:13px;line-height:1.7;background:#fff;padding:10px;border-radius:7px;margin-bottom:10px}.case{font-size:12px;color:#5f6368;margin-bottom:8px}textarea{box-sizing:border-box;width:100%;height:330px;padding:10px;font:12px/1.45 monospace;border:1px solid #bdc1c6;border-radius:7px;background:#fff}.actions{display:flex;gap:8px;justify-content:flex-end;margin-top:10px}button{padding:9px 15px;border:0;border-radius:6px;font-weight:700;cursor:pointer}.primary{background:#1a73e8;color:#fff}.secondary{background:#e8eaed;color:#202124}.status{font-size:12px;margin-top:8px;color:#137333;white-space:pre-wrap}</style></head><body>'+
    '<h2>Doctorのロングテール探索結果を取り込む</h2><div class="flow">既存KWクラスター → Doctor（＋1語探索・SERP確認）→ Diagnosis → SBM → Creator</div>'+
    '<div class="why"><b>何のため？</b><br>既存記事と競合しない未開拓ロングテール候補を取り込み、CREATOR候補を一覧化します。候補が0件でも正常です。</div><div class="case">'+caseLabel.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')+'</div>'+
    '<textarea id="t" placeholder="Doctor回答全文またはSIMS_DOCTOR_LONGTAIL_DISCOVERY_RESULT_V1 JSONを貼り付け"></textarea><div class="status" id="s"></div><div class="actions"><button class="secondary" onclick="google.script.host.close()">キャンセル</button><button class="primary" id="go">探索結果を取り込む</button></div>'+
    '<script>document.getElementById("go").onclick=()=>{const t=document.getElementById("t").value.trim(),b=document.getElementById("go"),s=document.getElementById("s");if(!t){s.textContent="Doctor回答を貼り付けてください。";return;}b.disabled=true;s.textContent="取り込んでいます...";google.script.run.withSuccessHandler(r=>{s.textContent="取り込み完了：候補 "+r.count+"件 / CREATOR "+r.creator_count+"件";setTimeout(()=>google.script.host.close(),1000);}).withFailureHandler(e=>{b.disabled=false;s.textContent=(e&&e.message)?e.message:String(e||"エラー");}).sdsdRegisterLongtailDiscoveryResult(t);};</script></body></html>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(760).setHeight(560),'ロングテール探索結果の取込');
}

function sdsdOpenLongtailDiscovery() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SDSD_CONFIG.sheets.longtailDiscovery);
  if (!sh) {
    SpreadsheetApp.getUi().alert(
      'ロングテール探索はまだ作成されていません。\n\n' +
      '「Creator候補チェック」でYELLOW案件を選択し、\n' +
      '設定・管理 → 保守・復旧操作 → Creator候補：選択案件のロングテール探索紹介状を作成\n' +
      'を実行してください。'
    );
    return;
  }
  ss.setActiveSheet(sh);
}

function sdsdPromoteSelectedLongtailCandidateToCreator() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getActiveSheet();
  if (!sh || sh.getName() !== SDSD_CONFIG.sheets.longtailDiscovery) throw new Error('「ロングテール探索」シートでCREATOR候補の行を選択してください。');
  const row = sh.getActiveRange().getRow();
  if (row < 2) throw new Error('見出しではなく候補行を選択してください。');
  const values = sh.getRange(row,1,1,13).getValues()[0];
  const sourceCaseId = String(values[0] || '').trim();
  const sourceKeyword = String(values[1] || '').trim();
  const keyword = String(values[2] || '').trim();
  const decision = String(values[4] || '').toUpperCase();
  if (decision !== 'CREATOR') throw new Error('CREATOR判定の候補だけをSBM引き渡し対象にできます。');
  const stored = sdsdReadStoredSiteWideResult_();
  const source = (stored.diagnosis_cases || []).find(c => String(c.diagnosis_case_id || '') === sourceCaseId);
  if (!source) throw new Error('元案件を保存済みサイト治療計画から確認できません。');
  const slug = Utilities.base64EncodeWebSafe(keyword).replace(/[^A-Za-z0-9]/g,'').slice(0,10) || String(new Date().getTime());
  const newCaseId = `${sourceCaseId}-LT-${slug}`;
  const existing = (stored.diagnosis_cases || []).find(c => String(c.diagnosis_case_id || '') === newCaseId);
  if (existing) throw new Error('このロングテール候補はすでにCreator案件として登録されています。');
  const internalLinksText = String(values[10] || '');
  const internalLinks = internalLinksText.split('\n').map(line => line.trim()).filter(Boolean).map(line => ({note:line}));
  const targetArticles = Array.isArray(source.target_articles) ? source.target_articles : [];
  // v0.7.3: These are existing articles used only as reference/identity evidence.
  // They are NOT the new Creator article itself. Keep them for backward-compatible
  // SBM identity verification, but separate their semantic role explicitly.
  const referenceArticles = targetArticles.map(a => {
    if (typeof a === 'string') return {article_url:String(a), role:'REFERENCE_ONLY'};
    return Object.assign({}, a || {}, {role:'REFERENCE_ONLY'});
  });
  // The parent NEEDS_EVIDENCE case is now resolved by this derived Creator case.
  // Keep the original case for traceability, but exclude it from future Creator validation.
  source.creator_validation_status = 'RESOLVED_TO_LONGTAIL_CREATOR';
  source.creator_resolution_case_id = newCaseId;
  source.creator_resolution_keyword = keyword;
  source.creator_resolution_at = new Date().toISOString();

  const c = {
    diagnosis_case_id:newCaseId,
    diagnosis_theme:keyword,
    diagnosis_type:'LONGTAIL_DISCOVERY_CREATOR',
    absorbed_source_case_ids:[sourceCaseId],
    // Compatibility identity only. Site Diagnosis UI and the explicit Creator fields
    // below treat the new article as the true target.
    target_articles:referenceArticles,
    reference_articles:referenceArticles,
    new_article_target:{
      target_type:'NEW_ARTICLE',
      candidate_keyword:keyword,
      article_id:'',
      article_url:'',
      publication_status:'NOT_PUBLISHED'
    },
    doctor_decision:'CREATOR_APPROVED_AFTER_LONGTAIL_DISCOVERY',
    confidence:String(values[11] || ''),
    site_impact:'成長機会',
    treatment_strategy:'CREATE_NEW_ARTICLE_AND_MONITOR',
    route_to:'CREATOR',
    eventual_route:'CREATOR',
    reason:String(values[7] || ''),
    additional_evidence_needed:[],
    internal_link_recommendations:internalLinks,
    creator_plan:{
      candidate_keyword:keyword,
      source_keyword:sourceKeyword,
      added_intent_term:String(values[3] || ''),
      search_intent:String(values[6] || ''),
      serp_independence:String(values[5] || ''),
      new_article_reason:String(values[7] || ''),
      role_with_existing_articles:String(values[8] || ''),
      do_not_target:String(values[9] || '').split(' / ').filter(Boolean),
      internal_link_candidates:internalLinks,
      monitor_days:parseInt(String(values[12] || '30').replace(/\D/g,''),10) || 30,
      post_publish_policy:'公開後はSBMで約30日を初回評価点としてモニター。データ不足ならMONITOR延長。重大なカニバリが確認された場合はDoctor/Writer/Mergeへ再診断。',
      discovery_source:'SIMS_DOCTOR_LONGTAIL_DISCOVERY_RESULT_V1',
      new_article_target:{
        target_type:'NEW_ARTICLE',
        candidate_keyword:keyword,
        publication_status:'NOT_PUBLISHED'
      },
      reference_articles:referenceArticles,
      bridge_identity_policy:'REFERENCE_ARTICLE_FOR_LEGACY_SBM_IDENTITY_ONLY'
    }
  };
  stored.diagnosis_cases = (stored.diagnosis_cases || []).concat([c]);
  // A newly-finalized Creator case creates new work for SBM even if an earlier
  // Writer/Merge handoff in this diagnosis session had already been marked complete.
  sdsdSetSbmHandoffState_('PENDING_NEW_ACTIONABLE');
  PropertiesService.getDocumentProperties().deleteProperty('SDSD_SBM_HANDOFF_FINGERPRINT');
  sdsdStoreSiteWideResult_(stored);
  sdsdWriteTreatmentPlan_(stored);
  try { sdsdWriteCreatorValidationSheet_(sdsdCreatorCandidateValidations_()); } catch(e) {}
  try { sdsdRenderHome_(); } catch(e) {}
  SpreadsheetApp.getUi().alert(`Creator案件として登録しました。\n\n候補キーワード: ${keyword}\n次はSite DiagnosisからSBMへ引き渡してください。`);
  return c;
}
