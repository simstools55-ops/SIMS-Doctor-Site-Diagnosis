function sdsdNormalizeQuery_(q) {
  return String(q || '').toLowerCase().replace(/[　\s]+/g, ' ').trim();
}

function sdsdBuildQueryUrlIndex_() {
  const rows = sdsdReadObjects_(SDSD_CONFIG.sheets.evidencePageQuery);
  const index = {};
  rows.forEach(r => {
    const url = sdsdNormalizeUrl_(sdsdObjectValue_(r, ['page','key','url','URL','記事URL']));
    const query = String(sdsdObjectValue_(r, ['query','クエリ','検索クエリ']) || '').trim();
    const nq = sdsdNormalizeQuery_(query);
    if (!url || !nq) return;
    (index[nq] = index[nq] || []).push({
      query: query, url: url,
      clicks: Number(sdsdObjectValue_(r, ['clicks','クリック数']) || 0),
      impressions: Number(sdsdObjectValue_(r, ['impressions','表示回数']) || 0),
      position: Number(sdsdObjectValue_(r, ['position','掲載順位','平均掲載順位']) || 0)
    });
  });
  return index;
}

function sdsdDetectSiteOpportunities_() {
  const index = sdsdBuildQueryUrlIndex_();
  const articleMap = sdsdArticleTitleMap_();
  const out = [];

  Object.keys(index).forEach(nq => {
    const byUrl = {};
    index[nq].forEach(x => {
      const a = byUrl[x.url] || {url:x.url,query:x.query,clicks:0,impressions:0,weightedPosition:0};
      a.clicks += x.clicks;
      a.impressions += x.impressions;
      a.weightedPosition += x.position * Math.max(x.impressions,1);
      byUrl[x.url] = a;
    });

    const urls = Object.keys(byUrl).map(u => {
      const a=byUrl[u];
      a.position=a.impressions ? a.weightedPosition/a.impressions : 0;
      return a;
    }).sort((a,b)=>b.impressions-a.impressions);

    const totalImp=urls.reduce((s,x)=>s+x.impressions,0);
    if (totalImp < 10) return;

    if (urls.length >= 2) {
      const second=urls[1];
      const share=totalImp ? second.impressions/totalImp : 0;
      if (second.impressions >= 5 && share >= 0.20) {
        out.push({
          type:'カニバリ疑い', priority:totalImp>=100?1:2, theme:urls[0].query,
          targets:[urls[0].url,urls[1].url],
          evidence:`同じ検索クエリで複数記事に表示が分散（合計${Math.round(totalImp)}表示、2番目の記事の比率${Math.round(share*100)}%）`,
          confidence:totalImp>=100?'高':'中',
          action:'Doctorで検索意図と記事本文を比較', destination:'Doctor'
        });
      }
    }

    const lead=urls[0];
    if (lead && totalImp >= 30 && lead.position >= 10) {
      const titleNorm=sdsdNormalizeQuery_(articleMap[lead.url] || '');
      const terms=nq.split(' ').filter(t=>t.length>=2);
      const titleHit=terms.length ? terms.filter(t=>titleNorm.indexOf(t)>=0).length/terms.length : 0;

      if (titleHit < 0.5) {
        out.push({
          type:'新規記事機会', priority:totalImp>=100?1:2, theme:lead.query,
          targets:[lead.url],
          evidence:`検索需要${Math.round(totalImp)}表示に対し、主な表示記事の順位は${lead.position.toFixed(1)}位。記事タイトルとの一致も弱い`,
          confidence:totalImp>=100?'高':'中',
          action:'既存記事との検索意図重複を確認後、新記事候補として評価', destination:'Creator'
        });
      } else {
        out.push({
          type:'コンテンツギャップ', priority:totalImp>=100?1:2, theme:lead.query,
          targets:[lead.url],
          evidence:`既存記事に関連する検索需要${Math.round(totalImp)}表示があるが、主な掲載順位は${lead.position.toFixed(1)}位`,
          confidence:totalImp>=100?'高':'中',
          action:'既存記事が検索意図へ十分回答しているか確認', destination:'Writer / Doctor'
        });
      }
    }
  });

  return out.sort((a,b)=>a.priority-b.priority || a.type.localeCompare(b.type,'ja'));
}

function sdsdWriteSiteOpportunities_(rows) {
  const ss=SpreadsheetApp.getActive();
  let sh=ss.getSheetByName(SDSD_CONFIG.sheets.opportunities);
  if (!sh) sh=ss.insertSheet(SDSD_CONFIG.sheets.opportunities);
  sh.clear();

  const titleMap=sdsdArticleTitleMap_();
  const headers=['優先順位','改善テーマ','検索テーマ','対象記事','根拠','確信度','推奨対応','担当'];
  const values=rows.map((x,i)=>[
    i+1,x.type,x.theme,
    x.targets.map(u=>sdsdDisplayTitle_(u,titleMap)+'\n'+u).join('\n\n'),
    x.evidence,x.confidence,x.action,x.destination
  ]);

  sh.getRange(1,1,1,headers.length).setValues([headers]);
  if (values.length) sh.getRange(2,1,values.length,headers.length).setValues(values);
  sh.setFrozenRows(1);
  sh.getRange(1,1,1,headers.length).setFontWeight('bold');
  sh.getRange(1,1,Math.max(1,values.length+1),headers.length).setWrap(true);
  [70,150,260,430,430,90,360,120].forEach((w,i)=>sh.setColumnWidth(i+1,w));
  if (!values.length) {
    sh.getRange(2,1,1,headers.length).merge();
    sh.getRange(2,1).setValue('現在のEvidenceから、中以上の確信度で提示できる横断改善候補は見つかりませんでした。');
  }
}

function sdsdRunSiteOpportunityDiagnosis() {
  try {
    if (!sdsdQueryEvidenceSourceCount_()) {
      SpreadsheetApp.getUi().alert('サイト横断診断を実行できません。\n\npage_query_top のEvidenceがありません。');
      return;
    }
    const rows=sdsdDetectSiteOpportunities_();
    sdsdWriteSiteOpportunities_(rows);
    const counts={};
    rows.forEach(x=>counts[x.type]=(counts[x.type]||0)+1);
    const sh=SpreadsheetApp.getActive().getSheetByName(SDSD_CONFIG.sheets.opportunities);
    if (sh) SpreadsheetApp.getActive().setActiveSheet(sh);
    SpreadsheetApp.getUi().alert(
      `サイト横断診断が完了しました。\n\n`+
      `カニバリ疑い: ${counts['カニバリ疑い']||0}件\n`+
      `新規記事機会: ${counts['新規記事機会']||0}件\n`+
      `コンテンツギャップ: ${counts['コンテンツギャップ']||0}件\n\n`+
      `これは一次候補です。自動で記事統合・新規作成・リライトは行いません。`
    );
  } catch(e) {
    SpreadsheetApp.getUi().alert(`サイト横断診断を完了できませんでした。\n\n${e.message||e}`);
    throw e;
  }
}
