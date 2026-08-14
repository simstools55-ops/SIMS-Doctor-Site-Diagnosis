const SDSD_VERSION = '0.4.0-RC5';

const SDSD_CONFIG = Object.freeze({
  sheets: {
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
    treatmentPlan: 'サイト治療計画'
  },
  score: {
    demandMax: 30,
    opportunityMax: 30,
    urgencyMax: 25,
    assetMax: 15
  },
  guardDays: 35
});
