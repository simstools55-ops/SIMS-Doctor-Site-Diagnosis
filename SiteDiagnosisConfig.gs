const SDSD_VERSION = '0.3.3-Sprint3-Query-Evidence-Hotfix';

const SDSD_CONFIG = Object.freeze({
  sheets: {
    evidencePageSummary: '_SDSD_PAGE_SUMMARY',
    evidencePageWeekly: '_SDSD_PAGE_WEEKLY',
    evidencePageQuery: '_SDSD_PAGE_QUERY_TOP',
    sbmHistory: '_SDSD_SBM_HISTORY',
    candidates: 'Site Diagnosis Candidates',
    selectedCases: 'Selected Treatment Cases',
    articleMaster: '_SDSD_ARTICLE_MASTER'
  },
  score: {
    demandMax: 30,
    opportunityMax: 30,
    urgencyMax: 25,
    assetMax: 15
  },
  guardDays: 35
});
