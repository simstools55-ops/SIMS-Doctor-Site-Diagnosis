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
