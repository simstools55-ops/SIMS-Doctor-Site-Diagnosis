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
