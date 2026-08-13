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
