function sdsdNormalizeUrl_(url) {
  const s = String(url || '').trim().split('#')[0].split('?')[0];
  return s.replace(/\/+$/, '');
}
