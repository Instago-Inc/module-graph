// Thin Microsoft Graph client for hgi-v8 runtime
// - configure({ tenant?, clientId?, clientSecret?, refreshToken?, scope? })
// - ensureAccessToken(authOverrides?) -> Promise<string>
// - json({ path, method='GET', headers, bodyObj, debug }) -> { ok, data, status, error? }

(function(){
  const httpx = require('http@1.0.0');
  const log = require('log@1.0.0').create('graph');
  let oauth; try { oauth = require('msauth@1.0.0'); } catch {}

  const cfg = {
    tenant: null,
    clientId: null,
    clientSecret: null,
    refreshToken: null,
    scope: null,
    baseUrl: 'https://graph.microsoft.com/v1.0'
  };

  function configure(opts){
    if (!opts || typeof opts !== 'object') return;
    if (opts.tenant) cfg.tenant = String(opts.tenant);
    if (opts.clientId) cfg.clientId = String(opts.clientId);
    if (opts.clientSecret) cfg.clientSecret = String(opts.clientSecret);
    if (opts.refreshToken) cfg.refreshToken = String(opts.refreshToken);
    if (opts.scope) cfg.scope = String(opts.scope);
    if (opts.baseUrl) cfg.baseUrl = (''+opts.baseUrl).replace(/\/$/, '');
  }

  async function ensureAccessToken(over){
    // Prefer overrides, then cfg, then env
    const directToken =
      (over && over.accessToken) ||
      sys.env.get('ms.accessToken') ||
      sys.env.get('msAccessToken') ||
      sys.env.get('graph.accessToken');
    if (directToken) return directToken;
    const tenant = (over && over.tenant) || cfg.tenant || sys.env.get('ms.tenant') || sys.env.get('msTenant') || sys.env.get('graph.tenant') || 'common';
    const clientId = (over && over.clientId) || cfg.clientId || sys.env.get('ms.clientId') || sys.env.get('msClientId') || sys.env.get('graph.clientId') || null;
    const clientSecret = (over && over.clientSecret) || cfg.clientSecret || sys.env.get('ms.clientSecret') || sys.env.get('msClientSecret') || sys.env.get('graph.clientSecret') || null;
    const refreshToken = (over && over.refreshToken) || cfg.refreshToken || sys.env.get('ms.refreshToken') || sys.env.get('msRefreshToken') || sys.env.get('graph.refreshToken') || null;
    const scopeR = (over && over.scope) || cfg.scope || sys.env.get('ms.scope') || sys.env.get('msScope') || sys.env.get('graph.scope') || 'offline_access Files.ReadWrite.All';
    if (!oauth) return '';
    // Try refresh token first
    if (refreshToken && clientId){
      try { const t = await oauth.refresh({ tenant, clientId, refresh_token: refreshToken, scope: scopeR }); if (t && t.access_token) return t.access_token; } catch {}
    }
    // Client credentials
    if (clientId && clientSecret){
      try { const t = await oauth.clientCredentialsToken({ tenant, clientId, clientSecret, scope: 'https://graph.microsoft.com/.default' }); if (t && t.access_token) return t.access_token; } catch {}
    }
    // As a last resort, try stored tokens (if oauth module or caller put one in env)
    const envToken = sys.env.get('ms.accessToken') || sys.env.get('msAccessToken') || sys.env.get('graph.accessToken');
    if (envToken) return envToken;
    return '';
  }

  function joinUrl(base, path){
    if (!path) return base;
    if (/^https?:\/\//i.test(path)) return path;
    const b = (base||'').replace(/\/$/, '');
    const p = (''+path).replace(/^\//, '');
    return b + '/' + p;
  }

  async function json({ path, method='GET', headers, bodyObj, debug, auth } = {}){
    try {
      const url = joinUrl(cfg.baseUrl, path || '');
      const hdr = (headers && typeof headers==='object') ? Object.assign({}, headers) : {};
      if (!('Authorization' in hdr)){
        const tok = await ensureAccessToken(auth);
        if (!tok) return { ok:false, error:'graph: no access token', status: 0 };
        hdr['Authorization'] = 'Bearer ' + tok;
      }
      if (!hdr['Content-Type']) hdr['Content-Type'] = 'application/json';
      const r = await httpx.json({ url, method, headers: hdr, bodyObj, debug: !!debug });
      const status = r && r.status;
      const ok = r && (typeof r.ok === 'boolean') ? r.ok : !(status >= 400);
      const data = (r && (r.json || r.raw)) || null;
      const errMsg = (!ok && r && r.json && r.json.error && r.json.error.message)
        ? r.json.error.message
        : (!ok ? (r && r.raw) : undefined);
      if (!ok) {
        const raw = (r && r.raw) ? String(r.raw) : '';
        const detail = raw && raw.length > 2000 ? raw.slice(0, 2000) + '…' : raw;
        const msg = `status=${status || 0} error=${errMsg || ''} raw=${detail || ''}`;
        log.error('graph:error', msg.trim());
      }
      return { ok, data, status, raw: r && r.raw, error: errMsg };
    } catch (e){
      log.error('json:error', (e && (e.message||e)) || 'unknown');
      return { ok:false, error: (e && (e.message||String(e))) || 'unknown', status: 0 };
    }
  }

  module.exports = { configure, ensureAccessToken, json };
})();
