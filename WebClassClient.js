/**
 * WebClassClient.gs
 * WebClass認証・通信クラス
 */

class WebClassClient {
  constructor() {
    this.cookies = {};
  }

  login(userid, password) {
    log('--- WebClassログイン開始 ---');
    this.cookies = {};

    // 1. SSO
    const ssoRes = this._fetch(SSO_URL, { method: 'post' });
    const authId = JSON.parse(ssoRes.getContentText()).authId;

    const authPayload = {
      'authId': authId,
      'callbacks': [
        { 'type': 'NameCallback', 'output': [{ 'name': 'prompt', 'value': 'ユーザー名:' }], 'input': [{ 'name': 'IDToken1', 'value': userid }] },
        { 'type': 'PasswordCallback', 'output': [{ 'name': 'prompt', 'value': 'パスワード:' }], 'input': [{ 'name': 'IDToken2', 'value': password }], 'echoPassword': false }
      ]
    };

    const loginRes = this._fetch(SSO_URL, {
      method: 'post', contentType: 'application/json', payload: JSON.stringify(authPayload)
    });

    const authResponse = JSON.parse(loginRes.getContentText());
    if (!authResponse.tokenId && !authResponse.successUrl) {
      throw new Error('認証失敗: パスワードを確認してください。');
    }
    if (authResponse.tokenId) this.cookies['iPlanetDirectoryPro'] = authResponse.tokenId;

    // 2. SAML Manual Follow
    const loginUrl = WEBCLASS_BASE_URL + '/webclass/login.php?auth_mode=SAML';
    const ssoBase = (SSO_URL.match(/^https?:\/\/[^\/]+/) || [])[0] || '';
    const wcBase = WEBCLASS_BASE_URL.replace(/\/$/, '');

    const result = this._followRedirects(loginUrl, ssoBase, wcBase);
    log(`✅ WebClassセッション確立: ${result.finalUrl}`);
    return result.finalUrl;
  }

  fetchWithSession(url, count = 0) {
    if (count >= MAX_REDIRECTS) throw new Error('リダイレクトループ');
    const res = this._fetch(url, { method: 'get', followRedirects: false });
    const code = res.getResponseCode();
    const html = res.getContentText();

    if (code === 301 || code === 302) {
      return this.fetchWithSession(res.getHeaders()['Location'], count + 1);
    }
    if (code >= 200 && code < 300) {
      const match = html.match(REGEX.REDIRECT);
      if (match && html.length < 500) {
        let path = match[1];
        const base = WEBCLASS_BASE_URL.replace(/\/$/, '');
        const next = base + (path.startsWith('/') ? path : '/' + path);
        return this.fetchWithSession(next, count + 1);
      }
    }
    return html;
  }

  _fetch(url, options = {}) {
    const cleanUrl = this._decode(url);
    const headers = this._headers(cleanUrl);
    const opts = { 'headers': headers, 'muteHttpExceptions': true, 'followRedirects': options.followRedirects === true, ...options };

    const res = UrlFetchApp.fetch(cleanUrl, opts);
    this._updateCookies(res);
    return res;
  }

  _headers(url) {
    const ua = USER_AGENTS[Math.floor(Math.random() * USER_AGENTS.length)];
    const headers = { 'User-Agent': ua, 'Referer': url };
    const cookieStr = Object.keys(this.cookies).map(k => `${k}=${this.cookies[k]}`).join('; ');
    if (cookieStr) headers['Cookie'] = cookieStr;
    return headers;
  }

  _updateCookies(res) {
    const h = res.getAllHeaders();
    if (!h['Set-Cookie']) return;
    const cs = Array.isArray(h['Set-Cookie']) ? h['Set-Cookie'] : [h['Set-Cookie']];
    cs.forEach(c => {
      const parts = c.split(';');
      if (parts[0].includes('=')) {
        const [k, v] = parts[0].split('=').map(s => s.trim());
        this.cookies[k] = v;
      }
    });
  }

  _followRedirects(startUrl, ssoBase, wcBase) {
    let current = startUrl;
    let postData = null;
    let acsUrl = '';

    for (let i = 0; i < MAX_REDIRECTS; i++) {
      let res;
      if (postData) {
        res = this._fetch(acsUrl, { method: 'post', contentType: 'application/x-www-form-urlencoded', payload: postData, followRedirects: false });
        postData = null;
      } else {
        res = this._fetch(current, { method: 'get', followRedirects: false });
      }

      const body = res.getContentText();
      const code = res.getResponseCode();

      if (code === 200 && (body.includes('コースリスト') || body.includes('cl-courseList_courseLink'))) {
        return { response: res, finalUrl: current };
      }

      const samlMatch = body.match(REGEX.SAML_RESPONSE);
      if (samlMatch) {
        const saml = this._cleanBase64(samlMatch[1]);
        const relay = this._cleanRelay((body.match(REGEX.RELAY_STATE) || [])[1] || '');
        const action = body.match(REGEX.FORM_ACTION);
        acsUrl = this._correctUrl(this._decode(action ? action[1] : ACS_URL), wcBase);

        postData = `SAMLResponse=${encodeURIComponent(saml)}&RelayState=${encodeURIComponent(relay)}`;
        current = res.getHeaders()['Location'] || current;
        continue;
      }

      let loc = res.getHeaders()['Location'];
      if (loc) {
        loc = this._stripPort(this._decode(loc));
        const base = loc.includes(wcBase) ? wcBase : ssoBase;
        current = this._correctUrl(loc, base);
        continue;
      }

      const js = body.match(REGEX.REDIRECT);
      if (js) {
        let path = js[1];
        if (!path.startsWith('http')) path = this._correctUrl(path, wcBase);
        current = path;
        continue;
      }

      if (i >= 5 && code === 200) return { response: res, finalUrl: current };
      throw new Error(`リダイレクト追跡失敗: ${code} ${current}`);
    }
    throw new Error('リダイレクト回数超過');
  }

  _decode(t) { return t ? t.replace(/&#x3a;/g, ':').replace(/&#x2f;/g, '/').replace(/&amp;/g, '&') : t; }
  _stripPort(u) { return (u && u.includes(':443/')) ? u.replace(':443/', '/') : u; }
  _correctUrl(u, b) { if (!u || u.startsWith('http')) return u; const bc = b.endsWith('/') ? b.slice(0, -1) : b; return bc + (u.startsWith('/') ? '' : '/') + u; }
  _cleanBase64(s) { if (!s) return ''; let c = s.replace(/&#x([0-9a-fA-F]+);/g, (_, h) => String.fromCharCode(parseInt(h, 16))).replace(/&amp;/g, '&').replace(/[^A-Za-z0-9+/=]/g, ''); const p = (4 - (c.replace(/=/g, '').length % 4)) % 4; return c.replace(/=/g, '') + '='.repeat(p); }
  _cleanRelay(s) { return s ? s.replace(/&#x([0-9a-fA-F]+);/g, (_, h) => String.fromCharCode(parseInt(h, 16))).replace(/[\r\n]/g, '').trim() : ''; }
}