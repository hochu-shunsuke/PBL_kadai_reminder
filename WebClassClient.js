/**
 * WebClassClient.gs
 * WebClassとの通信、認証、セッション管理を行うクラス
 */

class WebClassClient {
  constructor() {
    this.cookies = {}; // セッションCookieをこのインスタンスで保持
  }

  /**
   * ログイン処理を実行し、ダッシュボードへのアクセスを確立する
   */
  login(userid, password) {
    log('--- WebClassログイン処理開始 ---');
    this.cookies = {}; 

    // 1. SSO認証
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
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(authPayload)
    });
    
    const authResponse = JSON.parse(loginRes.getContentText());
    if (!authResponse.tokenId && !authResponse.successUrl) {
      throw new Error('SSO認証失敗: ' + (authResponse.message || '不明なエラー'));
    }

    if (authResponse.tokenId) {
      this.cookies['iPlanetDirectoryPro'] = authResponse.tokenId;
    }
    log('SSO認証成功。WebClassへの接続を開始します。');

    // 2. SAMLフローの手動追跡
    const initialLoginUrl = WEBCLASS_BASE_URL + '/webclass/login.php?auth_mode=SAML';
    const ssoBaseUri = (SSO_URL.match(/^https?:\/\/[^\/]+/) || [])[0] || '';
    const webClassBaseUri = WEBCLASS_BASE_URL.replace(/\/$/, '');

    const finalResult = this._followManualRedirects(initialLoginUrl, ssoBaseUri, webClassBaseUri);
    log(`✅ WebClassセッション確立成功！最終URL: ${finalResult.finalUrl}`);
    
    return finalResult.finalUrl;
  }

  /**
   * セッションを維持してURLを取得（リダイレクト対応）
   */
  fetchWithSession(url, redirectCount = 0) {
    if (redirectCount >= MAX_REDIRECTS) throw new Error('リダイレクトループ検出');

    const res = this._fetch(url, { method: 'get', followRedirects: false });
    const code = res.getResponseCode();
    const content = res.getContentText();

    // HTTPリダイレクト
    if (code === 301 || code === 302) {
      const nextUrl = res.getHeaders()['Location'];
      return this.fetchWithSession(nextUrl, redirectCount + 1);
    }

    // JSリダイレクト
    if (code >= 200 && code < 300) {
      const match = content.match(REGEX.REDIRECT);
      if (match && content.length < 500) {
        let nextPath = match[1];
        const baseUrl = WEBCLASS_BASE_URL.replace(/\/$/, '');
        const nextUrl = baseUrl + (nextPath.startsWith('/') ? nextPath : '/' + nextPath);
        return this.fetchWithSession(nextUrl, redirectCount + 1);
      }
    }

    return content;
  }

  // --- 内部ヘルパーメソッド ---

  _fetch(url, options = {}) {
    const cleanUrl = this._decodeHtmlEntities(url);
    const headers = this._buildHeaders(cleanUrl);
    
    const defaults = {
      'headers': headers,
      'muteHttpExceptions': true,
      'followRedirects': options.followRedirects === true
    };
    const finalOpts = { ...defaults, ...options };

    try {
      const res = UrlFetchApp.fetch(cleanUrl, finalOpts);
      this._updateCookies(res);
      return res;
    } catch (e) {
      log(`[Error] Fetch failed: ${cleanUrl} - ${e.message}`);
      throw e;
    }
  }

  _buildHeaders(url) {
    const ua = USER_AGENTS[Math.floor(Math.random() * USER_AGENTS.length)];
    const headers = {
      'User-Agent': ua,
      'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
      'Referer': url
    };
    const cookieStr = Object.keys(this.cookies)
      .map(k => `${k}=${this.cookies[k]}`).join('; ');
    if (cookieStr) headers['Cookie'] = cookieStr;
    return headers;
  }

  _updateCookies(res) {
    const headers = res.getAllHeaders();
    if (!headers['Set-Cookie']) return;
    
    const cookies = Array.isArray(headers['Set-Cookie']) ? headers['Set-Cookie'] : [headers['Set-Cookie']];
    cookies.forEach(c => {
      const parts = c.split(';');
      if (parts[0].includes('=')) {
        const [k, v] = parts[0].split('=').map(s => s.trim());
        this.cookies[k] = v;
      }
    });
  }

  /**
   * 複雑なSAMLリダイレクトを手動で処理する
   */
  _followManualRedirects(startUrl, ssoBaseUri, webclassBaseUri) {
    let currentUrl = startUrl;
    let samlPostData = null;
    let samlAcsUrl = '';

    for (let i = 0; i < MAX_REDIRECTS; i++) {
      let res;
      if (samlPostData) {
        log(`SAML POST実行...`);
        res = this._fetch(samlAcsUrl, {
          method: 'post',
          contentType: 'application/x-www-form-urlencoded',
          payload: samlPostData,
          followRedirects: false
        });
        samlPostData = null;
      } else {
        res = this._fetch(currentUrl, { method: 'get', followRedirects: false });
      }

      const body = res.getContentText();
      const code = res.getResponseCode();

      // ダッシュボード到達判定
      if (code === 200 && (body.includes('コースリスト') || body.includes('cl-courseList_courseLink'))) {
        return { response: res, finalUrl: currentUrl };
      }

      // SAML Response POST判定
      const samlMatch = body.match(REGEX.SAML_RESPONSE);
      if (samlMatch) {
        const samlResponse = this._cleanBase64(samlMatch[1]);
        const relayState = this._cleanRelayState((body.match(REGEX.RELAY_STATE) || [])[1] || '');
        const actionMatch = body.match(REGEX.FORM_ACTION);
        let acsUrl = actionMatch ? actionMatch[1] : ACS_URL;
        
        samlAcsUrl = this._correctUrl(this._decodeHtmlEntities(acsUrl), webclassBaseUri);
        samlPostData = `SAMLResponse=${encodeURIComponent(samlResponse)}&RelayState=${encodeURIComponent(relayState)}`;
        currentUrl = res.getHeaders()['Location'] || currentUrl; // ロケーションがなければ現在地維持で次へ
        continue;
      }

      // 通常リダイレクト
      let nextLoc = res.getHeaders()['Location'];
      if (nextLoc) {
        nextLoc = this._stripPort443(this._decodeHtmlEntities(nextLoc));
        const base = nextLoc.includes(webclassBaseUri) ? webclassBaseUri : ssoBaseUri;
        currentUrl = this._correctUrl(nextLoc, base);
        continue;
      }
      
      // JavaScriptリダイレクト
      const jsMatch = body.match(REGEX.REDIRECT);
      if (jsMatch) {
        let path = jsMatch[1];
        if(!path.startsWith('http')) path = this._correctUrl(path, webclassBaseUri);
        currentUrl = path;
        continue;
      }

      // ここまで来て進展がなければ完了とみなす（安全策）
      if (i >= 5 && code === 200) return { response: res, finalUrl: currentUrl };

      throw new Error(`リダイレクト追跡不能: Status ${code}, URL ${currentUrl}`);
    }
    throw new Error('リダイレクト回数超過');
  }

  // --- 文字列処理ヘルパー (SAML特有の処理) ---

  _decodeHtmlEntities(text) {
    if (!text) return text;
    return text.replace(/&#x3a;/g, ':').replace(/&#x2f;/g, '/')
      .replace(/&amp;/g, '&').replace(/&quot;/g, '"')
      .replace(/&gt;/g, '>').replace(/&lt;/g, '<');
  }

  _stripPort443(url) {
    return (url && url.includes(':443/')) ? url.replace(':443/', '/') : url;
  }

  _correctUrl(url, base) {
    if (!url || url.startsWith('http')) return url;
    const baseClean = base.endsWith('/') ? base.slice(0, -1) : base;
    return baseClean + (url.startsWith('/') ? '' : '/') + url;
  }

  _cleanBase64(str) {
    if (!str) return '';
    let c = str.replace(/&#x([0-9a-fA-F]+);/g, (_, h) => String.fromCharCode(parseInt(h, 16)))
               .replace(/&#(\d+);/g, (_, d) => String.fromCharCode(parseInt(d, 10)))
               .replace(/&amp;/g, '&').replace(/&quot;/g, '"').replace(/&lt;/g, '<').replace(/&gt;/g, '>')
               .replace(/[\x00-\x1F\x7F-\x9F\s]/g, '')
               .replace(/[^A-Za-z0-9+/=]/g, '');
    const pad = (4 - (c.replace(/=/g, '').length % 4)) % 4;
    return c.replace(/=/g, '') + '='.repeat(pad);
  }

  _cleanRelayState(str) {
    if (!str) return '';
    return str.replace(/&#x([0-9a-fA-F]+);/g, (_, h) => String.fromCharCode(parseInt(h, 16)))
              .replace(/&#(\d+);/g, (_, d) => String.fromCharCode(parseInt(d, 10)))
              .replace(/&amp;/g, '&').replace(/&quot;/g, '"').replace(/&lt;/g, '<').replace(/&gt;/g, '>')
              .replace(/[\r\n]/g, '').trim();
  }
}