/**
 * Config.gs
 * システム全体で使用する共通定数。
 */

// --- WebClass関連 URL ---
const WEBCLASS_BASE_URL = 'https://rpwebcls.meijo-u.ac.jp';
const SSO_URL = 'https://slbsso.meijo-u.ac.jp/opensso/json/authenticate';
const ACS_URL = WEBCLASS_BASE_URL + '/simplesaml/module.php/saml/sp/saml2-acs.php/default-sp';

// --- スプレッドシート設定 ---
const SHEET_NAME_WEBCLASS = 'WebClass課題';
const SHEET_NAME_CLASSROOM = 'Classroom課題';
const SHEET_NAME_LOG = 'ログ';
const HEADER = ['ソース', '授業名', '課題タイトル', '開始日時', '終了日時', '課題リンク (URL)', 'Tasks ID', '登録済みフラグ'];

// --- システム定数 ---
const MAX_REDIRECTS = 15;

// --- 正規表現 ---
const REGEX = {
  ID: /id=([a-f0-9]+)/,
  REDIRECT: /(?:window\.location\.href\s*=\s*|content\s*=\s*[\"']0;\s*URL=)['"]([^\"']+)[\"']/,
  COURSE_LINK: /\/webclass\/course\.php\/[a-f0-9]+/,
  SAML_RESPONSE: /<input type="hidden" name="SAMLResponse" value="([^"]+)"/,
  RELAY_STATE: /<input type="hidden" name="RelayState" value="([^"]+)"/,
  FORM_ACTION: /<form method="post" action="([^"]+)"/
};

// --- User-Agent リスト ---
const USER_AGENTS = [
  'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/129.0.0.0 Safari/537.36',
  'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/129.0.0.0 Safari/537.36',
  'Mozilla/5.0 (iPhone; CPU iPhone OS 17_6 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.6 Mobile/15E148 Safari/604.1'
];