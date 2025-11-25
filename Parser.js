/**
 * Parser.gs
 * WebClassのHTML解析ロジック
 */

const WebClassParser = {
  /**
   * ダッシュボードから科目リンクを抽出
   */
  parseDashboard(html) {
    const result = new Map();
    // タイムテーブルまたはコース一覧のリンクを抽出
    const regex = /<a href='(\/webclass\/course\.php\/[a-f0-9]+\/login\?acs_=[^']*)' Target='_top'>([\s\S]*?)<\/a>/g;
    let match;
    const baseUrl = WEBCLASS_BASE_URL.replace(/\/$/, '');

    while ((match = regex.exec(html)) !== null) {
      const relativeUrl = match[1];
      let rawName = match[2].replace(/&raquo;/, '').replace(/<[^>]+>/g, '').trim();
      const uniqueKey = relativeUrl.split('?')[0];

      if (relativeUrl && !result.has(uniqueKey)) {
        result.set(uniqueKey, {
          url: baseUrl + relativeUrl,
          name: rawName
        });
      }
    }
    return Array.from(result.values());
  },

  /**
   * コースページから課題リストを抽出
   */
  parseCourseContents(html) {
    const results = [];
    const itemsMatch = html.match(/<section class=\"list-group-item cl-contentsList_listGroupItem\"[\s\S]*?<\/section>/g);
    if (!itemsMatch) return results;

    itemsMatch.forEach(item => {
      // タイトル抽出
      const titleMatch = item.match(/<h4\s+[^>]*?class=\"cm-contentsList_contentName\"[^>]*?>([\s\S]*?)<\/h4>/);
      if (!titleMatch) return;

      let content = titleMatch[1].replace(/<div class=\"cl-contentsList_new\">[\s\S]*?<\/div>/g, '').trim();
      const title = content.replace(/<[^>]+>/g, '').trim();

      // リンク抽出
      const urlMatch = content.match(/<a href=\"([^\"]+)\">/);
      let shareLink = "";
      let rawEnd = "";
      if (urlMatch) {
        const idMatch = urlMatch[1].match(REGEX.ID);
        if (idMatch) {
          shareLink = `${WEBCLASS_BASE_URL}/webclass/login.php?id=${idMatch[1]}&page=1&auth_mode=SAML`;
        }
      }

      // 期間抽出
      let start = "";
      const periodMatch = item.match(/利用可能期間<\/div>\s*<div\s+[^>]*?class=['"]cm-contentsList_contentDetailListItemData['"][^>]*?>\s*([\s\S]*?)\s*<\/div>/);
      if (periodMatch) {
        const rawPeriod = periodMatch[1].replace(/<[^>]+>/g, '').trim();
        const parts = rawPeriod.split(' - ');
        if (parts.length === 2) {
          start = parts[0].trim();
          rawEnd = parts[1].trim();
        } else {
          rawEnd = rawPeriod; // 終了日のみの場合
        }
      }

      if (title && shareLink) {
        results.push({ title, shareLink, start, end: rawEnd });
      }
    });
    return results;
  }
};