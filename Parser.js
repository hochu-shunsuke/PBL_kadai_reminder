/**
 * Parser.gs
 */
const WebClassParser = {
  parseDashboard(html) {
    const result = new Map();
    const regex = /<a href='(\/webclass\/course\.php\/[a-f0-9]+\/login\?acs_=[^']*)' Target='_top'>([\s\S]*?)<\/a>/g;
    let match;
    const baseUrl = WEBCLASS_BASE_URL.replace(/\/$/, '');
    while ((match = regex.exec(html)) !== null) {
      const rel = match[1];
      const name = match[2].replace(/&raquo;/, '').replace(/<[^>]+>/g, '').trim();
      const key = rel.split('?')[0];
      if (!result.has(key)) result.set(key, { url: baseUrl + rel, name: name });
    }
    return Array.from(result.values());
  },

  parseCourseContents(html) {
    const results = [];
    const items = html.match(/<section class=\"list-group-item cl-contentsList_listGroupItem\"[\s\S]*?<\/section>/g);
    if (!items) return results;

    items.forEach(item => {
      const tMatch = item.match(/<h4\s+[^>]*?class=\"cm-contentsList_contentName\"[^>]*?>([\s\S]*?)<\/h4>/);
      if (!tMatch) return;
      const content = tMatch[1].replace(/<div class=\"cl-contentsList_new\">[\s\S]*?<\/div>/g, '').trim();
      const title = content.replace(/<[^>]+>/g, '').trim();

      const uMatch = content.match(/<a href=\"([^\"]+)\">/);
      let link = "";
      if (uMatch) {
        const id = uMatch[1].match(REGEX.ID);
        if (id) link = `${WEBCLASS_BASE_URL}/webclass/login.php?id=${id[1]}&page=1&auth_mode=SAML`;
      }

      let start = "", end = "";
      const pMatch = item.match(/利用可能期間<\/div>\s*<div\s+[^>]*?class=['"]cm-contentsList_contentDetailListItemData['"][^>]*?>\s*([\s\S]*?)\s*<\/div>/);
      if (pMatch) {
        const raw = pMatch[1].replace(/<[^>]+>/g, '').trim();
        const parts = raw.split(' - ');
        if (parts.length === 2) { start = parts[0].trim(); end = parts[1].trim(); }
        else { end = raw; }
      }

      if (title && link) results.push({ title, shareLink: link, start, end });
    });
    return results;
  }
};