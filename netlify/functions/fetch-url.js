const fetch = require('node-fetch');

exports.handler = async (event) => {
  if (event.httpMethod !== 'POST') {
    return { statusCode: 405, body: 'Method not allowed' };
  }

  let url;
  try {
    ({ url } = JSON.parse(event.body));
  } catch {
    return { statusCode: 400, body: JSON.stringify({ error: 'Invalid JSON' }) };
  }

  if (!url || !url.startsWith('http')) {
    return { statusCode: 400, body: JSON.stringify({ error: 'Invalid URL' }) };
  }

  try {
    const res = await fetch(url, {
      headers: { 'User-Agent': 'Mozilla/5.0 (compatible; Sourcepoint/1.0)' },
      timeout: 8000,
    });

    const contentType = res.headers.get('content-type') || '';
    if (!contentType.includes('text')) {
      return {
        statusCode: 200,
        body: JSON.stringify({ success: false, error: 'Non-text content type' }),
      };
    }

    const html = await res.text();
    // Strip HTML tags and collapse whitespace
    const text = html
      .replace(/<script[\s\S]*?<\/script>/gi, '')
      .replace(/<style[\s\S]*?<\/style>/gi, '')
      .replace(/<[^>]+>/g, ' ')
      .replace(/\s+/g, ' ')
      .trim()
      .slice(0, 8000);

    // Extract title
    const titleMatch = html.match(/<title[^>]*>([^<]*)<\/title>/i);
    const title = titleMatch ? titleMatch[1].trim() : '';

    return {
      statusCode: 200,
      body: JSON.stringify({ success: true, text, title }),
    };
  } catch (err) {
    return {
      statusCode: 200,
      body: JSON.stringify({ success: false, error: err.message }),
    };
  }
};
