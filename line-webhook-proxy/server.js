const http = require('http');
const https = require('https');
const crypto = require('crypto');
const fs = require('fs');
const path = require('path');
const { URL } = require('url');

loadEnvFile();

const PORT = Number(process.env.PORT || 8787);
const LINE_CHANNEL_SECRET = String(process.env.LINE_CHANNEL_SECRET || '').trim();
const LINE_CHANNEL_ACCESS_TOKEN = String(process.env.LINE_CHANNEL_ACCESS_TOKEN || '').trim();
const LINE_DEFAULT_USER_ID = String(process.env.LINE_DEFAULT_USER_ID || '').trim();
const GAS_BASE_URL = String(process.env.GAS_BASE_URL || '').trim();
const APP_TIMEZONE = String(process.env.APP_TIMEZONE || 'Asia/Taipei').trim();
const WEEKDAY_NAMES = ['日', '一', '二', '三', '四', '五', '六'];
const LINE_HOMEROOM = ['國語', '數學', '社會', '健康', '樂活'];
const LINE_SUBJECT = ['自然', '藝專', '閩語', '視覺', '英語', '分部課', '樂理'];

if (!LINE_CHANNEL_SECRET) console.warn('Missing LINE_CHANNEL_SECRET');
if (!LINE_CHANNEL_ACCESS_TOKEN) console.warn('Missing LINE_CHANNEL_ACCESS_TOKEN');
if (!GAS_BASE_URL) console.warn('Missing GAS_BASE_URL');

const server = http.createServer(async (req, res) => {
  try {
    const url = new URL(req.url, `http://${req.headers.host || 'localhost'}`);

    if (req.method === 'GET' && url.pathname === '/') {
      return sendText(res, 200, 'LINE webhook proxy is running');
    }

    if (req.method === 'GET' && url.pathname === '/health') {
      return sendJson(res, 200, { ok: true, service: 'line-webhook-proxy' });
    }

    if (req.method === 'POST' && url.pathname === '/webhook') {
      const rawBody = await readRawBody(req);
      if (!verifyLineSignature(req.headers['x-line-signature'], rawBody, LINE_CHANNEL_SECRET)) {
        return sendJson(res, 401, { ok: false, error: 'Invalid LINE signature' });
      }

      let payload = {};
      try {
        payload = rawBody ? JSON.parse(rawBody) : {};
      } catch (error) {
        return sendJson(res, 400, { ok: false, error: 'Invalid JSON body' });
      }

      sendText(res, 200, 'OK');
      setImmediate(() => {
        processWebhookPayload(payload).catch((error) => {
          console.error('[processWebhookPayload]', error && error.stack ? error.stack : error);
        });
      });
      return;
    }

    return sendJson(res, 404, { ok: false, error: 'Not found' });
  } catch (error) {
    console.error('[server]', error && error.stack ? error.stack : error);
    if (!res.headersSent) {
      return sendJson(res, 500, { ok: false, error: 'Internal server error' });
    }
  }
});

server.listen(PORT, () => {
  console.log(`LINE webhook proxy listening on http://localhost:${PORT}`);
});

async function processWebhookPayload(payload) {
  const events = Array.isArray(payload && payload.events) ? payload.events : [];
  if (!events.length) return;

  for (const event of events) {
    if (!event || event.type !== 'message' || !event.message || event.message.type !== 'text') {
      continue;
    }

    const dateStr = parseQueryDate(event.message.text);
    if (!dateStr) continue;

    const replyToken = event.replyToken ? String(event.replyToken).trim() : '';
    const userId = event.source && event.source.userId ? String(event.source.userId).trim() : LINE_DEFAULT_USER_ID;

    try {
      const digest = await fetchLineDigest(dateStr);
      const flex = buildFlexMessage(digest);
      await respondToLineEvent(replyToken, userId, flex);
      console.log(`[event ${dateStr}] delivered`);
    } catch (error) {
      console.error(`[event ${dateStr}]`, error && error.stack ? error.stack : error);
      const fallbackMessage = {
        type: 'text',
        text: `查詢 ${dateStr} 失敗：${error.message || error}`
      };
      await respondToLineEvent(replyToken, userId, fallbackMessage).catch((deliveryError) => {
        console.error('[delivery fallback]', deliveryError && deliveryError.stack ? deliveryError.stack : deliveryError);
      });
    }
  }
}

function parseQueryDate(text) {
  const value = String(text || '').trim();
  const today = getTodayInTimeZone();

  if (value === '今天' || value === '今日') return formatMdDate(today);
  if (value === '明天') {
    const next = new Date(today);
    next.setDate(next.getDate() + 1);
    return formatMdDate(next);
  }
  if (value === '昨天') {
    const prev = new Date(today);
    prev.setDate(prev.getDate() - 1);
    return formatMdDate(prev);
  }

  const prevMatch = value.match(/^__prev__(\d{1,2})\/(\d{1,2})/);
  if (prevMatch) {
    const prev = new Date(today.getFullYear(), Number(prevMatch[1]) - 1, Number(prevMatch[2]) - 1);
    return formatMdDate(prev);
  }

  const nextMatch = value.match(/^__next__(\d{1,2})\/(\d{1,2})/);
  if (nextMatch) {
    const next = new Date(today.getFullYear(), Number(nextMatch[1]) - 1, Number(nextMatch[2]) + 1);
    return formatMdDate(next);
  }

  const mdMatch = value.match(/^(\d{1,2})[/-](\d{1,2})$/);
  if (mdMatch) {
    return `${Number(mdMatch[1])}/${Number(mdMatch[2])}`;
  }

  return null;
}

function getTodayInTimeZone() {
  const formatter = new Intl.DateTimeFormat('en-CA', {
    timeZone: APP_TIMEZONE,
    year: 'numeric',
    month: '2-digit',
    day: '2-digit'
  });
  const parts = formatter.formatToParts(new Date());
  const map = {};
  for (const part of parts) {
    if (part.type !== 'literal') {
      map[part.type] = part.value;
    }
  }
  return new Date(Number(map.year), Number(map.month) - 1, Number(map.day));
}

function formatMdDate(dateObj) {
  return `${dateObj.getMonth() + 1}/${dateObj.getDate()}`;
}

async function fetchLineDigest(dateStr) {
  const gasUrl = new URL(GAS_BASE_URL);
  gasUrl.searchParams.set('action', 'getLineDigest');
  gasUrl.searchParams.set('date', dateStr);

  const response = await requestJson(gasUrl, { method: 'GET' });
  if (!response.ok) {
    throw new Error(response.error || 'GAS returned an error');
  }
  return response;
}

function buildFlexMessage(digest) {
  const schedule = digest.schedule || {};
  const periods = Array.isArray(schedule.periods) ? schedule.periods : [];
  const weeklyNotes = Array.isArray(digest.weeklyNotes) ? digest.weeklyNotes : [];
  const contactText = String(digest.contact || '').trim();
  const headerText = `📅 ${digest.dayLabel || digest.date || ''}${digest.todayKey ? `　${digest.todayKey}` : ''}`;

  const tableRows = [
    {
      type: 'box',
      layout: 'horizontal',
      backgroundColor: '#dce8fb',
      paddingAll: '8px',
      contents: [
        { type: 'text', text: '節', size: 'lg', flex: 1, align: 'center', weight: 'bold', color: '#2a4070' },
        { type: 'text', text: '進度', size: 'lg', flex: 5, weight: 'bold', color: '#2a4070' },
        { type: 'text', text: '備註', size: 'lg', flex: 4, weight: 'bold', color: '#2a4070' }
      ]
    }
  ];

  if (periods.length) {
    for (const period of periods) {
      tableRows.push({
        type: 'box',
        layout: 'horizontal',
        backgroundColor: getRowColor(period),
        paddingAll: '8px',
        contents: [
          {
            type: 'text',
            text: period.period === '晨光' ? '晨' : String(period.period || ''),
            size: 'lg',
            flex: 1,
            align: 'center',
            color: '#5c6f92',
            weight: 'bold'
          },
          {
            type: 'text',
            text: String(period.content || '－'),
            size: 'lg',
            flex: 5,
            wrap: true,
            color: '#243b63'
          },
          {
            type: 'text',
            text: String(period.note || '　'),
            size: 'lg',
            flex: 4,
            wrap: true,
            color: '#50627f'
          }
        ]
      });
      tableRows.push({ type: 'separator', color: '#c8d5e8' });
    }
  } else {
    tableRows.push({
      type: 'box',
      layout: 'vertical',
      paddingAll: '8px',
      contents: [{ type: 'text', text: '（今日無課程進度）', size: 'lg', color: '#9aaabb', align: 'center' }]
    });
  }

  const footerItems = [
    { type: 'text', text: '📖 聯絡簿', weight: 'bold', size: 'xl', color: '#243b63' }
  ];

  const contactLines = contactText.split('\n').map((line) => line.trim()).filter(Boolean);
  if (contactLines.length) {
    for (const line of contactLines) {
      footerItems.push({ type: 'text', text: line, size: 'lg', color: '#50627f', wrap: true, margin: 'xs' });
    }
  } else {
    footerItems.push({ type: 'text', text: '（今日無聯絡簿內容）', size: 'lg', color: '#aaaaaa' });
  }

  if (weeklyNotes.length) {
    footerItems.push({ type: 'separator', color: '#c8d5e8', margin: 'md' });
    footerItems.push({ type: 'text', text: '📋 本週記事', weight: 'bold', size: 'xl', color: '#243b63', margin: 'md' });
    for (const note of weeklyNotes) {
      if (!note) continue;
      footerItems.push({ type: 'text', text: `• ${note}`, size: 'lg', color: '#50627f', wrap: true, margin: 'xs' });
    }
  }

  return {
    type: 'flex',
    altText: `${digest.dayLabel || digest.date || ''} 今日進度`,
    quickReply: {
      items: [
        { type: 'action', action: { type: 'message', label: '⬅ 前一天', text: `__prev__${digest.dayLabel || digest.date || ''}` } },
        { type: 'action', action: { type: 'message', label: '今天', text: '今天' } },
        { type: 'action', action: { type: 'message', label: '明天 ➡', text: `__next__${digest.dayLabel || digest.date || ''}` } }
      ]
    },
    contents: {
      type: 'bubble',
      size: 'mega',
      header: {
        type: 'box',
        layout: 'horizontal',
        backgroundColor: '#e3eeff',
        paddingAll: '12px',
        contents: [{ type: 'text', text: headerText, weight: 'bold', size: 'xl', color: '#1a3a6a', wrap: true }]
      },
      body: {
        type: 'box',
        layout: 'vertical',
        spacing: 'none',
        paddingTop: '0px',
        paddingBottom: '0px',
        paddingStart: 'lg',
        paddingEnd: 'lg',
        contents: tableRows
      },
      footer: {
        type: 'box',
        layout: 'vertical',
        spacing: 'sm',
        paddingTop: '10px',
        paddingBottom: '10px',
        paddingStart: 'lg',
        paddingEnd: 'lg',
        backgroundColor: '#f7faff',
        contents: footerItems
      }
    }
  };
}

function getRowColor(period) {
  const subject = String((period && period.subject) || '');
  const content = String((period && period.content) || '');
  const note = String((period && period.note) || '');
  const text = `${subject} ${content} ${note}`;

  if (/調課|借課|還課/.test(text)) return '#fff4a8';
  if (/(^|[^A-Za-z])X([^A-Za-z]|$)/i.test(content)) return '#d8dce4';
  if (LINE_HOMEROOM.some((keyword) => subject.includes(keyword) || content.includes(keyword))) return '#f5f0eb';
  if (LINE_SUBJECT.some((keyword) => subject.includes(keyword) || content.includes(keyword))) return '#dde6dc';
  return '#ffffff';
}

async function respondToLineEvent(replyToken, userId, message) {
  if (replyToken) {
    return replyLineMessage(replyToken, message);
  }
  if (userId) {
    return pushLineMessage(userId, message);
  }
  throw new Error('No replyToken or userId available for delivery');
}

async function replyLineMessage(replyToken, message) {
  const body = JSON.stringify({
    replyToken,
    messages: [message]
  });

  return requestJson('https://api.line.me/v2/bot/message/reply', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      Authorization: `Bearer ${LINE_CHANNEL_ACCESS_TOKEN}`
    },
    body
  });
}

async function pushLineMessage(userId, message) {
  const body = JSON.stringify({
    to: userId,
    messages: [message]
  });

  return requestJson('https://api.line.me/v2/bot/message/push', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      Authorization: `Bearer ${LINE_CHANNEL_ACCESS_TOKEN}`
    },
    body
  });
}

function verifyLineSignature(signature, rawBody, secret) {
  if (!signature || !secret) return false;
  const digest = crypto.createHmac('sha256', secret).update(rawBody).digest('base64');
  const actual = Buffer.from(String(signature));
  const expected = Buffer.from(digest);
  return actual.length === expected.length && crypto.timingSafeEqual(actual, expected);
}

function readRawBody(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    req.on('data', (chunk) => chunks.push(chunk));
    req.on('end', () => resolve(Buffer.concat(chunks).toString('utf8')));
    req.on('error', reject);
  });
}

function requestJson(targetUrl, options, redirectCount) {
  return new Promise((resolve, reject) => {
    const currentRedirectCount = Number(redirectCount || 0);
    const url = typeof targetUrl === 'string' ? new URL(targetUrl) : targetUrl;
    const client = url.protocol === 'https:' ? https : http;
    const request = client.request(url, options || {}, (response) => {
      const chunks = [];
      response.on('data', (chunk) => chunks.push(chunk));
      response.on('end', () => {
        const raw = Buffer.concat(chunks).toString('utf8');
        const statusCode = response.statusCode || 500;

        if ([301, 302, 303, 307, 308].includes(statusCode) && response.headers.location) {
          if (currentRedirectCount >= 5) {
            return reject(new Error(`HTTP ${statusCode}: too many redirects`));
          }
          const nextUrl = new URL(response.headers.location, url);
          const nextOptions = Object.assign({}, options || {});
          if (statusCode === 303) {
            nextOptions.method = 'GET';
            delete nextOptions.body;
          }
          return resolve(requestJson(nextUrl, nextOptions, currentRedirectCount + 1));
        }

        if (!raw) {
          if (statusCode >= 200 && statusCode < 300) return resolve({ ok: true });
          return reject(new Error(`HTTP ${statusCode}`));
        }

        let parsed;
        try {
          parsed = JSON.parse(raw);
        } catch (error) {
          if (statusCode >= 200 && statusCode < 300) return resolve({ ok: true, raw });
          const compactRaw = raw
            .replace(/\s+/g, ' ')
            .replace(/<[^>]+>/g, ' ')
            .trim()
            .slice(0, 180);
          return reject(new Error(`HTTP ${statusCode}: ${compactRaw || 'Non-JSON response'}`));
        }

        if (statusCode >= 200 && statusCode < 300) return resolve(parsed);
        reject(new Error(`HTTP ${statusCode}: ${parsed.error || raw}`));
      });
    });

    request.on('error', reject);
    if (options && options.body) request.write(options.body);
    request.end();
  });
}

function sendJson(res, statusCode, payload) {
  res.writeHead(statusCode, { 'Content-Type': 'application/json; charset=utf-8' });
  res.end(JSON.stringify(payload));
}

function sendText(res, statusCode, text) {
  res.writeHead(statusCode, { 'Content-Type': 'text/plain; charset=utf-8' });
  res.end(text);
}

function loadEnvFile() {
  const envPath = path.join(__dirname, '.env');
  if (!fs.existsSync(envPath)) return;

  const lines = fs.readFileSync(envPath, 'utf8').split(/\r?\n/);
  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith('#')) continue;
    const equalIndex = trimmed.indexOf('=');
    if (equalIndex === -1) continue;
    const key = trimmed.slice(0, equalIndex).trim();
    const value = trimmed.slice(equalIndex + 1).trim();
    if (!process.env[key]) {
      process.env[key] = value;
    }
  }
}
