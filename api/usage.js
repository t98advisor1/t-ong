const KEY_TOTAL = 'notetong:total-visits';
const KEY_DAY_PREFIX = 'notetong:daily-visitors:';

function credentials() {
  return {
    url: process.env.UPSTASH_REDIS_REST_URL || process.env.KV_REST_API_URL,
    token: process.env.UPSTASH_REDIS_REST_TOKEN || process.env.KV_REST_API_TOKEN
  };
}

function koreaDateKey() {
  const parts = new Intl.DateTimeFormat('en-CA', {
    timeZone: 'Asia/Seoul',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit'
  }).formatToParts(new Date());
  const map = {};
  for (const p of parts) map[p.type] = p.value;
  return map.year + '-' + map.month + '-' + map.day;
}

async function pipeline(commands) {
  const config = credentials();
  if (!config.url || !config.token) throw new Error('STATS_STORAGE_NOT_CONFIGURED');
  const response = await fetch(config.url.replace(/\/$/, '') + '/pipeline', {
    method: 'POST',
    headers: {
      Authorization: 'Bearer ' + config.token,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify(commands)
  });
  if (!response.ok) throw new Error('UPSTASH_' + response.status);
  return response.json();
}

function setCors(req, res) {
  const origin = (req.headers && req.headers.origin) || '*';
  res.setHeader('Access-Control-Allow-Origin', origin);
  res.setHeader('Vary', 'Origin');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  res.setHeader('Access-Control-Max-Age', '86400');
}

function send(res, status, body) {
  res.setHeader('Cache-Control', 'no-store, max-age=0');
  res.setHeader('Content-Type', 'application/json; charset=utf-8');
  res.status(status).json(body);
}

function parseBody(req) {
  if (!req || !req.body) return {};
  if (typeof req.body === 'string') {
    try { return JSON.parse(req.body); } catch (e) { return {}; }
  }
  return req.body;
}

module.exports = async function handler(req, res) {
  setCors(req, res);
  if (req.method === 'OPTIONS') {
    return res.status(204).end();
  }
  if (req.method !== 'GET' && req.method !== 'POST') {
    res.setHeader('Allow', 'GET, POST, OPTIONS');
    return send(res, 405, { error: 'METHOD_NOT_ALLOWED' });
  }
  try {
    const bodyObj = parseBody(req);
    // scope=smartboard 이면 스마트 판서 전용 카운터를 씁니다.
    const rawScope = String((bodyObj && bodyObj.scope) || (req.query && req.query.scope) || '');
    const scope = rawScope === 'smartboard' ? 'smartboard' : '';
    const prefix = scope ? 'notetong:' + scope + ':' : 'notetong:';
    const totalKey = scope ? prefix + 'total-visits' : KEY_TOTAL;
    const dayKey = (scope ? prefix + 'daily-visitors:' : KEY_DAY_PREFIX) + koreaDateKey();
    let commands;
    if (req.method === 'POST') {
      const token = String((bodyObj && bodyObj.visitorToken) || '');
      if (!/^[a-f0-9]{64}$/.test(token)) return send(res, 400, { error: 'INVALID_VISITOR_TOKEN' });
      commands = [['SADD', dayKey, token], ['EXPIRE', dayKey, 259200], ['INCR', totalKey], ['SCARD', dayKey]];
      const result = await pipeline(commands);
      return send(res, 200, {
        todayVisitors: Number(result[3] && result[3].result) || 0,
        totalVisits: Number(result[2] && result[2].result) || 0
      });
    }
    commands = [['SCARD', dayKey], ['GET', totalKey]];
    const result = await pipeline(commands);
    return send(res, 200, {
      todayVisitors: Number(result[0] && result[0].result) || 0,
      totalVisits: Number(result[1] && result[1].result) || 0
    });
  } catch (error) {
    const missing = error && error.message === 'STATS_STORAGE_NOT_CONFIGURED';
    return send(res, missing ? 503 : 500, { error: missing ? 'STATS_STORAGE_NOT_CONFIGURED' : 'STATS_UNAVAILABLE' });
  }
};
