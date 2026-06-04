import { neon } from '@netlify/neon';

// Allowed origins for CORS. Add your production URL(s) here as a comma-separated list.
const ALLOWED_ORIGINS = (process.env.ALLOWED_ORIGINS || '')
  .split(',')
  .map((s) => s.trim())
  .filter(Boolean);

// Optional shared secret. If set, clients must send `x-admin-token: <value>`.
// Recommended in production because scan images contain student PII.
const ADMIN_TOKEN = process.env.ADMIN_TOKEN || '';

// Hard cap on the number of rows returned in one request.
const MAX_ROWS = 1000;

const corsHeaders = (origin) => {
  const headers = {
    'Access-Control-Allow-Methods': 'GET, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, x-admin-token',
    'Access-Control-Max-Age': '86400',
    'Vary': 'Origin',
  };
  if (origin && ALLOWED_ORIGINS.length === 0) {
    headers['Access-Control-Allow-Origin'] = origin;
  } else if (origin && ALLOWED_ORIGINS.includes(origin)) {
    headers['Access-Control-Allow-Origin'] = origin;
  }
  return headers;
};

const safeRespond = (statusCode, body, event, extraHeaders = {}) => ({
  statusCode,
  headers: {
    'Content-Type': 'application/json',
    ...corsHeaders(event?.headers?.origin || event?.headers?.Origin),
    ...extraHeaders,
  },
  body: typeof body === 'string' ? body : JSON.stringify(body),
});

export async function handler(event) {
  // CORS preflight
  if (event.httpMethod === 'OPTIONS') {
    return { statusCode: 204, headers: corsHeaders(event.headers.origin || event.headers.Origin), body: '' };
  }

  // SECURITY: The previous implementation accepted an unauthenticated `DELETE`
  // request that wiped the entire `scans` table. That behavior has been
  // removed. If you need a way to clear scans, do it via an authenticated
  // admin script or a dedicated, token-gated function.
  if (event.httpMethod === 'DELETE') {
    return safeRespond(405, { error: 'Method Not Allowed' }, event, { Allow: 'GET, OPTIONS' });
  }

  if (event.httpMethod !== 'GET') {
    return safeRespond(405, { error: 'Method Not Allowed' }, event, { Allow: 'GET, OPTIONS' });
  }

  // --- Optional admin-token gate ---
  if (ADMIN_TOKEN) {
    const provided = event.headers['x-admin-token'] || event.headers['X-Admin-Token'];
    if (provided !== ADMIN_TOKEN) {
      return safeRespond(401, { error: 'Unauthorized' }, event);
    }
  }

  const connectionString = process.env.NETLIFY_DATABASE_URL || process.env.DATABASE_URL;
  if (!connectionString) {
    console.error('Database connection string is missing. Check your environment variables.');
    return safeRespond(500, { error: 'Internal Server Error' }, event);
  }
  const sql = neon(connectionString);

  // Optional `?limit=N` query-string support (capped at MAX_ROWS).
  let limit = MAX_ROWS;
  if (event.queryStringParameters?.limit) {
    const parsed = parseInt(event.queryStringParameters.limit, 10);
    if (Number.isFinite(parsed) && parsed > 0) {
      limit = Math.min(parsed, MAX_ROWS);
    }
  }

  try {
    const rows = await sql`
      SELECT id, filename, image_data, created_at
      FROM scans
      ORDER BY created_at DESC
      LIMIT ${limit}
    `;

    return safeRespond(200, rows, event);
  } catch (err) {
    console.error('Database Error:', err);
    return safeRespond(500, { error: 'Internal Server Error' }, event);
  }
}
