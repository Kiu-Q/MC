import { neon } from '@netlify/neon';

// Allowed origins for CORS. Add your production URL(s) here as a comma-separated list.
const ALLOWED_ORIGINS = (process.env.ALLOWED_ORIGINS || '')
  .split(',')
  .map((s) => s.trim())
  .filter(Boolean);

const corsHeaders = (origin) => {
  const headers = {
    'Access-Control-Allow-Methods': 'GET, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type',
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

  if (event.httpMethod !== 'GET') {
    return safeRespond(405, { error: 'Method Not Allowed' }, event, { Allow: 'GET, OPTIONS' });
  }

  const connectionString = process.env.NETLIFY_DATABASE_URL || process.env.DATABASE_URL;
  if (!connectionString) {
    console.error('Database connection string is missing. Check your environment variables.');
    return safeRespond(500, { error: 'Internal Server Error' }, event);
  }
  const sql = neon(connectionString);

  try {
    // Return at most one link row (table is expected to hold a single config row).
    const rows = await sql`
      SELECT link
      FROM link
      LIMIT 1
    `;

    return safeRespond(200, rows, event);
  } catch (err) {
    console.error('Database Error:', err);
    return safeRespond(500, { error: 'Internal Server Error' }, event);
  }
}
