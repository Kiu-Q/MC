import { neon } from '@netlify/neon';

// --- Configuration ---
// Max body size accepted by this function (5 MB). Netlify Functions already
// cap payloads at ~6 MB, but we enforce a smaller, app-appropriate limit.
const MAX_BODY_BYTES = 5 * 1024 * 1024;
const MAX_FILENAME_LEN = 255;
const MAX_IMAGE_DATA_LEN = 4 * 1024 * 1024; // 4 MB worth of base64 (~3 MB image)

// Allowed origins for CORS. Add your production URL(s) here.
// Falls back to Netlify's automatic same-origin handling when unset.
const ALLOWED_ORIGINS = (process.env.ALLOWED_ORIGINS || '')
  .split(',')
  .map((s) => s.trim())
  .filter(Boolean);

// Optional shared secret for write operations. If set, clients must send
// `x-admin-token: <value>` to POST. Skip check if not configured (so local
// dev still works), but RECOMMEND setting it in production.
const ADMIN_TOKEN = process.env.ADMIN_TOKEN || '';

const corsHeaders = (origin) => {
  const headers = {
    'Access-Control-Allow-Methods': 'POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, x-admin-token',
    'Access-Control-Max-Age': '86400',
    'Vary': 'Origin',
  };
  if (origin && ALLOWED_ORIGINS.length === 0) {
    // No allow-list configured: allow same-origin by default (Netlify behavior).
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
  // Handle CORS preflight
  if (event.httpMethod === 'OPTIONS') {
    return { statusCode: 204, headers: corsHeaders(event.headers.origin || event.headers.Origin), body: '' };
  }

  // Only allow POST
  if (event.httpMethod !== 'POST') {
    return safeRespond(405, { error: 'Method Not Allowed' }, event, { Allow: 'POST, OPTIONS' });
  }

  // --- Optional admin-token gate ---
  if (ADMIN_TOKEN) {
    const provided = event.headers['x-admin-token'] || event.headers['X-Admin-Token'];
    if (provided !== ADMIN_TOKEN) {
      return safeRespond(401, { error: 'Unauthorized' }, event);
    }
  }

  // --- Body size check BEFORE parsing ---
  const bodyText = event.body || '';
  if (typeof event.isBase64Encoded === 'boolean' && event.isBase64Encoded) {
    // base64-encoded body: size in chars is a reasonable proxy
    if (bodyText.length > MAX_BODY_BYTES * 1.4) {
      return safeRespond(413, { error: 'Payload Too Large' }, event);
    }
  } else if (bodyText.length > MAX_BODY_BYTES) {
    return safeRespond(413, { error: 'Payload Too Large' }, event);
  }

  let body;
  try {
    body = JSON.parse(bodyText || '{}');
  } catch {
    return safeRespond(400, { error: 'Invalid JSON' }, event);
  }

  const { filename, image_data, created_at } = body;

  // --- Field validation ---
  if (typeof filename !== 'string' || filename.length === 0) {
    return safeRespond(400, { error: 'Missing or invalid filename' }, event);
  }
  if (filename.length > MAX_FILENAME_LEN) {
    return safeRespond(400, { error: 'Filename too long' }, event);
  }
  // Allow letters, numbers, spaces, dash, underscore, dot, parens. Strip control chars.
  if (/[\x00-\x1f\x7f]/.test(filename)) {
    return safeRespond(400, { error: 'Invalid characters in filename' }, event);
  }

  if (typeof image_data !== 'string' || image_data.length === 0) {
    return safeRespond(400, { error: 'Missing or invalid image_data' }, event);
  }
  if (image_data.length > MAX_IMAGE_DATA_LEN) {
    return safeRespond(413, { error: 'image_data too large' }, event);
  }
  // Must look like a data URL for an image
  if (!/^data:image\/(png|jpe?g|webp|gif);base64,[A-Za-z0-9+/=]+$/.test(image_data)) {
    return safeRespond(400, { error: 'image_data must be a base64 data URL for an image' }, event);
  }

  // --- DB connection ---
  const connectionString = process.env.NETLIFY_DATABASE_URL || process.env.DATABASE_URL;
  if (!connectionString) {
    console.error('Database connection string is missing. Check your environment variables.');
    return safeRespond(500, { error: 'Internal Server Error' }, event);
  }
  const sql = neon(connectionString);

  try {
    const [row] = await sql`
      INSERT INTO scans (filename, image_data, created_at)
      VALUES (${filename}, ${image_data}, ${created_at ?? new Date().toISOString()})
      RETURNING id, filename, created_at
    `;

    return safeRespond(200, row, event);
  } catch (err) {
    // Log full error server-side; return generic message to client.
    console.error('Database Insert Error:', err);
    return safeRespond(500, { error: 'Internal Server Error' }, event);
  }
}
