import { neon } from '@netlify/neon';

export async function handler(event) {
  // Only allow POST requests
  if (event.httpMethod !== 'POST') {
    return { statusCode: 405, body: 'Method Not Allowed' };
  }

  try {
    // 1. Parse Data
    const body = JSON.parse(event.body || '{}');
    const { filename, image_data, created_at } = body;

    // Validation
    if (!filename || !image_data) {
      console.error("Missing fields:", body);
      return { statusCode: 400, body: JSON.stringify({ error: "Missing filename or image_data" }) };
    }

    // 2. Initialize DB Connection
    // explicitly trying process.env.DATABASE_URL as a fallback if NETLIFY_DATABASE_URL is missing in local dev
    const connectionString = process.env.NETLIFY_DATABASE_URL || process.env.DATABASE_URL;
    
    if (!connectionString) {
      throw new Error("Database connection string is missing. Check your environment variables.");
    }

    const sql = neon(connectionString);

    // 3. Execute Insert
    // Note: This assumes a table named 'scans' exists with columns: filename (TEXT), image_data (TEXT), created_at (TIMESTAMP)
    const [row] = await sql`
      INSERT INTO scans (filename, image_data, created_at)
      VALUES (${filename}, ${image_data}, ${created_at})
      RETURNING *
    `;

    // 4. Success Response
    return {
      statusCode: 200,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(row),
    };

  } catch (err) {
    // Log actual error to Netlify function logs
    console.error("Database Insert Error:", err);

    return { 
      statusCode: 500, 
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ 
        error: err.message || 'Internal Server Error',
        details: err.toString() 
      }) 
    };
  }
}