import { neon } from '@netlify/neon';

export async function handler(event) {
  const connectionString = process.env.NETLIFY_DATABASE_URL || process.env.DATABASE_URL;
  
  if (!connectionString) {
    return { 
      statusCode: 500, 
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ error: "Database connection string is missing." }) 
    };
  }

  const sql = neon(connectionString);

  try {
    // GET: Retrieve all scans for download
    if (event.httpMethod === 'GET') {
      const rows = await sql`
        SELECT filename, image_data 
        FROM scans 
        ORDER BY created_at DESC 
        LIMIT 500
      `;

      return {
        statusCode: 200,
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(rows),
      };
    }

    // DELETE: Clear the scans table
    if (event.httpMethod === 'DELETE') {
      await sql`DELETE FROM scans`;

      return {
        statusCode: 200,
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ message: "Database cleared successfully" }),
      };
    }

    return { statusCode: 405, body: 'Method Not Allowed' };

  } catch (err) {
    console.error("Database Error:", err);
    return { 
      statusCode: 500, 
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ error: err.message || 'Internal Server Error' }) 
    };
  }
}