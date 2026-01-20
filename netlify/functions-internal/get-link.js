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
        SELECT link 
        FROM link
      `;

      return {
        statusCode: 200,
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(rows),
      };
    }

  } catch (err) {
    console.error("Database Error:", err);
    return { 
      statusCode: 500, 
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ error: err.message || 'Internal Server Error' }) 
    };
  }
}