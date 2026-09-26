// /api/notion-export.js
// Serverless function (Vercel) — perantara aman ke Notion API.
// API key Notion disimpan sebagai environment variable di Vercel
// (NOTION_API_KEY), TIDAK PERNAH dikirim ke browser.
//
// PENTING: file ini WAJIB berada persis di folder "api/" di root project
// (yaitu di path "api/notion-export.js") supaya Vercel otomatis
// mengenalinya sebagai endpoint "/api/notion-export". Kalau file ini
// ditaruh di root project (bukan di dalam folder "api/"), endpoint-nya
// tidak akan ada sama sekali dan setiap request dari index.html akan gagal
// (fetch ke "/api/notion-export" akan mendapat 404 dari Vercel).
//
// Terima 1 baris data lewat POST, buat 1 page baru di database Notion
// "Data Scraping" dengan properti yang sudah disesuaikan skemanya.

const NOTION_VERSION = '2025-09-03';

// Data source ID dari database Notion "Data Scraping".
// Bisa dioverride lewat env var NOTION_DATA_SOURCE_ID kalau suatu saat
// datanya dipindah ke database/workspace lain.
const DEFAULT_DATA_SOURCE_ID = '3c3dcd14-e2c8-8040-8462-000bc13161ec';

module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    res.status(204).end();
    return;
  }

  const apiKey = process.env.NOTION_API_KEY;
  const dataSourceId = process.env.NOTION_DATA_SOURCE_ID || DEFAULT_DATA_SOURCE_ID;

  // ---- Diagnostic route: buka URL ini langsung di browser ----
  // https://<domain-vercel-kamu>/api/notion-export
  // Kalau ini menampilkan 404 dari Vercel (bukan JSON di bawah), berarti
  // function belum ter-deploy dengan benar (biasanya karena file tidak ada
  // di folder "api/"). Kalau menampilkan JSON tapi hasApiKey: false, berarti
  // env var belum tersimpan / belum ter-redeploy di Vercel.
  if (req.method === 'GET') {
    res.status(200).json({
      ok: true,
      info: 'Function notion-export aktif dan bisa dijangkau.',
      hasApiKey: Boolean(apiKey),
      dataSourceIdPreview: dataSourceId ? `${dataSourceId.slice(0, 8)}...` : null,
      usingDefaultDataSourceId: !process.env.NOTION_DATA_SOURCE_ID
    });
    return;
  }

  if (req.method !== 'POST') {
    res.status(405).json({ ok: false, error: 'Method not allowed' });
    return;
  }

  if (!apiKey) {
    res.status(500).json({
      ok: false,
      error: 'NOTION_API_KEY belum diset di environment variable Vercel (atau belum redeploy setelah diisi).'
    });
    return;
  }

  let body = req.body;
  if (typeof body === 'string') {
    try { body = JSON.parse(body); } catch (e) { body = {}; }
  }
  const row = (body && body.row) || {};

  const name = (row.name || '').toString().trim();
  const phone = (row.phone || '').toString().trim();
  const alamat = (row.alamat || '').toString().trim();
  const maps = (row.maps || '').toString().trim();
  const coverage = (row.coverage || '').toString().trim(); // 'Cover' | 'Uncover' | ''
  const category = (row.category || '').toString().trim();

  const properties = {
    'Name': { title: [{ text: { content: name.slice(0, 2000) || '(tanpa nama)' } }] },
    'Alamat': { rich_text: alamat ? [{ text: { content: alamat.slice(0, 2000) } }] : [] },
    'Category': { rich_text: category ? [{ text: { content: category.slice(0, 2000) } }] : [] }
  };

  if (phone) properties['Phone'] = { phone_number: phone };
  if (maps) {
    try {
      new URL(maps);
      properties['Maps'] = { url: maps };
    } catch (e) {
      // bukan URL valid, dilewati supaya Notion API tidak menolak seluruh request
    }
  }
  if (coverage === 'Cover' || coverage === 'Uncover') {
    properties['Coverage '] = { select: { name: coverage } };
  }

  try {
    const response = await fetch('https://api.notion.com/v1/pages', {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${apiKey}`,
        'Notion-Version': NOTION_VERSION,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        parent: { type: 'data_source_id', data_source_id: dataSourceId },
        properties
      })
    });

    const data = await response.json();

    if (!response.ok) {
      res.status(response.status).json({ ok: false, error: data.message || 'Notion API error', notion: data });
      return;
    }

    res.status(200).json({ ok: true, id: data.id, url: data.url });
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message || 'Unknown error' });
  }
};
