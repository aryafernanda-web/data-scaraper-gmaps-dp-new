// /api/import-excel.js
// Vercel Serverless Function — POST /api/import-excel
//
// Input : multipart/form-data
//   - file   : file Excel (.xlsx / .xls) atau CSV
//   - radius : radius dalam meter (default 350)
//
// Output: JSON
//   {
//     ok: true,
//     summary: { total, covered, notCovered, noCoord },
//     rows: [
//       { nama, nomor, address, status, distance }
//       ...
//     ]
//   }
//
// Logic yang dijalankan sama persis dengan client-side di index.html:
//   1. Parse file dengan XLSX (kolom yang sama: nama/nomor/address/link/lat/lng)
//   2. Ekstrak koordinat dari kolom lat/lng atau dari URL Google Maps (parseLatLngFromUrl)
//   3. Geocoding via Nominatim untuk baris yang belum punya koordinat tapi ada address
//   4. Hitung jarak Haversine ke setiap titik DP, ambil yang terdekat
//   5. Tentukan status: Tercover / Tidak Tercover / Alamat Tidak Ditemukan / Tidak Ada Koordinat

'use strict';

const formidable = require('formidable');
const XLSX = require('xlsx');

// ---------- Titik-titik DP (lat, lng) — sama persis dengan index.html ----------
const DP_POINTS = [[-8.078741,111.900203],[-8.080025,111.898795],[-8.080503,111.896551],[-8.080434,111.89377],[-8.085942,111.901543],[-8.086814,111.905963],[-8.088842,111.904101],[-8.088909,111.904915],[-8.081272,111.902941],[-8.083022,111.902424],[-8.087978,111.900949],[-8.089542,111.900517],[-8.089124,111.900489],[-8.087889,111.899258],[-8.087601,111.900055],[-8.078665,111.898331],[-8.076459,111.896314],[-8.075778,111.894027],[-8.081471,111.902768],[-8.087871,111.900863],[-8.076404,111.899862],[-8.075839,111.898666],[-8.075749,111.897242],[-8.074424,111.897432],[-8.075092,111.90227],[-8.074128,111.90215],[-8.074658,111.9046],[-8.076452,111.902355],[-8.078814,111.900411],[-8.07887,111.902698],[-8.078646,111.904803],[-8.078621,111.905941],[-8.081326,111.900265],[-8.07817,111.887151],[-8.076667,111.885417],[-8.076162,111.883495],[-8.075777,111.881874],[-8.075574,111.887951],[-8.076189,111.889763],[-8.074281,111.888395],[-8.074669,111.889643],[-8.080211,111.886526],[-8.081633,111.88612],[-8.081454,111.884428],[-8.081161,111.88233],[-8.081685,111.892481],[-8.08161,111.894148],[-8.081528,111.896325],[-8.081367,111.898512],[-8.080046,111.890208],[-8.081909,111.889408],[-8.083324,111.889052],[-8.083274,111.890965],[-8.076537,111.891975],[-8.075893,111.892285],[-8.074578,111.892948],[-8.077204,111.887469],[-8.078533,111.89688],[-8.07846,111.895754],[-8.078219,111.892018],[-8.077807,111.889649],[-8.058631,111.916766],[-8.049531,111.906436],[-8.051053,111.906031],[-8.052609,111.90555],[-8.054225,111.90478],[-8.0471425,111.9094017],[-8.045442,111.909844],[-8.042866,111.91112],[-8.041389,111.912003],[-8.054281,111.907865],[-8.053285,111.908085],[-8.05048,111.908758],[-8.048278,111.909291],[-8.054358,111.907687],[-8.052594,111.908103],[-8.049727,111.908797],[-8.048678,111.909041],[-8.053945,111.913457],[-8.053749,111.912148],[-8.053464,111.909468],[-8.052097,111.913455],[-8.051696,111.91158],[-8.051313,111.909218],[-8.050052,111.913419],[-8.048385,111.913257],[-8.051665,111.922584],[-8.050997,111.921893],[-8.051004,111.921213],[-8.051467,111.923613],[-8.050081,111.924131],[-8.051591,111.924718],[-8.050896,111.924737],[-8.050127,111.925378],[-8.054886,111.91736],[-8.053885,111.917378],[-8.052186,111.917415],[-8.051779,111.916128],[-8.05164,111.918389],[-8.051649,111.920427],[-8.050526,111.920406],[-8.050501,111.922354],[-8.058613,111.928673],[-8.059295,111.929104],[-8.058042,111.929055],[-8.059518,111.928267],[-8.05403,111.916321],[-8.053117,111.916639],[-8.054621,111.916716],[-8.05396,111.917086],[-8.05673,111.919992],[-8.05837,111.92081],[-8.058652,111.918298],[-8.058822,111.916899],[-8.058198,111.922688],[-8.057825,111.925341],[-8.056149,111.922868],[-8.0552,111.922149],[-8.057452,111.928046],[-8.05585,111.905881],[-8.055909,111.904476],[-8.058264,111.904547],[-8.059427,111.904612],[-8.055342,111.91825],[-8.055167,111.9210775],[-8.05512,111.923527],[-8.057357,111.91659],[-8.055432,111.916456],[-8.055594,111.915062],[-8.055678,111.913421],[-8.055795,111.910625],[-8.057478,111.911129],[-8.055823,111.909529],[-8.055772,111.908491],[-8.059932,111.938191],[-8.059762,111.939148],[-8.060638,111.934122],[-8.060227,111.936402],[-8.060401,111.920725],[-8.060457,111.922216],[-8.059881,111.92155],[-8.062494,111.922943],[-8.070079,111.915738],[-8.070314,111.916712],[-8.069967,111.917319],[-8.069859,111.918819],[-8.072272,111.918993],[-8.072705,111.918338],[-8.071861,111.921194],[-8.072036,111.92254],[-8.0676077,111.9137618],[-8.067325,111.911707],[-8.068256,111.917602],[-8.068534,111.919231],[-8.062272,111.924254],[-8.061888,111.926587],[-8.061619,111.928262],[-8.061084,111.931433],[-8.066296,111.920775],[-8.064549,111.921058],[-8.062869,111.9208],[-8.063293,111.918088],[-8.071613,111.919361],[-8.070876,111.920052],[-8.068561,111.920444],[-8.067484,111.920577],[-8.067246,111.915266],[-8.068626,111.914957],[-8.070994,111.914452],[-8.07142,111.916614],[-8.060751,111.916334],[-8.062613,111.916038],[-8.064479,111.915735],[-8.065745,111.915527],[-8.064905,111.908347],[-8.064223,111.912465],[-8.063897,111.914481],[-8.063014,111.908862],[-8.062546,111.912205],[-8.062226,111.914197],[-8.061967,111.915575],[-8.061768,111.90865],[-8.061243,111.912677],[-8.062226,111.914197],[-8.060806,111.915768],[-8.06719,111.910363],[-8.069766,111.909875],[-8.070101,111.910432],[-8.070517,111.912206],[-8.06148,111.910981],[-8.062757,111.9109],[-8.06448,111.910767],[-8.066306,111.910529],[-8.063425,111.905687],[-8.063309,111.9065],[-8.06538,111.905449],[-8.065225,111.906315],[-8.0599818,111.9054022],[-8.060161,111.904752],[-8.061092,111.904838],[-8.062018,111.906436],[-8.06575,111.907438],[-8.06716,111.907422],[-8.068545,111.907429],[-8.067296,111.908809],[-8.057234,111.908542],[-8.057586,111.907394],[-8.061568,111.907288],[-8.063268,111.907413],[-8.065554,111.907275],[-8.067013,111.907273],[-8.068222,111.907246],[-8.06971,111.908436],[-8.059549,111.907392],[-8.057882,111.905564],[-8.06161,111.907413],[-8.063133,111.907272],[-8.0588506,111.9161655],[-8.059122,111.914399],[-8.059369,111.912031],[-8.059479,111.910975],[-8.055028,111.896391],[-8.054809,111.895034],[-8.0544773,111.8936177],[-8.054215,111.892187],[-8.053928,111.890731],[-8.053555,111.88891],[-8.053346,111.887371],[-8.053084,111.886039],[-8.05285,111.897182],[-8.050823,111.897979],[-8.04997,111.898305],[-8.048985,111.898754],[-8.048301,111.899031],[-8.047712,111.899265],[-8.046335,111.899823],[-8.044959,111.90041],[-8.053352,111.890951],[-8.053538,111.891934],[-8.05366,111.892567],[-8.053213,111.892009],[-8.052478,111.891128],[-8.052668,111.892098],[-8.052962,111.892003],[-8.053084,111.892509],[-8.051795,111.891266],[-8.052025,111.892177],[-8.052342,111.891476],[-8.052243,111.892139],[-8.052359,111.893247],[-8.051641,111.892231],[-8.051107,111.890063],[-8.050782,111.889465],[-8.053067,111.889617],[-8.052845,111.888551],[-8.052249,111.887061],[-8.052392,111.88811],[-8.052645,111.889329],[-8.052241,111.889889],[-8.051921,111.888201],[-8.051724,111.887202],[-8.052006,111.887144],[-8.052152,111.888161],[-8.052387,111.889388],[-8.051656,111.890557],[-8.051722,111.88936],[-8.051389,111.887989],[-8.051505,111.887297],[-8.051689,111.888243],[-8.051972,111.889414],[-8.055175,111.897495],[-8.056607,111.896751],[-8.057499,111.896993],[-8.058631,111.89733],[-8.059315,111.896219],[-8.060565,111.897906],[-8.061039,111.895827],[-8.061355,111.893832],[-8.061345,111.893088],[-8.06193,111.89131],[-8.05868,111.89365],[-8.057819,111.893975],[-8.059949,111.893259],[-8.059631,111.891438],[-8.058838,111.889853],[-8.057822,111.887922],[-8.060929,111.888888],[-8.060514,111.887499],[-8.059585,111.884987],[-8.058806,111.883479],[-8.056856,111.890724],[-8.057055,111.891596],[-8.057287,111.892709],[-8.0579,111.892696],[-8.057474,111.890617],[-8.058136,111.890481],[-8.056128,111.890886],[-8.055697,111.890606],[-8.05832,111.891345],[-8.057675,111.891475],[-8.05627,111.891751],[-8.056342,111.892909],[-8.080461,111.914091],[-8.079876,111.911988],[-8.079386,111.910251],[-8.079004,111.908907],[-8.078608,111.907285],[-8.079718,111.907308],[-8.081483,111.907311],[-8.083446,111.907311],[-8.083547,111.90916],[-8.083613,111.910373],[-8.083901,111.91256],[-8.085949,111.907329],[-8.086193,111.909018],[-8.08636,111.910617],[-8.089194,111.90758],[-8.091253,111.907778],[-8.090903,111.910299],[-8.093075,111.908081],[-8.092937,111.91042],[-8.095646,111.908809],[-8.075924,111.909174],[-8.076377,111.911138],[-8.0769232,111.9135064],[-8.078713,111.913959],[-8.077133,111.914099],[-8.077493,111.915674],[-8.078229,111.91557],[-8.079219,111.915528],[-8.078162,111.915339],[-8.07785,111.916979],[-8.078558,111.916837],[-8.079326,111.916688],[-8.079875,111.916768],[-8.080179,111.917755],[-8.08032,111.917243],[-8.080464,111.916794],[-8.078338,111.918641],[-8.0822,111.913947],[-8.084095,111.913869],[-8.086714,111.913832],[-8.086471,111.912016],[-8.079594,111.919173],[-8.07983,111.921835],[-8.080857,111.915679],[-8.081522,111.918053],[-8.082801,111.917794],[-8.084697,111.91748],[-8.085008,111.920229],[-8.08296,111.919869],[-8.083125,111.922564],[-8.085226,111.922236],[-8.083328,111.925996],[-8.081503,111.926074],[-8.078878,111.927262],[-8.079673,111.931771],[-8.085521,111.925996],[-8.085105,111.928449]];

// ---------- Bounding box DP (untuk bias viewbox geocoding) ----------
function dpBoundingBox() {
  let minLat = Infinity, maxLat = -Infinity, minLng = Infinity, maxLng = -Infinity;
  for (const [lat, lng] of DP_POINTS) {
    if (lat < minLat) minLat = lat;
    if (lat > maxLat) maxLat = lat;
    if (lng < minLng) minLng = lng;
    if (lng > maxLng) maxLng = lng;
  }
  const pad = 0.15;
  return { minLat: minLat - pad, maxLat: maxLat + pad, minLng: minLng - pad, maxLng: maxLng + pad };
}
const DP_BBOX = dpBoundingBox();

// ---------- Utilitas — sama dengan index.html ----------
function findColumn(headers, candidates) {
  const norm = h => String(h).toLowerCase().replace(/[^a-z]/g, '');
  const normHeaders = headers.map(norm);
  for (const cand of candidates) {
    const idx = normHeaders.indexOf(norm(cand));
    if (idx !== -1) return headers[idx];
  }
  for (const cand of candidates) {
    const c = norm(cand);
    const idx = normHeaders.findIndex(h => h.includes(c));
    if (idx !== -1) return headers[idx];
  }
  return null;
}

function parseLatLngFromUrl(url) {
  if (!url) return null;
  url = String(url);
  let m = url.match(/!3d(-?\d+\.\d+)!4d(-?\d+\.\d+)/);
  if (m) return { lat: parseFloat(m[1]), lng: parseFloat(m[2]) };
  m = url.match(/@(-?\d+\.\d+),(-?\d+\.\d+)/);
  if (m) return { lat: parseFloat(m[1]), lng: parseFloat(m[2]) };
  m = url.match(/[?&](?:q|ll)=(-?\d+\.\d+),(-?\d+\.\d+)/);
  if (m) return { lat: parseFloat(m[1]), lng: parseFloat(m[2]) };
  return null;
}

function haversine(lat1, lon1, lat2, lon2) {
  const R = 6371000;
  const toRad = d => d * Math.PI / 180;
  const dLat = toRad(lat2 - lat1);
  const dLon = toRad(lon2 - lon1);
  const a = Math.sin(dLat / 2) ** 2 + Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) * Math.sin(dLon / 2) ** 2;
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}

function nearestDpDistance(lat, lng) {
  let best = Infinity;
  for (const [dpLat, dpLng] of DP_POINTS) {
    const d = haversine(lat, lng, dpLat, dpLng);
    if (d < best) best = d;
  }
  return best;
}

function sleep(ms) {
  return new Promise(r => setTimeout(r, ms));
}

async function geocodeAddress(address) {
  const viewbox = `${DP_BBOX.minLng},${DP_BBOX.maxLat},${DP_BBOX.maxLng},${DP_BBOX.minLat}`;
  const url = 'https://nominatim.openstreetmap.org/search?format=json&limit=1&countrycodes=id'
    + `&viewbox=${viewbox}&bounded=0&q=${encodeURIComponent(address)}`;
  const res = await fetch(url, {
    headers: {
      'Accept': 'application/json',
      // Nominatim mewajibkan User-Agent untuk identifikasi
      'User-Agent': 'cek-coverage-dp-fixed/1.0 (serverless; contact=admin)'
    }
  });
  if (!res.ok) throw new Error('Nominatim HTTP ' + res.status);
  const data = await res.json();
  if (!data.length) return null;
  return { lat: parseFloat(data[0].lat), lng: parseFloat(data[0].lon) };
}

// ---------- Parse multipart/form-data ----------
function parseForm(req) {
  return new Promise((resolve, reject) => {
    const form = formidable({ maxFileSize: 20 * 1024 * 1024 }); // 20 MB
    form.parse(req, (err, fields, files) => {
      if (err) return reject(err);
      resolve({ fields, files });
    });
  });
}

// ---------- Handler utama ----------
module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    res.status(204).end();
    return;
  }

  if (req.method !== 'POST') {
    res.status(405).json({ ok: false, error: 'Method not allowed. Use POST multipart/form-data.' });
    return;
  }

  // ---- 1. Parse form ----
  let fields, files;
  try {
    ({ fields, files } = await parseForm(req));
  } catch (err) {
    res.status(400).json({ ok: false, error: 'Gagal parse form: ' + err.message });
    return;
  }

  // formidable v3 bisa wrap nilai dalam array
  const fileObj = Array.isArray(files.file) ? files.file[0] : files.file;
  if (!fileObj) {
    res.status(400).json({ ok: false, error: 'Field "file" tidak ditemukan di request.' });
    return;
  }

  const rawRadius = Array.isArray(fields.radius) ? fields.radius[0] : fields.radius;
  const radius = parseFloat(rawRadius) || 350;

  // ---- 2. Baca file Excel/CSV ----
  let jsonRows;
  try {
    const wb = XLSX.readFile(fileObj.filepath || fileObj.path);
    const sheet = wb.Sheets[wb.SheetNames[0]];
    jsonRows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
  } catch (err) {
    res.status(400).json({ ok: false, error: 'Gagal membaca file Excel/CSV: ' + err.message });
    return;
  }

  if (!jsonRows || !jsonRows.length) {
    res.status(400).json({ ok: false, error: 'File kosong atau tidak ada baris data.' });
    return;
  }

  // ---- 3. Deteksi kolom (sama dengan client-side) ----
  const headers = Object.keys(jsonRows[0]);
  const colNama    = findColumn(headers, ['nama', 'name', 'title']);
  const colNomor   = findColumn(headers, ['nomor', 'no', 'phone', 'telepon', 'notelp']);
  const colAddress = findColumn(headers, ['address', 'alamat']);
  const colLink    = findColumn(headers, ['linkmaps', 'link', 'maps', 'googlemaps', 'urlmaps', 'url']);
  const colLat     = findColumn(headers, ['latitude', 'lat', 'locationlat']);
  const colLng     = findColumn(headers, ['longitude', 'lng', 'lon', 'long', 'locationlng', 'locationlon']);

  if (!colLink && !(colLat && colLng)) {
    res.status(400).json({
      ok: false,
      error: 'Tidak ditemukan kolom link maps ataupun kolom latitude/longitude di file.'
    });
    return;
  }

  const parsedRows = jsonRows.map(row => ({
    nama:    colNama    ? row[colNama]              : '',
    nomor:   colNomor   ? row[colNomor]             : '',
    address: colAddress ? row[colAddress]           : '',
    link:    colLink    ? row[colLink]              : '',
    lat:     colLat     ? parseFloat(row[colLat])   : null,
    lng:     colLng     ? parseFloat(row[colLng])   : null
  }));

  // ---- 4. Tahap 1: tentukan koordinat dari kolom lat/lng atau dari URL ----
  const working = parsedRows.map(row => {
    let coord = null;
    if (row.lat != null && row.lng != null && !isNaN(row.lat) && !isNaN(row.lng)) {
      coord = { lat: row.lat, lng: row.lng };
    } else {
      coord = parseLatLngFromUrl(row.link);
    }
    return { ...row, coord, geocoded: false };
  });

  // ---- 5. Tahap 2: geocoding Nominatim untuk baris tanpa koordinat tapi ada address ----
  const needGeocode = working.filter(r => !r.coord && r.address && String(r.address).trim());
  for (let i = 0; i < needGeocode.length; i++) {
    const row = needGeocode[i];
    try {
      const coord = await geocodeAddress(String(row.address));
      if (coord) { row.coord = coord; row.geocoded = true; }
    } catch (_) {
      // coord tetap null, akan ditandai gagal di bawah
    }
    // patuhi rate-limit Nominatim: 1 request/detik
    if (i < needGeocode.length - 1) await sleep(1100);
  }

  // ---- 6. Tahap 3: hitung jarak Haversine & tentukan status ----
  let covered = 0, notCovered = 0, invalid = 0;

  const resultRows = working.map(row => {
    if (!row.coord) {
      invalid++;
      const status = (row.address && String(row.address).trim())
        ? 'Alamat Tidak Ditemukan'
        : 'Tidak Ada Koordinat';
      return { nama: row.nama, nomor: row.nomor, address: row.address, status, distance: null };
    }
    const dist = nearestDpDistance(row.coord.lat, row.coord.lng);
    const isCovered = dist <= radius;
    if (isCovered) covered++; else notCovered++;
    const status = (isCovered ? 'Tercover' : 'Tidak Tercover') + (row.geocoded ? ' (est. alamat)' : '');
    return { nama: row.nama, nomor: row.nomor, address: row.address, status, distance: Math.round(dist) };
  });

  // ---- 7. Kirim respons ----
  res.status(200).json({
    ok: true,
    summary: {
      total: resultRows.length,
      covered,
      notCovered,
      noCoord: invalid
    },
    rows: resultRows
  });
};
