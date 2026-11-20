const http = require('node:http');
const fs = require('node:fs');
const path = require('node:path');
const zlib = require('node:zlib');
const { DatabaseSync } = require('node:sqlite');

const PORT = process.env.PORT || 3001;
const dbPath = path.join(__dirname, 'stok.db');
const db = new DatabaseSync(dbPath);

// Skema simpel: kategori, barang, transaksi
db.exec(`
  CREATE TABLE IF NOT EXISTS kategori (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nama TEXT NOT NULL UNIQUE,
    urutan INTEGER DEFAULT 0
  );
  CREATE TABLE IF NOT EXISTS barang (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nama TEXT NOT NULL,
    subkategori_id INTEGER,
    varian TEXT,
    ukuran TEXT,
    aktif INTEGER DEFAULT 1,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP
  );
  CREATE TABLE IF NOT EXISTS transaksi (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    barang_id INTEGER NOT NULL,
    tanggal TEXT NOT NULL,
    tipe TEXT NOT NULL,
    qty REAL DEFAULT 1,
    pic TEXT DEFAULT '',
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP
  );
  CREATE INDEX IF NOT EXISTS idx_transaksi_barang_tgl ON transaksi(barang_id, tanggal);
  CREATE INDEX IF NOT EXISTS idx_transaksi_tgl ON transaksi(tanggal);
`);

// 6 Kategori tepat sesuai sheet Excel asli
const defaultCategories = ['Kardus', 'Nampan', 'Standing', 'Kantor', 'Pita', 'Box'];
const existCat = db.prepare('SELECT COUNT(*) as cnt FROM kategori').get();
if (!existCat || existCat.cnt === 0) {
  const insertCat = db.prepare('INSERT INTO kategori (nama, urutan) VALUES (?, ?)');
  defaultCategories.forEach((c, i) => insertCat.run(c, i + 1));
}

function sendJson(res, statusCode, data) {
  const payload = JSON.stringify(data);
  res.writeHead(statusCode, {
    'Content-Type': 'application/json; charset=utf-8',
    'Content-Length': Buffer.byteLength(payload)
  });
  res.end(payload);
}

function sendHtml(res, filePath) {
  if (!fs.existsSync(filePath)) {
    filePath = path.join(__dirname, 'public', 'index.html');
  }
  const content = fs.readFileSync(filePath, 'utf-8');
  res.writeHead(200, {
    'Content-Type': 'text/html; charset=utf-8',
    'Content-Length': Buffer.byteLength(content)
  });
  res.end(content);
}

function parseBody(req) {
  return new Promise((resolve, reject) => {
    let body = '';
    req.on('data', chunk => { body += chunk; });
    req.on('end', () => {
      try {
        resolve(body ? JSON.parse(body) : {});
      } catch (err) {
        reject(new Error('Format JSON tidak valid'));
      }
    });
    req.on('error', reject);
  });
}

function parseQty(val) {
  if (typeof val === 'number') return val;
  if (!val) return 0;
  const num = parseFloat(String(val).replace(',', '.').replace(/[^\d.-]/g, ''));
  return isNaN(num) ? 0 : num;
}

// OpenXML Builder (6 Sheet Sesuai Excel Asli)
const crcTable = new Uint32Array(256);
for (let i = 0; i < 256; i++) {
  let c = i;
  for (let k = 0; k < 8; k++) c = (c & 1) ? (0xedb88320 ^ (c >>> 1)) : (c >>> 1);
  crcTable[i] = c >>> 0;
}
function crc32(buf) {
  let crc = 0 ^ (-1);
  for (let i = 0; i < buf.length; i++) crc = (crc >>> 8) ^ crcTable[(crc ^ buf[i]) & 0xff];
  return (crc ^ (-1)) >>> 0;
}

function createZip(files) {
  const localHeaders = [];
  const centralHeaders = [];
  let offset = 0;

  for (const file of files) {
    const data = Buffer.isBuffer(file.data) ? file.data : Buffer.from(file.data, 'utf-8');
    const compressed = zlib.deflateRawSync(data);
    const useCompressed = compressed.length < data.length;
    const body = useCompressed ? compressed : data;
    const method = useCompressed ? 8 : 0;
    const crc = crc32(data);
    const nameBuf = Buffer.from(file.name, 'utf-8');

    const lh = Buffer.alloc(30 + nameBuf.length);
    lh.writeUInt32LE(0x04034b50, 0);
    lh.writeUInt16LE(20, 4);
    lh.writeUInt16LE(0, 6);
    lh.writeUInt16LE(method, 8);
    lh.writeUInt16LE(0, 10);
    lh.writeUInt16LE(0, 12);
    lh.writeUInt32LE(crc, 14);
    lh.writeUInt32LE(body.length, 18);
    lh.writeUInt32LE(data.length, 22);
    lh.writeUInt16LE(nameBuf.length, 26);
    lh.writeUInt16LE(0, 28);
    nameBuf.copy(lh, 30);
    localHeaders.push(lh, body);

    const ch = Buffer.alloc(46 + nameBuf.length);
    ch.writeUInt32LE(0x02014b50, 0);
    ch.writeUInt16LE(20, 4);
    ch.writeUInt16LE(20, 6);
    ch.writeUInt16LE(0, 8);
    ch.writeUInt16LE(method, 10);
    ch.writeUInt16LE(0, 12);
    ch.writeUInt16LE(0, 14);
    ch.writeUInt32LE(crc, 16);
    ch.writeUInt32LE(body.length, 20);
    ch.writeUInt32LE(data.length, 24);
    ch.writeUInt16LE(nameBuf.length, 28);
    ch.writeUInt16LE(0, 30);
    ch.writeUInt16LE(0, 32);
    ch.writeUInt16LE(0, 34);
    ch.writeUInt16LE(0, 36);
    ch.writeUInt32LE(0, 38);
    ch.writeUInt32LE(offset, 42);
    nameBuf.copy(ch, 46);
    centralHeaders.push(ch);

    offset += lh.length + body.length;
  }

  const centralOffset = offset;
  let centralSize = 0;
  for (const ch of centralHeaders) centralSize += ch.length;

  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0);
  eocd.writeUInt16LE(0, 4);
  eocd.writeUInt16LE(0, 6);
  eocd.writeUInt16LE(files.length, 8);
  eocd.writeUInt16LE(files.length, 10);
  eocd.writeUInt32LE(centralSize, 12);
  eocd.writeUInt32LE(centralOffset, 16);
  eocd.writeUInt16LE(0, 20);

  return Buffer.concat([...localHeaders, ...centralHeaders, eocd]);
}

function escapeXml(str) {
  if (str === null || str === undefined) return '';
  return String(str)
    .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function buildSheetXml(rows) {
  const lastRow = Math.max(rows.length + 1, 1);
  let xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="A1:F${lastRow}"/>
  <sheetViews>
    <sheetView workbookViewId="0">
      <pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>
      <selection pane="bottomLeft" activeCell="A2" sqref="A2"/>
    </sheetView>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="20"/>
  <cols>
    <col min="1" max="1" width="7" customWidth="1"/>
    <col min="2" max="2" width="14" customWidth="1"/>
    <col min="3" max="3" width="12" customWidth="1"/>
    <col min="4" max="4" width="38" customWidth="1"/>
    <col min="5" max="5" width="12" customWidth="1"/>
    <col min="6" max="6" width="22" customWidth="1"/>
  </cols>
  <sheetData>
    <row r="1" ht="26" customHeight="1">
      <c r="A1" t="inlineStr" s="1"><is><t>No</t></is></c>
      <c r="B1" t="inlineStr" s="1"><is><t>Tanggal</t></is></c>
      <c r="C1" t="inlineStr" s="1"><is><t>Tipe</t></is></c>
      <c r="D1" t="inlineStr" s="1"><is><t>Nama Barang</t></is></c>
      <c r="E1" t="inlineStr" s="1"><is><t>Qty</t></is></c>
      <c r="F1" t="inlineStr" s="1"><is><t>PIC</t></is></c>
    </row>`;

  let rNum = 2;
  for (const r of rows) {
    const isAlt = (rNum % 2 === 1);
    const sCenter = isAlt ? 9 : 2;
    const sLeft = isAlt ? 10 : 3;
    const sQty = isAlt ? 11 : 4;
    const sPic = isAlt ? 12 : 5;

    let sTipe = sCenter;
    const tUpper = (r.tipe || '').toUpperCase();
    if (tUpper === 'MASUK') sTipe = 6;
    else if (tUpper === 'KELUAR') sTipe = 7;
    else if (tUpper === 'AWAL') sTipe = 8;

    xml += `
    <row r="${rNum}" ht="20" customHeight="1">
      <c r="A${rNum}" s="${sCenter}"><v>${rNum - 1}</v></c>
      <c r="B${rNum}" t="inlineStr" s="${sCenter}"><is><t>${escapeXml(r.tanggal || '-')}</t></is></c>
      <c r="C${rNum}" t="inlineStr" s="${sTipe}"><is><t>${escapeXml(r.tipe || '-')}</t></is></c>
      <c r="D${rNum}" t="inlineStr" s="${sLeft}"><is><t>${escapeXml(r.nama || '-')}</t></is></c>
      <c r="E${rNum}" s="${sQty}"><v>${r.qty || 0}</v></c>
      <c r="F${rNum}" t="inlineStr" s="${sPic}"><is><t>${escapeXml(r.pic || '-')}</t></is></c>
    </row>`;
    rNum++;
  }

  xml += `
  </sheetData>`;
  if (rows.length > 0) {
    xml += `
  <autoFilter ref="A1:F${lastRow}"/>`;
  }
  xml += `
</worksheet>`;
  return xml;
}

function generateMultiSheetXlsx(categories, rowsByCategory) {
  const sheetList = categories.map(c => ({ name: c, rows: rowsByCategory[c] || [] }));

  let contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>`;

  sheetList.forEach((s, idx) => {
    contentTypesXml += `\n  <Override PartName="/xl/worksheets/sheet${idx + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`;
  });
  contentTypesXml += `\n</Types>`;

  const relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;

  let wbRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`;

  sheetList.forEach((s, idx) => {
    wbRelsXml += `\n  <Relationship Id="rId${idx + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${idx + 1}.xml"/>`;
  });
  wbRelsXml += `\n  <Relationship Id="rId${sheetList.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`;
  wbRelsXml += `\n</Relationships>`;

  let workbookXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>`;

  sheetList.forEach((s, idx) => {
    workbookXml += `\n    <sheet name="${escapeXml(s.name)}" sheetId="${idx + 1}" r:id="rId${idx + 1}"/>`;
  });
  workbookXml += `\n  </sheets>\n</workbook>`;

  const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="5">
    <font><sz val="11"/><color rgb="FF1E293B"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
    <font><b/><sz val="10"/><color rgb="FF166534"/><name val="Calibri"/></font>
    <font><b/><sz val="10"/><color rgb="FF991B1B"/><name val="Calibri"/></font>
    <font><b/><sz val="10"/><color rgb="FF854D0E"/><name val="Calibri"/></font>
  </fonts>
  <fills count="7">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF1B382B"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFF8FAFC"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFDCFCE7"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFEE2E2"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFEF3C7"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border>
      <left style="thin"><color rgb="FFE2E8F0"/></left>
      <right style="thin"><color rgb="FFE2E8F0"/></right>
      <top style="thin"><color rgb="FFE2E8F0"/></top>
      <bottom style="thin"><color rgb="FFE2E8F0"/></bottom>
      <diagonal/>
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="13">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center"/>
    </xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center"/>
    </xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1">
      <alignment horizontal="left" vertical="center"/>
    </xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1">
      <alignment horizontal="right" vertical="center"/>
    </xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1">
      <alignment horizontal="left" vertical="center"/>
    </xf>
    <xf numFmtId="0" fontId="2" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center"/>
    </xf>
    <xf numFmtId="0" fontId="3" fillId="5" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center"/>
    </xf>
    <xf numFmtId="0" fontId="4" fillId="6" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center"/>
    </xf>
    <xf numFmtId="0" fontId="0" fillId="3" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center"/>
    </xf>
    <xf numFmtId="0" fontId="0" fillId="3" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="left" vertical="center"/>
    </xf>
    <xf numFmtId="0" fontId="0" fillId="3" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="right" vertical="center"/>
    </xf>
    <xf numFmtId="0" fontId="0" fillId="3" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="left" vertical="center"/>
    </xf>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
</styleSheet>`;

  const files = [
    { name: '[Content_Types].xml', data: contentTypesXml },
    { name: '_rels/.rels', data: relsXml },
    { name: 'xl/_rels/workbook.xml.rels', data: wbRelsXml },
    { name: 'xl/workbook.xml', data: workbookXml },
    { name: 'xl/styles.xml', data: stylesXml }
  ];

  sheetList.forEach((s, idx) => {
    files.push({
      name: `xl/worksheets/sheet${idx + 1}.xml`,
      data: buildSheetXml(s.rows)
    });
  });

  return createZip(files);
}

// Server
const server = http.createServer(async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    res.writeHead(204);
    return res.end();
  }

  try {
    const url = new URL(req.url, `http://${req.headers.host || 'localhost'}`);
    const pathname = url.pathname;
    const method = req.method;

    // Web
    if (method === 'GET' && (pathname === '/' || pathname === '/index.html')) {
      return sendHtml(res, path.join(__dirname, 'public', 'index.html'));
    }

    // Status
    if (method === 'GET' && pathname === '/api/status') {
      const bRow = db.prepare('SELECT COUNT(*) as cnt FROM barang WHERE aktif = 1').get();
      const tRow = db.prepare('SELECT COUNT(*) as cnt FROM transaksi').get();
      const sRow = db.prepare(`
        SELECT COALESCE(SUM(CASE WHEN tipe IN ('AWAL', 'MASUK') THEN qty WHEN tipe = 'KELUAR' THEN -qty ELSE 0 END), 0) as total_stok
        FROM transaksi
      `).get();

      return sendJson(res, 200, {
        barang: bRow ? bRow.cnt : 0,
        transaksi: tRow ? tRow.cnt : 0,
        total_stok: sRow ? sRow.total_stok : 0
      });
    }

    // 6 Kategori Sheet
    if (method === 'GET' && pathname === '/api/kategori') {
      const rows = db.prepare(`
        SELECT k.id, k.nama,
          COUNT(b.id) as total_barang,
          COALESCE(SUM(s.stok), 0) as total_stok
        FROM kategori k
        LEFT JOIN barang b ON b.aktif = 1 AND b.subkategori_id = k.id
        LEFT JOIN (
          SELECT barang_id,
            COALESCE(SUM(CASE WHEN tipe IN ('AWAL', 'MASUK') THEN qty WHEN tipe = 'KELUAR' THEN -qty ELSE 0 END), 0) as stok
          FROM transaksi
          GROUP BY barang_id
        ) s ON b.id = s.barang_id
        GROUP BY k.id
        ORDER BY k.urutan ASC
      `).all();
      return sendJson(res, 200, rows);
    }

    // PIC Suggestions
    if (method === 'GET' && pathname === '/api/pic') {
      const rows = db.prepare(`
        SELECT pic, COUNT(*) as cnt
        FROM transaksi
        WHERE pic IS NOT NULL AND TRIM(pic) != '' AND pic != 'Sistem' AND pic != 'Saldo Awal'
        GROUP BY pic
        ORDER BY cnt DESC
        LIMIT 30
      `).all();
      return sendJson(res, 200, rows.map(r => r.pic));
    }

    // Barang per Kategori
    if (method === 'GET' && pathname === '/api/barang') {
      const search = url.searchParams.get('search') || '';
      const kategori = url.searchParams.get('kategori') || 'Kardus';

      let sql = `
        SELECT
          b.id, b.nama,
          k.nama as kategori_nama,
          COALESCE(s.stok, 0) as stok
        FROM barang b
        JOIN kategori k ON b.subkategori_id = k.id
        LEFT JOIN (
          SELECT barang_id,
            COALESCE(SUM(CASE WHEN tipe IN ('AWAL', 'MASUK') THEN qty WHEN tipe = 'KELUAR' THEN -qty ELSE 0 END), 0) as stok
          FROM transaksi
          GROUP BY barang_id
        ) s ON b.id = s.barang_id
        WHERE b.aktif = 1
      `;
      const params = [];

      if (kategori && kategori !== 'Semua') {
        sql += ` AND k.nama = ?`;
        params.push(kategori);
      }

      if (search) {
        sql += ` AND b.nama LIKE ?`;
        params.push(`%${search}%`);
      }

      sql += ` ORDER BY b.id ASC`;

      const rows = db.prepare(sql).all(...params);
      return sendJson(res, 200, rows);
    }

    // Tambah Barang Baru
    if (method === 'POST' && pathname === '/api/barang') {
      const body = await parseBody(req);
      const { kategori, nama, saldo_awal } = body;

      if (!nama || !nama.trim()) {
        return sendJson(res, 400, { error: 'Nama barang wajib diisi' });
      }

      const catRow = db.prepare('SELECT id FROM kategori WHERE nama = ?').get(kategori || 'Kardus');
      const catId = catRow ? catRow.id : 1;

      const insertRes = db.prepare(`
        INSERT INTO barang (nama, subkategori_id, aktif)
        VALUES (?, ?, 1)
      `).run(nama.trim(), catId);

      const newId = Number(insertRes.lastInsertRowid);
      const initialStock = parseQty(saldo_awal);
      if (initialStock > 0) {
        const today = new Date().toISOString().slice(0, 10);
        db.prepare(`
          INSERT INTO transaksi (barang_id, tanggal, tipe, qty, pic)
          VALUES (?, ?, 'AWAL', ?, 'Saldo Awal')
        `).run(newId, today, initialStock);
      }

      return sendJson(res, 201, { id: newId, message: 'Barang berhasil disimpan' });
    }

    // Input Mutasi (Masuk / Keluar)
    if (method === 'POST' && pathname === '/api/transaksi') {
      const body = await parseBody(req);
      const { barang_id, tanggal, tipe, qty, pic } = body;

      if (!barang_id || !tanggal || !tipe || !qty) {
        return sendJson(res, 400, { error: 'Semua kolom wajib diisi' });
      }

      const parsedQ = parseQty(qty);
      if (parsedQ <= 0) {
        return sendJson(res, 400, { error: 'Qty harus lebih dari 0' });
      }

      const bId = parseInt(barang_id, 10);
      const stockRow = db.prepare(`
        SELECT COALESCE(SUM(CASE WHEN tipe IN ('AWAL', 'MASUK') THEN qty WHEN tipe = 'KELUAR' THEN -qty ELSE 0 END), 0) as stok
        FROM transaksi WHERE barang_id = ?
      `).get(bId);
      const currentStock = stockRow ? stockRow.stok : 0;

      db.prepare(`
        INSERT INTO transaksi (barang_id, tanggal, tipe, qty, pic)
        VALUES (?, ?, ?, ?, ?)
      `).run(bId, String(tanggal), tipe, parsedQ, (pic || '').trim());

      const sisaStok = tipe === 'KELUAR' ? (currentStock - parsedQ) : (currentStock + parsedQ);

      return sendJson(res, 201, {
        message: `Mutasi ${tipe} berhasil dicatat`,
        sisa_stok: sisaStok
      });
    }

    // Riwayat Transaksi per Sheet Kategori
    if (method === 'GET' && pathname === '/api/transaksi') {
      const kategori = url.searchParams.get('kategori') || 'Kardus';
      const search = url.searchParams.get('search') || '';

      let sql = `
        SELECT
          t.id, t.tanggal, t.tipe, t.qty, t.pic,
          b.id as barang_id, b.nama,
          k.nama as kategori_nama
        FROM transaksi t
        JOIN barang b ON t.barang_id = b.id
        JOIN kategori k ON b.subkategori_id = k.id
        WHERE 1=1
      `;
      const params = [];

      if (kategori && kategori !== 'Semua') {
        sql += ` AND k.nama = ?`;
        params.push(kategori);
      }

      if (search) {
        sql += ` AND (b.nama LIKE ? OR t.pic LIKE ?)`;
        const s = `%${search}%`;
        params.push(s, s);
      }

      sql += ` ORDER BY t.tanggal DESC, t.id DESC LIMIT 150`;

      const rows = db.prepare(sql).all(...params);
      return sendJson(res, 200, rows);
    }

    // Export Excel (6 Sheet Terpisah Persis Sesuai Excel Asli)
    if (method === 'GET' && pathname === '/api/laporan') {
      const allRows = db.prepare(`
        SELECT
          t.id, t.tanggal, t.tipe, t.qty, t.pic,
          b.nama,
          k.nama as kategori_nama
        FROM transaksi t
        JOIN barang b ON t.barang_id = b.id
        JOIN kategori k ON b.subkategori_id = k.id
        ORDER BY t.tanggal DESC, t.id DESC
      `).all();

      const categories = ['Kardus', 'Nampan', 'Standing', 'Kantor', 'Pita', 'Box'];
      const rowsByCategory = {};
      categories.forEach(c => rowsByCategory[c] = []);

      allRows.forEach(r => {
        if (rowsByCategory[r.kategori_nama]) {
          rowsByCategory[r.kategori_nama].push(r);
        }
      });

      const xlsxBuf = generateMultiSheetXlsx(categories, rowsByCategory);
      const filename = `Stok_DQueen_${new Date().toISOString().slice(0, 10)}.xlsx`;

      res.writeHead(200, {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': `attachment; filename="${filename}"`,
        'Content-Length': xlsxBuf.length
      });
      return res.end(xlsxBuf);
    }

    // Backup Database
    if (method === 'POST' && pathname === '/api/backup') {
      const backupDir = path.join(__dirname, 'backup');
      if (!fs.existsSync(backupDir)) fs.mkdirSync(backupDir, { recursive: true });

      const todayStr = new Date().toISOString().slice(0, 10);
      const backupFile = path.join(backupDir, `stok-${todayStr}.db`);
      fs.copyFileSync(dbPath, backupFile);

      return sendJson(res, 200, {
        message: 'Cadangan berhasil disimpan',
        filename: path.basename(backupFile)
      });
    }

    sendJson(res, 404, { error: 'Endpoint tidak ditemukan' });
  } catch (err) {
    console.error('Server error:', err);
    sendJson(res, 500, { error: err.message || 'Kesalahan server' });
  }
});

server.listen(PORT, () => {
  console.log(`===============================================`);
  console.log(`  Stok DQueen — Atelier Persediaan`);
  console.log(`  Port: ${PORT}`);
  console.log(`  Akses di browser: http://localhost:${PORT}`);
  console.log(`===============================================`);
});
