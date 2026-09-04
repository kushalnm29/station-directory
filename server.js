const express = require('express');
const { createClient } = require('@libsql/client');
const multer = require('multer');
const XLSX = require('xlsx');
const cors = require('cors');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = process.env.PORT || 3000;
const ADMIN_PIN = 'Kushal';

// Turso Cloud Database - REPLACE THESE WITH YOUR VALUES
const dbClient = createClient({
  url: process.env.TURSO_URL || 'YOUR_TURSO_URL_HERE',
  authToken: process.env.TURSO_TOKEN || 'YOUR_TURSO_TOKEN_HERE'
});

app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

const uploadDir = path.join(__dirname, 'uploads');
if (!fs.existsSync(uploadDir)) fs.mkdirSync(uploadDir);
const upload = multer({ dest: uploadDir });

function normalizeColumn(col) {
  var c = col.toString().trim().toLowerCase().replace(/[\s\-\/]+/g, '_');
  var map = {
    'station_code': 'station_code', 'station': 'station_code', 'code': 'station_code',
    'location': 'location', 'city': 'location',
    'state': 'state',
    'zone': 'zone',
    'station_manager_name': 'sm_name', 'sm_name': 'sm_name', 'sm': 'sm_name', 'station_manager': 'sm_name',
    'station_manager_number': 'sm_number', 'sm_number': 'sm_number',
    'ctl': 'ctl_name', 'ctl_name': 'ctl_name',
    'ctl_number': 'ctl_number',
    'crm': 'crm', 'crm_name': 'crm',
    'com': 'com', 'com_name': 'com',
    'zm': 'zm', 'zm_name': 'zm', 'zone_manager': 'zm'
  };
  return map[c] || null;
}

async function importExcelData(filePath) {
  var workbook = XLSX.readFile(filePath);
  var sheetName = workbook.SheetNames[0];
  var rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: '' });

  if (rows.length === 0) return 0;

  var excelCols = Object.keys(rows[0]);
  var colMap = {};
  excelCols.forEach(function(ec) { var mapped = normalizeColumn(ec); if (mapped) colMap[ec] = mapped; });

  console.log('Column mapping:', colMap);

  await dbClient.execute('DELETE FROM stations');

  for (var i = 0; i < rows.length; i++) {
    var mapped = { station_code: '', location: '', state: '', zone: '', sm_name: '', sm_number: '', ctl_name: '', ctl_number: '', crm: '', com: '', zm: '' };
    for (var excelCol in colMap) { mapped[colMap[excelCol]] = String(rows[i][excelCol] || '').trim(); }
    await dbClient.execute({
      sql: "INSERT INTO stations (station_code, location, state, zone, sm_name, sm_number, ctl_name, ctl_number, crm, com, zm) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
      args: [mapped.station_code, mapped.location, mapped.state, mapped.zone, mapped.sm_name, mapped.sm_number, mapped.ctl_name, mapped.ctl_number, mapped.crm, mapped.com, mapped.zm]
    });
  }

  return rows.length;
}

async function startServer() {
  // Create table
  await dbClient.execute(`
    CREATE TABLE IF NOT EXISTS stations (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      station_code TEXT, location TEXT, state TEXT, zone TEXT,
      sm_name TEXT, sm_number TEXT, ctl_name TEXT, ctl_number TEXT,
      crm TEXT, com TEXT, zm TEXT,
      created_at TEXT DEFAULT (datetime('now')),
      updated_at TEXT DEFAULT (datetime('now'))
    )
  `);

  // Auto-seed if database is empty and seed file exists
  var countResult = await dbClient.execute('SELECT COUNT(*) as c FROM stations');
  var seedCount = countResult.rows[0].c;
  var seedFile = path.join(__dirname, 'seed-data.xlsx');

  if (seedCount === 0 && fs.existsSync(seedFile)) {
    console.log('📦 Database empty — auto-seeding from seed-data.xlsx...');
    var imported = await importExcelData(seedFile);
    console.log('✅ Auto-seeded ' + imported + ' stations');
  }

  // GET /api/stations
  app.get('/api/stations', async function(req, res) {
    try {
      var search = req.query.search || '';
      var com = req.query.com || '';
      var ctl = req.query.ctl || '';
      var page = parseInt(req.query.page) || 1;
      var limit = parseInt(req.query.limit) || 50;
      var where = [];
      var args = [];

      if (search) {
        where.push("(station_code LIKE ? OR location LIKE ? OR state LIKE ? OR sm_name LIKE ? OR ctl_name LIKE ? OR crm LIKE ? OR com LIKE ? OR zm LIKE ? OR sm_number LIKE ? OR ctl_number LIKE ?)");
        var s = '%' + search + '%';
        args.push(s, s, s, s, s, s, s, s, s, s);
      }
      if (com) { where.push('com = ?'); args.push(com); }
      if (ctl) { where.push('ctl_name = ?'); args.push(ctl); }

      var whereClause = where.length ? 'WHERE ' + where.join(' AND ') : '';
      var offset = (page - 1) * limit;

      var countResult = await dbClient.execute({ sql: 'SELECT COUNT(*) as total FROM stations ' + whereClause, args: args });
      var total = countResult.rows[0].total;

      var dataArgs = args.slice();
      dataArgs.push(limit, offset);
      var dataResult = await dbClient.execute({ sql: 'SELECT * FROM stations ' + whereClause + ' ORDER BY station_code ASC LIMIT ? OFFSET ?', args: dataArgs });

      res.json({ stations: dataResult.rows, total: total, page: page, totalPages: Math.ceil(total / limit) });
    } catch (err) { console.error(err); res.status(500).json({ error: err.message }); }
  });

  // GET /api/stations/:id
  app.get('/api/stations/:id', async function(req, res) {
    try {
      var result = await dbClient.execute({ sql: 'SELECT * FROM stations WHERE id = ?', args: [parseInt(req.params.id)] });
      if (result.rows.length === 0) return res.status(404).json({ error: 'Station not found' });
      res.json(result.rows[0]);
    } catch (err) { res.status(500).json({ error: err.message }); }
  });

  // POST /api/stations
  app.post('/api/stations', async function(req, res) {
    try {
      var b = req.body;
      await dbClient.execute({
        sql: "INSERT INTO stations (station_code, location, state, zone, sm_name, sm_number, ctl_name, ctl_number, crm, com, zm) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
        args: [b.station_code || '', b.location || '', b.state || '', b.zone || '', b.sm_name || '', b.sm_number || '', b.ctl_name || '', b.ctl_number || '', b.crm || '', b.com || '', b.zm || '']
      });
      res.json({ message: 'Station added' });
    } catch (err) { res.status(500).json({ error: err.message }); }
  });

  // PUT /api/stations/:id
  app.put('/api/stations/:id', async function(req, res) {
    try {
      var fields = ['station_code','location','state','zone','sm_name','sm_number','ctl_name','ctl_number','crm','com','zm'];
      var updates = [];
      var args = [];
      fields.forEach(function(f) {
        if (req.body[f] !== undefined) { updates.push(f + ' = ?'); args.push(req.body[f]); }
      });
      if (updates.length === 0) return res.status(400).json({ error: 'No fields to update' });
      updates.push("updated_at = datetime('now')");
      args.push(parseInt(req.params.id));
      await dbClient.execute({ sql: 'UPDATE stations SET ' + updates.join(', ') + ' WHERE id = ?', args: args });
      res.json({ message: 'Station updated' });
    } catch (err) { res.status(500).json({ error: err.message }); }
  });

  // DELETE /api/stations/:id
  app.delete('/api/stations/:id', async function(req, res) {
    try {
      await dbClient.execute({ sql: 'DELETE FROM stations WHERE id = ?', args: [parseInt(req.params.id)] });
      res.json({ message: 'Station deleted' });
    } catch (err) { res.status(500).json({ error: err.message }); }
  });

  // POST /api/verify-admin
  app.post('/api/verify-admin', function(req, res) {
    var pin = req.body.password || '';
    console.log('Login attempt:', pin, '| Match:', pin === ADMIN_PIN);
    if (pin === ADMIN_PIN) {
      res.json({ success: true });
    } else {
      res.status(403).json({ success: false, error: 'Wrong password' });
    }
  });

  // POST /api/upload
  app.post('/api/upload', upload.single('file'), async function(req, res) {
    try {
      var pin = req.query.key || '';
      if (pin !== ADMIN_PIN) {
        if (req.file && fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
        return res.status(403).json({ error: 'Access denied.' });
      }
      if (!req.file) return res.status(400).json({ error: 'No file uploaded' });

      console.log('Importing file:', req.file.originalname);
      var count = await importExcelData(req.file.path);
      fs.unlinkSync(req.file.path);

      console.log('Import complete:', count, 'stations');
      res.json({ message: 'Successfully imported ' + count + ' stations', count: count });
    } catch (err) {
      console.error('Upload error:', err);
      if (req.file && fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
      res.status(500).json({ error: err.message });
    }
  });

  // GET /api/stats
  app.get('/api/stats', async function(req, res) {
    try {
      var totalR = await dbClient.execute('SELECT COUNT(*) as c FROM stations');
      var northR = await dbClient.execute('SELECT COUNT(*) as c FROM stations WHERE zone = "North"');
      var statesR = await dbClient.execute('SELECT state, COUNT(*) as count FROM stations WHERE state != "" GROUP BY state ORDER BY count DESC');
      var comsR = await dbClient.execute('SELECT com, COUNT(*) as count FROM stations WHERE com != "" GROUP BY com ORDER BY count DESC');
      var ctlsR = await dbClient.execute('SELECT ctl_name, COUNT(*) as count FROM stations WHERE ctl_name != "" AND LOWER(ctl_name) NOT LIKE "%no data%" GROUP BY ctl_name ORDER BY count DESC');
      res.json({ total: totalR.rows[0].c, northCount: northR.rows[0].c, states: statesR.rows, coms: comsR.rows, ctls: ctlsR.rows });
    } catch (err) { res.status(500).json({ error: err.message }); }
  });

  // GET /api/export
  app.get('/api/export', async function(req, res) {
    try {
      var result = await dbClient.execute('SELECT * FROM stations ORDER BY station_code');
      res.setHeader('Content-Disposition', 'attachment; filename=stations_export.json');
      res.json(result.rows);
    } catch (err) { res.status(500).json({ error: err.message }); }
  });

  // GET /api/filters
  app.get('/api/filters', async function(req, res) {
    try {
      var statesR = await dbClient.execute('SELECT DISTINCT state FROM stations WHERE state != "" ORDER BY state');
      var comsR = await dbClient.execute('SELECT DISTINCT com FROM stations WHERE com != "" ORDER BY com');
      var ctlsR = await dbClient.execute('SELECT DISTINCT ctl_name FROM stations WHERE ctl_name != "" AND LOWER(ctl_name) NOT LIKE "%no data%" ORDER BY ctl_name');
      res.json({
        states: statesR.rows.map(function(r) { return r.state; }),
        coms: comsR.rows.map(function(r) { return r.com; }),
        ctls: ctlsR.rows.map(function(r) { return r.ctl_name; })
      });
    } catch (err) { res.status(500).json({ error: err.message }); }
  });

  app.listen(PORT, async function() {
    var countResult = await dbClient.execute('SELECT COUNT(*) as c FROM stations');
    console.log('\n🚀 Station Directory running at http://localhost:' + PORT);
    console.log('📊 Stations in DB: ' + countResult.rows[0].c);
    console.log('🔐 Admin pin: ' + ADMIN_PIN);
    console.log('☁️  Database: Turso Cloud (permanent storage)');
  });
}

startServer().catch(function(err) {
  console.error('Failed to start:', err);
  process.exit(1);
});