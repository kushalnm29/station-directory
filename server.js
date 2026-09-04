const express = require('express');
const initSqlJs = require('sql.js');
const multer = require('multer');
const XLSX = require('xlsx');
const cors = require('cors');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = process.env.PORT || 3000;
const ADMIN_PIN = 'Kushal';
const DB_PATH = path.join(__dirname, 'stations.db');
const SEED_FILE = path.join(__dirname, 'seed-data.xlsx');

app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

const uploadDir = path.join(__dirname, 'uploads');
if (!fs.existsSync(uploadDir)) fs.mkdirSync(uploadDir);
const upload = multer({ dest: uploadDir });

var db;

function saveDB() {
  var data = db.export();
  var buffer = Buffer.from(data);
  fs.writeFileSync(DB_PATH, buffer);
}

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

function importExcelData(filePath) {
  var workbook = XLSX.readFile(filePath);
  var sheetName = workbook.SheetNames[0];
  var rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: '' });

  if (rows.length === 0) return 0;

  var excelCols = Object.keys(rows[0]);
  var colMap = {};
  excelCols.forEach(function(ec) { var mapped = normalizeColumn(ec); if (mapped) colMap[ec] = mapped; });

  console.log('Column mapping:', colMap);

  db.run('DELETE FROM stations');

  for (var i = 0; i < rows.length; i++) {
    var mapped = { station_code: '', location: '', state: '', zone: '', sm_name: '', sm_number: '', ctl_name: '', ctl_number: '', crm: '', com: '', zm: '' };
    for (var excelCol in colMap) { mapped[colMap[excelCol]] = String(rows[i][excelCol] || '').trim(); }
    db.run("INSERT INTO stations (station_code, location, state, zone, sm_name, sm_number, ctl_name, ctl_number, crm, com, zm) VALUES ($a,$b,$c,$d,$e,$f,$g,$h,$i,$j,$k)", {
      '$a': mapped.station_code, '$b': mapped.location, '$c': mapped.state, '$d': mapped.zone,
      '$e': mapped.sm_name, '$f': mapped.sm_number, '$g': mapped.ctl_name, '$h': mapped.ctl_number,
      '$i': mapped.crm, '$j': mapped.com, '$k': mapped.zm
    });
  }

  saveDB();
  return rows.length;
}

function getAll(sql, params) {
  var stmt = db.prepare(sql);
  if (params) stmt.bind(params);
  var rows = [];
  while (stmt.step()) { rows.push(stmt.getAsObject()); }
  stmt.free();
  return rows;
}

function getOne(sql, params) {
  var stmt = db.prepare(sql);
  if (params) stmt.bind(params);
  var row = null;
  if (stmt.step()) { row = stmt.getAsObject(); }
  stmt.free();
  return row;
}

function runSql(sql, params) {
  if (params) { db.run(sql, params); }
  else { db.run(sql); }
  saveDB();
}

initSqlJs().then(function(SQL) {
  if (fs.existsSync(DB_PATH)) {
    var fileBuffer = fs.readFileSync(DB_PATH);
    db = new SQL.Database(fileBuffer);
  } else {
    db = new SQL.Database();
  }

  db.run(`CREATE TABLE IF NOT EXISTS stations (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    station_code TEXT, location TEXT, state TEXT, zone TEXT,
    sm_name TEXT, sm_number TEXT, ctl_name TEXT, ctl_number TEXT,
    crm TEXT, com TEXT, zm TEXT,
    created_at TEXT DEFAULT (datetime('now')),
    updated_at TEXT DEFAULT (datetime('now'))
  )`);
  saveDB();

  // AUTO-SEED: If database is empty and seed file exists, import it
  var seedCount = getOne('SELECT COUNT(*) as c FROM stations').c;
  if (seedCount === 0 && fs.existsSync(SEED_FILE)) {
    console.log('📦 Database empty — auto-seeding from seed-data.xlsx...');
    var imported = importExcelData(SEED_FILE);
    console.log('✅ Auto-seeded ' + imported + ' stations from seed-data.xlsx');
  }

  // GET /api/stations
  app.get('/api/stations', function(req, res) {
    try {
      var search = req.query.search || '';
      var state = req.query.state || '';
      var com = req.query.com || '';
      var ctl = req.query.ctl || '';
      var page = parseInt(req.query.page) || 1;
      var limit = parseInt(req.query.limit) || 50;
      var where = [];
      var params = {};

      if (search) {
        where.push("(station_code LIKE $search OR location LIKE $search OR state LIKE $search OR sm_name LIKE $search OR ctl_name LIKE $search OR crm LIKE $search OR com LIKE $search OR zm LIKE $search OR sm_number LIKE $search OR ctl_number LIKE $search)");
        params['$search'] = '%' + search + '%';
      }
      if (state) { where.push('state = $state'); params['$state'] = state; }
      if (com) { where.push('com = $com'); params['$com'] = com; }
      if (ctl) { where.push('ctl_name = $ctl'); params['$ctl'] = ctl; }

      var whereClause = where.length ? 'WHERE ' + where.join(' AND ') : '';
      var offset = (page - 1) * limit;

      var countRow = getOne('SELECT COUNT(*) as total FROM stations ' + whereClause, params);
      var total = countRow ? countRow.total : 0;

      var queryParams = Object.assign({}, params);
      queryParams['$limit'] = limit;
      queryParams['$offset'] = offset;
      var rows = getAll('SELECT * FROM stations ' + whereClause + ' ORDER BY station_code ASC LIMIT $limit OFFSET $offset', queryParams);

      res.json({ stations: rows, total: total, page: page, totalPages: Math.ceil(total / limit) });
    } catch (err) { res.status(500).json({ error: err.message }); }
  });

  // GET /api/stations/:id
  app.get('/api/stations/:id', function(req, res) {
    try {
      var row = getOne('SELECT * FROM stations WHERE id = $id', { '$id': parseInt(req.params.id) });
      if (!row) return res.status(404).json({ error: 'Station not found' });
      res.json(row);
    } catch (err) { res.status(500).json({ error: err.message }); }
  });

  // POST /api/stations
  app.post('/api/stations', function(req, res) {
    try {
      var b = req.body;
      runSql("INSERT INTO stations (station_code, location, state, zone, sm_name, sm_number, ctl_name, ctl_number, crm, com, zm) VALUES ($a,$b,$c,$d,$e,$f,$g,$h,$i,$j,$k)", {
        '$a': b.station_code || '', '$b': b.location || '', '$c': b.state || '', '$d': b.zone || '',
        '$e': b.sm_name || '', '$f': b.sm_number || '', '$g': b.ctl_name || '', '$h': b.ctl_number || '',
        '$i': b.crm || '', '$j': b.com || '', '$k': b.zm || ''
      });
      res.json({ message: 'Station added' });
    } catch (err) { res.status(500).json({ error: err.message }); }
  });

  // PUT /api/stations/:id
  app.put('/api/stations/:id', function(req, res) {
    try {
      var fields = ['station_code','location','state','zone','sm_name','sm_number','ctl_name','ctl_number','crm','com','zm'];
      var updates = [];
      var params = { '$id': parseInt(req.params.id) };
      var idx = 0;
      fields.forEach(function(f) {
        if (req.body[f] !== undefined) {
          var key = '$v' + idx;
          updates.push(f + ' = ' + key);
          params[key] = req.body[f];
          idx++;
        }
      });
      if (updates.length === 0) return res.status(400).json({ error: 'No fields to update' });
      updates.push("updated_at = datetime('now')");
      runSql('UPDATE stations SET ' + updates.join(', ') + ' WHERE id = $id', params);
      res.json({ message: 'Station updated' });
    } catch (err) { res.status(500).json({ error: err.message }); }
  });

  // DELETE /api/stations/:id
  app.delete('/api/stations/:id', function(req, res) {
    try {
      runSql('DELETE FROM stations WHERE id = $id', { '$id': parseInt(req.params.id) });
      res.json({ message: 'Station deleted' });
    } catch (err) { res.status(500).json({ error: err.message }); }
  });

  // POST /api/verify-admin
  app.post('/api/verify-admin', function(req, res) {
    var pin = req.body.password || '';
    console.log('Login attempt:', pin, '| Expected:', ADMIN_PIN, '| Match:', pin === ADMIN_PIN);
    if (pin === ADMIN_PIN) {
      res.json({ success: true });
    } else {
      res.status(403).json({ success: false, error: 'Wrong password' });
    }
  });

  // POST /api/upload
  app.post('/api/upload', upload.single('file'), function(req, res) {
    try {
      var pin = req.query.key || '';
      if (pin !== ADMIN_PIN) {
        if (req.file && fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
        return res.status(403).json({ error: 'Access denied.' });
      }
      if (!req.file) return res.status(400).json({ error: 'No file uploaded' });

      console.log('Importing file:', req.file.originalname);
      var count = importExcelData(req.file.path);
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
  app.get('/api/stats', function(req, res) {
    try {
      var total = getOne('SELECT COUNT(*) as c FROM stations').c;
      var northCount = getOne('SELECT COUNT(*) as c FROM stations WHERE zone = "North"').c;
      var states = getAll('SELECT state, COUNT(*) as count FROM stations WHERE state != "" GROUP BY state ORDER BY count DESC');
      var coms = getAll('SELECT com, COUNT(*) as count FROM stations WHERE com != "" GROUP BY com ORDER BY count DESC');
      var ctls = getAll('SELECT ctl_name, COUNT(*) as count FROM stations WHERE ctl_name != "" AND LOWER(ctl_name) NOT LIKE "%no data%" GROUP BY ctl_name ORDER BY count DESC');
      res.json({ total: total, northCount: northCount, states: states, coms: coms, ctls: ctls });
    } catch (err) { res.status(500).json({ error: err.message }); }
  });

  // GET /api/export
  app.get('/api/export', function(req, res) {
    try {
      var rows = getAll('SELECT * FROM stations ORDER BY station_code');
      res.setHeader('Content-Disposition', 'attachment; filename=stations_export.json');
      res.json(rows);
    } catch (err) { res.status(500).json({ error: err.message }); }
  });

  // GET /api/filters
  app.get('/api/filters', function(req, res) {
    try {
      var states = getAll('SELECT DISTINCT state FROM stations WHERE state != "" ORDER BY state').map(function(r) { return r.state; });
      var coms = getAll('SELECT DISTINCT com FROM stations WHERE com != "" ORDER BY com').map(function(r) { return r.com; });
      var ctls = getAll('SELECT DISTINCT ctl_name FROM stations WHERE ctl_name != "" AND LOWER(ctl_name) NOT LIKE "%no data%" ORDER BY ctl_name').map(function(r) { return r.ctl_name; });
      res.json({ states: states, coms: coms, ctls: ctls });
    } catch (err) { res.status(500).json({ error: err.message }); }
  });

  app.listen(PORT, function() {
    var count = getOne('SELECT COUNT(*) as c FROM stations').c;
    console.log('\n🚀 Station Directory running at http://localhost:' + PORT);
    console.log('📊 Stations in DB: ' + count);
    console.log('🔐 Admin pin: ' + ADMIN_PIN);
  });
});