const express = require('express');
const low = require('lowdb');
const FileSync = require('lowdb/adapters/FileSync');
const XLSX = require('xlsx');
const QRCode = require('qrcode');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;

// DB setup (JSON file)
const adapter = new FileSync('db.json');
const db = low(adapter);
db.defaults({ records: [], nextId: 1 }).write();

app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, 'public')));

// Submit form (student)
app.post('/api/submit', (req, res) => {
  const { prefix, fname, lname, dept, level, remark } = req.body;
  if (!prefix || !fname || !lname || !dept || !level) {
    return res.status(400).json({ ok: false, message: 'กรุณากรอกข้อมูลให้ครบ' });
  }
  const now = new Date();
  const pad = n => String(n).padStart(2, '0');
  const timestamp = `${now.getMonth()+1}/${now.getDate()}/${now.getFullYear()} ${pad(now.getHours())}:${pad(now.getMinutes())}:${pad(now.getSeconds())}`;
  const id = db.get('nextId').value();
  const record = { id, timestamp, prefix, fname, lname, dept, level, remark: remark || '', created_at: now.toISOString() };
  db.get('records').push(record).write();
  db.set('nextId', id + 1).write();
  res.json({ ok: true, id });
});

// Get all records
app.get('/api/records', (req, res) => {
  res.json(db.get('records').value());
});

// Delete one record
app.delete('/api/records/:id', (req, res) => {
  db.get('records').remove({ id: parseInt(req.params.id) }).write();
  res.json({ ok: true });
});

// Delete all records
app.delete('/api/records', (req, res) => {
  db.set('records', []).set('nextId', 1).write();
  res.json({ ok: true });
});

// Export Excel
app.get('/api/export', (req, res) => {
  const rows = db.get('records').value();
  const today = new Date();
  const thDate = `${today.getDate()}/${today.getMonth()+1}/${today.getFullYear()+543}`;
  const wsData = [
    [`สรุปรายชื่อนักศึกษามาสาย วันที่ ${thDate}`],
    ['ผู้รายงาน: นางสาวภัทร อรุณส่ง'],
    [],
    ['ลำดับที่','ประทับเวลา','คำนำหน้า','ชื่อ','นามสกุล','แผนกวิชา','ระดับชั้น','หมายเหตุ']
  ];
  rows.forEach((r, i) => wsData.push([i+1, r.timestamp, r.prefix, r.fname, r.lname, r.dept, r.level, r.remark || '']));
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(wsData);
  ws['!cols'] = [{wch:10},{wch:22},{wch:10},{wch:14},{wch:18},{wch:22},{wch:14},{wch:35}];
  ws['!merges'] = [{s:{r:0,c:0},e:{r:0,c:7}},{s:{r:1,c:0},e:{r:1,c:7}}];
  XLSX.utils.book_append_sheet(wb, ws, 'นักศึกษามาสาย');
  const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  const filename = `นักศึกษามาสาย_${today.getDate()}-${today.getMonth()+1}-${today.getFullYear()}.xlsx`;
  res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(filename)}`);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.send(buf);
});

// Generate QR
app.get('/api/qr', async (req, res) => {
  const host = process.env.RAILWAY_PUBLIC_DOMAIN
    ? `https://${process.env.RAILWAY_PUBLIC_DOMAIN}`
    : `http://localhost:${PORT}`;
  const url = `${host}/form`;
  try {
    const qr = await QRCode.toDataURL(url, { width: 400, margin: 2, color: { dark: '#0D2B55' } });
    res.json({ ok: true, qr, url });
  } catch(e) { res.status(500).json({ ok: false }); }
});

app.get('/form', (req, res) => res.sendFile(path.join(__dirname, 'public', 'form.html')));
app.get('/admin', (req, res) => res.sendFile(path.join(__dirname, 'public', 'admin.html')));
app.get('/', (req, res) => res.redirect('/admin'));

app.listen(PORT, () => console.log(`✅ Server on port ${PORT}`));
