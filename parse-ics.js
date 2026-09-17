const fs = require('fs');
const https = require('https');
function get(url) {
  return new Promise((resolve, reject) => {
    https.get(url, r => {
      if (r.statusCode >= 300 && r.statusCode < 400 && r.headers.location) {
        r.resume();
        get(r.headers.location).then(resolve, reject);
        return;
      }
      if (r.statusCode >= 400) { r.resume(); reject(new Error('HTTP ' + r.statusCode)); return; }
      let data = '';
      r.on('data', c => data += c);
      r.on('end', () => resolve(data));
    }).on('error', reject);
  });
}
function unfold(t) { return t.replace(/\r\n[ \t]/g, '').replace(/\n[ \t]/g, ''); }
function parseICS(text) {
  const lines = unfold(text).split(/\r?\n/);
  const events = [];
  let cur = null;
  for (const line of lines) {
    if (line.indexOf('BEGIN:VEVENT') === 0) cur = { raw: [] };
    else if (line.indexOf('END:VEVENT') === 0) { if (cur) events.push(cur); cur = null; }
    else if (cur) cur.raw.push(line);
  }
  return events.map(ev => {
    const o = {};
    for (const line of ev.raw) {
      const idx = line.indexOf(':');
      if (idx < 0) continue;
      const key = line.slice(0, idx);
      const val = line.slice(idx + 1);
      const k = key.split(';')[0].toUpperCase();
      if (k === 'SUMMARY') o.summary = val;
      else if (k === 'LOCATION') o.location = val;
      else if (k === 'UID') o.uid = val;
      else if (k === 'DESCRIPTION') o.description = val;
      else if (k === 'DTSTART') { o.dtstart = val; const m = key.match(/TZID=([^;]+)/); if (m) o.tzid = m[1]; }
      else if (k === 'DTEND') o.dtend = val;
      else if (k === 'RRULE') o.rrule = val;
      else if (k === 'RECURRENCE-ID') o.recurrenceId = val;
      else if (k === 'STATUS') o.status = val;
    }
    o.allDay = /^\d{8}$/.test(o.dtstart || '');
    return o;
  });
}
(async () => {
  const ics = await get(process.env.ICS_URL);
  fs.writeFileSync('raw.ics', ics);
  const evs = parseICS(ics);
  const json = JSON.stringify(evs);
  fs.writeFileSync('surf-events.json', json);
  fs.rmSync('parts', { recursive: true, force: true });
  fs.mkdirSync('parts', { recursive: true });
  const size = 6000;
  let i = 0;
  for (let off = 0; off < ics.length; off += size) { i++; fs.writeFileSync('parts/raw-' + String(i).padStart(2, '0') + '.txt', ics.slice(off, off + size)); }
  let j = 0;
  for (let off = 0; off < json.length; off += size) { j++; fs.writeFileSync('parts/ev-' + String(j).padStart(2, '0') + '.json', json.slice(off, off + size)); }
  fs.writeFileSync('parts/count.txt', 'raw=' + i + ' ev=' + j + ' bytes=' + ics.length);
  console.log('events:', evs.length, 'bytes:', ics.length);
})().catch(e => { console.error(e); process.exit(1); });