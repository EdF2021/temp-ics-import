const https = require('https');
const fs = require('fs');

const URL = 'https://outlook.office365.com/owa/calendar/0f17d7dd19d54501bd794b502804a3c8@surf.nl/ef608d548b7f48228037c262897bf2029943935942078287197/calendar.ics';

function fetchText(u) {
  return new Promise((resolve, reject) => {
    https.get(u, r => {
      if (r.statusCode !== 200) return reject(new Error('HTTP ' + r.statusCode));
      let d = '';
      r.on('data', c => d += c);
      r.on('end', () => resolve(d));
    }).on('error', reject);
  });
}

function unfold(text) {
  const lines = text.replace(/\r\n/g, '\n').split('\n');
  const out = [];
  for (const l of lines) {
    if ((l.startsWith(' ') || l.startsWith('\t')) && out.length) out[out.length - 1] += l.slice(1);
    else out.push(l);
  }
  return out;
}

function parse(text) {
  const lines = unfold(text);
  const events = [];
  let cur = null;
  for (const l of lines) {
    if (l.startsWith('BEGIN:VEVENT')) { cur = {}; continue; }
    if (l.startsWith('END:VEVENT')) { if (cur) events.push(cur); cur = null; continue; }
    if (!cur) continue;
    const ci = l.indexOf(':'); if (ci < 0) continue;
    const name = l.slice(0, ci), val = l.slice(ci + 1);
    const base = name.split(';')[0].toUpperCase();
    if (base === 'UID') cur.uid = val;
    else if (base === 'SUMMARY') cur.summary = val;
    else if (base === 'LOCATION') cur.location = val;
    else if (base === 'RRULE') cur.rrule = val;
    else if (base === 'STATUS') cur.status = val;
    else if (base === 'DTSTART') cur.dtstartRaw = l;
    else if (base === 'DTEND') cur.dtendRaw = l;
    else if (base === 'RECURRENCE-ID') cur.recurrenceIdRaw = l;
  }
  return events;
}

function conv(raw) {
  if (!raw) return null;
  const ci = raw.indexOf(':');
  const name = raw.slice(0, ci), val = raw.slice(ci + 1);
  const tz = (name.split(';').find(p => p.startsWith('TZID=')) || '').slice(5);
  if (val.length === 8) return { date: val.slice(0,4)+'-'+val.slice(4,6)+'-'+val.slice(6,8), allDay: true };
  const clean = val.replace(/Z$/, '');
  const iso = clean.slice(0,4)+'-'+clean.slice(4,6)+'-'+clean.slice(6,8)+'T'+clean.slice(9,11)+':'+clean.slice(11,13)+':'+clean.slice(13,15);
  if (val.endsWith('Z')) return { date: iso + 'Z', allDay: false };
  const isWEurope = tz.indexOf('W. Europe') === 0 || tz === 'CET' || tz === 'CEST';
  if (isWEurope) {
    const y = +clean.slice(0,4), mo = +clean.slice(4,6) - 1, dd = +clean.slice(6,8), t = clean.slice(9,15);
    const lastSun = m => { const d = new Date(Date.UTC(y, m + 1, 0)); return d.getUTCDay() === 0 ? d.getUTCDate() : d.getUTCDate() - d.getUTCDay(); };
    let dst;
    if (mo > 2 && mo < 9) dst = true;
    else if (mo === 2) dst = dd > lastSun(2) || (dd === lastSun(2) && t >= '020000');
    else if (mo === 9) dst = !(dd > lastSun(9) || (dd === lastSun(9) && t >= '030000'));
    else dst = false;
    const off = dst ? 120 : 60;
    const d2 = new Date(iso + 'Z');
    d2.setUTCMinutes(d2.getUTCMinutes() - off);
    return { date: d2.toISOString().replace('.000Z', 'Z'), allDay: false, converted: true };
  }
  return { date: iso, allDay: false, note: 'unconverted ' + tz };
}

function occ(orig, newD) {
  const day = newD.toISOString().slice(0, 10);
  return { date: day + orig.date.slice(10), allDay: false };
}

function expand(ev) {
  const s = ev.dtstart;
  if (s.allDay || !ev.rrule) return [];
  const freq = /FREQ=([A-Z]+)/.exec(ev.rrule);
  const interval = /INTERVAL=(\d+)/.exec(ev.rrule);
  const until = /UNTIL=(\d{8})/.exec(ev.rrule);
  const byday = /BYDAY=([A-Z, ]+)/.exec(ev.rrule);
  if (!freq || !interval || !byday) return [];
  const map = { MO: 1, TU: 2, WE: 3, TH: 4, FR: 5, SA: 6, SU: 0 };
  const days = byday[1].split(',').map(x => map[x.trim()]).filter(x => x !== undefined);
  if (freq[1] !== 'WEEKLY' || !days.length) return [];
  const d0 = new Date(s.date);
  const end = until ? new Date(until[1].slice(0,4)+'-'+until[1].slice(4,6)+'-'+until[1].slice(6,8)+'T23:59:59Z') : new Date('2026-12-31T23:59:59Z');
  const out = [];
  const d = new Date(d0);
  let guard = 0;
  while (d <= end && guard++ < 700) {
    if (days.includes(d.getUTCDay())) {
      const dayDiff = Math.round((d - d0) / 86400000);
      if (dayDiff % (7 * +interval[1]) === 0) out.push(occ(s, d));
    }
    d.setUTCDate(d.getUTCDate() + 1);
  }
  return out;
}

(async () => {
  const text = await fetchText(URL);
  fs.writeFileSync('raw.ics', text);
  const evs = parse(text).filter(e => e.status !== 'CANCELLED');
  const all = [];
  const exceptions = [];
  for (const e of evs) {
    const item = { uid: (e.uid || '').slice(-24), summary: e.summary, dtstart: conv(e.dtstartRaw), dtend: conv(e.dtendRaw), location: e.location || '' };
    if (e.recurrenceIdRaw) { item.recurrenceId = conv(e.recurrenceIdRaw); exceptions.push(item); }
    else all.push(item);
  }
  const masterKeys = new Set(all.map(e => e.summary));
  const final = [];
  for (const e of all) {
    if (e.rrule) { for (const o of expand(e)) final.push(Object.assign({}, e, { dtstart: o, rrule: undefined })); }
    else final.push(e);
  }
  for (const x of exceptions) {
    final.push(x);
    for (let i = final.length - 1; i >= 0; i--) {
      const f = final[i];
      if (f.summary === x.summary && !f.recurrenceId && f.dtstart.date.slice(0, 10) === x.recurrenceId.date.slice(0, 10)) final.splice(i, 1);
    }
  }
  const seen = new Set();
  const dedup = [];
  for (const e of final.sort((a, b) => (a.dtstart.date < b.dtstart.date ? -1 : 1))) {
    const k = e.summary + '|' + e.dtstart.date + '|' + (e.dtend ? e.dtend.date : '');
    if (seen.has(k)) continue;
    seen.add(k); dedup.push(e);
  }
  fs.writeFileSync('surf-events.json', JSON.stringify(dedup, null, 1));
  console.log('TOTAL ' + dedup.length);
  for (const e of dedup) {
    const t = e.dtstart.allDay ? e.dtstart.date : e.dtstart.date;
    const t2 = e.dtend ? e.dtend.date : '';
    console.log('EVENT ' + t + ' .. ' + t2 + ' | ' + e.summary + ' | ' + e.location);
  }
})().catch(e => { console.error(e); process.exit(1); });