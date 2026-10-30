// Snapshot of production deterministic case parsing, without model calls.
// Source: Step 2 Draft & Review febec03b-26c5-4442-98d7-dfe5a67d81f5.
function caseContext(fields) {
const f = fields;

const toStr = (v) => (v == null ? "" : String(v).trim());
const toYesNo = (v) => {
  const s = String(v ?? "").trim().toLowerCase();
  if (["yes", "true", "y", "1"].includes(s)) return "Yes";
  if (["no", "false", "n", "0"].includes(s)) return "No";
  return toStr(v);
};

// --- County lookup from city or zip ---
function lookupCounty(address, cityStateZip) {
  const cityCountyMap = {
    'san jose': 'Santa Clara County',
    'santa clara': 'Santa Clara County',
    'milpitas': 'Santa Clara County',
    'sunnyvale': 'Santa Clara County',
    'mountain view': 'Santa Clara County',
    'palo alto': 'Santa Clara County',
    'campbell': 'Santa Clara County',
    'los gatos': 'Santa Clara County',
    'los altos': 'Santa Clara County',
    'cupertino': 'Santa Clara County',
    'saratoga': 'Santa Clara County',
    'gilroy': 'Santa Clara County',
    'morgan hill': 'Santa Clara County',
    'fremont': 'Alameda County',
    'oakland': 'Alameda County',
    'berkeley': 'Alameda County',
    'san leandro': 'Alameda County',
    'hayward': 'Alameda County',
    'union city': 'Alameda County',
    'newark': 'Alameda County',
    'pleasanton': 'Alameda County',
    'livermore': 'Alameda County',
    'dublin': 'Alameda County',
    'san francisco': 'San Francisco County',
    'san mateo': 'San Mateo County',
    'redwood city': 'San Mateo County',
    'daly city': 'San Mateo County',
    'south san francisco': 'San Mateo County',
    'east palo alto': 'San Mateo County',
    'burlingame': 'San Mateo County',
    'menlo park': 'San Mateo County',
    'foster city': 'San Mateo County',
    'san carlos': 'San Mateo County',
    'belmont': 'San Mateo County',
    'millbrae': 'San Mateo County',
    'alameda': 'Alameda County',
    'emeryville': 'Alameda County',
    'albany': 'Alameda County',
    'santa cruz': 'Santa Cruz County',
    'scotts valley': 'Santa Cruz County',
    'capitola': 'Santa Cruz County',
    'watsonville': 'Santa Cruz County',
    'aptos': 'Santa Cruz County',
    'walnut creek': 'Contra Costa County',
    'concord': 'Contra Costa County',
    'richmond': 'Contra Costa County',
    'antioch': 'Contra Costa County',
    'pittsburg': 'Contra Costa County',
    'martinez': 'Contra Costa County',
    'san ramon': 'Contra Costa County',
    'danville': 'Contra Costa County',
    'lafayette': 'Contra Costa County',
    'orinda': 'Contra Costa County',
    'hollister': 'San Benito County',
    'san juan bautista': 'San Benito County'
  };

  const combined = (cityStateZip || '').trim();
  const cityFromCSZ = combined.split(',')[0]?.trim().toLowerCase() || '';
  if (cityFromCSZ && cityCountyMap[cityFromCSZ]) return cityCountyMap[cityFromCSZ];

  const addrParts = (address || '').split(',');
  if (addrParts.length >= 2) {
    const cityPart = addrParts[1]?.trim().replace(/\s+(CA|ca|California)\s+\d{5}.*$/i, '').trim().toLowerCase();
    if (cityPart && cityCountyMap[cityPart]) return cityCountyMap[cityPart];
  }

  const allText = (address || '') + ' ' + (cityStateZip || '');
  const zipMatch = allText.match(/\b(9[0-9]{4})\b/);
  if (zipMatch) {
    const zip = parseInt(zipMatch[1], 10);
    if (zip >= 95110 && zip <= 95141) return 'Santa Clara County';
  }

  return '';
}

// --- Parse compound tenant address field ---
// Airtable 'Tenant’s Property' (base appg3kN6Z3JNX7bnF, table tblZbUGz8OTvFNh9i)
// holds the full address as one string, e.g. '123 Main St Santa Cruz CA 98598'.
// Parse into discrete parts so the tenant CSZ line comes from this field ONLY.
function parseTenantProperty(raw) {
  const result = { street: '', city: '', state: '', zip: '', city_state_zip: '' };
  if (!raw) return result;
  let s = String(raw).replace(/\r\n|\n/g, ' ').replace(/\s+/g, ' ').trim();
  let m = s.match(/^(.*?),\s*([^,]+),\s*([A-Za-z]{2})\s+(\d{5}(?:-\d{4})?)$/);
  if (m) {
    result.street = m[1].trim(); result.city = m[2].trim();
    result.state = m[3].toUpperCase(); result.zip = m[4];
    result.city_state_zip = result.city + ', ' + result.state + ' ' + result.zip;
    return result;
  }
  const CA_CITIES = ['East Palo Alto','Palo Alto','Los Altos Hills','Los Altos','Los Gatos','Monte Sereno','Morgan Hill','Mountain View','San Juan Bautista','San Jose','San José','Santa Clara','Santa Cruz','Scotts Valley','San Mateo','San Carlos','San Bruno','San Francisco','San Leandro','San Ramon','Union City','Daly City','South San Francisco','Foster City','Half Moon Bay','Menlo Park','Redwood City','Walnut Creek','Sunnyvale','Saratoga','Milpitas','Campbell','Cupertino','Gilroy','Watsonville','Capitola','Aptos','Fremont','Hayward','Oakland','Berkeley','Alameda','Pleasanton','Dublin','Livermore','Newark','Emeryville','Albany','Hollister','Belmont','Millbrae','Burlingame','Atherton','Portola Valley','Woodside','Hillsborough','Brisbane','Colma','Pacifica','Antioch','Concord','Richmond','Pittsburg','Martinez','Danville','Lafayette','Orinda'].sort((a,b)=>b.length-a.length);
  for (const c of CA_CITIES) {
    const re = new RegExp('^(.+?)\\s+' + c.replace(/[.*+?^${}()|[\\]\\\\]/g, '\\$&') + '\\s+(?:CA|California)\\s+(\\d{5}(?:-\\d{4})?)\\s*$', 'i');
    m = s.match(re);
    if (m) {
      result.street = m[1].trim(); result.city = c; result.state = 'CA'; result.zip = m[2];
      result.city_state_zip = c + ', CA ' + m[2];
      return result;
    }
  }
  m = s.match(/^(.+?)\s+(?:CA|California)\s+(\d{5}(?:-\d{4})?)\s*$/i);
  if (m) {
    const left = m[1].trim();
    const parts = left.split(/\s+/);
    if (parts.length >= 2) { result.city = parts.slice(-1).join(' '); result.street = parts.slice(0,-1).join(' ').trim(); }
    else { result.street = left; }
    result.state = 'CA'; result.zip = m[2];
    result.city_state_zip = (result.city ? result.city + ', ' : '') + 'CA ' + m[2];
    return result;
  }
  result.street = s;
  return result;
}
let _rawTenantProp = "";
let _tenantPropFieldUsed = "";
// Apostrophe-agnostic key match: straight ', curly ’, backtick, or missing
for (const _k of Object.keys(f)) {
  const _norm = String(_k).toLowerCase().replace(/[‘’ʼ`´]/g, "'").replace(/\s+/g, ' ').trim();
  if (/^tenant'?s?\s+property(\s+address)?$/.test(_norm)) {
    const _v = toStr(f[_k]);
    if (_v) { _rawTenantProp = _v; _tenantPropFieldUsed = _k; break; }
  }
}
if (!_rawTenantProp) {
  const _fallback = ["Tenant's Property Address", "Tenant’s Property Address", "Tenant's Property", "Tenant’s Property", "Tenants Property Address", "Tenants Property", "tenant's property address", "Tenant Property Address"];
  for (const _k of _fallback) { const _v = toStr(f[_k]); if (_v) { _rawTenantProp = _v; _tenantPropFieldUsed = _k; break; } }
}
console.log('[build context] tenant lookup:', JSON.stringify({ fieldUsed: _tenantPropFieldUsed, raw: _rawTenantProp, tenantKeys: Object.keys(f).filter(k => /tenant/i.test(k)) }));
// Configured review path: parse the tenant's own address without borrowing the landlord's city.
function parseConfiguredTenantProperty(raw) {
  const blank = { street: '', city: '', state: '', zip: '', city_state_zip: '' };
  const text = String(raw || '').replace(/\s*[\r\n]+\s*/g, ', ').replace(/\s+/g, ' ').trim();
  if (!text) return blank;
  const tail = text.match(/^(.*?)(?:,\s*|\s+)(CA|California)\.?(?:\s*,\s*|\s*)(\d{5}(?:-\d{4})?)?\s*$/i);
  if (!tail) return { ...blank, street: text };
  const prefix = tail[1].trim().replace(/,+$/, '').trim();
  let street = '', city = '';
  const parts = prefix.split(',').map(s => s.trim()).filter(Boolean);
  if (parts.length >= 2) {
    city = parts.pop();
    street = parts.join(', ');
  } else {
    const cities = ['East Palo Alto','South San Francisco','Los Altos Hills','San Juan Bautista','Mountain View','Morgan Hill','Redwood City','San Francisco','Santa Clara','Santa Cruz','Scotts Valley','San Mateo','San Carlos','San Jose','San José','Los Gatos','Los Altos','San Leandro','San Ramon','Union City','Daly City','Foster City','Half Moon Bay','Menlo Park','Walnut Creek','Sunnyvale','Saratoga','Milpitas','Campbell','Cupertino','Gilroy','Watsonville','Capitola','Aptos','Fremont','Hayward','Oakland','Berkeley','Alameda','Pleasanton','Dublin','Livermore','Newark','Emeryville','Albany','Hollister','Belmont','Millbrae','Burlingame','Atherton','Portola Valley','Woodside','Hillsborough','Brisbane','Colma','Pacifica','Antioch','Concord','Richmond','Pittsburg','Martinez','Danville','Lafayette','Orinda','Monte Sereno','San Bruno'].sort((a,b) => b.length-a.length);
    for (const name of cities) {
      const m = prefix.match(new RegExp('^(.+?)\\s+(' + name + ')$', 'i'));
      if (m) { street = m[1].trim(); city = m[2].trim(); break; }
    }
  }
  // Repeated full locations or an apartment mistaken for a city need review.
  if (!street || !city || !/[A-Za-zÀ-ÿ]/.test(city) || /\d/.test(city)
      || /^(?:apt|apartment|unit|suite|ste|#)\b/i.test(city)
      || /\b(?:CA|California)\s*,?\s*\d{5}\b/i.test(street)) {
    return { ...blank, street: text };
  }
  const zip = tail[3] || '';
  return { street, city, state: 'CA', zip,
    city_state_zip: city + ', CA' + (zip ? ' ' + zip : '') };
}

const _parsedTenant = f._noticePolicy?.enabled === true ? parseConfiguredTenantProperty(_rawTenantProp) : parseTenantProperty(_rawTenantProp);
console.log('[build context] parsed tenant:', JSON.stringify(_parsedTenant));

const rawCounty = toStr(f["County"]);
const resolvedCounty = rawCounty || lookupCounty(
  _parsedTenant.street,
  _parsedTenant.city_state_zip
);
const countyValue = resolvedCounty || '\u2014';
const countyAutoResolved = !rawCounty && resolvedCounty !== '' && resolvedCounty !== '\u2014';
const countyUnknown = countyValue === '\u2014';

// ---- CALIFORNIA JUDICIAL HOLIDAYS: DERIVED FROM STATUTE, NOT LISTED ---------
// CCP 135 makes every Gov. Code 6700 holiday a judicial holiday EXCEPT Lunar New
// Year, Genocide Remembrance Day (Apr 24), Admission Day (Sep 9) and Columbus Day,
// and adds the day after Thanksgiving. CCP 135 observance: a holiday falling on a
// Saturday is observed the preceding Friday; on a Sunday, the following Monday.
// Every rule below is a fixed date or an nth-weekday, so it computes for ANY year.
// There is no list to update and nothing to go stale.
// Verified against the Judicial Council schedule (courts.ca.gov/holidays.htm).
// Replaced a hardcoded list on 2026-09-03 that was missing Lincoln Day (Feb 12)
// and Native American Day, wrongly included Columbus Day, and — in Step 2 — held
// no 2027 dates at all, so every deadline next year computed too early.
const _pad2 = (n) => String(n).padStart(2, '0');
const _iso = (d) => `${d.getUTCFullYear()}-${_pad2(d.getUTCMonth() + 1)}-${_pad2(d.getUTCDate())}`;
const _utc = (y, m, d) => new Date(Date.UTC(y, m - 1, d));
const _plus = (d, n) => new Date(d.getTime() + n * 86400000);
function _nthWeekday(year, month, weekday, n) {
  if (n === -1) { let d = _plus(_utc(year, month + 1, 1), -1);
    while (d.getUTCDay() !== weekday) d = _plus(d, -1); return d; }
  let d = _utc(year, month, 1);
  while (d.getUTCDay() !== weekday) d = _plus(d, 1);
  return _plus(d, (n - 1) * 7);
}
const _observed = (d) => d.getUTCDay() === 6 ? _plus(d, -1) : d.getUTCDay() === 0 ? _plus(d, 1) : d;
const _HOLIDAY_CACHE = {};
function _holidays(year) {
  if (_HOLIDAY_CACHE[year]) return _HOLIDAY_CACHE[year];
  const out = new Set();
  for (const [m, d] of [[1,1],[2,12],[3,31],[6,19],[7,4],[11,11],[12,25]]) out.add(_iso(_observed(_utc(year, m, d))));
  out.add(_iso(_observed(_utc(year + 1, 1, 1))));   // next New Year observed on Dec 31
  out.add(_iso(_nthWeekday(year, 1, 1, 3)));        // MLK
  out.add(_iso(_nthWeekday(year, 2, 1, 3)));        // Washington
  out.add(_iso(_nthWeekday(year, 5, 1, -1)));       // Memorial, LAST Monday
  out.add(_iso(_nthWeekday(year, 9, 1, 1)));        // Labor
  out.add(_iso(_nthWeekday(year, 9, 5, 4)));        // Native American Day
  const _tg = _nthWeekday(year, 11, 4, 4);
  out.add(_iso(_tg)); out.add(_iso(_plus(_tg, 1))); // Thanksgiving + day after
  const r = new Set([...out].filter((s) => s.startsWith(String(year) + '-')));
  _HOLIDAY_CACHE[year] = r; return r;
}
const _isCourtDay = (d) =>
  d.getUTCDay() !== 0 && d.getUTCDay() !== 6 && !_holidays(d.getUTCFullYear()).has(_iso(d));
// Day 1 is the day AFTER `start`. CCP 1161(2).
function _addCourtDays(start, n) {
  let d = start, c = 0;
  while (c < n) { d = _plus(d, 1); if (_isCourtDay(d)) c++; }
  return d;
}

// Serve By: CCP 1161(2) — three COURT days from today, counting from tomorrow.
const _serveByDate = _addCourtDays(new Date(new Date().toISOString().slice(0,10)+'T00:00:00Z'), 3);
const _serveByFormatted = new Intl.DateTimeFormat('en-US', {
  year: 'numeric', month: 'long', day: 'numeric'
}).format(_serveByDate);
const _serveByIso = _serveByDate.toISOString().slice(0, 10);

// === RENT DUE DATE INFERENCE (AM call — non-negotiable) ===
// Landlord-supplied 'Rent Due Day' is messy day-of-month text
// ('1st', 'first', '1st late on the 5th', 'first to 31st'). Parse a
// day 1-31 deterministically and build full 'Month Day, Year' strings
// for each month owed. The DUE DATE column in the rendered notice
// pulls these so the LLM never has to infer them from raw text.
const ORDINAL_WORDS = {
  first: 1, second: 2, third: 3, fourth: 4, fifth: 5,
  sixth: 6, seventh: 7, eighth: 8, ninth: 9, tenth: 10,
  eleventh: 11, twelfth: 12, thirteenth: 13, fourteenth: 14,
  fifteenth: 15, sixteenth: 16, seventeenth: 17, eighteenth: 18,
  nineteenth: 19, twentieth: 20,
  'twenty-first': 21, 'twenty-second': 22, 'twenty-third': 23,
  'twenty-fourth': 24, 'twenty-fifth': 25, 'twenty-sixth': 26,
  'twenty-seventh': 27, 'twenty-eighth': 28, 'twenty-ninth': 29,
  thirtieth: 30, 'thirty-first': 31
};

function parseRentDueDay(raw) {
  if (raw == null || raw === '') return 1;
  if (typeof raw === 'number') {
    const d = Math.floor(raw);
    return (d >= 1 && d <= 31) ? d : 1;
  }
  if (typeof raw !== 'string') raw = String(raw);
  const cleaned = raw.toLowerCase().trim();
  if (!cleaned) return 1;
  let bestPos = Infinity;
  let bestVal = 1;
  // Longer words first so 'twenty-first' wins over embedded 'first'.
  const words = Object.keys(ORDINAL_WORDS).sort((a, b) => b.length - a.length);
  for (const word of words) {
    const idx = cleaned.indexOf(word);
    if (idx !== -1 && idx < bestPos) { bestPos = idx; bestVal = ORDINAL_WORDS[word]; }
  }
  const numMatch = cleaned.match(/(\d{1,2})(st|nd|rd|th)?/);
  if (numMatch && numMatch.index < bestPos) {
    const day = parseInt(numMatch[1], 10);
    if (day >= 1 && day <= 31) { bestPos = numMatch.index; bestVal = day; }
  }
  return bestVal;
}

function buildRentDueDates(dayOfMonth, monthsOwed, today = new Date()) {
  const MONTH_NAMES = ['January','February','March','April','May','June','July','August','September','October','November','December'];
  const safeDay = Math.max(1, Math.min(31, parseInt(dayOfMonth, 10) || 1));
  const daysInMonth = (y, m) => new Date(y, m + 1, 0).getDate();
  let monthsList = [];
  if (Array.isArray(monthsOwed)) {
    monthsList = monthsOwed.slice();
  } else {
    const count = Math.max(1, parseInt(monthsOwed, 10) || 1);
    // Anchor = current month; if the due day hasn't arrived yet, roll back.
    let anchorYear = today.getFullYear();
    let anchorMonth = today.getMonth();
    if (safeDay > today.getDate()) {
      anchorMonth -= 1;
      if (anchorMonth < 0) { anchorMonth = 11; anchorYear -= 1; }
    }
    for (let i = count - 1; i >= 0; i--) {
      let y = anchorYear, m = anchorMonth - i;
      while (m < 0) { m += 12; y -= 1; }
      monthsList.push({ year: y, month: m });
    }
  }
  return monthsList
    .map(({ year, month }) => {
      const clampedDay = Math.min(safeDay, daysInMonth(year, month));
      const candidate = new Date(year, month, clampedDay);
      // Future-date guardrail: AM's rule — 'it can't be due until that
      // date actually happens'. Shift forward-dated rows back one month.
      if (candidate > today) {
        let py = year, pm = month - 1;
        if (pm < 0) { pm = 11; py -= 1; }
        const pday = Math.min(safeDay, daysInMonth(py, pm));
        return { year: py, month: pm, day: pday };
      }
      return { year, month, day: clampedDay };
    })
    .sort((a, b) => (a.year - b.year) || (a.month - b.month))
    .map(({ year, month, day }) => MONTH_NAMES[month] + ' ' + day + ', ' + year);
}

// Deterministic reader for the free-text "Itemized Rent Due" field.
// It never guesses. If the text does not yield unambiguous money, it says so
// and the caller refuses to build a notice. No LLM touches these numbers.
function parseItemizedAmounts(raw) {
  // Itemized Rent Due is an amount-bearing field. Accept plain amounts and
  // clearly labeled amount rows without treating free-text dates/counts as money.
  const rawText = String(raw == null ? '' : raw).trim();
  const amount = '(?:\\d{1,3}(?:,\\d{3})+|\\d+)(?:\\.\\d{1,2})?';
  const months = '(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)';
  const bareRow = new RegExp('^(?:(Rent(?: due)?|Balance(?: due)?|Total(?: rent)?(?: due)?|Amount due|'+months+'(?:\\s+20\\d{2})?)\\s*[:=–-]?\\s*)?('+amount+')$', 'i');
  const text = rawText.split(/([;\n])/).map(part => {
    if (new RegExp('^'+months+'\\s+20\\d{2}$','i').test(part.trim())) return part;
    const m = part.trim().match(bareRow);
    return m ? (m[1] ? m[1]+' ' : '')+'$'+m[2] : part;
  }).join('');
  if (!text) return { status: 'empty', rows: [], sum: 0, statedTotal: null, reason: 'field is blank' };

  // A reference to another document is not a missing itemization. Its balance
  // must be read before a fallback can be considered.
  if (/\b(?:see|refer(?:\s+to)?|per|attached|enclosed)\b[^\n]{0,60}\b(?:ledger|rent\s+statement|balance\s+statement)\b|\b(?:ledger|rent\s+statement)\b[^\n]{0,40}\b(?:attached|enclosed)\b/i.test(text)) {
    return { status: 'document-required', rows: [], sum: 0, statedTotal: null, reason: text };
  }

  // A placeholder means "nothing itemised here", which is a genuine single-month
  // case. That is different from a human writing "I'll send it later", which
  // means the amount is not yet known and must never be guessed at.
  const stripped = text.replace(/[\s\-_.,:;"'\u201c\u201d\u2018\u2019()]/g, '').toLowerCase();
  if (stripped === '' || ['na','n/a','none','nil','null','0','tbd'].includes(stripped)) {
    return { status: 'none', rows: [], sum: 0, statedTotal: null, reason: 'no itemisation provided' };
  }

  // Blank out anything date-shaped first so 3/18/26 is never read as money.
  const masked = text
    .replace(/\b\d{1,4}[\/\-]\d{1,2}[\/\-]\d{2,4}\b/g, m => ' '.repeat(m.length))
    .replace(/\b\d{4}-\d{2}-\d{2}\b/g, m => ' '.repeat(m.length));

  // A number counts as money only if it carries a $ or explicit cents.
  // That keeps bare years ("2026") and counts ("5 months") out of the maths.
  const MONEY = /\$\s*\d[\d,]*(?:\.\d{1,2})?|\b\d[\d,]*\.\d{2}\b/g;
  const tokens = [];
  let m;
  while ((m = MONEY.exec(masked)) !== null) {
    const value = parseFloat(m[0].replace(/[$,\s]/g, ''));
    if (!isFinite(value)) continue;
    tokens.push({ value, start: m.index, text: m[0].trim() });
  }
  if (tokens.length === 0) {
    return { status: 'unparseable', rows: [], sum: 0, statedTotal: null,
             reason: 'no dollar amounts found in: ' + JSON.stringify(text.slice(0, 90)) };
  }

  // A token is the stated total when a total-ish word sits just before it.
  const TOTAL_WORD = /(total|balance|amount\s+due|due\s+in\s+full)[^$]{0,25}$/i;
  let totalIdx = -1;
  tokens.forEach((tok, i) => {
    const lead = masked.slice(Math.max(0, tok.start - 40), tok.start);
    if (TOTAL_WORD.test(lead)) totalIdx = i;
  });

  const statedTotal = totalIdx >= 0 ? tokens[totalIdx].value : null;
  const rowTokens = tokens.filter((_, i) => i !== totalIdx);

  // A row's label is only the text between the previous amount and this one.
  // Slicing a fixed window swallowed the row before it.
  const MONTHS = ['january','february','march','april','may','june','july','august',
                  'september','october','november','december'];
  const labelFor = (tok, prevEnd) => text.slice(prevEnd, tok.start)
    .replace(/[\s:,\/\-&.]+$/, '').replace(/^[\s:,\/\-&.]+/, '').trim();
  const monthOf = lbl => {
    const m = String(lbl).toLowerCase().match(/\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\b/);
    if (!m) return null;
    const idx = MONTHS.findIndex(x => x.startsWith(m[1]));
    if (idx < 0) return null;
    const y = String(lbl).match(/\b(20\d{2})\b/);
    return { month: idx, year: y ? parseInt(y[1], 10) : null };
  };

  const decorate = list => {
    let prevEnd = 0;
    return list.map(t => {
      const lbl = labelFor(t, prevEnd);
      prevEnd = t.start + t.text.length;
      const mo = monthOf(lbl);
      return { label: lbl, amount: t.value, isMonth: !!mo, month: mo };
    });
  };

  if (rowTokens.length === 0) {
    const only = decorate([tokens[totalIdx]]);
    return { status: 'ok', rows: only, allMonths: only.every(r => r.isMonth),
             sum: statedTotal, statedTotal, reason: 'single stated amount' };
  }

  const rows = decorate(rowTokens);
  const sum = Math.round(rows.reduce((a, r) => a + r.amount, 0) * 100) / 100;

  if (statedTotal !== null && Math.abs(sum - statedTotal) > 0.01) {
    return { status: 'inconsistent', rows, allMonths: rows.every(r => r.isMonth), sum, statedTotal,
             reason: `rows add to $${sum.toFixed(2)} but the field states $${statedTotal.toFixed(2)}` };
  }
  return { status: 'ok', rows, allMonths: rows.every(r => r.isMonth), sum, statedTotal, reason: statedTotal !== null ? 'rows agree with stated total' : 'sum of rows' };
}


// Itemized unpaid rent is authoritative. Monthly Rent is used only for explicit unpaid periods.
function reconcileTotal(itemized, rentField) {
  const money = v => {
    if (v === null || v === undefined || v === '') return null;
    const n = parseFloat(String(v).replace(/[$,\s]/g, ''));
    return isFinite(n) && n > 0 ? n : null;
  };
  const rent = money(rentField);
  const near = (a, b) => a != null && b != null && Math.abs(a - b) < 0.01;
  const figures = [];
  if (rent != null) figures.push({ source: 'Monthly rent rate (not the unpaid balance)', value: rent });
  if (itemized.statedTotal != null) figures.push({ source: 'total typed in the itemised text', value: itemized.statedTotal });
  if (itemized.rows && itemized.rows.length) {
    // List the individual rows too - the email has to name every number it
    // found, not just the arithmetic it did with them.
    itemized.rows.forEach(r => figures.push({
      source: 'itemised row' + (r.label ? ' (' + r.label + ')' : ''), value: r.amount }));
    if (itemized.rows.length > 1) figures.push({ source: 'itemised rows added up', value: itemized.sum });
  }
  const money$ = n => '$' + Number(n).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const ask = (lead) => lead + ' Please review the dated charges, correct the rent figures in Airtable, and have the notice generated again.';

  if (itemized.status === 'document-required') {
    return { status: 'ambiguous', total: null, source: null, figures,
      question: 'The itemized rent refers to a ledger or statement. Its dated charges and balance must be reviewed before a notice can be generated.' };
  }

  // The rows already contradict a total typed beside them.
  if (itemized.status === 'inconsistent') {
    return { status: 'ambiguous', total: null, source: null, figures,
      question: ask('The itemised rows add to ' + money$(itemized.sum)
        + ' but the same field states ' + money$(itemized.statedTotal) + '. Which is right?') };
  }

  if (itemized.status === 'ok' && itemized.rows.length > 0) {
    return { status: 'ok', total: itemized.sum,
      source: itemized.statedTotal != null ? 'row-sum-confirmed' : 'row-sum', figures };
  }

  return { status: 'ambiguous', total: null, source: null, figures,
    question: ask('The unpaid rent could not be determined from Itemized Rent Due ('
      + String(itemized.reason || 'missing information') + '). Monthly Rent alone does not establish the balance owed.') };
}

// Monthly rent is a rate, never a rival arrears total.
function parseItemized(raw, monthlyRent, today = new Date()) {
  let text = String(raw == null ? '' : raw).trim();
  const excludedCharges = [];
  text = text.replace(/\b(late\s+fees?|late\s+charges?)\s*(?:of|:|=)?\s*(\$?\s*\d[\d,]*(?:\.\d{1,2})?)/gi, (whole,label,amount) => {
    excludedCharges.push({label,amount:Number(amount.replace(/[$,\s]/g,''))});
    return ' '.repeat(whole.length);
  });
  if (excludedCharges.length) text = text.replace(/(?:\bplus|\band|[+])\s*$/i, '').trim();
  const stop = reason => ({status:'ambiguous',rows:[],sum:0,statedTotal:null,reason,excludedCharges});
  // Do not silently add payments/credits or negative amounts as charges.
  if (/\b(?:credit|payment|paid|deduct|less|minus|deposit|utilities|damages)\b|[-−]\s*\$?\s*\d|\(\s*\$\s*\d/i.test(text))
    return stop('Charges or credits need an explicit remaining-rent balance.');
  const months = ['january','february','march','april','may','june','july','august','september','october','november','december'];
  const monthPattern = '('+months.join('|')+')';
  const since = text.match(new RegExp('^(?:(?:unpaid\\s+)?rent\\s+)?(?:owed\\s+)?(?:since|from)\\s+'+monthPattern+'(?:\\s+(20\\d{2}))?\\s*[.!]?$', 'i'));
  const range = text.match(new RegExp('^(?:(?:unpaid\\s+)?rent\\s+)?(?:owed\\s+)?(?:from\\s+)?'+monthPattern+'(?:\\s+(20\\d{2}))?\\s*(?:through|thru|to|[-–])\\s*'+monthPattern+'(?:\\s+(20\\d{2}))?\\s*[.!]?$', 'i'));
  if (since || range) {
    const rate = Number(String(monthlyRent ?? '').replace(/[$,\s]/g,''));
    if (!(rate>0) || !Number.isFinite(rate)) return stop('Monthly rent is required for the specified unpaid months.');
    const m=since||range, startMonth=months.indexOf(m[1].toLowerCase());
    const startYear=m[2]?Number(m[2]):today.getFullYear();
    const endMonth=since?today.getMonth():months.indexOf(m[3].toLowerCase());
    const endYear=since?today.getFullYear():(m[4]?Number(m[4]):startYear);
    const first=startYear*12+startMonth,last=endYear*12+endMonth;
    if(last<first || last-first>119 || last>today.getFullYear()*12+today.getMonth())return stop('Unpaid month range needs clarification.');
    const rows=[];
    for(let n=first;n<=last;n++){const year=Math.floor(n/12),month=n%12;rows.push({label:months[month]+' '+year,amount:rate,isMonth:true,month:{month,year}});}
    return {status:'ok',rows,allMonths:true,sum:Math.round(rows.length*rate*100)/100,statedTotal:null,reason:'explicit unpaid months multiplied by monthly rent',excludedCharges};
  }
  // Ranges with extra wording/amounts need interpretation; never count a rate as a total.
  if (/\b(?:since|per\s+month|monthly|months?\s+(?:at|@)|through|thru)\b/i.test(text))return stop('Please provide dated unpaid rent amounts or a clear unpaid month range.');
  const result=parseItemizedAmounts(text);
  return {...result,excludedCharges};
}

module.exports = { parseItemized, reconcileTotal };

const _rentDueDay = parseRentDueDay(f["Rent Due Day"]);
// Itemised rent is a free-text field typed by humans. Read it deterministically.
// An LLM must never do arithmetic on money that goes into a legal notice.
const _itemized = parseItemized(f["Itemized Rent Due"], f["Rent"]);
// "Rent" (fldmNhzUpFBktNhbO) is the only currency field on the record. Reconcile
// it against the itemised rows and against any total typed inside that text.
const _money = reconcileTotal(_itemized, f["Rent"]);
const _rentAmountDefault = (f["Amount Due"] && String(f["Amount Due"]).trim())
  || (f["Rent"] && String(f["Rent"]).trim())
  || '';
// Show the itemised rows only when the total actually came from them.
const _useItemized = _money.status === 'ok'
  && (_money.source === 'row-sum' || _money.source === 'row-sum-confirmed')
  && _itemized.rows.length > 0;
// Each row gets the date IT names. Generating N consecutive months instead
// would invent a rent period for any row that is a late fee rather than rent.
const _MONTH_NAMES = ['January','February','March','April','May','June','July',
                      'August','September','October','November','December'];
function _rowDateLabel(row, dueDay, today) {
  if (!row.isMonth || !row.month) return row.label || 'Amount due';
  const day = Math.max(1, Math.min(31, parseInt(dueDay, 10) || 1));
  let year = row.month.year;
  if (!year) {
    year = today.getFullYear();
    if (row.month.month > today.getMonth()) year -= 1;
  }
  const dim = new Date(year, row.month.month + 1, 0).getDate();
  return _MONTH_NAMES[row.month.month] + ' ' + Math.min(day, dim) + ', ' + year;
}
const _today = new Date();
const _rentDueDatesDefault = _useItemized
  ? _itemized.rows.map(r => _rowDateLabel(r, _rentDueDay, _today))
  : buildRentDueDates(_rentDueDay, 1);
const _amountsDueDefault = _useItemized
  ? _itemized.rows.map(r => String(r.amount))
  : _rentDueDatesDefault.map(() => (_money.source === 'rent-field'
      ? String(_money.total) : _rentAmountDefault));
const _totalComputed = (_money.status === 'ok' && _money.total > 0)
  ? _money.total
  : (parseFloat(String(_rentAmountDefault).replace(/[$,]/g, '')) || 0);

const fullContext = {
  serve_by_date: _serveByFormatted,
  serve_by_iso: _serveByIso,
  airtable: {
    record_id: toStr(f.id),
    created_time: toStr(f.createdTime)
  },
  case: {
    case_id: toStr(f["Case ID"]),
    case_name: toStr(f["Case Name"]),
    notice_type: toStr(f["Notice Type"]),
    product: toStr(f["Product"]),
    status: toStr(f["Status"]),
    created_date: toStr(f["Created"]),
    last_modified: toStr(f["Last Modified"]),
    notes: toStr(f["Notes"])
  },
  // April 10 / May 11 intake contract: the named intake person signs;
  // property owner / payment recipient is a separate field.
  signer: {
    full_name: [toStr(f["First Name"]), toStr(f["Last Name"])].filter(Boolean).join(' ').trim(),
    title: toStr(f["Signer Title"]),
    source: 'designated-signatory-intake'
  },
  landlord: {
    full_name: toStr(f["Property Owner's Name"]),
    first_name: toStr(f["First Name"]),
    last_name: toStr(f["Last Name"]),
    phone: toStr(f["Landlords Preferred Phone"] || f["Phone Number"]),
    email: toStr(f["Landlord's Preferred Email"] || f["Email"]),
    address: toStr(f["Landlord Address"])
  },
  tenant: {
    names: toStr(f["Tenant Name(s)"]),
    contact_information: toStr(f["Tenant's Contact Information"]),
    language_used: toStr(f["Tenant Language"])
  },
  property: {
    address: _parsedTenant.street,
    city_state_zip: _parsedTenant.city_state_zip,
    city: _parsedTenant.city,
    zip: _parsedTenant.zip,
    county: countyValue,
    county_auto_resolved: countyAutoResolved,
    county_unknown: countyUnknown,
    state: _parsedTenant.state,
    section_8: toYesNo(f["Section 8"]),
    number_of_units: f["Number of Units"] ?? null,
    year_built: f["Year Built"] ?? null,
    tpo_override: f["TPO Override"] === true
  },
  financials: {
    current_monthly_rent: f["Rent"] ?? "",
    amount_due: f["Amount Due"] || f["Rent"] || "",
    total_amount_due: _totalComputed > 0 ? String(_totalComputed) : (f["Rent"] || ""),
    total_amount_due_computed: _totalComputed,
    total_source: _money.source || 'airtable-rent',
    itemized: _itemized,
    money: _money,
    rent_due_date: f["Due Date"] || f["Rent Due Date"] || "",
    rent_due_day_raw: toStr(f["Rent Due Day"]),
    rent_due_day: _rentDueDay,
    itemized_rent_due_raw: toStr(f["Itemized Rent Due"])
  },
  attachments: {
    leases: f["Lease"] ?? [],
    prior_notices: f["Prior Notices"] ?? [],
    other: f["Attachments"] ?? []
  },
  rent_due: {
    day: _rentDueDay,
    raw_day_string: toStr(f["Rent Due Day"]),
    itemized_raw: toStr(f["Itemized Rent Due"]),
    dates: _rentDueDatesDefault,
    amounts: _amountsDueDefault
  },
  RENT_DUE_DATES: _rentDueDatesDefault,
  AMOUNTS_DUE: _amountsDueDefault
};

return fullContext;
}
// Exact rent enumeration only. Source dates suggest a question, never establish a debt year.
function resolveRentYear(ctx, record={}, mailResponses=[], now=new Date()) {
 const c=JSON.parse(JSON.stringify(ctx)), f=record.fields||record;
 const raw=String(c.financials.itemized_rent_due_raw||'').trim();
 const mn=['January','February','March','April','May','June','July','August','September','October','November','December'];
 const pat='(Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:t(?:ember)?)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)';
 const hits=[...raw.matchAll(new RegExp('\\b'+pat+'\\b','gi'))];
 if(!hits.length)return c;
 // Preserve the wider parser's holds for ranges, credits and other adjustments.
 if(/\b(?:paid|payment|credit|deposit|late|fee|fees|partial|prorat\w*|since|through|thru|until|less|minus|ignore|instructions?)\b|[−]|\$\s*-\s*\d/i.test(raw))return c;
 const amount='(?:\\d{1,3}(?:,\\d{3})+|\\d+)(?:\\.\\d{1,2})?';
 const re=new RegExp('\\b'+pat+'\\b(?:\\s+(20\\d{2}))?\\s*(?:[:=]\\s*)?(?:\\(\\s*\\$?('+amount+')\\s*\\)|\\$('+amount+')|('+amount+'))?','gi');
 const matches=[...raw.matchAll(re)];
 let rest=raw.replace(re,'').replace(new RegExp('\\btotal(?:\\s+rent)?\\s*[:=]?\\s*\\$?('+amount+')','gi'),'').replace(/\b(?:and|of|rents?|20\d{2})\b/gi,'').replace(/[\s,;&.]/g,'');
 if(rest||matches.length!==hits.length)return c;
 const rows=matches.map(m=>({month:mn.findIndex(n=>n.slice(0,3).toLowerCase()===m[1].slice(0,3).toLowerCase())+1,year:m[2]?Number(m[2]):null,amount:m[3]||m[4]||m[5]||null}));
 if(new Set(rows.map(r=>r.month)).size!==rows.length)return c;
 const hold=(question,reason)=>{c.financials.money={...c.financials.money,status:'ambiguous',total:null,source:null,question};c.financials.year_resolution={status:'needs_confirmation',reason,evidence:[]};return c;};
 const years=[...new Set((raw.match(/\b20\d{2}\b/g)||[]).map(Number))];
 const evidence=[];
 if(years.length===1)for(const r of rows)if(!r.year)r.year=years[0];
 if(rows.some(r=>!r.year)){
  const texts=[{source:'Airtable Notes',text:String(f.Notes||'')}];
  for(const key of ['Rent Year','Unpaid Rent Year'])if(/^20\d{2}$/.test(String(f[key]||'')))texts.push({source:'Airtable '+key,text:'Rent year: '+f[key]});
  const order=String(f['Order #']||'');
  for(const response of mailResponses){
   if(response?.error||response?.['@odata.nextLink'])continue;
   for(const m of response?.value||[]){
    const text=String(m.body?.content||'').replace(/<[^>]*>/g,' ').replace(/&nbsp;/g,' ').replace(/\s+/g,' ');
    // No sender-name or subject matching: require the exact purchase reference.
    const ids=[...text.matchAll(/(?:Purchase\s*ID|Order\s*(?:#|ID|number))\s*[:#]?\s*(\d{6,})/gi)].map(x=>x[1]);
    if(!order||!ids.includes(order)||ids.some(id=>id!==order))continue;
    texts.push({source:'Order email '+String(m.internetMessageId||m.id||''),text});
   }
  }
  for(const {source,text} of texts){
   for(const m of text.matchAll(/\b(?:unpaid\s+)?rent\s+year\s*[:=]\s*(20\d{2})\b/gi))evidence.push({source,year:Number(m[1]),months:rows.map(r=>r.month)});
   // Explicit rent sentences only. A sent date, trust date or lease date is not debt evidence.
   for(const sentence of text.split(/[\n.!?]/)){
    if(!/\b(?:rent|rents|owed|unpaid)\b/i.test(sentence))continue;
    const ys=[...new Set((sentence.match(/\b20\d{2}\b/g)||[]).map(Number))];
    const ms=[...sentence.matchAll(new RegExp('\\b'+pat+'\\b','gi'))].map(m=>mn.findIndex(n=>n.slice(0,3).toLowerCase()===m[1].slice(0,3).toLowerCase())+1);
    if(ys.length===1&&ms.length&&rows.every(r=>ms.includes(r.month)))evidence.push({source,year:ys[0],months:ms});
   }
  }
  const candidates=[...new Set(evidence.map(e=>e.year))];
  if(candidates.length===1){for(const r of rows)if(!r.year)r.year=candidates[0];}
  else {
   const anchor=String(f['Order Date']||'').match(/^(20\d{2})-(\d{2})/);
   const labels=rows.map(r=>mn[r.month-1]).join(' and ');
   const crossYear=rows.some((r,i)=>i&&r.month<rows[i-1].month);
   const question=candidates.length>1?'I found different rent years in the case or matching email. Which year applies to '+labels+'?':anchor&&!crossYear?'Are '+labels+' both in '+anchor[1]+'?':'Which year applies to '+labels+'?';
   return hold(question,'missing_or_conflicting_rent_year');
  }
 }
 if(rows.some(r=>!r.year))return hold('Please confirm the year for each unpaid month.','missing_year');
 const amounts=rows.map(r=>r.amount===null?Number(c.financials.current_monthly_rent):Number(r.amount.replace(/,/g,'')));
 if(amounts.some(a=>!Number.isFinite(a)||a<=0||Math.abs(a*100-Math.round(a*100))>1e-5))return hold('What rent amount is owed for each listed month?','invalid_amount');
 const sum=amounts.reduce((a,b)=>a+Math.round(b*100),0)/100;
 const totals=[...raw.matchAll(new RegExp('\\btotal(?:\\s+rent)?\\s*[:=]?\\s*\\$?('+amount+')','gi'))];
 if(totals.length>1||totals.some(t=>Math.round(Number(t[1].replace(/,/g,''))*100)!==Math.round(sum*100)))return hold('The listed rent amounts add up to $'+sum.toFixed(2)+', but the stated total differs. Which amount should I use?','total_conflict');
 if(rows.some(r=>r.year>now.getFullYear()||(r.year===now.getFullYear()&&r.month>now.getMonth()+1)))return hold('One listed rent month is in the future. Please confirm the unpaid months.','future_month');
 const day=Math.max(1,Math.min(31,Number(c.financials.rent_due_day)||1));
 const out=rows.map((r,i)=>({label:mn[r.month-1]+' '+Math.min(day,new Date(r.year,r.month,0).getDate())+', '+r.year,amount:amounts[i],isMonth:true,month:{month:r.month-1,year:r.year}}));
 c.financials.itemized={status:'ok',rows:out,sum,statedTotal:totals.length?sum:null,excludedCharges:[]};
 c.financials.money={status:'ok',total:sum,source:'case-grounded-rent',figures:out.map(r=>({source:r.label,value:r.amount}))};
 c.financials.total_amount_due=String(sum);c.financials.total_amount_due_computed=sum;c.financials.total_source='case-grounded-rent';
 c.RENT_DUE_DATES=out.map(r=>r.label);c.AMOUNTS_DUE=amounts.map(String);c.rent_due={...c.rent_due,dates:c.RENT_DUE_DATES,amounts:c.AMOUNTS_DUE};
 c.financials.year_resolution={status:'resolved',evidence:years.length?[{source:'Itemized Rent Due',years}]:evidence};
 return c;
}


function noticePayload(context) {
// !! service_date IS A DEAD KEY. The renderer reads serve_date, never
// !! service_date (grep src/ -- no matches). Today this value is ignored and the
// !! notice prints a blank that the process server fills in by hand.
// !! DO NOT wire it through to the renderer: it is a SERVE-BY target computed from
// !! today, not the date the notice was actually served. Printing it would put a
// !! future date on the notice as the date of service.
// Build the JSON payload for the Railway /generate-notice endpoint
const ctx = context;
const editFields = {};
const draft = {};

// Currency formatter
function formatCurrency(val) {
  const num = parseFloat(String(val).replace(/[$,]/g, ''));
  if (isNaN(num)) return '$0.00';
  return '$' + num.toFixed(2).replace(/\d(?=(\d{3})+\.)/g, '$&,');
}

// Phone formatter
function formatPhone(raw) {
  if (!raw) return '';
  const digits = String(raw).replace(/\D/g, '');
  if (digits.length === 10) {
    return `(${digits.slice(0,3)}) ${digits.slice(3,6)}-${digits.slice(6)}`;
  }
  if (digits.length === 11 && digits[0] === '1') {
    return `(${digits.slice(1,4)}) ${digits.slice(4,7)}-${digits.slice(7)}`;
  }
  return raw;
}

// Helper: use Airtable value first, GPT fallback
const airtableFirst = (airtableVal, gptVal) => {
  const a = airtableVal ? String(airtableVal).trim() : '';
  if (a && a !== 'NOT_FOUND' && a !== 'not_found') return a;
  const g = gptVal ? String(gptVal).trim() : '';
  if (g && g !== 'NOT_FOUND' && g !== 'not_found') return g;
  return '';
};

// ========================================
// ADDRESS PARSER — fix smashed-together addresses from Zapier
// ========================================
function parseAddress(raw) {
  if (!raw) return '';
  let addr = String(raw).trim();
  // Fix smashed street suffix + city: "Cedar StreetAptos" -> "Cedar Street\nAptos"
  addr = addr.replace(/(Street|St|Ave|Avenue|Dr|Drive|Rd|Road|Ln|Lane|Blvd|Boulevard|Way|Ct|Court|Pl|Place|Cir|Circle)([A-Z])/g, '$1\n$2');
  // If it already has commas or newlines, it's properly formatted
  if (addr.includes(',') || addr.includes('\n')) return addr;
  // Detect state abbreviation pattern to split smashed addresses like "87 Main StSanta ClaraCA12345"
  const stateAbbrs = ['AL','AK','AZ','AR','CA','CO','CT','DE','FL','GA','HI','ID','IL','IN','IA','KS','KY','LA','ME','MD','MA','MI','MN','MS','MO','MT','NE','NV','NH','NJ','NM','NY','NC','ND','OH','OK','OR','PA','RI','SC','SD','TN','TX','UT','VT','VA','WA','WV','WI','WY'];
  for (const st of stateAbbrs) {
    const re = new RegExp('([A-Za-z ]+?)\\s*(' + st + ')\\s*(\\d{5}(?:-\\d{4})?)', 'i');
    const m = addr.match(re);
    if (m) {
      const idx = addr.indexOf(m[0]);
      const street = addr.substring(0, idx).trim();
      const city = m[1].trim();
      const state = m[2].toUpperCase();
      const zip = m[3];
      if (street && street.toLowerCase() !== city.toLowerCase()) {
        return street + '\n' + city + ', ' + state + ' ' + zip;
      }
      return city + ', ' + state + ' ' + zip;
    }
  }
  return addr;
}

// Determine notice_type — AIRTABLE FIRST
const rawNoticeType = airtableFirst(
  ctx.case?.notice_type || editFields.case?.notice_type,
  draft.NOTICE_TYPE
).toLowerCase();
let noticeType = 'residential';
if (rawNoticeType.includes('commercial')) {
  noticeType = 'commercial';
}

// Day count
let dayCount = parseInt(draft.DAY_COUNT) || 3;
if (isNaN(dayCount) || dayCount < 1) {
  const dayMatch = (ctx.case?.notice_type || editFields.case?.notice_type || '').match(/(\d+)/);
  dayCount = dayMatch ? parseInt(dayMatch[1]) : 3;
}

// Normalize notice type to header — handles all Airtable name variants
function resolveNoticeHeader(noticeType) {
  const t = (noticeType || '').toLowerCase();
  if (t.includes('pay') && t.includes('quit')) return 'THREE (3) DAY NOTICE TO PAY RENT OR QUIT';
  if (t.includes('perform') && t.includes('quit')) return 'THREE (3) DAY NOTICE TO PERFORM COVENANTS OR QUIT';
  if (t.includes('quit')) return 'THREE (3) DAY NOTICE TO QUIT';
  if (t.includes('tpo') || t.includes('tenant protection ordinance')) return 'TENANT PROTECTION ORDINANCE WARNING';
  if (t.includes('tpa') || t.includes('tenant protection act')) return 'TENANT PROTECTION ACT WARNING';
  if (t.includes('warning')) return 'TENANT PROTECTION ORDINANCE WARNING';
  if (t.includes('30-day') || t.includes('30 day')) return 'THIRTY (30) DAY NOTICE OF TERMINATION OF TENANCY';
  if (t.includes('60-day') || t.includes('60 day')) return 'SIXTY (60) DAY NOTICE OF TERMINATION OF TENANCY';
  if (t.includes('3-day') || t.includes('3 day')) return 'THREE (3) DAY NOTICE TO PAY RENT OR QUIT';
  return noticeType.toUpperCase();
}
const fullNoticeType = ctx.case?.notice_type || editFields.case?.notice_type || '';
const noticeHeader = draft.NOTICE_HEADER || resolveNoticeHeader(fullNoticeType);

// --- AIRTABLE-FIRST fields ---
let tenantNames = airtableFirst(ctx.tenant?.names, draft.TENANT_NAMES);
// ALWAYS append "and all other occupants" if not already present
if (tenantNames && !tenantNames.toLowerCase().includes('all other occupant') && !tenantNames.toLowerCase().includes('all others in possession')) {
  tenantNames += ', and all other occupants';
}
const propertyAddress = airtableFirst(ctx.property?.address, draft.PROPERTY_ADDRESS_STREET || draft.PROPERTY_ADDRESS);
const propertyCityStateZip = airtableFirst(ctx.property?.city_state_zip, draft.PROPERTY_ADDRESS_CITY);
const caseName = airtableFirst(ctx.case?.case_name || editFields.case?.case_name, draft.CASE_NAME);
const landlordName = airtableFirst(ctx.landlord?.full_name, draft.LANDLORD_NAME);
const landlordPhone = formatPhone(airtableFirst(ctx.landlord?.phone, draft.LANDLORD_PHONE));

// County: derive from city in Airtable address, fallback to Airtable county field, then GPT
const SERVICE_AREA = {
  "Santa Clara": ["Alum Rock","Campbell","Cupertino","Gilroy","Los Altos","Los Altos Hills","Los Gatos","Milpitas","Monte Sereno","Morgan Hill","Mountain View","Palo Alto","San Jose","San Jos\u00e9","Santa Clara","Saratoga","Sunnyvale"],
  "San Mateo": ["Atherton","Belmont","Brisbane","Burlingame","Colma","Daly City","East Palo Alto","Foster City","Half Moon Bay","Hillsborough","Menlo Park","Millbrae","Pacifica","Portola Valley","Redwood City","San Bruno","San Carlos","San Mateo","South San Francisco","Woodside"],
  "Santa Cruz": ["Aptos","Capitola","Santa Cruz","Scotts Valley","Watsonville"],
  "San Benito": ["Hollister","San Juan Bautista"],
  "Alameda": ["Alameda","Albany","Berkeley","Dublin","Emeryville","Fremont","Hayward","Livermore","Newark","Oakland","Pleasanton","San Leandro","Union City"],
  "Contra Costa": ["Antioch","Concord","Danville","Lafayette","Martinez","Orinda","Pittsburg","Richmond","San Ramon","Walnut Creek"],
  "San Francisco": ["San Francisco"]
};

function deriveCounty(cityStateZip) {
  if (!cityStateZip) return '';
  const cityMatch = String(cityStateZip).match(/^([^,]+)/);
  if (!cityMatch) return '';
  const city = cityMatch[1].trim().toLowerCase();
  for (const [county, cities] of Object.entries(SERVICE_AREA)) {
    for (const c of cities) {
      if (city === c.toLowerCase()) return county;
    }
  }
  return '';
}

const derivedCounty = deriveCounty(propertyCityStateZip);
const county = derivedCounty || airtableFirst(ctx.property?.county, draft.COUNTY);

// --- Airtable rent fallback (Rent field from Kajabi via Zapier) ---
const airtableRent = ctx.financials?.amount_due || ctx.financials?.current_monthly_rent || '';
const airtableTotal = ctx.financials?.total_amount_due || ctx.financials?.current_monthly_rent || '';
const airtableDueDate = ctx.financials?.rent_due_date || '';

// --- GPT-extracted fields (amounts, dates) with Airtable fallback ---
// Both line items and total come from the same reconciled rent context.
const reconciled = ctx.financials?.money;
const dates = ctx.RENT_DUE_DATES;
const amounts = ctx.AMOUNTS_DUE;
if(reconciled?.status !== 'ok' || !Array.isArray(dates) || !Array.isArray(amounts)
   || dates.length !== amounts.length || !dates.length) throw Error('Verified unpaid rent rows required');
const amounts_due = dates.map((date,i)=>({due_date:String(date),amount:formatCurrency(amounts[i])}));
const rowSum=amounts.reduce((sum,a)=>sum+Number(String(a).replace(/[$,]/g,'')),0);
if(!Number.isFinite(rowSum)||Math.abs(rowSum-reconciled.total)>0.01)throw Error('Rent rows do not equal unpaid balance');
const totalFormatted=formatCurrency(reconciled.total);

// Landlord address — parse smashed addresses from Zapier
const rawLandlordAddr = airtableFirst(ctx.landlord?.address, draft.LANDLORD_ADDRESS);
const formattedLandlordAddr = parseAddress(rawLandlordAddr);

const cleanField = (val) => {
  if (!val) return '';
  const s = String(val);
  return (s === 'NOT_FOUND' || s === 'not_found') ? '' : s;
};

// Format notice_date
const months = ['January','February','March','April','May','June','July','August','September','October','November','December'];
const today = new Date();
const noticeDate = months[today.getMonth()] + ' ' + today.getDate() + ', ' + today.getFullYear();

// Boolean flags
const isSection8 = draft.IS_SECTION_8 === true || String(draft.IS_SECTION_8).toLowerCase() === 'true'
  || String(ctx.property?.section_8).toLowerCase() === 'yes';

// ========================================
// ATTACHMENT RULES - Deterministic evaluation
// ========================================
const attachments_required = [];

function extractCityFromPropertyAddress(address) {
  if (!address) return '';
  const parts = String(address).split(',');
  if (parts.length < 2) return '';
  let cityPart = parts.slice(1).join(',').trim();
  cityPart = cityPart.replace(/\s*,?\s*[A-Z]{2}\s+\d{5}(-\d{4})?\s*$/, '').trim();
  return cityPart.toLowerCase();
}
const cityRaw = String(ctx.property?.city || '').trim().toLowerCase()
  || String(propertyCityStateZip || '').split(',')[0].replace(/\s+CA(?:\s+\d{5}(?:-\d{4})?)?$/i, '').trim().toLowerCase()
  || extractCityFromPropertyAddress(propertyAddress);

const units = ctx.property?.number_of_units;
const yearBuilt = ctx.property?.year_built;
const tpoOverride = ctx.property?.tpo_override === true;

// RULE 1: 30-Day + residential -> HUD attachments
if (rawNoticeType.includes('30-day') && noticeType === 'residential') {
  attachments_required.push('Attachment_1_HUD-5380.pdf');
  attachments_required.push('Attachment_2_HUD-5382.pdf');
}

// RULE 2: Nonpayment + San Jose + 3+ units + (year_built <= 1979 OR TPO Override)
const isNonpayment = rawNoticeType.includes('pay');
if (isNonpayment) {
  const isSanJose = (cityRaw === 'san jose' || cityRaw === 'san jos\u00e9');
  if (isSanJose && units >= 3 && ((yearBuilt != null && yearBuilt <= 1979) || tpoOverride)) {
    attachments_required.push('TPO_Required_Attachment.pdf');
  }

  // RULE 3: Nonpayment + Mountain View + 3+ units + year_built < 2016
  const isMtnView = (cityRaw === 'mountain view');
  if (isMtnView && units >= 3 && yearBuilt != null && yearBuilt < 2016) {
    attachments_required.push('Mountain_View_Attachment.pdf');
  }
}

const payload = {
  notice_type: noticeType,
  notice_header: noticeHeader,
  day_count: dayCount,
  tenant_names: tenantNames,
  property_address_street: propertyAddress.trim(),
  property_address_city: propertyCityStateZip || propertyAddress.split(',').slice(1).join(',').trim(),
  property_address: propertyAddress + ', ' + propertyCityStateZip,
  county: county,
  amounts_due: amounts_due,
  total_amount_due: totalFormatted,
  service_date: ctx.serve_by_date || '________',
  payment_address: formattedLandlordAddr,
  landlord_name: landlordName,
  signer_name: ctx.signer?.full_name || '',
  signer_title: ctx.signer?.title || '', 
  landlord_email: airtableFirst(ctx.landlord?.email || editFields.landlord?.email, draft.LANDLORD_EMAIL),
  landlord_phone: landlordPhone,
  landlord_address: formattedLandlordAddr,
  notice_date: noticeDate,
  case_name: caseName.replace(/\bv\./g, 'v'),
  is_section_8: isSection8,
  attachments_required: attachments_required
};

// A total is not a rent period. Ask the reviewer rather than inventing dates.
payload._notice_input_issues = amounts_due.some(row => !String(row.due_date || '').trim() || /^(?:amount|balance|rent|total)(?:\s+rent)?(?:\s+due)?$/i.test(String(row.due_date).trim()))
  ? ['Please provide the unpaid rent periods/date ranges for the listed amounts.'] : [];

delete payload.service_date; return payload;
}
module.exports={caseContext,resolveRentYear,noticePayload};
