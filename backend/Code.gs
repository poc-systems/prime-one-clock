/**
 * Prime One Clock — backend for the Prime One time clock app.
 *
 * Two web-app deployments of this one project:
 *  1. Kiosk API (the iPads and Jacksonville phones). Execute as: Me. Who has access: Anyone.
 *     Only doPost is used. It holds NO office functions.
 *  2. Office page (Juliana and Andre, on a laptop or phone). Execute as: Me.
 *     Who has access: Anyone within primeoneinc.com. Google sign-in, and every action checks the
 *     signed-in email against OFFICE_EMAILS, not just when the page opens.
 *
 * Rules (decided by Andre, Oct 29, 2026)
 *  - A device only works after it is set up with a one-time code from the Office page.
 *  - Every punch needs the person's own 6-digit PIN and takes a small photo.
 *  - PINs are checked only here, stored hashed with a secret pepper, in their own restricted sheet
 *    (SYSTEM · Clock PINs). The iPad never receives anything about a PIN.
 *  - First sign-in: the person enters the shared setup code, then must choose their own PIN.
 *  - Wrong PINs: after 5 in a row, 30 seconds between tries; at 10 the PIN locks and Andre and
 *    Juliana get an email. Only the office unlocks or resets.
 *  - Punch time is the server's clock. A punch made offline keeps the time it was made, measured
 *    from the last server sync with a stopwatch the iPad's clock setting can't change. If the app
 *    restarted while offline, the iPad clock is used and the punch is marked for review.
 *  - Anyone active can work at another store for the day: phone number + PIN on that store's iPad.
 *    They stay on that store's list until midnight. Never clocked in at two stores at once.
 *  - Only the office (Juliana, Andre) fixes punches. Nothing on the iPads edits time, except
 *    "Not me? Undo" within 2 minutes of a punch. Every change is in the Edit Log with a reason.
 *    Punches are never deleted, only voided.
 *
 * Data: SYSTEM · Time Clock (punches, devices, visits, logs). People, stores, schedules:
 * SYSTEM · Core Data (read only). PINs: SYSTEM · Clock PINs (created by setupV2).
 */

const IDS = {
  CORE: '1UzZAZtnMGV46KGU1c-1kLcXw2M7utqaj__My_jgRppU',
  CLOCK: '1CJ7Uk0ne56aNxvKkgL-p2QDyJlkDoH_e9pZHELu2X68',
  MEDIA_LIBRARY: '1DqrdJSpoeGdXKGwqFLcet6ocTd8eLzSC'
};
const TZ = 'America/New_York';
/** Who can open the Office page. Email → Prime One ID. */
const OFFICE_EMAILS = { 'andre@primeoneinc.com': 'PO-0001', 'juliana@primeoneinc.com': 'PO-0031' };
const ALERT_EMAILS = ['andre@primeoneinc.com', 'juliana@primeoneinc.com'];
const OVERHEAD_STORE = 'PRIMEHQ';
const PIN_LEN = 6;
const SLOW_AFTER = 5, SLOW_SECONDS = 30, LOCK_AFTER = 10;
const SETUP_TRIES = 8, SETUP_WINDOW_SEC = 900;   // wrong setup codes per iPad per 15 minutes
const PUNCH_TOKEN_MIN = 3;
const PIN_ROUNDS = 100;           // HMAC rounds per PIN check (slows offline cracking of a leaked sheet)
const DEVICE_CODE_FAILS = 20;     // wrong iPad codes per hour (all devices) before iPad setup turns off
const SETUP_FAILS_GLOBAL = 30;    // wrong PIN setup codes per hour (all iPads) before an email alert
const UNKNOWN_PHONE_TRIES = 10;   // unknown cell numbers per iPad per 15 minutes
const OFFLINE_TRUST_MIN = 15;     // an offline punch sent within this many minutes keeps its time unflagged
/** Shown on the iPads wherever people are told to text Juliana. Set Juliana's cell here, e.g. '(919) 555-0100'. */
const HELP_PHONE = '';
/** Final answers: an offline punch refused with one of these is dropped and reported. Anything else is retried. */
const FINAL_ERRORS = /^(PIN_NOT_SET|PIN_LOCKED|PIN_WAIT|NOT_ON_FILE|Wrong PIN|Enter your|This phone is set up|You.re not set up to clock in|Tap your own name)/;
const OPEN_SHIFT_HOURS = 16;      // an IN older than this with no OUT is a missed clock-out
const GRACE_MIN = 5;              // minutes before someone counts as late / left early
const CODE_HOURS = 48;            // device setup codes expire after this
const OFFLINE_MAX_DAYS = 7;       // an offline punch older than this is recorded at receive time
const CONFIRM_OUT_MIN = 15;
const W1_MONDAY = Date.UTC(2025, 11, 29); // Week 1 of 2026 starts Mon Dec 29, 2025
/** Raise with the app's VERSION. Devices reload themselves when idle after it changes. */
const APP_VERSION = '2.0.0';

/* ============================ entry points ============================ */

function doGet() {
  let email = '';
  try { email = String(Session.getActiveUser().getEmail() || '').toLowerCase(); } catch (e) {}
  if (OFFICE_EMAILS[email]) {
    return HtmlService.createHtmlOutputFromFile('Office')
      .setTitle('Prime One Clock · Office')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  // Nobody else gets an HTML page (so nobody else gets a way to call the office functions).
  return json_({ ok: true, service: 'Prime One Clock', time: now_() });
}

function doPost(e) {
  let b = {};
  try {
    b = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    return json_(Object.assign({ ok: true, appVersion: APP_VERSION, serverMs: Date.now() }, route_(b)));
  } catch (err) {
    const msg = String((err && err.message) || err);
    return json_({ ok: false, appVersion: APP_VERSION, serverMs: Date.now(), error: msg, final: FINAL_ERRORS.test(msg) || msg === 'DEVICE_NOT_SET_UP' });
  }
}

/** Kiosk actions only. */
function route_(b) {
  switch (b.action) {
    case 'ping': return { time: now_() };
    case 'setupDevice': return setupDevice_(b.code, b.name);
    case 'roster': { const d = device_(b.device); noteOnline_(d); return roster_(d); }
    case 'clock': { const d = device_(b.device); if (!b.queued) noteOnline_(d); return clock_(d, b); }
    case 'checkCode': { const d = device_(b.device); noteOnline_(d); return checkCode_(d, b); }
    case 'setPin': { const d = device_(b.device); noteOnline_(d); return setPin_(d, b); }
    case 'undo': return undo_(device_(b.device), b.ticket, b.punchId);
    default: throw new Error('Unknown action.');
  }
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ============================ devices ============================ */

function newDeviceCode_(who, storeId, name, employeeId) {
  const store = activeStores_().find(s => s.id === storeId);
  if (!store) throw new Error('Choose a store.');
  let person = null;
  if (employeeId) {
    person = clockPeople_().find(p => p.id === employeeId);
    if (!person) throw new Error('Choose the person this phone belongs to.');
  }
  return withLock_(() => {
    const code = rand6_();
    const id = nextId_(IDS.CLOCK, 'Devices', 'Device ID', 'DEV-', 4);
    appendRow_(IDS.CLOCK, 'Devices', {
      'Device ID': id, 'Store ID': storeId, 'Device Name': text_(name, 60) || (person ? person.name + ' phone' : store.name + ' iPad'),
      'Employee ID': person ? person.id : '',
      'Token Hash': '', 'Setup Code': code, 'Code Expires': stamp_(new Date(Date.now() + CODE_HOURS * 3600000)),
      'Created On': now_(), 'Created By': who.name, 'Status': 'Waiting for setup', 'Last Seen': ''
    });
    return { deviceId: id, code: code, expires: CODE_HOURS + ' hours', store: store.name };
  });
}

function setupDevice_(code, name) {
  const c = String(code || '').replace(/\D/g, '');
  if (c.length !== 6) throw new Error('Enter the 6-digit iPad code.');
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('DEVICE_SETUP_OFF') === '1') throw new Error('iPad setup is paused. Text Andre or Juliana to turn it back on.');
  // Look for the code without the lock, so wrong guesses never slow down real punches.
  const rows = readTable_(IDS.CLOCK, 'Devices');
  const ok = rows.some(r => String(r['Setup Code']) === c && r['Status'] === 'Waiting for setup' && asTime_(r['Code Expires']) > Date.now());
  if (!ok) {
    withLock_(() => {
      const fails = bumpWindow_('DEVFAIL', 3600);
      if (fails === DEVICE_CODE_FAILS) {
        props.setProperty('DEVICE_SETUP_OFF', '1');
        accessLog_('', '', 'device setup', '', 'Paused after ' + fails + ' wrong codes in an hour', '');
        alert_('Prime One Clock: iPad setup paused', fails + ' wrong iPad codes were tried in one hour, so new iPad setup is paused. If that was you setting up iPads, turn it back on in Prime One Clock Office › Devices. If not, someone is guessing codes.');
      } else if (fails < DEVICE_CODE_FAILS) accessLog_('', '', 'device setup', '', 'Bad code', '');
    });
    throw new Error('That code is not valid, or it expired. Ask the office for a new one.');
  }
  return withLock_(() => {
    const rows = readTable_(IDS.CLOCK, 'Devices', true);
    const i = rows.findIndex(r => String(r['Setup Code']) === c && r['Status'] === 'Waiting for setup' && asTime_(r['Code Expires']) > Date.now());
    if (i === -1) throw new Error('That code is not valid, or it expired. Ask the office for a new one.');
    const r = rows[i];
    const token = Utilities.getUuid() + Utilities.getUuid();
    setCell_(IDS.CLOCK, 'Devices', i, 'Token Hash', hash_('dev', token));
    setCell_(IDS.CLOCK, 'Devices', i, 'Setup Code', '');
    setCell_(IDS.CLOCK, 'Devices', i, 'Status', 'Active');
    setCell_(IDS.CLOCK, 'Devices', i, 'Last Seen', now_());
    if (name) setCell_(IDS.CLOCK, 'Devices', i, 'Device Name', text_(name, 60));
    accessLog_(r['Device ID'], r['Store ID'], 'device setup', '', 'OK', '');
    return { device: r['Device ID'] + '.' + token, deviceId: r['Device ID'], storeId: r['Store ID'], storeName: storeName_(r['Store ID']) };
  });
}

/** The last time this device talked to the server while online (used to judge "offline" punches). */
function noteOnline_(dev) {
  try {
    const props = PropertiesService.getScriptProperties(), k = 'ON_' + dev.id;
    if (Date.now() - Number(props.getProperty(k) || 0) > 60000) props.setProperty(k, String(Date.now()));
  } catch (e) {}
}

/** Device token is "DEV-0001.<secret>". Every kiosk call carries it; a revoked device stops at once. */
function device_(tok) {
  const p = String(tok || '').split('.');
  if (p.length !== 2) throw new Error('DEVICE_NOT_SET_UP');
  const rows = readTable_(IDS.CLOCK, 'Devices');
  const i = rows.findIndex(r => r['Device ID'] === p[0]);
  if (i === -1 || rows[i]['Status'] !== 'Active' || rows[i]['Token Hash'] !== hash_('dev', p[1])) throw new Error('DEVICE_NOT_SET_UP');
  const r = rows[i];
  if (Date.now() - asTime_(r['Last Seen']) > 600000) setCell_(IDS.CLOCK, 'Devices', i, 'Last Seen', now_());
  return { id: r['Device ID'], storeId: r['Store ID'], storeName: storeName_(r['Store ID']), employeeId: r['Employee ID'] || '' };
}

function revokeDevice_(who, deviceId) {
  return withLock_(() => {
    const rows = readTable_(IDS.CLOCK, 'Devices');
    const i = rows.findIndex(r => r['Device ID'] === deviceId);
    if (i === -1) throw new Error('Unknown device.');
    setCell_(IDS.CLOCK, 'Devices', i, 'Status', 'Revoked ' + today_() + ' by ' + who.name);
    setCell_(IDS.CLOCK, 'Devices', i, 'Token Hash', '');
    setCell_(IDS.CLOCK, 'Devices', i, 'Setup Code', '');
    accessLog_(deviceId, rows[i]['Store ID'], 'device revoked', '', 'OK', who.name);
    return { message: 'Removed. That device can no longer punch.' };
  });
}

/* ============================ kiosk ============================ */

/**
 * The store's list: people assigned there, plus anyone working there today. No PIN data at all,
 * only whether the person has set their PIN yet ("set", "none" or "locked").
 */
function roster_(dev) {
  const all = clockPeople_();
  const home = all.filter(p => dev.employeeId ? p.id === dev.employeeId : p.storeId === dev.storeId);
  const homeIds = {}; home.forEach(p => homeIds[p.id] = true);
  const t = today_();
  const visitIds = {};
  if (!dev.employeeId) readTable_(IDS.CLOCK, 'Visits').forEach(v => {
    if (asDay_(v['Date']) === t && v['Store ID'] === dev.storeId && !homeIds[v['Employee ID']]) visitIds[v['Employee ID']] = true;
  });
  const visitors = all.filter(p => visitIds[p.id]);
  const last = lastPunches_(), creds = credStates_();
  const people = home.concat(visitors).map(p => {
    const st = stateOf_(last[p.id]);
    return {
      id: p.id, name: p.name, role: p.role, status: st.status, since: st.since, sinceLabel: st.sinceLabel,
      pin: creds[p.id] || 'none',
      visiting: visitIds[p.id] ? storeName_(p.storeId) : '',
      elsewhere: st.status === 'in' && st.storeId && st.storeId !== dev.storeId ? storeName_(st.storeId) : ''
    };
  }).sort((a, b) => a.name.localeCompare(b.name));
  return { storeId: dev.storeId, storeName: dev.storeName, deviceId: dev.id, serverTime: now_(), personal: !!dev.employeeId, helpPhone: HELP_PHONE, people: people };
}

function readTicket_(dev, ticket) {
  const v = readToken_('T', ticket).split(':');
  if (v[1] !== dev.id) throw new Error('Please enter your PIN again.');
  return v[0];
}

/**
 * One step at the kiosk: PIN + photo in, punch saved. The server decides IN or OUT from the
 * person's last punch. Every reply carries the store's fresh list, so the iPad never polls.
 * An offline punch (queued) that is refused for a final reason is logged on the Refused tab.
 */
function clock_(dev, b) {
  const pin = String(b.pin || '');
  if (!new RegExp('^\\d{' + PIN_LEN + '}$').test(pin)) throw new Error('Enter your ' + PIN_LEN + '-digit PIN.');
  const cid = String(b.cid || '');
  if (cid && !/^[a-z0-9]{8,40}$/.test(cid)) throw new Error('Enter your PIN again.');
  return withLock_(() => {
    // The same punch sent twice (slow answer + saved copy, or a resend after a reload) counts once.
    if (cid) {
      const cache = CacheService.getScriptCache(), seen = cache.get('cid_' + cid);
      if (seen) { const o = JSON.parse(seen); if (o.dev === dev.id) return o.r; }
      const dup = recentRows_(600).find(x => String(x['Client ID']) === cid);
      if (dup) return { duplicate: true, punchId: dup['Punch ID'], type: dup['Type'], time: hm_(str_(dup['Punch Time'])), name: String(dup['Employee Name']).split(' ')[0], roster: roster_(dev) };
    }
    let p;
    try {
      p = findPerson_(dev, b.employeeId, b.phone);
      if (!p.clocksIn) throw new Error('You’re not set up to clock in. Text Juliana.');
      checkPin_(p, pin, dev);
    } catch (e) {
      if (b.queued && FINAL_ERRORS.test(String(e.message))) logRefused_(dev, p, b, e.message);
      throw e;
    }
    const r = punchNow_(dev, p, b, cid);
    if (cid && !r.confirm && !r.switchStore) { try { CacheService.getScriptCache().put('cid_' + cid, JSON.stringify({ dev: dev.id, r: r }), 21600); } catch (e) {} }
    return r;
  });
}

function punchNow_(dev, p, b, cid) {
  const first = p.first || p.name.split(' ')[0];
  const last = lastPunches_()[p.id];
  const when = punchTime_(dev, b, last);
  const at = when.at || new Date();
  const notes = when.note ? [when.note] : [];
  let st = stateOf_(last, at.getTime());

  // Still clocked in at another store: one tap closes that and opens this one.
  // Offline: only if the other store's clock-in came before this punch; otherwise it's flagged.
  if (st.status === 'in' && st.storeId && st.storeId !== dev.storeId && !dev.employeeId) {
    if (!b.queued && !b.switchStore) return { switchStore: true, name: first, otherStore: storeName_(st.storeId), sinceLabel: st.sinceLabel };
    if (!when.outOfOrder) {
      doPunch_({ id: dev.id, storeId: st.storeId }, p, 'OUT', '', st, at,
        notes.concat(['Clocked out by clocking in at ' + dev.storeName + '.']).join(' '), cid ? cid + 'x' : '');
      st = { status: 'out', since: '', sinceLabel: '', openSince: '', storeId: '' };
    }
  }

  let type = st.status === 'in' ? 'OUT' : 'IN';
  if (when.outOfOrder) {
    // Sent after a newer punch already arrived: keep what the iPad showed, and ask Juliana to check.
    if (b.guess === 'IN' || b.guess === 'OUT') type = b.guess;
    notes.push('REVIEW: arrived after a newer punch; check the order of this day.');
  } else if (b.queued && (b.guess === 'IN' || b.guess === 'OUT') && b.guess !== type) {
    notes.push('REVIEW: the iPad showed ' + (b.guess === 'IN' ? 'clock in' : 'clock out') + '.');
  }
  if (type === 'OUT' && !b.confirm && !b.queued && Date.now() - asTime_(st.since) < CONFIRM_OUT_MIN * 60000) {
    return { confirm: true, name: first, sinceLabel: st.sinceLabel };
  }
  if (type === 'IN' && st.openSince) notes.push('Previous clock-in ' + st.openSince + ' has no clock-out.');
  if (!dev.employeeId && p.storeId !== dev.storeId) recordVisit_(p, dev, at);
  const r = doPunch_(dev, p, type, b.photo, st, at, notes.join(' '), cid);
  return Object.assign(r, {
    ticket: makeToken_('T', p.id + ':' + dev.id, PUNCH_TOKEN_MIN), employeeId: p.id, fullName: p.name,
    hoursToday: type === 'OUT' ? hoursToday_(p.id) : null,
    visiting: !dev.employeeId && p.storeId !== dev.storeId,
    roster: roster_(dev)
  });
}

/**
 * When the punch happened. Online: the server's clock, always. Offline (queued): the time the iPad
 * recorded, but marked REVIEW for Juliana when any of these is true —
 *  - it was made on a personal phone (that phone's owner controls it);
 *  - it arrived more than 15 minutes after it was made;
 *  - the iPad wasn't really offline (it talked to the server online after that time);
 *  - the iPad's stopwatch and clock disagreed ("device" time).
 * A punch older than the person's newest punch is "out of order" and never moves their state.
 */
function punchTime_(dev, b, last) {
  if (!b.queued) return { at: null, note: '' };
  const t = Number(b.at), now = Date.now();
  if (!t || t > now + 60000 || t < now - OFFLINE_MAX_DAYS * 86400000) {
    return { at: null, note: 'REVIEW: made offline, time out of range; recorded when received ' + now_() + '.' };
  }
  const at = new Date(Math.min(t, now));
  const why = [];
  if (b.timeSource !== 'anchored') why.push('time is from the iPad clock and could not be checked');
  if (dev.employeeId) why.push('made on a personal phone');
  if (now - at.getTime() > OFFLINE_TRUST_MIN * 60000) why.push('sent ' + dur_((now - at.getTime()) / 60000) + ' after it was made');
  const online = Number(PropertiesService.getScriptProperties().getProperty('ON_' + dev.id) || 0);
  if (online && online > at.getTime() + 60000) why.push('this device was online after that time');
  const outOfOrder = !!last && at.getTime() < asTime_(last['Punch Time']);
  return {
    at: at, outOfOrder: outOfOrder,
    note: 'Made offline; sent ' + now_() + '.' + (why.length ? ' REVIEW: ' + why.join('; ') + '.' : '')
  };
}

/**
 * The person punching. A tapped name must be on this iPad's list (its store, or here today);
 * anyone else uses their cell number. One answer for every miss, so numbers can't be tested.
 */
function findPerson_(dev, employeeId, phone) {
  const everyone = clockPeople_(true);
  let p;
  if (employeeId) {
    p = everyone.find(x => x.id === employeeId);
    if (p && dev.employeeId && p.id !== dev.employeeId) throw new Error('This phone is set up for someone else.');
    if (p && !dev.employeeId && p.storeId !== dev.storeId && !visitingToday_(p.id, dev.storeId)) p = null;
    if (!p) throw new Error('Tap your own name, or use “Not on the list?”.');
    return p;
  }
  const d = digits_(phone);
  if (d.length !== 10) throw new Error('Enter your 10-digit cell number.');
  const k = 'UNK_' + dev.id;
  if (windowCount_(k) >= UNKNOWN_PHONE_TRIES) throw new Error('NOT_ON_FILE');
  p = everyone.find(x => digits_(x.phone) === d);
  if (!p || (dev.employeeId && p.id !== dev.employeeId)) {
    bumpWindow_(k, 900);
    accessLog_(dev.id, dev.storeId, 'sign in', '', 'Unknown phone', d.slice(-4));
    throw new Error('NOT_ON_FILE');
  }
  return p;
}
function visitingToday_(empId, storeId) {
  const t = today_();
  return readTable_(IDS.CLOCK, 'Visits').some(v => asDay_(v['Date']) === t && v['Employee ID'] === empId && v['Store ID'] === storeId);
}

/** A punch saved offline that was refused (wrong PIN, locked…). Juliana sees it on Today. */
function logRefused_(dev, p, b, why) {
  try {
    appendRow_(IDS.CLOCK, 'Refused', {
      'Refused At': now_(), 'Made At': Number(b.at) ? stamp_(new Date(Number(b.at))) : '', 'Store ID': dev.storeId, 'Device ID': dev.id,
      'Employee ID': p ? p.id : '', 'Employee Name': p ? p.name : '', 'Cell (last 4)': p ? '' : digits_(b.phone).slice(-4),
      'Showed As': b.guess || '', 'Reason': plainReason_(why), 'Status': 'Open', 'Handled By': ''
    });
  } catch (e) {}
}
function plainReason_(m) {
  m = String(m || '');
  if (/^PIN_NOT_SET/.test(m)) return 'No PIN set up yet';
  if (/^PIN_LOCKED/.test(m)) return 'PIN locked';
  if (/^PIN_WAIT|^Wrong PIN/.test(m)) return 'Wrong PIN';
  if (/^NOT_ON_FILE/.test(m)) return 'Cell number not on file';
  return m.slice(0, 120);
}

/** One row per person, per store, per day they worked away from their own store. */
function recordVisit_(p, dev, at) {
  const day = Utilities.formatDate(at, TZ, 'yyyy-MM-dd');
  const seen = readTable_(IDS.CLOCK, 'Visits').some(v => asDay_(v['Date']) === day && v['Employee ID'] === p.id && v['Store ID'] === dev.storeId);
  if (seen) return;
  appendRow_(IDS.CLOCK, 'Visits', {
    'Date': day, 'Employee ID': p.id, 'Employee Name': p.name, 'Home Store': p.storeId, 'Store ID': dev.storeId,
    'Device ID': dev.id, 'First Punch': stamp_(at)
  });
}

/** Undo a kiosk punch within 2 minutes (wrong name tapped, tapped twice). Kept in the Edit Log. */
function undo_(dev, ticket, punchId) {
  const empId = readTicket_(dev, ticket);
  return withLock_(() => {
    const rows = recentRows_(300);
    const pu = rows.find(x => x['Punch ID'] === punchId);
    if (!pu || pu['Employee ID'] !== empId || pu['Source'] !== 'Kiosk' || pu['Status'] !== 'Active') throw new Error('That punch cannot be undone here.');
    if (Date.now() - asTime_(pu['Recorded At']) > 2 * 60000) throw new Error('Too late to undo. Text Juliana.');
    const p = clockPeople_(true).find(x => x.id === empId) || { name: empId };
    const by = { name: p.name + ' (undo at kiosk)' };
    voidPunchCore_(by, punchId, 'Undone at the kiosk right after punching', '');
    const cid = String(pu['Client ID'] || '');
    const partner = cid ? rows.find(x => String(x['Client ID']) === cid + 'x' && x['Status'] === 'Active') : null;
    if (partner) voidPunchCore_(by, partner['Punch ID'], 'Undone at the kiosk (automatic clock-out when switching stores)', '');
    rebuildOne_(empId);
    return { message: 'Undone.', roster: roster_(dev) };
  });
}

function hoursToday_(empId) {
  const t = today_();
  const list = recentRows_(1500).filter(x => x['Employee ID'] === empId && x['Status'] === 'Active' && asDay_(x['Punch Time']) === t)
    .sort((a, b) => asTime_(a['Punch Time']) - asTime_(b['Punch Time']));
  return round2_(pair_(list).minutes / 60);
}

function doPunch_(dev, p, type, photo, st, at, notes, cid) {
  const empId = p.id;
  const when = at || new Date();
  const id = nextPunchId_();
  let photoUrl = 'No photo';
  if (photo) { try { photoUrl = savePhoto_(photo, dev.storeId, empId, type, when); } catch (e) { photoUrl = 'Photo failed: ' + e.message; } }
  appendRow_(IDS.CLOCK, 'Punches', {
    'Punch ID': id, 'Employee ID': empId, 'Employee Name': p.name, 'Store ID': dev.storeId, 'Type': type,
    'Punch Time': stamp_(when), 'Recorded At': stamp_(new Date()), 'Photo': photoUrl, 'Device ID': dev.id,
    'Source': 'Kiosk', 'Status': 'Active', 'Fix Request ID': '', 'Notes': notes || '', 'Client ID': cid || ''
  });
  // Only move "last punch" forward: an old offline punch never overrides a newer one.
  const last = lastPunches_()[empId];
  if (!last || when.getTime() >= asTime_(last['Punch Time'])) setLast_(empId, id, type, stamp_(when), dev.storeId);
  return {
    punchId: id, type: type, time: Utilities.formatDate(when, TZ, 'h:mm a'), name: p.first || p.name,
    noPhoto: photoUrl.indexOf('http') !== 0,
    openWarning: type === 'IN' ? (st.openSince || '') : ''
  };
}

/** Small JPEG to Media Library / Time Clock Photos / <store> / <yyyy-MM>. */
function savePhoto_(photo, storeId, empId, type, at) {
  const m = String(photo || '').match(/^data:image\/jpeg;base64,(.+)$/);
  if (!m) return 'No photo';
  const bytes = Utilities.base64Decode(m[1]);
  if (bytes.length > 250000) throw new Error('too large');
  const month = Utilities.formatDate(at, TZ, 'yyyy-MM');
  const props = PropertiesService.getScriptProperties(), key = 'PF_' + storeId + '_' + month;
  let folder = null;
  const known = props.getProperty(key);
  if (known) { try { folder = DriveApp.getFolderById(known); } catch (e) { folder = null; } }
  if (!folder) {
    const root = childFolder_(childFolder_(DriveApp.getFolderById(IDS.MEDIA_LIBRARY), 'Time Clock Photos'), storeId);
    folder = childFolder_(root, month);
    props.setProperty(key, folder.getId());
  }
  const name = Utilities.formatDate(at, TZ, 'yyyy-MM-dd_HHmmss') + '_' + empId + '_' + type + '.jpg';
  return folder.createFile(Utilities.newBlob(bytes, 'image/jpeg', name)).getUrl();
}

/* ============================ PINs ============================ */

function credId_() {
  const id = PropertiesService.getScriptProperties().getProperty('CRED_ID');
  if (!id) throw new Error('The clock is not finished setting up. Text Andre.');
  return id;
}
function pinPepper_() {
  const props = PropertiesService.getScriptProperties();
  let s = props.getProperty('PIN_PEPPER');
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); props.setProperty('PIN_PEPPER', s); }
  return s;
}
/** Iterated HMAC with a pepper kept only in this script's settings (never in the sheet). */
function pinHash_(salt, pin) {
  const key = pinPepper_();
  let h = 'pin:' + salt + ':' + pin;
  for (let i = 0; i < PIN_ROUNDS; i++) h = Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(h, key));
  return h;
}

/** { 'PO-0047': 'set' | 'locked' } — anyone missing has not set a PIN yet. */
function credStates_() {
  const out = {};
  readTable_(credId_(), 'PINs').forEach(r => {
    if (r['Status'] !== 'Active') return;
    out[r['Employee ID']] = r['Locked'] === 'Yes' ? 'locked' : 'set';
  });
  return out;
}

/**
 * The only place a PIN is checked. After 5 wrong in a row: 30 seconds between tries.
 * At 10: locked until the office unlocks it, and Andre and Juliana get an email.
 */
function checkPin_(p, pin, dev) {
  const cid = credId_();
  const rows = readTable_(cid, 'PINs');
  const i = rows.findIndex(r => r['Employee ID'] === p.id && r['Status'] === 'Active');
  if (i === -1) throw new Error('PIN_NOT_SET');
  const r = rows[i];
  if (r['Locked'] === 'Yes') throw new Error('PIN_LOCKED');
  let tries = Number(r['Failed Tries'] || 0);
  const lastFail = asTime_(r['Last Fail']);
  if (tries && Date.now() - lastFail > 86400000) tries = 0;           // old typos don't add up over weeks
  if (tries >= SLOW_AFTER && Date.now() - lastFail < SLOW_SECONDS * 1000) {
    throw new Error('PIN_WAIT:' + Math.ceil((SLOW_SECONDS * 1000 - (Date.now() - lastFail)) / 1000) + ':' + (LOCK_AFTER - tries));
  }
  if (r['PIN Hash'] !== pinHash_(String(r['Salt']), pin)) {
    tries++;
    setCell_(cid, 'PINs', i, 'Failed Tries', tries);
    setCell_(cid, 'PINs', i, 'Last Fail', stamp_(new Date()));
    accessLog_(dev ? dev.id : '', dev ? dev.storeId : '', 'pin', p.id, 'Wrong PIN (' + tries + ')', '');
    if (tries >= LOCK_AFTER) {
      setCell_(cid, 'PINs', i, 'Locked', 'Yes');
      setCell_(cid, 'PINs', i, 'Locked At', now_());
      alertLocked_(p, dev);
      throw new Error('PIN_LOCKED');
    }
    if (tries >= SLOW_AFTER) throw new Error('PIN_WAIT:' + SLOW_SECONDS + ':' + (LOCK_AFTER - tries));
    throw new Error('Wrong PIN. Check that you tapped your own name, then try again.');
  }
  if (tries || r['Failed Tries']) setCell_(cid, 'PINs', i, 'Failed Tries', 0);
  setCell_(cid, 'PINs', i, 'Last Used', now_());
}

function alertLocked_(p, dev) {
  alert_('Prime One Clock: ' + p.name + '’s PIN is locked',
    p.name + ' entered a wrong PIN 10 times' + (dev ? ' at ' + dev.storeName : '') + ' (' + now_() + ').\n\n' +
    'They can’t clock in until it’s unlocked. Open Prime One Clock Office › Today, then Unlock (same PIN) or Reset PIN (you get a one-time code to text them).');
}
function alert_(subject, body) {
  try { MailApp.sendEmail({ to: ALERT_EMAILS.join(','), subject: subject, body: body }); } catch (e) { /* email never blocks the clock */ }
}

/** The shared code people use the first time, to choose their own PIN. */
function setupCode_() {
  const props = PropertiesService.getScriptProperties();
  let c = props.getProperty('PIN_SETUP_CODE');
  if (!c) { c = newSetupCodeValue_(); props.setProperty('PIN_SETUP_CODE', c); }
  return c;
}
function newSetupCodeValue_() {
  let c;
  do { c = rand6_(); } while (weakPin_(c, ''));
  return c;
}

/** Simple rules: not one digit repeated, not a run (123456 or 654321), not a repeated pair, not the
 * setup code, not the last 6 digits of their phone. */
function weakPin_(pin, phone) {
  if (/^(\d)\1+$/.test(pin)) return 'Pick a PIN that isn’t the same digit over and over.';
  if ('01234567890'.indexOf(pin) !== -1 || '09876543210'.indexOf(pin) !== -1) return 'Pick a PIN that isn’t a run of numbers like 123456.';
  if (/^(\d\d)\1\1$/.test(pin) || /^(\d\d\d)\1$/.test(pin)) return 'Pick a PIN that isn’t a repeating pattern.';
  if (phone && digits_(phone).slice(-6) === pin) return 'Pick a PIN that isn’t part of your phone number.';
  const props = PropertiesService.getScriptProperties();
  if (pin === props.getProperty('PIN_SETUP_CODE')) return 'Pick your own PIN, not the setup code.';
  return '';
}

/**
 * Which code this person needs: the shared setup code the first time, or the one-time reset code
 * Juliana got when she reset their PIN. Wrong codes are counted per iPad and across all iPads.
 */
function codeCheck_(dev, p, code) {
  const k = 'SC_' + dev.id;
  if (windowCount_(k) >= SETUP_TRIES) throw new Error('Too many wrong codes on this iPad. Wait 15 minutes, or text Juliana' + helpSuffix_() + '.');
  const cid = credId_(), rows = readTable_(cid, 'PINs');
  const i = rows.findIndex(r => r['Employee ID'] === p.id);
  const row = i === -1 ? null : rows[i];
  if (row && row['Status'] === 'Active') throw new Error('You already have a PIN. If you forgot it, text Juliana' + helpSuffix_() + ' to reset it.');
  const isReset = row && /^Reset/.test(String(row['Status']));
  let good;
  if (isReset) good = String(row['Reset Code']) === hash_('reset', p.id + ':' + code) && asTime_(row['Reset Expires']) > Date.now();
  else good = String(code || '') === setupCode_();
  if (!good) {
    bumpWindow_(k, SETUP_WINDOW_SEC);
    const all = bumpWindow_('SCALL', 3600);
    if (all === SETUP_FAILS_GLOBAL) alert_('Prime One Clock: many wrong setup codes', all + ' wrong PIN setup codes were entered in the last hour across the store iPads. If that wasn’t a busy setup day, change the setup code in Office › People & PINs.');
    accessLog_(dev.id, dev.storeId, 'pin setup', p.id, 'Wrong code', '');
    throw new Error(isReset ? 'That code is not right. Use the code Juliana texted you, or text her' + helpSuffix_() + ' for a new one.' : 'That setup code is not right. Ask your manager, or text Juliana' + helpSuffix_() + '.');
  }
  return { i: i, isReset: isReset };
}
function helpSuffix_() { return HELP_PHONE ? ' at ' + HELP_PHONE : ''; }

/** Step 1 of first sign-in: is the code right? (So nobody picks a PIN twice for nothing.) */
function checkCode_(dev, b) {
  return withLock_(() => {
    const p = findPerson_(dev, b.employeeId, b.phone);
    const c = codeCheck_(dev, p, String(b.setupCode || ''));
    return { ok: true, reset: c.isReset, name: p.first || p.name.split(' ')[0] };
  });
}

/**
 * First sign-in (or after the office resets a PIN): the code, then the person's own PIN.
 * The iPad then clocks them in with the new PIN.
 */
function setPin_(dev, b) {
  const pin = String(b.newPin || '');
  if (!new RegExp('^\\d{' + PIN_LEN + '}$').test(pin)) throw new Error('Your PIN must be ' + PIN_LEN + ' digits.');
  return withLock_(() => {
    const p = findPerson_(dev, b.employeeId, b.phone);
    const c = codeCheck_(dev, p, String(b.setupCode || ''));
    const weak = weakPin_(pin, p.phone);
    if (weak) throw new Error(weak);
    const cid = credId_();
    const salt = Utilities.getUuid();
    const vals = {
      'Employee ID': p.id, 'Employee Name': p.name, 'PIN Hash': pinHash_(salt, pin), 'Salt': salt, 'Status': 'Active',
      'Set On': now_(), 'Set By': 'Self, ' + dev.storeId + ' (' + dev.id + ')' + (c.isReset ? ' after reset' : ''), 'Failed Tries': 0, 'Last Fail': '', 'Locked': '', 'Locked At': '', 'Last Used': '',
      'Reset Code': '', 'Reset Expires': ''
    };
    if (c.i === -1) appendRow_(cid, 'PINs', vals);
    else Object.keys(vals).forEach(k => setCell_(cid, 'PINs', c.i, k, vals[k]));
    accessLog_(dev.id, dev.storeId, 'pin setup', p.id, 'PIN set', '');
    return { message: 'PIN saved.' };
  });
}

/* ============================ office (Google sign-in) ============================ */

/** Every office action starts here. The signed-in Google account must be on OFFICE_EMAILS. */
function officeWho_() {
  let email = '';
  try { email = String(Session.getActiveUser().getEmail() || '').toLowerCase(); } catch (e) {}
  const id = OFFICE_EMAILS[email];
  if (!id) throw new Error('Office access only. Sign in with your primeoneinc.com account.');
  const e = coreTable_('Employees').find(x => x['Employee ID'] === id) || {};
  return { id: id, name: e['Display Name'] || email, email: email };
}

/** The Office page calls only this (google.script.run.officeCall). */
function officeCall(action, data) {
  const who = officeWho_();
  const d = data || {};
  switch (action) {
    case 'home': { const home = officeHome_(who); return { home: home, sheet: timesheet_(who, 'ALL', home.currentWeek) }; }
    case 'timesheet': return timesheet_(who, d.storeId, d.weekStart);
    case 'editPunch': must_(d.reason, 'Reason'); return withLock_(() => { const e = editPunchCore_(who, d.punchId, d.time, d.type, d.reason, ''); rebuildOne_(e); return { message: 'Saved.' }; });
    case 'okPunch': return withLock_(() => { okPunchCore_(who, d.punchId); return { message: 'Marked as checked.' }; });
    case 'voidPunch': must_(d.reason, 'Reason'); return withLock_(() => { const e = voidPunchCore_(who, d.punchId, d.reason, ''); rebuildOne_(e); return { message: 'Removed from hours (kept in the log).' }; });
    case 'addPunch': must_(d.reason, 'Reason'); return withLock_(() => { const id = addPunchCore_(who, d.employeeId, d.storeId, d.type, d.time, d.reason, ''); rebuildOne_(d.employeeId); return { punchId: id, message: 'Added.' }; });
    case 'closeRefused': return withLock_(() => { closeRefused_(who, d.row); return { message: 'Done.' }; });
    case 'deviceSetupOn': return withLock_(() => { const pr = PropertiesService.getScriptProperties(); pr.deleteProperty('DEVICE_SETUP_OFF'); pr.deleteProperty('W_DEVFAIL'); accessLog_('', '', 'device setup', who.id, 'Turned back on', ''); return { message: 'iPad setup is on again.' }; });
    case 'newDeviceCode': return newDeviceCode_(who, d.storeId, d.name, d.employeeId);
    case 'revokeDevice': return revokeDevice_(who, d.deviceId);
    case 'pins': return pinList_();
    case 'resetPin': return resetPin_(who, d.employeeId);
    case 'unlockPin': return unlockPin_(who, d.employeeId);
    case 'newSetupCode': return withLock_(() => { const c = newSetupCodeValue_(); PropertiesService.getScriptProperties().setProperty('PIN_SETUP_CODE', c); accessLog_('', '', 'setup code changed', who.id, 'OK', ''); return { code: c }; });
    default: throw new Error('Unknown action.');
  }
}

function officeHome_(who) {
  return withLock_(() => {
    const stores = activeStores_().filter(s => s.id !== OVERHEAD_STORE);
    const devices = readTable_(IDS.CLOCK, 'Devices').filter(d => d['Status'] === 'Active' || d['Status'] === 'Waiting for setup').map(d => ({
      id: d['Device ID'], storeId: d['Store ID'], name: d['Device Name'], status: d['Status'],
      code: d['Status'] === 'Waiting for setup' ? str_(d['Setup Code']) : '', lastSeen: str_(d['Last Seen'])
    }));
    const last = lastPunches_(), creds = credStates_();
    const people = clockPeople_().map(p => Object.assign({ id: p.id, name: p.name, storeId: p.storeId, role: p.role, pin: creds[p.id] || 'none' }, stateOf_(last[p.id])));
    const t = today_();
    const visits = readTable_(IDS.CLOCK, 'Visits').filter(v => asDay_(v['Date']) === t)
      .map(v => ({ employeeId: v['Employee ID'], name: v['Employee Name'], homeStore: v['Home Store'], storeId: v['Store ID'], first: hm_(str_(v['First Punch'])) }));
    const refused = readTable_(IDS.CLOCK, 'Refused').map((r, i) => Object.assign({ row: i }, r)).filter(r => r['Status'] === 'Open').map(r => ({
      row: r.row, at: str_(r['Made At']), storeId: r['Store ID'], employeeId: r['Employee ID'], name: r['Employee Name'] || ('Cell ending ' + r['Cell (last 4)']),
      showed: r['Showed As'], reason: r['Reason']
    }));
    return {
      me: who.name, today: t, currentWeek: mondayOf_(t), stores: stores, devices: devices, people: people, visits: visits,
      setupCode: setupCode_(), refused: refused,
      deviceSetupOff: PropertiesService.getScriptProperties().getProperty('DEVICE_SETUP_OFF') === '1'
    };
  });
}

function pinList_() {
  const rows = {};
  readTable_(credId_(), 'PINs').forEach(r => rows[r['Employee ID']] = r);
  return {
    setupCode: setupCode_(),
    people: clockPeople_().map(p => {
      const r = rows[p.id];
      const active = r && r['Status'] === 'Active';
      return {
        id: p.id, name: p.name, storeId: p.storeId,
        state: !active ? 'none' : r['Locked'] === 'Yes' ? 'locked' : 'set',
        setOn: active ? str_(r['Set On']) : '', lastUsed: active ? str_(r['Last Used']) : '',
        tries: active ? Number(r['Failed Tries'] || 0) : 0, lockedAt: active ? str_(r['Locked At']) : '',
        resetPending: !!r && /^Reset/.test(String(r['Status']))
      };
    }).sort((a, b) => a.name.localeCompare(b.name))
  };
}

/**
 * Clears the PIN and makes a one-time code for this person only (good for 3 days). They tap their
 * name on any store iPad, enter that code, and choose a new PIN. The shared setup code won't work
 * for them, so nobody else can take over the account after a reset.
 */
function resetPin_(who, empId) {
  return withLock_(() => {
    const cid = credId_(), rows = readTable_(cid, 'PINs');
    const p = clockPeople_(true).find(x => x.id === empId);
    if (!p) throw new Error('Unknown employee.');
    const i = rows.findIndex(r => r['Employee ID'] === empId);
    if (i === -1) return { message: (p.first || p.name) + ' has no PIN yet. They set one on any store iPad with the setup code.', code: '', text: '' };
    const code = rand6_();
    ['PIN Hash', 'Salt', 'Last Fail', 'Locked', 'Locked At'].forEach(k => setCell_(cid, 'PINs', i, k, ''));
    setCell_(cid, 'PINs', i, 'Failed Tries', 0);
    setCell_(cid, 'PINs', i, 'Reset Code', hash_('reset', empId + ':' + code));
    setCell_(cid, 'PINs', i, 'Reset Expires', stamp_(new Date(Date.now() + 3 * 86400000)));
    setCell_(cid, 'PINs', i, 'Status', 'Reset ' + now_() + ' by ' + who.name);
    accessLog_('', '', 'pin reset', empId, 'OK', who.name);
    const first = p.first || p.name.split(' ')[0];
    return {
      message: 'PIN reset. Text ' + first + ' the code below.', code: code,
      text: 'Hi ' + first + ', your Prime One Clock PIN was reset. On the store iPad, tap your name, enter code ' + code + ', then choose a new 6-digit PIN. The code works for 3 days.'
    };
  });
}

function unlockPin_(who, empId) {
  return withLock_(() => {
    const cid = credId_(), rows = readTable_(cid, 'PINs');
    const i = rows.findIndex(r => r['Employee ID'] === empId && r['Status'] === 'Active');
    if (i === -1) throw new Error('No PIN on file to unlock.');
    setCell_(cid, 'PINs', i, 'Locked', '');
    setCell_(cid, 'PINs', i, 'Locked At', '');
    setCell_(cid, 'PINs', i, 'Failed Tries', 0);
    setCell_(cid, 'PINs', i, 'Last Fail', '');
    accessLog_('', '', 'pin unlock', empId, 'OK', who.name);
    return { message: 'Unlocked. Same PIN as before.' };
  });
}

/** One week (Mon–Sat) for one store or ALL. Per person per day: punches, hours, late/early flags. */
function timesheet_(who, storeId, weekStart) {
  const ws = mondayOf_(weekStart || today_());
  const days = [0, 1, 2, 3, 4, 5, 6].map(n => addDays_(ws, n));   // Sunday too, so no punch is ever left out
  const sched = schedules_();
  const people = clockPeople_(true);
  const byId = {}; people.forEach(p => byId[p.id] = p);
  const punches = readTable_(IDS.CLOCK, 'Punches').filter(p => {
    const d = asDay_(p['Punch Time']);
    return d >= ws && d <= days[6] && (storeId === 'ALL' || !storeId || p['Store ID'] === storeId || (byId[p['Employee ID']] || {}).storeId === storeId);
  });
  const ids = {};
  people.filter(p => p.clocksIn && (storeId === 'ALL' || !storeId || p.storeId === storeId)).forEach(p => ids[p.id] = true);
  punches.forEach(p => ids[p['Employee ID']] = true);
  const rows = Object.keys(ids).map(id => {
    const p = byId[id] || { id: id, name: (punches.find(x => x['Employee ID'] === id) || {})['Employee Name'] || id, storeId: '' };
    const mine = punches.filter(x => x['Employee ID'] === id);
    let total = 0, late = 0, early = 0, open = 0, review = 0;
    const storeMin = {};
    const dayRows = days.map((day, n) => {
      const list = mine.filter(x => asDay_(x['Punch Time']) === day)
        .sort((a, b) => asTime_(a['Punch Time']) - asTime_(b['Punch Time']));
      const active = list.filter(x => x['Status'] === 'Active');
      const r = pair_(active);
      Object.keys(r.byStore).forEach(s => storeMin[s] = (storeMin[s] || 0) + r.byStore[s]);
      const s = n < 6 ? schedFor_(sched, id, p, n) : null;
      const flags = [];
      if (r.open && day === today_()) flags.push('On the clock');
      else if (r.open) { flags.push('No clock-out'); open++; }
      if (r.orphans) flags.push('Clock-out without clock-in');
      const nRev = active.filter(x => /REVIEW/.test(String(x['Notes']))).length;
      if (nRev) { flags.push('Check offline time'); review += nRev; }
      if (Object.keys(r.byStore).length > 1 || Object.keys(r.byStore).some(k => k !== p.storeId && p.storeId)) flags.push('Worked at ' + Object.keys(r.byStore).map(storeName_).join(' + '));
      let lateMin = 0, earlyMin = 0;
      if (s && r.firstIn) {
        lateMin = Math.round((minsOf_(r.firstIn) - s.start));
        if (lateMin > GRACE_MIN) { flags.push('Late ' + dur_(lateMin)); late++; } else lateMin = 0;
      }
      if (s && r.lastOut && !r.open) {
        earlyMin = Math.round(s.end - minsOf_(r.lastOut));
        if (earlyMin > GRACE_MIN) { flags.push('Left ' + dur_(earlyMin) + ' early'); early++; } else earlyMin = 0;
      }
      if (s && !active.length && day < today_()) flags.push('No punches');
      if (s && !active.length && day === today_() && nowMin_() > s.start + GRACE_MIN) flags.push('Not in yet');
      total += r.minutes;
      return {
        date: day, dow: ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'][n], schedule: s ? s.label : '',
        firstIn: r.firstIn ? hm_(r.firstIn) : '', lastOut: r.lastOut ? hm_(r.lastOut) : '',
        hours: round2_(r.minutes / 60), lateMin: lateMin, earlyMin: earlyMin, flags: flags,
        punches: list.map(x => ({
          id: x['Punch ID'], type: x['Type'], time: hm_(str_(x['Punch Time'])), at: str_(x['Punch Time']), status: x['Status'],
          source: x['Source'], photo: str_(x['Photo']).indexOf('http') === 0 ? x['Photo'] : '', storeId: x['Store ID'], notes: x['Notes'],
          review: /REVIEW/.test(String(x['Notes']))
        }))
      };
    });
    const storeHours = Object.keys(storeMin).map(s => ({ storeId: s, hours: round2_(storeMin[s] / 60) }));
    return { id: id, name: p.name, storeId: p.storeId, role: p.role || '', totalHours: round2_(total / 60), lateDays: late, earlyDays: early, openDays: open, reviewPunches: review, storeHours: storeHours, days: dayRows };
  }).sort((a, b) => (a.storeId + a.name).localeCompare(b.storeId + b.name));
  return { weekStart: ws, weekEnd: days[5], week: weekNum_(ws), storeId: storeId || 'ALL', rows: rows };
}

/** IN/OUT pairing for one day's active punches (sorted). Hours count to the store of the IN. */
function pair_(list) {
  let minutes = 0, openIn = null, openStore = '', firstIn = '', lastOut = '', orphans = 0;
  const byStore = {};
  list.forEach(p => {
    const t = str_(p['Punch Time']);
    if (p['Type'] === 'IN') {
      if (!firstIn) firstIn = t;
      if (openIn) orphans++;
      openIn = t; openStore = p['Store ID'];
    } else {
      if (openIn) {
        const m = (asTime_(t) - asTime_(openIn)) / 60000;
        minutes += m; byStore[openStore] = (byStore[openStore] || 0) + m; openIn = null;
      } else orphans++;
      lastOut = t;
    }
  });
  return { minutes: Math.max(0, minutes), firstIn: firstIn, lastOut: lastOut, open: !!openIn, orphans: orphans, byStore: byStore };
}

function editPunchCore_(who, punchId, time, type, reason, fixId) {
  const rows = readTable_(IDS.CLOCK, 'Punches');
  const i = rows.findIndex(p => p['Punch ID'] === punchId);
  if (i === -1) throw new Error('Unknown punch.');
  if (rows[i]['Status'] !== 'Active') throw new Error('That punch was removed.');
  clearReview_(who, rows, i);
  if (time) {
    const t = checkStamp_(time);
    if (asTime_(t) > Date.now() + 60000) throw new Error('The time cannot be in the future.');
    const old = str_(rows[i]['Punch Time']);
    if (old !== t) { setCell_(IDS.CLOCK, 'Punches', i, 'Punch Time', t); editLog_(who, 'Edit', punchId, 'Punch Time', old, t, reason, fixId); }
  }
  if (type && type !== rows[i]['Type']) {
    oneOf_(type, ['IN', 'OUT']);
    setCell_(IDS.CLOCK, 'Punches', i, 'Type', type);
    editLog_(who, 'Edit', punchId, 'Type', rows[i]['Type'], type, reason, fixId);
  }
  return rows[i]['Employee ID'];
}
/** "Time is OK": Juliana checked an offline punch flagged REVIEW. */
function okPunchCore_(who, punchId) {
  const rows = readTable_(IDS.CLOCK, 'Punches');
  const i = rows.findIndex(p => p['Punch ID'] === punchId);
  if (i === -1) throw new Error('Unknown punch.');
  if (!clearReview_(who, rows, i)) throw new Error('That punch was not flagged.');
  editLog_(who, 'Checked', punchId, 'Notes', 'REVIEW', 'Checked', 'Offline time checked', '');
}
function clearReview_(who, rows, i) {
  const n = String(rows[i]['Notes'] || '');
  if (!/REVIEW:/.test(n)) return false;
  setCell_(IDS.CLOCK, 'Punches', i, 'Notes', n.replace(/REVIEW:/g, 'Checked by ' + who.name + ' ' + today_() + ':'));
  return true;
}
function closeRefused_(who, row) {
  const rows = readTable_(IDS.CLOCK, 'Refused');
  if (!rows[row] || rows[row]['Status'] !== 'Open') throw new Error('Already handled.');
  setCell_(IDS.CLOCK, 'Refused', row, 'Status', 'Handled');
  setCell_(IDS.CLOCK, 'Refused', row, 'Handled By', who.name + ' ' + now_());
}
function voidPunchCore_(who, punchId, reason, fixId) {
  const rows = readTable_(IDS.CLOCK, 'Punches');
  const i = rows.findIndex(p => p['Punch ID'] === punchId);
  if (i === -1) throw new Error('Unknown punch.');
  if (rows[i]['Status'] !== 'Active') throw new Error('Already removed.');
  setCell_(IDS.CLOCK, 'Punches', i, 'Status', 'Voided');
  editLog_(who, 'Void', punchId, 'Status', 'Active', 'Voided', reason, fixId);
  return rows[i]['Employee ID'];
}
function addPunchCore_(who, employeeId, storeId, type, time, reason, fixId) {
  oneOf_(type, ['IN', 'OUT']);
  const t = checkStamp_(time);
  if (asTime_(t) > Date.now() + 60000) throw new Error('The time cannot be in the future.');
  const p = clockPeople_(true).find(x => x.id === employeeId);
  if (!p) throw new Error('Unknown employee.');
  const id = nextPunchId_();
  appendRow_(IDS.CLOCK, 'Punches', {
    'Punch ID': id, 'Employee ID': employeeId, 'Employee Name': p.name, 'Store ID': storeId || p.storeId, 'Type': type,
    'Punch Time': t, 'Recorded At': stamp_(new Date()), 'Photo': '', 'Device ID': '', 'Source': 'Office (' + who.name + ')',
    'Status': 'Active', 'Fix Request ID': fixId || '', 'Notes': text_(reason, 300), 'Client ID': ''
  });
  editLog_(who, 'Add', id, 'Punch', '', type + ' ' + t, reason, fixId);
  return id;
}

function editLog_(who, action, punchId, field, oldV, newV, reason, fixId) {
  appendRow_(IDS.CLOCK, 'Edit Log', {
    'Log ID': nextId_(IDS.CLOCK, 'Edit Log', 'Log ID', 'LOG-', 6), 'Timestamp': now_(), 'By': who.name, 'Action': action,
    'Punch ID': punchId, 'Field': field, 'Old Value': String(oldV), 'New Value': String(newV), 'Reason': text_(reason, 500), 'Fix Request ID': fixId || ''
  });
}

/* ============================ one-time setup ============================ */

/**
 * Run once from the editor (Andre), after pasting this version:
 *  - creates SYSTEM · Clock PINs (restricted, next to the Time Clock sheet), with a PINs tab;
 *  - adds the Visits tab and the Client ID column to SYSTEM · Time Clock;
 *  - creates the shared setup code;
 *  - asks for permission to send the locked-PIN email.
 * Safe to run again: it only adds what's missing. Returns the setup code.
 */
function setupV2() {
  officeWho_();
  secret_(); pinPepper_();
  const props = PropertiesService.getScriptProperties();
  const notes = [];
  let credId = props.getProperty('CRED_ID');
  const credHead = ['Employee ID', 'Employee Name', 'PIN Hash', 'Salt', 'Status', 'Set On', 'Set By', 'Failed Tries', 'Last Fail', 'Locked', 'Locked At', 'Last Used', 'Reset Code', 'Reset Expires'];
  if (!credId) {
    // Stays in Andre's own Drive (not in a shared folder), so only he and this script can open it.
    const ss = SpreadsheetApp.create('SYSTEM · Clock PINs (restricted)');
    const sh = ss.getSheets()[0].setName('PINs');
    sh.getRange(1, 1, 1, credHead.length).setValues([credHead]);
    sh.setFrozenRows(1);
    credId = ss.getId();
    props.setProperty('CRED_ID', credId);
  }
  const cred = SpreadsheetApp.openById(credId);
  const ps = cred.getSheetByName('PINs');
  const ch = ps.getRange(1, 1, 1, ps.getLastColumn()).getValues()[0].map(String);
  credHead.forEach(k => { if (ch.indexOf(k) === -1) { ps.getRange(1, ch.length + 1).setValue(k); ch.push(k); } });
  plainText_(ps, ch, credHead);
  try { cred.setSpreadsheetTimeZone(TZ); } catch (e) {}
  try {
    const f = DriveApp.getFileById(credId);
    const extra = f.getEditors().concat(f.getViewers()).map(u => u.getEmail()).filter(e => e && e.toLowerCase() !== 'andre@primeoneinc.com');
    if (extra.length) notes.push('WARNING: the PINs sheet is shared with ' + extra.join(', ') + '. Remove them.');
  } catch (e) {}

  const clock = SpreadsheetApp.openById(IDS.CLOCK);
  try { if (clock.getSpreadsheetTimeZone() !== TZ) { clock.setSpreadsheetTimeZone(TZ); notes.push('Time Clock sheet time zone set to ' + TZ + '.'); } } catch (e) {}
  try { const core = SpreadsheetApp.openById(IDS.CORE); if (core.getSpreadsheetTimeZone() !== TZ) notes.push('NOTE: Core Data time zone is ' + core.getSpreadsheetTimeZone() + ', not ' + TZ + '.'); } catch (e) {}
  const tab = (name, head) => {
    let sh = clock.getSheetByName(name);
    if (!sh) { sh = clock.insertSheet(name); sh.getRange(1, 1, 1, head.length).setValues([head]); sh.setFrozenRows(1); }
    return sh;
  };
  tab('Visits', ['Date', 'Employee ID', 'Employee Name', 'Home Store', 'Store ID', 'Device ID', 'First Punch']);
  tab('Refused', ['Refused At', 'Made At', 'Store ID', 'Device ID', 'Employee ID', 'Employee Name', 'Cell (last 4)', 'Showed As', 'Reason', 'Status', 'Handled By']);
  const pu = clock.getSheetByName('Punches');
  const h = pu.getRange(1, 1, 1, pu.getLastColumn()).getValues()[0].map(String);
  if (h.indexOf('Client ID') === -1) { pu.getRange(1, h.length + 1).setValue('Client ID'); h.push('Client ID'); }
  plainText_(pu, h, ['Client ID']);
  const dv = clock.getSheetByName('Devices');
  plainText_(dv, dv.getRange(1, 1, 1, dv.getLastColumn()).getValues()[0].map(String), ['Setup Code', 'Token Hash']);
  childFolder_(DriveApp.getFolderById(IDS.MEDIA_LIBRARY), 'Time Clock Photos');
  props.setProperty('STATE_READY', '');   // rebuild "last punch" with stores
  rebuildState_();
  MailApp.getRemainingDailyQuota();       // asks for the email permission now, not at the first lockout
  return ['Ready. PIN setup code: ' + setupCode_() + '.', 'PINs sheet: ' + cred.getUrl()].concat(notes).join('\n');
}
/** Columns that must stay exactly as typed (codes, hashes, IDs): plain text, so Sheets never turns them into numbers or dates. */
function plainText_(sh, headers, names) {
  names.forEach(n => { const c = headers.indexOf(n) + 1; if (c) sh.getRange(1, c, sh.getMaxRows ? sh.getMaxRows() : 1000, 1).setNumberFormat('@'); });
}

/* ============================ people, stores, schedules ============================ */

function activeStores_() {
  return coreTable_('Stores').filter(s => s['Status'] === 'Active').map(s => ({ id: s['Store ID'], name: s['Store Name'] }))
    .sort((a, b) => a.name.localeCompare(b.name));
}

/** Active people with their current store. includeAll=false → only people who clock in. */
function clockPeople_(includeAll) {
  const cur = {};
  coreTable_('Assignments').forEach(a => { if (a['Employee ID'] && !asDay_(a['End Date'])) cur[a['Employee ID']] = a; });
  return coreTable_('Employees').filter(e => e['Status'] === 'Active').map(e => {
    const a = cur[e['Employee ID']] || {};
    const clocks = String(e['Clocks In']).toUpperCase() === 'TRUE' || e['Clocks In'] === true;
    return {
      id: e['Employee ID'], name: e['Display Name'], first: e['Preferred Name'] || e['First Name'], phone: e['Phone'],
      storeId: a['Store ID'] || '', role: a['Role'] || '', clocksIn: clocks, schedText: e['Schedule (current)'] || ''
    };
  }).filter(p => includeAll || p.clocksIn);
}

/** Schedules tab first (per-day times), else the Employees "Schedule (current)" text. */
function schedules_() {
  const out = {};
  const t = today_();
  coreTable_('Schedules').forEach(s => {
    if (asDay_(s['Effective Start']) && asDay_(s['Effective Start']) > t) return;
    if (asDay_(s['Effective End']) && asDay_(s['Effective End']) < t) return;
    out[s['Employee ID']] = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'].map(d => parseRange_(str_(s[d])));
  });
  return out;
}
function schedFor_(sched, id, p, dayIndex) {
  if (sched[id]) return sched[id][dayIndex];
  return parseScheduleText_(p.schedText || '')[dayIndex];
}

/** "Mon–Fri | 10–8; Sat | 8–6" → 6 day slots (Mon..Sat) of {start,end,label} minutes. */
function parseScheduleText_(txt) {
  const out = [null, null, null, null, null, null];
  const names = ['mon', 'tue', 'wed', 'thu', 'fri', 'sat'];
  String(txt || '').split(';').forEach(part => {
    const bits = part.split('|');
    if (bits.length !== 2) return;
    const range = parseRange_(bits[1]);
    if (!range) return;
    const dd = bits[0].toLowerCase().replace(/\s/g, '').split(/[–—-]/);
    const a = names.indexOf(dd[0].slice(0, 3)), b = dd[1] ? names.indexOf(dd[1].slice(0, 3)) : a;
    if (a === -1 || b === -1) return;
    for (let i = a; i <= b; i++) out[i] = range;
  });
  return out;
}
function parseRange_(s) {
  s = String(s || '').trim();
  if (!s || /off/i.test(s)) return null;
  const m = s.match(/^(\d{1,2})(?::(\d{2}))?\s*(am|pm)?\s*[–—-]\s*(\d{1,2})(?::(\d{2}))?\s*(am|pm)?$/i);
  if (!m) return null;
  const conv = (h, mi, ap, isEnd, startH) => {
    h = +h; mi = +(mi || 0);
    if (ap) { ap = ap.toLowerCase(); if (ap === 'pm' && h < 12) h += 12; if (ap === 'am' && h === 12) h = 0; }
    else if (!isEnd && h <= 5) h += 12;             // "4–7": a start at 4 means 4 PM
    return h * 60 + mi;                              // an end before the start is moved to PM below
  };
  const start = conv(m[1], m[2], m[3] || (m[6] && +m[1] <= +m[4] ? m[6] : ''), false, 0);
  let end = conv(m[4], m[5], m[6], true, Math.floor(start / 60));
  if (end <= start) end += 12 * 60;
  return { start: start, end: end, label: hmMin_(start) + '–' + hmMin_(end) };
}

/* ============================ punch state ============================ */

function storeName_(id) {
  if (!id) return '';
  const s = activeStores_().find(x => x.id === id);
  return s ? s.name : id;
}

/** The person's state at refMs (default now): in or out, and since when. */
function stateOf_(last, refMs) {
  const ref = refMs || Date.now();
  if (!last) return { status: 'out', since: '', sinceLabel: '', openSince: '', storeId: '' };
  const t = asTime_(last['Punch Time']);
  const label = Utilities.formatDate(new Date(t), TZ, asDay_(last['Punch Time']) === today_() ? 'h:mm a' : 'EEE M/d h:mm a');
  const storeId = last['Store ID'] || '';
  if (last['Type'] === 'IN') {
    if (ref - t > OPEN_SHIFT_HOURS * 3600000) return { status: 'out', since: '', sinceLabel: '', openSince: label, openDate: asDay_(last['Punch Time']), storeId: '' };
    return { status: 'in', since: str_(last['Punch Time']), sinceLabel: label, openSince: '', storeId: storeId };
  }
  return { status: 'out', since: str_(last['Punch Time']), sinceLabel: label, openSince: '', storeId: storeId };
}

/* ============================ tokens ============================ */

function secret_() {
  const props = PropertiesService.getScriptProperties();
  let s = props.getProperty('TOKEN_SECRET');
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); props.setProperty('TOKEN_SECRET', s); }
  return s;
}
function sign_(text) { return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(text, secret_())); }
function makeToken_(kind, value, minutes) {
  const body = kind + '~' + value + '~' + (Date.now() + minutes * 60000);
  return body + '~' + sign_(body);
}
function readToken_(kind, token) {
  const p = String(token || '').split('~');
  if (p.length !== 4 || p[0] !== kind) throw new Error('Please enter your PIN again.');
  if (sign_(p.slice(0, 3).join('~')) !== p[3]) throw new Error('Please enter your PIN again.');
  if (Number(p[2]) < Date.now()) throw new Error(kind === 'O' ? 'Signed out. Please sign in again.' : 'Took too long. Please enter your PIN again.');
  return p[1];
}
function hash_(kind, v) { return sha_(kind + ':' + v + ':' + secret_()); }
function sha_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s)
    .map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}

/* ============================ speed: caches ============================ */

/** Core Data tabs (people, stores, assignments, schedules) cached 5 minutes. PINs are never cached. */
function coreTable_(name) {
  const key = 'core_' + name;
  if (ROWS_[key]) return ROWS_[key];
  const c = CacheService.getScriptCache();
  const hit = c.get(key);
  if (hit) return (ROWS_[key] = JSON.parse(hit));
  const rows = readTable_(IDS.CORE, name).map(r => {
    const o = {}; Object.keys(r).forEach(k => o[k] = isDate_(r[k]) ? stamp_(r[k]) : r[k]); return o;
  });
  try { c.put(key, JSON.stringify(rows), 300); } catch (e) { /* too big to cache: read each time */ }
  return (ROWS_[key] = rows);
}

/**
 * Each person's last punch is kept in script properties so the kiosk never has to read the whole
 * Punches tab. It is rebuilt from the sheet whenever the office screen loads or a punch is edited.
 */
let LAST_ = null;
function lastPunches_() {
  if (LAST_) return LAST_;
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('STATE_READY') !== '3') return rebuildState_();
  const all = props.getProperties(), out = {};
  Object.keys(all).forEach(k => { if (k.indexOf('LAST_') === 0) out[k.slice(5)] = JSON.parse(all[k]); });
  return (LAST_ = out);
}
function rebuildState_() {
  const props = PropertiesService.getScriptProperties();
  const out = {}; let max = Number(props.getProperty('PUNCH_SEQ') || 0);
  readTable_(IDS.CLOCK, 'Punches').forEach(p => {
    const n = parseInt(String(p['Punch ID']).replace('PCH-', ''), 10);
    if (!isNaN(n) && n > max) max = n;
    if (p['Status'] !== 'Active') return;
    const cur = out[p['Employee ID']];
    if (!cur || asTime_(p['Punch Time']) >= asTime_(cur['Punch Time'])) out[p['Employee ID']] = { 'Punch ID': p['Punch ID'], 'Type': p['Type'], 'Punch Time': str_(p['Punch Time']), 'Store ID': p['Store ID'] };
  });
  const old = props.getProperties();
  Object.keys(old).forEach(k => { if (k.indexOf('LAST_') === 0 && !out[k.slice(5)]) props.deleteProperty(k); });
  const set = { PUNCH_SEQ: String(max), STATE_READY: '3' };
  Object.keys(out).forEach(id => set['LAST_' + id] = JSON.stringify(out[id]));
  props.setProperties(set);
  return (LAST_ = out);
}
/** After an office edit or an undo: recompute one person's last punch from the newest rows. */
function rebuildOne_(empId) {
  if (!empId) return;
  let best = null;
  recentRows_(4000).forEach(p => {
    if (p['Employee ID'] !== empId || p['Status'] !== 'Active') return;
    if (!best || asTime_(p['Punch Time']) >= asTime_(best['Punch Time'])) best = p;
  });
  const props = PropertiesService.getScriptProperties();
  if (!best) { props.deleteProperty('LAST_' + empId); if (LAST_) delete LAST_[empId]; return; }
  setLast_(empId, best['Punch ID'], best['Type'], str_(best['Punch Time']), best['Store ID']);
}
function setLast_(empId, punchId, type, time, storeId) {
  const v = { 'Punch ID': punchId, 'Type': type, 'Punch Time': time, 'Store ID': storeId || '' };
  PropertiesService.getScriptProperties().setProperty('LAST_' + empId, JSON.stringify(v));
  if (LAST_) LAST_[empId] = v;
}
function nextPunchId_() {
  lastPunches_();
  const props = PropertiesService.getScriptProperties();
  const n = Number(props.getProperty('PUNCH_SEQ') || 0) + 1;
  props.setProperty('PUNCH_SEQ', String(n));
  return 'PCH-' + String(n).padStart(6, '0');
}
/** Only the newest rows of the Punches tab (for a person's recent punches). */
function recentRows_(n) {
  const sh = sheet_(IDS.CLOCK, 'Punches'), last = sh.getLastRow();
  if (last < 2) return [];
  const start = Math.max(2, last - n + 1), h = headers_(IDS.CLOCK, 'Punches');
  return sh.getRange(start, 1, last - start + 1, h.length).getValues().map(r => { const o = {}; h.forEach((k, i) => o[k] = r[i]); return o; });
}

/* ============================ sheet helpers ============================ */

const CACHE_ = {};
function sheet_(fileId, name) {
  const key = fileId + '|' + name;
  if (!CACHE_[key]) {
    const sh = SpreadsheetApp.openById(fileId).getSheetByName(name);
    if (!sh) throw new Error('Missing tab: ' + name);
    CACHE_[key] = sh;
  }
  return CACHE_[key];
}
const HEAD_ = {};
function headers_(fileId, name) {
  const key = fileId + '|' + name;
  if (!HEAD_[key]) { const sh = sheet_(fileId, name); HEAD_[key] = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String); }
  return HEAD_[key];
}
const ROWS_ = {};
/** Rows as objects. Each remembers its real sheet row (__row), so a blank row never shifts writes. */
function readTable_(fileId, name, fresh) {
  const key = fileId + '|' + name;
  if (ROWS_[key] && !fresh) return ROWS_[key];
  const sh = sheet_(fileId, name);
  const last = sh.getLastRow();
  const h = headers_(fileId, name);
  ROWS_[key] = last < 2 ? [] : sh.getRange(2, 1, last - 1, h.length).getValues()
    .map((r, n) => { const o = {}; h.forEach((k, i) => o[k] = r[i]); Object.defineProperty(o, '__row', { value: n + 2 }); return o; })
    .filter(o => o[h[0]] !== '');
  return ROWS_[key];
}
function appendRow_(fileId, name, obj) {
  const h = headers_(fileId, name);
  const sh = sheet_(fileId, name);
  sh.appendRow(h.map(k => (k in obj ? obj[k] : '')));
  const key = fileId + '|' + name;
  if (ROWS_[key]) { const o = {}; h.forEach(k => o[k] = k in obj ? obj[k] : ''); Object.defineProperty(o, '__row', { value: sh.getLastRow() }); ROWS_[key].push(o); }
}
/** dataIndex = position in the readTable_ result; the write goes to that row's real sheet row. */
function setCell_(fileId, name, dataIndex, header, value) {
  const col = headers_(fileId, name).indexOf(header) + 1;
  if (!col) throw new Error('Missing column ' + header + ' in ' + name);
  const rows = readTable_(fileId, name);
  const r = rows[dataIndex];
  if (!r || !r.__row) throw new Error('Row not found in ' + name + '. Try again.');
  sheet_(fileId, name).getRange(r.__row, col).setValue(value);
  r[header] = value;
}
function nextId_(fileId, name, header, prefix, width) {
  let max = 0;
  readTable_(fileId, name).forEach(r => {
    const n = parseInt(String(r[header]).replace(prefix, ''), 10);
    if (!isNaN(n) && n > max) max = n;
  });
  return prefix + String(max + 1).padStart(width, '0');
}
function childFolder_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}
function accessLog_(devId, storeId, action, empId, result, detail) {
  try {
    appendRow_(IDS.CLOCK, 'Access Log', { 'Timestamp': now_(), 'Device ID': devId, 'Store ID': storeId, 'Action': action, 'Employee ID': empId, 'Result': result, 'Detail': detail });
  } catch (e) { /* logging never blocks */ }
}
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) throw new Error('RETRY: Busy. Try again in a few seconds.');
  try { return fn(); } finally { SpreadsheetApp.flush(); lock.releaseLock(); }
}

/* ============================ small helpers ============================ */

function must_(v, label) { if (v === undefined || v === null || String(v).trim() === '') throw new Error(label + ' is required.'); }
function oneOf_(v, list) { if (list.indexOf(v) === -1) throw new Error('Invalid choice: ' + v); return v; }
/** Free text from people: trimmed, and never starts like a spreadsheet formula. */
function text_(v, max) { return String(v === undefined || v === null ? '' : v).trim().replace(/^[=+\-@\t\r]+/, '').slice(0, max); }
/** 6 random digits from a UUID (not Math.random). */
function rand6_() { return String(100000 + parseInt(Utilities.getUuid().replace(/-/g, '').slice(0, 12), 16) % 900000); }
/** Counters that survive across requests (Script Properties). Call bumpWindow_ inside withLock_. */
function windowCount_(key) {
  const v = String(PropertiesService.getScriptProperties().getProperty('W_' + key) || '').split('|');
  return Number(v[1] || 0) > Date.now() ? Number(v[0] || 0) : 0;
}
function bumpWindow_(key, seconds) {
  const props = PropertiesService.getScriptProperties();
  const v = String(props.getProperty('W_' + key) || '').split('|');
  const live = Number(v[1] || 0) > Date.now();
  const n = (live ? Number(v[0] || 0) : 0) + 1;
  props.setProperty('W_' + key, n + '|' + (live ? v[1] : Date.now() + seconds * 1000));
  return n;
}
function digits_(v) { return String(v || '').replace(/\D/g, '').slice(-10); }
function checkDate_(v, label) { if (!/^\d{4}-\d{2}-\d{2}$/.test(String(v))) throw new Error(label + ' must be a date.'); }
/** "2026-10-28 08:02" or "...:ss" → "2026-10-28 08:02:00" */
function checkStamp_(v) {
  const m = String(v || '').trim().replace('T', ' ').match(/^(\d{4}-\d{2}-\d{2}) (\d{2}):(\d{2})(?::(\d{2}))?$/);
  if (!m) throw new Error('Enter a date and time.');
  if (+m[2] > 23 || +m[3] > 59) throw new Error('Invalid time.');
  return m[1] + ' ' + m[2] + ':' + m[3] + ':' + (m[4] || '00');
}
function now_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'); }
function today_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function stamp_(d) { return Utilities.formatDate(d, TZ, 'yyyy-MM-dd HH:mm:ss'); }
function isDate_(v) { return Object.prototype.toString.call(v) === '[object Date]'; }
function str_(v) { return isDate_(v) ? stamp_(v) : String(v === undefined || v === null ? '' : v); }
function asDay_(v) { return isDate_(v) ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd') : String(v || '').trim().slice(0, 10); }
function asTime_(v) {
  if (!v) return 0;
  if (isDate_(v)) return v.getTime();
  const m = String(v).match(/^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2})(?::(\d{2}))?/);
  if (!m) return 0;
  const guess = new Date(Date.UTC(+m[1], +m[2] - 1, +m[3], +m[4], +m[5], +(m[6] || 0)));
  const off = Utilities.formatDate(guess, TZ, 'Z'); // e.g. -0400
  const sign = off[0] === '-' ? 1 : -1;
  const mins = (+off.slice(1, 3)) * 60 + (+off.slice(3, 5));
  return guess.getTime() + sign * mins * 60000;
}
function minsOf_(stampStr) { const m = String(stampStr).match(/ (\d{2}):(\d{2})/); return m ? (+m[1]) * 60 + (+m[2]) : 0; }
function hm_(stampStr) { return hmMin_(minsOf_(stampStr)); }
function hmMin_(mins) {
  const h = Math.floor(mins / 60) % 24, mi = mins % 60;
  return ((h % 12) || 12) + ':' + String(mi).padStart(2, '0') + (h < 12 ? ' AM' : ' PM');
}
function addDays_(iso, n) {
  const p = iso.split('-').map(Number);
  return new Date(Date.UTC(p[0], p[1] - 1, p[2] + n)).toISOString().slice(0, 10);
}
function mondayOf_(iso) {
  checkDate_(iso, 'Week');
  const p = iso.split('-').map(Number);
  const d = new Date(Date.UTC(p[0], p[1] - 1, p[2]));
  const dow = d.getUTCDay();
  d.setUTCDate(d.getUTCDate() - (dow === 0 ? 6 : dow - 1));
  return d.toISOString().slice(0, 10);
}
function weekNum_(mondayIso) {
  const p = mondayIso.split('-').map(Number);
  return Math.round((Date.UTC(p[0], p[1] - 1, p[2]) - W1_MONDAY) / (7 * 86400000)) + 1;
}
function nowMin_() { const t = Utilities.formatDate(new Date(), TZ, 'HH:mm').split(':'); return (+t[0]) * 60 + (+t[1]); }
function dur_(m) { m = Math.round(m); return m >= 60 ? Math.floor(m / 60) + 'h ' + (m % 60) + 'm' : m + ' min'; }
function round2_(n) { return Math.round(n * 100) / 100; }
