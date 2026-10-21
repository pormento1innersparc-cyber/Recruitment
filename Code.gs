/**
 * Code.gs  (complete, single file)
 * Recruitment page + Day 0 confirmation + owner alert for Cynthia M. Martin (Mommy Property Finder).
 * Works on its own. The 30-day email sequence lives in a separate file, Nurture.gs.
 *
 * On every survey submit:
 *  - answers are logged to the "Recruits" tab
 *  - Day 0: branded confirmation email (with the welcome picture) goes out right away
 *  - owner gets a branded alert email
 *
 * Setup: bound script (Sheet > Extensions > Apps Script), HTML file named RecruitPage,
 * Deploy > Web app (Execute as Me, Access Anyone).
 */

/* ============ CONFIG ============ */
const OWNER_EMAIL = 'mommypropertyfinder@gmail.com'; // TODO: change to the live address
const SENDER_NAME = 'Cynthia M. Martin';

const LOGO_ID = '1KdxbH74HfnrhZMXaigXV-m8MzuK2TQ4N';      // brand logo
const HERO_ID = '1st9mDJcHanBFHlxiWI3Z-NnGfrbmsWJi';      // Cynthia's portrait (signature)
const WELCOME_ID = '1r_CkzR2ZASejcXi4uZSAoAh7Cn5m__bB';   // welcome picture (Day 0 email)
// Day 0 CTA target. utm tags let you see email clicks in Google Analytics. (fbclid left out on purpose: it's a per-click Facebook ID)
const AGENT_PORTAL_URL = 'https://innersparcagentportal.vercel.app/?utm_source=recruit_email&utm_medium=email&utm_campaign=day0_welcome';

const NAVY = '#0B1F4D', GOLD = '#C8A24D';
const SERIF = "Georgia,'Times New Roman',serif";
const SANS = "Arial,Helvetica,sans-serif";
const DISCLAIMER = 'Ang income sa real estate ay commission-based at iba-iba. Walang guaranteed na income.';

const RECRUIT_LOG_TAB = 'Recruits';
const RECRUIT_LOG_HEADERS = ['Timestamp', 'Name', 'Email', 'Phone', 'Location', 'CurrentWork',
  'Experience', 'LicenseJourney', 'HoursPerDay', 'StartTiming', 'Why', 'BestTime', 'HeardFrom', 'Question', 'Consent', 'CommissionAck', 'Source'];

/* ============ WEB APP ============ */

function doGet(e) {
  // unsubscribe links only work when Nurture.gs is in the project
  if (e && e.parameter && e.parameter.action === 'unsub' && typeof handleUnsubscribe_ === 'function') return handleUnsubscribe_(e);
  return HtmlService.createHtmlOutputFromFile('RecruitPage')
    .setTitle('Become a Real Estate Agent | Cynthia M. Martin')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// For the Vercel version later (safe to leave in)
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.type === 'recruit') submitRecruitSurvey(body);
  } catch (err) { console.error(err); }
  return ContentService.createTextOutput('ok');
}

/* ============ SURVEY SUBMIT ============ */

function submitRecruitSurvey(d) {
  try {
    d = d || {};
    const email = String(d.email || '').trim().toLowerCase();
    if (!isValidEmail_(email)) return { ok: false, error: 'Please enter a valid email.' };
    if (!d.consent) return { ok: false, error: 'Consent is required.' };
    if (!d.ack) return { ok: false, error: 'Please confirm that you understand this is commission-based.' };
    if (d.website) return { ok: true }; // honeypot hit, pretend success

    const now = new Date();
    // If Nurture.gs is in the project it decides whether this recruit starts the 30-day sequence.
    const nurture = (typeof nurtureEnrollment_ === 'function') ? nurtureEnrollment_(email, now) : {};

    logRecruit_(Object.assign({
      Timestamp: now, Name: d.name, Email: email, Phone: d.phone, Location: d.location,
      CurrentWork: d.work, Experience: d.experience, LicenseJourney: d.license, HoursPerDay: d.hours, StartTiming: d.start, BestTime: d.bestTime, HeardFrom: d.heard,
      Why: d.why, Question: d.question, Consent: 'Yes', CommissionAck: 'Yes', Source: d.source || 'recruitment-page'
    }, nurture));

    let confirmationSent = false;
    try { confirmationSent = sendRecruitConfirmation_(d, email) === true; } catch (e1) { console.error('Confirmation failed: ' + e1); }
    try { sendOwnerAlert_(d, email, confirmationSent); } catch (e2) { console.error('Owner alert failed: ' + e2); }

    return { ok: true };
  } catch (err) {
    console.error(err);
    return { ok: false, error: 'Something went wrong. Please try again.' };
  }
}

function isValidEmail_(s) { return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(s || '')); }

/* ============ SHEET HELPERS (header-name based, column order never matters) ============ */

function recruitSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(RECRUIT_LOG_TAB);
  if (!sh) {
    sh = ss.insertSheet(RECRUIT_LOG_TAB);
  }
  ensureRecruitHeaders_(sh);
  return sh;
}

function recruitCols_(sh) {
  const map = {};
  sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].forEach(function (h, i) { map[String(h).trim()] = i; });
  return map;
}

function logRecruit_(obj) {
  const sh = recruitSheet_();
  const cols = recruitCols_(sh);
  const row = new Array(sh.getLastColumn()).fill('');
  Object.keys(obj).forEach(function (k) { if (cols[k] !== undefined) row[cols[k]] = obj[k]; });
  sh.appendRow(row);
  // Formatting must never stop a valid survey from being saved.
  try { formatRecruitSheet_(sh); } catch (formatErr) { console.warn('Sheet formatting skipped: ' + formatErr); }
}

/* ============ RECRUITS SHEET CLEANUP ============ */

// Run setupRecruitSheet() once from the Apps Script editor to clean the existing sheet.
// It also runs automatically before every new survey response.
function setupRecruitSheet() {
  const sh = recruitSheet_();
  formatRecruitSheet_(sh);
  return 'Recruits sheet is organized and formatted.';
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Recruitment Tools')
    .addItem('Organize Recruits Sheet', 'setupRecruitSheet')
    .addToUi();
}

function ensureRecruitHeaders_(sh) {
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const firstRow = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  let headers = firstRow.map(function (h) { return String(h).trim(); });
  if (!headers.some(function (h) { return h; })) {
    sh.getRange(1, 1, 1, RECRUIT_LOG_HEADERS.length).setValues([RECRUIT_LOG_HEADERS]);
    return;
  }

  // Keep old submissions working while correcting the old header name in place.
  const oldHours = headers.indexOf('HoursPerWeek');
  const newHours = headers.indexOf('HoursPerDay');
  if (oldHours !== -1 && newHours === -1) {
    sh.getRange(1, oldHours + 1).setValue('HoursPerDay');
    headers[oldHours] = 'HoursPerDay';
  }

  const oldLicense = headers.indexOf('License');
  const oldPrLicense = headers.indexOf('PRCLicense');
  const newLicense = headers.indexOf('LicenseJourney');
  if (newLicense === -1 && oldLicense !== -1) {
    sh.getRange(1, oldLicense + 1).setValue('LicenseJourney');
    headers[oldLicense] = 'LicenseJourney';
  } else if (newLicense === -1 && oldPrLicense !== -1) {
    sh.getRange(1, oldPrLicense + 1).setValue('LicenseJourney');
    headers[oldPrLicense] = 'LicenseJourney';
  }

  RECRUIT_LOG_HEADERS.forEach(function (h) {
    if (headers.indexOf(h) === -1) {
      sh.getRange(1, sh.getLastColumn() + 1).setValue(h);
      headers.push(h);
    }
  });
}

function organizeRecruitSheet_(sh) {
  const values = sh.getDataRange().getValues();
  const existingHeaders = values.length && values[0].some(function (h) { return String(h).trim(); })
    ? values[0].map(function (h) { return String(h).trim(); })
    : [];

  // Migrate the old column name without deleting any existing responses.
  const aliases = {
    HoursPerWeek: 'HoursPerDay',
    'Hours per day': 'HoursPerDay',
    'Start timing': 'StartTiming',
    'Best time': 'BestTime',
    'Heard from': 'HeardFrom',
    'Commission acknowledgement': 'CommissionAck',
    License: 'LicenseJourney',
    PRCLicense: 'LicenseJourney'
  };
  const normalized = existingHeaders.map(function (h) { return aliases[h] || h; });
  const extras = normalized.filter(function (h, i) {
    return h && RECRUIT_LOG_HEADERS.indexOf(h) === -1 && normalized.indexOf(h) === i;
  });
  const headers = RECRUIT_LOG_HEADERS.concat(extras);
  const index = {};
  normalized.forEach(function (h, i) { if (h && index[h] === undefined) index[h] = i; });
  const rows = values.slice(1).map(function (row) {
    return headers.map(function (h) { return index[h] === undefined ? '' : row[index[h]]; });
  });
  const output = [headers].concat(rows);
  const oldRows = Math.max(sh.getLastRow(), 1);
  const oldCols = Math.max(sh.getLastColumn(), headers.length, 1);
  sh.getRange(1, 1, oldRows, oldCols).clearContent();
  sh.getRange(1, 1, output.length, headers.length).setValues(output);
}

function formatRecruitSheet_(sh) {
  const lastRow = Math.max(sh.getLastRow(), 1);
  const lastCol = Math.max(sh.getLastColumn(), RECRUIT_LOG_HEADERS.length);
  const header = sh.getRange(1, 1, 1, lastCol);
  header.setBackground(NAVY).setFontColor('#FFFFFF').setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
  header.setRowHeight(34);

  if (lastRow > 1) {
    sh.getRange(2, 1, lastRow - 1, lastCol).setVerticalAlignment('top').setWrap(true);
  }

  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h).trim(); });
  const timestampCol = headers.indexOf('Timestamp') + 1;
  if (timestampCol) sh.getRange(2, timestampCol, Math.max(lastRow - 1, 1), 1).setNumberFormat('mmm d, yyyy h:mm:ss');

  const widths = {
    Timestamp: 155, Name: 165, Email: 245, Phone: 140, Location: 185,
    CurrentWork: 220, Experience: 180, LicenseJourney: 245, HoursPerDay: 170,
    StartTiming: 160, Why: 265, BestTime: 145, HeardFrom: 165,
    Question: 300, Consent: 100, CommissionAck: 150, Source: 150
  };
  headers.forEach(function (h, i) {
    sh.setColumnWidth(i + 1, widths[h] || 160);
  });

  const filter = sh.getFilter();
  if (filter) filter.remove();
  sh.getRange(1, 1, lastRow, lastCol).createFilter();
  sh.setFrozenRows(1);
  sh.setFrozenColumns(2);
}


/* ============ EMAIL DESIGN (shared by every email) ============ */

function esc_(s) {
  return String(s === undefined || s === null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/\n/g, '<br>');
}

const blobCache_ = {};
function driveBlob_(id, name) {
  try {
    if (!blobCache_[id]) blobCache_[id] = DriveApp.getFileById(id).getBlob();
    return blobCache_[id].copyBlob().setName(name);
  } catch (e) { console.warn('Image ' + name + ' not loaded: ' + e); return null; }
}

// logo + signature photo for every email. Optional extra "visual" picture by Drive ID.
function brandImages_(visualId) {
  const images = {};
  const logo = driveBlob_(LOGO_ID, 'logo'); if (logo) images.logo = logo;
  const photo = driveBlob_(HERO_ID, 'photo'); if (photo) images.photo = photo;
  if (visualId) { const v = driveBlob_(visualId, 'visual'); if (v) images.visual = v; }
  return images;
}

// opts: { visual: 'banner' | 'portrait' | undefined, signature: bool, footer: html }
function emailShell_(title, bodyHtml, images, opts) {
  opts = opts || {};
  const header = images.logo
    ? '<img src="cid:logo" alt="Mommy Property Finder" height="52" style="height:52px;display:inline-block">'
    : '<span style="font:bold 22px ' + SERIF + ';color:' + GOLD + '">Mommy Property Finder</span>';

  let visual = '';
  if (images.visual && opts.visual === 'banner') {
    visual = '<tr><td style="padding:0"><img src="cid:visual" alt="" width="580" style="width:100%;max-width:580px;height:auto;display:block"></td></tr>';
  } else if (images.visual && opts.visual === 'portrait') {
    visual = '<tr><td align="center" style="padding:26px 30px 0"><img src="cid:visual" alt="Cynthia M. Martin" width="260" style="width:260px;max-width:100%;height:auto;border-radius:14px;border:3px solid ' + GOLD + ';display:block"></td></tr>';
  }

  const signature = opts.signature
    ? '<tr><td style="padding:14px 30px 0"><div style="height:1px;background:#eadfc4"></div></td></tr>' +
      '<tr><td align="center" style="padding:24px 30px 30px;text-align:center">' +
      (images.photo
        ? '<table role="presentation" align="center" cellpadding="0" cellspacing="0" style="margin:0 auto 12px"><tr><td align="center">' +
          '<img src="cid:photo" alt="Cynthia M. Martin" width="88" height="88" style="width:88px;height:88px;border-radius:50%;border:3px solid ' + GOLD + ';object-fit:cover;display:block"></td></tr></table>'
        : '') +
      '<div style="font:bold 19px ' + SERIF + ';color:' + NAVY + ';text-align:center">Cynthia M. Martin</div>' +
      '<div style="font:14px/1.5 ' + SANS + ';color:#2a3042;text-align:center;margin-top:2px">Mommy Property Finder</div>' +
      '<div style="font:13px/1.5 ' + SANS + ';color:#7b8397;text-align:center">InnerSparc Realty Corporation</div>' +
      '</td></tr>'
    : '<tr><td style="padding-bottom:24px"></td></tr>';

  return '<div style="background:#eef0f5;padding:26px 12px">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:580px;margin:0 auto;background:#ffffff;border-radius:12px;overflow:hidden">' +
    '<tr><td style="background:' + NAVY + ';padding:24px 20px;text-align:center;border-bottom:4px solid ' + GOLD + '">' + header +
    '<div style="font:11px ' + SANS + ';letter-spacing:2px;text-transform:uppercase;color:#9fadd0;margin-top:8px">Agent Recruitment</div></td></tr>' +
    visual +
    '<tr><td style="padding:28px 30px 6px">' +
    '<h1 style="margin:0 0 18px;font:bold 26px/1.25 ' + SERIF + ';color:' + NAVY + '">' + title + '</h1>' +
    bodyHtml + '</td></tr>' +
    signature +
    '<tr><td style="background:' + NAVY + ';padding:18px 24px;text-align:center;font:12px/1.6 ' + SANS + ';color:#8d9bc0">' +
    (opts.footer || '') + '<br>' + DISCLAIMER + '</td></tr>' +
    '</table></div>';
}

function para_(text) {
  return '<p style="margin:0 0 18px;font:16px/1.7 ' + SANS + ';color:#2a3042">' + text + '</p>';
}

// Big, centered gold CTA button (email-safe, tappable on phones)
function ctaButton_(url, label) {
  return '<table role="presentation" align="center" cellpadding="0" cellspacing="0" style="margin:6px auto 26px"><tr>' +
    '<td align="center" style="background:' + GOLD + ';border-radius:10px;box-shadow:0 4px 12px rgba(200,162,77,.45)">' +
    '<a href="' + String(url).replace(/&/g, '&amp;') + '" style="display:block;padding:17px 34px;font-family:' + SANS + ';font-size:17px;font-weight:bold;color:' + NAVY +
    ';text-decoration:none;letter-spacing:.2px">' + label + ' &rarr;</a></td></tr></table>';
}

function stepRow_(n, heading, text) {
  return '<tr><td valign="middle" style="padding:0 16px 20px 0">' +
    '<table role="presentation" cellpadding="0" cellspacing="0"><tr>' +
    '<td width="48" height="48" align="center" valign="middle" style="width:48px;height:48px;border-radius:50%;background:' + NAVY +
    ';color:' + GOLD + ';text-align:center;font-family:' + SANS + ';font-size:22px;font-weight:bold;line-height:48px">' + n + '</td></tr></table></td>' +
    '<td valign="middle" style="padding-bottom:20px;font:16px/1.6 ' + SANS + ';color:#2a3042">' +
    '<strong style="color:' + NAVY + ';font-size:17px">' + heading + '</strong><br>' + text + '</td></tr>';
}

/* ============ DAY 0: CONFIRMATION (with welcome picture) ============ */

function sendRecruitConfirmation_(d, email) {
  const first = String(d.name || '').trim().split(/\s+/)[0] || 'there';
  const images = brandImages_(WELCOME_ID);
  const replySubject = encodeURIComponent('Question about the real estate opportunity');
  const replyBody = encodeURIComponent('Hi Cynthia,\n\nI have a question about the real estate opportunity.\n\nThank you.');
  const replyUrl = 'mailto:' + OWNER_EMAIL + '?subject=' + replySubject + '&body=' + replyBody;

  const body =
    para_('Hi ' + esc_(first) + '!') +
    '<div style="font:bold 18px/1.45 ' + SANS + ';color:' + NAVY + ';margin:4px 0 20px">💭 MINSAN, HINDI NAMAN DAHIL WALANG OPPORTUNITY. BAKA HINDI MO PA LANG NAI-EXPLORE ANG MGA ITO.</div>' +
    para_('May mga pagkakataong naiisip natin…') +
    '<div style="padding:0 0 4px 18px;border-left:3px solid ' + GOLD + ';margin:4px 0 22px;font:italic 16px/1.9 ' + SANS + ';color:#2a3042">“May iba pa kaya akong puwedeng gawin?”<br><br>“Paano kaya kung magkaroon ako ng additional source of income?”<br><br>“What if I try something new?”</div>' +
    para_('At minsan, kahit interesado tayo, may pumipigil pa rin sa atin. Baka iniisip mong kailangan mo ng experience. Baka feeling mo hindi ka magaling sa sales. O baka hindi mo lang talaga alam kung paano magsisimula.') +
    para_('And that’s completely understandable. Dahil kapag papasok ka sa isang bagong industry, natural lang na magkaroon ng questions.') +
    '<div style="background:#fbf6e8;border-left:4px solid ' + GOLD + ';padding:16px 18px;margin:4px 0 22px;font:16px/1.7 ' + SANS + ';color:#2a3042"><strong style="color:' + NAVY + '">Pero alam mo kung ano ang magandang simula?</strong><br>Ang bigyan ang sarili mo ng pagkakataong malaman kung ano ang puwede mong matutunan.</div>' +
    para_('<strong style="color:' + NAVY + '">🏡 Have you considered a career in real estate?</strong>') +
    para_('Sa Inner SPARC Realty Corporation, may opportunity kang maging Real Estate Agent at i-explore ang world of property sales. Matututo kang makipag-connect sa prospective buyers, mag-present ng property options, at tumulong sa mga taong naghahanap ng bahay na babagay sa kanilang needs at budget.') +
    para_('Habang nagkakaroon ka ng experience, maaari mo ring ma-develop ang communication skills mo, sales confidence, at professional network.') +
    para_('Siyempre, hindi overnight ang success. Kailangan pa rin ng effort, consistency, at patience.') +
    para_('Pero paano mo malalaman kung bagay sa’yo ang isang opportunity kung hindi mo man lang susubukang alamin?') +
    para_('<strong style="color:' + NAVY + '">You don’t have to figure everything out today. You just have to take the first step. 🚀</strong>') +
    para_('Kung curious kang malaman kung paano maging part ng Inner SPARC, puwede kang magsimula rito:') +
    ctaButton_(AGENT_PORTAL_URL, 'Start Your Real Estate Journey') +
    para_('<strong style="color:' + NAVY + '">Important:</strong> Commission-based ang income sa real estate. Walang fixed salary at walang guaranteed income. Explore the opportunity, ask questions, at tingnan kung aligned ito sa goals mo.') +
    para_('Malay mo, ito ang bagong path na gusto mong i-pursue.') +
    para_('— Inner SPARC Realty Corporation') +
    para_('May tanong ka? Click the button below para mabilis kang makapag-reply sa akin.') +
    ctaButton_(replyUrl, 'Reply to Cynthia');

  const html = emailShell_('Nakuha ko na ang survey mo, ' + esc_(first) + '. Baka ito ang next step para sa iyo.', body, images,
    { visual: 'banner', signature: true, footer: 'Natanggap mo ito dahil sinagutan mo ang agent survey. Pwede kang magtanong bago mag-register.' });

  MailApp.sendEmail({
    to: email,
    subject: first + ', baka ito ang next step para sa iyo',
    body: 'Hi ' + first + '! Salamat sa pagsagot ng agent survey.\n\n' +
      'Baka may opportunity na hindi mo pa na-e-explore. Kung curious kang malaman kung puwedeng bagay sa iyo ang real estate, tingnan ang next step dito: ' + AGENT_PORTAL_URL + '\n\n' +
      'Commission-based ang income sa real estate. Walang fixed salary at walang guaranteed income. Magtanong muna at magdesisyon kapag ready ka na.\n\n' +
      'May tanong ka? Mag-reply sa email na ito o gamitin ang Reply to Cynthia button sa email.\n\n' +
      'Salamat,\nCynthia M. Martin',
    htmlBody: html, inlineImages: images, name: SENDER_NAME, replyTo: OWNER_EMAIL
  });
  return true;
}

/* ============ OWNER ALERT ============ */

function quote_(label, text) {
  return '<div style="font:bold 11px ' + SANS + ';letter-spacing:1.5px;text-transform:uppercase;color:' + GOLD + ';margin:4px 0 6px">' + label + '</div>' +
    '<div style="background:#fbf6e8;border-left:4px solid ' + GOLD + ';padding:12px 16px;margin-bottom:16px;font:15px/1.6 ' + SANS + ';color:#2a3042">' + esc_(text) + '</div>';
}

function alertButton_(href, label, primary) {
  return '<td style="padding-right:10px"><table role="presentation" cellpadding="0" cellspacing="0"><tr>' +
    '<td style="background:' + (primary ? GOLD : '#ffffff') + ';border:2px solid ' + (primary ? GOLD : NAVY) + ';border-radius:8px">' +
    '<a href="' + String(href).replace(/&/g, '&amp;') + '" style="display:inline-block;padding:12px 22px;font-family:' + SANS + ';font-size:15px;font-weight:bold;color:' + NAVY +
    ';text-decoration:none">' + label + '</a></td></tr></table></td>';
}

// Alert to the owner. confirmationSent tells her if the recruit's welcome email actually went out.
function sendOwnerAlert_(d, email, confirmationSent) {
  const images = brandImages_();
  const name = String(d.name || '').trim() || email;
  const first = name.split(/\s+/)[0];
  const priority = /bulacan|cavite/i.test(String(d.location || ''));
  const phoneDigits = String(d.phone || '').replace(/[^\d+]/g, '');
  const sheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();

  const pill = function (t) { return ' <span style="display:inline-block;background:' + NAVY + ';color:' + GOLD + ';font:bold 10px ' + SANS + ';letter-spacing:1.5px;padding:4px 10px;border-radius:20px;vertical-align:middle">' + t + '</span>'; };
  const badge = (priority ? pill('PRIORITY AREA') : '') + (/^Within 1 to 2 days/i.test(String(d.start || '')) ? pill('READY TO START SOON') : '');

  const card = '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#f4f6fb;border-left:4px solid ' + GOLD + ';border-radius:8px;margin-bottom:18px"><tr><td style="padding:16px 18px">' +
    '<div style="font:bold 22px ' + SERIF + ';color:' + NAVY + '">' + esc_(name) + '</div>' +
    '<div style="font:15px/1.8 ' + SANS + ';color:#2a3042;margin-top:4px">' + esc_(d.location || '-') + badge + '<br>' +
    esc_(d.phone || '-') + ' &nbsp;&middot;&nbsp; ' + esc_(email) + '</div></td></tr></table>';

  const buttons = '<table role="presentation" cellpadding="0" cellspacing="0" style="margin:0 0 22px"><tr>' +
    alertButton_('mailto:' + email + '?subject=' + encodeURIComponent('Kumusta, ' + first + '! Si Cynthia ito'), 'Reply to ' + esc_(first), true) +
    (phoneDigits.length >= 7 ? alertButton_('tel:' + phoneDigits, 'Call ' + esc_(first), false) : '') +
    '</tr></table>';

  const rows = [['Trabaho ngayon', d.work], ['Experience sa real estate', d.experience], ['PRC license', d.license], ['Oras kada araw', d.hours], ['Gusto mag-start', d.start], ['Best time para kontakin', d.bestTime], ['Saan nalaman', d.heard], ['Commission-based', d.ack ? 'Nag-confirm siya na naiintindihan niya' : '-']];
  const table = '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-top:1px solid #e6e9f0;margin-bottom:18px">' +
    rows.map(function (r) {
      return '<tr><td valign="top" width="42%" style="padding:10px 10px 10px 0;border-bottom:1px solid #e6e9f0;font:bold 13px ' + SANS + ';color:' + NAVY + '">' + r[0] + '</td>' +
        '<td style="padding:10px 0;border-bottom:1px solid #e6e9f0;font:15px/1.5 ' + SANS + ';color:#2a3042">' + esc_(r[1] || '-') + '</td></tr>';
    }).join('') + '</table>';

  const status = confirmationSent
    ? 'Na-send na ang welcome email sa kanya. Ikaw na ang next.'
    : 'Hindi na-send ang welcome email sa kanya, kaya i-contact mo na siya agad.';

  const body =
    para_('May sumagot lang ng agent survey: si <strong>' + esc_(first) + '</strong>. Nandito ang mga sagot niya, at isang click lang para makausap mo siya.') +
    card + buttons + table +
    quote_('Bakit gusto niyang maging agent', d.why || '-') +
    quote_('Tanong niya para kay Cynthia', d.question || 'Wala siyang tanong.') +
    para_(status) +
    '<p style="margin:0 0 6px;font:14px ' + SANS + '"><a href="' + sheetUrl + '" style="color:' + NAVY + ';font-weight:bold">Buksan ang Recruits sheet &rarr;</a></p>';

  MailApp.sendEmail({
    to: OWNER_EMAIL,
    subject: 'Bagong agent recruit: ' + name + (d.location ? ' (' + d.location + ')' : ''),
    body: 'Bagong agent recruit: ' + name + '\nLokasyon: ' + (d.location || '-') + '\nPhone: ' + (d.phone || '-') + '\nEmail: ' + email +
      '\nTrabaho: ' + (d.work || '-') + '\nExperience: ' + (d.experience || '-') + '\nPRC license: ' + (d.license || '-') + '\nOras kada araw: ' + (d.hours || '-') + '\nGusto mag-start: ' + (d.start || '-') + '\nBest time para kontakin: ' + (d.bestTime || '-') + '\nSaan nalaman: ' + (d.heard || '-') + '\nCommission-based: ' + (d.ack ? 'Nag-confirm siya' : '-') +
      '\n\nBakit gusto niyang maging agent:\n' + (d.why || '-') + '\n\nTanong:\n' + (d.question || '-') + '\n\n' + status,
    htmlBody: emailShell_('May bagong agent recruit!', body, images, { signature: false, footer: 'Awtomatikong ipinadala mula sa recruitment page.' }),
    inlineImages: images, name: 'Recruitment Page', replyTo: email
  });
}
