/**
 * Резюме договору · ESCORE — серверна частина (Apps Script).
 *
 * Що робить:
 *   action "parse" — приймає скан договору (PDF або фото) і текст інструкції,
 *                    перевіряє вхід Google, зберігає скан у папку Google Drive (якщо задана)
 *                    і надсилає все в Claude API. Повертає відповідь AI.
 *   action "save"  — дописує перевірений рядок у реєстр договорів (Google Sheets).
 * Ключ API зберігається тут, у властивостях скрипта, і не потрапляє на сайт.
 *
 * Властивості скрипта (Project Settings → Script properties):
 *   ANTHROPIC_API_KEY — ключ з console.anthropic.com (обов'язково для розбору)
 *   CLIENT_ID         — OAuth Client ID сайту (той самий, що в index.html)
 *   ALLOWED_EMAILS    — дозволені email через кому (порожньо = будь-хто з доступом до OAuth-клієнта)
 *   SPREADSHEET_ID    — ID таблиці-реєстру (з її адреси: /spreadsheets/d/<ID>/edit) (обов'язково для реєстру)
 *   SHEET_NAME        — назва аркуша (необов'язково, за замовчуванням «Реєстр»)
 *   DRIVE_FOLDER_ID   — ID папки Google Drive для сканів (необов'язково; порожньо = скани не зберігаються)
 *   MODEL             — модель Claude (необов'язково, за замовчуванням нижче)
 */

const MODEL_DEFAULT = 'claude-sonnet-5-5';
const SHEET_DEFAULT = 'Реєстр';
const MAX_FILES_BYTES = 22 * 1024 * 1024;   // ліміт запиту Claude API ~32 МБ, base64 додає третину
const MAX_PROMPT_CHARS = 60000;
const PDF = 'application/pdf';
const IMAGES = ['image/jpeg', 'image/png', 'image/webp', 'image/gif'];

function doPost(e) {
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const props = PropertiesService.getScriptProperties();

    const email = verifyUser_(req.token, props.getProperty('CLIENT_ID'));
    const allowed = (props.getProperty('ALLOWED_EMAILS') || '')
      .toLowerCase().split(/[\s,;]+/).filter(String);
    if (!email || (allowed.length && allowed.indexOf(email) === -1)) {
      return json_({ ok: false, error: 'forbidden' });
    }

    if (req.action === 'save') return json_(save_(req, email, props));
    return json_(parse_(req, email, props));
  } catch (err) {
    return json_({ ok: false, error: 'server', message: String(err) });
  }
}

/* ---------- розбір договору ---------- */
function parse_(req, email, props) {
  const key = props.getProperty('ANTHROPIC_API_KEY');
  if (!key) return { ok: false, error: 'no_key' };

  const prompt = String(req.prompt || '');
  const files = Array.isArray(req.files) ? req.files : [];
  if (!prompt || prompt.length > MAX_PROMPT_CHARS || !files.length) return { ok: false, error: 'bad_request' };

  let total = 0;
  const content = [];
  const blobs = [];
  for (let i = 0; i < files.length; i++) {
    const f = files[i] || {};
    const mime = String(f.mime || '');
    const data = String(f.data || '');
    if (mime !== PDF && IMAGES.indexOf(mime) === -1) return { ok: false, error: 'bad_file' };
    total += Math.floor(data.length * 3 / 4);
    if (total > MAX_FILES_BYTES) return { ok: false, error: 'too_big' };
    content.push(mime === PDF
      ? { type: 'document', source: { type: 'base64', media_type: PDF, data: data } }
      : { type: 'image', source: { type: 'base64', media_type: mime, data: data } });
    blobs.push({ name: String(f.name || ('скан_' + (i + 1))), mime: mime, data: data });
  }
  content.push({ type: 'text', text: prompt });

  const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify({
      model: props.getProperty('MODEL') || MODEL_DEFAULT,
      max_tokens: 8000,
      messages: [{ role: 'user', content: content }]
    }),
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  const body = JSON.parse(res.getContentText() || '{}');
  if (code !== 200) {
    return { ok: false, error: 'anthropic', message: (body.error && body.error.message) || ('HTTP ' + code) };
  }
  const text = (body.content || [])
    .filter(function (b) { return b.type === 'text'; })
    .map(function (b) { return b.text; }).join('');

  // Скан — у папку Drive (лише після успішного розбору, щоб не накопичувати сміття)
  let fileUrls = [];
  const folderId = props.getProperty('DRIVE_FOLDER_ID');
  if (folderId) {
    try {
      const folder = DriveApp.getFolderById(folderId);
      const stamp = Utilities.formatDate(new Date(), 'Europe/Kyiv', 'yyyy-MM-dd_HH-mm');
      fileUrls = blobs.map(function (b) {
        const blob = Utilities.newBlob(Utilities.base64Decode(b.data), b.mime, stamp + '_' + b.name);
        return folder.createFile(blob).getUrl();
      });
    } catch (err) {
      console.warn('Drive save failed: ' + err);
    }
  }

  console.log('Contract parsed by ' + email + ', tokens in/out: ' +
    (body.usage ? body.usage.input_tokens + '/' + body.usage.output_tokens : '?'));

  return { ok: true, text: text, usage: body.usage || null, fileUrls: fileUrls };
}

/* ---------- запис у реєстр ---------- */
function save_(req, email, props) {
  const id = props.getProperty('SPREADSHEET_ID');
  if (!id) return { ok: false, error: 'no_sheet' };
  const cols = Array.isArray(req.columns) ? req.columns : [];
  if (!cols.length) return { ok: false, error: 'bad_request' };

  const ss = SpreadsheetApp.openById(id);
  const name = props.getProperty('SHEET_NAME') || SHEET_DEFAULT;
  const sh = ss.getSheetByName(name) || ss.insertSheet(name);

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const now = Utilities.formatDate(new Date(), 'Europe/Kyiv', 'dd.MM.yyyy HH:mm');
    const all = [['Додано', now], ['Хто додав', email]].concat(
      cols.map(function (c) { return [String(c[0] || ''), String(c[1] == null ? '' : c[1])]; }));

    // Заголовки: якщо аркуш порожній — пишемо; якщо з'явилась нова колонка — додаємо в кінець.
    let headers = sh.getLastColumn() ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String) : [];
    if (!headers.filter(String).length) headers = [];
    all.forEach(function (c) { if (headers.indexOf(c[0]) === -1) headers.push(c[0]); });
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#e2efe8');
    sh.setFrozenRows(1);

    const row = headers.map(function () { return ''; });
    all.forEach(function (c) { row[headers.indexOf(c[0])] = cell_(c[0], c[1]); });
    const r = sh.getLastRow() + 1;
    sh.getRange(r, 1, 1, row.length).setValues([row]).setWrap(true).setVerticalAlignment('top');

    return { ok: true, row: r, sheetUrl: ss.getUrl() + '#gid=' + sh.getSheetId() };
  } finally {
    lock.releaseLock();
  }
}

/** Значення клітинки: посилання на скан — клікабельні; решта — як текст (без формул). */
function cell_(header, v) {
  v = String(v || '');
  if (header === 'Скан' && /^https:\/\/\S+$/.test(v.split('\n')[0])) {
    const urls = v.split('\n').filter(String);
    if (urls.length === 1) return '=HYPERLINK("' + urls[0].replace(/"/g, '') + '","відкрити скан")';
  }
  if (/^[=+\-@]/.test(v)) return "'" + v;   // захист від формул у тексті
  return v;
}

/** Перевірка: це справжній вхід Google через наш сайт. Повертає email або null. */
function verifyUser_(token, clientId) {
  if (!token) return null;
  const r = UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?access_token=' +
    encodeURIComponent(token), { muteHttpExceptions: true });
  if (r.getResponseCode() !== 200) return null;
  const t = JSON.parse(r.getContentText());
  if (clientId && t.azp !== clientId && t.aud !== clientId) return null;
  if (String(t.email_verified) !== 'true') return null;
  return String(t.email || '').toLowerCase() || null;
}

/** Перевірка, що веб-додаток працює: відкрийте його URL у браузері. */
function doGet() {
  return json_({ ok: true, service: 'contract-summary' });
}

/** Запустіть один раз вручну в редакторі, щоб надати скрипту доступ до Drive і Sheets. */
function authorize() {
  const p = PropertiesService.getScriptProperties();
  if (p.getProperty('SPREADSHEET_ID')) SpreadsheetApp.openById(p.getProperty('SPREADSHEET_ID')).getName();
  if (p.getProperty('DRIVE_FOLDER_ID')) DriveApp.getFolderById(p.getProperty('DRIVE_FOLDER_ID')).getName();
  UrlFetchApp.fetch('https://www.google.com', { muteHttpExceptions: true });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
