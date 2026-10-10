/**
 * 중고거래 장부 — Google Apps Script 웹 앱
 *
 * 시트 [중고거래 장부 - Google Sheets 이전] 의 확장 프로그램 > Apps Script 에
 * 이 파일 전체를 붙여 넣습니다. 시트에 묶인 스크립트이므로
 * SpreadsheetApp.getActiveSpreadsheet() 를 사용합니다.
 *
 * 비밀번호는 코드에 적지 않습니다. 스크립트 속성 TRADE_DB_PASSWORD 에만 두고,
 * 웹 앱은 읽기·추가·수정·삭제 모두 그 값과 비교한 뒤에만 시트를 다룹니다.
 * GET 은 데이터를 돌려주지 않습니다. 페이지는 POST 본문에 비밀번호를 넣습니다.
 *
 * 열 이름은 헤더로 찾습니다. 없는 열을 새로 만들거나, 헤더 밖의 칸을 지우지 않습니다.
 * 처음 한 번: 스프레드시트를 새로고침한 뒤 메뉴 [중고거래 장부] > [초기 설정]
 * 또는 편집기에서 setup 함수를 실행합니다. 다시 실행해도 기존 행은 바꾸지 않습니다.
 *
 * 시간대는 바꾸지 마세요. 구매일·판매일이 하루 밀릴 수 있습니다.
 */

var PASSWORD_KEY = 'TRADE_DB_PASSWORD';
var APP_VERSION = '1';
var HEADER_SCAN_ROWS = 8;
var HEADER_SCAN_COLS = 40;

var FIELDS = [
  { key: 'id', names: ['id', '아이디'], kind: 'id' },
  { key: 'buy_date', names: ['buy_date', '구매일'], kind: 'date' },
  { key: 'category', names: ['category', '카테고리'], kind: 'text' },
  { key: 'item_name', names: ['item_name', '품목명'], kind: 'text' },
  { key: 'buy_type', names: ['buy_type', '구매유형'], kind: 'text' },
  { key: 'buy_platform', names: ['buy_platform', '구매플랫폼'], kind: 'text' },
  { key: 'buy_price', names: ['buy_price', '구매가'], kind: 'number' },
  { key: 'sell_date', names: ['sell_date', '판매일'], kind: 'date' },
  { key: 'sell_platform', names: ['sell_platform', '판매플랫폼'], kind: 'text' },
  { key: 'sell_price', names: ['sell_price', '판매가'], kind: 'number' },
  { key: 'status', names: ['status', '상태'], kind: 'text' },
  { key: 'memo', names: ['memo', '메모'], kind: 'text' },
  { key: 'created_at', names: ['created_at', '생성시각'], kind: 'raw' },
  { key: 'updated_at', names: ['updated_at', '수정시각'], kind: 'raw' }
];

var REQUIRED_KEYS = ['id', 'item_name', 'buy_date'];
var CLIENT_KEYS = [
  'buy_date', 'category', 'item_name', 'buy_type', 'buy_platform', 'buy_price',
  'sell_date', 'sell_platform', 'sell_price', 'status', 'memo'
];

// ── 웹 앱 ────────────────────────────────────────────

function doGet(e) {
  if (e && e.parameter && e.parameter.password) {
    return json_({ data: null, error: '비밀번호는 주소에 넣지 마세요. 화면에서 입력하세요.' });
  }
  return json_({ data: null, error: '비밀번호가 필요합니다.' });
}

function doPost(e) {
  var data = {};
  try {
    data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return json_({ data: null, error: 'JSON 형식이 올바르지 않습니다.' });
  }
  if (!data || typeof data !== 'object' || Object.prototype.toString.call(data) === '[object Array]') {
    return json_({ data: null, error: 'JSON 형식이 올바르지 않습니다.' });
  }
  return handle_(data);
}

function handle_(data) {
  try {
    authenticate_(data.password);
    var action = String(data.action || 'list').trim().toLowerCase();
    if (action === 'auth' || action === 'check') return json_({ data: { ok: true }, error: null });
    if (action === 'list' || action === 'read' || action === 'get' || action === 'getall' || action === 'all') {
      return json_({ data: listTrades_(), error: null });
    }
    if (action === 'insert' || action === 'add' || action === 'create') {
      return json_({ data: withLock_(function() { return insertTrade_(data.record); }), error: null });
    }
    if (action === 'update' || action === 'edit') {
      return json_({ data: withLock_(function() { return updateTrade_(data.id, data.record); }), error: null });
    }
    if (action === 'delete' || action === 'remove') {
      return json_({ data: withLock_(function() { return deleteTrade_(data.id); }), error: null });
    }
    return json_({ data: null, error: '알 수 없는 요청입니다.' });
  } catch (err) {
    var message = (err && err.safe) ? String(err.message) : '요청을 처리하지 못했습니다.';
    return json_({ data: null, error: message });
  }
}

function authenticate_(password) {
  var expected = PropertiesService.getScriptProperties().getProperty(PASSWORD_KEY) || '';
  if (!expected) fail_('비밀번호가 아직 설정되지 않았습니다. 시트 메뉴의 초기 설정을 실행하세요.');
  var given = String(password == null ? '' : password).trim();
  if (!passwordsMatch_(given, expected)) fail_('비밀번호가 올바르지 않습니다.');
}

function passwordsMatch_(given, expected) {
  given = String(given == null ? '' : given);
  expected = String(expected == null ? '' : expected);
  if (given.length > 500) given = given.slice(0, 500);
  var n = Math.max(given.length, expected.length);
  var diff = given.length ^ expected.length;
  var i;
  for (i = 0; i < n; i++) {
    var a = i < given.length ? given.charCodeAt(i) : 0;
    var b = i < expected.length ? expected.charCodeAt(i) : 0;
    diff |= a ^ b;
  }
  return diff === 0;
}

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) fail_('다른 저장이 진행 중입니다. 잠시 후 다시 시도하세요.');
  try { return fn(); }
  finally { lock.releaseLock(); }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function fail_(message) {
  var err = new Error(message);
  err.safe = true;
  throw err;
}

// ── 시트 찾기 (헤더만 보고, 데이터는 고치지 않음) ─────

function activeSs_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) fail_('스프레드시트에 묶인 스크립트가 아닙니다. 중고거래 장부 시트에서 확장 프로그램 → Apps Script 로 연 코드를 바꾸세요.');
  return ss;
}

function listTrades_() {
  return readLocated_(locate_(activeSs_()));
}

function insertTrade_(record) {
  var located = locate_(activeSs_());
  var saved = writeRecord_(located, null, record, true);
  return saved;
}

function updateTrade_(id, record) {
  var located = locate_(activeSs_());
  var row = findDataRow_(located, id);
  return writeRecord_(located, row, record, false);
}

function deleteTrade_(id) {
  var located = locate_(activeSs_());
  var row = findDataRow_(located, id);
  if (row <= located.headerRow) fail_('헤더 행은 지울 수 없습니다.');
  located.sheet.deleteRow(row);
  return { id: idKeyOut_(id) };
}

function locate_(ss) {
  if (!ss) fail_('스프레드시트를 찾지 못했습니다.');
  var sheets = ss.getSheets();
  var matches = [];
  var i;
  for (i = 0; i < sheets.length; i++) {
    var found = locateHeader_(sheets[i]);
    if (found) matches.push(found);
  }
  if (!matches.length) {
    fail_('거래 시트를 찾지 못했습니다. 첫 8행 안에 id(또는 아이디), item_name(또는 품목명), buy_date(또는 구매일) 헤더가 있는 시트가 필요합니다. 데이터는 바꾸지 않았습니다.');
  }
  if (matches.length === 1) return matches[0];
  var named = matches.filter(function(m) { return normHeader_(m.sheet.getName()) === 'jonggotrades'; });
  if (named.length === 1) return named[0];
  var korean = matches.filter(function(m) { return String(m.sheet.getName()).indexOf('중고') !== -1; });
  if (korean.length === 1) return korean[0];
  fail_('헤더가 맞는 시트가 여러 개입니다: ' + matches.map(function(m) { return m.sheet.getName(); }).join(', ') + '. 사용할 시트 이름을 jonggo_trades 로 바꿔 주세요. 데이터는 바꾸지 않았습니다.');
}

function locateHeader_(sh) {
  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (!lastRow || !lastCol) return null;
  var scanRows = Math.min(lastRow, HEADER_SCAN_ROWS);
  var scanCols = Math.min(lastCol, HEADER_SCAN_COLS);
  var values = sh.getRange(1, 1, scanRows, scanCols).getValues();
  var r;
  for (r = 0; r < values.length; r++) {
    var map = {};
    var duplicates = [];
    var c;
    for (c = 0; c < values[r].length; c++) {
      var header = normHeader_(values[r][c]);
      if (!header) continue;
      var f;
      for (f = 0; f < FIELDS.length; f++) {
        var names = FIELDS[f].names;
        var n;
        var hit = false;
        for (n = 0; n < names.length; n++) {
          if (normHeader_(names[n]) === header) { hit = true; break; }
        }
        if (!hit) continue;
        if (map[FIELDS[f].key] != null) duplicates.push(String(values[r][c]));
        else map[FIELDS[f].key] = c + 1;
      }
    }
    var ok = true;
    var req;
    for (req = 0; req < REQUIRED_KEYS.length; req++) {
      if (map[REQUIRED_KEYS[req]] == null) { ok = false; break; }
    }
    if (!ok) continue;
    return {
      sheet: sh,
      headerRow: r + 1,
      map: map,
      duplicates: duplicates,
      scanCols: scanCols
    };
  }
  return null;
}

function readLocated_(located) {
  var sh = located.sheet;
  var lastRow = sh.getLastRow();
  if (lastRow <= located.headerRow) return [];
  var lastCol = Math.max(sh.getLastColumn(), located.scanCols);
  var values = sh.getRange(located.headerRow + 1, 1, lastRow - located.headerRow, lastCol).getValues();
  var out = [];
  var i;
  for (i = 0; i < values.length; i++) {
    var obj = rowToObject_(values[i], located.map);
    if (obj) out.push(obj);
  }
  return out;
}

function rowToObject_(row, map) {
  var obj = {};
  var any = false;
  var f;
  for (f = 0; f < FIELDS.length; f++) {
    var key = FIELDS[f].key;
    var col = map[key];
    var value = null;
    if (col != null && col <= row.length) value = readCell_(row[col - 1], FIELDS[f].kind);
    obj[key] = value;
    if (value != null && value !== '') any = true;
  }
  if (!any) return null;
  return obj;
}

function readCell_(value, kind) {
  if (value == null || value === '') return null;
  if (kind === 'date') return readDateCell_(value);
  if (kind === 'number') return readNumberCell_(value);
  if (kind === 'id') return readIdCell_(value);
  if (Object.prototype.toString.call(value) === '[object Date]') {
    if (isNaN(value.getTime())) return null;
    return value.toISOString();
  }
  if (typeof value === 'number' && isFinite(value)) return value;
  var text = String(value);
  var trimmed = text.trim();
  if (!trimmed || trimmed.toLowerCase() === 'null' || trimmed === 'NaN') return null;
  return text;
}

function readDateCell_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]') {
    if (isNaN(value.getTime())) return null;
    return value.toISOString();
  }
  if (typeof value === 'string') {
    var s = value.trim();
    if (!s || s.toLowerCase() === 'null' || s === 'NaN') return null;
    return s;
  }
  return null;
}

function readNumberCell_(value) {
  if (typeof value === 'number' && isFinite(value)) return value;
  if (typeof value === 'boolean') return null;
  var s = String(value).trim();
  if (!s || s.toLowerCase() === 'null' || s === 'NaN' || s.charAt(0) === '#') return null;
  var n = Number(s.replace(/,/g, ''));
  return isFinite(n) ? n : null;
}

function readIdCell_(value) {
  if (typeof value === 'number' && isFinite(value)) return value;
  var s = String(value).trim();
  if (!s || s.toLowerCase() === 'null' || s === 'NaN') return null;
  if (/^-?\d+(\.0+)?$/.test(s)) return Number(s);
  return s;
}

function findDataRow_(located, id) {
  var hits = idHits_(located, id);
  if (!idKey_(id)) fail_('수정할 거래를 지정하지 않았습니다.');
  if (hits.length !== 1) {
    fail_(hits.length ? '같은 id 가 여러 행입니다. 시트에서 직접 확인해 주세요.' : '해당 거래를 찾지 못했습니다.');
  }
  return hits[0];
}

function idHits_(located, id) {
  var want = idKey_(id);
  if (!want) return [];
  var sh = located.sheet;
  var lastRow = sh.getLastRow();
  if (lastRow <= located.headerRow) return [];
  var values = sh.getRange(located.headerRow + 1, located.map.id, lastRow - located.headerRow, 1).getValues();
  var hits = [];
  var i;
  for (i = 0; i < values.length; i++) {
    if (idKey_(values[i][0]) === want) hits.push(located.headerRow + 1 + i);
  }
  return hits;
}

function idExists_(located, id) {
  return idHits_(located, id).length > 0;
}

function idKey_(value) {
  if (value == null || value === '') return '';
  if (typeof value === 'number' && isFinite(value)) return String(value);
  var s = String(value).trim();
  if (!s || s.toLowerCase() === 'null') return '';
  if (/^-?\d+(\.0+)?$/.test(s)) return String(Number(s));
  return s;
}

function idKeyOut_(value) {
  var key = idKey_(value);
  if (/^-?\d+$/.test(key)) return Number(key);
  return key;
}

function writeRecord_(located, row, record, isInsert) {
  if (!record || typeof record !== 'object' || Object.prototype.toString.call(record) === '[object Array]') {
    fail_('저장할 내용이 없습니다.');
  }
  var sh = located.sheet;
  var map = located.map;
  var tz = sh.getParent().getSpreadsheetTimeZone();
  var prepared = {};
  var k;
  for (k = 0; k < CLIENT_KEYS.length; k++) {
    var key = CLIENT_KEYS[k];
    if (!Object.prototype.hasOwnProperty.call(record, key)) continue;
    var cleaned = cleanForKind_(record[key], kindOf_(key), tz);
    if (cleaned != null && map[key] == null) fail_('시트에 없는 열이라 저장하지 않았습니다: ' + key);
    prepared[key] = cleaned;
  }
  if (isInsert && !prepared.item_name) fail_('품목명을 입력해 주세요.');
  if (!isInsert && Object.prototype.hasOwnProperty.call(record, 'item_name') && !prepared.item_name) fail_('품목명을 입력해 주세요.');

  if (isInsert) {
    row = Math.max(sh.getLastRow() + 1, located.headerRow + 1);
    var newId = Utilities.getUuid();
    var guard = 0;
    while (idExists_(located, newId) && guard < 5) {
      newId = Utilities.getUuid();
      guard++;
    }
    if (idExists_(located, newId)) fail_('새 id 를 만들지 못했습니다.');
    writeCell_(sh, row, map.id, newId);
    stamp_(sh, map, row, 'created_at');
  }
  stamp_(sh, map, row, 'updated_at');

  for (k = 0; k < CLIENT_KEYS.length; k++) {
    var field = CLIENT_KEYS[k];
    if (!Object.prototype.hasOwnProperty.call(prepared, field)) continue;
    var col = map[field];
    if (col == null) continue;
    if (prepared[field] == null) sh.getRange(row, col).clearContent();
    else writeCell_(sh, row, col, prepared[field]);
  }

  var width = Math.max(sh.getLastColumn(), 1);
  var values = sh.getRange(row, 1, 1, width).getValues();
  var obj = rowToObject_(values[0], map);
  if (!obj) fail_('저장한 행을 다시 읽지 못했습니다.');
  return obj;
}

function stamp_(sh, map, row, key) {
  var col = map[key];
  if (col == null) return;
  writeCell_(sh, row, col, new Date().toISOString());
}

function writeCell_(sh, row, col, value) {
  sh.getRange(row, col).setValue(value);
}

function kindOf_(key) {
  var i;
  for (i = 0; i < FIELDS.length; i++) if (FIELDS[i].key === key) return FIELDS[i].kind;
  return 'text';
}

function cleanForKind_(value, kind, tz) {
  if (kind === 'date') return cleanDate_(value, tz);
  if (kind === 'number') return cleanNumber_(value);
  return cleanText_(value, 5000);
}

function cleanText_(value, maxLen) {
  if (value == null) return null;
  var s = String(value).trim();
  if (!s || s.toLowerCase() === 'null' || s === 'NaN') return null;
  if (s.length > maxLen) fail_('입력 내용이 너무 깁니다.');
  return s;
}

function cleanNumber_(value) {
  if (value == null || value === '') return null;
  if (typeof value === 'string') {
    var s = value.trim();
    if (!s || s.toLowerCase() === 'null' || s === 'NaN') return null;
    value = s.replace(/,/g, '');
  }
  var n = Number(value);
  if (!isFinite(n)) fail_('금액이 올바르지 않습니다.');
  return n;
}

function cleanDate_(value, tz) {
  if (value == null || value === '') return null;
  var s = String(value).trim();
  if (!s || s.toLowerCase() === 'null' || s === 'NaN') return null;
  var m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!m) fail_('날짜 형식이 올바르지 않습니다.');
  var y = Number(m[1]);
  var mo = Number(m[2]);
  var d = Number(m[3]);
  var check = new Date(Date.UTC(y, mo - 1, d));
  if (check.getUTCFullYear() !== y || check.getUTCMonth() !== mo - 1 || check.getUTCDate() !== d) {
    fail_('날짜 형식이 올바르지 않습니다.');
  }
  var iso = y + '-' + (mo < 10 ? '0' : '') + mo + '-' + (d < 10 ? '0' : '') + d;
  return Utilities.parseDate(iso, tz || Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function normHeader_(value) {
  return String(value == null ? '' : value).trim().toLowerCase().replace(/[\s_\-./]+/g, '');
}

// ── 초기 설정 ────────────────────────────────────────

/**
 * 비밀번호를 스크립트 속성에 저장하고, 거래 시트가 보이는지만 확인합니다.
 * 기존 행과 헤더는 바꾸지 않습니다.
 * @param {string=} initialPassword
 * @param {boolean=} alreadyPrompted 메뉴에서 이미 물어봤으면 true
 */
function setup(initialPassword, alreadyPrompted) {
  var lines = [];
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) fail_('활성 스프레드시트가 없습니다. 중고거래 장부 시트에서 확장 프로그램 > Apps Script 로 연 뒤 실행하세요.');
    lines.push(describeLocation_(ss));
  } catch (err) {
    lines.push('시트 확인 실패: ' + ((err && err.message) ? err.message : err));
  }
  lines.push(storePassword_(initialPassword, !alreadyPrompted));
  lines.push('버전 ' + APP_VERSION + '. 웹 앱은 배포 관리에서 새 버전으로 올려야 이 코드가 적용됩니다.');
  var text = lines.join('\n');
  Logger.log(text);
  return text;
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('중고거래 장부')
    .addItem('초기 설정', 'setupFromMenu')
    .addItem('비밀번호 설정', 'setPasswordFromMenu')
    .addItem('시트 점검', 'auditSheet')
    .addToUi();
}

function setupFromMenu() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt('중고거래 장부', '웹 장부 비밀번호를 입력하세요. 이미 설정돼 있으면 비워 두고 확인을 누르세요.', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  try {
    ui.alert(setup(String(res.getResponseText() || ''), true));
  } catch (err) {
    ui.alert(String(err && err.message ? err.message : err).slice(0, 1500));
  }
}

function setPasswordFromMenu() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt('비밀번호 설정', '웹 장부에 사용할 비밀번호를 입력하세요. 4자 이상 20자 이하입니다.', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  ui.alert(storePassword_(String(res.getResponseText() || ''), false));
}

function auditSheet() {
  var lines = [];
  try {
    var ss = activeSs_();
    lines.push('스프레드시트: ' + ss.getName());
    lines.push('시트 시간대: ' + ss.getSpreadsheetTimeZone());
    lines.push('스크립트 시간대: ' + Session.getScriptTimeZone());
    lines.push(describeLocation_(ss));
  } catch (err) {
    lines.push('시트 확인 실패: ' + ((err && err.message) ? err.message : err));
  }
  var secret = PropertiesService.getScriptProperties().getProperty(PASSWORD_KEY);
  lines.push(secret ? '비밀번호: 설정됨' : '비밀번호: 없음');
  var text = lines.join('\n');
  Logger.log(text);
  try { SpreadsheetApp.getUi().alert(text.slice(0, 1500)); } catch (e) {}
  return text;
}

function describeLocation_(ss) {
  var located = locate_(ss);
  var lines = [];
  lines.push('사용 시트: ' + located.sheet.getName());
  lines.push('헤더 행: ' + located.headerRow);
  var found = [];
  var missing = [];
  var i;
  for (i = 0; i < FIELDS.length; i++) {
    if (located.map[FIELDS[i].key] != null) found.push(FIELDS[i].key);
    else missing.push(FIELDS[i].key);
  }
  lines.push('찾은 열: ' + found.join(', '));
  lines.push(missing.length ? ('없는 열: ' + missing.join(', ')) : '없는 열: 없음');
  if (located.duplicates.length) lines.push('헤더가 두 번 있습니다. 왼쪽 열만 사용합니다: ' + located.duplicates.join(', '));
  var rows = readLocated_(located);
  var blankId = 0;
  for (i = 0; i < rows.length; i++) if (rows[i].id == null) blankId++;
  lines.push('데이터 행: ' + rows.length);
  lines.push('id 가 비어 있는 행: ' + blankId);
  lines.push('시간대는 바꾸지 마세요. 구매일 표시가 하루 달라질 수 있습니다.');
  return lines.join('\n');
}

function storePassword_(initialPassword, allowPrompt) {
  var props = PropertiesService.getScriptProperties();
  var existing = props.getProperty(PASSWORD_KEY) || '';
  var incoming = initialPassword == null ? '' : String(initialPassword).trim();
  if (!incoming && allowPrompt) {
    try {
      var ui = SpreadsheetApp.getUi();
      var res = ui.prompt('중고거래 장부', '웹 장부 비밀번호를 입력하세요. 이미 설정돼 있으면 비워 두고 확인을 누르세요.', ui.ButtonSet.OK_CANCEL);
      if (res.getSelectedButton() === ui.Button.OK) incoming = String(res.getResponseText() || '').trim();
    } catch (e) {}
  }
  if (incoming) {
    if (incoming.length < 4) return '비밀번호는 4자 이상으로 설정하세요. 저장하지 않았습니다.';
    if (incoming.length > 20) return '비밀번호는 20자 이하로 설정하세요. 화면 입력칸이 20자까지입니다. 저장하지 않았습니다.';
    props.setProperty(PASSWORD_KEY, incoming);
    return '비밀번호를 스크립트 속성 TRADE_DB_PASSWORD 에 저장했습니다.';
  }
  if (existing) return '기존 비밀번호를 그대로 사용합니다.';
  return '비밀번호가 없습니다. 프로젝트 설정 → 스크립트 속성에 TRADE_DB_PASSWORD 를 추가하거나, 시트 메뉴 중고거래 장부 → 비밀번호 설정을 실행하세요.';
}
