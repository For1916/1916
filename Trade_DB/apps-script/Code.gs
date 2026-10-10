/**
 * 중고거래 장부 — 독립 Apps Script 웹 앱
 *
 * 시트에 묶인 스크립트가 아닙니다. SpreadsheetApp.openById 로 스프레드시트를 엽니다.
 * 배포는 지금 /exec 주소를 가진 그 프로젝트에서 새 버전으로 올립니다.
 *
 * 비밀번호는 코드에 적지 않습니다. 스크립트 속성 TRADE_DB_PASSWORD 에만 둡니다.
 * 읽기·추가·수정·삭제는 그 값과 맞은 뒤에만 시트를 엽니다. GET 은 행을 돌려주지 않습니다.
 * 페이지는 POST {action:'list'|'auth'|'insert'|'update'|'delete', password, id, record} 를 보냅니다.
 *
 * 저장 방식은 예전 스크립트와 같습니다. setup 은 편집기에서 실행하고, 셀을 바꾸지 않습니다.
 */

const SHEET_NAME = 'jonggo_trades';
const SPREADSHEET_ID = '1IdfcAJQj1O9IgxxKUa4J-_hMSR-A6Vp-kSONFg78FRQ';
const PASSWORD_KEY = 'TRADE_DB_PASSWORD';

function getSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['id','buy_date','category','item_name','buy_type','buy_platform','buy_price','sell_date','sell_platform','sell_price','status','memo','created_at','updated_at']);
  }
  return sheet;
}

function json_(value) {
  return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  return json_({data: null, error: '비밀번호가 필요합니다.'});
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return json_({data: null, error: String(err)});
  }
  if (!body || typeof body !== 'object') return json_({data: null, error: String(new Error('JSON 형식이 올바르지 않습니다.'))});
  const auth = checkPassword_(body.password);
  if (!auth.ok) return json_({data: null, error: auth.error});
  if (body.action === 'auth') return json_({data: {ok: true}, error: null});
  if (body.action === 'list') {
    try {
      return json_(readAll_());
    } catch (err) {
      return json_({data: null, error: String(err)});
    }
  }
  return mutate_(body);
}

function readAll_() {
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return {data: [], error: null};
  const headers = values.shift();
  const data = values.filter(row => row.some(v => v !== '')).map(row => Object.fromEntries(headers.map((h, i) => {
    let value = row[i] === '' ? null : row[i];
    if (value instanceof Date) value = Utilities.formatDate(value, Session.getScriptTimeZone() || 'Asia/Seoul', 'yyyy-MM-dd');
    return [h, value];
  })));
  return {data, error: null};
}

function mutate_(body) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sheet = getSheet_();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const now = new Date().toISOString();
    const action = body.action;
    if (action === 'insert') {
      const record = body.record || {};
      const row = headers.map(h => h === 'id' ? Utilities.getUuid() : h === 'created_at' || h === 'updated_at' ? now : record[h] ?? '');
      sheet.appendRow(row);
      return json_({data: row, error: null});
    }
    const idIndex = headers.indexOf('id');
    const ids = sheet.getRange(2, idIndex + 1, Math.max(sheet.getLastRow() - 1, 1), 1).getValues().flat();
    const rowIndex = ids.findIndex(id => String(id) === String(body.id));
    if (rowIndex < 0) return json_({data: null, error: 'Record not found'});
    const actualRow = rowIndex + 2;
    if (action === 'update') {
      const record = body.record || {};
      const old = sheet.getRange(actualRow, 1, 1, headers.length).getValues()[0];
      const row = headers.map((h, i) => h === 'id' || h === 'created_at' ? old[i] : h === 'updated_at' ? now : record[h] ?? '');
      sheet.getRange(actualRow, 1, 1, headers.length).setValues([row]);
      return json_({data: row, error: null});
    }
    if (action === 'delete') {
      sheet.deleteRow(actualRow);
      return json_({data: true, error: null});
    }
    return json_({data: null, error: 'Unknown action'});
  } catch (err) {
    return json_({data: null, error: String(err)});
  } finally {
    lock.releaseLock();
  }
}

function checkPassword_(password) {
  const expected = PropertiesService.getScriptProperties().getProperty(PASSWORD_KEY) || '';
  if (!expected) {
    return {ok: false, error: '비밀번호가 아직 설정되지 않았습니다. 프로젝트 설정 → 스크립트 속성에 TRADE_DB_PASSWORD 를 넣으세요.'};
  }
  if (!passwordsMatch_(password, expected)) return {ok: false, error: '비밀번호가 올바르지 않습니다.'};
  return {ok: true};
}

function passwordsMatch_(given, expected) {
  given = String(given == null ? '' : given).trim();
  expected = String(expected == null ? '' : expected).trim();
  if (given.length > 500) given = given.slice(0, 500);
  const n = Math.max(given.length, expected.length);
  let diff = given.length ^ expected.length;
  for (let i = 0; i < n; i++) {
    const a = i < given.length ? given.charCodeAt(i) : 0;
    const b = i < expected.length ? expected.charCodeAt(i) : 0;
    diff |= a ^ b;
  }
  return diff === 0;
}

/**
 * 편집기 함수 목록에서 실행합니다. 셀을 만들거나 고치지 않습니다.
 * 비밀번호는 프로젝트 설정 → 스크립트 속성 TRADE_DB_PASSWORD 에 넣습니다.
 */
function setup() {
  const lines = [];
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    lines.push('스프레드시트: ' + ss.getName());
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      lines.push('탭 jonggo_trades 가 없습니다. 탭을 만들거나 행을 바꾸지 않았습니다.');
    } else {
      lines.push('탭: ' + SHEET_NAME);
      lines.push('마지막 행: ' + sheet.getLastRow());
      if (sheet.getLastRow() > 0 && sheet.getLastColumn() > 0) {
        const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        lines.push('헤더: ' + headers.join(', '));
      }
    }
  } catch (err) {
    lines.push('시트를 열지 못했습니다: ' + err);
  }
  const secret = PropertiesService.getScriptProperties().getProperty(PASSWORD_KEY) || '';
  if (!secret) {
    lines.push('비밀번호: 없음. 프로젝트 설정 → 스크립트 속성에 TRADE_DB_PASSWORD 를 추가한 뒤 setup 을 다시 실행하세요.');
  } else if (secret.trim().length < 4 || secret.trim().length > 20) {
    lines.push('비밀번호: 설정됨. 화면 입력칸은 4자 이상 20자 이하입니다. 길이가 다르면 속성 값을 그 범위로 바꾸세요.');
  } else {
    lines.push('비밀번호: 설정됨');
  }
  lines.push('셀은 변경하지 않았습니다.');
  const text = lines.join('\n');
  Logger.log(text);
  return text;
}
