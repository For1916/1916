/**
 * 황소 워크북 장부 — Google Apps Script 웹 앱
 *
 * 시트 [황소 워크북 장부] 의 확장 프로그램 > Apps Script 에 이 파일 전체를 붙여 넣습니다.
 * 시트에 묶인 스크립트이므로 SpreadsheetApp.getActiveSpreadsheet() 를 사용합니다.
 *
 * 판매는 주문 연동의 금액(취소·환불 제외)과, 수입 시트에 직접 적은 금액입니다.
 * 수입 시트는 날짜·금액·메모만 둡니다. 주문에 이미 있는 판매를 여기에 다시 적으면 두 번 합산됩니다.
 * 지출 분류는 제본, AI, 광고, 박스 네 가지입니다.
 *
 * 비밀번호는 코드에 적지 않습니다. 스크립트 속성 LEDGER_PASSWORD 에만 둡니다.
 * 웹 앱은 매 요청마다 그 값과 비교하고, 주문자·연락처·주소·유입 경로는 반환하지 않습니다.
 *
 * 처음 한 번: 스프레드시트를 새로고침한 뒤 메뉴 [황소 장부] > [초기 설정]
 * 또는 편집기에서 setup 함수를 실행합니다. 다시 실행해도 됩니다.
 */

var APP_VERSION = '2';
var PASSWORD_KEY = 'LEDGER_PASSWORD';
var SOURCE_SPREADSHEET_ID = '1s_QC5gRuU7E07WGZrBtbMS_S80YHexlPPD_qYVDqTpM';
var SOURCE_SHEET_NAME = '주문 관리';
var ORDER_LAST_ROW_DEFAULT = 1000;
var MANUAL_LAST_ROW = 5000;

var PERSONAL_HEADER_RE = /주문자|고객명|수취인|받는\s*분|받는분|연락처|전화|휴대폰|핸드폰|휴대전화|주소|우편번호|이메일|유입|모아폼|moaform|답변\s*id|응답\s*id|answer\s*id|submission/i;
var DOC_MARKER = '[장부 안내]';
var EXPENSE_CATEGORIES = ['제본', 'AI', '광고', '박스'];
var SUMMARY_HEADERS = ['월', '판매', '제본', 'AI', '광고', '박스', '총지출', '순이익', '이익률'];
var SUMMARY_METRICS = { '월': 1, '판매': 1, '주문 수입': 1, '기타 수입': 1, '총수입': 1, '총지출': 1, '순이익': 1, '이익률': 1 };

var formulaSep_ = ',';

// ── 웹 앱 ────────────────────────────────────────────

function doGet(e) {
  return handleRequest_((e && e.parameter) || {});
}

function doPost(e) {
  var data = {};
  try {
    data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return json_({ ok: false, error: 'JSON 형식이 올바르지 않습니다.' });
  }
  if (e && e.parameter) {
    var key;
    for (key in e.parameter) {
      if (Object.prototype.hasOwnProperty.call(e.parameter, key) && data[key] == null) data[key] = e.parameter[key];
    }
  }
  return handleRequest_(data);
}

function handleRequest_(data) {
  try {
    data = data || {};
    var auth = authenticate_(data.password);
    if (!auth.ok) return json_({ ok: false, error: auth.error });
    var action = String(data.action || 'load').trim();
    if (action === 'auth' || action === 'check') return json_({ ok: true, version: APP_VERSION });
    if (action === 'load') return json_(loadLedger_());
    if (typeof data.record === 'string') data.record = parseJsonField_(data.record, 'record');
    if (typeof data.match === 'string') data.match = parseJsonField_(data.match, 'match');
    if (action === 'add' || action === 'update' || action === 'delete') {
      return json_(withLock_(function() { return mutate_(action, data); }));
    }
    return json_({ ok: false, error: '알 수 없는 요청입니다.' });
  } catch (err) {
    var message = (err && err.ledgerSafe) ? String(err.message) : '요청을 처리하지 못했습니다.';
    return json_({ ok: false, error: message });
  }
}

function authenticate_(password) {
  var expected = PropertiesService.getScriptProperties().getProperty(PASSWORD_KEY) || '';
  if (!expected) {
    return { ok: false, error: '비밀번호가 아직 설정되지 않았습니다. 시트 메뉴의 초기 설정을 실행하세요.' };
  }
  if (!passwordsMatch_(password, expected)) {
    return { ok: false, error: '비밀번호가 올바르지 않습니다.' };
  }
  return { ok: true };
}

function passwordsMatch_(given, expected) {
  given = String(given == null ? '' : given);
  expected = String(expected == null ? '' : expected);
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

function loadLedger_() {
  var ss = activeSs_();
  var incomeSheet = mustSheet_(ss, '수입');
  var expenseSheet = mustSheet_(ss, '지출');
  var orderSheet = mustSheet_(ss, '주문 연동');
  var summarySheet = ss.getSheetByName('월별 요약');
  var classSheet = ss.getSheetByName('분류');
  var income = readIncome_(incomeSheet);
  var expense = readExpense_(expenseSheet);
  var orders = readOrders_(orderSheet);
  var meta = readMeta_(classSheet);
  var summary = readSheetSummary_(summarySheet);
  if (!summaryHasNumbers_(summary)) {
    summary = computeSummary_(income, expense, orders);
    summary.source = 'computed';
  } else {
    summary.source = 'sheet';
    if (!summary.total) summary.total = totalFromMonths_(summary.months);
  }
  return {
    ok: true,
    version: APP_VERSION,
    meta: meta,
    summary: summary,
    transactions: { income: income, expense: expense, orders: orders }
  };
}

function mutate_(action, data) {
  var sheetName = String(data.sheet || '').trim();
  if (sheetName !== '수입' && sheetName !== '지출') fail_('수입 또는 지출만 수정할 수 있습니다.');
  var ss = activeSs_();
  var sh = mustSheet_(ss, sheetName);
  var map = headerMap_(sh);
  if (action === 'add') {
    var row = nextDataRow_(sh, map);
    writeRecord_(sh, map, row, sheetName, data.record || {}, null);
    return { ok: true, version: APP_VERSION, sheet: sheetName, row: row };
  }
  var rowNum = parseRow_(data.row);
  assertRowMatch_(sh, map, rowNum, data.match);
  if (action === 'delete') {
    sh.deleteRow(rowNum);
    return { ok: true, version: APP_VERSION, sheet: sheetName, row: rowNum };
  }
  writeRecord_(sh, map, rowNum, sheetName, data.record || {});
  return { ok: true, version: APP_VERSION, sheet: sheetName, row: rowNum };
}

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) fail_('다른 저장이 진행 중입니다. 잠시 후 다시 시도하세요.');
  try { return fn(); }
  finally { lock.releaseLock(); }
}

// ── 읽기 (개인정보·교재·구분 제외) ───────────────────

function readIncome_(sh) {
  return readMapped_(sh, [
    { key: 'date', names: ['날짜'], kind: 'date' },
    { key: 'type', names: ['구분'], kind: 'text' },
    { key: 'amount', names: ['금액'], kind: 'number' },
    { key: 'memo', names: ['메모'], kind: 'text' }
  ], function(row) {
    return row.date || row.amount != null;
  }).map(function(row) {
    var manual = compact_(row.type) !== '주문입금';
    return { row: row.row, date: row.date, amount: row.amount, memo: row.memo, counts: manual };
  });
}

function readExpense_(sh) {
  return readMapped_(sh, [
    { key: 'date', names: ['날짜'], kind: 'date' },
    { key: 'category', names: ['분류'], kind: 'text' },
    { key: 'item', names: ['항목'], kind: 'text' },
    { key: 'amount', names: ['금액'], kind: 'number' },
    { key: 'memo', names: ['메모'], kind: 'text' }
  ], function(row) {
    return row.date || row.item || row.category || row.amount != null;
  });
}

function readOrders_(sh) {
  return readMapped_(sh, [
    { key: 'date', names: ['날짜'], kind: 'date' },
    { key: 'status', names: ['상태'], kind: 'text' },
    { key: 'amount', names: ['금액'], kind: 'number' },
    { key: 'includeLabel', names: ['집계 포함', '집계포함'], kind: 'text' }
  ], function(row) {
    return row.date || row.amount != null || row.status;
  }).map(function(row) {
    return {
      row: row.row,
      date: row.date,
      status: row.status,
      amount: row.amount,
      included: orderIncluded_(row.status, row.includeLabel)
    };
  });
}

function readMapped_(sh, fields, keep) {
  var map = headerMap_(sh);
  var cols = [];
  var c;
  for (c = 0; c < fields.length; c++) cols.push(findHeader_(map, fields[c].names));
  var last = sh.getLastRow();
  if (last < 2) return [];
  var height = Math.min(last, MANUAL_LAST_ROW) - 1;
  var width = Math.max(sh.getLastColumn(), 1);
  var values = sh.getRange(2, 1, height, width).getValues();
  var out = [];
  var r;
  for (r = 0; r < values.length; r++) {
    var item = { row: r + 2 };
    for (c = 0; c < fields.length; c++) {
      var raw = cols[c] ? values[r][cols[c] - 1] : '';
      item[fields[c].key] = castCell_(raw, fields[c].kind);
    }
    if (keep(item)) out.push(item);
  }
  return out;
}

function castCell_(value, kind) {
  if (kind === 'number') return asNumber_(value);
  if (kind === 'date') return formatCellDate_(value);
  if (value == null) return '';
  return String(value).trim();
}

function readMeta_(sh) {
  var categories = sh ? readListColumn_(sh, 1) : [];
  if (!categories.length) categories = EXPENSE_CATEGORIES.slice();
  return { expenseCategories: categories };
}

function readListColumn_(sh, col) {
  var last = Math.min(Math.max(sh.getLastRow(), 1), 30);
  var values = sh.getRange(1, col, last, 1).getValues();
  var out = [];
  var i;
  for (i = 0; i < values.length; i++) {
    var v = String(values[i][0] == null ? '' : values[i][0]).trim();
    if (!v) continue;
    if (i === 0 && /^(분류|지출분류|수입구분|결제수단|구분)$/.test(compact_(v))) continue;
    if (out.indexOf(v) === -1) out.push(v);
  }
  return out;
}

function readSheetSummary_(sh) {
  if (!sh) return null;
  var map = headerMap_(sh);
  var monthCol = findHeader_(map, ['월']);
  if (!monthCol) return null;
  var last = Math.min(Math.max(sh.getLastRow(), 1), 40);
  var width = Math.max(sh.getLastColumn(), 1);
  var values = sh.getRange(1, 1, last, width).getValues();
  var headers = [];
  var i;
  for (i = 0; i < values[0].length; i++) headers.push(String(values[0][i] || '').trim());
  var months = [];
  var total = null;
  for (i = 1; i < values.length; i++) {
    var label = values[i][monthCol - 1];
    var labelText = String(label == null ? '' : label).trim();
    var month = parseMonthLabel_(label);
    var isTotal = compact_(labelText) === '합계';
    if (!month && !isTotal) continue;
    var row = summaryRowFromValues_(headers, values[i], month ? month.key : '합계', isTotal);
    if (isTotal) total = row;
    else months.push(row);
  }
  if (!months.length) return null;
  return { months: months, total: total };
}

function summaryRowFromValues_(headers, row, key, isTotal) {
  var expenses = {};
  var i;
  for (i = 0; i < headers.length; i++) {
    var name = headers[i];
    if (!name || SUMMARY_METRICS[name]) continue;
    if (!isHeaderLabel_(name)) continue;
    var n = asNumber_(row[i]);
    expenses[name] = n == null ? 0 : n;
  }
  var sales = numHeader_(headers, row, '판매');
  if (sales == null) {
    var orderIncome = numHeader_(headers, row, '주문 수입');
    var otherIncome = numHeader_(headers, row, '기타 수입');
    if (orderIncome != null || otherIncome != null) sales = (orderIncome || 0) + (otherIncome || 0);
  }
  var totalExpense = numHeader_(headers, row, '총지출');
  var profit = numHeader_(headers, row, '순이익');
  if (totalExpense == null) totalExpense = sumObj_(expenses);
  if (profit == null && sales != null && totalExpense != null) profit = sales - totalExpense;
  var margin = numHeader_(headers, row, '이익률');
  if (margin != null && Math.abs(margin) > 1.5) margin = margin / 100;
  return {
    month: key,
    label: isTotal ? '합계' : key,
    sales: sales,
    totalExpense: totalExpense,
    profit: profit,
    margin: margin == null ? marginOf_(profit, sales) : margin,
    expenses: expenses
  };
}

function numHeader_(headers, row, name) {
  var i;
  for (i = 0; i < headers.length; i++) {
    if (headers[i] === name || compact_(headers[i]) === compact_(name)) return asNumber_(row[i]);
  }
  return null;
}

function summaryHasNumbers_(summary) {
  if (!summary || !summary.months) return false;
  var rows = summary.months.slice();
  if (summary.total) rows.push(summary.total);
  var i;
  for (i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (r.sales != null || r.totalExpense != null || r.profit != null) return true;
  }
  return false;
}

function computeSummary_(income, expense, orders) {
  var keys = fiscalMonthKeys_();
  var seen = {};
  var i;
  for (i = 0; i < keys.length; i++) seen[keys[i]] = true;
  function touch(date) {
    if (!date || String(date).length < 7) return;
    var key = String(date).slice(0, 7);
    if (!seen[key]) { seen[key] = true; keys.push(key); }
  }
  income.forEach(function(row) { if (row.counts) touch(row.date); });
  expense.forEach(function(row) { touch(row.date); });
  orders.forEach(function(row) { touch(row.date); });
  keys.sort();
  var months = keys.map(function(key) { return blankMonth_(key); });
  var byKey = {};
  months.forEach(function(m) { byKey[m.month] = m; });
  orders.forEach(function(row) {
    var bucket = byKey[String(row.date || '').slice(0, 7)];
    if (!row.included || row.amount == null || !bucket) return;
    bucket.sales += row.amount;
  });
  income.forEach(function(row) {
    var bucket = byKey[String(row.date || '').slice(0, 7)];
    if (!row.counts || row.amount == null || !bucket) return;
    bucket.sales += row.amount;
  });
  expense.forEach(function(row) {
    var bucket = byKey[String(row.date || '').slice(0, 7)];
    if (row.amount == null || !bucket) return;
    var cat = canonicalCategory_(row.category);
    if (!cat) return;
    bucket.expenses[cat] += row.amount;
    bucket.totalExpense += row.amount;
  });
  months.forEach(finishMonth_);
  return { months: months, total: totalFromMonths_(months) };
}

function blankMonth_(key) {
  var expenses = {};
  EXPENSE_CATEGORIES.forEach(function(cat) { expenses[cat] = 0; });
  return {
    month: key,
    label: key,
    sales: 0,
    totalExpense: 0,
    profit: 0,
    margin: null,
    expenses: expenses
  };
}

function finishMonth_(row) {
  if (row.totalExpense == null) row.totalExpense = sumObj_(row.expenses);
  row.profit = (row.sales || 0) - row.totalExpense;
  row.margin = marginOf_(row.profit, row.sales);
}

function totalFromMonths_(months) {
  var total = blankMonth_('합계');
  total.label = '합계';
  months.forEach(function(m) {
    total.sales += m.sales || 0;
    total.totalExpense += m.totalExpense || 0;
    EXPENSE_CATEGORIES.forEach(function(cat) {
      total.expenses[cat] += (m.expenses && m.expenses[cat]) || 0;
    });
    Object.keys(m.expenses || {}).forEach(function(cat) {
      if (EXPENSE_CATEGORIES.indexOf(cat) !== -1) return;
      if (total.expenses[cat] == null) total.expenses[cat] = 0;
      total.expenses[cat] += m.expenses[cat] || 0;
    });
  });
  finishMonth_(total);
  total.month = '합계';
  return total;
}

function fiscalMonthKeys_() {
  var out = [];
  var y = 2026;
  var m = 7;
  var i;
  for (i = 0; i < 12; i++) {
    out.push(y + '-' + (m < 10 ? '0' + m : String(m)));
    m++;
    if (m === 13) { m = 1; y++; }
  }
  return out;
}

function orderIncluded_(status, includeLabel) {
  if (/취소|환불/.test(String(status || ''))) return false;
  var label = String(includeLabel == null ? '' : includeLabel).trim();
  if (!label) return true;
  return includeState_(label);
}

function includeState_(value) {
  if (value === true) return true;
  if (value === false) return false;
  var s = String(value == null ? '' : value).trim().toLowerCase();
  if (!s) return false;
  if (/^(제외|아니오|아니요|취소|환불|n|no|false|0)$/.test(s)) return false;
  if (/^(포함|예|y|yes|true|1|o|○|◯)$/.test(s)) return true;
  return false;
}

function canonicalCategory_(name) {
  var raw = String(name == null ? '' : name).trim();
  var c = compact_(raw);
  if (!c) return '';
  if (c.toUpperCase() === 'AI') return 'AI';
  if (c === '제본' || c === '광고' || c === '박스') return c;
  if (c.indexOf('제본') !== -1 || c.indexOf('제작') !== -1 || c.indexOf('구매') !== -1) return '제본';
  if (c.indexOf('광고') !== -1) return '광고';
  if (c.indexOf('박스') !== -1 || c.indexOf('포장') !== -1 || c.indexOf('택배') !== -1) return '박스';
  if (c.indexOf('구독') !== -1 || /ai/i.test(raw)) return 'AI';
  return '';
}

// ── 쓰기 ─────────────────────────────────────────────

function writeRecord_(sh, map, row, sheetName, record) {
  record = record || {};
  rejectBlockedKeys_(record);
  rejectDroppedFields_(record, sheetName);
  if (sheetName === '지출') writeExpense_(sh, map, row, record);
  else writeIncome_(sh, map, row, record);
  clearPersonalOnRow_(sh, row, headerMap_(sh));
}

function rejectDroppedFields_(record, sheetName) {
  var blocked = ['거래처', 'vendor', '결제수단', '결제 수단', 'method', '내용', 'detail', '교재', '권수', '구분', 'type'];
  if (sheetName === '지출') blocked = ['거래처', 'vendor', '결제수단', '결제 수단', 'method'];
  var i;
  for (i = 0; i < blocked.length; i++) {
    if (!hasField_(record, [blocked[i]])) continue;
    var value = record[blocked[i]];
    if (value == null || String(value).trim() === '') continue;
    if (sheetName === '수입') fail_('수입은 날짜, 금액, 메모만 저장합니다.');
    fail_('거래처와 결제수단은 저장하지 않습니다.');
  }
}

function writeExpense_(sh, map, row, record) {
  var date = parseDateInput_(firstField_(record, ['날짜', 'date']));
  var amount = parseAmount_(firstField_(record, ['금액', 'amount']));
  if (amount == null) fail_('금액을 입력하세요.');
  var category = canonicalCategory_(firstField_(record, ['분류', 'category']));
  if (!category) fail_('분류는 제본, AI, 광고, 박스 중 하나여야 합니다.');
  var pairs = [
    { names: ['날짜'], kind: 'date', value: date, required: true },
    { names: ['분류'], kind: 'text', value: category, required: true },
    { names: ['금액'], kind: 'number', value: amount, required: true },
    { names: ['메모'], kind: 'text', value: cleanText_(firstField_(record, ['메모', 'memo']), 2000) }
  ];
  if (hasField_(record, ['항목', 'item'])) {
    pairs.push({ names: ['항목'], kind: 'text', value: cleanText_(firstField_(record, ['항목', 'item']), 200) });
  }
  writeMapped_(sh, row, map, pairs);
}

function writeIncome_(sh, map, row, record) {
  var date = parseDateInput_(firstField_(record, ['날짜', 'date']));
  var amount = parseAmount_(firstField_(record, ['금액', 'amount']));
  if (amount == null) fail_('금액을 입력하세요.');
  writeMapped_(sh, row, map, [
    { names: ['날짜'], kind: 'date', value: date, required: true },
    { names: ['금액'], kind: 'number', value: amount, required: true },
    { names: ['메모'], kind: 'text', value: cleanText_(firstField_(record, ['메모', 'memo']), 2000) }
  ]);
}

function writeMapped_(sh, row, map, pairs) {
  var i;
  for (i = 0; i < pairs.length; i++) {
    var col = findHeader_(map, pairs[i].names);
    if (!col) {
      if (pairs[i].required) fail_(pairs[i].names[0] + ' 열을 찾지 못했습니다. 헤더 이름을 확인해 주세요.');
      continue;
    }
    var header = headerAt_(sh, col);
    if (isPersonalHeader_(header)) continue;
    var cell = sh.getRange(row, col);
    if (pairs[i].kind === 'date') {
      cell.setValue(pairs[i].value);
      cell.setNumberFormat('yyyy-mm-dd');
    } else if (pairs[i].kind === 'number') {
      cell.setValue(pairs[i].value);
      cell.setNumberFormat('#,##0');
    } else {
      cell.setValue(pairs[i].value);
    }
  }
}

function assertRowMatch_(sh, map, row, match) {
  if (!match) fail_('수정·삭제에는 match.date 와 match.amount 가 필요합니다.');
  var dateCol = findHeader_(map, ['날짜']);
  var amountCol = findHeader_(map, ['금액']);
  if (!dateCol || !amountCol) fail_('날짜 또는 금액 열을 찾지 못했습니다.');
  var actualDate = formatCellDate_(sh.getRange(row, dateCol).getValue());
  var actualAmount = asNumber_(sh.getRange(row, amountCol).getValue());
  var expectDate = formatCellDate_(parseDateInput_(firstField_(match, ['date', '날짜'])));
  var expectAmount = parseAmount_(firstField_(match, ['amount', '금액']));
  if (!actualDate || actualDate !== expectDate || actualAmount == null || expectAmount == null || Math.abs(actualAmount - expectAmount) > 0.001) {
    fail_('행 내용이 일치하지 않습니다. 목록을 새로고침한 뒤 다시 시도하세요.');
  }
}

function nextDataRow_(sh, map) {
  var dateCol = findHeader_(map, ['날짜']);
  var amountCol = findHeader_(map, ['금액']);
  var last = Math.max(sh.getLastRow(), 1);
  if (last < 2) return 2;
  var height = Math.min(last, MANUAL_LAST_ROW) - 1;
  var dates = dateCol ? sh.getRange(2, dateCol, height, 1).getValues() : [];
  var amounts = amountCol ? sh.getRange(2, amountCol, height, 1).getValues() : [];
  var lastData = 1;
  var i;
  for (i = 0; i < height; i++) {
    var d = dates.length ? dates[i][0] : '';
    var a = amounts.length ? amounts[i][0] : '';
    var hasAmount = typeof a === 'number' || String(a).trim() !== '';
    if (String(d).trim() !== '' || hasAmount) lastData = i + 2;
  }
  return lastData + 1;
}

function clearPersonalOnRow_(sh, row, map) {
  Object.keys(map).forEach(function(name) {
    if (isPersonalHeader_(name)) sh.getRange(row, map[name]).clearContent();
  });
}

function rejectBlockedKeys_(record) {
  Object.keys(record || {}).forEach(function(key) {
    if (isPersonalHeader_(key)) fail_('개인정보 항목은 저장하지 않습니다.');
  });
}

function hasField_(record, keys) {
  var i;
  for (i = 0; i < keys.length; i++) {
    if (record && Object.prototype.hasOwnProperty.call(record, keys[i])) return true;
  }
  return false;
}

function firstField_(record, keys) {
  var i;
  for (i = 0; i < keys.length; i++) {
    if (record && Object.prototype.hasOwnProperty.call(record, keys[i]) && record[keys[i]] != null && record[keys[i]] !== '') {
      return record[keys[i]];
    }
  }
  return '';
}

function parseJsonField_(value, label) {
  try { return JSON.parse(value); }
  catch (e) { fail_(label + ' JSON이 올바르지 않습니다.'); }
}

function parseRow_(value) {
  var n = Number(value);
  if (!isFinite(n) || Math.floor(n) !== n || n < 2 || n > 20000) fail_('행 번호가 올바르지 않습니다.');
  return n;
}

function cleanText_(value, max) {
  var s = String(value == null ? '' : value).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '').trim();
  if (s.length > max) fail_('글이 너무 깁니다. ' + max + '자 이하로 적어 주세요.');
  return s;
}

function parseAmount_(value) {
  if (typeof value === 'number' && isFinite(value)) return value;
  var s = String(value == null ? '' : value).trim();
  if (!s) return null;
  var negative = /^\(/.test(s) || /^-/.test(s);
  var digits = s.replace(/[^\d.]/g, '');
  if (!digits) return null;
  var n = Number(digits);
  if (!isFinite(n)) return null;
  return negative ? -Math.abs(n) : n;
}

function parseDateInput_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.parseDate(Utilities.formatDate(value, sheetTz_(), 'yyyy-MM-dd'), sheetTz_(), 'yyyy-MM-dd');
  }
  var s = String(value == null ? '' : value).trim();
  var m = s.match(/^(\d{4})\s*[-./]\s*(\d{1,2})\s*[-./]\s*(\d{1,2})/);
  if (!m) fail_('날짜는 YYYY-MM-DD 형식이어야 합니다.');
  var y = Number(m[1]);
  var mo = Number(m[2]);
  var d = Number(m[3]);
  if (y < 2000 || y > 2100 || mo < 1 || mo > 12 || d < 1 || d > 31) fail_('날짜가 올바르지 않습니다.');
  var iso = m[1] + '-' + (mo < 10 ? '0' + mo : String(mo)) + '-' + (d < 10 ? '0' + d : String(d));
  var parsed = Utilities.parseDate(iso, sheetTz_(), 'yyyy-MM-dd');
  var back = Utilities.formatDate(parsed, sheetTz_(), 'yyyy-MM-dd');
  if (back !== iso) fail_('날짜가 올바르지 않습니다.');
  return parsed;
}

function formatCellDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return Utilities.formatDate(value, sheetTz_(), 'yyyy-MM-dd');
  var s = String(value == null ? '' : value).trim();
  if (!s) return '';
  var iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return iso[1] + '-' + iso[2] + '-' + iso[3];
  var dot = s.match(/^(\d{4})\s*[./]\s*(\d{1,2})\s*[./]\s*(\d{1,2})/);
  if (!dot) return '';
  return dot[1] + '-' + ('0' + dot[2]).slice(-2) + '-' + ('0' + dot[3]).slice(-2);
}

function parseMonthLabel_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    var key = Utilities.formatDate(value, sheetTz_(), 'yyyy-MM');
    return { key: key, y: Number(key.slice(0, 4)), m: Number(key.slice(5, 7)) };
  }
  var s = String(value == null ? '' : value).trim();
  var m = s.match(/^(\d{4})\s*[-./년]\s*(\d{1,2})/);
  if (!m) return null;
  var y = Number(m[1]);
  var mo = Number(m[2]);
  if (mo < 1 || mo > 12) return null;
  return { key: y + '-' + (mo < 10 ? '0' + mo : String(mo)), y: y, m: mo };
}

function asNumber_(value) {
  if (typeof value === 'number' && isFinite(value)) return value;
  if (value instanceof Date) return null;
  if (typeof value === 'boolean') return null;
  var s = String(value == null ? '' : value).trim();
  if (!s || s.charAt(0) === '#') return null;
  return parseAmount_(s);
}

function marginOf_(profit, sales) {
  if (profit == null || sales == null || Math.abs(sales) < 0.0001) return null;
  return profit / sales;
}

function sumObj_(obj) {
  var n = 0;
  Object.keys(obj || {}).forEach(function(k) { n += obj[k] || 0; });
  return n;
}

// ── 초기 설정 ────────────────────────────────────────

/**
 * 시트를 판매 / 제본 / AI / 광고 / 박스 구조로 맞추고, 비밀번호가 있으면 저장합니다.
 * 편집기에서 그냥 실행하면 비밀번호 창은 뜨지 않을 수 있습니다.
 * @param {string=} initialPassword
 * @param {boolean=} alreadyPrompted 메뉴에서 이미 물어봤으면 true
 */
function setup(initialPassword, alreadyPrompted) {
  var report = [];
  var error = null;
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) fail_('활성 스프레드시트가 없습니다. 황소 워크북 장부에서 확장 프로그램 > Apps Script 로 연 뒤 실행하세요.');
    detectFormulaSep_(ss);
    var ctx = captureOrderContext_(ss);
    migrateIncome_(mustSheet_(ss, '수입'), report);
    migrateExpense_(mustSheet_(ss, '지출'), report);
    migrateCategories_(ss, report);
    migrateOrders_(mustSheet_(ss, '주문 연동'), ctx, report);
    rebuildSummary_(ss, ctx, report);
    fixNotes_(ss, report);
    report.push('장부 구조를 판매와 지출 네 분류로 맞췄습니다.');
  } catch (err) {
    error = err;
    report.push('오류: ' + ((err && err.message) ? err.message : err));
  }
  report.push(storePassword_(initialPassword, !alreadyPrompted));
  Logger.log(report.join('\n'));
  if (error) throw new Error(report.join('\n'));
  return report.join('\n');
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('황소 장부')
    .addItem('초기 설정', 'setupFromMenu')
    .addItem('비밀번호 설정', 'setPasswordFromMenu')
    .addItem('개인정보 열 점검', 'auditLedger')
    .addToUi();
}

function setupFromMenu() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt('황소 워크북 장부', '웹앱 비밀번호를 입력하세요. 이미 설정돼 있으면 비워 두고 확인을 누르세요.', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  try {
    ui.alert(setup(String(res.getResponseText() || ''), true));
  } catch (err) {
    ui.alert(String(err && err.message ? err.message : err).slice(0, 1500));
  }
}

function setPasswordFromMenu() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt('비밀번호 설정', '웹앱에 사용할 비밀번호를 입력하세요.', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  ui.alert(storePassword_(String(res.getResponseText() || ''), false));
}

function auditLedger() {
  var ss = activeSs_();
  var lines = [];
  lines.push('스프레드시트: ' + ss.getName());
  ['수입', '지출', '주문 연동', '월별 요약', '분류'].forEach(function(name) {
    lines.push(ss.getSheetByName(name) ? (name + ' 시트 있음') : (name + ' 시트 없음'));
  });
  var order = ss.getSheetByName('주문 연동');
  if (order) {
    describePrivacy_(order, '주문 연동', lines);
    orderSheetHealth_(order).forEach(function(problem) { lines.push('주문 연동: ' + problem); });
  }
  var income = ss.getSheetByName('수입');
  if (income) describePrivacy_(income, '수입', lines);
  var expense = ss.getSheetByName('지출');
  if (expense) describePrivacy_(expense, '지출', lines);
  var summary = ss.getSheetByName('월별 요약');
  if (summary) {
    var formulas = summary.getDataRange().getFormulas();
    var broken = 0;
    formulas.forEach(function(row) {
      row.forEach(function(f) { if (f && f.indexOf('#REF!') !== -1) broken++; });
    });
    lines.push(broken ? ('월별 요약에 깨진 수식 ' + broken + '개') : '월별 요약에 깨진 수식 없음');
  }
  var secret = PropertiesService.getScriptProperties().getProperty(PASSWORD_KEY);
  lines.push(secret ? '비밀번호: 설정됨' : '비밀번호: 없음');
  var text = lines.join('\n');
  Logger.log(text);
  try { SpreadsheetApp.getUi().alert(text.slice(0, 1500)); } catch (e) {}
  return text;
}

function describePrivacy_(sh, title, lines) {
  var map = headerMap_(sh);
  Object.keys(map).forEach(function(name) {
    if (isPersonalHeader_(name)) lines.push('개인정보 열이 남아 있습니다: ' + title + ' / ' + name);
  });
  if (title !== '주문 연동') return;
  listImports_(sh).forEach(function(imp) {
    lines.push('가져오기 ' + indexToCol_(imp.col) + imp.row + ' → ' + (imp.sheetName || '') + ' ' + indexToCol_(imp.startCol) + ':' + indexToCol_(imp.endCol));
    if (imp.startCol <= 3 && imp.endCol >= 3) lines.push('가져오기 범위에 주문자 열이 포함되어 있습니다.');
    if (imp.startCol <= 4 && imp.endCol >= 4) lines.push('가져오기 범위에 연락처 열이 포함되어 있습니다.');
    if (imp.startCol <= 5 && imp.endCol >= 5) lines.push('가져오기 범위에 교재 열이 포함되어 있습니다.');
  });
}

function storePassword_(initialPassword, allowPrompt) {
  var props = PropertiesService.getScriptProperties();
  var existing = props.getProperty(PASSWORD_KEY) || '';
  var incoming = initialPassword == null ? '' : String(initialPassword).trim();
  if (!incoming && allowPrompt) {
    try {
      var ui = SpreadsheetApp.getUi();
      var res = ui.prompt('황소 워크북 장부', '웹앱 비밀번호를 입력하세요. 이미 설정돼 있으면 비워 두고 확인을 누르세요.', ui.ButtonSet.OK_CANCEL);
      if (res.getSelectedButton() === ui.Button.OK) incoming = String(res.getResponseText() || '').trim();
    } catch (e) {}
  }
  if (incoming) {
    if (incoming.length < 4) return '비밀번호는 4자 이상으로 설정하세요. 저장하지 않았습니다.';
    if (incoming.length > 200) return '비밀번호가 너무 깁니다. 저장하지 않았습니다.';
    props.setProperty(PASSWORD_KEY, incoming);
    return '비밀번호를 스크립트 속성에 저장했습니다.';
  }
  if (existing) return '기존 비밀번호를 그대로 사용합니다.';
  return '비밀번호가 없습니다. 프로젝트 설정 → 스크립트 속성에 LEDGER_PASSWORD 를 추가하거나, 스프레드시트 메뉴 황소 장부 → 비밀번호 설정을 실행하세요.';
}

function captureOrderContext_(ss) {
  var order = mustSheet_(ss, '주문 연동');
  var ctx = { sourceId: SOURCE_SPREADSHEET_ID, endRow: ORDER_LAST_ROW_DEFAULT };
  listImports_(order).forEach(function(imp) {
    if (imp.sourceId) ctx.sourceId = imp.sourceId;
    var endRow = Number(imp.r2);
    if (endRow > ctx.endRow && endRow <= 5000) ctx.endRow = endRow;
  });
  return ctx;
}

function migrateIncome_(sh, report) {
  var map = headerMap_(sh);
  if (incomeLayoutOk_(map)) {
    report.push('수입 시트는 이미 날짜·금액·메모입니다.');
    return;
  }
  var dateCol = findHeader_(map, ['날짜']);
  var amountCol = findHeader_(map, ['금액']);
  var memoCol = findHeader_(map, ['메모']);
  var typeCol = findHeader_(map, ['구분']);
  if (!dateCol || !amountCol) fail_('수입 시트에서 날짜 또는 금액 열을 찾지 못했습니다.');
  var last = Math.max(sh.getLastRow(), 1);
  var width = Math.max(sh.getLastColumn(), 1);
  var values = last >= 2 ? sh.getRange(2, 1, last - 1, width).getValues() : [];
  var kept = [];
  var dropped = 0;
  var r;
  for (r = 0; r < values.length; r++) {
    var type = typeCol ? values[r][typeCol - 1] : '';
    if (compact_(type) === '주문입금') { dropped++; continue; }
    var date = values[r][dateCol - 1];
    var amount = values[r][amountCol - 1];
    var memo = memoCol ? values[r][memoCol - 1] : '';
    if (!isFilledDate_(date) && asNumber_(amount) == null) continue;
    kept.push([date, amount, memo]);
  }
  var headerCols = Object.keys(map).map(function(name) { return map[name]; });
  var maxCol = 3;
  headerCols.forEach(function(col) { if (col > maxCol) maxCol = col; });
  sh.getRange(1, 1, Math.max(last, 1), maxCol).clearContent();
  sh.getRange(1, 1, 1, 3).setValues([['날짜', '금액', '메모']]);
  if (kept.length) {
    sh.getRange(2, 1, kept.length, 3).setValues(kept);
    sh.getRange(2, 1, kept.length, 1).setNumberFormat('yyyy-mm-dd');
    sh.getRange(2, 2, kept.length, 1).setNumberFormat('#,##0');
  }
  report.push('수입 시트를 날짜·금액·메모만 남겼습니다. 주문 입금 ' + dropped + '행은 주문과 겹치므로 지웠습니다.');
}

function incomeLayoutOk_(map) {
  var keys = Object.keys(map);
  if (keys.length !== 3) return false;
  return findHeader_(map, ['날짜']) === 1 && findHeader_(map, ['금액']) === 2 && findHeader_(map, ['메모']) === 3;
}

function isFilledDate_(value) {
  return value instanceof Date || !!formatCellDate_(value);
}

function migrateExpense_(sh, report) {
  remapExpenseCategories_(sh, report);
  var map = headerMap_(sh);
  var remove = [];
  Object.keys(map).forEach(function(name) {
    if (keptExpenseHeader_(name)) return;
    remove.push({ col: map[name], name: name });
  });
  remove.sort(function(a, b) { return b.col - a.col; });
  if (remove.length > 8) fail_('지출 시트에서 지울 열이 너무 많아 멈췄습니다.');
  remove.forEach(function(item) {
    sh.deleteColumn(item.col);
    report.push('지출 시트에서 [' + item.name + '] 열을 삭제했습니다.');
  });
  map = headerMap_(sh);
  if (!findHeader_(map, ['항목'])) {
    var catCol = findHeader_(map, ['분류']);
    if (!catCol) fail_('지출 시트에서 분류 열을 찾지 못했습니다.');
    sh.insertColumnAfter(catCol);
    sh.getRange(1, catCol + 1).setValue('항목');
    report.push('지출 시트에 항목 열을 추가했습니다.');
  }
  if (!remove.length) report.push('지출 시트 열은 날짜·분류·항목·금액·메모입니다.');
}

function keptExpenseHeader_(name) {
  var c = compact_(name);
  return c === '날짜' || c === '분류' || c === '항목' || c === '금액' || c === '메모';
}

function remapExpenseCategories_(sh, report) {
  var map = headerMap_(sh);
  var catCol = findHeader_(map, ['분류']);
  var memoCol = findHeader_(map, ['메모']);
  if (!catCol) return;
  var last = sh.getLastRow();
  if (last < 2) return;
  var height = Math.min(last, MANUAL_LAST_ROW) - 1;
  var cats = sh.getRange(2, catCol, height, 1).getValues();
  var memos = memoCol ? sh.getRange(2, memoCol, height, 1).getValues() : null;
  var changed = 0;
  var unmapped = [];
  var i;
  for (i = 0; i < cats.length; i++) {
    var prev = String(cats[i][0] == null ? '' : cats[i][0]).trim();
    if (!prev) continue;
    var next = canonicalCategory_(prev);
    if (next && next !== prev) {
      cats[i][0] = next;
      changed++;
    } else if (!next) {
      unmapped.push(prev);
      if (memos) {
        var memo = String(memos[i][0] == null ? '' : memos[i][0]);
        if (memo.indexOf('이전 분류:') === -1) memos[i][0] = (memo ? memo + ' ' : '') + '이전 분류: ' + prev;
      }
    }
  }
  if (changed) sh.getRange(2, catCol, height, 1).setValues(cats);
  if (memos && unmapped.length) sh.getRange(2, memoCol, height, 1).setValues(memos);
  if (changed) report.push('지출 분류 ' + changed + '행을 제본·AI·광고·박스로 옮겼습니다.');
  if (unmapped.length) {
    report.push('네 분류로 옮기지 못한 지출 ' + unmapped.length + '행이 있습니다. 분류를 제본, AI, 광고, 박스 중 하나로 바꿔 주세요. 예: ' + uniqueTexts_(unmapped).join(', '));
  }
}

function migrateCategories_(ss, report) {
  var sh = ss.getSheetByName('분류');
  if (!sh) sh = ss.insertSheet('분류');
  var rows = Math.min(Math.max(sh.getMaxRows(), 4), 30);
  sh.getRange(1, 1, rows, 3).clearContent();
  var values = EXPENSE_CATEGORIES.map(function(name) { return [name]; });
  sh.getRange(1, 1, values.length, 1).setValues(values);
  report.push('분류 시트를 제본, AI, 광고, 박스로 맞췄습니다.');
}

function migrateOrders_(sh, ctx, report) {
  var problems = orderSheetHealth_(sh);
  if (!problems.length) {
    report.push('주문 연동은 이미 제출일시·상태·금액만 가져옵니다.');
    return;
  }
  problems.forEach(function(problem) { report.push('주문 연동: ' + problem); });
  rebuildOrderSheet_(sh, ctx);
  var again = orderSheetHealth_(sh);
  if (again.length) fail_('주문 연동을 다시 만든 뒤에도 확인이 필요합니다. ' + again.join(' '));
  report.push('주문 연동을 제출일시·상태·금액만 가져오도록 다시 썼습니다. 교재와 주문자는 가져오지 않습니다.');
}

function rebuildSummary_(ss, ctx, report) {
  var sh = mustSheet_(ss, '월별 요약');
  var orderSh = mustSheet_(ss, '주문 연동');
  var incomeSh = mustSheet_(ss, '수입');
  var expenseSh = mustSheet_(ss, '지출');
  var letters = {
    orderDate: letterOf_(orderSh, ['날짜']),
    orderStatus: letterOf_(orderSh, ['상태']),
    orderAmount: letterOf_(orderSh, ['금액']),
    incomeDate: letterOf_(incomeSh, ['날짜']),
    incomeAmount: letterOf_(incomeSh, ['금액']),
    expenseDate: letterOf_(expenseSh, ['날짜']),
    expenseCategory: letterOf_(expenseSh, ['분류']),
    expenseAmount: letterOf_(expenseSh, ['금액'])
  };
  var months = readExistingMonths_(sh);
  var orderEnd = ctx.endRow || ORDER_LAST_ROW_DEFAULT;
  var incomeEnd = formulaLastRow_(incomeSh);
  var expenseEnd = formulaLastRow_(expenseSh);
  var width = Math.max(sh.getLastColumn(), SUMMARY_HEADERS.length);
  var clearRows = Math.min(Math.max(sh.getMaxRows(), 23), 40);
  sh.getRange(1, 1, clearRows, width).clearContent();
  sh.getRange(1, 1, 1, SUMMARY_HEADERS.length).setValues([SUMMARY_HEADERS]);
  sh.getRange(2, 1, months.length, 1).setNumberFormat('@');
  sh.getRange(2, 1, months.length, 1).setValues(months.map(function(key) { return [key]; }));
  var formulas = [];
  var i;
  for (i = 0; i < months.length; i++) {
    var month = parseMonthLabel_(months[i]);
    var row = i + 2;
    formulas.push([
      salesFormula_(letters, month, orderEnd, incomeEnd),
      categoryFormula_(letters, month, expenseEnd, '제본'),
      categoryFormula_(letters, month, expenseEnd, 'AI'),
      categoryFormula_(letters, month, expenseEnd, '광고'),
      categoryFormula_(letters, month, expenseEnd, '박스'),
      '=C' + row + '+D' + row + '+E' + row + '+F' + row,
      '=B' + row + '-G' + row,
      '=IF(B' + row + '=0,"",H' + row + '/B' + row + ')'
    ]);
  }
  sh.getRange(2, 2, formulas.length, 8).setFormulas(localizeGrid_(formulas));
  var totalRow = months.length + 2;
  sh.getRange(totalRow, 1).setValue('합계');
  var totals = [];
  var c;
  for (c = 2; c <= 8; c++) {
    var letter = indexToCol_(c);
    totals.push('=SUM(' + letter + '2:' + letter + (totalRow - 1) + ')');
  }
  totals.push('=IF(B' + totalRow + '=0,"",H' + totalRow + '/B' + totalRow + ')');
  sh.getRange(totalRow, 2, 1, 8).setFormulas(localizeGrid_([totals]));
  sh.getRange(2, 2, totalRow - 1, 7).setNumberFormat('#,##0');
  sh.getRange(2, 9, totalRow - 1, 1).setNumberFormat('0.0%');
  report.push('월별 요약을 판매, 제본, AI, 광고, 박스, 총지출, 순이익, 이익률로 다시 썼습니다.');
}

function readExistingMonths_(sh) {
  var keys = [];
  var last = Math.min(Math.max(sh.getLastRow(), 1), 20);
  var values = sh.getRange(1, 1, last, 1).getValues();
  var i;
  for (i = 1; i < values.length; i++) {
    var parsed = parseMonthLabel_(values[i][0]);
    if (!parsed) continue;
    if (compact_(String(values[i][0] == null ? '' : values[i][0])) === '합계') break;
    keys.push(parsed.key);
    if (keys.length === 12) break;
  }
  if (keys.length === 12) return keys;
  return fiscalMonthKeys_();
}

function formulaLastRow_(sh) {
  return Math.min(Math.max(sh.getMaxRows(), 2), MANUAL_LAST_ROW);
}

function letterOf_(sh, names) {
  var col = findHeader_(headerMap_(sh), names);
  if (!col) fail_(sh.getName() + ' 시트에서 ' + names[0] + ' 열을 찾지 못했습니다.');
  return indexToCol_(col);
}

function salesFormula_(letters, month, orderEnd, incomeEnd) {
  var orders = sumifsMonth_(
    rangeOf_('주문 연동', letters.orderAmount, orderEnd),
    monthCriteria_(rangeOf_('주문 연동', letters.orderDate, orderEnd), month).concat([
      [rangeOf_('주문 연동', letters.orderStatus, orderEnd), '"<>*취소*"'],
      [rangeOf_('주문 연동', letters.orderStatus, orderEnd), '"<>*환불*"']
    ])
  );
  var manual = sumifsMonth_(
    rangeOf_('수입', letters.incomeAmount, incomeEnd),
    monthCriteria_(rangeOf_('수입', letters.incomeDate, incomeEnd), month)
  );
  return orders + '+' + manual.replace(/^=/, '');
}

function categoryFormula_(letters, month, expenseEnd, category) {
  var criteria = monthCriteria_(rangeOf_('지출', letters.expenseDate, expenseEnd), month);
  criteria.push([rangeOf_('지출', letters.expenseCategory, expenseEnd), '"' + String(category).replace(/"/g, '""') + '"']);
  return sumifsMonth_(rangeOf_('지출', letters.expenseAmount, expenseEnd), criteria);
}

function sumifsMonth_(amountRange, criteria) {
  var parts = [amountRange];
  criteria.forEach(function(pair) {
    parts.push(pair[0]);
    parts.push(pair[1]);
  });
  return '=SUMIFS(' + parts.join(',') + ')';
}

function monthCriteria_(dateRange, month) {
  var end = nextMonth_(month.y, month.m);
  return [
    [dateRange, '">="&DATE(' + month.y + ',' + month.m + ',1)'],
    [dateRange, '"<"&DATE(' + end.y + ',' + end.m + ',1)']
  ];
}

function rangeOf_(sheetName, letter, end) {
  return quoteSheet_(sheetName) + '!' + letter + '2:' + letter + end;
}

function nextMonth_(y, m) {
  m += 1;
  if (m === 13) { m = 1; y += 1; }
  return { y: y, m: m };
}

function fixNotes_(ss, report) {
  var orderSh = mustSheet_(ss, '주문 연동');
  var incomeSh = mustSheet_(ss, '수입');
  var summary = ss.getSheetByName('월별 요약');
  var guide = columnGuide_();
  var cleared = 0;
  cleared += clearPersonalNotes_(orderSh, 1);
  cleared += clearPersonalNotes_(incomeSh, 1);
  if (summary) cleared += clearPersonalNotes_(summary, 16);
  orderSh.getRange(1, 8).setValue(guide);
  if (summary) summary.getRange(16, 1).setValue(guide).setWrap(true);
  report.push(cleared ? ('이전 안내 문구 ' + cleared + '곳을 지우고 새 안내를 넣었습니다.') : '안내 문구를 현재 구조로 넣었습니다.');
}

function columnGuide_() {
  return [
    DOC_MARKER,
    '판매는 주문 연동의 금액입니다. 취소·환불은 제외합니다.',
    '가져오는 값은 제출일시, 상태, 금액뿐입니다.',
    '수입 시트는 날짜, 금액, 메모입니다. 주문 연동에 없는 판매만 적습니다.',
    '같은 판매를 주문 연동과 수입 시트에 모두 적으면 두 번 합산됩니다.',
    '지출 분류는 제본, AI, 광고, 박스입니다.',
    '이익률은 순이익을 판매로 나눈 값입니다.'
  ].join('\n');
}

function clearPersonalNotes_(sh, minRow) {
  var map = headerMap_(sh);
  var dataCols = {};
  Object.keys(map).forEach(function(name) {
    if (!isPersonalHeader_(name)) dataCols[map[name]] = true;
  });
  var lastRow = Math.min(Math.max(sh.getLastRow(), minRow), 80);
  var lastCol = Math.max(sh.getLastColumn(), 1);
  if (lastRow < minRow) return 0;
  var range = sh.getRange(minRow, 1, lastRow - minRow + 1, lastCol);
  var values = range.getValues();
  var formulas = range.getFormulas();
  var changed = 0;
  var r, c;
  for (r = 0; r < values.length; r++) {
    for (c = 0; c < values[r].length; c++) {
      if (formulas[r][c]) continue;
      var sheetRow = minRow + r;
      if (sheetRow >= 2 && dataCols[c + 1] && sh.getName() !== '월별 요약') continue;
      var text = String(values[r][c] == null ? '' : values[r][c]);
      if (!shouldScrubDoc_(text)) continue;
      sh.getRange(sheetRow, c + 1).clearContent();
      changed++;
    }
  }
  return changed;
}

function shouldScrubDoc_(text) {
  var s = String(text || '').trim();
  if (!s) return false;
  if (s.indexOf(DOC_MARKER) === 0 && !PERSONAL_HEADER_RE.test(s)) return false;
  if (PERSONAL_HEADER_RE.test(s)) return true;
  if (/IMPORTRANGE/i.test(s)) return true;
  if (/[A-Z]{1,2}\s*열/.test(s) && /주문|금액|집계|교재|상태/.test(s) && s.length > 15) return true;
  return false;
}

function rebuildOrderSheet_(sh, ctx) {
  var endRow = ctx.endRow || ORDER_LAST_ROW_DEFAULT;
  ensureSize_(sh, endRow, 8);
  listImports_(sh).forEach(function(imp) {
    sh.getRange(imp.row, imp.col).clearContent();
  });
  var clearCols = Math.min(Math.max(sh.getLastColumn(), 12), 12);
  sh.getRange(1, 1, endRow, clearCols).clearContent();
  var lastCol = sh.getLastColumn();
  if (lastCol > 12) {
    var headers = sh.getRange(1, 13, 1, lastCol - 12).getValues()[0];
    var c;
    for (c = 0; c < headers.length; c++) {
      if (!obsoleteOrderHeader_(headers[c])) continue;
      sh.getRange(1, 13 + c, Math.min(sh.getMaxRows(), endRow), 1).clearContent();
    }
  }
  sh.getRange(1, 1, 1, 7).setValues([['날짜', '상태', '금액', '', '원본 제출일시', '원본 상태', '원본 금액']]);
  var id = ctx.sourceId || SOURCE_SPREADSHEET_ID;
  sh.getRange(2, 5).setFormula(localizeFormula_(importFormula_(id, 'A', 'B', endRow)));
  sh.getRange(2, 7).setFormula(localizeFormula_(importFormula_(id, 'F', 'F', endRow)));
  var height = endRow - 1;
  var dates = [];
  var statuses = [];
  var amounts = [];
  var r;
  for (r = 2; r <= endRow; r++) {
    dates.push([dateFormula_('E' + r)]);
    statuses.push([echoFormula_('F' + r)]);
    amounts.push([amountFormula_('G' + r)]);
  }
  sh.getRange(2, 1, height, 1).setFormulas(localizeGrid_(dates));
  sh.getRange(2, 2, height, 1).setFormulas(localizeGrid_(statuses));
  sh.getRange(2, 3, height, 1).setFormulas(localizeGrid_(amounts));
  sh.getRange(2, 1, height, 1).setNumberFormat('yyyy-mm-dd');
  sh.getRange(2, 3, height, 1).setNumberFormat('#,##0');
}

function importFormula_(id, c1, c2, endRow) {
  return '=IMPORTRANGE("' + id + '","\'' + SOURCE_SHEET_NAME + '\'!' + c1 + '2:' + c2 + endRow + '")';
}

function dateFormula_(cell) {
  return '=IF(' + cell + '="","",IFERROR(DATE(' +
    'VALUE(REGEXEXTRACT(' + cell + ',"(\\d{4})\\s*[.]\\s*\\d{1,2}\\s*[.]\\s*\\d{1,2}")),' +
    'VALUE(REGEXEXTRACT(' + cell + ',"\\d{4}\\s*[.]\\s*(\\d{1,2})\\s*[.]\\s*\\d{1,2}")),' +
    'VALUE(REGEXEXTRACT(' + cell + ',"\\d{4}\\s*[.]\\s*\\d{1,2}\\s*[.]\\s*(\\d{1,2})"))' +
    '),IFERROR(DATEVALUE(SUBSTITUTE(SUBSTITUTE(' + cell + ',". ","-"),".","-")),)))';
}

function echoFormula_(cell) {
  return '=IF(' + cell + '="","",' + cell + ')';
}

function amountFormula_(cell) {
  return '=IF(' + cell + '="","",IFERROR(VALUE(REGEXREPLACE(TO_TEXT(' + cell + '),"[^0-9.]",""))*1,))';
}

function orderSheetHealth_(sh) {
  var problems = [];
  var map = headerMap_(sh);
  ['날짜', '상태', '금액'].forEach(function(name) {
    if (!findHeader_(map, [name])) problems.push(name + ' 헤더가 없습니다.');
  });
  Object.keys(map).forEach(function(name) {
    if (obsoleteOrderHeader_(name) || isPersonalHeader_(name)) problems.push('쓰지 않는 헤더가 남아 있습니다: ' + name);
  });
  var imports = listImports_(sh);
  var dates = null;
  var amount = null;
  imports.forEach(function(imp) {
    if (imp.sheetName && imp.sheetName !== SOURCE_SHEET_NAME) return;
    if (imp.startCol <= 3 && imp.endCol >= 3) problems.push('가져오기에 주문자 열이 남아 있습니다.');
    if (imp.startCol <= 4 && imp.endCol >= 4) problems.push('가져오기에 연락처 열이 남아 있습니다.');
    if (imp.startCol <= 5 && imp.endCol >= 5) problems.push('가져오기에 교재 열이 남아 있습니다.');
    if (imp.startCol === 1 && imp.endCol === 2) dates = imp;
    else if (imp.startCol === 6 && imp.endCol === 6) amount = imp;
    else problems.push('가져오기 범위가 제출일시·상태 또는 금액이 아닙니다.');
  });
  if (!dates) problems.push('제출일시·상태 가져오기가 없습니다.');
  if (!amount) problems.push('금액 가져오기가 없습니다.');
  if (dates && amount) {
    problems = problems.concat(expectFormulaRef_(sh, findHeader_(map, ['날짜']), dates.col, '날짜'));
    problems = problems.concat(expectFormulaRef_(sh, findHeader_(map, ['상태']), dates.col + 1, '상태'));
    problems = problems.concat(expectFormulaRef_(sh, findHeader_(map, ['금액']), amount.col, '금액'));
  }
  return uniqueTexts_(problems);
}

function obsoleteOrderHeader_(name) {
  if (!isHeaderLabel_(name)) return false;
  var c = compact_(name);
  if (c === '날짜' || c === '상태' || c === '금액') return false;
  if (c === '원본제출일시' || c === '원본상태' || c === '원본금액') return false;
  return /교재|주문자|연락처|집계|구분|내용|거래처|결제/.test(c);
}

function expectFormulaRef_(sh, col, expectedCol, label) {
  if (!col) return [label + ' 열이 없습니다.'];
  var formula = sh.getRange(2, col).getFormula();
  if (!formula) return [label + ' 변환 수식이 없습니다.'];
  if (formula.indexOf('#REF!') !== -1) return [label + ' 변환 수식이 깨졌습니다.'];
  if (!formulaReferencesCol_(formula, expectedCol)) return [label + ' 수식이 원본 열을 가리키지 않습니다.'];
  return [];
}

function formulaReferencesCol_(formula, colIndex) {
  var letter = indexToCol_(colIndex);
  return new RegExp('(^|[^A-Z])\\$?' + letter + '\\$?\\d+', 'i').test(formula);
}

// ── 시트 도우미 ──────────────────────────────────────

function activeSs_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) fail_('활성 스프레드시트가 없습니다.');
  return ss;
}

function mustSheet_(ss, name) {
  var sh = ss.getSheetByName(name);
  if (!sh) fail_(name + ' 시트를 찾지 못했습니다.');
  return sh;
}

function headerMap_(sh) {
  var last = Math.max(sh.getLastColumn(), 1);
  var row = sh.getRange(1, 1, 1, last).getValues()[0];
  var map = {};
  var i;
  for (i = 0; i < row.length; i++) {
    var name = String(row[i] == null ? '' : row[i]).trim();
    if (!isHeaderLabel_(name)) continue;
    if (!map[name]) map[name] = i + 1;
  }
  return map;
}

function headerAt_(sh, col) {
  if (!col) return '';
  var name = String(sh.getRange(1, col).getValue() || '').trim();
  return isHeaderLabel_(name) ? name : '';
}

function findHeader_(map, names) {
  var keys = Object.keys(map);
  var i, k;
  for (i = 0; i < names.length; i++) if (map[names[i]]) return map[names[i]];
  for (i = 0; i < names.length; i++) {
    var want = compact_(names[i]);
    for (k = 0; k < keys.length; k++) if (compact_(keys[k]) === want) return map[keys[k]];
  }
  for (i = 0; i < names.length; i++) {
    if (names[i].length < 2) continue;
    for (k = 0; k < keys.length; k++) if (keys[k].indexOf(names[i]) === 0) return map[keys[k]];
  }
  return 0;
}

function isHeaderLabel_(value) {
  var s = String(value || '').trim();
  if (!s || s.length > 40) return false;
  if (/[\r\n]/.test(s)) return false;
  if (/입니다|합니다|하세요|가져옵|IMPORTRANGE|참고|주의|안내/.test(s)) return false;
  return true;
}

function isPersonalHeader_(name) {
  return PERSONAL_HEADER_RE.test(String(name || ''));
}

function listImports_(sh) {
  var maxRows = Math.min(sh.getMaxRows(), 30);
  var maxCols = Math.min(Math.max(sh.getLastColumn(), 1), 26);
  var formulas = sh.getRange(1, 1, maxRows, maxCols).getFormulas();
  var found = [];
  var r, c;
  for (r = 0; r < formulas.length; r++) {
    for (c = 0; c < formulas[r].length; c++) {
      var f = formulas[r][c];
      if (!f || f.indexOf('IMPORTRANGE') === -1) continue;
      var parsed = firstImport_(f);
      if (!parsed) continue;
      found.push({
        row: r + 1,
        col: c + 1,
        formula: f,
        sourceId: parsed.id,
        sheetName: parsed.range.sheet,
        startCol: colToIndex_(parsed.range.c1),
        endCol: colToIndex_(parsed.range.c2),
        r1: parsed.range.r1,
        r2: parsed.range.r2
      });
    }
  }
  return found;
}

function firstImport_(formula) {
  var re = /IMPORTRANGE\(\s*"((?:[^"\\]|\\.)*)"\s*,\s*"((?:[^"\\]|\\.)*)"\s*\)/i;
  var m = re.exec(formula);
  if (!m) return null;
  var range = parseA1Range_(m[2]);
  if (!range) return null;
  return { id: m[1], range: range };
}

function ensureSize_(sh, rows, cols) {
  if (sh.getMaxRows() < rows) sh.insertRowsAfter(sh.getMaxRows(), rows - sh.getMaxRows());
  if (sh.getMaxColumns() < cols) sh.insertColumnsAfter(sh.getMaxColumns(), cols - sh.getMaxColumns());
}

function sheetTz_() {
  try { return SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() || 'Asia/Seoul'; }
  catch (e) { return 'Asia/Seoul'; }
}

function detectFormulaSep_(ss) {
  formulaSep_ = ',';
  var sh = ss.getSheetByName('분류') || ss.getSheets()[0];
  var cell = sh.getRange(sh.getMaxRows(), sh.getMaxColumns());
  if (cell.getFormula() || String(cell.getValue() || '') !== '') return;
  try {
    cell.setFormula('=SUM(1,2)');
    SpreadsheetApp.flush();
    if (Number(cell.getValue()) !== 3) formulaSep_ = ';';
  } catch (e) {
    formulaSep_ = ';';
  } finally {
    cell.clearContent();
  }
}

function localizeFormula_(formula) {
  if (!formula || formulaSep_ !== ';') return formula;
  return replaceTopLevel_(formula, ',', ';');
}

function localizeGrid_(grid) {
  return grid.map(function(row) { return row.map(localizeFormula_); });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function fail_(message) {
  var err = new Error(message);
  err.ledgerSafe = true;
  throw err;
}

function compact_(s) {
  return String(s || '').replace(/\s+/g, '');
}

function quoteSheet_(name) {
  return "'" + String(name).replace(/'/g, "''") + "'";
}

function indexToCol_(n) {
  var s = '';
  var num = Number(n);
  while (num > 0) {
    var rem = (num - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    num = Math.floor((num - 1) / 26);
  }
  return s;
}

function colToIndex_(letters) {
  var n = 0;
  var s = String(letters || '').toUpperCase();
  var i;
  for (i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64);
  return n;
}

function parseA1Range_(ref) {
  var m = String(ref || '').trim().match(/^(?:(?:'((?:[^']|'')*)'|([^'!]+))!)?\$?([A-Za-z]{1,3})\$?(\d+)\s*:\s*\$?([A-Za-z]{1,3})\$?(\d+)$/);
  if (!m) return null;
  return {
    sheet: (m[1] || m[2] || '').replace(/''/g, "'"),
    c1: m[3].toUpperCase(),
    r1: m[4],
    c2: m[5].toUpperCase(),
    r2: m[6]
  };
}

function replaceTopLevel_(s, from, to) {
  var out = '';
  var depth = 0;
  var inStr = false;
  var i;
  for (i = 0; i < s.length; i++) {
    var ch = s.charAt(i);
    if (inStr) {
      out += ch;
      if (ch === '"') inStr = false;
      continue;
    }
    if (ch === '"') { inStr = true; out += ch; continue; }
    if (ch === '(') depth++;
    if (ch === ')') depth = Math.max(0, depth - 1);
    out += (ch === from && depth > 0) ? to : ch;
  }
  return out;
}

function uniqueTexts_(list) {
  var seen = {};
  var out = [];
  list.forEach(function(item) {
    var key = String(item);
    if (seen[key]) return;
    seen[key] = true;
    out.push(key);
  });
  return out;
}
