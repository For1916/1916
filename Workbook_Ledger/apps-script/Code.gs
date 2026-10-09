/**
 * 황소 워크북 장부 — Google Apps Script 웹 앱
 *
 * 이 파일 전체를 스프레드시트 [황소 워크북 장부]의
 * 확장 프로그램 > Apps Script 에 붙여 넣습니다.
 * 시트에 묶인 스크립트이므로 SpreadsheetApp.getActiveSpreadsheet() 를 사용합니다.
 *
 * 비밀번호는 코드에 적지 않습니다. 스크립트 속성 LEDGER_PASSWORD 에만 둡니다.
 * 웹 앱은 매 요청마다 그 값과 비교하고, 주문자·연락처·주소·유입 경로는 반환하지 않습니다.
 *
 * 처음 한 번: 스프레드시트를 새로고침한 뒤 메뉴 [황소 장부] > [초기 설정]
 * 또는 편집기에서 setup 함수를 실행합니다.
 */

var APP_VERSION = '1';
var PASSWORD_KEY = 'LEDGER_PASSWORD';
var SOURCE_SPREADSHEET_ID = '1s_QC5gRuU7E07WGZrBtbMS_S80YHexlPPD_qYVDqTpM';
var SOURCE_SHEET_NAME = '주문 관리';
var ORDER_LAST_ROW_DEFAULT = 1000;
var MANUAL_LAST_ROW = 5000;

var PERSONAL_HEADER_RE = /주문자|고객명|수취인|받는\s*분|받는분|연락처|전화|휴대폰|핸드폰|휴대전화|주소|우편번호|이메일|유입|모아폼|moaform|답변\s*id|응답\s*id|answer\s*id|submission/i;
var DOC_MARKER = '[장부 안내]';

var DEFAULT_EXPENSE_CATEGORIES = ['교재 제작비', '교재 구매비', '박스·포장재', '택배비', '광고비', '구독료(AI·툴)', '수수료', '기타'];
var SUMMARY_METRICS = { '월': 1, '주문 수입': 1, '기타 수입': 1, '총수입': 1, '총지출': 1, '순이익': 1, '이익률': 1 };

var SOURCE_FIELD_BY_COL = {
  1: '원본 제출일시',
  2: '원본 상태',
  3: '원본 주문자',
  4: '원본 연락처',
  5: '원본 교재',
  6: '원본 금액'
};

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
    summary = computeSummary_(income, expense, orders, meta.expenseCategories);
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
  var existing = readRowSnapshot_(sh, map, rowNum, sheetName);
  if (action === 'delete') {
    sh.deleteRow(rowNum);
    return { ok: true, version: APP_VERSION, sheet: sheetName, row: rowNum };
  }
  writeRecord_(sh, map, rowNum, sheetName, data.record || {}, existing);
  return { ok: true, version: APP_VERSION, sheet: sheetName, row: rowNum };
}

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) fail_('다른 저장이 진행 중입니다. 잠시 후 다시 시도하세요.');
  try { return fn(); }
  finally { lock.releaseLock(); }
}

// ── 읽기 (개인정보 열 제외) ──────────────────────────

function readIncome_(sh) {
  return readMapped_(sh, [
    { key: 'date', names: ['날짜'], kind: 'date' },
    { key: 'type', names: ['구분'], kind: 'text' },
    { key: 'detail', names: ['내용'], kind: 'text' },
    { key: 'amount', names: ['금액'], kind: 'number' },
    { key: 'memo', names: ['메모'], kind: 'text' }
  ], function(row) {
    return row.date || row.detail || row.amount != null;
  }).map(function(row) {
    row.counts = compact_(row.type) === '기타수입';
    return row;
  });
}

function readExpense_(sh) {
  return readMapped_(sh, [
    { key: 'date', names: ['날짜'], kind: 'date' },
    { key: 'category', names: ['분류'], kind: 'text' },
    { key: 'item', names: ['항목'], kind: 'text' },
    { key: 'vendor', names: ['거래처'], kind: 'text' },
    { key: 'amount', names: ['금액'], kind: 'number' },
    { key: 'method', names: ['결제수단', '결제 수단'], kind: 'text' },
    { key: 'memo', names: ['메모'], kind: 'text' }
  ], function(row) {
    return row.date || row.item || row.category || row.amount != null;
  });
}

function readOrders_(sh) {
  return readMapped_(sh, [
    { key: 'date', names: ['날짜'], kind: 'date' },
    { key: 'status', names: ['상태'], kind: 'text' },
    { key: 'book', names: ['교재'], kind: 'text' },
    { key: 'amount', names: ['금액'], kind: 'number' },
    { key: 'includeLabel', names: ['집계 포함', '집계포함'], kind: 'text' }
  ], function(row) {
    return row.date || row.book || row.amount != null || row.status;
  }).map(function(row) {
    row.included = includeState_(row.includeLabel);
    return row;
  });
}

function readMapped_(sh, fields, keep) {
  var map = headerMap_(sh);
  var cols = [];
  var c;
  for (c = 0; c < fields.length; c++) {
    var col = findHeader_(map, fields[c].names);
    cols.push(col);
  }
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
  var incomeTypes = sh ? readListColumn_(sh, 2) : [];
  var methods = sh ? readListColumn_(sh, 3) : [];
  if (!categories.length) categories = DEFAULT_EXPENSE_CATEGORIES.slice();
  if (!incomeTypes.length) incomeTypes = ['기타 수입', '주문 입금'];
  return { expenseCategories: categories, incomeTypes: incomeTypes, paymentMethods: methods };
}

function readListColumn_(sh, col) {
  var last = sh.getLastRow();
  if (last < 1) return [];
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
    if (!name || SUMMARY_METRICS[name] || SUMMARY_METRICS[compact_(name)]) continue;
    if (!isHeaderLabel_(name)) continue;
    var n = asNumber_(row[i]);
    expenses[name] = n == null ? 0 : n;
  }
  var orderIncome = numHeader_(headers, row, '주문 수입');
  var otherIncome = numHeader_(headers, row, '기타 수입');
  var totalIncome = numHeader_(headers, row, '총수입');
  var totalExpense = numHeader_(headers, row, '총지출');
  var profit = numHeader_(headers, row, '순이익');
  if (totalIncome == null && orderIncome != null && otherIncome != null) totalIncome = orderIncome + otherIncome;
  if (totalExpense == null) totalExpense = sumObj_(expenses);
  if (profit == null && totalIncome != null && totalExpense != null) profit = totalIncome - totalExpense;
  return {
    month: key,
    label: isTotal ? '합계' : key,
    orderIncome: orderIncome,
    otherIncome: otherIncome,
    totalIncome: totalIncome,
    totalExpense: totalExpense,
    profit: profit,
    margin: marginOf_(profit, totalIncome),
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
    if (r.orderIncome != null || r.otherIncome != null || r.totalExpense != null || r.totalIncome != null) return true;
  }
  return false;
}

function computeSummary_(income, expense, orders, categories) {
  var keys = fiscalMonthKeys_();
  var seen = {};
  var i;
  for (i = 0; i < keys.length; i++) seen[keys[i]] = true;
  function touch(date) {
    if (!date || String(date).length < 7) return;
    var key = String(date).slice(0, 7);
    if (!seen[key]) { seen[key] = true; keys.push(key); }
  }
  income.forEach(function(row) { touch(row.date); });
  expense.forEach(function(row) { touch(row.date); });
  orders.forEach(function(row) { touch(row.date); });
  keys.sort();
  var cats = categories && categories.length ? categories.slice() : DEFAULT_EXPENSE_CATEGORIES.slice();
  var months = keys.map(function(key) {
    return blankMonth_(key, cats);
  });
  var byKey = {};
  months.forEach(function(m) { byKey[m.month] = m; });
  orders.forEach(function(row) {
    if (!row.included || row.amount == null || !byKey[String(row.date).slice(0, 7)]) return;
    byKey[String(row.date).slice(0, 7)].orderIncome += row.amount;
  });
  income.forEach(function(row) {
    if (!row.counts || row.amount == null || !byKey[String(row.date).slice(0, 7)]) return;
    byKey[String(row.date).slice(0, 7)].otherIncome += row.amount;
  });
  expense.forEach(function(row) {
    if (row.amount == null || !byKey[String(row.date).slice(0, 7)]) return;
    var bucket = byKey[String(row.date).slice(0, 7)];
    var cat = row.category || '기타';
    if (bucket.expenses[cat] == null) bucket.expenses[cat] = 0;
    bucket.expenses[cat] += row.amount;
    bucket.totalExpense += row.amount;
  });
  months.forEach(finishMonth_);
  return { months: months, total: totalFromMonths_(months) };
}

function blankMonth_(key, cats) {
  var expenses = {};
  cats.forEach(function(cat) { expenses[cat] = 0; });
  return {
    month: key,
    label: key,
    orderIncome: 0,
    otherIncome: 0,
    totalIncome: 0,
    totalExpense: 0,
    profit: 0,
    margin: null,
    expenses: expenses
  };
}

function finishMonth_(row) {
  row.totalIncome = (row.orderIncome || 0) + (row.otherIncome || 0);
  if (row.totalExpense == null) row.totalExpense = sumObj_(row.expenses);
  row.profit = row.totalIncome - row.totalExpense;
  row.margin = marginOf_(row.profit, row.totalIncome);
}

function totalFromMonths_(months) {
  var cats = {};
  var total = blankMonth_('합계', []);
  total.label = '합계';
  total.orderIncome = 0;
  total.otherIncome = 0;
  total.totalExpense = 0;
  months.forEach(function(m) {
    total.orderIncome += m.orderIncome || 0;
    total.otherIncome += m.otherIncome || 0;
    total.totalExpense += m.totalExpense || 0;
    Object.keys(m.expenses || {}).forEach(function(cat) {
      if (cats[cat] == null) cats[cat] = 0;
      cats[cat] += m.expenses[cat] || 0;
    });
  });
  total.expenses = cats;
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

function includeState_(value) {
  if (value === true) return true;
  if (value === false) return false;
  var s = String(value == null ? '' : value).trim().toLowerCase();
  if (!s) return false;
  if (/^(제외|아니오|아니요|취소|환불|n|no|false|0)$/.test(s)) return false;
  if (/^(포함|예|y|yes|true|1|o|○|◯)$/.test(s)) return true;
  return false;
}

// ── 쓰기 ─────────────────────────────────────────────

function writeRecord_(sh, map, row, sheetName, record, existing) {
  record = record || {};
  rejectBlockedKeys_(record);
  if (sheetName === '지출') writeExpense_(sh, map, row, record, existing);
  else writeIncome_(sh, map, row, record, existing);
  clearPersonalOnRow_(sh, row, headerMap_(sh));
}

function writeExpense_(sh, map, row, record, existing) {
  var date = parseDateInput_(firstField_(record, ['날짜', 'date']));
  var amount = parseAmount_(firstField_(record, ['금액', 'amount']));
  if (amount == null) fail_('금액을 입력하세요.');
  var category = cleanText_(firstField_(record, ['분류', 'category']), 40);
  if (!category) fail_('분류를 선택하세요.');
  assertAllowed_(category, listOrEmpty_(sh, '분류', 1), '분류', existing && existing.category);
  var method = cleanText_(firstField_(record, ['결제수단', '결제 수단', 'method']), 40);
  if (method) assertAllowed_(method, listOrEmpty_(sh, '분류', 3), '결제수단', existing && existing.method);
  writeMapped_(sh, row, map, [
    { names: ['날짜'], kind: 'date', value: date, required: true },
    { names: ['분류'], kind: 'text', value: category, required: true },
    { names: ['항목'], kind: 'text', value: cleanText_(firstField_(record, ['항목', 'item']), 200) },
    { names: ['거래처'], kind: 'text', value: cleanText_(firstField_(record, ['거래처', 'vendor']), 200) },
    { names: ['금액'], kind: 'number', value: amount, required: true },
    { names: ['결제수단', '결제 수단'], kind: 'text', value: method },
    { names: ['메모'], kind: 'text', value: cleanText_(firstField_(record, ['메모', 'memo']), 2000) }
  ]);
}

function writeIncome_(sh, map, row, record, existing) {
  var date = parseDateInput_(firstField_(record, ['날짜', 'date']));
  var amount = parseAmount_(firstField_(record, ['금액', 'amount']));
  if (amount == null) fail_('금액을 입력하세요.');
  var type = '기타 수입';
  if (existing && compact_(existing.type) === '주문입금') type = existing.type || '주문 입금';
  var requested = cleanText_(firstField_(record, ['구분', 'type']), 30);
  if (requested) {
    if (compact_(requested) === '기타수입') type = '기타 수입';
    else if (existing && compact_(requested) === compact_(existing.type)) type = existing.type;
    else if (compact_(requested) !== '기타수입') fail_('새 수입은 기타 수입만 등록할 수 있습니다. 주문 수입은 주문 연동에서 집계됩니다.');
  }
  writeMapped_(sh, row, map, [
    { names: ['날짜'], kind: 'date', value: date, required: true },
    { names: ['구분'], kind: 'text', value: type, required: true },
    { names: ['내용'], kind: 'text', value: cleanText_(firstField_(record, ['내용', 'detail']), 200) },
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

function readRowSnapshot_(sh, map, row, sheetName) {
  function textAt(names) {
    var col = findHeader_(map, names);
    return col ? String(sh.getRange(row, col).getValue() || '').trim() : '';
  }
  if (sheetName === '수입') return { type: textAt(['구분']) };
  return { category: textAt(['분류']), method: textAt(['결제수단', '결제 수단']) };
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

function listOrEmpty_(contextSheet, classSheetName, col) {
  var ss = contextSheet.getParent();
  var sh = ss.getSheetByName(classSheetName);
  if (!sh) return [];
  return readListColumn_(sh, col);
}

function assertAllowed_(value, list, label, previous) {
  if (previous && value === previous) return;
  if (!list || !list.length) return;
  if (list.indexOf(value) === -1) fail_(label + ' 값이 분류 시트에 없습니다: ' + value);
}

function rejectBlockedKeys_(record) {
  Object.keys(record || {}).forEach(function(key) {
    if (isPersonalHeader_(key)) fail_('개인정보 항목은 저장하지 않습니다.');
  });
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
  var n = Number(s.replace(/[^\d.]/g, ''));
  if (!isFinite(n) || s.replace(/[^\d.]/g, '') === '') return null;
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

function marginOf_(profit, income) {
  if (profit == null || income == null || Math.abs(income) < 0.0001) return null;
  return profit / income;
}

function sumObj_(obj) {
  var n = 0;
  Object.keys(obj || {}).forEach(function(k) { n += obj[k] || 0; });
  return n;
}

// ── 초기 설정 ────────────────────────────────────────

/**
 * 개인정보 열을 정리하고, 비밀번호가 넘어오면 스크립트 속성에 저장합니다.
 * 편집기에서 그냥 실행하면 비밀번호 창은 뜨지 않을 수 있습니다.
 * 그 경우 메뉴 [황소 장부] > [비밀번호 설정] 또는 스크립트 속성 LEDGER_PASSWORD 를 사용하세요.
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
    migrateOrders_(mustSheet_(ss, '주문 연동'), ctx, report);
    fixSummary_(ss, ctx, report);
    fixNotes_(ss, report);
    report.push('개인정보 열 정리를 마쳤습니다.');
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
  if (order) describePrivacy_(order, '주문 연동', lines);
  var income = ss.getSheetByName('수입');
  if (income) describePrivacy_(income, '수입', lines);
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
  var imports = listImports_(order);
  var roleByLetter = letterRoleMap_(order);
  annotateImportRoles_(imports, roleByLetter);
  var includeCol = findHeader_(headerMap_(order), ['집계 포함', '집계포함']);
  var includeFormula = includeCol ? order.getRange(2, includeCol).getFormula() : '';
  var ctx = {
    roleByLetter: roleByLetter,
    includeFormula: includeFormula,
    tokenInclude: tokenizeFormula_(includeFormula, roleByLetter),
    criterion: '',
    referencedInclude: false,
    sourceId: SOURCE_SPREADSHEET_ID,
    endRow: ORDER_LAST_ROW_DEFAULT,
    oldMaps: {
      '주문 연동': roleByLetter,
      '수입': letterRoleMap_(mustSheet_(ss, '수입')),
      '지출': letterRoleMap_(mustSheet_(ss, '지출'))
    },
    summaryFormulas: []
  };
  imports.forEach(function(imp) {
    if (imp.sourceId) ctx.sourceId = imp.sourceId;
    var endRow = Number(imp.r2);
    if (endRow > ctx.endRow && endRow <= 5000) ctx.endRow = endRow;
  });
  var includeLetter = '';
  Object.keys(roleByLetter).forEach(function(letter) {
    if (compact_(roleByLetter[letter]) === '집계포함') includeLetter = letter;
  });
  var summary = ss.getSheetByName('월별 요약');
  if (summary) {
    var last = Math.min(Math.max(summary.getLastRow(), 1), 40);
    var width = Math.min(Math.max(summary.getLastColumn(), 1), 30);
    ctx.summaryFormulas = summary.getRange(1, 1, last, width).getFormulas();
    ctx.summaryValues = summary.getRange(1, 1, last, width).getValues();
    var sMap = headerMap_(summary);
    var orderCol = findHeader_(sMap, ['주문 수입']);
    if (orderCol && ctx.summaryFormulas.length > 1) {
      var sample = ctx.summaryFormulas[1][orderCol - 1] || '';
      ctx.criterion = extractIncludeCriterion_(sample, roleByLetter) || '';
      if (includeLetter && new RegExp('(^|[^A-Z])\\$?' + includeLetter + '(?![A-Z])').test(sample)) ctx.referencedInclude = true;
    }
  }
  return ctx;
}

function migrateIncome_(sh, report) {
  var map = headerMap_(sh);
  var cols = [];
  Object.keys(map).forEach(function(name) {
    if (isPersonalHeader_(name)) cols.push({ col: map[name], name: name });
  });
  cols.sort(function(a, b) { return b.col - a.col; });
  if (!cols.length) {
    report.push('수입 시트에는 삭제할 개인정보 열이 없습니다.');
    return;
  }
  if (cols.length > 6) fail_('수입 시트에서 개인정보로 보이는 열이 너무 많아 삭제를 멈췄습니다.');
  cols.forEach(function(item) {
    sh.deleteColumn(item.col);
    report.push('수입 시트에서 [' + item.name + '] 열을 삭제했습니다.');
  });
}

function migrateOrders_(sh, ctx, report) {
  var imports = listImports_(sh);
  var deleteCols = {};
  imports.forEach(function(imp) {
    var rewritten = rewriteImportFormula_(imp.formula, ctx.sourceId || SOURCE_SPREADSHEET_ID);
    if (!rewritten.changed) return;
    if (rewritten.formula) sh.getRange(imp.row, imp.col).setFormula(localizeFormula_(rewritten.formula));
    else sh.getRange(imp.row, imp.col).clearContent();
    rewritten.dropOffsets.forEach(function(offset) {
      deleteCols[imp.col + offset] = '가져오기에서 뺀 개인정보 열';
    });
    report.push(indexToCol_(imp.col) + imp.row + ' 가져오기 범위를 제출일시·상태 또는 교재·금액만 남기도록 고쳤습니다.');
  });
  var map = headerMap_(sh);
  Object.keys(map).forEach(function(name) {
    if (isPersonalHeader_(name)) deleteCols[map[name]] = name;
  });
  var cols = Object.keys(deleteCols).map(Number).sort(function(a, b) { return b - a; });
  if (cols.length > 8) fail_('주문 연동에서 삭제할 열이 너무 많아 멈췄습니다.');
  cols.forEach(function(col) {
    if (columnHasImport_(sh, col)) {
      report.push(indexToCol_(col) + '열에는 가져오기 수식이 있어 삭제하지 않고 값만 비웁니다.');
      clearDataRows_(sh, col);
      return;
    }
    sh.deleteColumn(col);
    report.push('주문 연동 ' + indexToCol_(col) + '열을 삭제했습니다 (' + deleteCols[col] + ').');
  });
  SpreadsheetApp.flush();
  clearNearImportOrphans_(sh, report);
  var problems = orderSheetHealth_(sh);
  if (problems.length) {
    report.push('주문 연동 변환 수식을 기준 형태로 다시 썼습니다.');
    problems.forEach(function(p) { report.push('- ' + p); });
    rebuildOrderSheet_(sh, ctx);
    clearNearImportOrphans_(sh, report);
  } else {
    report.push('주문 연동의 기존 변환 수식은 유지했습니다.');
  }
}

function clearNearImportOrphans_(sh, report) {
  var imports = listImports_(sh);
  var spill = spillCols_(imports);
  var map = headerMap_(sh);
  var headerCols = {};
  Object.keys(map).forEach(function(name) { headerCols[map[name]] = name; });
  var last = sh.getLastColumn();
  var toDelete = [];
  var toClear = [];
  var c;
  for (c = 1; c <= last; c++) {
    if (headerCols[c] && !isPersonalHeader_(headerCols[c])) continue;
    if (spill[c] || columnHasImport_(sh, c)) continue;
    if (!nearImport_(c, imports)) continue;
    if (!columnHasValues_(sh, c)) continue;
    if (headerCols[c] && isPersonalHeader_(headerCols[c])) toDelete.push({ col: c, name: headerCols[c] });
    else toClear.push(c);
  }
  toClear.forEach(function(col) {
    clearDataRows_(sh, col);
    report.push(indexToCol_(col) + '열에 남아 있던 원본 값을 비웠습니다.');
  });
  toDelete.sort(function(a, b) { return b.col - a.col; }).forEach(function(item) {
    sh.deleteColumn(item.col);
    report.push('남아 있던 개인정보 열 [' + item.name + '] 을 삭제했습니다.');
  });
}

function fixSummary_(ss, ctx, report) {
  var sh = mustSheet_(ss, '월별 요약');
  var orderSh = mustSheet_(ss, '주문 연동');
  var incomeSh = mustSheet_(ss, '수입');
  var expenseSh = mustSheet_(ss, '지출');
  var sheets = { '주문 연동': orderSh, '수입': incomeSh, '지출': expenseSh };
  var newMaps = {
    '주문 연동': invertRoleMap_(letterRoleMap_(orderSh)),
    '수입': invertRoleMap_(letterRoleMap_(incomeSh)),
    '지출': invertRoleMap_(letterRoleMap_(expenseSh))
  };
  var current = sh.getDataRange().getFormulas();
  var headers = headerList_(sh);
  var fixed = 0;
  var r, c;
  for (r = 0; r < current.length; r++) {
    for (c = 0; c < current[r].length; c++) {
      var formula = current[r][c];
      if (!formula) continue;
      if (!formulaNeedsRebuild_(formula, sheets)) continue;
      var snapshot = (ctx.summaryFormulas[r] && ctx.summaryFormulas[r][c]) || formula;
      var replacement = replacementSummaryFormula_(sh, r + 1, c + 1, headers, ctx, newMaps, snapshot, sheets);
      if (!replacement) {
        report.push('월별 요약 ' + indexToCol_(c + 1) + (r + 1) + ' 수식은 자동으로 고치지 못했습니다. 실행 로그를 확인해 주세요.');
        continue;
      }
      sh.getRange(r + 1, c + 1).setFormula(localizeFormula_(replacement));
      fixed++;
    }
  }
  report.push(fixed ? ('월별 요약 수식 ' + fixed + '개를 고쳤습니다.') : '월별 요약 수식은 열 삭제에 맞춰 유지됩니다.');
}

function replacementSummaryFormula_(sh, row, col, headers, ctx, newMaps, snapshot, sheets) {
  var header = headerAt_(sh, col);
  var values = ctx.summaryValues || [];
  var monthLabel = values[row - 1] ? values[row - 1][0] : sh.getRange(row, 1).getValue();
  var month = parseMonthLabel_(monthLabel);
  var isTotal = compact_(String(monthLabel == null ? '' : monthLabel)) === '합계';
  if (row >= 2 && row <= 13 && month && header) {
    var built = metricFormula_(header, month, headers, row, ctx, newMaps, snapshot);
    if (built) return built;
  }
  if ((isTotal || row === 14) && header && compact_(header) !== '월') {
    return '=SUM(' + indexToCol_(col) + '2:' + indexToCol_(col) + '13)';
  }
  var remapped = remapQualifiedRefs_(snapshot, ctx.oldMaps, newMaps);
  if (remapped && !formulaNeedsRebuild_(remapped, sheets)) return remapped;
  return '';
}

function metricFormula_(header, month, headers, row, ctx, newMaps, snapshot) {
  var order = newMaps['주문 연동'] || {};
  var income = newMaps['수입'] || {};
  var expense = newMaps['지출'] || {};
  if (header === '주문 수입') {
    var includeLetter = order['집계 포함'] || order['집계포함'];
    var oldInclude = includeLetterFrom_((ctx.oldMaps && ctx.oldMaps['주문 연동']) || {});
    var snapshotUsesInclude = oldInclude && new RegExp('(^|[^A-Z])\\$?' + oldInclude + '(?![A-Z])').test(snapshot);
    if (snapshotUsesInclude && !ctx.criterion) return '';
    return sumifsMonth_(quoteSheet_('주문 연동'), order['금액'], order['날짜'], month, snapshotUsesInclude ? includeLetter : '', snapshotUsesInclude ? (ctx.criterion || '"포함"') : '');
  }
  if (header === '기타 수입') {
    return sumifsMonth_(quoteSheet_('수입'), income['금액'], income['날짜'], month, income['구분'], '"기타 수입"');
  }
  if (header === '총수입') return localAdd_(headers, row, ['주문 수입', '기타 수입']);
  if (header === '총지출') return localAdd_(headers, row, expenseHeaders_(headers));
  if (header === '순이익') return localSub_(headers, row, '총수입', '총지출');
  if (!SUMMARY_METRICS[header]) {
    return sumifsMonth_(quoteSheet_('지출'), expense['금액'], expense['날짜'], month, expense['분류'], '"' + String(header).replace(/"/g, '""') + '"');
  }
  return '';
}

function sumifsMonth_(sheetQuoted, amountLetter, dateLetter, month, extraLetter, extraCrit) {
  if (!amountLetter || !dateLetter) return '';
  var amount = sheetQuoted + '!' + amountLetter + '2:' + amountLetter + (sheetQuoted.indexOf('주문') !== -1 ? '1000' : '5000');
  var dates = sheetQuoted + '!' + dateLetter + '2:' + dateLetter + (sheetQuoted.indexOf('주문') !== -1 ? '1000' : '5000');
  var end = nextMonth_(month.y, month.m);
  var formula = '=SUMIFS(' + amount + ',' + dates + ',">="&DATE(' + month.y + ',' + month.m + ',1),' + dates + ',"<"&DATE(' + end.y + ',' + end.m + ',1)';
  if (extraLetter && extraCrit) {
    var extra = sheetQuoted + '!' + extraLetter + '2:' + extraLetter + (sheetQuoted.indexOf('주문') !== -1 ? '1000' : '5000');
    formula += ',' + extra + ',' + extraCrit;
  }
  return formula + ')';
}

function localAdd_(headers, row, names) {
  var letters = [];
  var i;
  for (i = 0; i < names.length; i++) {
    var col = headerColFromList_(headers, names[i]);
    if (col) letters.push(indexToCol_(col) + row);
  }
  if (!letters.length) return '';
  return '=' + letters.join('+');
}

function localSub_(headers, row, plusName, minusName) {
  var a = headerColFromList_(headers, plusName);
  var b = headerColFromList_(headers, minusName);
  if (!a || !b) return '';
  return '=' + indexToCol_(a) + row + '-' + indexToCol_(b) + row;
}

function headerColFromList_(headers, name) {
  var i;
  for (i = 0; i < headers.length; i++) {
    if (headers[i] === name || compact_(headers[i]) === compact_(name)) return i + 1;
  }
  return 0;
}

function expenseHeaders_(headers) {
  var out = [];
  headers.forEach(function(name) {
    if (!name || SUMMARY_METRICS[name] || !isHeaderLabel_(name)) return;
    out.push(name);
  });
  return out;
}

function headerList_(sh) {
  var last = Math.max(sh.getLastColumn(), 1);
  var row = sh.getRange(1, 1, 1, last).getValues()[0];
  return row.map(function(h) { return String(h || '').trim(); });
}

function nextMonth_(y, m) {
  m += 1;
  if (m === 13) { m = 1; y += 1; }
  return { y: y, m: m };
}

function includeLetterFrom_(roleByLetter) {
  var letters = Object.keys(roleByLetter || {});
  var i;
  for (i = 0; i < letters.length; i++) {
    if (compact_(roleByLetter[letters[i]]) === '집계포함') return letters[i];
  }
  return '';
}

function formulaNeedsRebuild_(formula, sheets) {
  if (!formula) return false;
  if (formula.indexOf('#REF!') !== -1) return true;
  if (/INDIRECT/i.test(formula)) return true;
  if (PERSONAL_HEADER_RE.test(formula)) return true;
  var refs = extractSheetRefs_(formula);
  var i;
  for (i = 0; i < refs.length; i++) {
    var ref = refs[i];
    var sh = sheets[ref.sheet];
    if (!sh) continue;
    var header = headerAt_(sh, ref.col);
    if (ref.sheet === '주문 연동' && !allowedOrderHeader_(header)) return true;
    if (ref.sheet === '수입' && !allowedIncomeHeader_(header)) return true;
    if (ref.sheet === '지출' && !allowedExpenseHeader_(header)) return true;
  }
  return false;
}

function allowedOrderHeader_(header) {
  if (!header) return false;
  if (isPersonalHeader_(header)) return false;
  var c = compact_(header);
  return c === '날짜' || c === '상태' || c === '교재' || c === '금액' || c === '집계포함';
}

function allowedIncomeHeader_(header) {
  if (!header || isPersonalHeader_(header)) return false;
  var c = compact_(header);
  return c === '날짜' || c === '구분' || c.indexOf('내용') === 0 || c === '금액' || c === '메모';
}

function allowedExpenseHeader_(header) {
  if (!header || isPersonalHeader_(header)) return false;
  var c = compact_(header);
  return c === '날짜' || c === '분류' || c === '항목' || c === '거래처' || c === '금액' || c === '결제수단' || c === '메모';
}

function extractSheetRefs_(formula) {
  var refs = [];
  var re = /(?:'((?:[^']|'')*)'|([^\s'!]+))!\$?([A-Z]{1,3})(?![A-Z])(?:\$?\d*)?(?::\$?([A-Z]{1,3})(?![A-Z]))?/g;
  var m;
  while ((m = re.exec(formula))) {
    var sheet = (m[1] || m[2] || '').replace(/''/g, "'");
    refs.push({ sheet: sheet, col: colToIndex_(m[3]) });
    if (m[4]) refs.push({ sheet: sheet, col: colToIndex_(m[4]) });
  }
  return refs;
}

function mapRef_(sheet, letter, dol1, dol2, row, oldMaps, newMaps) {
  var role = oldMaps[sheet] && oldMaps[sheet][String(letter).toUpperCase()];
  if (!role || isPersonalHeader_(role)) return { bad: true };
  var neu = newMaps[sheet] && (newMaps[sheet][role] || newMaps[sheet][compact_(role)]);
  if (!neu) return { cell: (dol1 || '') + letter + (dol2 || '') + (row || '') };
  return { cell: (dol1 || '') + neu + (dol2 || '') + (row || '') };
}

function remapQualifiedRefs_(formula, oldMaps, newMaps) {
  if (!formula) return '';
  var names = Object.keys(oldMaps || {}).sort(function(a, b) { return b.length - a.length; });
  var out = formula;
  names.forEach(function(name) {
    var quoted = name.replace(/'/g, "''");
    var re = new RegExp("'" + quoted + "'!(\\$?)([A-Z]{1,3})(\\$?)(\\d*)(?::(\\$?)([A-Z]{1,3})(\\$?)(\\d*))?", 'g');
    out = out.replace(re, function(full, dol1, letter1, dol2, row1, dol3, letter2, dol4, row2) {
      var first = mapRef_(name, letter1, dol1, dol2, row1, oldMaps, newMaps);
      if (first.bad) return '#REF!';
      var head = "'" + quoted + "'!" + first.cell;
      if (!letter2) return head;
      var second = mapRef_(name, letter2, dol3, dol4, row2, oldMaps, newMaps);
      if (second.bad) return '#REF!';
      return head + ':' + second.cell;
    });
    if (name.indexOf(' ') === -1) {
      var plain = new RegExp("(^|[^A-Za-z0-9_'])" + name + "!(\\$?)([A-Z]{1,3})(\\$?)(\\d*)(?::(\\$?)([A-Z]{1,3})(\\$?)(\\d*))?", 'g');
      out = out.replace(plain, function(full, pre, dol1, letter1, dol2, row1, dol3, letter2, dol4, row2) {
        var first = mapRef_(name, letter1, dol1, dol2, row1, oldMaps, newMaps);
        if (first.bad) return pre + '#REF!';
        var head = pre + name + '!' + first.cell;
        if (!letter2) return head;
        var second = mapRef_(name, letter2, dol3, dol4, row2, oldMaps, newMaps);
        if (second.bad) return pre + '#REF!';
        return head + ':' + second.cell;
      });
    }
  });
  return out;
}

function fixNotes_(ss, report) {
  var orderSh = mustSheet_(ss, '주문 연동');
  var incomeSh = mustSheet_(ss, '수입');
  var summary = ss.getSheetByName('월별 요약');
  var guide = columnGuide_(orderSh, incomeSh);
  var changed = 0;
  changed += scrubDocSheet_(orderSh, guide, 1);
  changed += scrubDocSheet_(incomeSh, guide, 1);
  if (summary) changed += scrubDocSheet_(summary, guide, 16);
  report.push(changed ? ('안내 문구 ' + changed + '곳을 개인정보가 없는 설명으로 고쳤습니다.') : '고칠 안내 문구는 없습니다.');
}

function columnGuide_(orderSh, incomeSh) {
  function describe(sh, title) {
    var map = headerMap_(sh);
    var parts = [];
    Object.keys(map).sort(function(a, b) { return map[a] - map[b]; }).forEach(function(name) {
      if (isPersonalHeader_(name)) return;
      parts.push(indexToCol_(map[name]) + ' ' + name);
    });
    return title + ': ' + parts.join(', ');
  }
  return [
    DOC_MARKER,
    describe(orderSh, '주문 연동'),
    '원본 가져오기는 제출일시, 상태, 교재, 금액만 사용합니다.',
    describe(incomeSh, '수입'),
    '월별 요약의 주문 수입은 주문 연동에서 집계에 포함하는 금액만 더합니다.',
    '기타 수입은 수입 시트에서 구분이 기타 수입인 행만 더합니다.'
  ].join('\n');
}

function scrubDocSheet_(sh, guide, minRow) {
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
  var placed = false;
  var r, c;
  for (r = 0; r < values.length; r++) {
    for (c = 0; c < values[r].length; c++) {
      if (formulas[r][c]) continue;
      var sheetRow = minRow + r;
      if (sheetRow >= 2 && dataCols[c + 1] && sh.getName() !== '월별 요약') continue;
      var text = String(values[r][c] == null ? '' : values[r][c]);
      if (!shouldScrubDoc_(text)) continue;
      var cell = sh.getRange(sheetRow, c + 1);
      if (!placed) {
        cell.setValue(guide);
        placed = true;
      } else if (text.length > 40 || /IMPORTRANGE|[A-Z]\s*열/.test(text)) {
        cell.clearContent();
      } else {
        var shortText = removePersonalLines_(text);
        cell.setValue(shortText || '검산');
      }
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

function removePersonalLines_(text) {
  return String(text || '').split(/\n/).filter(function(line) {
    return !PERSONAL_HEADER_RE.test(line);
  }).join('\n').trim();
}

function rebuildOrderSheet_(sh, ctx) {
  var endRow = ctx.endRow || ORDER_LAST_ROW_DEFAULT;
  ensureSize_(sh, endRow, 10);
  sh.getRange(1, 1, endRow, 10).clearContent();
  sh.getRange(1, 1, 1, 10).setValues([[
    '날짜', '상태', '교재', '금액', '집계 포함', '', '원본 제출일시', '원본 상태', '원본 교재', '원본 금액'
  ]]);
  var id = ctx.sourceId || SOURCE_SPREADSHEET_ID;
  sh.getRange(2, 7).setFormula(localizeFormula_(importFormula_(id, 'A', 'B', endRow)));
  sh.getRange(2, 9).setFormula(localizeFormula_(importFormula_(id, 'E', 'F', endRow)));
  var height = endRow - 1;
  var dateF = [];
  var statusF = [];
  var bookF = [];
  var amtF = [];
  var incF = [];
  var r;
  for (r = 2; r <= endRow; r++) {
    dateF.push([dateFormula_('G' + r)]);
    statusF.push([echoFormula_('H' + r)]);
    bookF.push([echoFormula_('I' + r)]);
    amtF.push([amountFormula_('J' + r)]);
    incF.push([includeFormulaForRow_(ctx.tokenInclude, r)]);
  }
  sh.getRange(2, 1, height, 1).setFormulas(localizeGrid_(dateF));
  sh.getRange(2, 2, height, 1).setFormulas(localizeGrid_(statusF));
  sh.getRange(2, 3, height, 1).setFormulas(localizeGrid_(bookF));
  sh.getRange(2, 4, height, 1).setFormulas(localizeGrid_(amtF));
  sh.getRange(2, 5, height, 1).setFormulas(localizeGrid_(incF));
  sh.getRange(2, 1, height, 1).setNumberFormat('yyyy-mm-dd');
  sh.getRange(2, 4, height, 1).setNumberFormat('#,##0');
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

function defaultIncludeFormula_(row) {
  return '=IF(D' + row + '="","",IF(REGEXMATCH(B' + row + '&"","취소|환불|미입금"),"제외","포함"))';
}

function includeFormulaForRow_(template, row) {
  if (!template || template.indexOf('{{') === -1) return defaultIncludeFormula_(row);
  var map = {
    '날짜': 'A' + row,
    '상태': 'B' + row,
    '교재': 'C' + row,
    '금액': 'D' + row,
    '집계 포함': 'E' + row,
    '집계포함': 'E' + row,
    '원본 제출일시': 'G' + row,
    '원본 상태': 'H' + row,
    '원본 교재': 'I' + row,
    '원본 금액': 'J' + row
  };
  var out = template.replace(/\{\{([^}]+)\}\}/g, function(_, role) {
    return map[role] || map[compact_(role)] || '""';
  });
  if (out.indexOf('#REF!') !== -1 || PERSONAL_HEADER_RE.test(out)) return defaultIncludeFormula_(row);
  return out;
}

function orderSheetHealth_(sh) {
  var problems = [];
  var map = headerMap_(sh);
  ['날짜', '상태', '교재', '금액'].forEach(function(name) {
    if (!findHeader_(map, [name])) problems.push(name + ' 헤더가 없습니다.');
  });
  if (!findHeader_(map, ['집계 포함', '집계포함'])) problems.push('집계 포함 헤더가 없습니다.');
  Object.keys(map).forEach(function(name) {
    if (isPersonalHeader_(name)) problems.push('개인정보 헤더가 남아 있습니다: ' + name);
  });
  var imports = listImports_(sh);
  var ab = null;
  var ef = null;
  imports.forEach(function(imp) {
    if ((imp.sheetName || SOURCE_SHEET_NAME) !== SOURCE_SHEET_NAME && imp.sheetName) return;
    if (imp.startCol <= 3 && imp.endCol >= 3) problems.push('가져오기 범위에 주문자 열이 남아 있습니다.');
    if (imp.startCol <= 4 && imp.endCol >= 4) problems.push('가져오기 범위에 연락처 열이 남아 있습니다.');
    if (imp.startCol === 1 && imp.endCol === 2) ab = imp;
    if (imp.startCol === 5 && imp.endCol === 6) ef = imp;
  });
  if (!ab) problems.push('제출일시·상태 가져오기가 없습니다.');
  if (!ef) problems.push('교재·금액 가져오기가 없습니다.');
  if (ab && ef && !problems.length) {
    problems = problems.concat(expectFormulaRef_(sh, findHeader_(map, ['날짜']), ab.col, '날짜'));
    problems = problems.concat(expectFormulaRef_(sh, findHeader_(map, ['상태']), ab.col + 1, '상태'));
    problems = problems.concat(expectFormulaRef_(sh, findHeader_(map, ['교재']), ef.col, '교재'));
    problems = problems.concat(expectFormulaRef_(sh, findHeader_(map, ['금액']), ef.col + 1, '금액'));
    var includeCol = findHeader_(map, ['집계 포함', '집계포함']);
    if (includeCol) {
      var f = sh.getRange(2, includeCol).getFormula();
      var v = sh.getRange(2, includeCol).getValue();
      if ((!f && (v === '' || v == null)) || (f && f.indexOf('#REF!') !== -1)) problems.push('집계 포함 수식이 비어 있거나 깨졌습니다.');
    }
  }
  if (ab && ef) {
    var spill = spillCols_(imports);
    var headerCols = {};
    Object.keys(map).forEach(function(name) { headerCols[map[name]] = true; });
    var last = Math.min(sh.getLastColumn(), ef.col + 3);
    var c;
    for (c = Math.min(ab.col, ef.col); c <= last; c++) {
      if (headerCols[c] || spill[c] || columnHasImport_(sh, c)) continue;
      if (columnHasValues_(sh, c)) problems.push(indexToCol_(c) + '열에 원본으로 보이지 않는 값이 남아 있습니다.');
    }
  }
  return problems;
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

function letterRoleMap_(sh) {
  var map = headerMap_(sh);
  var out = {};
  Object.keys(map).forEach(function(name) { out[indexToCol_(map[name])] = name; });
  return out;
}

function invertRoleMap_(roleByLetter) {
  var out = {};
  Object.keys(roleByLetter).forEach(function(letter) {
    out[roleByLetter[letter]] = letter;
    out[compact_(roleByLetter[letter])] = letter;
  });
  return out;
}

function annotateImportRoles_(imports, roleByLetter) {
  imports.forEach(function(imp) {
    if (imp.sheetName && imp.sheetName !== SOURCE_SHEET_NAME) return;
    var src;
    for (src = imp.startCol; src <= imp.endCol; src++) {
      var letter = indexToCol_(imp.col + (src - imp.startCol));
      if (!roleByLetter[letter]) roleByLetter[letter] = SOURCE_FIELD_BY_COL[src] || ('원본' + src);
    }
  });
  return roleByLetter;
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

function spillCols_(imports) {
  var spill = {};
  imports.forEach(function(imp) {
    var width = imp.endCol - imp.startCol + 1;
    var i;
    for (i = 0; i < width; i++) spill[imp.col + i] = true;
  });
  return spill;
}

function nearImport_(col, imports) {
  var i;
  for (i = 0; i < imports.length; i++) {
    var imp = imports[i];
    var width = Math.max(imp.endCol - imp.startCol + 1, 1);
    if (col >= imp.col && col <= imp.col + width + 2) return true;
  }
  return false;
}

function columnHasImport_(sh, col) {
  var height = Math.min(sh.getMaxRows(), 30);
  var formulas = sh.getRange(1, col, height, 1).getFormulas();
  var i;
  for (i = 0; i < formulas.length; i++) {
    if (formulas[i][0] && String(formulas[i][0]).indexOf('IMPORTRANGE') !== -1) return true;
  }
  return false;
}

function columnHasValues_(sh, col) {
  var height = Math.min(Math.max(sh.getLastRow(), 2), 1000) - 1;
  if (height < 1) return false;
  var values = sh.getRange(2, col, height, 1).getValues();
  var i;
  for (i = 0; i < values.length; i++) {
    if (String(values[i][0] == null ? '' : values[i][0]).trim() !== '') return true;
  }
  return false;
}

function clearDataRows_(sh, col) {
  var height = Math.min(sh.getMaxRows(), 1000) - 1;
  if (height > 0) sh.getRange(2, col, height, 1).clearContent();
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

function rewriteImportFormula_(formula, sourceId) {
  var dropOffsets = [];
  var changed = false;
  if (!formula || String(formula).indexOf('IMPORTRANGE') === -1) {
    return { formula: formula || '', changed: false, dropOffsets: [] };
  }
  var re = /IMPORTRANGE\(\s*"((?:[^"\\]|\\.)*)"\s*,\s*"((?:[^"\\]|\\.)*)"\s*\)/gi;
  var newFormula = String(formula).replace(re, function(full, id, ref) {
    var parsed = parseA1Range_(ref);
    if (!parsed) return full;
    if (parsed.sheet && parsed.sheet !== SOURCE_SHEET_NAME) return full;
    var start = colToIndex_(parsed.c1);
    var end = colToIndex_(parsed.c2);
    if (end < start) return full;
    var personal = [];
    var c;
    for (c = start; c <= end; c++) if (c === 3 || c === 4) personal.push(c - start);
    if (!personal.length) return full;
    changed = true;
    personal.forEach(function(offset) { dropOffsets.push(offset); });
    var useId = id || sourceId || SOURCE_SPREADSHEET_ID;
    if (start >= 3 && end <= 4) return '';
    var newStart = start;
    var newEnd = end;
    if (start <= 2) newEnd = Math.min(end, 2);
    else if (start < 5) return '';
    if (newEnd < newStart) return '';
    var sheetPart = "'" + (parsed.sheet || SOURCE_SHEET_NAME).replace(/'/g, "''") + "'!";
    return 'IMPORTRANGE("' + useId + '","' + sheetPart + indexToCol_(newStart) + parsed.r1 + ':' + indexToCol_(newEnd) + parsed.r2 + '")';
  });
  if (changed && newFormula.indexOf('IMPORTRANGE') === -1) newFormula = '';
  return { formula: newFormula, changed: changed, dropOffsets: uniqueNums_(dropOffsets) };
}

function extractIncludeCriterion_(formula, roleByLetter) {
  if (!formula) return '';
  var includeLetters = [];
  Object.keys(roleByLetter || {}).forEach(function(letter) {
    if (compact_(roleByLetter[letter]) === '집계포함') includeLetters.push(letter);
  });
  if (!includeLetters.length) return '';
  var body = formula.match(/SUMIFS\s*\(([\s\S]*)\)\s*$/i);
  if (!body) return '';
  var sep = argSepOf_(body[1]);
  var args = splitTopLevel_(body[1], sep);
  var i;
  for (i = 1; i + 1 < args.length; i += 2) {
    var letterMatch = args[i].match(/!\$?([A-Z]{1,3})(?![A-Z])/i);
    if (!letterMatch) continue;
    if (includeLetters.indexOf(letterMatch[1].toUpperCase()) !== -1) return args[i + 1];
  }
  return '';
}

function tokenizeFormula_(formula, roleByLetter) {
  if (!formula) return '';
  return String(formula).replace(/(^|[^A-Z$])(\$?)([A-Z]{1,3})(\$?)(\d+)/g, function(full, pre, dol1, col, dol2, row) {
    var role = roleByLetter[String(col).toUpperCase()];
    if (!role) return full;
    if (isPersonalHeader_(role)) return pre + '""';
    return pre + '{{' + role + '}}';
  });
}

function argSepOf_(formula) {
  var inStr = false;
  var semis = 0;
  var commas = 0;
  var i;
  for (i = 0; i < formula.length; i++) {
    var ch = formula.charAt(i);
    if (ch === '"') { inStr = !inStr; continue; }
    if (inStr) continue;
    if (ch === ';') semis++;
    if (ch === ',') commas++;
  }
  return semis > commas ? ';' : ',';
}

function splitTopLevel_(s, sep) {
  var args = [];
  var cur = '';
  var depth = 0;
  var inStr = false;
  var i;
  for (i = 0; i < s.length; i++) {
    var ch = s.charAt(i);
    if (inStr) {
      cur += ch;
      if (ch === '"') inStr = false;
      continue;
    }
    if (ch === '"') { inStr = true; cur += ch; continue; }
    if (ch === '(') { depth++; cur += ch; continue; }
    if (ch === ')') { depth = Math.max(0, depth - 1); cur += ch; continue; }
    if (ch === sep && depth === 0) { args.push(cur.trim()); cur = ''; continue; }
    cur += ch;
  }
  if (cur.trim()) args.push(cur.trim());
  return args;
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

function uniqueNums_(list) {
  var seen = {};
  var out = [];
  list.forEach(function(n) {
    if (seen[n]) return;
    seen[n] = true;
    out.push(n);
  });
  return out;
}
