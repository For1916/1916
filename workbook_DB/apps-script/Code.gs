/**
 * 황소 워크북 장부 — Google Apps Script 웹 앱
 *
 * 시트 [황소 워크북 장부] 의 확장 프로그램 > Apps Script 에 이 파일 전체를 붙여 넣습니다.
 * 시트에 묶인 스크립트이므로 SpreadsheetApp.getActiveSpreadsheet() 를 사용합니다.
 *
 * 입금은 주문 연동의 금액(취소·환불 제외)과, 수입 시트에 직접 적은 금액입니다. 월별 요약 헤더는 입금입니다.
 * 월별 요약의 달은 시트에 YY-MM(26-07)으로 보이고, 누적 수입·누적 지출·누적 순이익·누적 수익률을 더합니다.
 * 수입 시트는 날짜·금액·메모만 둡니다. 주문에 이미 있는 입금을 여기에 다시 적으면 두 번 합산됩니다.
 * 지출 분류는 제본, AI, 광고, 박스 네 가지입니다.
 * 재고는 같은 스프레드시트의 재고 품목·입출고·현재 재고·주문 출고 탭에서 따로 계산합니다.
 * 이익용 주문 연동에는 교재를 넣지 않고, 재고용 주문 출고만 교재 글자를 가져옵니다.
 * 주문 출고는 주문일·입고일·출고일을 따로 둡니다. 재고 입고는 입고일, 재고 출고는 출고일로만 움직입니다.
 * 출고일이 비어 있으면 아직 발송하지 않은 주문이므로 재고를 빼지 않습니다. 주문일을 출고일에 복사하지 않습니다.
 *
 * 비밀번호는 코드에 적지 않습니다. 스크립트 속성 LEDGER_PASSWORD 에만 둡니다.
 * 웹 앱은 매 요청마다 그 값과 비교하고, 주문자·연락처·주소·유입 경로는 반환하지 않습니다.
 *
 * 처음 한 번: 스프레드시트를 새로고침한 뒤 메뉴 [황소 장부] > [초기 설정]
 * 또는 편집기에서 setup 함수를 실행합니다. 다시 실행해도 됩니다.
 */

var APP_VERSION = '4';
var PASSWORD_KEY = 'LEDGER_PASSWORD';
var STOCK_RULES_KEY = 'STOCK_RULES_INITIALIZED';
var SOURCE_SPREADSHEET_ID = '1s_QC5gRuU7E07WGZrBtbMS_S80YHexlPPD_qYVDqTpM';
var SOURCE_SHEET_NAME = '주문 관리';
var ORDER_LAST_ROW_DEFAULT = 1000;
var MANUAL_LAST_ROW = 5000;

var PERSONAL_HEADER_RE = /주문자|고객명|수취인|받는\s*분|받는분|연락처|전화|휴대폰|핸드폰|휴대전화|주소|우편번호|이메일|유입|모아폼|moaform|답변\s*id|응답\s*id|answer\s*id|submission/i;
var DOC_MARKER = '[장부 안내]';
var EXPENSE_CATEGORIES = ['제본', 'AI', '광고', '박스'];
var SUMMARY_FLOOR = '2026-06';
var SUMMARY_HEADERS = ['월', '입금', '제본', 'AI', '광고', '박스', '총지출', '순이익', '이익률', '누적 수입', '누적 지출', '누적 순이익', '누적 수익률'];
var SUMMARY_METRICS = {
  '월': 1, '판매': 1, '입금': 1, '주문 수입': 1, '기타 수입': 1, '총수입': 1, '총지출': 1, '순이익': 1, '이익률': 1,
  '누적 수입': 1, '누적 판매': 1, '누적 입금': 1, '누적 지출': 1, '누적 순이익': 1, '누적 수익률': 1, '누적 이익률': 1
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
    if (action === 'stock') return json_(withLockMs_(20000, function() { return loadStock_(); }));
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
    summary = alignSummary_(summary, income, expense, orders);
    summary.source = 'sheet';
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
  if (sheetName === '입출고' || sheetName === '재고 품목' || sheetName === '박스규칙' || sheetName === '주문 출고') {
    return mutateStock_(action, sheetName, data);
  }
  if (sheetName !== '수입' && sheetName !== '지출') fail_('수입, 지출, 재고만 수정할 수 있습니다.');
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
  return withLockMs_(15000, fn);
}

function withLockMs_(waitMs, fn) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(waitMs)) fail_('다른 저장이 진행 중입니다. 잠시 후 다시 시도하세요.');
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
  return withCumulative_({ months: months, total: total });
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
  var sales = numHeader_(headers, row, '입금');
  if (sales == null) sales = numHeader_(headers, row, '판매');
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
  var dates = [];
  income.forEach(function(row) { dates.push(row.date); });
  expense.forEach(function(row) { dates.push(row.date); });
  orders.forEach(function(row) { dates.push(row.date); });
  var keys = summaryMonthKeys_(dates);
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
  return withCumulative_({ months: months, total: totalFromMonths_(months) });
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
  total.cumSales = total.sales;
  total.cumExpense = total.totalExpense;
  total.cumProfit = total.profit;
  total.cumMargin = total.margin;
  return total;
}

function applyCumulative_(months) {
  var sales = 0;
  var expense = 0;
  (months || []).forEach(function(row) {
    sales += Number(row.sales) || 0;
    expense += Number(row.totalExpense) || 0;
    row.cumSales = sales;
    row.cumExpense = expense;
    row.cumProfit = sales - expense;
    row.cumMargin = marginOf_(row.cumProfit, sales);
  });
  return months;
}

function withCumulative_(summary) {
  if (!summary) return summary;
  applyCumulative_(summary.months || []);
  if (summary.total) {
    summary.total.cumSales = summary.total.sales;
    summary.total.cumExpense = summary.total.totalExpense;
    summary.total.cumProfit = summary.total.profit;
    summary.total.cumMargin = marginOf_(summary.total.profit, summary.total.sales);
  }
  return summary;
}

function alignSummary_(summary, income, expense, orders) {
  var computed = computeSummary_(income, expense, orders);
  var byKey = {};
  (summary.months || []).forEach(function(row) {
    if (row && row.month) byKey[row.month] = row;
  });
  var added = false;
  var months = computed.months.map(function(row) {
    if (byKey[row.month]) return byKey[row.month];
    added = true;
    return row;
  });
  var same = !added && (summary.months || []).length === months.length;
  return withCumulative_({
    months: months,
    total: same && summary.total ? summary.total : totalFromMonths_(months)
  });
}

function summaryMonthKeys_(dates, todayKey) {
  var keys = [];
  var i;
  for (i = 0; i < (dates || []).length; i++) {
    var key = monthKeyFromCell_(dates[i]);
    if (key) keys.push(key);
  }
  return summaryMonthKeysFromList_(keys, todayKey);
}

function summaryMonthKeysFromList_(keys, todayKey) {
  var today = todayKey || seoulMonthKey_();
  if (!/^\d{4}-\d{2}$/.test(today)) today = '2026-10';
  var start = SUMMARY_FLOOR;
  var end = addMonthsKey_(today, 12);
  var i;
  for (i = 0; i < (keys || []).length; i++) {
    var key = keys[i];
    if (!/^\d{4}-\d{2}$/.test(key)) continue;
    if (key < start) start = key;
    if (key > end) end = key;
  }
  return enumerateMonthKeys_(start, end);
}

function enumerateMonthKeys_(start, end) {
  var out = [];
  var cursor = start;
  var guard = 0;
  while (cursor && cursor <= end && guard < 400) {
    out.push(cursor);
    cursor = addMonthsKey_(cursor, 1);
    guard++;
  }
  return out;
}

function addMonthsKey_(key, delta) {
  var y = Number(String(key).slice(0, 4));
  var m = Number(String(key).slice(5, 7));
  if (!y || m < 1 || m > 12) return '';
  var index = y * 12 + (m - 1) + delta;
  var ny = Math.floor(index / 12);
  var nm = index % 12;
  return ny + '-' + (nm + 1 < 10 ? '0' + (nm + 1) : String(nm + 1));
}

function seoulMonthKey_() {
  try {
    return Utilities.formatDate(new Date(), sheetTz_(), 'yyyy-MM');
  } catch (e) {
    return '2026-10';
  }
}

function monthKeyFromCell_(value) {
  var iso = formatCellDate_(value);
  var key = '';
  if (iso && iso.length >= 7) key = iso.slice(0, 7);
  else {
    var parsed = parseMonthLabel_(value);
    if (parsed) key = parsed.key;
  }
  if (!/^\d{4}-\d{2}$/.test(key)) return '';
  var y = Number(key.slice(0, 4));
  if (y < 2020 || y > 2045) return '';
  return key;
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
  fitColumns_(sh);
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
      if (pairs[i].plain) cell.setNumberFormat('@');
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
  if (typeof value === 'number' && isFinite(value) && value >= 30000 && value < 80000) return sheetsSerialToIso_(value);
  if (value instanceof Date && !isNaN(value.getTime())) return Utilities.formatDate(value, sheetTz_(), 'yyyy-MM-dd');
  var s = String(value == null ? '' : value).trim();
  if (!s) return '';
  var iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return iso[1] + '-' + iso[2] + '-' + iso[3];
  var dot = s.match(/^(\d{4})\s*[./]\s*(\d{1,2})\s*[./]\s*(\d{1,2})/);
  if (!dot) return '';
  return dot[1] + '-' + ('0' + dot[2]).slice(-2) + '-' + ('0' + dot[3]).slice(-2);
}

function sheetsSerialToIso_(serial) {
  var utc = Date.UTC(1899, 11, 30) + Math.round(Number(serial)) * 86400000;
  var d = new Date(utc);
  var y = d.getUTCFullYear();
  var m = d.getUTCMonth() + 1;
  var day = d.getUTCDate();
  return y + '-' + (m < 10 ? '0' + m : String(m)) + '-' + (day < 10 ? '0' + day : String(day));
}

function parseMonthLabel_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    var key = Utilities.formatDate(value, sheetTz_(), 'yyyy-MM');
    return { key: key, y: Number(key.slice(0, 4)), m: Number(key.slice(5, 7)) };
  }
  var s = String(value == null ? '' : value).trim();
  var m = s.match(/^(\d{4})\s*[-./년]\s*(\d{1,2})/);
  if (m) {
    var y4 = Number(m[1]);
    var mo4 = Number(m[2]);
    if (mo4 < 1 || mo4 > 12) return null;
    return { key: y4 + '-' + (mo4 < 10 ? '0' + mo4 : String(mo4)), y: y4, m: mo4 };
  }
  var yy = s.match(/^(\d{2})\s*[-./]\s*(\d{1,2})$/);
  if (!yy) return null;
  var y = 2000 + Number(yy[1]);
  var mo = Number(yy[2]);
  if (mo < 1 || mo > 12) return null;
  return { key: y + '-' + (mo < 10 ? '0' + mo : String(mo)), y: y, m: mo };
}

function formatYyMm_(key) {
  var parsed = parseMonthLabel_(key);
  if (!parsed) return String(key || '');
  var mm = parsed.m < 10 ? '0' + parsed.m : String(parsed.m);
  return String(parsed.y).slice(-2) + '-' + mm;
}

function monthKeyToDate_(key) {
  var parsed = parseMonthLabel_(key);
  if (!parsed) return '';
  var mm = parsed.m < 10 ? '0' + parsed.m : String(parsed.m);
  return Utilities.parseDate(parsed.y + '-' + mm + '-01', sheetTz_(), 'yyyy-MM-dd');
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
 * 시트를 입금 / 제본 / AI / 광고 / 박스 구조로 맞추고, 비밀번호가 있으면 저장합니다.
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
    migrateInventory_(ss, ctx, report);
    fitLedgerColumns_(ss);
    report.push('장부 구조를 입금과 지출 네 분류로 맞췄습니다. 월별 요약 열 너비를 내용에 맞췄습니다.');
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
    .addItem('재고 새로고침', 'refreshStockFromMenu')
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
  ['수입', '지출', '주문 연동', '월별 요약', '분류', '재고 품목', '입출고', '현재 재고', '주문 출고'].forEach(function(name) {
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
  ['재고 품목', '입출고', '현재 재고', '주문 출고'].forEach(function(name) {
    var sh = ss.getSheetByName(name);
    if (sh) describePrivacy_(sh, name, lines);
  });
  var ship = ss.getSheetByName('주문 출고');
  if (ship) shipSheetHealth_(ship).forEach(function(problem) { lines.push('주문 출고: ' + problem); });
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
  if (problems.length) {
    problems.forEach(function(problem) { report.push('주문 연동: ' + problem); });
    rebuildOrderSheet_(sh, ctx);
    var again = orderSheetHealth_(sh);
    if (again.length) fail_('주문 연동을 다시 만든 뒤에도 확인이 필요합니다. ' + again.join(' '));
    report.push('주문 연동을 제출일시·상태·금액만 가져오도록 다시 썼습니다. 교재와 주문자는 가져오지 않습니다.');
    return;
  }
  if (refreshOrderConversions_(sh, ctx)) {
    report.push('주문 연동 날짜·금액 변환을 고쳤습니다. 날짜 숫자와 2026. 7. 13 같은 글자, 금액 숫자와 40,000원 같은 글자를 모두 받습니다.');
  } else {
    report.push('주문 연동은 이미 제출일시·상태·금액만 가져옵니다.');
  }
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
  var months = summaryMonthKeys_(collectSummaryDates_(ss));
  if (!months.length) months = summaryMonthKeysFromList_([], seoulMonthKey_());
  var orderEnd = ctx.endRow || ORDER_LAST_ROW_DEFAULT;
  var incomeEnd = formulaLastRow_(incomeSh);
  var expenseEnd = formulaLastRow_(expenseSh);
  var width = Math.max(sh.getLastColumn(), SUMMARY_HEADERS.length);
  var needed = months.length + 6;
  var clearRows = Math.min(Math.max(sh.getLastRow(), needed, 30), 240);
  ensureSize_(sh, Math.max(needed, clearRows), SUMMARY_HEADERS.length);
  sh.getRange(1, 1, clearRows, width).clearContent();
  sh.getRange(1, 1, 1, SUMMARY_HEADERS.length).setValues([SUMMARY_HEADERS]);
  sh.getRange(2, 1, months.length, 1).setNumberFormat('yy-mm');
  sh.getRange(2, 1, months.length, 1).setValues(months.map(function(key) { return [monthKeyToDate_(key)]; }));
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
      '=IF(B' + row + '=0,"–",H' + row + '/B' + row + ')',
      '=SUM($B$2:B' + row + ')',
      '=SUM($G$2:G' + row + ')',
      '=J' + row + '-K' + row,
      '=IF(J' + row + '=0,"–",L' + row + '/J' + row + ')'
    ]);
  }
  sh.getRange(2, 2, formulas.length, 12).setFormulas(localizeGrid_(formulas));
  var totalRow = months.length + 2;
  sh.getRange(totalRow, 1).setNumberFormat('@');
  sh.getRange(totalRow, 1).setValue('합계');
  var totals = [];
  var c;
  for (c = 2; c <= 8; c++) {
    var letter = indexToCol_(c);
    totals.push('=SUM(' + letter + '2:' + letter + (totalRow - 1) + ')');
  }
  totals.push('=IF(B' + totalRow + '=0,"–",H' + totalRow + '/B' + totalRow + ')');
  totals.push('=B' + totalRow);
  totals.push('=G' + totalRow);
  totals.push('=H' + totalRow);
  totals.push('=IF(B' + totalRow + '=0,"–",H' + totalRow + '/B' + totalRow + ')');
  sh.getRange(totalRow, 2, 1, 12).setFormulas(localizeGrid_([totals]));
  sh.getRange(2, 2, totalRow - 1, 7).setNumberFormat('#,##0');
  sh.getRange(2, 9, totalRow - 1, 1).setNumberFormat('0.0%');
  sh.getRange(2, 10, totalRow - 1, 3).setNumberFormat('#,##0');
  sh.getRange(2, 13, totalRow - 1, 1).setNumberFormat('0.0%');
  fitColumns_(sh);
  report.push('월별 요약을 ' + formatYyMm_(months[0]) + '부터 ' + formatYyMm_(months[months.length - 1]) + '까지 다시 썼습니다. 입금, 누적 열, YY-MM 표시를 넣었고 수입·지출·주문 연동의 데이터는 바꾸지 않았습니다.');
}

function collectSummaryDates_(ss) {
  var dates = [];
  ['수입', '지출', '주문 연동'].forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) return;
    var col = findHeader_(headerMap_(sheet), ['날짜']);
    if (!col) return;
    var last = Math.min(Math.max(sheet.getLastRow(), 1), MANUAL_LAST_ROW);
    if (last < 2) return;
    var values = sheet.getRange(2, col, last - 1, 1).getValues();
    var i;
    for (i = 0; i < values.length; i++) dates.push(values[i][0]);
  });
  return dates;
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
  var ship = ss.getSheetByName('주문 출고');
  var guide = columnGuide_();
  var cleared = 0;
  cleared += clearPersonalNotes_(orderSh, 1);
  cleared += clearPersonalNotes_(incomeSh, 1);
  if (summary) cleared += clearPersonalNotes_(summary, 16);
  if (ship) cleared += clearPersonalNotes_(ship, 1);
  orderSh.getRange(1, 8).setValue(guide);
  if (summary) summary.getRange(summaryGuideRow_(summary), 1).setValue(guide).setWrap(true);
  if (ship) placeShipGuide_(ship);
  report.push(cleared ? ('이전 안내 문구 ' + cleared + '곳을 지우고 새 안내를 넣었습니다.') : '안내 문구를 현재 구조로 넣었습니다.');
}

function summaryGuideRow_(sh) {
  var last = Math.min(Math.max(sh.getLastRow(), 1), 240);
  var values = sh.getRange(1, 1, last, 1).getValues();
  var i;
  for (i = values.length - 1; i >= 1; i--) {
    if (compact_(String(values[i][0] == null ? '' : values[i][0])) === '합계') return i + 3;
  }
  return 16;
}

function columnGuide_() {
  return [
    DOC_MARKER,
    '입금은 주문 연동의 금액입니다. 취소·환불은 제외합니다. 월별 요약은 주문일(제출일시)로 달을 나눕니다.',
    '가져오는 값은 제출일시, 상태, 금액뿐입니다.',
    '수입 시트는 날짜, 금액, 메모입니다. 주문 연동에 없는 입금만 적습니다.',
    '같은 입금을 주문 연동과 수입 시트에 모두 적으면 두 번 합산됩니다.',
    '지출 분류는 제본, AI, 광고, 박스입니다.',
    '이익률은 순이익을 입금으로 나눈 값입니다. 입금이 0이면 – 입니다.',
    '누적 수입·누적 지출·누적 순이익·누적 수익률은 첫 달부터 그 달까지의 합입니다.',
    '재고는 재고 품목, 입출고, 현재 재고 시트에서 이익과 따로 계산합니다.',
    '재고 입고는 입고일, 재고 출고는 출고일로만 움직입니다. 주문일을 출고일로 쓰지 않습니다.'
  ].join('\n');
}

function clearPersonalNotes_(sh, minRow) {
  var map = headerMap_(sh);
  var dataCols = {};
  Object.keys(map).forEach(function(name) {
    if (!isPersonalHeader_(name)) dataCols[map[name]] = true;
  });
  var lastRow = Math.min(Math.max(sh.getLastRow(), minRow), 80);
  var lastCol = Math.min(Math.max(sh.getLastColumn(), 16), sh.getMaxColumns());
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
      var outside = !dataCols[c + 1];
      if (sheetRow >= 2 && !outside && sh.getName() !== '월별 요약') continue;
      var text = String(values[r][c] == null ? '' : values[r][c]);
      if (!shouldScrubDoc_(text) && !(outside && isLooseNote_(text))) continue;
      sh.getRange(sheetRow, c + 1).clearContent();
      changed++;
    }
  }
  return changed;
}

function shouldScrubDoc_(text) {
  var s = String(text || '').trim();
  if (!s) return false;
  if (s.indexOf(DOC_MARKER) === 0 && !legacyNote_(s) && !PERSONAL_HEADER_RE.test(s)) return false;
  if (PERSONAL_HEADER_RE.test(s)) return true;
  if (legacyNote_(s)) return true;
  if (/IMPORTRANGE/i.test(s)) return true;
  if (/[A-Z]{1,2}\s*열/.test(s) && /주문|금액|집계|교재|상태/.test(s) && s.length > 15) return true;
  return false;
}

function legacyNote_(text) {
  var s = String(text || '');
  if (/네이버\s*폼|naver\s*form/i.test(s)) return true;
  if (/주문\s*입금/.test(s)) return true;
  if (/기타\s*수입/.test(s)) return true;
  if (/집계\s*포함/.test(s)) return true;
  if (/직접\s*입력\s*금지/.test(s)) return true;
  if (/자동\s*연동/.test(s)) return true;
  if (/액세스\s*허용/.test(s)) return true;
  if (/[A-Z]\d+\s*[·・]\s*[A-Z]\d+/.test(s)) return true;
  if (/[A-Z]\d+\s*[~～]\s*[A-Z]\d+/.test(s)) return true;
  if (/※/.test(s) && /안내/.test(s)) return true;
  return false;
}

function isLooseNote_(text) {
  var s = String(text || '').trim();
  if (!s || s.indexOf(DOC_MARKER) === 0) return false;
  if (/[\r\n]/.test(s) && s.length > 12) return true;
  if (/^※/.test(s)) return true;
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
  applyImportNumberFormats_(sh, endRow, { dateCol: 1, rawDateCol: 5, amountCol: 3, rawAmountCol: 7 });
}

function refreshOrderConversions_(sh, ctx) {
  var map = headerMap_(sh);
  var endRow = (ctx && ctx.endRow) || ORDER_LAST_ROW_DEFAULT;
  var rawDateCol = findHeader_(map, ['원본 제출일시']) || 5;
  var rawAmountCol = findHeader_(map, ['원본 금액']) || 7;
  var dateCol = findHeader_(map, ['날짜']);
  var amountCol = findHeader_(map, ['금액']);
  var changed = rewriteConversionColumn_(sh, endRow, dateCol, rawDateCol, dateFormula_, dateFormulaReady_);
  changed = rewriteConversionColumn_(sh, endRow, amountCol, rawAmountCol, amountFormula_, amountFormulaReady_) || changed;
  applyImportNumberFormats_(sh, endRow, {
    dateCol: dateCol,
    rawDateCol: rawDateCol,
    amountCol: amountCol,
    rawAmountCol: rawAmountCol
  });
  return changed;
}

function refreshShipConversions_(sh, ctx) {
  var map = headerMap_(sh);
  var endRow = (ctx && ctx.endRow) || ORDER_LAST_ROW_DEFAULT;
  var rawDateCol = findHeader_(map, ['원본 제출일시']) || 5;
  var dateCol = findHeader_(map, ['주문일', '날짜']);
  var changed = rewriteConversionColumn_(sh, endRow, dateCol, rawDateCol, dateFormula_, dateFormulaReady_);
  applyImportNumberFormats_(sh, endRow, { dateCol: dateCol, rawDateCol: rawDateCol });
  return changed;
}

function rewriteConversionColumn_(sh, endRow, col, sourceCol, formulaFn, readyFn) {
  if (!col || !sourceCol) return false;
  if (readyFn(sh.getRange(2, col).getFormula())) return false;
  var height = endRow - 1;
  var letter = indexToCol_(sourceCol);
  var grid = [];
  var r;
  for (r = 2; r <= endRow; r++) grid.push([formulaFn(letter + r)]);
  ensureSize_(sh, endRow, Math.max(col, sourceCol));
  sh.getRange(2, col, height, 1).setFormulas(localizeGrid_(grid));
  return true;
}

function applyImportNumberFormats_(sh, endRow, spec) {
  var height = Math.max(endRow - 1, 1);
  ensureSize_(sh, endRow, 8);
  function paint(col, pattern) {
    if (!col) return;
    sh.getRange(2, col, height, 1).setNumberFormat(pattern);
  }
  paint(spec.dateCol, 'yyyy-mm-dd');
  paint(spec.rawDateCol, 'yyyy-mm-dd');
  paint(spec.amountCol, '#,##0');
  paint(spec.rawAmountCol, '#,##0');
}

function dateFormulaReady_(formula) {
  var s = String(formula || '');
  return s.indexOf('ISNUMBER') !== -1 && s.indexOf('REGEXEXTRACT') !== -1;
}

function amountFormulaReady_(formula) {
  return String(formula || '').indexOf('ISNUMBER') !== -1;
}

function importFormula_(id, c1, c2, endRow) {
  return '=IMPORTRANGE("' + id + '","\'' + SOURCE_SHEET_NAME + '\'!' + c1 + '2:' + c2 + endRow + '")';
}

function dateFormula_(cell) {
  var text = 'TO_TEXT(' + cell + ')';
  var parsed = 'DATE(' +
    'VALUE(REGEXEXTRACT(' + text + ',"(\\d{4})\\s*[.]\\s*\\d{1,2}\\s*[.]\\s*\\d{1,2}")),' +
    'VALUE(REGEXEXTRACT(' + text + ',"\\d{4}\\s*[.]\\s*(\\d{1,2})\\s*[.]\\s*\\d{1,2}")),' +
    'VALUE(REGEXEXTRACT(' + text + ',"\\d{4}\\s*[.]\\s*\\d{1,2}\\s*[.]\\s*(\\d{1,2})"))' +
    ')';
  var fromText = 'IFERROR(' + parsed + ',IFERROR(DATEVALUE(SUBSTITUTE(SUBSTITUTE(' + text + ',". ","-"),".","-")),))';
  return '=IF(' + cell + '="","",IF(AND(ISNUMBER(' + cell + '),' + cell + '>=30000,' + cell + '<80000),' + cell + ',' + fromText + '))';
}

function echoFormula_(cell) {
  return '=IF(' + cell + '="","",' + cell + ')';
}

function amountFormula_(cell) {
  return '=IF(' + cell + '="","",IF(ISNUMBER(' + cell + '),' + cell + ',IFERROR(VALUE(REGEXREPLACE(TO_TEXT(' + cell + '),"[^0-9.]",""))*1,)))';
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

function fitColumns_(sh) {
  if (!sh) return;
  var cols = Math.max(1, Math.min(sh.getLastColumn(), 30));
  try { sh.autoResizeColumns(1, cols); } catch (err) {}
}

function fitLedgerColumns_(ss) {
  ['수입', '지출', '주문 연동', '월별 요약', '분류', '재고 품목', '입출고', '현재 재고', '주문 출고'].forEach(function(name) {
    fitColumns_(ss.getSheetByName(name));
  });
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

// ── 재고 (이익 장부와 분리) ──────────────────────────

var ITEM_HEADERS = ['코드', '이름', '종류', '구매단위', '단위당개수', '기초수량', '최소재고', '메모'];
var RULE_HEADERS = ['최소권수', '최대권수', '박스코드', '박스개수'];
var MOVE_HEADERS = ['날짜', '구분', '품목', '수량', '단위', '환산수량', '메모', '출처', '주문행', '주문일', '입고일', '출고일'];
var STOCK_HEADERS = ['코드', '이름', '종류', '기초', '입고', '출고', '조정', '현재', '최소재고', '부족'];
var SHIP_HEADERS = ['주문일', '상태', '교재', '', '원본 제출일시', '원본 상태', '원본 교재', '입고일', '출고일'];
var STOCK_ITEM_LAST = 200;
var STOCK_MOVE_LAST = 5000;

function refreshStockFromMenu() {
  var ui = SpreadsheetApp.getUi();
  try {
    var data = withLockMs_(20000, function() { return loadStock_(); });
    ui.alert('재고를 다시 계산했습니다. 부족 ' + (data.low ? data.low.length : 0) + '개.');
  } catch (err) {
    ui.alert(String(err && err.message ? err.message : err).slice(0, 1500));
  }
}

function loadStock_() {
  var ss = activeSs_();
  detectFormulaSep_(ss);
  var ctx = orderContextOrDefault_(ss);
  migrateInventory_(ss, ctx, []);
  var snapshot = stockSnapshot_(ss);
  snapshot.ok = true;
  snapshot.version = APP_VERSION;
  return snapshot;
}

function orderContextOrDefault_(ss) {
  try { return captureOrderContext_(ss); }
  catch (err) { return { sourceId: SOURCE_SPREADSHEET_ID, endRow: ORDER_LAST_ROW_DEFAULT }; }
}

function migrateInventory_(ss, ctx, report) {
  report = report || [];
  var items = ensureStockSheet_(ss, '재고 품목');
  var moves = ensureStockSheet_(ss, '입출고');
  var current = ensureStockSheet_(ss, '현재 재고');
  var ship = ensureStockSheet_(ss, '주문 출고');
  ensureHeaderGroup_(items, ITEM_HEADERS, 1);
  ensureHeaderGroup_(items, RULE_HEADERS, 10);
  ensureHeaderGroup_(moves, MOVE_HEADERS, 1);
  ensureHeaderGroup_(current, STOCK_HEADERS, 1);
  migrateMovementDates_(moves);
  var added = appendItems_(items, itemsToAdd_(readItemRecords_(items)));
  if (added) report.push('재고 품목 ' + added + '개를 채웠습니다. 박스 팩당 개수는 1로 두었으니, 1팩이 여러 개면 단위당개수를 고치세요.');
  else report.push('재고 품목은 이미 있습니다. 기초·최소·팩 수량은 바꾸지 않았습니다.');
  seedRulesOnce_(items, report);
  ensureShipImport_(ship, ctx || orderContextOrDefault_(ss), report);
  var replaced = replaceOrderMovements_(ss);
  if (replaced.importReady) report.push('주문 재고 이동 ' + replaced.count + '줄을 다시 맞췄습니다. 출고일이 비어 있는 주문은 재고에서 빼지 않았습니다.');
  else report.push('주문 출고를 아직 읽지 못했습니다. 액세스 허용 뒤 재고 새로고침을 실행하세요.');
  rebuildCurrentStock_(ss);
  report.push('현재 재고 수식을 품목별로 다시 썼습니다.');
  return report;
}

function ensureStockSheet_(ss, name) {
  var sh = ss.getSheetByName(name);
  if (sh) return sh;
  return ss.insertSheet(name);
}

function ensureHeaderGroup_(sh, names, startCol) {
  var map = headerMap_(sh);
  var missing = [];
  var i;
  for (i = 0; i < names.length; i++) {
    if (!names[i]) continue;
    if (!findHeader_(map, [names[i]])) missing.push(names[i]);
  }
  if (!missing.length) return;
  var anchor = String(sh.getRange(1, startCol).getValue() || '').trim();
  var expected = names.filter(function(name) { return !!name; });
  if (!anchor && missing.length === expected.length) {
    ensureSize_(sh, 2, startCol + names.length - 1);
    sh.getRange(1, startCol, 1, names.length).setValues([names]);
    return;
  }
  var col = Math.max(sh.getLastColumn(), 1) + 1;
  ensureSize_(sh, 2, col + missing.length - 1);
  sh.getRange(1, col, 1, missing.length).setValues([missing]);
}

function defaultWorkbookItems_() {
  return ['4-1', '4-2', '5-1', '5-2', '6-1', '6-2'].map(function(code) {
    return { code: code, name: '워크북 ' + code, kind: '교재', purchaseUnit: '개', perPack: 1, opening: 0, minimum: 0, memo: '' };
  });
}

function defaultBoxItems_() {
  return [
    { code: '박스-소', name: '작은 박스', kind: '박스', purchaseUnit: '팩', perPack: 1, opening: 0, minimum: 0, memo: '' },
    { code: '박스-중', name: '중간 박스', kind: '박스', purchaseUnit: '팩', perPack: 1, opening: 0, minimum: 0, memo: '' },
    { code: '박스-대', name: '큰 박스', kind: '박스', purchaseUnit: '팩', perPack: 1, opening: 0, minimum: 0, memo: '' }
  ];
}

function defaultBoxRules_() {
  return [
    { min: 1, max: 1, code: '박스-소', count: 1 },
    { min: 2, max: 3, code: '박스-중', count: 1 },
    { min: 4, max: 20, code: '박스-대', count: 1 }
  ];
}

function itemsToAdd_(existing) {
  var have = {};
  var boxCount = 0;
  (existing || []).forEach(function(it) {
    if (!it || !it.code) return;
    have[String(it.code).trim()] = true;
    if (compact_(it.kind) === '박스') boxCount++;
  });
  var add = [];
  defaultWorkbookItems_().forEach(function(it) {
    if (!have[it.code]) add.push(it);
  });
  if (boxCount === 0) {
    defaultBoxItems_().forEach(function(it) {
      if (!have[it.code]) add.push(it);
    });
  }
  return add;
}

function rulesToSeed_(initialized, existingRules, items) {
  if (initialized) return [];
  if (existingRules && existingRules.length) return [];
  var codes = {};
  (items || []).forEach(function(it) { if (it && it.code) codes[it.code] = true; });
  var defaults = defaultBoxRules_();
  var i;
  for (i = 0; i < defaults.length; i++) if (!codes[defaults[i].code]) return [];
  return defaults;
}

function rulesFlag_() {
  return PropertiesService.getScriptProperties().getProperty(STOCK_RULES_KEY) === '1';
}

function markRulesFlag_() {
  PropertiesService.getScriptProperties().setProperty(STOCK_RULES_KEY, '1');
}

function seedRulesOnce_(sh, report) {
  if (rulesFlag_()) {
    report.push('박스 규칙은 그대로 둡니다.');
    return;
  }
  var rules = readRuleRecords_(sh);
  var items = readItemRecords_(sh);
  var seeded = rulesToSeed_(false, rules, items);
  if (seeded.length) {
    appendRules_(sh, seeded);
    report.push('박스 규칙 예시를 넣었습니다. 1권은 작은 박스 1개, 2–3권은 중간 박스 1개, 4–20권은 큰 박스 1개입니다. 주문 권수에 맞는 첫 규칙을 씁니다.');
  } else if (rules.length) report.push('박스 규칙이 이미 있어 예시를 넣지 않았습니다.');
  else report.push('박스 코드가 예시와 달라 규칙 예시를 넣지 않았습니다. 재고 화면에서 규칙을 추가하세요.');
  markRulesFlag_();
}

function appendItems_(sh, items) {
  if (!items || !items.length) return 0;
  var map = headerMap_(sh);
  var codeCol = findHeader_(map, ['코드']);
  if (!codeCol) fail_('재고 품목 시트에서 코드 열을 찾지 못했습니다.');
  var row = lastFilledRow_(sh, codeCol, STOCK_ITEM_LAST) + 1;
  var i;
  for (i = 0; i < items.length; i++) writeItemValues_(sh, map, row + i, items[i]);
  return items.length;
}

function writeItemValues_(sh, map, row, item) {
  writeMapped_(sh, row, map, [
    { names: ['코드'], kind: 'text', value: item.code, required: true, plain: true },
    { names: ['이름'], kind: 'text', value: item.name || item.code, plain: true },
    { names: ['종류'], kind: 'text', value: item.kind || '교재' },
    { names: ['구매단위'], kind: 'text', value: item.purchaseUnit || '개' },
    { names: ['단위당개수'], kind: 'number', value: item.perPack > 0 ? item.perPack : 1, required: true },
    { names: ['기초수량'], kind: 'number', value: item.opening || 0, required: true },
    { names: ['최소재고'], kind: 'number', value: item.minimum || 0, required: true },
    { names: ['메모'], kind: 'text', value: item.memo || '' }
  ]);
}

function appendRules_(sh, rules) {
  if (!rules || !rules.length) return 0;
  var map = headerMap_(sh);
  var minCol = findHeader_(map, ['최소권수']);
  if (!minCol) fail_('재고 품목 시트에서 최소권수 열을 찾지 못했습니다.');
  var row = lastFilledRow_(sh, minCol, STOCK_ITEM_LAST) + 1;
  var i;
  for (i = 0; i < rules.length; i++) writeRuleValues_(sh, map, row + i, rules[i]);
  return rules.length;
}

function writeRuleValues_(sh, map, row, rule) {
  writeMapped_(sh, row, map, [
    { names: ['최소권수'], kind: 'number', value: rule.min, required: true },
    { names: ['최대권수'], kind: 'number', value: rule.max, required: true },
    { names: ['박스코드'], kind: 'text', value: rule.code, required: true, plain: true },
    { names: ['박스개수'], kind: 'number', value: rule.count, required: true }
  ]);
}

function lastFilledRow_(sh, col, maxRow) {
  if (!col) return 1;
  var last = Math.min(Math.max(sh.getLastRow(), 1), maxRow);
  if (last < 2) return 1;
  var values = sh.getRange(2, col, last - 1, 1).getValues();
  var lastData = 1;
  var i;
  for (i = 0; i < values.length; i++) {
    if (String(values[i][0] == null ? '' : values[i][0]).trim() !== '') lastData = i + 2;
  }
  return lastData;
}

function readItemRecords_(sh) {
  if (!sh) return [];
  var map = headerMap_(sh);
  var codeCol = findHeader_(map, ['코드']);
  if (!codeCol) return [];
  var nameCol = findHeader_(map, ['이름']);
  var kindCol = findHeader_(map, ['종류']);
  var unitCol = findHeader_(map, ['구매단위']);
  var packCol = findHeader_(map, ['단위당개수']);
  var openCol = findHeader_(map, ['기초수량']);
  var minCol = findHeader_(map, ['최소재고']);
  var memoCol = findHeader_(map, ['메모']);
  var last = lastFilledRow_(sh, codeCol, STOCK_ITEM_LAST);
  if (last < 2) return [];
  var width = Math.max(sh.getLastColumn(), codeCol);
  var values = sh.getRange(2, 1, last - 1, width).getValues();
  var out = [];
  var r;
  for (r = 0; r < values.length; r++) {
    var code = String(cellAt_(values[r], codeCol) || '').trim();
    if (!code || code.charAt(0) === '#') continue;
    var per = packCol ? asNumber_(cellAt_(values[r], packCol)) : 1;
    var opening = openCol ? asNumber_(cellAt_(values[r], openCol)) : 0;
    var minimum = minCol ? asNumber_(cellAt_(values[r], minCol)) : 0;
    out.push({
      row: r + 2,
      code: code,
      name: nameCol ? String(cellAt_(values[r], nameCol) || '').trim() : code,
      kind: kindCol ? String(cellAt_(values[r], kindCol) || '').trim() : '교재',
      purchaseUnit: unitCol ? (String(cellAt_(values[r], unitCol) || '').trim() || '개') : '개',
      perPack: per != null && per > 0 ? per : 1,
      opening: opening == null ? 0 : opening,
      minimum: minimum == null ? 0 : minimum,
      memo: memoCol ? String(cellAt_(values[r], memoCol) || '').trim() : ''
    });
  }
  return out;
}

function readRuleRecords_(sh) {
  if (!sh) return [];
  var map = headerMap_(sh);
  var minCol = findHeader_(map, ['최소권수']);
  var maxCol = findHeader_(map, ['최대권수']);
  var codeCol = findHeader_(map, ['박스코드']);
  var countCol = findHeader_(map, ['박스개수']);
  if (!minCol || !maxCol || !codeCol || !countCol) return [];
  var last = lastFilledRow_(sh, minCol, STOCK_ITEM_LAST);
  if (last < 2) return [];
  var width = Math.max(sh.getLastColumn(), countCol, codeCol);
  var values = sh.getRange(2, 1, last - 1, width).getValues();
  var out = [];
  var r;
  for (r = 0; r < values.length; r++) {
    var min = asNumber_(cellAt_(values[r], minCol));
    var max = asNumber_(cellAt_(values[r], maxCol));
    var code = String(cellAt_(values[r], codeCol) || '').trim();
    var count = asNumber_(cellAt_(values[r], countCol));
    if (min == null && max == null && !code) continue;
    out.push({ row: r + 2, min: min, max: max, code: code, count: count == null ? 0 : count });
  }
  return out;
}

function migrateMovementDates_(sh) {
  if (!sh) return;
  ensureHeaderGroup_(sh, ['주문일', '입고일', '출고일'], 1);
  var map = headerMap_(sh);
  var dateCol = findHeader_(map, ['날짜']);
  var inCol = findHeader_(map, ['입고일']);
  var outCol = findHeader_(map, ['출고일']);
  var kindCol = findHeader_(map, ['구분']);
  var sourceCol = findHeader_(map, ['출처']);
  var codeCol = findHeader_(map, ['품목']);
  if (!kindCol || !codeCol || !inCol || !outCol) return;
  var last = lastFilledRow_(sh, codeCol, STOCK_MOVE_LAST);
  if (last < 2) return;
  var width = Math.max(sh.getLastColumn(), dateCol || 1, inCol, outCol, kindCol, sourceCol || 1);
  var values = sh.getRange(2, 1, last - 1, width).getValues();
  var inValues = sh.getRange(2, inCol, last - 1, 1).getValues();
  var outValues = sh.getRange(2, outCol, last - 1, 1).getValues();
  var changedIn = false;
  var changedOut = false;
  var r;
  for (r = 0; r < values.length; r++) {
    var kind = compact_(cellAt_(values[r], kindCol));
    var source = sourceCol ? String(cellAt_(values[r], sourceCol) || '') : '';
    var legacy = dateCol ? cellAt_(values[r], dateCol) : '';
    if (!isFilledDate_(legacy)) continue;
    if (kind === '입고' && !isFilledDate_(inValues[r][0])) {
      inValues[r][0] = legacy;
      changedIn = true;
    }
    if (kind === '출고' && !isOrderSource_(source) && !isFilledDate_(outValues[r][0])) {
      outValues[r][0] = legacy;
      changedOut = true;
    }
  }
  if (changedIn) {
    sh.getRange(2, inCol, last - 1, 1).setValues(inValues);
    sh.getRange(2, inCol, last - 1, 1).setNumberFormat('yyyy-mm-dd');
  }
  if (changedOut) {
    sh.getRange(2, outCol, last - 1, 1).setValues(outValues);
    sh.getRange(2, outCol, last - 1, 1).setNumberFormat('yyyy-mm-dd');
  }
}

function readMovementRecords_(sh) {
  if (!sh) return [];
  var map = headerMap_(sh);
  var dateCol = findHeader_(map, ['날짜']);
  var orderDateCol = findHeader_(map, ['주문일']);
  var inCol = findHeader_(map, ['입고일']);
  var outCol = findHeader_(map, ['출고일']);
  var kindCol = findHeader_(map, ['구분']);
  var codeCol = findHeader_(map, ['품목']);
  var qtyCol = findHeader_(map, ['수량']);
  var unitCol = findHeader_(map, ['단위']);
  var convCol = findHeader_(map, ['환산수량']);
  var memoCol = findHeader_(map, ['메모']);
  var sourceCol = findHeader_(map, ['출처']);
  var orderCol = findHeader_(map, ['주문행']);
  if (!codeCol) return [];
  var last = lastFilledRow_(sh, codeCol, STOCK_MOVE_LAST);
  if (last < 2) return [];
  var width = Math.max(sh.getLastColumn(), codeCol);
  var values = sh.getRange(2, 1, last - 1, width).getValues();
  var out = [];
  var r;
  for (r = 0; r < values.length; r++) {
    var code = String(cellAt_(values[r], codeCol) || '').trim();
    if (!code || code.charAt(0) === '#') continue;
    var qty = qtyCol ? asNumber_(cellAt_(values[r], qtyCol)) : null;
    var converted = convCol ? asNumber_(cellAt_(values[r], convCol)) : null;
    var orderRow = orderCol ? asNumber_(cellAt_(values[r], orderCol)) : null;
    var kind = kindCol ? String(cellAt_(values[r], kindCol) || '').trim() : '';
    var source = sourceCol ? String(cellAt_(values[r], sourceCol) || '').trim() : '';
    var legacy = dateCol ? formatCellDate_(cellAt_(values[r], dateCol)) : '';
    var orderDate = orderDateCol ? formatCellDate_(cellAt_(values[r], orderDateCol)) : '';
    var inDate = inCol ? formatCellDate_(cellAt_(values[r], inCol)) : '';
    var outDate = outCol ? formatCellDate_(cellAt_(values[r], outCol)) : '';
    var kindKey = compact_(kind);
    if (kindKey === '입고' && !inDate) inDate = legacy;
    if (kindKey === '출고' && !outDate && !isOrderSource_(source)) outDate = legacy;
    var effective = legacy;
    if (kindKey === '입고') effective = inDate || legacy;
    else if (kindKey === '출고') effective = outDate || '';
    out.push({
      row: r + 2,
      date: effective,
      orderDate: orderDate,
      inDate: inDate,
      outDate: outDate,
      kind: kind,
      code: code,
      qty: qty,
      unit: unitCol ? String(cellAt_(values[r], unitCol) || '').trim() : '개',
      converted: converted,
      memo: memoCol ? String(cellAt_(values[r], memoCol) || '').trim() : '',
      source: source,
      orderRow: orderRow
    });
  }
  return out;
}

function cellAt_(row, col) {
  if (!col || !row || col > row.length) return '';
  return row[col - 1];
}

function parseWorkbookText_(text) {
  var re = /(?:^|[^\d])([4-6])\s*[-~./]\s*([12])(?!\d)(?:\s*[\(（]\s*(\d+)\s*권?\s*[\)）]|\s+(\d+)\s*권)?/g;
  var merged = {};
  var order = [];
  var s = String(text == null ? '' : text);
  var m;
  while ((m = re.exec(s))) {
    var code = m[1] + '-' + m[2];
    var qty = Number(m[3] || m[4] || 1);
    if (!isFinite(qty) || qty <= 0) qty = 1;
    qty = Math.floor(qty);
    if (!merged[code]) {
      merged[code] = 0;
      order.push(code);
    }
    merged[code] += qty;
  }
  return order.map(function(code) { return { code: code, qty: merged[code] }; });
}

function boxesForOrder_(total, rules) {
  var books = Number(total);
  if (!isFinite(books) || books <= 0) return null;
  var i;
  for (i = 0; i < (rules || []).length; i++) {
    var rule = rules[i];
    if (!rule || !rule.code || !(rule.count > 0)) continue;
    if (books >= rule.min && books <= rule.max) return { code: rule.code, qty: rule.count };
  }
  return null;
}

function convertedQty_(kind, qty, unit, perPack, purchaseUnit) {
  var n = Number(qty);
  if (!isFinite(n)) return null;
  var pack = Number(perPack);
  if (!isFinite(pack) || pack <= 0) pack = 1;
  var u = compact_(unit || '개');
  var purchase = compact_(purchaseUnit || '개');
  var multiply = u === '팩' || (u && u !== '개' && purchase !== '개' && u === purchase);
  var pieces = multiply ? Math.abs(n) * pack : Math.abs(n);
  var k = compact_(kind);
  if (k === '입고') return pieces;
  if (k === '출고') return -pieces;
  if (k === '조정') return n < 0 ? -pieces : pieces;
  return null;
}

function movementsFromOrders_(rows, rules) {
  var out = [];
  (rows || []).forEach(function(row) {
    if (!row) return;
    if (/취소|환불/.test(String(row.status || ''))) return;
    var books = parseWorkbookText_(row.text);
    if (!books.length) return;
    var total = 0;
    books.forEach(function(book) { total += book.qty; });
    var memo = clipText_(row.text, 500);
    var orderDate = row.orderDate || row.date || '';
    function pushMove(kind, date, box) {
      if (!date) return;
      var sign = kind === '입고' ? 1 : -1;
      out.push({
        date: date,
        orderDate: orderDate,
        inDate: kind === '입고' ? date : '',
        outDate: kind === '출고' ? date : '',
        kind: kind,
        code: box ? box.code : '',
        qty: box ? box.qty : 0,
        unit: '개',
        converted: sign * (box ? box.qty : 0),
        memo: box ? ('주문 ' + total + '권') : memo,
        source: '주문',
        orderRow: row.row,
        box: !!box
      });
    }
    books.forEach(function(book) {
      if (row.inDate) {
        out.push({
          date: row.inDate,
          orderDate: orderDate,
          inDate: row.inDate,
          outDate: '',
          kind: '입고',
          code: book.code,
          qty: book.qty,
          unit: '개',
          converted: book.qty,
          memo: memo,
          source: '주문',
          orderRow: row.row,
          box: false
        });
      }
      if (row.outDate) {
        out.push({
          date: row.outDate,
          orderDate: orderDate,
          inDate: '',
          outDate: row.outDate,
          kind: '출고',
          code: book.code,
          qty: book.qty,
          unit: '개',
          converted: -book.qty,
          memo: memo,
          source: '주문',
          orderRow: row.row,
          box: false
        });
      }
    });
    var box = boxesForOrder_(total, rules);
    if (box && row.outDate) pushMove('출고', row.outDate, box);
  });
  out.sort(function(a, b) {
    var ar = Number(a.orderRow) || 0;
    var br = Number(b.orderRow) || 0;
    if (ar !== br) return ar - br;
    if (a.kind !== b.kind) return a.kind === '입고' ? -1 : 1;
    return String(a.code).localeCompare(String(b.code), 'ko');
  });
  return out;
}

function buildStock_(items, movements) {
  var byCode = {};
  var order = [];
  function bucket(code) {
    if (!byCode[code]) {
      byCode[code] = {
        code: code, name: code, kind: '교재', purchaseUnit: '개', perPack: 1,
        opening: 0, minimum: 0, inbound: 0, outbound: 0, adjustment: 0
      };
      order.push(code);
    }
    return byCode[code];
  }
  (items || []).forEach(function(it) {
    if (!it || !it.code) return;
    var b = bucket(String(it.code).trim());
    b.name = it.name || b.code;
    b.kind = it.kind || '교재';
    b.purchaseUnit = it.purchaseUnit || '개';
    b.perPack = it.perPack > 0 ? it.perPack : 1;
    b.opening = Number(it.opening) || 0;
    b.minimum = Number(it.minimum) || 0;
  });
  (movements || []).forEach(function(mv) {
    if (!mv || !mv.code) return;
    var n = Number(mv.converted);
    if (!isFinite(n)) return;
    var kind = compact_(mv.kind);
    if (kind === '입고' && !mv.inDate) return;
    if (kind === '출고' && !mv.outDate) return;
    var b = bucket(String(mv.code).trim());
    if (kind === '입고') b.inbound += n;
    else if (kind === '출고') b.outbound += -n;
    else if (kind === '조정') b.adjustment += n;
  });
  var list = order.map(function(code) {
    var b = byCode[code];
    var onHand = b.opening + b.inbound - b.outbound + b.adjustment;
    return {
      code: b.code,
      name: b.name,
      kind: b.kind,
      purchaseUnit: b.purchaseUnit,
      perPack: b.perPack,
      opening: b.opening,
      inbound: b.inbound,
      outbound: b.outbound,
      adjustment: b.adjustment,
      onHand: onHand,
      minimum: b.minimum,
      low: onHand < b.minimum
    };
  });
  return { items: list, low: list.filter(function(it) { return it.low; }) };
}

function stockSnapshot_(ss) {
  var itemSheet = ss.getSheetByName('재고 품목');
  var moveSheet = ss.getSheetByName('입출고');
  var ship = ss.getSheetByName('주문 출고');
  var items = readItemRecords_(itemSheet);
  var movements = readMovementRecords_(moveSheet);
  var rules = readRuleRecords_(itemSheet).map(function(rule) {
    return { row: rule.row, min: rule.min, max: rule.max, boxCode: rule.code, count: rule.count };
  });
  var stock = buildStock_(items, movements);
  var manual = movements.filter(function(mv) { return !isOrderSource_(mv.source); });
  var shipments = movements.filter(function(mv) { return isOrderSource_(mv.source); });
  var orderDates = orderDateRows_(ship);
  return {
    asOf: Utilities.formatDate(new Date(), sheetTz_(), "yyyy-MM-dd'T'HH:mm:ss"),
    importReady: shipImportReady_(ship),
    items: stock.items,
    low: stock.low,
    rules: rules,
    manual: manual.slice(-80),
    shipments: shipments.slice(-40),
    orderMovements: shipments.length,
    orderDates: orderDates.rows,
    orderDateCount: orderDates.count
  };
}

function orderDateRows_(ship) {
  var read = readShipOrders_(ship);
  if (!read.ready) return { rows: [], count: 0 };
  var rows = [];
  read.rows.forEach(function(row) {
    if (/취소|환불/.test(String(row.status || ''))) return;
    if (!parseWorkbookText_(row.text).length) return;
    rows.push({
      row: row.row,
      orderDate: row.orderDate || '',
      inDate: row.inDate || '',
      outDate: row.outDate || '',
      status: row.status || '',
      text: clipText_(row.text, 180),
      shipped: !!row.outDate
    });
  });
  var count = rows.length;
  rows.reverse();
  return { rows: rows.slice(0, 60), count: count };
}

function isOrderSource_(source) {
  return compact_(source) === '주문';
}

function ensureShipImport_(sh, ctx, report) {
  ensureShipDateColumns_(sh, report);
  var problems = shipSheetHealth_(sh);
  if (problems.length) {
    rebuildShipSheet_(sh, ctx);
    report.push('주문 출고를 제출일시·상태·교재만 가져오도록 맞췄습니다. 입고일과 출고일은 비워 둡니다.');
    return;
  }
  if (refreshShipConversions_(sh, ctx)) report.push('주문 출고의 주문일 변환을 날짜 숫자와 글자 모두 받게 고쳤습니다. 출고일에는 복사하지 않습니다.');
  else report.push('주문 출고는 주문일·상태·교재를 가져옵니다. 출고일이 비어 있으면 재고를 빼지 않습니다.');
  placeShipGuide_(sh);
}

function shipSheetHealth_(sh) {
  var problems = [];
  if (!sh) return ['시트가 없습니다.'];
  var map = headerMap_(sh);
  if (!findHeader_(map, ['주문일', '날짜'])) problems.push('주문일 헤더가 없습니다.');
  ['상태', '교재'].forEach(function(name) {
    if (!findHeader_(map, [name])) problems.push(name + ' 헤더가 없습니다.');
  });
  Object.keys(map).forEach(function(name) {
    if (isPersonalHeader_(name)) problems.push('개인정보 헤더가 있습니다: ' + name);
  });
  var imports = listImports_(sh);
  var dates = null;
  var books = null;
  imports.forEach(function(imp) {
    if (imp.sheetName && imp.sheetName !== SOURCE_SHEET_NAME) return;
    if (imp.startCol <= 3 && imp.endCol >= 3) problems.push('가져오기에 주문자 열이 남아 있습니다.');
    if (imp.startCol <= 4 && imp.endCol >= 4) problems.push('가져오기에 연락처 열이 남아 있습니다.');
    if (imp.startCol <= 6 && imp.endCol >= 6) problems.push('가져오기에 금액 열이 남아 있습니다.');
    if (imp.startCol === 1 && imp.endCol === 2) dates = imp;
    else if (imp.startCol === 5 && imp.endCol === 5) books = imp;
    else problems.push('가져오기 범위가 제출일시·상태 또는 교재가 아닙니다.');
  });
  if (!dates) problems.push('제출일시·상태 가져오기가 없습니다.');
  if (!books) problems.push('교재 가져오기가 없습니다.');
  if (dates && books) {
    problems = problems.concat(expectFormulaRef_(sh, findHeader_(map, ['주문일', '날짜']), dates.col, '주문일'));
    problems = problems.concat(expectFormulaRef_(sh, findHeader_(map, ['상태']), dates.col + 1, '상태'));
    problems = problems.concat(expectFormulaRef_(sh, findHeader_(map, ['교재']), books.col, '교재'));
  }
  return uniqueTexts_(problems);
}

function rebuildShipSheet_(sh, ctx) {
  var endRow = (ctx && ctx.endRow) || ORDER_LAST_ROW_DEFAULT;
  ensureSize_(sh, endRow, SHIP_HEADERS.length);
  listImports_(sh).forEach(function(imp) { sh.getRange(imp.row, imp.col).clearContent(); });
  var clearCols = Math.min(Math.max(sh.getLastColumn(), SHIP_HEADERS.length), 12);
  sh.getRange(1, 1, endRow, clearCols).clearContent();
  sh.getRange(1, 1, 1, SHIP_HEADERS.length).setValues([SHIP_HEADERS]);
  var id = (ctx && ctx.sourceId) || SOURCE_SPREADSHEET_ID;
  sh.getRange(2, 5).setFormula(localizeFormula_(importFormula_(id, 'A', 'B', endRow)));
  sh.getRange(2, 7).setFormula(localizeFormula_(importFormula_(id, 'E', 'E', endRow)));
  var height = endRow - 1;
  var dates = [];
  var statuses = [];
  var books = [];
  var r;
  for (r = 2; r <= endRow; r++) {
    dates.push([dateFormula_('E' + r)]);
    statuses.push([echoFormula_('F' + r)]);
    books.push([echoFormula_('G' + r)]);
  }
  sh.getRange(2, 1, height, 1).setFormulas(localizeGrid_(dates));
  sh.getRange(2, 2, height, 1).setFormulas(localizeGrid_(statuses));
  sh.getRange(2, 3, height, 1).setFormulas(localizeGrid_(books));
  applyImportNumberFormats_(sh, endRow, { dateCol: 1, rawDateCol: 5 });
  var inCol = findHeader_(headerMap_(sh), ['입고일']);
  var outCol = findHeader_(headerMap_(sh), ['출고일']);
  if (inCol) sh.getRange(2, inCol, height, 1).setNumberFormat('yyyy-mm-dd');
  if (outCol) sh.getRange(2, outCol, height, 1).setNumberFormat('yyyy-mm-dd');
  placeShipGuide_(sh);
}

function ensureShipDateColumns_(sh, report) {
  if (!sh) return;
  report = report || [];
  var map = headerMap_(sh);
  var orderCol = findHeader_(map, ['주문일']);
  var legacy = findHeader_(map, ['날짜']);
  if (!orderCol && legacy) {
    sh.getRange(1, legacy).setValue('주문일');
    report.push('주문 출고의 날짜 헤더를 주문일로 바꿨습니다. 이 열은 제출일시이고 출고일이 아닙니다.');
    map = headerMap_(sh);
  }
  var missing = [];
  if (!findHeader_(map, ['입고일'])) missing.push('입고일');
  if (!findHeader_(map, ['출고일'])) missing.push('출고일');
  if (!missing.length) {
    placeShipGuide_(sh);
    return;
  }
  var start = nextDataHeaderCol_(sh);
  var c;
  for (c = start; c < start + missing.length + 2; c++) {
    if (isGuideCell_(sh, c)) sh.getRange(1, c).clearContent();
  }
  ensureSize_(sh, Math.max(sh.getMaxRows(), 2), start + missing.length - 1);
  sh.getRange(1, start, 1, missing.length).setValues([missing]);
  var height = Math.max(Math.min(sh.getMaxRows(), ORDER_LAST_ROW_DEFAULT) - 1, 1);
  sh.getRange(2, start, height, missing.length).setNumberFormat('yyyy-mm-dd');
  report.push('주문 출고에 ' + missing.join(', ') + ' 열을 추가했습니다. 주문일을 출고일로 복사하지 않습니다. 출고일이 비어 있으면 재고를 빼지 않습니다.');
  placeShipGuide_(sh);
}

function nextDataHeaderCol_(sh) {
  var map = headerMap_(sh);
  var col = 0;
  Object.keys(map).forEach(function(name) { col = Math.max(col, map[name]); });
  return col + 1;
}

function isGuideCell_(sh, col) {
  if (!col || col > sh.getMaxColumns()) return false;
  var text = String(sh.getRange(1, col).getValue() || '');
  if (!text) return false;
  if (text.indexOf(DOC_MARKER) === 0) return true;
  return text.length > 40;
}

function placeShipGuide_(sh) {
  if (!sh) return;
  var existing = 0;
  var last = Math.min(Math.max(sh.getLastColumn(), 1), 20);
  var c;
  for (c = 1; c <= last; c++) {
    var text = String(sh.getRange(1, c).getValue() || '');
    if (text.indexOf(DOC_MARKER) === 0) existing = c;
  }
  var col = existing || (nextDataHeaderCol_(sh) + 1);
  ensureSize_(sh, 1, col);
  sh.getRange(1, col).setValue(shipGuide_()).setWrap(true);
}

function shipGuide_() {
  return [
    DOC_MARKER,
    '재고 출고는 이 시트에서 교재 글자만 읽습니다.',
    '이익 계산용 주문 연동에는 교재를 넣지 않습니다. 월별 입금은 주문일(제출일시)을 씁니다.',
    '가져오는 값은 제출일시, 상태, 교재뿐입니다.',
    '입고일과 출고일은 직접 적습니다. 제출일시를 출고일에 복사하지 않습니다.',
    '입고일이 있으면 그 날짜에 교재를 입고로 더합니다. 박스는 주문 입고에 넣지 않습니다.',
    '출고일이 있으면 그 날짜에 교재와 박스를 출고로 뺍니다.',
    '출고일이 비어 있으면 아직 발송하지 않은 주문이므로 재고를 빼지 않습니다.',
    '취소와 환불은 재고에서 빼지 않습니다.'
  ].join('\n');
}

function readShipOrders_(sh) {
  if (!sh) return { ready: false, rows: [] };
  if (shipSheetHealth_(sh).length) return { ready: false, rows: [] };
  var map = headerMap_(sh);
  var dateCol = findHeader_(map, ['주문일', '날짜']);
  var inCol = findHeader_(map, ['입고일']);
  var outCol = findHeader_(map, ['출고일']);
  var statusCol = findHeader_(map, ['상태']);
  var textCol = findHeader_(map, ['교재']);
  var last = Math.min(Math.max(sh.getLastRow(), 1), ORDER_LAST_ROW_DEFAULT);
  if (last < 2) return { ready: true, rows: [] };
  var width = Math.max(sh.getLastColumn(), textCol || 1, outCol || 1, inCol || 1);
  var values = sh.getRange(2, 1, last - 1, width).getValues();
  var rows = [];
  var r;
  for (r = 0; r < values.length; r++) {
    var dateRaw = cellAt_(values[r], dateCol);
    var statusRaw = cellAt_(values[r], statusCol);
    var textRaw = cellAt_(values[r], textCol);
    if (isSheetError_(dateRaw) || isSheetError_(statusRaw) || isSheetError_(textRaw)) return { ready: false, rows: [] };
    var text = String(textRaw == null ? '' : textRaw).trim();
    var status = String(statusRaw == null ? '' : statusRaw).trim();
    var orderDate = formatCellDate_(dateRaw);
    var inDate = optionalDate_(inCol ? cellAt_(values[r], inCol) : '');
    var outDate = optionalDate_(outCol ? cellAt_(values[r], outCol) : '');
    if (!text && !status && !orderDate && !inDate && !outDate) continue;
    if (!orderDate && !text) continue;
    rows.push({
      row: r + 2,
      date: orderDate,
      orderDate: orderDate,
      inDate: inDate,
      outDate: outDate,
      status: status,
      text: text
    });
  }
  return { ready: true, rows: rows };
}

function optionalDate_(value) {
  if (isSheetError_(value)) return '';
  return formatCellDate_(value);
}

function shipImportReady_(sh) {
  return readShipOrders_(sh).ready;
}

function isSheetError_(value) {
  return String(value == null ? '' : value).charAt(0) === '#';
}

function replaceOrderMovements_(ss) {
  flushSheet_();
  var ship = ss.getSheetByName('주문 출고');
  var moveSh = ss.getSheetByName('입출고');
  var itemSh = ss.getSheetByName('재고 품목');
  var read = readShipOrders_(ship);
  if (!read.ready) {
    return { importReady: false, count: readMovementRecords_(moveSh).filter(function(mv) { return isOrderSource_(mv.source); }).length };
  }
  var generated = movementsFromOrders_(read.rows, readRuleRecords_(itemSh));
  ensureMissingItems_(itemSh, generated);
  var manual = readMovementRecords_(moveSh).filter(function(mv) { return !isOrderSource_(mv.source); });
  rewriteMovements_(moveSh, manual.concat(generated));
  return { importReady: true, count: generated.length };
}

function ensureMissingItems_(sh, moves) {
  var items = readItemRecords_(sh);
  var have = {};
  items.forEach(function(it) { have[it.code] = true; });
  var add = [];
  (moves || []).forEach(function(mv) {
    if (!mv || !mv.code || have[mv.code]) return;
    have[mv.code] = true;
    if (mv.box) add.push({ code: mv.code, name: mv.code, kind: '박스', purchaseUnit: '팩', perPack: 1, opening: 0, minimum: 0, memo: '' });
    else add.push({ code: mv.code, name: '워크북 ' + mv.code, kind: '교재', purchaseUnit: '개', perPack: 1, opening: 0, minimum: 0, memo: '' });
  });
  appendItems_(sh, add);
}

function rewriteMovements_(sh, moves) {
  ensureHeaderGroup_(sh, MOVE_HEADERS, 1);
  var last = Math.max(sh.getLastRow(), 1);
  if (last >= 2) sh.getRange(2, 1, last - 1, Math.max(sh.getLastColumn(), 1)).clearContent();
  appendMovementValues_(sh, moves || []);
}

function appendMovementValues_(sh, moves) {
  if (!moves.length) return;
  var map = headerMap_(sh);
  var codeCol = findHeader_(map, ['품목']);
  var start = lastFilledRow_(sh, codeCol, STOCK_MOVE_LAST) + 1;
  var width = Math.max(sh.getLastColumn(), MOVE_HEADERS.length);
  ensureSize_(sh, start + moves.length - 1, width);
  var grid = moves.map(function(mv) {
    var row = [];
    var c;
    for (c = 0; c < width; c++) row.push('');
    function put(names, value) {
      var col = findHeader_(map, names);
      if (col && col <= width) row[col - 1] = value == null ? '' : value;
    }
    put(['날짜'], mv.date || '');
    put(['주문일'], mv.orderDate || '');
    put(['입고일'], mv.inDate || '');
    put(['출고일'], mv.outDate || '');
    put(['구분'], mv.kind || '');
    put(['품목'], mv.code || '');
    put(['수량'], mv.qty == null ? '' : mv.qty);
    put(['단위'], mv.unit || '개');
    put(['환산수량'], mv.converted == null ? '' : mv.converted);
    put(['메모'], mv.memo || '');
    put(['출처'], mv.source || '직접');
    put(['주문행'], mv.orderRow == null ? '' : mv.orderRow);
    return row;
  });
  sh.getRange(start, 1, grid.length, width).setValues(grid);
  if (codeCol) {
    var codes = moves.map(function(mv) { return [String(mv.code || '')]; });
    sh.getRange(start, codeCol, codes.length, 1).setNumberFormat('@');
    sh.getRange(start, codeCol, codes.length, 1).setValues(codes);
  }
  ['날짜', '주문일', '입고일', '출고일'].forEach(function(name) {
    var col = findHeader_(map, [name]);
    if (col) sh.getRange(start, col, grid.length, 1).setNumberFormat('yyyy-mm-dd');
  });
  var qtyCol = findHeader_(map, ['수량']);
  var convCol = findHeader_(map, ['환산수량']);
  if (qtyCol) sh.getRange(start, qtyCol, grid.length, 1).setNumberFormat('#,##0.##');
  if (convCol) sh.getRange(start, convCol, grid.length, 1).setNumberFormat('#,##0.##');
}

function rebuildCurrentStock_(ss) {
  var items = ss.getSheetByName('재고 품목');
  var current = ss.getSheetByName('현재 재고');
  var moves = ss.getSheetByName('입출고');
  ensureHeaderGroup_(current, STOCK_HEADERS, 1);
  var src = {
    code: colLetter_(items, '코드'),
    name: colLetter_(items, '이름'),
    kind: colLetter_(items, '종류'),
    opening: colLetter_(items, '기초수량'),
    minimum: colLetter_(items, '최소재고'),
    moveQty: colLetter_(moves, '환산수량'),
    moveKind: colLetter_(moves, '구분'),
    moveCode: colLetter_(moves, '품목'),
    moveIn: colLetter_(moves, '입고일'),
    moveOut: colLetter_(moves, '출고일')
  };
  var dest = {
    code: colLetter_(current, '코드'),
    name: colLetter_(current, '이름'),
    kind: colLetter_(current, '종류'),
    opening: colLetter_(current, '기초'),
    inbound: colLetter_(current, '입고'),
    outbound: colLetter_(current, '출고'),
    adjustment: colLetter_(current, '조정'),
    onHand: colLetter_(current, '현재'),
    minimum: colLetter_(current, '최소재고'),
    low: colLetter_(current, '부족')
  };
  var codeCol = findHeader_(headerMap_(items), ['코드']);
  var itemRows = [];
  if (codeCol) {
    var last = lastFilledRow_(items, codeCol, STOCK_ITEM_LAST);
    if (last >= 2) {
      var values = items.getRange(2, codeCol, last - 1, 1).getValues();
      var i;
      for (i = 0; i < values.length; i++) {
        var code = String(values[i][0] == null ? '' : values[i][0]).trim();
        if (code) itemRows.push(i + 2);
      }
    }
  }
  var maxCol = STOCK_HEADERS.length;
  Object.keys(dest).forEach(function(key) { maxCol = Math.max(maxCol, colToIndex_(dest[key])); });
  ensureSize_(current, Math.max(itemRows.length + 5, 20), maxCol);
  var clearRows = Math.min(Math.max(current.getMaxRows(), 2), STOCK_ITEM_LAST + 1) - 1;
  if (clearRows >= 1) current.getRange(2, 1, clearRows, maxCol).clearContent();
  if (!itemRows.length) return;
  var formulas = itemRows.map(function(itemRow, index) {
    var byName = currentStockFormulas_(itemRow, index + 2, STOCK_MOVE_LAST, src, dest);
    var row = [];
    var c;
    for (c = 0; c < maxCol; c++) row.push('');
    STOCK_HEADERS.forEach(function(name) {
      var col = findHeader_(headerMap_(current), [name]);
      if (col) row[col - 1] = byName[name];
    });
    return row;
  });
  current.getRange(2, 1, formulas.length, maxCol).setFormulas(localizeGrid_(formulas));
  ['기초', '입고', '출고', '조정', '현재', '최소재고'].forEach(function(name) {
    var col = findHeader_(headerMap_(current), [name]);
    if (col) current.getRange(2, col, formulas.length, 1).setNumberFormat('#,##0.##');
  });
  var shownCode = findHeader_(headerMap_(current), ['코드']);
  if (shownCode) current.getRange(2, shownCode, formulas.length, 1).setNumberFormat('@');
}

function currentStockFormulas_(itemRow, destRow, moveEnd, src, dest) {
  var codeCell = dest.code + destRow;
  var codeRef = quoteSheet_('재고 품목') + '!' + src.code + itemRow;
  var itemCodes = quoteSheet_('재고 품목') + '!$' + src.code + '$2:$' + src.code + '$' + STOCK_ITEM_LAST;
  var names = quoteSheet_('재고 품목') + '!$' + src.name + '$2:$' + src.name + '$' + STOCK_ITEM_LAST;
  var kinds = quoteSheet_('재고 품목') + '!$' + src.kind + '$2:$' + src.kind + '$' + STOCK_ITEM_LAST;
  var opening = quoteSheet_('재고 품목') + '!$' + src.opening + '$2:$' + src.opening + '$' + STOCK_ITEM_LAST;
  var minimum = quoteSheet_('재고 품목') + '!$' + src.minimum + '$2:$' + src.minimum + '$' + STOCK_ITEM_LAST;
  var moveQty = quoteSheet_('입출고') + '!$' + src.moveQty + '$2:$' + src.moveQty + '$' + moveEnd;
  var moveKind = quoteSheet_('입출고') + '!$' + src.moveKind + '$2:$' + src.moveKind + '$' + moveEnd;
  var moveCode = quoteSheet_('입출고') + '!$' + src.moveCode + '$2:$' + src.moveCode + '$' + moveEnd;
  var moveIn = src.moveIn ? (quoteSheet_('입출고') + '!$' + src.moveIn + '$2:$' + src.moveIn + '$' + moveEnd) : '';
  var moveOut = src.moveOut ? (quoteSheet_('입출고') + '!$' + src.moveOut + '$2:$' + src.moveOut + '$' + moveEnd) : '';
  function idx(col) {
    return 'IFERROR(INDEX(' + col + ',MATCH(' + codeCell + ',' + itemCodes + ',0)),"")';
  }
  var out = {};
  out['코드'] = '=IF(' + codeRef + '="","",' + codeRef + ')';
  out['이름'] = '=IF(' + codeCell + '="","",' + idx(names) + ')';
  out['종류'] = '=IF(' + codeCell + '="","",' + idx(kinds) + ')';
  out['기초'] = '=IF(' + codeCell + '="","",IFERROR(INDEX(' + opening + ',MATCH(' + codeCell + ',' + itemCodes + ',0))*1,0))';
  out['입고'] = '=IF(' + codeCell + '="","",SUMIFS(' + moveQty + ',' + moveKind + ',"입고",' + moveCode + ',' + codeCell + (moveIn ? (',' + moveIn + ',"<>"') : '') + '))';
  out['출고'] = '=IF(' + codeCell + '="","",-SUMIFS(' + moveQty + ',' + moveKind + ',"출고",' + moveCode + ',' + codeCell + (moveOut ? (',' + moveOut + ',"<>"') : '') + '))';
  out['조정'] = '=IF(' + codeCell + '="","",SUMIFS(' + moveQty + ',' + moveKind + ',"조정",' + moveCode + ',' + codeCell + '))';
  out['현재'] = '=IF(' + codeCell + '="","",' + dest.opening + destRow + '+' + dest.inbound + destRow + '-' + dest.outbound + destRow + '+' + dest.adjustment + destRow + ')';
  out['최소재고'] = '=IF(' + codeCell + '="","",IFERROR(INDEX(' + minimum + ',MATCH(' + codeCell + ',' + itemCodes + ',0))*1,0))';
  out['부족'] = '=IF(' + codeCell + '="","",IF(' + dest.onHand + destRow + '<' + dest.minimum + destRow + ',"부족",""))';
  return out;
}

function colLetter_(sh, name) {
  var col = findHeader_(headerMap_(sh), [name]);
  if (!col) fail_(sh.getName() + ' 시트에서 ' + name + ' 열을 찾지 못했습니다.');
  return indexToCol_(col);
}

function flushSheet_() {
  try { SpreadsheetApp.flush(); } catch (err) {}
}

function clipText_(value, max) {
  var s = String(value == null ? '' : value).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '').trim();
  if (s.length > max) return s.slice(0, max);
  return s;
}

function mutateStock_(action, sheetName, data) {
  var ss = activeSs_();
  detectFormulaSep_(ss);
  if (sheetName === '입출고') return mutateMovement_(ss, action, data);
  if (sheetName === '재고 품목') return mutateItem_(ss, action, data);
  if (sheetName === '주문 출고') return updateShipDates_(ss, action, data);
  return mutateRule_(ss, action, data);
}

function updateShipDates_(ss, action, data) {
  if (action !== 'update') fail_('주문 출고에서는 입고일과 출고일만 고칠 수 있습니다.');
  var sh = ensureStockSheet_(ss, '주문 출고');
  ensureShipDateColumns_(sh, []);
  var map = headerMap_(sh);
  var row = parseRow_(data.row);
  var record = data.record || {};
  rejectBlockedKeys_(record);
  var orderCol = findHeader_(map, ['주문일', '날짜']);
  var textCol = findHeader_(map, ['교재']);
  var actualOrder = orderCol ? formatCellDate_(sh.getRange(row, orderCol).getValue()) : '';
  var text = textCol ? String(sh.getRange(row, textCol).getValue() || '').trim() : '';
  if (!text && !actualOrder) fail_('주문이 없는 행입니다.');
  if (data.match && String(firstField_(data.match, ['orderDate', '주문일']) || '').trim()) {
    var expect = formatCellDate_(parseDateInput_(firstField_(data.match, ['orderDate', '주문일'])));
    if (!actualOrder || actualOrder !== expect) fail_('행 내용이 일치하지 않습니다. 재고를 새로고침한 뒤 다시 시도하세요.');
  }
  if (hasField_(record, ['입고일', 'inDate'])) writeOptionalDate_(sh, row, findHeader_(map, ['입고일']), firstField_(record, ['입고일', 'inDate']));
  if (hasField_(record, ['출고일', 'outDate'])) writeOptionalDate_(sh, row, findHeader_(map, ['출고일']), firstField_(record, ['출고일', 'outDate']));
  replaceOrderMovements_(ss);
  rebuildCurrentStock_(ss);
  fitColumns_(sh);
  return { ok: true, version: APP_VERSION, sheet: '주문 출고', row: row };
}

function writeOptionalDate_(sh, row, col, value) {
  if (!col) fail_('날짜 열을 찾지 못했습니다. 초기 설정을 다시 실행하세요.');
  var cell = sh.getRange(row, col);
  if (value == null || String(value).trim() === '') {
    cell.clearContent();
    return;
  }
  cell.setValue(parseDateInput_(value));
  cell.setNumberFormat('yyyy-mm-dd');
}

function mutateMovement_(ss, action, data) {
  var sh = ensureStockSheet_(ss, '입출고');
  ensureHeaderGroup_(sh, MOVE_HEADERS, 1);
  var map = headerMap_(sh);
  var record = data.record || {};
  rejectBlockedKeys_(record);
  if (action === 'add') {
    var row = lastFilledRow_(sh, findHeader_(map, ['품목']), STOCK_MOVE_LAST) + 1;
    var written = movementFromRecord_(ss, record);
    writeMovementRow_(sh, map, row, written);
    return { ok: true, version: APP_VERSION, sheet: '입출고', row: row, converted: written.converted };
  }
  var rowNum = parseRow_(data.row);
  assertMovementMatch_(sh, map, rowNum, data.match);
  if (action === 'delete') {
    sh.deleteRow(rowNum);
    return { ok: true, version: APP_VERSION, sheet: '입출고', row: rowNum };
  }
  var updated = movementFromRecord_(ss, record);
  writeMovementRow_(sh, map, rowNum, updated);
  return { ok: true, version: APP_VERSION, sheet: '입출고', row: rowNum, converted: updated.converted };
}

function movementFromRecord_(ss, record) {
  var kind = compact_(firstField_(record, ['구분', 'kind']));
  if (kind !== '입고' && kind !== '출고' && kind !== '조정') fail_('구분은 입고, 출고, 조정 중 하나여야 합니다.');
  var code = cleanCode_(firstField_(record, ['품목', '코드', 'code']));
  var item = findItemRecord_(ss, code);
  if (!item) fail_('품목 목록에 없는 코드입니다. 재고 품목에 먼저 추가하세요.');
  var qty = parseAmount_(firstField_(record, ['수량', 'qty']));
  if (qty == null) fail_('수량을 입력하세요.');
  if (kind !== '조정' && qty <= 0) fail_('입고와 출고 수량은 0보다 커야 합니다.');
  if (kind === '조정' && qty === 0) fail_('조정 수량은 0이 아니어야 합니다.');
  var unit = cleanText_(firstField_(record, ['단위', 'unit']) || '개', 20) || '개';
  var converted = convertedQty_(kind, qty, unit, item.perPack, item.purchaseUnit);
  if (converted == null) fail_('수량을 환산하지 못했습니다.');
  var rawDate = firstField_(record, ['날짜', 'date']);
  if (kind === '입고') rawDate = firstField_(record, ['입고일', 'inDate']) || rawDate;
  if (kind === '출고') rawDate = firstField_(record, ['출고일', 'outDate']) || rawDate;
  var date = parseDateInput_(rawDate);
  return {
    date: date,
    orderDate: '',
    inDate: kind === '입고' ? date : '',
    outDate: kind === '출고' ? date : '',
    kind: kind,
    code: code,
    qty: kind === '조정' ? qty : Math.abs(qty),
    unit: unit,
    converted: converted,
    memo: cleanText_(firstField_(record, ['메모', 'memo']), 500),
    source: '직접',
    orderRow: ''
  };
}

function writeMovementRow_(sh, map, row, mv) {
  var date = mv.date instanceof Date ? mv.date : parseDateInput_(mv.date);
  var kind = compact_(mv.kind);
  var inDate = kind === '입고' ? date : '';
  var outDate = kind === '출고' ? date : '';
  writeMapped_(sh, row, map, [
    { names: ['날짜'], kind: 'date', value: date, required: true },
    { names: ['주문일'], kind: 'text', value: '' },
    { names: ['입고일'], kind: kind === '입고' ? 'date' : 'text', value: inDate },
    { names: ['출고일'], kind: kind === '출고' ? 'date' : 'text', value: outDate },
    { names: ['구분'], kind: 'text', value: mv.kind, required: true },
    { names: ['품목'], kind: 'text', value: mv.code, required: true, plain: true },
    { names: ['수량'], kind: 'number', value: mv.qty, required: true },
    { names: ['단위'], kind: 'text', value: mv.unit || '개' },
    { names: ['환산수량'], kind: 'number', value: mv.converted, required: true },
    { names: ['메모'], kind: 'text', value: mv.memo || '' },
    { names: ['출처'], kind: 'text', value: '직접', required: true },
    { names: ['주문행'], kind: 'text', value: '' }
  ]);
  fitColumns_(sh);
}

function assertMovementMatch_(sh, map, row, match) {
  if (!match) fail_('수정·삭제에는 날짜, 품목, 환산수량 또는 수량이 필요합니다.');
  var sourceCol = findHeader_(map, ['출처']);
  var source = sourceCol ? sh.getRange(row, sourceCol).getValue() : '';
  if (isOrderSource_(source)) fail_('주문에서 온 출고는 지우거나 고치지 않습니다. 주문이 바뀌면 재고 새로고침이 다시 계산합니다.');
  var dateCol = findHeader_(map, ['날짜']);
  var codeCol = findHeader_(map, ['품목']);
  var convCol = findHeader_(map, ['환산수량']);
  var qtyCol = findHeader_(map, ['수량']);
  var actualDate = dateCol ? formatCellDate_(sh.getRange(row, dateCol).getValue()) : '';
  var actualCode = codeCol ? String(sh.getRange(row, codeCol).getValue() || '').trim() : '';
  var actualConverted = convCol ? asNumber_(sh.getRange(row, convCol).getValue()) : null;
  var actualQty = qtyCol ? asNumber_(sh.getRange(row, qtyCol).getValue()) : null;
  var expectDate = formatCellDate_(parseDateInput_(firstField_(match, ['date', '날짜'])));
  var expectCode = String(firstField_(match, ['code', '품목']) || '').trim();
  var expectConverted = hasField_(match, ['converted', '환산수량']) ? parseAmount_(firstField_(match, ['converted', '환산수량'])) : null;
  var expectQty = hasField_(match, ['qty', '수량']) ? parseAmount_(firstField_(match, ['qty', '수량'])) : null;
  var convertedOk = expectConverted != null && actualConverted != null && Math.abs(actualConverted - expectConverted) < 0.001;
  var qtyOk = expectQty != null && actualQty != null && Math.abs(actualQty - expectQty) < 0.001;
  if (!actualDate || actualDate !== expectDate || !expectCode || actualCode !== expectCode || (!convertedOk && !qtyOk)) {
    fail_('행 내용이 일치하지 않습니다. 재고를 새로고침한 뒤 다시 시도하세요.');
  }
}

function findItemRecord_(ss, code) {
  var items = readItemRecords_(ss.getSheetByName('재고 품목'));
  var i;
  for (i = 0; i < items.length; i++) if (items[i].code === code) return items[i];
  return null;
}

function cleanCode_(value) {
  var s = cleanText_(value, 40);
  if (!s) fail_('품목 코드를 입력하세요.');
  return s;
}

function cleanKind_(value) {
  var s = compact_(value);
  if (s === '교재' || s === '워크북') return '교재';
  if (s === '박스' || s === '포장') return '박스';
  fail_('종류는 교재 또는 박스여야 합니다.');
}

function mutateItem_(ss, action, data) {
  var sh = ensureStockSheet_(ss, '재고 품목');
  ensureHeaderGroup_(sh, ITEM_HEADERS, 1);
  ensureHeaderGroup_(sh, RULE_HEADERS, 10);
  var map = headerMap_(sh);
  var record = data.record || {};
  rejectBlockedKeys_(record);
  if (action === 'add') {
    var code = cleanCode_(firstField_(record, ['코드', 'code']));
    if (findItemRecord_(ss, code)) fail_('이미 있는 품목입니다.');
    var item = itemFromRecord_(record, code, true);
    var row = lastFilledRow_(sh, findHeader_(map, ['코드']), STOCK_ITEM_LAST) + 1;
    writeItemValues_(sh, map, row, item);
    rebuildCurrentStock_(ss);
    return { ok: true, version: APP_VERSION, sheet: '재고 품목', row: row, code: code };
  }
  var codeKey = String(firstField_(record, ['코드', 'code']) || firstField_(data.match || {}, ['code', '코드']) || '').trim();
  var existing = findItemRecord_(ss, codeKey);
  if (!existing) fail_('품목을 찾지 못했습니다.');
  if (data.match && String(firstField_(data.match, ['code', '코드']) || '').trim() && String(firstField_(data.match, ['code', '코드'])).trim() !== existing.code) {
    fail_('행 내용이 일치하지 않습니다. 재고를 새로고침한 뒤 다시 시도하세요.');
  }
  if (action === 'delete') {
    clearItemRow_(sh, map, existing.row);
    rebuildCurrentStock_(ss);
    return { ok: true, version: APP_VERSION, sheet: '재고 품목', row: existing.row, code: existing.code };
  }
  var next = itemFromRecord_(record, existing.code, false, existing);
  writeItemValues_(sh, map, existing.row, next);
  rebuildCurrentStock_(ss);
  return { ok: true, version: APP_VERSION, sheet: '재고 품목', row: existing.row, code: existing.code };
}

function itemFromRecord_(record, code, creating, existing) {
  existing = existing || {};
  var kind = existing.kind || (creating ? '박스' : '교재');
  if (hasField_(record, ['종류', 'kind']) && String(firstField_(record, ['종류', 'kind']) || '').trim()) {
    kind = cleanKind_(firstField_(record, ['종류', 'kind']));
  }
  var per = existing.perPack > 0 ? existing.perPack : 1;
  if (hasField_(record, ['단위당개수', 'perPack'])) {
    per = parseAmount_(firstField_(record, ['단위당개수', 'perPack']));
    if (per == null || per <= 0) fail_('단위당개수는 0보다 커야 합니다.');
  }
  var opening = existing.opening || 0;
  if (hasField_(record, ['기초수량', 'opening'])) {
    opening = parseAmount_(firstField_(record, ['기초수량', 'opening']));
    if (opening == null) fail_('기초수량을 입력하세요.');
  }
  var minimum = existing.minimum || 0;
  if (hasField_(record, ['최소재고', 'minimum'])) {
    minimum = parseAmount_(firstField_(record, ['최소재고', 'minimum']));
    if (minimum == null || minimum < 0) fail_('최소재고는 0 이상이어야 합니다.');
  }
  var purchase = existing.purchaseUnit || (kind === '박스' ? '팩' : '개');
  if (hasField_(record, ['구매단위', 'purchaseUnit'])) purchase = cleanText_(firstField_(record, ['구매단위', 'purchaseUnit']), 20) || purchase;
  var name = existing.name || code;
  if (hasField_(record, ['이름', 'name'])) name = cleanText_(firstField_(record, ['이름', 'name']), 80) || code;
  var memo = existing.memo || '';
  if (hasField_(record, ['메모', 'memo'])) memo = cleanText_(firstField_(record, ['메모', 'memo']), 500);
  return { code: code, name: name, kind: kind, purchaseUnit: purchase, perPack: per, opening: opening, minimum: minimum, memo: memo };
}

function clearItemRow_(sh, map, row) {
  ['코드', '이름', '종류', '구매단위', '단위당개수', '기초수량', '최소재고', '메모'].forEach(function(name) {
    var col = findHeader_(map, [name]);
    if (col) sh.getRange(row, col).clearContent();
  });
}

function mutateRule_(ss, action, data) {
  var sh = ensureStockSheet_(ss, '재고 품목');
  ensureHeaderGroup_(sh, ITEM_HEADERS, 1);
  ensureHeaderGroup_(sh, RULE_HEADERS, 10);
  var record = data.record || {};
  rejectBlockedKeys_(record);
  var rules = readRuleRecords_(sh);
  if (action === 'add') {
    rules.push(ruleFromRecord_(ss, record));
    replaceRules_(sh, rules);
    return { ok: true, version: APP_VERSION, sheet: '박스규칙', count: rules.length };
  }
  var rowNum = parseRow_(data.row);
  var index = -1;
  var i;
  for (i = 0; i < rules.length; i++) if (rules[i].row === rowNum) index = i;
  if (index < 0) fail_('박스 규칙을 찾지 못했습니다.');
  assertRuleMatch_(rules[index], data.match);
  if (action === 'delete') rules.splice(index, 1);
  else rules[index] = ruleFromRecord_(ss, record);
  replaceRules_(sh, rules);
  return { ok: true, version: APP_VERSION, sheet: '박스규칙', row: rowNum };
}

function ruleFromRecord_(ss, record) {
  var min = requireInt_(firstField_(record, ['최소권수', 'min']), '최소권수', 1, 999);
  var max = requireInt_(firstField_(record, ['최대권수', 'max']), '최대권수', 1, 999);
  if (min > max) fail_('최소권수는 최대권수보다 클 수 없습니다.');
  var code = cleanCode_(firstField_(record, ['박스코드', 'code']));
  var item = findItemRecord_(ss, code);
  if (!item || compact_(item.kind) !== '박스') fail_('박스 규칙에는 종류가 박스인 품목만 쓸 수 있습니다.');
  var count = requireInt_(firstField_(record, ['박스개수', 'count']), '박스개수', 1, 999);
  return { min: min, max: max, code: code, count: count };
}

function assertRuleMatch_(rule, match) {
  if (!match) fail_('수정·삭제에는 최소권수, 최대권수, 박스코드가 필요합니다.');
  var min = asNumber_(firstField_(match, ['min', '최소권수']));
  var max = asNumber_(firstField_(match, ['max', '최대권수']));
  var code = String(firstField_(match, ['code', '박스코드', 'boxCode']) || '').trim();
  if (min == null || max == null || !code || rule.min !== min || rule.max !== max || rule.code !== code) {
    fail_('행 내용이 일치하지 않습니다. 재고를 새로고침한 뒤 다시 시도하세요.');
  }
}

function replaceRules_(sh, rules) {
  var map = headerMap_(sh);
  var cols = ['최소권수', '최대권수', '박스코드', '박스개수'].map(function(name) { return findHeader_(map, [name]); });
  var last = 1;
  cols.forEach(function(col) { if (col) last = Math.max(last, lastFilledRow_(sh, col, STOCK_ITEM_LAST)); });
  if (last >= 2) {
    cols.forEach(function(col) {
      if (col) sh.getRange(2, col, last - 1, 1).clearContent();
    });
  }
  var i;
  for (i = 0; i < rules.length; i++) writeRuleValues_(sh, map, i + 2, rules[i]);
}

function requireInt_(value, label, min, max) {
  var n = typeof value === 'number' ? value : Number(String(value == null ? '' : value).trim());
  if (!isFinite(n) || Math.floor(n) !== n) fail_(label + '는 정수여야 합니다.');
  if (n < min || n > max) fail_(label + ' 범위를 확인해 주세요.');
  return n;
}
