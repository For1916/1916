/**
 * 황소 워크북 주문 — 독립 실행형 Google Apps Script 웹 앱
 *
 * script.google.com 에서 새 프로젝트를 만들고 이 파일 전체를 붙여 넣습니다.
 * 시트에 묶인 스크립트가 아닙니다. 스프레드시트는 openById 로 엽니다.
 *
 * 하는 일: 공개 페이지의 새 주문만 받습니다.
 * 시트에 있는 기존 주문은 어떤 응답으로도 돌려주지 않습니다.
 * doGet 은 안내 한 줄만 반환하고 시트를 열지 않습니다.
 *
 * 배포: 실행 계정 = 나, 액세스 = 모든 사용자(Anyone).
 * 실행 계정은 스프레드시트 [황소 워크북 주문 관리] 를 수정할 수 있어야 하고,
 * 주문 직후 그 계정 메일로 알림을 보냅니다.
 *
 * ★ 사장님 확인: PRODUCTS 의 price 는 모두 임시 기본값 40000원입니다.
 * 권종별 실제 가격이 다르면 숫자만 바꾸세요.
 * 페이지 Workbook_Order/index.html 의 ORDER_CONFIG.products 와 금액을 같게 유지해야
 * 화면 금액과 저장 금액이 어긋나지 않습니다.
 */

var SPREADSHEET_ID = '1s_QC5gRuU7E07WGZrBtbMS_S80YHexlPPD_qYVDqTpM';
var SHEET_NAME = '주문 관리';
var TIMEZONE = 'Asia/Seoul';
var BANK_NAME = '신한은행';
var BANK_ACCOUNT = '110-351-100566';
var DEPOSIT_WINDOW_MINUTES = 30;
var CONSENT_PHRASE = '동의합니다';
var MIN_FILL_MS = 5000;
var MAX_FILL_MS = 3 * 60 * 60 * 1000;
var CLOCK_SKEW_MS = 2 * 60 * 1000;
var MAX_QTY = 30;
var MAX_TOTAL_QTY = 60;
var RATE_LIMIT_COUNT = 3;
var RATE_LIMIT_SECONDS = 30 * 60;
var SCAN_LAST_ROW = 5000;

var PRODUCTS = [
  { code: '4-1', name: '황소 워크북 4-1 (일품·실력 겸용)', price: 40000 },
  { code: '4-2', name: '황소 워크북 4-2', price: 40000 },
  { code: '5-1', name: '황소 워크북 5-1', price: 40000 },
  { code: '5-2', name: '황소 워크북 5-2', price: 40000 },
  { code: '6-1', name: '황소 워크북 6-1', price: 40000 },
  { code: '6-2', name: '황소 워크북 6-2', price: 40000 }
];

var PLATFORMS = ['인스타그램', '블로그', '카페', '지인 추천', '기타'];
var PAID_ANSWERS = ['네, 완료했습니다', '아니요, 곧 송금하겠습니다'];

var HEADER_FIELDS = [
  { key: 'submittedAt', names: ['제출일시'] },
  { key: 'status', names: ['상태'] },
  { key: 'payer', names: ['주문자(예금주)'] },
  { key: 'phone', names: ['연락처'] },
  { key: 'books', names: ['교재'] },
  { key: 'amount', names: ['금액'] },
  { key: 'zip', names: ['우편번호'] },
  { key: 'address', names: ['주소'] },
  { key: 'platform', names: ['유입 플랫폼', '유입플랫폼'] },
  { key: 'memo', names: ['메모'] },
  { key: 'answerId', names: ['모아폼 답변 ID', '모아폼 답변ID', '모아폼답변ID'] },
  { key: 'resultUrl', names: ['모아폼 결과 URL', '모아폼 결과URL', '모아폼결과URL'] },
  { key: 'payerRaw', names: ['주문자 원본 (숨김)', '주문자 원본', '주문자원본(숨김)', '주문자 원본(숨김)'] },
  { key: 'booksRaw', names: ['교재 원본 (숨김)', '교재 원본', '교재원본(숨김)', '교재 원본(숨김)'] }
];

// 장부 IMPORTRANGE 가 A:B, E(교재), F(금액)를 위치로 읽습니다.
// 이 헤더는 없을 때만 마지막 헤더 뒤에 만들고, 열을 끼워 넣지 않습니다.
var EXTRA_FIELDS = [
  { key: 'os', names: ['운영체제'] },
  { key: 'device', names: ['기기'] },
  { key: 'browser', names: ['브라우저'] },
  { key: 'sido', names: ['시도'] },
  { key: 'sigungu', names: ['시군구'] }
];

// ── 웹 앱 ────────────────────────────────────────────

function doGet() {
  return json_({ ok: false, error: '주문은 이 주소로 전송만 할 수 있습니다.' });
}

function doPost(e) {
  try {
    var parsed = readBody_(e);
    if (!parsed.ok) return json_(parsed);
    var checked = validateOrder_(parsed.data, Date.now());
    if (!checked.ok) return json_({ ok: false, error: checked.error });
    var saved = withLock_(function() {
      var limited = rateLimit_(checked.value.phoneDigits);
      if (limited) return limited;
      return { ok: true, saved: appendOrder_(checked.value) };
    });
    if (!saved.ok) return json_({ ok: false, error: saved.error });
    notifyOwner_(saved.saved);
    return json_({
      ok: true,
      orderId: saved.saved.id,
      amount: saved.saved.amount,
      books: saved.saved.books,
      payerName: saved.saved.payerName,
      submittedAt: saved.saved.submittedText
    });
  } catch (err) {
    return json_({ ok: false, error: publicError_(err) });
  }
}

function readBody_(e) {
  var raw = e && e.postData && e.postData.contents;
  if (!raw) return { ok: false, error: '주문 형식이 올바르지 않습니다.' };
  if (String(raw).length > 20000) return { ok: false, error: '주문 내용이 너무 깁니다.' };
  var data;
  try { data = JSON.parse(raw); }
  catch (err) { return { ok: false, error: '주문 형식이 올바르지 않습니다.' }; }
  if (!data || typeof data !== 'object' || Array.isArray(data)) {
    return { ok: false, error: '주문 형식이 올바르지 않습니다.' };
  }
  return { ok: true, data: data };
}

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) fail_('다른 주문을 저장하는 중입니다. 잠시 후 다시 시도해 주세요.');
  try { return fn(); }
  finally { lock.releaseLock(); }
}

function rateLimit_(phoneDigits) {
  var cache = CacheService.getScriptCache();
  var key = 'wo_' + phoneDigits;
  var hits = Number(cache.get(key) || '0');
  if (!isFinite(hits) || hits < 0) hits = 0;
  if (hits >= RATE_LIMIT_COUNT) {
    return { ok: false, error: '같은 연락처로 짧은 시간에 여러 번 주문할 수 없습니다. 잠시 후 다시 시도해 주세요.' };
  }
  cache.put(key, String(hits + 1), RATE_LIMIT_SECONDS);
  return null;
}

// ── 검증 (시트에 쓰기 전에 끝냄) ─────────────────────

function validateOrder_(data, nowMs) {
  data = data || {};
  if (trimText_(data.hp_leave_blank)) {
    return { ok: false, error: '주문을 접수하지 못했습니다. 페이지를 새로고침한 뒤 다시 시도해 주세요.' };
  }
  var opened = Number(data.openedAt);
  if (!isFinite(opened)) {
    return { ok: false, error: '주문 시간을 확인하지 못했습니다. 페이지를 새로고침한 뒤 다시 시도해 주세요.' };
  }
  if (opened > nowMs + CLOCK_SKEW_MS) {
    return { ok: false, error: '주문 시간을 확인하지 못했습니다. 페이지를 새로고침한 뒤 다시 시도해 주세요.' };
  }
  var elapsed = nowMs - opened;
  if (elapsed < MIN_FILL_MS) {
    return { ok: false, error: '내용을 조금 더 확인한 뒤 다시 제출해 주세요.' };
  }
  if (elapsed > MAX_FILL_MS) {
    return { ok: false, error: '접수가 너무 오래 열려 있었습니다. 페이지를 새로고침한 뒤 다시 주문해 주세요.' };
  }

  var items = normalizeItems_(data.items);
  if (items.error) return { ok: false, error: items.error };
  var claimed = strictInt_(data.amount);
  if (claimed == null || claimed !== items.amount) {
    return { ok: false, error: '총 입금 금액이 맞지 않습니다. 페이지를 새로고침한 뒤 다시 주문해 주세요.' };
  }

  var payer = normalizePayer_(data.payer);
  if (!payer) {
    return { ok: false, error: '주문자(예금주)는 이름_전화번호뒤4자리 형식이어야 합니다. 예: 도진우_1916' };
  }
  var phone = normalizePhone_(data.phone);
  if (!phone) return { ok: false, error: '연락처는 한국 휴대폰 번호여야 합니다.' };

  var zip = toAsciiDigits_(trimText_(data.zip)).replace(/[^0-9]/g, '');
  if (!/^\d{5}$/.test(zip)) {
    return { ok: false, error: '우편번호가 올바르지 않습니다. 주소 검색으로 다시 선택하세요.' };
  }
  var road = trimText_(data.road).replace(/\s+/g, ' ');
  if (road.length < 8 || road.length > 150 || !/[가-힣]/.test(road) || !/\d/.test(road) || /[<>]/.test(road)) {
    return { ok: false, error: '주소 검색으로 도로명주소를 선택하세요.' };
  }
  var detail = trimText_(data.detail).replace(/\s+/g, ' ');
  if (!detail || detail.length > 80 || !/[가-힣0-9]/.test(detail) || /[<>]/.test(detail)) {
    return { ok: false, error: detail ? '상세주소를 확인해 주세요.' : '상세주소를 입력하세요.' };
  }
  var region = parseRoadRegion_(road);
  if (!region) {
    return { ok: false, error: '주소에서 시도를 확인하지 못했습니다. 주소 검색으로 다시 선택하세요.' };
  }
  if (normSpace_(data.sido) !== region.sido || normSpace_(data.sigungu) !== region.sigungu) {
    return { ok: false, error: '주소의 시도와 시군구가 맞지 않습니다. 주소 검색으로 다시 선택하세요.' };
  }
  var ua = normSpace_(data.userAgent);
  var client = classifyClient_(ua);
  if (!client) {
    return { ok: false, error: '접속 환경을 확인하지 못했습니다. 페이지를 새로고침한 뒤 다시 주문해 주세요.' };
  }
  if (normSpace_(data.os) !== client.os || normSpace_(data.device) !== client.device || normSpace_(data.browser) !== client.browser) {
    return { ok: false, error: '접속 환경 정보가 맞지 않습니다. 페이지를 새로고침한 뒤 다시 주문해 주세요.' };
  }

  var platform = trimText_(data.platform);
  if (PLATFORMS.indexOf(platform) === -1) return { ok: false, error: '유입 플랫폼을 선택하세요.' };
  var note = trimText_(data.platformNote).replace(/\s+/g, ' ');
  var platformValue = platform;
  if (platform === '기타') {
    if (!note) return { ok: false, error: '기타 경로를 입력하세요.' };
    if (note.length > 40) return { ok: false, error: '기타 경로는 40자 이하로 적어 주세요.' };
    platformValue = '기타: ' + note;
  } else if (note) {
    return { ok: false, error: '유입 플랫폼 입력이 올바르지 않습니다.' };
  }

  var paid = trimText_(data.paid);
  if (PAID_ANSWERS.indexOf(paid) === -1) return { ok: false, error: '송금하셨나요?에 답해 주세요.' };
  if (!exactConsent_(data.consentPrivacy)) {
    return { ok: false, error: '개인정보 동의를 칸에 동의합니다 라고 정확히 입력하세요.' };
  }
  if (!exactConsent_(data.consentNotice)) {
    return { ok: false, error: '주문 안내 동의를 칸에 동의합니다 라고 정확히 입력하세요.' };
  }

  return {
    ok: true,
    value: {
      submittedMs: nowMs,
      payer: payer.value,
      payerRaw: payer.raw,
      payerName: payer.name,
      phone: phone.display,
      phoneDigits: phone.digits,
      zip: zip,
      road: road,
      detail: detail,
      address: road + ' ' + detail,
      sido: region.sido,
      sigungu: region.sigungu,
      os: client.os,
      device: client.device,
      browser: client.browser,
      platform: platformValue,
      memo: '송금: ' + paid,
      paid: paid,
      books: formatBooks_(items.lines),
      booksRaw: formatBooksRaw_(items.lines),
      amount: items.amount
    }
  };
}

function normalizeItems_(items) {
  if (!Array.isArray(items) || !items.length || items.length > 12) {
    return { error: '교재를 1권 이상 선택하세요.' };
  }
  var byCode = {};
  var i;
  for (i = 0; i < items.length; i++) {
    var it = items[i];
    if (!it || typeof it !== 'object') return { error: '교재 선택이 올바르지 않습니다.' };
    var code = trimText_(it.code);
    if (!productByCode_(code)) return { error: '판매하지 않는 교재가 포함되어 있습니다.' };
    var qty = strictInt_(it.qty);
    if (qty == null) return { error: '수량은 정수여야 합니다.' };
    if (qty === 0) continue;
    if (qty < 1 || qty > MAX_QTY) return { error: '수량은 1에서 ' + MAX_QTY + ' 사이의 정수여야 합니다.' };
    byCode[code] = (byCode[code] || 0) + qty;
  }
  var lines = [];
  var totalQty = 0;
  var amount = 0;
  for (i = 0; i < PRODUCTS.length; i++) {
    var product = PRODUCTS[i];
    var q = byCode[product.code] || 0;
    if (!q) continue;
    if (q > MAX_QTY) return { error: '수량은 1에서 ' + MAX_QTY + ' 사이의 정수여야 합니다.' };
    if (strictInt_(product.price) == null || product.price < 1) return { error: '교재 가격 설정을 확인해 주세요.' };
    lines.push({ code: product.code, name: product.name, qty: q, price: product.price });
    totalQty += q;
    amount += q * product.price;
  }
  if (!lines.length) return { error: '교재를 1권 이상 선택하세요.' };
  if (totalQty > MAX_TOTAL_QTY) return { error: '한 번에 주문할 수 있는 권수를 넘었습니다.' };
  return { lines: lines, amount: amount, totalQty: totalQty };
}

function productByCode_(code) {
  var i;
  for (i = 0; i < PRODUCTS.length; i++) if (PRODUCTS[i].code === code) return PRODUCTS[i];
  return null;
}

function formatBooks_(lines) {
  return lines.map(function(line) { return line.code + ' (' + line.qty + '권)'; }).join(', ');
}

function formatBooksRaw_(lines) {
  return lines.map(function(line) { return line.name + ' ' + line.qty + '권'; }).join(', ');
}

function exactConsent_(value) {
  var s = trimText_(value);
  try { s = s.normalize('NFC'); } catch (e) {}
  return s === CONSENT_PHRASE;
}

function normalizePayer_(raw) {
  var original = trimText_(raw);
  try { original = original.normalize('NFC'); } catch (e) {}
  var s = toAsciiDigits_(original).replace(/\uFF3F/g, '_').replace(/\s+/g, '');
  if (!/^[가-힣]{2,20}_[0-9]{4}$/.test(s)) return null;
  return { raw: original, value: s, name: s.split('_')[0] };
}

function normalizePhone_(raw) {
  var s = toAsciiDigits_(trimText_(raw)).replace(/[\s.\-()]/g, '');
  if (s.indexOf('+82') === 0) s = '0' + s.slice(3);
  else if (s.indexOf('82') === 0 && (s.length === 11 || s.length === 12)) s = '0' + s.slice(2);
  if (!/^01[016789]\d{7,8}$/.test(s)) return null;
  if (s.indexOf('010') === 0 && s.length !== 11) return null;
  if (s.length === 10) return { digits: s, display: s.slice(0, 3) + '-' + s.slice(3, 6) + '-' + s.slice(6) };
  if (s.length === 11) return { digits: s, display: s.slice(0, 3) + '-' + s.slice(3, 7) + '-' + s.slice(7) };
  return null;
}

function strictInt_(value) {
  if (typeof value === 'number' && isFinite(value) && Math.floor(value) === value) return value;
  if (typeof value === 'string' && /^-?\d+$/.test(value.trim())) return Number(value.trim());
  return null;
}

function trimText_(value) {
  return String(value == null ? '' : value)
    .replace(/\u00a0/g, ' ')
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '')
    .trim();
}

function toAsciiDigits_(value) {
  return String(value == null ? '' : value).replace(/[０-９]/g, function(ch) {
    return String.fromCharCode(ch.charCodeAt(0) - 0xFF10 + 0x30);
  });
}

function normSpace_(value) {
  return trimText_(value).replace(/\s+/g, ' ');
}

// classifyClient_ 와 parseRoadRegion_ 는 index.html 의 classifyClient, parseRoadRegion 과
// 같은 규칙이어야 합니다. 한쪽만 고치지 마세요. IP 와 User-Agent 원문은 저장하지 않습니다.
var SIDO_NAMES_ = [
  '강원특별자치도', '전북특별자치도', '제주특별자치도', '세종특별자치시',
  '서울특별시', '부산광역시', '대구광역시', '인천광역시', '광주광역시', '대전광역시', '울산광역시',
  '충청북도', '충청남도', '전라북도', '전라남도', '경상북도', '경상남도',
  '경기도', '강원도'
];

function parseRoadRegion_(road) {
  var text = normSpace_(road);
  var sido = '';
  var i;
  for (i = 0; i < SIDO_NAMES_.length; i++) {
    var name = SIDO_NAMES_[i];
    if (text.indexOf(name) !== 0) continue;
    if (text.length !== name.length && text.charAt(name.length) !== ' ') continue;
    sido = name;
    break;
  }
  if (!sido) return null;
  var sigungu = '';
  if (sido !== '세종특별자치시') {
    var tokens = text.slice(sido.length).trim().split(' ');
    var parts = [];
    for (i = 0; i < tokens.length && parts.length < 3; i++) {
      if (!isSigunguToken_(tokens[i])) break;
      parts.push(tokens[i]);
    }
    sigungu = parts.join(' ');
  }
  return { sido: sido, sigungu: sigungu };
}

function isSigunguToken_(token) {
  if (!token || token.length < 2 || token.length > 10) return false;
  if (!/^[가-힣]+$/.test(token)) return false;
  var last = token.charAt(token.length - 1);
  return last === '시' || last === '군' || last === '구';
}

function classifyClient_(ua) {
  var text = normSpace_(ua);
  if (!text || text.length > 1000 || /[<>]/.test(text)) return null;
  return {
    os: classifyOs_(text),
    device: classifyDevice_(text),
    browser: classifyBrowser_(text)
  };
}

function classifyOs_(ua) {
  if (/iPad/.test(ua)) return 'iPadOS';
  if (/iPhone|iPod/.test(ua)) return 'iOS';
  if (/Android/.test(ua)) return 'Android';
  if (/Windows|Win64|Win32/.test(ua)) return 'Windows';
  if (/CrOS|Chrome OS/.test(ua)) return 'ChromeOS';
  if (/Macintosh|Mac OS X/.test(ua)) return 'macOS';
  if (/Linux/.test(ua)) return 'Linux';
  return '기타';
}

function classifyDevice_(ua) {
  if (/iPad/.test(ua)) return '태블릿';
  if (/iPhone|iPod/.test(ua)) return '모바일';
  if (/Android/.test(ua)) return /Mobile/.test(ua) ? '모바일' : '태블릿';
  if (/Mobile/.test(ua)) return '모바일';
  return '데스크톱';
}

function classifyBrowser_(ua) {
  if (/Edg(?:e|A|iOS)?\//.test(ua)) return 'Edge';
  if (/SamsungBrowser\//.test(ua)) return 'Samsung Internet';
  if (/OPR\/|Opera\//.test(ua)) return 'Opera';
  if (/Firefox\/|FxiOS\//.test(ua)) return 'Firefox';
  if (/Whale\//.test(ua)) return 'Whale';
  if (/KAKAOTALK/i.test(ua)) return '카카오톡';
  if (/Chrome\/|CriOS\//.test(ua)) return 'Chrome';
  if (/Safari\//.test(ua)) return 'Safari';
  return '기타';
}

function seoulParts_(ms) {
  var fmt = new Intl.DateTimeFormat('en-US', {
    timeZone: TIMEZONE,
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
    hourCycle: 'h23'
  });
  var map = {};
  fmt.formatToParts(new Date(ms)).forEach(function(part) {
    if (part.type !== 'literal') map[part.type] = part.value;
  });
  if (map.hour === '24') map.hour = '00';
  return map;
}

function seoulText_(ms) {
  var p = seoulParts_(ms);
  return p.year + '-' + p.month + '-' + p.day + ' ' + p.hour + ':' + p.minute + ':' + p.second;
}

function makeOrderId_(ms) {
  var p = seoulParts_(ms);
  var alphabet = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  var out = '';
  var i;
  for (i = 0; i < 4; i++) out += alphabet.charAt(Math.floor(Math.random() * alphabet.length));
  return 'WB' + String(p.year).slice(2) + p.month + p.day + '-' + out;
}

// ── 시트에 한 줄 추가 ────────────────────────────────

function appendOrder_(order) {
  var ss;
  try { ss = SpreadsheetApp.openById(SPREADSHEET_ID); }
  catch (err) { fail_('주문 시트를 열 수 없습니다. 배포 계정이 스프레드시트를 수정할 수 있는지 확인해 주세요.'); }
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) fail_('주문 관리 탭을 찾지 못했습니다.');
  var map = headerMap_(sh);
  var cols = {};
  var i;
  for (i = 0; i < HEADER_FIELDS.length; i++) {
    var field = HEADER_FIELDS[i];
    var col = findCol_(map, field.names);
    if (!col) fail_('시트 헤더 [' + field.names[0] + '] 를 찾지 못했습니다.');
    cols[field.key] = col;
  }
  for (i = 0; i < EXTRA_FIELDS.length; i++) {
    var extra = EXTRA_FIELDS[i];
    cols[extra.key] = ensureHeader_(sh, map, extra.names);
  }
  var row = nextRow_(sh, cols.submittedAt);
  ensureRow_(sh, row);
  var id = uniqueId_(sh, cols.answerId, order.submittedMs);
  var width = Math.max(sh.getLastColumn(), 1);
  var colKeys = Object.keys(cols);
  for (i = 0; i < colKeys.length; i++) width = Math.max(width, cols[colKeys[i]]);
  var values = [];
  for (i = 0; i < width; i++) values.push('');
  function put(key, value) {
    var col = cols[key];
    if (col > width) fail_('시트 헤더 위치를 확인하지 못했습니다.');
    values[col - 1] = value;
  }
  put('submittedAt', new Date(order.submittedMs));
  put('status', '신규');
  put('payer', order.payer);
  put('phone', order.phone);
  put('books', order.books);
  put('amount', order.amount);
  put('zip', order.zip);
  put('address', order.address);
  put('os', order.os);
  put('device', order.device);
  put('browser', order.browser);
  put('sido', order.sido);
  put('sigungu', order.sigungu);
  put('platform', order.platform);
  put('memo', order.memo);
  put('answerId', id);
  put('resultUrl', '');
  put('payerRaw', order.payerRaw);
  put('booksRaw', order.booksRaw);

  ['payer', 'phone', 'books', 'zip', 'address', 'platform', 'memo', 'answerId', 'resultUrl', 'payerRaw', 'booksRaw', 'os', 'device', 'browser', 'sido', 'sigungu'].forEach(function(key) {
    sh.getRange(row, cols[key]).setNumberFormat('@');
  });
  sh.getRange(row, cols.submittedAt).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  sh.getRange(row, cols.amount).setNumberFormat('#,##0');
  sh.getRange(row, 1, 1, width).setValues([values]);
  sh.getRange(row, cols.submittedAt).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  sh.getRange(row, cols.amount).setNumberFormat('#,##0');
  ['payer', 'phone', 'books', 'zip', 'address', 'platform', 'memo', 'answerId', 'payerRaw', 'booksRaw', 'os', 'device', 'browser', 'sido', 'sigungu'].forEach(function(key) {
    sh.getRange(row, cols[key]).setNumberFormat('@');
    sh.getRange(row, cols[key]).setValue(values[cols[key] - 1]);
  });

  return {
    id: id,
    amount: order.amount,
    books: order.books,
    payerName: order.payerName,
    payer: order.payer,
    phone: order.phone,
    zip: order.zip,
    address: order.address,
    os: order.os,
    device: order.device,
    browser: order.browser,
    sido: order.sido,
    sigungu: order.sigungu,
    platform: order.platform,
    memo: order.memo,
    submittedText: seoulText_(order.submittedMs)
  };
}

function notifyOwner_(saved) {
  try {
    var owner = '';
    try { owner = Session.getEffectiveUser().getEmail(); } catch (e) { owner = ''; }
    if (!owner) return false;
    var amountText = Number(saved.amount).toLocaleString('ko-KR');
    var body = [
      '새 주문이 접수되었습니다.',
      '',
      '주문 번호: ' + saved.id,
      '제출일시: ' + saved.submittedText + ' (서울)',
      '주문자(예금주): ' + saved.payer,
      '입금자명: ' + saved.payerName,
      '연락처: ' + saved.phone,
      '교재: ' + saved.books,
      '금액: ' + amountText + '원',
      '우편번호: ' + saved.zip,
      '주소: ' + saved.address,
      '시도: ' + saved.sido,
      '시군구: ' + (saved.sigungu || '(없음)'),
      '운영체제: ' + saved.os,
      '기기: ' + saved.device,
      '브라우저: ' + saved.browser,
      '유입 플랫폼: ' + saved.platform,
      '메모: ' + saved.memo,
      '',
      '입금 계좌: ' + BANK_NAME + ' ' + BANK_ACCOUNT,
      '제출 시각 전후 ' + DEPOSIT_WINDOW_MINUTES + '분 안의 입금만 이 주문으로 확인합니다.',
      '입금자명과 금액이 같아야 합니다. 배송비는 착불입니다.'
    ].join('\n');
    MailApp.sendEmail(owner, '[황소 워크북] 새 주문 ' + saved.id, body);
    return true;
  } catch (err) {
    return false;
  }
}

function headerMap_(sh) {
  var last = Math.max(sh.getLastColumn(), 1);
  var row = sh.getRange(1, 1, 1, last).getValues()[0];
  var map = {};
  var i;
  for (i = 0; i < row.length; i++) {
    var name = String(row[i] == null ? '' : row[i]).trim();
    if (!name || name.length > 40 || /[\r\n]/.test(name)) continue;
    var key = compact_(name);
    if (!map[key]) map[key] = i + 1;
  }
  return map;
}

function findCol_(map, names) {
  var i;
  for (i = 0; i < names.length; i++) {
    var col = map[compact_(names[i])];
    if (col) return col;
  }
  return 0;
}

function ensureHeader_(sh, map, names) {
  var found = findCol_(map, names);
  if (found) return found;
  var lastHeader = 0;
  var keys = Object.keys(map);
  var i;
  for (i = 0; i < keys.length; i++) {
    if (map[keys[i]] > lastHeader) lastHeader = map[keys[i]];
  }
  var col = lastHeader + 1;
  var header = names[0];
  sh.getRange(1, col).setNumberFormat('@');
  sh.getRange(1, col).setValue(header);
  map[compact_(header)] = col;
  return col;
}

function nextRow_(sh, col) {
  var last = sh.getLastRow();
  if (last < 2) return 2;
  var height = Math.min(last, SCAN_LAST_ROW) - 1;
  var values = sh.getRange(2, col, height, 1).getValues();
  var lastData = 1;
  var i;
  for (i = 0; i < values.length; i++) {
    var v = values[i][0];
    if (v == null || v === '') continue;
    lastData = i + 2;
  }
  return Math.min(lastData + 1, SCAN_LAST_ROW + 1);
}

function uniqueId_(sh, col, ms) {
  var taken = {};
  var last = sh.getLastRow();
  if (col && last >= 2) {
    var height = Math.min(last, SCAN_LAST_ROW) - 1;
    var values = sh.getRange(2, col, height, 1).getValues();
    var i;
    for (i = 0; i < values.length; i++) {
      var v = String(values[i][0] == null ? '' : values[i][0]).trim();
      if (v) taken[v] = true;
    }
  }
  var n;
  for (n = 0; n < 8; n++) {
    var id = makeOrderId_(ms);
    if (!taken[id]) return id;
  }
  fail_('주문 번호를 만들지 못했습니다. 다시 시도해 주세요.');
}

function ensureRow_(sh, row) {
  if (sh.getMaxRows() < row) sh.insertRowsAfter(sh.getMaxRows(), row - sh.getMaxRows());
}

function compact_(value) {
  return String(value || '').replace(/\s+/g, '');
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function fail_(message) {
  var err = new Error(message);
  err.orderSafe = true;
  throw err;
}

function publicError_(err) {
  if (err && err.orderSafe) return String(err.message);
  return '주문을 저장하지 못했습니다. 잠시 후 다시 시도해 주세요.';
}
