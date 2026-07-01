// ══════════════════════════════════════════════════════════════
//  APEX 유소년 트레이닝 — Google Apps Script 백엔드
//  스프레드시트: APEX 유소년 트레이닝 베타 피드백(응답)
//  탭 구성:
//    ① 부모동의기록   — 앱에서 doPost 로 기록
//    ② 설문지 응답 시트1 — 구글 폼 연동(자동 생성)
//    ③ 설문지 응답 시트2 — 대시보드(buildSummary 자동 갱신)
//
//  ★ 배포 시 반드시 확인:
//    - 실행 계정 : 나(소유자)
//    - 액세스 권한 : 모든 사용자(익명 포함)
//    - 코드 변경 후 반드시 [새 버전 배포] 눌러야 반영됨
// ══════════════════════════════════════════════════════════════

var SPREADSHEET_ID = '1Oi7jzgWMOizO1Xla5NhjVLgA24U99fzzPLdIG8TsajA';

var SHEET_CONSENT  = '부모동의기록';       // 앱 → doPost
var SHEET_FEEDBACK = '설문지 응답 시트1';  // 구글 폼 자동 연결 탭 이름
var SHEET_SUMMARY  = '설문지 응답 시트2';  // buildSummary() 대시보드


// ── 연결 테스트 GET ─────────────────────────────────────────
// 브라우저에서 배포 URL 직접 열면 동작 확인 가능
function doGet(e) {
  var info = { ok: false, sheets: [] };
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    info.ok              = true;
    info.spreadsheetName = ss.getName();
    info.sheets          = ss.getSheets().map(function(s) { return s.getName(); });
    info.time            = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
  } catch(err) {
    info.error = err.message;
  }
  return ContentService
    .createTextOutput(JSON.stringify(info, null, 2))
    .setMimeType(ContentService.MimeType.JSON);
}


// ── 보호자 동의 로그 수신 POST ──────────────────────────────
function doPost(e) {
  var result = { ok: false };
  try {
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = getOrCreateSheet(ss, SHEET_CONSENT);

    // 헤더 없으면 초기화
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['No.', '서버시각', '동의유형', '동의버전', '세션토큰']);
      sheet.getRange(1, 1, 1, 5)
        .setFontWeight('bold')
        .setBackground('#4472C4')
        .setFontColor('#FFFFFF');
      sheet.setFrozenRows(1);
    }

    // ★ 핵심 수정: e.parameter(form-urlencoded) 우선 → JSON fallback
    var data = {};
    try {
      if (e && e.parameter && e.parameter.consentType) {
        // application/x-www-form-urlencoded (앱 v2+ 방식)
        data = e.parameter;
      } else if (e && e.postData && e.postData.contents) {
        // text/plain + JSON body (구버전 호환)
        data = JSON.parse(e.postData.contents);
      }
    } catch (_) { /* 파싱 실패 시 빈 객체로 기록 */ }

    var rowNo = Math.max(sheet.getLastRow(), 1); // 헤더 제외 순번
    sheet.appendRow([
      rowNo,
      new Date(),
      data.consentType    || '',
      data.consentVersion || '',
      data.sessionToken   || ''
    ]);

    result.ok  = true;
    result.row = sheet.getLastRow();

    try { buildSummary(); } catch (_) { /* 대시보드 실패해도 로그는 성공 */ }

  } catch (err) {
    result.error = err.message;
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}


// ── 헬퍼: 시트 없으면 생성 ─────────────────────────────────
function getOrCreateSheet(ss, name) {
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}


// ── 수동 연결 테스트 (Script Editor에서 직접 실행) ──────────
function testDoPost() {
  var mockE = {
    parameter: {
      consentType:    'parent_under14',
      consentVersion: 'parent-v1-2026-07',
      sessionToken:   'test-manual-001',
      clientTime:     new Date().toISOString()
    },
    postData: { contents: '' }
  };
  var res = doPost(mockE);
  Logger.log('결과: ' + res.getContent());
}


// ── 요약 대시보드 시트 갱신 ────────────────────────────────
function buildSummary() {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sumSh = getOrCreateSheet(ss, SHEET_SUMMARY);
  var fbSh  = ss.getSheetByName(SHEET_FEEDBACK);
  var conSh = ss.getSheetByName(SHEET_CONSENT);

  sumSh.clearContents();
  sumSh.clearFormats();

  var now = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
  var row = 1;

  // ── 타이틀 ──
  sumSh.getRange(row, 1, 1, 4).merge()
    .setValue('📊 APEX 유소년 트레이닝 — 응답 요약 대시보드')
    .setFontSize(14).setFontWeight('bold')
    .setBackground('#1F3864').setFontColor('#FFFFFF')
    .setHorizontalAlignment('center');
  row++;
  sumSh.getRange(row, 1, 1, 4).merge()
    .setValue('마지막 갱신: ' + now)
    .setFontSize(10).setFontColor('#666666')
    .setHorizontalAlignment('right');
  row += 2;

  // ── 섹션 A: 동의 현황 ──
  sumSh.getRange(row, 1, 1, 4).merge()
    .setValue('🔐 섹션 A — 보호자 동의 현황')
    .setFontWeight('bold').setBackground('#D9E1F2');
  row++;

  var conTotal = 0, conRows = [];
  if (conSh && conSh.getLastRow() > 1) {
    var conData = conSh.getDataRange().getValues();
    conTotal = conData.length - 1;
    conRows  = conData.slice(1);
  }

  sumSh.getRange(row, 1).setValue('총 동의 건수');
  sumSh.getRange(row, 2).setValue(conTotal + ' 건').setFontWeight('bold');
  row++;

  if (conRows.length > 0) {
    sumSh.getRange(row, 1).setValue('최근 5건');
    row++;
    sumSh.getRange(row, 1, 1, 4)
      .setValues([['No.', '서버시각', '동의유형', '세션토큰(앞8자)']])
      .setFontWeight('bold').setBackground('#EEF2F7');
    row++;
    conRows.slice(-5).reverse().forEach(function(r) {
      var ts = r[1] ? Utilities.formatDate(new Date(r[1]), 'Asia/Seoul', 'MM-dd HH:mm') : '';
      sumSh.getRange(row, 1, 1, 4).setValues([[r[0], ts, r[2], String(r[4]).slice(0, 8)]]);
      row++;
    });
  } else {
    sumSh.getRange(row, 1, 1, 4).merge()
      .setValue('아직 동의 기록이 없습니다. (부모동의 체크 후 자동 갱신)')
      .setFontColor('#999999');
    row++;
  }
  row++;

  // ── 섹션 B: 피드백 현황 ──
  sumSh.getRange(row, 1, 1, 4).merge()
    .setValue('📝 섹션 B — 피드백 응답 현황')
    .setFontWeight('bold').setBackground('#E2EFDA');
  row++;

  if (!fbSh || fbSh.getLastRow() <= 1) {
    sumSh.getRange(row, 1, 1, 4).merge()
      .setValue(!fbSh
        ? '⚠️ 구글 폼 → [응답] 탭 → 스프레드시트 아이콘 → "기존 스프레드시트 선택"으로 연결하면 이 탭이 생성됩니다.'
        : '아직 접수된 피드백이 없습니다.')
      .setFontColor('#C00000').setWrap(true);
    row++;
  } else {
    var fbData  = fbSh.getDataRange().getValues();
    var fbRows  = fbData.slice(1);
    var fbTotal = fbRows.length;
    var up = 0, down = 0, none = 0;
    var sportMap = {};

    fbRows.forEach(function(r) {
      var sport  = String(r[1] || '').trim();
      var rating = String(r[2] || '').trim();
      if (rating.indexOf('적합') !== -1 || rating.indexOf('👍') !== -1) up++;
      else if (rating.indexOf('아쉬') !== -1 || rating.indexOf('👎') !== -1) down++;
      else none++;
      if (sport) sportMap[sport] = (sportMap[sport] || 0) + 1;
    });

    sumSh.getRange(row, 1).setValue('총 피드백 수');
    sumSh.getRange(row, 2).setValue(fbTotal + ' 건').setFontWeight('bold');
    row++;
    sumSh.getRange(row, 1, 1, 4)
      .setValues([['👍 적합해요', '👎 아쉬워요', '미선택', '만족률']])
      .setFontWeight('bold').setBackground('#EEF2F7');
    row++;
    var satRate = fbTotal > 0 ? Math.round(up / fbTotal * 100) + '%' : '-';
    sumSh.getRange(row, 1, 1, 4).setValues([[up, down, none, satRate]]);
    sumSh.getRange(row, 4).setFontWeight('bold').setFontColor('#0E7A39');
    row += 2;

    sumSh.getRange(row, 1, 1, 2)
      .setValues([['종목', '응답 수']]).setFontWeight('bold').setBackground('#EEF2F7');
    row++;
    Object.keys(sportMap).sort(function(a, b) { return sportMap[b] - sportMap[a]; })
      .forEach(function(s) {
        sumSh.getRange(row, 1, 1, 2).setValues([[s, sportMap[s]]]);
        row++;
      });
    row++;

    sumSh.getRange(row, 1, 1, 4).merge()
      .setValue('📌 최근 의견 (최신 5건)')
      .setFontWeight('bold').setBackground('#FFF2CC');
    row++;
    sumSh.getRange(row, 1, 1, 4)
      .setValues([['시각', '종목', '적합도', '의견']]).setFontWeight('bold').setBackground('#EEF2F7');
    row++;
    fbRows.slice(-5).reverse().forEach(function(r) {
      var ts  = r[0] ? Utilities.formatDate(new Date(r[0]), 'Asia/Seoul', 'MM-dd HH:mm') : '';
      var txt = String(r[3] || '').replace(/\[sid:[^\]]+\]/g, '').trim();
      sumSh.getRange(row, 1, 1, 4).setValues([[ts, r[1], r[2], txt]]);
      row++;
    });
  }

  sumSh.autoResizeColumns(1, 4);
}
