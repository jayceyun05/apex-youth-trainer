// ── 타겟 스프레드시트 ID 고정 ──
// 파일제목: APEX 유소년 트레이닝 베타 피드백(응답)
var SPREADSHEET_ID = '1pIyjoneh2V3O6LnPOxsZDGMs2m1un0_fFtgPEcTrSHg';

// ── 시트(탭) 이름 상수 지정 (정확한 띄어쓰기 반영) ──
var SHEET_CONSENT  = '설문지 응답 시트1'; // 부모 동의 기록 탭
var SHEET_FEEDBACK = '설문지 응답 시트2';  // 구글 폼 피드백(응답) 데이터 탭
var SHEET_SUMMARY  = '요약대시보드';       // 이 두 데이터를 분석해서 자동으로 그려주는 요약 화면 탭

// ── 1. 보호자 동의 로그 수신 ──────────────────────
function doPost(e) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_CONSENT);
  
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_CONSENT);
  }

  // 첫 줄(헤더)이 비어있으면 컬럼명 자동 생성
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['No.', '서버시각', '동의유형', '동의버전', '세션토큰']);
    sheet.getRange(1, 1, 1, 5)
      .setFontWeight('bold')
      .setBackground('#4472C4')
      .setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
  }

  var data = JSON.parse(e.postData.contents);
  var no   = Math.max(sheet.getLastRow(), 1); 
  
  sheet.appendRow([
    no,
    new Date(),
    data.consentType    || '',
    data.consentVersion || '',
    data.sessionToken   || ''
  ]);

  try {
    buildSummary(); 
  } catch(err) { }

  return ContentService
    .createTextOutput('{"ok":true}')
    .setMimeType(ContentService.MimeType.JSON);
}

// ── 2. 요약 대시보드 시트 생성/갱신 ────────────────
function buildSummary() {
  var ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sumSh   = ss.getSheetByName(SHEET_SUMMARY);
  if (!sumSh) {
    sumSh = ss.insertSheet(SHEET_SUMMARY, 0); 
  }
  var fbSh    = ss.getSheetByName(SHEET_FEEDBACK);
  var conSh   = ss.getSheetByName(SHEET_CONSENT);

  sumSh.clearContents();
  sumSh.clearFormats();

  var now = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
  var row = 1;

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
  // 동의 시트 데이터가 있는지 확인 (헤더 제외 최소 1줄 이상)
  if (conSh && conSh.getLastRow() > 1) {
    var conData = conSh.getDataRange().getValues();
    conTotal = Math.max(conData.length - 1, 0);
    conRows  = conData.slice(1);
  } else if (conSh && conSh.getLastRow() === 1) {
    // 헤더만 있는 경우
    conTotal = 0;
    conRows = [];
  }

  sumSh.getRange(row, 1).setValue('총 동의 건수');
  sumSh.getRange(row, 2).setValue(conTotal + ' 건').setFontWeight('bold');
  row++;

  if (conRows.length > 0) {
    sumSh.getRange(row, 1).setValue('최근 5건');
    row++;
    sumSh.getRange(row, 1, 1, 4).setValues([['No.', '서버시각', '동의유형', '세션토큰(앞8자)']])
      .setFontWeight('bold').setBackground('#EEF2F7');
    row++;
    var recent5 = conRows.slice(-5).reverse();
    recent5.forEach(function(r) {
      var ts = r[1] ? Utilities.formatDate(new Date(r[1]), 'Asia/Seoul', 'MM-dd HH:mm') : '';
      sumSh.getRange(row, 1, 1, 4).setValues([[r[0], ts, r[2], String(r[4]).slice(0, 8)]]);
      row++;
    });
  }
  row++;

  // ── 섹션 B: 피드백 현황 ──
  sumSh.getRange(row, 1, 1, 4).merge()
    .setValue('📝 섹션 B — 피드백 응답 현황')
    .setFontWeight('bold').setBackground('#E2EFDA');
  row++;

  if (!fbSh || fbSh.getLastRow() <= 1) {
    var msg = !fbSh ? '⚠️ 피드백 시트("' + SHEET_FEEDBACK + '")를 찾을 수 없습니다.' : '아직 접수된 피드백 응답이 없습니다.';
    sumSh.getRange(row, 1, 1, 4).merge()
      .setValue(msg)
      .setFontColor('#C00000');
    row++;
  } else {
    var fbData = fbSh.getDataRange().getValues();
    var fbRows = fbData.slice(1); 
    var fbTotal = fbRows.length;
    var up = 0, down = 0, none = 0;
    var sportMap = {};

    fbRows.forEach(function(r) {
      // 구글폼 구조에 따라 종목, 적합도 컬럼 인덱스(1, 2) 확인
      var sport  = String(r[1] || '').trim();  
      var rating = String(r[2] || '').trim();  

      if (rating.indexOf('\uac1c\uc9dc') !== -1 || rating.indexOf('\ud83d\udc4d') !== -1) up++;
      else if (rating.indexOf('\uc544\uc26c') !== -1 || rating.indexOf('\ud83d\udc4e') !== -1) down++;
      else none++;

      if (sport) sportMap[sport] = (sportMap[sport] || 0) + 1;
    });

    sumSh.getRange(row, 1).setValue('총 피드백 수');
    sumSh.getRange(row, 2).setValue(fbTotal + ' 건').setFontWeight('bold');
    row++;
    sumSh.getRange(row, 1, 1, 4).setValues([['👍 적합해요', '👎 아쉬워요', '미선택', '만족률']])
      .setFontWeight('bold').setBackground('#EEF2F7');
    row++;
    var satRate = fbTotal > 0 ? Math.round(up / fbTotal * 100) + '%' : '-';
    sumSh.getRange(row, 1, 1, 4).setValues([[up, down, none, satRate]]);
    sumSh.getRange(row, 4).setFontWeight('bold').setFontColor('#0E7A39');
    row += 2;

    sumSh.getRange(row, 1, 1, 2).setValues([['종목', '응답 수']])
      .setFontWeight('bold').setBackground('#EEF2F7');
    row++;
    var sports = Object.keys(sportMap).sort(function(a, b) {
      return sportMap[b] - sportMap[a];
    });
    sports.forEach(function(s) {
      sumSh.getRange(row, 1, 1, 2).setValues([[s, sportMap[s]]]);
      row++;
    });
    row++;

    sumSh.getRange(row, 1, 1, 4).merge()
      .setValue('📌 최근 의견 (최신 5건)')
      .setFontWeight('bold').setBackground('#FFF2CC');
    row++;
    sumSh.getRange(row, 1, 1, 4).setValues([['시각', '종목', '적합도', '의견']])
      .setFontWeight('bold').setBackground('#EEF2F7');
    row++;
    var recent = fbRows.slice(-5).reverse();
    recent.forEach(function(r) {
      var ts  = r[0] ? Utilities.formatDate(new Date(r[0]), 'Asia/Seoul', 'MM-dd HH:mm') : '';
      var txt = String(r[3] || '').replace(/\[sid:[^\]]+\]/g, '').trim(); 
      sumSh.getRange(row, 1, 1, 4).setValues([[ts, r[1], r[2], txt]]);
      row++;
    });
  }

  sumSh.autoResizeColumns(1, 4);
}
