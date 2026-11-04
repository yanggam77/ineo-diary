/**
 * 『별일 없는 날에 일기 쓰기』 설문 · 서평단 신청 웹앱
 * - 이 스크립트는 응답을 받을 '구글 시트'에 연결(확장 프로그램 > Apps Script)해서 씁니다.
 * - Index.html 파일과 함께 배포하면, 배포 URL 하나로 설문 페이지가 열리고 응답이 시트에 쌓입니다.
 */

var SHEET_NAME = '응답';
var HEADERS = [
  '제출 시각', '학년', '일기 쓰는 빈도', '좋아하는 일기', '재미/어려움',
  '재미있는 이유', '어려운 이유', '직접 쓴 이유',
  '서평단', '보호자 성함', '연락처', '이메일', '주소', '서평 게시 주소', '개인정보 동의',
  '중복 의심', '기기 정보'
];

// 설문 페이지 열기
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('일기 쓰기, 어때요? — 별일 없는 날에 일기 쓰기')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 페이지(google.script.run)에서 호출
function submit(data) {
  return save_(data);
}

// 다른 곳에 HTML을 올렸을 때(fetch POST) 받는 입구
function doPost(e) {
  var data = JSON.parse(e.postData.contents);
  save_(data);
  return ContentService.createTextOutput('ok');
}

function save_(d) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = getSheet_();
    var phone = clean_(d.phone);
    var dup = '';
    if (phone) {
      var last = sh.getLastRow();
      if (last > 1) {
        var phones = sh.getRange(2, 11, last - 1, 1).getValues();
        for (var i = 0; i < phones.length; i++) {
          if (String(phones[i][0]).replace(/\D/g, '') === phone.replace(/\D/g, '')) { dup = '중복'; break; }
        }
      }
    }
    sh.appendRow([
      new Date(),
      clean_(d.grade), clean_(d.freq), clean_(d.diaryType), clean_(d.feel),
      clean_(d.funReasons), clean_(d.hardReasons), clean_(d.reasonText),
      clean_(d.reviewer), clean_(d.parentName), phone ? "'" + phone : '', clean_(d.email),
      clean_(d.address), clean_(d.sns), clean_(d.consent),
      dup, clean_(d.ua)
    ]);
    return 'ok';
  } finally {
    lock.releaseLock();
  }
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#FBD3D4');
  }
  return sh;
}

// 수식 주입 방지 + 길이 제한
function clean_(v) {
  var s = String(v == null ? '' : v).slice(0, 1000);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}
