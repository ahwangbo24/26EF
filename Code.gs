/**
 * 26EF VAS 설문 - Google Sheets 실시간 수집용 Apps Script
 * 15문항(식욕/충동성/정서·스트레스/피로·각성·집중/과제동기) VAS 구조 기준.
 *
 * 대상 시트:
 * https://docs.google.com/spreadsheets/d/1NE-bOcQhVB0MP_MkexHNVT6OGPTTOPESCM5_GWGNGfY/edit?gid=0#gid=0
 *
 * 배포 방법:
 * 1. 위 Google Sheet를 연다.
 * 2. 상단 메뉴 확장 프로그램(Extensions) > Apps Script 클릭.
 * 3. 기본 코드(Code.gs)를 모두 지우고 이 파일 내용을 붙여넣기.
 * 4. 저장 후 우측 상단 배포(Deploy) > 새 배포(New deployment) 클릭.
 * 5. 유형 선택(⚙) > 웹 앱(Web app) 선택.
 *    - 실행 계정(Execute as): 나(Me)
 *    - 액세스 권한(Who has access): 모든 사용자(Anyone)
 * 6. 배포(Deploy) 클릭 → 권한 승인(본인 Google 계정으로 승인) → 생성된
 *    웹 앱 URL(.../exec 로 끝남)을 복사한다.
 * 7. 이 URL을 index.html의 GOOGLE_SCRIPT_URL 상수에 붙여넣는다.
 *
 * 코드를 수정한 뒤에는 "새 배포"가 아니라 기존 배포에서 "관리(Manage
 * deployments) > 편집(연필 아이콘) > 버전: 새 버전 > 배포"로 갱신해야
 * 같은 URL이 유지된다.
 */

var HEADERS = [
  '제출시각(서버)', '연구등록번호', '이름', '응답날짜',
  '배고픔VAS', '포만감VAS', '배부름VAS', '추가섭취가능성VAS',
  '충동억제어려움VAS', '즉각만족욕구VAS', '신중판단어려움VAS', '기다림어려움VAS',
  '기분(불쾌-유쾌)VAS', '불안VAS', '스트레스VAS',
  '흥분도VAS', '집중도VAS', '피로도VAS',
  '과제동기VAS',
  '자유기술'
];

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Responses') || ss.insertSheet('Responses');

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
    }

    var data = JSON.parse(e.postData.contents);
    var val = function (x) { return (x !== undefined && x !== null) ? x : ''; };

    sheet.appendRow([
      Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss'),
      val(data.regNumber),
      val(data.participantName),
      val(data.responseDate),
      val(data.vHunger),
      val(data.vSatiety),
      val(data.vFullness),
      val(data.vMoreCapacity),
      val(data.vImpulseControl),
      val(data.vImmediateGrat),
      val(data.vJudgment),
      val(data.vWaiting),
      val(data.vMood),
      val(data.vAnxiety),
      val(data.vStress),
      val(data.vArousal),
      val(data.vAttention),
      val(data.vFatigue),
      val(data.vMotivation),
      val(data.freeText)
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', message: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  return ContentService.createTextOutput('OK - 26EF VAS survey endpoint is running.');
}
