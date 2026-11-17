/**
 * 투미 마케팅 요청 수신기 (구글 Apps Script)
 * 구글 시트에 붙여서 웹앱으로 배포합니다. 설치 방법은 README.md 참고.
 *
 * - 블로그 요청은 '블로그' 시트, 쇼츠 요청은 '쇼츠' 시트에 한 줄씩 쌓입니다.
 * - 쇼츠 요청에는 투미숏츠메이커용 JSON이 함께 저장됩니다.
 *   (그 칸을 복사해서 투미숏츠메이커\매물\구역명.json 으로 저장하면 바로 제작 가능)
 * - 새 요청이 오면 아래 NOTIFY_EMAIL 로 메일이 갑니다.
 */

// 알림 받을 메일. 비워 두면 이 스크립트를 만든 계정으로 갑니다.
const NOTIFY_EMAIL = '';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const type = data.type === 'shorts' ? '쇼츠' : '블로그';
    const fields = data.fields || {};
    if (!fields['담당자'] || !fields['구역명']) throw new Error('필수 항목 누락');

    const row = { '접수시각': new Date(), '상태': '접수' };
    Object.keys(fields).forEach(function (k) { row[k] = String(fields[k]).slice(0, 5000); });
    if (data.scenes) {
      row['장면 자막'] = data.scenes.map(function (s, i) { return (i + 1) + '. ' + s['그림'] + ' ' + s['자막']; }).join('\n');
    }
    if (data.shortsJson) row['숏츠메이커 JSON'] = JSON.stringify(data.shortsJson, null, 2);

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try { appendRow_(type, row); } finally { lock.releaseLock(); }

    const to = NOTIFY_EMAIL || Session.getEffectiveUser().getEmail();
    MailApp.sendEmail({
      to: to,
      subject: '[투미 ' + type + ' 요청] ' + fields['구역명'] + ' / ' + fields['담당자'],
      body: (data.text || '') + '\n\n시트: ' + SpreadsheetApp.getActiveSpreadsheet().getUrl()
    });

    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

// 헤더 이름으로 칸을 맞춰 넣습니다. 폼에 새 항목이 생기면 열이 자동으로 추가됩니다.
function appendRow_(sheetName, row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  let headers = sh.getLastColumn() ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] : [];
  Object.keys(row).forEach(function (k) {
    if (headers.indexOf(k) === -1) {
      headers.push(k);
      sh.getRange(1, headers.length).setValue(k).setFontWeight('bold');
    }
  });
  if (sh.getFrozenRows() === 0) sh.setFrozenRows(1);
  sh.appendRow(headers.map(function (h) { return row[h] !== undefined ? row[h] : ''; }));
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// 배포 전 테스트용: 편집기에서 이 함수를 한 번 실행하면 권한 승인 창이 뜨고, 시트에 테스트 줄이 들어갑니다.
function testRequest() {
  const res = doPost({ postData: { contents: JSON.stringify({
    type: 'blog',
    fields: { '담당자': '테스트', '업무폰 번호': '010-0000-0000', '구역명': '테스트구역', '사업진행 상세현황': '테스트', '구역 특장점': '테스트', '사실 확인': '확인함' },
    text: '[테스트 요청]'
  }) } });
  Logger.log(res.getContent());
}
