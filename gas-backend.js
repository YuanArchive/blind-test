/**
 * Google Apps Script 백엔드
 * ─────────────────────────────────────────────────────────
 * 배포 방법:
 *  1. https://script.google.com 접속
 *  2. 새 프로젝트 만들기
 *  3. 이 파일 전체를 붙여넣기 (기존 코드 덮어쓰기)
 *  4. 상단 메뉴: 배포 → 새 배포
 *  5. 유형: 웹 앱
 *     - 다음 사용자로 실행: 나 (본인)
 *     - 액세스 권한: 모든 사용자 (익명)
 *  6. 배포 → 웹 앱 URL 복사
 *  7. index.html 상단 CONFIG.scriptUrl 에 붙여넣기
 * ─────────────────────────────────────────────────────────
 */

function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data  = JSON.parse(e.postData.contents);

    // 헤더 행이 없으면 추가
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        '세션ID', '타임스탬프', '트랙ID', '트랙명',
        'A=오리지널', '투표(A/B/same)', '선호버전', 'UserAgent'
      ]);
    }

    // 각 트랙 결과 저장
    data.results.forEach(function(r) {
      sheet.appendRow([
        data.sessionId,
        data.timestamp,
        r.trackId,
        r.title,
        r.aIsOriginal,
        r.userVote,
        r.preferred,
        data.userAgent || ''
      ]);
    });

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', msg: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// 테스트용 (Apps Script 에디터에서 실행하면 동작 확인 가능)
function testPost() {
  var fake = {
    postData: {
      contents: JSON.stringify({
        sessionId: 'TEST01',
        timestamp: new Date().toISOString(),
        userAgent: 'TestAgent',
        results: [
          { trackId: 1, title: 'Track 01', aIsOriginal: true,  userVote: 'A', preferred: 'original'    },
          { trackId: 2, title: 'Track 02', aIsOriginal: false, userVote: 'B', preferred: 'remastered'  },
          { trackId: 3, title: 'Track 03', aIsOriginal: true,  userVote: 'same', preferred: 'same'     },
        ]
      })
    }
  };
  Logger.log(doPost(fake).getContent());
}
