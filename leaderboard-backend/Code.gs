/**
 * 신입사원 메이커 — 공유 리더보드 백엔드 (구글 시트 + Apps Script)
 *
 * 이 스크립트가 붙어있는 구글시트가 곧 데이터베이스입니다.
 * 시트에 "leaderboard" 탭을 자동으로 만들고, 게임에서 오는 점수를 그 탭에 쌓습니다.
 *
 * 설치 방법:
 * 1) 구글 드라이브에서 새 스프레드시트를 하나 만든다 (이름 자유, 예: "신입사원메이커_리더보드").
 * 2) 메뉴 [확장 프로그램] > [Apps Script] 를 연다.
 * 3) 기본으로 있는 코드를 지우고 이 파일 내용 전체를 붙여넣는다.
 * 4) 저장(Ctrl+S) 후 상단 [배포] > [새 배포] 클릭.
 * 5) 유형에서 톱니바퀴를 눌러 "웹 앱" 선택.
 * 6) "실행할 사용자" = 나(본인), "액세스 권한이 있는 사용자" = 전체(익명 사용자 포함) 로 설정.
 * 7) [배포] 클릭 → 처음엔 권한 승인 화면이 뜬다 → 본인 계정으로 승인.
 * 8) 발급된 웹 앱 URL(.../exec 로 끝남)을 복사한다.
 * 9) index.html 안의  const LB_API_URL = "";  괄호 안에 그 URL을 붙여넣는다.
 *    (LB_API_URL을 채워야 실제로 공유 리더보드가 켜집니다. 비어있으면 기존처럼 로컬 저장만 됨)
 *
 * 코드를 나중에 수정했다면: [배포] > [배포 관리] > 연필 아이콘 > 새 버전으로 다시 배포해야
 * 수정사항이 실제로 반영됩니다(URL은 그대로 유지됨).
 */

var SHEET_NAME = "leaderboard";

function _sheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(["id", "eventCode", "nick", "score", "grade", "char", "time", "ts"]);
  }
  return sh;
}

function _json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function _header(sh) {
  return sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
}

// 참가자가 게임을 마치면 호출됨: 점수 한 줄 추가 + 해당 대회 코드 안에서의 등수 계산
function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    if (body.action !== "submit") return _json({ ok: false, error: "unknown action" });

    var sh = _sheet();
    sh.appendRow([
      String(body.id || "").slice(0, 64),
      String(body.eventCode || "").slice(0, 64),
      String(body.nick || "익명").slice(0, 20),
      Number(body.score) || 0,
      String(body.grade || ""),
      String(body.char || "").slice(0, 40),
      Number(body.time) || 0,
      new Date()
    ]);

    var data = sh.getDataRange().getValues();
    var h = data[0];
    var iEvent = h.indexOf("eventCode"), iScore = h.indexOf("score"),
        iId = h.indexOf("id"), iTs = h.indexOf("ts");

    var rows = data.slice(1).filter(function (r) { return r[iEvent] === body.eventCode; });
    rows.sort(function (a, b) {
      if (b[iScore] !== a[iScore]) return b[iScore] - a[iScore];
      return new Date(a[iTs]) - new Date(b[iTs]);
    });
    var rank = rows.findIndex(function (r) { return r[iId] === body.id; }) + 1;

    return _json({ ok: true, rank: rank, total: rows.length });
  } catch (err) {
    return _json({ ok: false, error: String(err) });
  }
}

// 리더보드 조회: ?action=list&eventCode=XXXX&limit=10
function doGet(e) {
  try {
    var action = e.parameter.action || "list";
    if (action !== "list") return _json({ ok: false, error: "unknown action" });

    var eventCode = e.parameter.eventCode || "";
    var limit = Math.max(1, Math.min(50, parseInt(e.parameter.limit, 10) || 10));

    var sh = _sheet();
    var data = sh.getDataRange().getValues();
    var h = data[0];
    var iEvent = h.indexOf("eventCode"), iScore = h.indexOf("score"),
        iNick = h.indexOf("nick"), iGrade = h.indexOf("grade"),
        iChar = h.indexOf("char"), iId = h.indexOf("id"), iTs = h.indexOf("ts");

    var rows = data.slice(1).filter(function (r) { return r[iEvent] === eventCode; });
    rows.sort(function (a, b) {
      if (b[iScore] !== a[iScore]) return b[iScore] - a[iScore];
      return new Date(a[iTs]) - new Date(b[iTs]);
    });

    var top = rows.slice(0, limit).map(function (r) {
      return { id: r[iId], nick: r[iNick], score: r[iScore], grade: r[iGrade], char: r[iChar] };
    });

    return _json({ ok: true, entries: top, total: rows.length });
  } catch (err) {
    return _json({ ok: false, error: String(err) });
  }
}
