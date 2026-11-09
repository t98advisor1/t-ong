/**
 * ==============================================================================
 * [교사용 업무수첩 & NEIS Help] 관리자 승인 구글 시트 전용 Apps Script (경량 배포용)
 * ==============================================================================
 * 구글 시트 URL: https://docs.google.com/spreadsheets/d/1Tzo4r78gaFLMyGJOOBM7hTrmIZtEBopFOFlWWs1VQp4/edit
 * 
 * [기능]
 * 1. 구글 로그인 시 접속 교사 자동 행 추가 (A열 체크박스, 신청일시, 성명, 학교, 이메일, 만료일)
 * 2. 관리자가 A열 체크박스 [V] 클릭 시 즉시 한 학기(6개월/180일) 승인
 * 3. 만료일(G열) 수정 시 해당 날짜까지 승인 연장
 * 4. 웹앱 실시간 승인 여부 조회 (check_access)
 *
 * [배포 방법 - 1분 완성]
 * 1. 관리자 구글 시트 상단 메뉴 [확장 프로그램] > [Apps Script] 클릭
 * 2. 기존 코드를 모두 지우고 이 스크립트를 붙여넣은 후 저장(Ctrl+S)
 * 3. 우측 상단 파란색 [배포] > [새 배포] 클릭
 *    - 유형 선택(톱니바퀴): [웹 앱]
 *    - 설명: 교사 승인 웹앱 v1
 *    - 다음 권한으로 앱 실행: [나 (내 구글 계정)]
 *    - 액세스 권한: [모든 사용자 (Anyone)]  <-- 필수!
 * 4. [배포] 클릭 후 승인(액세스 허용) 진행
 * 5. 발급된 웹앱 URL (https://script.google.com/macros/s/.../exec)을
 *    관리자 대시보드(admin.html)의 [연동 주소 저장]에 붙여넣기만 하면 끝!
 * ==============================================================================
 */

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || "check_access";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getOrCreateTargetSheet(ss);

  if (action === "check_access") {
    var email = (e.parameter.email || "").trim().toLowerCase();
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var rowEmail = "";
      for (var c = 0; c < row.length; c++) {
        if (String(row[c]).toLowerCase().includes("@")) {
          rowEmail = String(row[c]).trim().toLowerCase();
          break;
        }
      }
      if (rowEmail === email) {
        var isApproved = (row[0] === true || String(row[0]).toLowerCase() === "true" || row[0] === "승인");
        var expStr = row[6] ? formatDateVal(row[6]) : "";
        return jsonOut({
          found: true,
          approved: isApproved,
          name: row[2] || email.split("@")[0],
          school: row[3] || "고등학교",
          expiresAt: expStr,
          status: isApproved ? "승인" : "대기"
        });
      }
    }
    return jsonOut({ found: false, approved: false });
  }

  return jsonOut({ status: "ok", service: "NEIS Help Admin GAS", time: new Date() });
}

function doPost(e) {
  try {
    var raw = (e && e.postData && e.postData.contents) || "{}";
    var body = JSON.parse(raw);
    var action = body.action || "apply_access";
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getOrCreateTargetSheet(ss);

    // 1. 구글 로그인 시 접속 교사 자동 행 추가
    if (action === "apply_access") {
      var email = String(body.email || "").trim().toLowerCase();
      var name = String(body.name || email.split("@")[0] + " 교사").trim();
      var school = String(body.school || "신규 접속 학교").trim();
      var now = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");

      var isAuto = (email.endsWith("@sch.hs.kr") || email.endsWith(".sch.kr") || email.endsWith("@goedu.kr")) || (body.approved === true) || (body.status === "승인");
      var expDate = new Date();
      expDate.setDate(expDate.getDate() + 180); // 한 학기(6개월)
      var expStr = Utilities.formatDate(expDate, "Asia/Seoul", "yyyy-MM-dd");

      var data = sheet.getDataRange().getValues();
      var foundRow = -1;
      for (var r = 1; r < data.length; r++) {
        for (var c = 0; c < data[r].length; c++) {
          if (String(data[r][c]).trim().toLowerCase() === email) {
            foundRow = r + 1;
            break;
          }
        }
        if (foundRow !== -1) break;
      }

      if (foundRow === -1) {
        sheet.appendRow([
          isAuto, // A열: 승인여부 (체크박스)
          now,    // B열: 신청일시
          name,   // C열: 교사성명
          school, // D열: 소속학교
          email,  // E열: 이메일
          isAuto ? "자동승인" : "신규접속", // F열: 구분
          isAuto ? expStr : "", // G열: 만료예정일
          0       // H열: 연장횟수
        ]);
        var lastR = sheet.getLastRow();
        try { sheet.getRange(lastR, 1).insertCheckboxes(); } catch(err) {}
      } else {
        if (isAuto) {
          sheet.getRange(foundRow, 1).setValue(true);
          sheet.getRange(foundRow, 7).setValue(expStr);
        } else if (body.approved === false || body.status === "대기") {
          sheet.getRange(foundRow, 1).setValue(false);
          sheet.getRange(foundRow, 7).setValue("");
        }
      }
      return jsonOut({ success: true, message: "구글 시트에 접속 기록이 성공적으로 등록되었습니다.", email: email, approved: isAuto });
    }

        // 1-2. 기존 대기/승인 교사 명단 일괄 동기화 (대시보드 기존 명단 전체 전송)
    if (action === "batch_apply_access") {
      var list = body.list || [];
      var added = 0;
      var updated = 0;
      var data = sheet.getDataRange().getValues();
      var emailMap = {};
      
      for (var r = 1; r < data.length; r++) {
        for (var c = 0; c < data[r].length; c++) {
          if (String(data[r][c]).toLowerCase().includes("@")) {
            emailMap[String(data[r][c]).trim().toLowerCase()] = r + 1;
            break;
          }
        }
      }

      var defExp = new Date();
      defExp.setDate(defExp.getDate() + 180);
      var defExpStr = Utilities.formatDate(defExp, "Asia/Seoul", "yyyy-MM-dd");

      list.forEach(function(item) {
        var email = String(item.email || "").trim().toLowerCase();
        if (!email) return;
        var name = String(item.name || email.split("@")[0] + " 교사").trim();
        var school = String(item.school || "신규 등록 학교").trim();
        var isApproved = (item.status === "승인");
        var expStr = item.expiresAt || (isApproved ? defExpStr : "");
        var reqDate = item.date || Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");

        if (emailMap[email]) {
          var rowIdx = emailMap[email];
          if (isApproved) {
            sheet.getRange(rowIdx, 1).setValue(true);
            if (expStr) sheet.getRange(rowIdx, 7).setValue(expStr);
          }
          updated++;
        } else {
          sheet.appendRow([
            isApproved,
            reqDate,
            name,
            school,
            email,
            item.note || (isApproved ? "기존승인" : "대기"),
            expStr,
            0
          ]);
          var lastR = sheet.getLastRow();
          try { sheet.getRange(lastR, 1).insertCheckboxes(); } catch(err) {}
          emailMap[email] = lastR;
          added++;
        }
      });

      return jsonOut({ success: true, message: "총 " + (added + updated) + "명의 기존 교사 명단이 구글 시트에 일괄 동기화되었습니다. (신규 추가: " + added + "건, 상태 갱신: " + updated + "건)", added: added, updated: updated });
    }

    // 2. 만료일 연장
    if (action === "update_expiry") {
      var targetEmail = String(body.email || "").trim().toLowerCase();
      var newDate = String(body.expiresAt || "").trim();
      var sheetData = sheet.getDataRange().getValues();
      for (var r2 = 1; r2 < sheetData.length; r2++) {
        for (var c2 = 0; c2 < sheetData[r2].length; c2++) {
          if (String(sheetData[r2][c2]).trim().toLowerCase() === targetEmail) {
            sheet.getRange(r2 + 1, 1).setValue(true);
            sheet.getRange(r2 + 1, 7).setValue(newDate);
            return jsonOut({ success: true, message: targetEmail + "의 만료일이 " + newDate + "로 갱신되었습니다." });
          }
        }
      }
      return jsonOut({ success: false, error: "해당 이메일을 시트에서 찾을 수 없습니다." });
    }

    return jsonOut({ success: true, message: "수신 완료" });
  } catch (err) {
    return jsonOut({ success: false, error: err.toString() });
  }
}

function getOrCreateTargetSheet(ss) {
  var sheet = ss.getSheetByName("2_교사승인명단") || ss.getSheetByName("사용자승인") || ss.getSheetByName("승인관리");
  if (!sheet) {
    var first = ss.getSheets()[0];
    if (first && first.getLastRow() >= 1) {
      var h = first.getRange(1, 1, 1, Math.min(first.getLastColumn(), 8)).getValues()[0].join(" ");
      if (h.includes("이메일") || h.includes("교사") || h.includes("승인")) {
        return first;
      }
    }
  }
  if (!sheet) {
    sheet = ss.insertSheet("2_교사승인명단");
    sheet.appendRow(["승인여부", "신청일시", "교사성명", "소속학교", "이메일", "구분", "만료예정일", "연장"]);
    sheet.getRange("A1:H1").setBackground("#1e293b").setFontColor("#ffffff").setFontWeight("bold");
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 80);
    sheet.setColumnWidth(2, 160);
    sheet.setColumnWidth(3, 110);
    sheet.setColumnWidth(4, 160);
    sheet.setColumnWidth(5, 220);
    sheet.setColumnWidth(6, 100);
    sheet.setColumnWidth(7, 120);
    sheet.setColumnWidth(8, 70);
  }
  return sheet;
}

function formatDateVal(val) {
  if (!val) return "";
  try {
    return Utilities.formatDate(new Date(val), "Asia/Seoul", "yyyy-MM-dd");
  } catch(e) {
    return String(val).slice(0, 10);
  }
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
