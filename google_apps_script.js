/**
 * ============================================================================
 * [문화놀이터 공조] 영상 및 사진 제작 의뢰서 - 구글 시트 자동 기록 & 이메일 알림 스크립트
 * ============================================================================
 * [사용 방법]
 * 1. 새 구글 스프레드시트(Google Sheets)를 만듭니다.
 * 2. 상단 메뉴에서 [확장 프로그램] -> [Apps Script]를 클릭합니다.
 * 3. 기존에 적혀있는 `function myFunction() {}`을 모두 지우고, 이 파일의 코드를 전체 복사하여 붙여넣습니다.
 * 4. 상단 [💾 프로젝트 저장] 아이콘을 누른 뒤, 우측 상단 파란색 [배포] -> [새 배포]를 클릭합니다.
 * 5. 왼쪽 톱니바퀴(유형 선택)에서 [웹 앱]을 선택합니다.
 *    - 다음 사용자 인증 정보로 실행: [나 (jeju.gongjo@gmail.com)]
 *    - 액세스 권한이 있는 사용자: [모든 사용자] (※ 반드시 '모든 사용자'로 해야 외부 고객이 제출 가능!)
 * 6. [배포] 버튼을 누르고 권한 승인(고급 -> 안전하지 않은 페이지로 이동 -> 허용)을 완료합니다.
 * 7. 생성된 [웹 앱 URL (https://script.google.com/macros/s/.../exec)]을 복사하여
 *    `video_request_form.html` 하단 스크립트의 `GOOGLE_SCRIPT_URL = "여기에 붙여넣기"`에 넣으면 끝납니다!
 * ============================================================================
 */

const NOTIFICATION_EMAIL = "jeju.gongjo@gmail.com";

function doPost(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const data = JSON.parse(e.postData.contents);

    // 기본 헤더 목록 (폼에서 _headersOrder가 넘어오지 않을 때 폴백)
    const defaultHeaders = [
      "제출일시",
      "1_의뢰기관명",
      "1_프로젝트명",
      "1_담당자",
      "1_연락처",
      "1_이메일",
      "2_촬영일시",
      "2_촬영장소",
      "2_예상촬영시간",
      "2_희망납품일",
      "2_예산범위",
      "3_의뢰콘텐츠_수량",
      "3_의뢰콘텐츠_상세내역",
      "4_원본데이터필요여부",
      "5_행사성격",
      "5_타겟시청자",
      "6_제공가능소스",
      "6_첨부파일목록",
      "6_소스공유링크",
      "6_초상권동의여부",
      "6_기타요청사항",
      "원본폼_열람_및_PDF저장_링크"
    ];

    // 폼(index.html)에서 동적으로 전달받은 헤더 목록이 있으면 최우선 반영
    const targetHeaders = Array.isArray(data._headersOrder) && data._headersOrder.length > 0
      ? data._headersOrder
      : defaultHeaders;

    // 첫 번째 행이 비어 있거나, 구버전 헤더인 경우 새 헤더로 자동 교체합니다.
    let sheetHeaders = targetHeaders;
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(targetHeaders);
      sheet.getRange(1, 1, 1, targetHeaders.length)
        .setBackground("#221E53")
        .setFontColor("#FFFFFF")
        .setFontWeight("bold");
      sheet.setFrozenRows(1);
    } else {
      const lastCol = sheet.getLastColumn();
      const existingRow1 = lastCol > 0
        ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String).filter(Boolean)
        : [];
      const isOldHeader =
        existingRow1.indexOf("4_특수촬영요청") !== -1 ||
        existingRow1.indexOf("6_자막내레이션") !== -1 ||
        existingRow1.indexOf("5_영상톤앤매너") !== -1 ||
        existingRow1.indexOf("7_제공가능소스") !== -1;

      if (sheet.getLastRow() === 1 || isOldHeader) {
        sheet.getRange(1, 1, 1, Math.max(sheet.getMaxColumns(), targetHeaders.length)).clearContent();
        sheet.getRange(1, 1, 1, targetHeaders.length)
          .setValues([targetHeaders])
          .setBackground("#221E53")
          .setFontColor("#FFFFFF")
          .setFontWeight("bold");
        sheet.setFrozenRows(1);
        sheetHeaders = targetHeaders;
      } else {
        // 이미 데이터가 쌓여있는 시트라면, 폼에 새로 추가된 열만 맨 뒤에 자동으로 덧붙여줍니다!
        const missingHeaders = targetHeaders.filter((h) => !existingRow1.includes(h));
        if (missingHeaders.length > 0) {
          const startCol = existingRow1.length + 1;
          sheet.getRange(1, startCol, 1, missingHeaders.length)
            .setValues([missingHeaders])
            .setBackground("#221E53")
            .setFontColor("#FFFFFF")
            .setFontWeight("bold");
          sheetHeaders = existingRow1.concat(missingHeaders);
        } else {
          sheetHeaders = existingRow1;
        }
      }
    }

    // 시트에 새 행으로 데이터 추가
    const row = sheetHeaders.map((key) => data[key] || "-");
    sheet.appendRow(row);

    // 7번 섹션에서 첨부한 소스 파일(Base64)을 이메일 첨부파일(Blob)로 변환
    const attachmentBlobs = [];
    if (Array.isArray(data._attachments)) {
      data._attachments.forEach((att) => {
        if (att && att.base64 && att.name) {
          const decodedBytes = Utilities.base64Decode(att.base64);
          const blob = Utilities.newBlob(
            decodedBytes,
            att.mimeType || "application/octet-stream",
            att.name
          );
          attachmentBlobs.push(blob);
        }
      });
    }

    // jeju.gongjo@gmail.com 으로 깔끔한 이메일 알림 + [원본 폼 확인] / [PDF 저장하기(표 문서)] + [첨부 소스 파일] 발송
    const subject = data["_subject"] || `[문화놀이터 공조] 신규 제작 의뢰서 접수 (${data["1_의뢰기관명"]})`;
    const viewLink = data["원본폼_열람_및_PDF저장_링크"] || "";
    const tablePdfLink = viewLink
      ? (viewLink.indexOf("?") !== -1 ? `${viewLink}&mode=table` : `${viewLink}?mode=table`)
      : "";

    let htmlBody = `
      <div style="font-family: 'Pretendard', sans-serif; max-width: 720px; border: 1px solid #E2DFDA; border-radius: 12px; overflow: hidden;">
        <div style="background: #221E53; color: #FFFFFF; padding: 24px 28px;">
          <div style="color: #FAAD34; font-size: 12px; font-weight: bold; margin-bottom: 6px;">NEW PRODUCTION REQUEST</div>
          <h2 style="margin: 0; font-size: 20px;">${subject}</h2>
        </div>
        <div style="padding: 24px 28px; background: #FFFFFF;">
          <div style="background: #F2F0ED; border-left: 5px solid #F05A22; padding: 16px; border-radius: 8px; margin-bottom: 24px;">
            <strong style="color: #221E53; font-size: 15px;">📋 의뢰서 열람 및 표 문서 PDF 저장</strong>
            <p style="margin: 6px 0 14px; font-size: 13px; color: #3A3668; line-height: 1.5;">
              • <strong>[원본 폼 확인]</strong>: 의뢰자가 체크하고 작성한 원본 폼 화면 그대로 확인합니다.<br/>
              • <strong>[PDF 저장하기]</strong>: 바로 아래에 보이는 <strong>표 형태의 요약 문서</strong>를 A4 PDF로 즉시 저장합니다.
            </p>
            <div>
              <a href="${viewLink}" target="_blank" style="display: inline-block; background: #221E53; color: #FFFFFF; text-decoration: none; padding: 10px 18px; border-radius: 8px; font-weight: bold; font-size: 13.5px; margin-right: 8px; margin-bottom: 6px;">
                📋 원본 폼 확인
              </a>
              <a href="${tablePdfLink}" target="_blank" style="display: inline-block; background: #F05A22; color: #FFFFFF; text-decoration: none; padding: 10px 18px; border-radius: 8px; font-weight: bold; font-size: 13.5px; margin-bottom: 6px;">
                🖨️ PDF 저장하기
              </a>
            </div>
          </div>
          <table style="width: 100%; border-collapse: collapse; font-size: 13.5px;">
    `;

    const displayHeaders = Array.isArray(data._headersOrder) && data._headersOrder.length > 0
      ? data._headersOrder
      : sheetHeaders;

    displayHeaders.forEach((key) => {
      if (key === "원본폼_열람_및_PDF저장_링크") return;
      htmlBody += `
        <tr>
          <th style="text-align: left; vertical-align: top; padding: 10px 12px; border-bottom: 1px solid #E2DFDA; background: #F9F8F6; color: #221E53; width: 30%;">${key}</th>
          <td style="padding: 10px 12px; border-bottom: 1px solid #E2DFDA; color: #3A3668; white-space: pre-line;">${data[key] || "-"}</td>
        </tr>
      `;
    });

    htmlBody += `
          </table>
        </div>
      </div>
    `;

    const mailOptions = {
      to: NOTIFICATION_EMAIL,
      subject: subject,
      htmlBody: htmlBody
    };

    if (attachmentBlobs.length > 0) {
      mailOptions.attachments = attachmentBlobs;
    }

    MailApp.sendEmail(mailOptions);

    return ContentService
      .createTextOutput(JSON.stringify({ result: "success" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: "error", message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
