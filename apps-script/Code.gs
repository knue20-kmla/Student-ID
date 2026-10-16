// ============================================================
//  민족사관고등학교 학생증 발급 신청 시스템 - Google Apps Script (보안 강화판)
//
//  [필수 설정] 관리자 비밀번호는 코드에 쓰지 않습니다.
//   Apps Script 편집기 > 프로젝트 설정(⚙) > 스크립트 속성 > 속성 추가
//     속성: ADMIN_PASSWORD   값: (새 비밀번호)
//   설정 후 반드시 '배포 > 배포 관리 > 수정 > 새 버전'으로 다시 배포하세요.
// ============================================================

const CONFIG = {
  SHEET_NAME:   '학생증신청',             // Google Sheets 시트 이름
  SCHOOL_NAME:  '민족사관고등학교',
  SCHOOL_EMAIL: 'office@minjok.hs.kr',   // 학교 대표 이메일 (reply-to 표시용)
  ADMIN_EMAIL:  'knue20@minjok.hs.kr',   // 신청 알림 받을 관리자 이메일
  // 학생 이메일을 이 도메인으로 제한하려면 'minjok.hs.kr' 처럼 입력 (빈 문자열이면 제한 없음)
  ALLOWED_EMAIL_DOMAIN: '',
  MAX_FAILED_LOGINS: 10,                 // 로그인 실패 허용 횟수
  LOCKOUT_SECONDS:   600,                // 초과 시 잠금 시간(초)
};

// ── 시트 헤더 컬럼 정의
const HEADERS = [
  '접수번호', '이름', '학번', '학년', '반', '사유',
  '이메일', '신청일시', '상태', '완료일시',
];

const REASONS = ['분실재발급', '훼손재발급'];

// ============================================================
//  GET: 인증이 필요한 데이터는 제공하지 않음
// ============================================================
function doGet(e) {
  return ContentService.createTextOutput('OK');
}

// ============================================================
//  POST: 신청 제출(공개) / 목록 조회·완료 처리(관리자 비밀번호 필요)
// ============================================================
function doPost(e) {
  let payload;
  try {
    payload = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonResponse({ success: false, error: '잘못된 요청 형식' });
  }

  switch (payload.action) {
    case 'submit':
      return handleSubmit(payload);
    case 'list':
    case 'complete': {
      const authError = checkAdmin(payload.password);
      if (authError) return jsonResponse({ success: false, error: authError });
      return payload.action === 'list' ? handleList() : handleComplete(payload);
    }
    default:
      return jsonResponse({ success: false, error: '알 수 없는 action' });
  }
}

// ── 관리자 인증 (실패 횟수 제한 포함). 통과하면 null, 실패하면 오류 메시지 반환
function checkAdmin(password) {
  const cache = CacheService.getScriptCache();
  const failed = Number(cache.get('failedLogins') || 0);
  if (failed >= CONFIG.MAX_FAILED_LOGINS) {
    return '시도 횟수를 초과했습니다. 잠시 후 다시 시도해 주세요.';
  }

  const expected = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  if (!expected) return '서버에 관리자 비밀번호가 설정되지 않았습니다.';

  if (!safeEqual(String(password || ''), expected)) {
    cache.put('failedLogins', String(failed + 1), CONFIG.LOCKOUT_SECONDS);
    return '비밀번호가 올바르지 않습니다.';
  }

  cache.remove('failedLogins');
  return null;
}

// 길이/내용에 관계없이 전체를 비교하는 문자열 비교 (타이밍 공격 완화)
function safeEqual(a, b) {
  let diff = a.length ^ b.length;
  const len = Math.max(a.length, b.length);
  for (let i = 0; i < len; i++) {
    diff |= (a.charCodeAt(i) || 0) ^ (b.charCodeAt(i) || 0);
  }
  return diff === 0;
}

function handleList() {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();

  if (data.length <= 1) return jsonResponse({ success: true, data: [] });

  const rows = data.slice(1).map((row, i) => ({
    rowIndex:  i + 2,
    rowNum:    row[0],
    name:      row[1],
    studentId: row[2],
    grade:     row[3],
    classNum:  row[4],
    reason:    row[5],
    email:     row[6],
    timestamp: row[7],
    status:    row[8],
    doneAt:    row[9],
  }));

  return jsonResponse({ success: true, data: rows });
}

// ── 신청 저장 (공개 엔드포인트: 입력 검증 필수)
function handleSubmit(data) {
  const lock = LockService.getScriptLock();
  try {
    const clean = validateSubmission(data);
    if (clean.error) return jsonResponse({ success: false, error: clean.error });

    lock.waitLock(10000);
    const sheet = getSheet();
    const lastRow = sheet.getLastRow();
    const receiptNum = 'ID-' + String(lastRow).padStart(4, '0');

    sheet.appendRow([
      receiptNum,
      cellSafe(clean.name),
      cellSafe(clean.studentId),
      clean.grade,
      clean.classNum,
      clean.reason,
      cellSafe(clean.email),
      new Date().toLocaleString('ko-KR', { timeZone: 'Asia/Seoul' }),  // 서버 시각 사용
      '대기중',
      '',
    ]);
    lock.releaseLock();

    // 관리자에게 신청 알림 발송
    sendAdminNotification(clean.name, clean.studentId, clean.grade, clean.classNum, clean.reason, clean.email, receiptNum);

    return jsonResponse({ success: true, receiptNum });
  } catch (err) {
    return jsonResponse({ success: false, error: err.message });
  } finally {
    try { lock.releaseLock(); } catch (e) { /* 이미 해제됨 */ }
  }
}

function validateSubmission(d) {
  const name      = String(d.name || '').trim();
  const studentId = String(d.studentId || '').trim();
  const grade     = Number(d.grade);
  const classNum  = Number(d.classNum);
  const reason    = String(d.reason || '').trim();
  const email     = String(d.email || '').trim();

  if (!name || name.length > 30 || /[<>]/.test(name)) return { error: '이름이 올바르지 않습니다.' };
  if (!/^[0-9]{5}$/.test(studentId))                  return { error: '학번이 올바르지 않습니다.' };
  if (![1, 2, 3].includes(grade))                     return { error: '학년이 올바르지 않습니다.' };
  if (!Number.isInteger(classNum) || classNum < 1 || classNum > 10) return { error: '반이 올바르지 않습니다.' };
  if (!REASONS.includes(reason))                      return { error: '신청 사유가 올바르지 않습니다.' };
  if (email.length > 100 || !/^[^\s@<>,;]+@[^\s@<>,;]+\.[^\s@<>,;]+$/.test(email)) {
    return { error: '이메일 형식이 올바르지 않습니다.' };
  }
  if (CONFIG.ALLOWED_EMAIL_DOMAIN &&
      !email.toLowerCase().endsWith('@' + CONFIG.ALLOWED_EMAIL_DOMAIN.toLowerCase())) {
    return { error: '학교 이메일만 사용할 수 있습니다.' };
  }

  return { name, studentId, grade, classNum, reason, email };
}

// ── 완료 처리 + 이메일 발송
function handleComplete(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet    = getSheet();
    const rowIndex = Number(data.rowIndex);

    if (!Number.isInteger(rowIndex) || rowIndex < 2 || rowIndex > sheet.getLastRow()) {
      return jsonResponse({ success: false, error: '잘못된 행 번호입니다.' });
    }

    const row = sheet.getRange(rowIndex, 1, 1, HEADERS.length).getValues()[0];

    const name      = row[1];
    const studentId = row[2];
    const grade     = row[3];
    const classNum  = row[4];
    const email     = row[6];
    const status    = row[8];

    if (status === '완료') {
      return jsonResponse({ success: false, error: '이미 완료 처리된 신청입니다.' });
    }

    // 상태 업데이트
    const doneAt = new Date().toLocaleString('ko-KR', { timeZone: 'Asia/Seoul' });
    sheet.getRange(rowIndex, 9).setValue('완료');
    sheet.getRange(rowIndex, 10).setValue(doneAt);
    lock.releaseLock();

    // 이메일 발송
    sendEmail(email, name, grade, classNum, studentId);

    return jsonResponse({ success: true, name, email });

  } catch (err) {
    return jsonResponse({ success: false, error: err.message });
  } finally {
    try { lock.releaseLock(); } catch (e) { /* 이미 해제됨 */ }
  }
}

// ============================================================
//  관리자 알림 이메일 (신청 접수 시)
// ============================================================
function sendAdminNotification(rawName, rawStudentId, grade, classNum, rawReason, rawEmail, receiptNum) {
  const name = esc(rawName), studentId = esc(rawStudentId), reason = esc(rawReason), email = esc(rawEmail);
  const subject = `[학생증 재발급 신청] ${grade}학년 ${classNum}반 ${oneLine(rawName)} 학생`;

  const htmlBody = `
<!DOCTYPE html>
<html lang="ko">
<head><meta charset="UTF-8"></head>
<body style="margin:0;padding:0;background:#f4f0e8;font-family:'Malgun Gothic',sans-serif;">
<table width="100%" cellpadding="0" cellspacing="0" style="background:#f4f0e8;padding:40px 0;">
  <tr><td align="center">
    <table width="560" cellpadding="0" cellspacing="0"
      style="background:#ffffff;border-radius:16px;overflow:hidden;box-shadow:0 4px 24px rgba(0,0,0,0.08);">

      <!-- 헤더 -->
      <tr>
        <td style="background:#0d1b3e;padding:28px 40px;text-align:center;">
          <div style="color:rgba(255,255,255,0.55);font-size:11px;letter-spacing:3px;margin-bottom:6px;">${CONFIG.SCHOOL_NAME}</div>
          <div style="color:#e8c96a;font-size:18px;font-weight:700;letter-spacing:0.5px;">
            📋 학생증 재발급 신청 접수
          </div>
        </td>
      </tr>

      <!-- 본문 -->
      <tr>
        <td style="padding:32px 40px 28px;">
          <p style="font-size:14px;color:#1a1a2e;line-height:1.8;margin:0 0 24px;">
            새로운 학생증 재발급 신청이 접수되었습니다.
          </p>

          <!-- 신청 정보 카드 -->
          <table width="100%" cellpadding="0" cellspacing="0"
            style="background:#f8f5ee;border-radius:12px;margin-bottom:24px;">
            <tr><td style="padding:24px 28px;">
              <table width="100%" cellpadding="0" cellspacing="0">
                <tr>
                  <td style="padding:8px 0;font-size:12px;color:#9ca3af;width:90px;">접수번호</td>
                  <td style="padding:8px 0;font-size:14px;color:#1a1a2e;font-weight:700;">${receiptNum}</td>
                </tr>
                <tr>
                  <td style="padding:8px 0;font-size:12px;color:#9ca3af;border-top:1px solid #ede8de;">이름</td>
                  <td style="padding:8px 0;font-size:14px;color:#1a1a2e;font-weight:600;border-top:1px solid #ede8de;">${name}</td>
                </tr>
                <tr>
                  <td style="padding:8px 0;font-size:12px;color:#9ca3af;border-top:1px solid #ede8de;">학번</td>
                  <td style="padding:8px 0;font-size:14px;color:#1a1a2e;font-weight:600;border-top:1px solid #ede8de;">${studentId}</td>
                </tr>
                <tr>
                  <td style="padding:8px 0;font-size:12px;color:#9ca3af;border-top:1px solid #ede8de;">학년 / 반</td>
                  <td style="padding:8px 0;font-size:14px;color:#1a1a2e;font-weight:600;border-top:1px solid #ede8de;">${grade}학년 ${classNum}반</td>
                </tr>
                <tr>
                  <td style="padding:8px 0;font-size:12px;color:#9ca3af;border-top:1px solid #ede8de;">신청 사유</td>
                  <td style="padding:8px 0;font-size:14px;color:#1a1a2e;font-weight:600;border-top:1px solid #ede8de;">${reason}</td>
                </tr>
                <tr>
                  <td style="padding:8px 0;font-size:12px;color:#9ca3af;border-top:1px solid #ede8de;">학생 이메일</td>
                  <td style="padding:8px 0;font-size:14px;color:#1a1a2e;font-weight:600;border-top:1px solid #ede8de;">${email}</td>
                </tr>
              </table>
            </td></tr>
          </table>

          <p style="font-size:12px;color:#c0b89a;line-height:1.8;margin:0;">
            본 메일은 학생증 재발급 신청 시스템에서 자동 발송된 메일입니다.
          </p>
        </td>
      </tr>

      <!-- 푸터 -->
      <tr>
        <td style="background:#f8f5ee;padding:18px 40px;text-align:center;border-top:1px solid #e8e0d0;">
          <div style="font-size:11px;color:#b0a898;">${CONFIG.SCHOOL_NAME} · 학생증 재발급 시스템</div>
        </td>
      </tr>

    </table>
  </td></tr>
</table>
</body>
</html>`;

  const plainBody =
    `[학생증 재발급 신청 접수]\n\n` +
    `접수번호: ${receiptNum}\n` +
    `이름:     ${rawName}\n` +
    `학번:     ${rawStudentId}\n` +
    `학년/반:  ${grade}학년 ${classNum}반\n` +
    `신청사유: ${rawReason}\n` +
    `이메일:   ${rawEmail}\n\n` +
    `── ${CONFIG.SCHOOL_NAME} 학생증 재발급 시스템`;

  GmailApp.sendEmail(CONFIG.ADMIN_EMAIL, subject, plainBody, {
    htmlBody: htmlBody,
    name: CONFIG.SCHOOL_NAME,
  });

  Logger.log(`관리자 알림 발송 → ${CONFIG.ADMIN_EMAIL}`);
}

// ============================================================
//  이메일 발송 (GmailApp — 별도 API Key 불필요)
// ============================================================
function sendEmail(email, rawName, grade, classNum, rawStudentId) {
  const name = esc(rawName), studentId = esc(rawStudentId);
  const subject = `[${CONFIG.SCHOOL_NAME}] 학생증 제작 완료 안내`;

  const htmlBody = `
<!DOCTYPE html>
<html lang="ko">
<head><meta charset="UTF-8"></head>
<body style="margin:0;padding:0;background:#f4f0e8;font-family:'Malgun Gothic',sans-serif;">
<table width="100%" cellpadding="0" cellspacing="0" style="background:#f4f0e8;padding:40px 0;">
  <tr><td align="center">
    <table width="560" cellpadding="0" cellspacing="0"
      style="background:#ffffff;border-radius:16px;overflow:hidden;box-shadow:0 4px 24px rgba(0,0,0,0.08);">

      <!-- 헤더 -->
      <tr>
        <td style="background:#0d1b3e;padding:36px 40px;text-align:center;">
          <div style="width:56px;height:56px;background:linear-gradient(135deg,#c9a84c,#e8c96a);
            border-radius:50%;display:inline-block;line-height:56px;
            font-size:22px;font-weight:700;color:#0d1b3e;margin-bottom:14px;">民</div>
          <div style="color:rgba(255,255,255,0.55);font-size:11px;letter-spacing:3px;
            text-transform:uppercase;margin-bottom:6px;">${CONFIG.SCHOOL_NAME}</div>
          <div style="color:#ffffff;font-size:20px;font-weight:700;letter-spacing:0.5px;">
            학생증 제작 완료 안내
          </div>
        </td>
      </tr>

      <!-- 본문 -->
      <tr>
        <td style="padding:36px 40px 28px;">
          <p style="font-size:15px;color:#1a1a2e;line-height:1.9;margin:0 0 28px;">
            안녕하세요, <strong>${name}</strong> 학생.<br>
            신청하신 학생증 제작이 완료되었습니다.
          </p>

          <!-- 정보 카드 -->
          <table width="100%" cellpadding="0" cellspacing="0"
            style="background:#f8f5ee;border-radius:12px;margin-bottom:24px;">
            <tr><td style="padding:24px 28px;">
              <table width="100%" cellpadding="0" cellspacing="0">
                <tr>
                  <td style="padding:8px 0;font-size:12px;color:#9ca3af;
                    text-transform:uppercase;letter-spacing:1px;width:90px;">이름</td>
                  <td style="padding:8px 0;font-size:14px;color:#1a1a2e;font-weight:600;">
                    ${name}</td>
                </tr>
                <tr>
                  <td style="padding:8px 0;font-size:12px;color:#9ca3af;
                    text-transform:uppercase;letter-spacing:1px;">학번</td>
                  <td style="padding:8px 0;font-size:14px;color:#1a1a2e;font-weight:600;">
                    ${studentId}</td>
                </tr>
                <tr>
                  <td style="padding:8px 0;font-size:12px;color:#9ca3af;
                    text-transform:uppercase;letter-spacing:1px;">학년 / 반</td>
                  <td style="padding:8px 0;font-size:14px;color:#1a1a2e;font-weight:600;">
                    ${grade}학년 ${classNum}반</td>
                </tr>
              </table>
            </td></tr>
          </table>

          <!-- 수령 안내 -->
          <table width="100%" cellpadding="0" cellspacing="0"
            style="border-left:4px solid #c9a84c;background:#fffbf0;
              border-radius:0 8px 8px 0;margin-bottom:32px;">
            <tr><td style="padding:16px 20px;font-size:13px;color:#7c5c00;line-height:1.9;">
              📍 <strong>수령 안내</strong><br>
              김태완 오피스(충무관 126호) 방문하여 학생증을 수령해 주세요.
            </td></tr>
          </table>

          <p style="font-size:12px;color:#c0b89a;line-height:1.8;margin:0;">
            본 메일은 학생증 발급 시스템에서 자동 발송된 메일입니다.<br>
            문의사항은 교무실로 연락해 주세요.
          </p>
        </td>
      </tr>

      <!-- 푸터 -->
      <tr>
        <td style="background:#f8f5ee;padding:18px 40px;text-align:center;
          border-top:1px solid #e8e0d0;">
          <div style="font-size:11px;color:#b0a898;letter-spacing:0.5px;">
            ${CONFIG.SCHOOL_NAME} · 학생증 발급 시스템
          </div>
        </td>
      </tr>

    </table>
  </td></tr>
</table>
</body>
</html>`;

  const plainBody =
    `[${CONFIG.SCHOOL_NAME}] 학생증 제작 완료 안내\n\n` +
    `${rawName} 학생, 신청하신 학생증 제작이 완료되었습니다.\n\n` +
    `이름:    ${rawName}\n` +
    `학번:    ${rawStudentId}\n` +
    `학년/반: ${grade}학년 ${classNum}반\n\n` +
    `김태완 오피스(충무관 126호) 방문하여 학생증을 수령해 주세요.\n` +
    `수령 시 본인 확인이 필요합니다.\n\n` +
    `── ${CONFIG.SCHOOL_NAME} 학생증 발급 시스템`;

  GmailApp.sendEmail(email, subject, plainBody, {
    htmlBody: htmlBody,
    name: CONFIG.SCHOOL_NAME,
    replyTo: CONFIG.SCHOOL_EMAIL,
  });

  Logger.log('완료 안내 이메일 발송 완료');
}

// ============================================================
//  유틸
// ============================================================
function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
    sheet.appendRow(HEADERS);

    const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
    headerRange.setBackground('#0d1b3e');
    headerRange.setFontColor('#e8c96a');
    headerRange.setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  return sheet;
}

// HTML 이스케이프 (메일 본문 삽입용)
function esc(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

// 메일 제목용: 줄바꿈 제거
function oneLine(value) {
  return String(value == null ? '' : value).replace(/[\r\n]+/g, ' ');
}

// 스프레드시트 수식 주입 방지
function cellSafe(value) {
  const text = String(value == null ? '' : value);
  return /^[=+\-@\t\r]/.test(text) ? "'" + text : text;
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
