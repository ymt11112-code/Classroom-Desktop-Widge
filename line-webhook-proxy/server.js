const http = require('http');
const https = require('https');
const crypto = require('crypto');
const fs = require('fs');
const path = require('path');
const { URL } = require('url');

loadEnvFile();

const PORT = Number(process.env.PORT || 8787);
const LINE_CHANNEL_SECRET = String(process.env.LINE_CHANNEL_SECRET || '').trim();
const LINE_CHANNEL_ACCESS_TOKEN = String(process.env.LINE_CHANNEL_ACCESS_TOKEN || '').trim();
const LINE_DEFAULT_USER_ID = String(process.env.LINE_DEFAULT_USER_ID || '').trim();
const GAS_BASE_URL = String(process.env.GAS_BASE_URL || '').trim();
const APP_TIMEZONE = String(process.env.APP_TIMEZONE || 'Asia/Taipei').trim();

const LINE_HOMEROOM = ['國語', '數學', '社會', '健康', '樂活'];
const LINE_SUBJECT = ['自然', '藝專', '閩語', '視覺', '英語', '分部課', '樂理'];
const LEAVE_TYPES = ['事假', '病假', '公假', '喪假', '曠課'];
const PENDING_DRAFT_TTL_MS = 30 * 60 * 1000;
const STUDENT_CACHE_TTL_MS = 5 * 60 * 1000;

const pendingAbsenceDrafts = new Map();
const pendingCounselingDrafts = new Map();
const studentCache = { expiresAt: 0, students: [] };

if (!LINE_CHANNEL_SECRET) console.warn('Missing LINE_CHANNEL_SECRET');
if (!LINE_CHANNEL_ACCESS_TOKEN) console.warn('Missing LINE_CHANNEL_ACCESS_TOKEN');
if (!GAS_BASE_URL) console.warn('Missing GAS_BASE_URL');

const server = http.createServer(async (req, res) => {
  try {
    const url = new URL(req.url, `http://${req.headers.host || 'localhost'}`);

    if (req.method === 'GET' && url.pathname === '/') {
      return sendText(res, 200, 'LINE webhook proxy is running');
    }

    if (req.method === 'GET' && url.pathname === '/health') {
      return sendJson(res, 200, { ok: true, service: 'line-webhook-proxy' });
    }

    if (req.method === 'POST' && url.pathname === '/webhook') {
      const rawBody = await readRawBody(req);
      if (!verifyLineSignature(req.headers['x-line-signature'], rawBody, LINE_CHANNEL_SECRET)) {
        return sendJson(res, 401, { ok: false, error: 'Invalid LINE signature' });
      }

      let payload = {};
      try {
        payload = rawBody ? JSON.parse(rawBody) : {};
      } catch (_error) {
        return sendJson(res, 400, { ok: false, error: 'Invalid JSON body' });
      }

      sendText(res, 200, 'OK');
      setImmediate(() => {
        processWebhookPayload(payload).catch((error) => {
          console.error('[processWebhookPayload]', error && error.stack ? error.stack : error);
        });
      });
      return;
    }

    return sendJson(res, 404, { ok: false, error: 'Not found' });
  } catch (error) {
    console.error('[server]', error && error.stack ? error.stack : error);
    if (!res.headersSent) return sendJson(res, 500, { ok: false, error: 'Internal server error' });
  }
});

server.listen(PORT, () => {
  console.log(`LINE webhook proxy listening on http://localhost:${PORT}`);
});

async function processWebhookPayload(payload) {
  const events = Array.isArray(payload && payload.events) ? payload.events : [];
  if (!events.length) return;

  cleanupExpiredDrafts();

  for (const event of events) {
    if (!event || event.type !== 'message' || !event.message || event.message.type !== 'text') {
      continue;
    }

    const text = String(event.message.text || '').trim();
    if (!text) continue;

    const replyToken = event.replyToken ? String(event.replyToken).trim() : '';
    const userId = event.source && event.source.userId ? String(event.source.userId).trim() : LINE_DEFAULT_USER_ID;

    try {
      if (isConfirmAbsenceCommand(text)) {
        await confirmPendingAbsence(replyToken, userId);
        continue;
      }

      if (isCancelAbsenceCommand(text)) {
        await cancelPendingAbsence(replyToken, userId);
        continue;
      }

      if (isConfirmCounselingCommand(text)) {
        await confirmPendingCounseling(replyToken, userId);
        continue;
      }

      if (isCancelCounselingCommand(text)) {
        await cancelPendingCounseling(replyToken, userId);
        continue;
      }

      if (isHelpCommand(text)) {
        await respondToLineEvent(replyToken, userId, buildCommandHelpFlexMessage());
        continue;
      }

      if (isTodoListCommand(text)) {
        const items = await fetchTodoItemsFromGas();
        await respondToLineEvent(replyToken, userId, buildTodoFlexMessage(items));
        continue;
      }

      const addTodoMatch = parseAddTodoCommand(text);
      if (addTodoMatch) {
        const items = await addTodoItemToGas(addTodoMatch.task, addTodoMatch.dueDate);
        await respondToLineEvent(replyToken, userId, buildTodoFlexMessage(items));
        continue;
      }

      const checkTodoNos = parseCheckTodoCommand(text);
      if (checkTodoNos !== null) {
        const items = await checkTodoItemsInGas(checkTodoNos);
        await respondToLineEvent(replyToken, userId, buildTodoFlexMessage(items));
        continue;
      }

      if (isClearCompletedTodoCommand(text)) {
        const items = await clearCompletedTodosFromGas();
        await respondToLineEvent(replyToken, userId, buildTodoFlexMessage(items));
        continue;
      }

      const deleteTodoNos = parseDeleteTodoCommand(text);
      if (deleteTodoNos !== null) {
        const items = await deleteTodoItemsFromGas(deleteTodoNos);
        await respondToLineEvent(replyToken, userId, buildTodoFlexMessage(items));
        continue;
      }

      const dateStr = parseQueryDate(text);
      if (dateStr) {
        const digest = await fetchLineDigest(dateStr);
        await respondToLineEvent(replyToken, userId, buildFlexMessage(digest));
        console.log(`[event ${dateStr}] delivered`);
        continue;
      }

      if (looksLikeCounselingMessage(text)) {
        await prepareCounselingDraft(replyToken, userId, text);
        continue;
      }

      if (looksLikeAbsenceMessage(text)) {
        await prepareAbsenceDraft(replyToken, userId, text);
      }
    } catch (error) {
      console.error('[event]', error && error.stack ? error.stack : error);
      await respondToLineEvent(replyToken, userId, {
        type: 'text',
        text: `處理失敗：${error.message || error}`
      }).catch((deliveryError) => {
        console.error('[delivery fallback]', deliveryError && deliveryError.stack ? deliveryError.stack : deliveryError);
      });
    }
  }
}

function cleanupExpiredDrafts() {
  const now = Date.now();
  cleanupExpiredDraftMap(pendingAbsenceDrafts, now);
  cleanupExpiredDraftMap(pendingCounselingDrafts, now);
}

function cleanupExpiredDraftMap(map, now) {
  for (const [userId, draft] of map.entries()) {
    if (!draft || !draft.createdAt || now - draft.createdAt > PENDING_DRAFT_TTL_MS) {
      map.delete(userId);
    }
  }
}

function parseQueryDate(text) {
  const value = String(text || '').trim();
  const today = getTodayInTimeZone();

  if (value === '今天' || value === '今日') return formatMdDate(today);
  if (value === '明天') return formatMdDate(addCalendarDays(today, 1));
  if (value === '昨天') return formatMdDate(addCalendarDays(today, -1));

  const prevMatch = value.match(/^__prev__(\d{1,2})\/(\d{1,2})/);
  if (prevMatch) {
    return formatMdDate(new Date(today.getFullYear(), Number(prevMatch[1]) - 1, Number(prevMatch[2]) - 1));
  }

  const nextMatch = value.match(/^__next__(\d{1,2})\/(\d{1,2})/);
  if (nextMatch) {
    return formatMdDate(new Date(today.getFullYear(), Number(nextMatch[1]) - 1, Number(nextMatch[2]) + 1));
  }

  const mdMatch = value.match(/^(\d{1,2})[/-](\d{1,2})$/);
  if (mdMatch) return `${Number(mdMatch[1])}/${Number(mdMatch[2])}`;

  return null;
}

function looksLikeAbsenceMessage(text) {
  const value = normalizeText(text);
  if (looksLikeCounselingMessage(value)) return false;
  const keywords = [
    '請假', '病假', '事假', '公假', '喪假', '曠課',
    '遲到', '早退', '退餐', '停餐', '不用餐', '不退餐',
    '看醫生', '發燒', '身體不舒服', '家裡有事', '比賽', '告別式'
  ];
  return keywords.some((keyword) => value.includes(keyword));
}

function looksLikeCounselingMessage(text) {
  const value = normalizeText(text);
  return value.startsWith('輔導：') || value.startsWith('輔導:');
}

function isConfirmAbsenceCommand(text) {
  const value = String(text || '').trim();
  return value === '確認請假' || value === '確認' || value === '送出請假';
}

function isCancelAbsenceCommand(text) {
  const value = String(text || '').trim();
  return value === '取消請假' || value === '取消';
}

function isConfirmCounselingCommand(text) {
  const value = String(text || '').trim();
  return value === '確認輔導' || value === '送出輔導';
}

function isCancelCounselingCommand(text) {
  const value = String(text || '').trim();
  return value === '取消輔導';
}

async function fetchLineDigest(dateStr) {
  const gasUrl = new URL(GAS_BASE_URL);
  gasUrl.searchParams.set('action', 'getLineDigest');
  gasUrl.searchParams.set('date', dateStr);

  const response = await requestJson(gasUrl, { method: 'GET' });
  if (!response.ok) throw new Error(response.error || 'GAS returned an error');
  return response;
}

async function fetchStudentList() {
  const now = Date.now();
  if (studentCache.expiresAt > now && Array.isArray(studentCache.students) && studentCache.students.length) {
    return studentCache.students;
  }

  const gasUrl = new URL(GAS_BASE_URL);
  gasUrl.searchParams.set('action', 'getStudentList');
  const response = await requestJson(gasUrl, { method: 'GET' });
  if (!response.ok) throw new Error(response.error || '無法取得學生名單');

  studentCache.students = Array.isArray(response.students) ? response.students : [];
  studentCache.expiresAt = now + STUDENT_CACHE_TTL_MS;
  return studentCache.students;
}

async function saveAbsenceRecordToGas(record) {
  const gasUrl = new URL(GAS_BASE_URL);
  gasUrl.searchParams.set('action', 'saveAbsenceRecord');
  gasUrl.searchParams.set('data', JSON.stringify(record));

  const response = await requestJson(gasUrl, { method: 'GET' });
  if (!response.ok) throw new Error(response.error || '寫入出缺席失敗');
  return response;
}

async function fetchAiCounselingStrategy(student, reason, related) {
  const gasUrl = new URL(GAS_BASE_URL);
  gasUrl.searchParams.set('action', 'getAiCounselingStrategy');
  gasUrl.searchParams.set('student', student || '');
  gasUrl.searchParams.set('reason', reason || '');
  gasUrl.searchParams.set('related', related || '');

  const response = await requestJson(gasUrl, { method: 'GET' });
  if (!response.ok) throw new Error(response.error || '取得 AI 輔導建議失敗');
  return String(response.strategy || '').trim();
}

async function saveCounselingRecordToGas(record) {
  const gasUrl = new URL(GAS_BASE_URL);
  gasUrl.searchParams.set('action', 'saveCounselingRecord');
  gasUrl.searchParams.set('data', JSON.stringify(record));

  const response = await requestJson(gasUrl, { method: 'GET' });
  if (!response.ok) throw new Error(response.error || '寫入輔導紀錄失敗');
  return response;
}

async function prepareAbsenceDraft(replyToken, userId, rawText) {
  const students = await fetchStudentList();
  const records = buildParsedAbsenceRecords(rawText, students);
  if (!records.length) throw new Error('無法辨識請假對象，請補上座號或姓名。');
  if (!userId) throw new Error('缺少 userId，無法建立請假草稿。');

  pendingAbsenceDrafts.set(userId, {
    createdAt: Date.now(),
    records
  });

  await respondToLineEvent(replyToken, userId, {
    type: 'text',
    text: buildAbsenceDraftSummary(records)
  });
}

async function confirmPendingAbsence(replyToken, userId) {
  if (!userId) throw new Error('缺少 userId，無法確認請假。');

  const pending = pendingAbsenceDrafts.get(userId);
  if (!pending || !Array.isArray(pending.records) || !pending.records.length) {
    await respondToLineEvent(replyToken, userId, {
      type: 'text',
      text: '目前沒有待確認的請假草稿。'
    });
    return;
  }

  for (const record of pending.records) {
    await saveAbsenceRecordToGas(record);
  }

  pendingAbsenceDrafts.delete(userId);
  await respondToLineEvent(replyToken, userId, {
    type: 'text',
    text: `已成功寫入 ${pending.records.length} 筆出缺席紀錄。`
  });
}

async function cancelPendingAbsence(replyToken, userId) {
  if (userId) pendingAbsenceDrafts.delete(userId);
  await respondToLineEvent(replyToken, userId, {
    type: 'text',
    text: '已取消本次請假草稿。'
  });
}

async function prepareCounselingDraft(replyToken, userId, rawText) {
  const students = await fetchStudentList();
  const record = buildParsedCounselingDraft(rawText, students);
  if (!record) throw new Error('無法辨識輔導對象，請補上座號或姓名。');
  if (!userId) throw new Error('缺少 userId，無法建立輔導草稿。');

  record.strategy = await fetchAiCounselingStrategy(record.student, record.reason, record.related);

  pendingCounselingDrafts.set(userId, {
    createdAt: Date.now(),
    record
  });

  await respondToLineEvent(replyToken, userId, {
    type: 'text',
    text: buildCounselingDraftSummary(record)
  });
}

async function confirmPendingCounseling(replyToken, userId) {
  if (!userId) throw new Error('缺少 userId，無法確認輔導紀錄。');

  const pending = pendingCounselingDrafts.get(userId);
  if (!pending || !pending.record) {
    await respondToLineEvent(replyToken, userId, {
      type: 'text',
      text: '目前沒有待確認的輔導紀錄草稿。'
    });
    return;
  }

  await saveCounselingRecordToGas(pending.record);
  pendingCounselingDrafts.delete(userId);
  await respondToLineEvent(replyToken, userId, {
    type: 'text',
    text: '已成功寫入 1 筆輔導紀錄。'
  });
}

async function cancelPendingCounseling(replyToken, userId) {
  if (userId) pendingCounselingDrafts.delete(userId);
  await respondToLineEvent(replyToken, userId, {
    type: 'text',
    text: '已取消本次輔導草稿。'
  });
}

function buildParsedAbsenceRecords(text, students) {
  const normalizedText = normalizeText(text);
  const mode = inferModeFromText(normalizedText);
  const studentLabels = inferStudentsFromText(normalizedText, students);
  if (!studentLabels.length) return [];

  const dates = parseDateTokens(normalizedText);
  const durationDays = inferDurationDays(normalizedText);
  const noticeDate = todayIso();
  const startDate = dates[0] || noticeDate;
  let endDate = dates[1] || startDate;

  if (mode === 'leave' && dates.length <= 1 && durationDays && durationDays > 1) {
    endDate = addWeekdays(startDate, durationDays - 1);
  }

  const inferredPeriods = inferPeriodsFromText(normalizedText);
  const leaveType = mode === 'leave' ? inferLeaveTypeFromText(normalizedText) : (mode === 'late' ? '遲到' : '早退');
  const meal = mode === 'leave' ? (inferMealChoiceFromText(normalizedText) || '否') : '';
  const days = mode === 'leave'
    ? calculateDays(startDate, endDate, inferredPeriods || '7')
    : String(countWeekdaysInclusive(startDate, endDate) || (startDate ? 1 : 0));
  const periods = resolvePeriodsValue(mode, startDate, endDate, inferredPeriods, days);
  const mealRule = evaluateMealRule({
    leaveType,
    meal,
    startDate,
    endDate,
    noticeDate
  });
  const note = String(text || '').trim();

  return studentLabels.map((label) => ({
    student: buildAbsenceStudentDisplay([label]),
    selectedStudents: [label],
    startDate,
    endDate,
    periods,
    days,
    leaveType,
    meal,
    mealReturned: mealRule.eligible,
    mealRuleMessage: mealRule.message,
    note
  }));
}

function buildParsedCounselingDraft(text, students) {
  const body = stripCounselingPrefix(text);
  const normalizedText = normalizeText(body);
  const selectedStudents = inferStudentsFromText(normalizedText, students);
  if (!selectedStudents.length) return null;

  const primaryStudent = selectedStudents[0];
  const relatedStudents = selectedStudents.slice(1);
  const period = inferCounselingPeriod(normalizedText);
  const dateIso = parseDateTokens(normalizedText)[0] || todayIso();
  const reason = buildCounselingReason(body, selectedStudents);
  const related = relatedStudents.join('、');

  return {
    date: dateIso,
    period,
    student: primaryStudent,
    reason,
    related,
    informant: 'LINE',
    contact: '',
    contactContent: '',
    strategy: '',
    selectedStudents
  };
}

function buildCounselingReason(text, selectedStudents) {
  let reason = String(text || '').trim();
  for (const label of selectedStudents) {
    reason = reason.replace(new RegExp(escapeRegExp(label), 'g'), ' ');
    const parts = String(label).split(/\s+/);
    if (parts[1]) {
      reason = reason.replace(new RegExp(escapeRegExp(parts[1]), 'g'), ' ');
    }
    if (parts[0]) {
      reason = reason.replace(new RegExp(`${escapeRegExp(parts[0])}\\s*號?`, 'g'), ' ');
    }
  }
  reason = reason
    .replace(/輔導[:：]?/g, ' ')
    .replace(/今天|今日|明天|昨天|後天/g, ' ')
    .replace(/第?\s*(晨光|[1-7])\s*節/g, ' ')
    .replace(/\d{1,2}[\/.-]\d{1,2}(?:\s*[-~～到至]\s*\d{1,2}[\/.-]\d{1,2})?/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  return reason || String(text || '').trim();
}

function inferCounselingPeriod(text) {
  if (/晨光/.test(text)) return '晨光';
  const numberedMatch = String(text || '').match(/第?\s*([1-7])\s*節/);
  if (numberedMatch) return numberedMatch[1];
  return '';
}

function buildAbsenceDraftSummary(records) {
  const lines = ['請假分析結果：'];
  records.forEach((record, index) => {
    lines.push(
      `${index + 1}. ${record.selectedStudents[0]}`,
      `日期：${record.startDate}${record.endDate && record.endDate !== record.startDate ? ` ~ ${record.endDate}` : ''}`,
      `類型：${record.leaveType}`,
      `節數：${record.periods || '未填'}`,
      `日數：${record.days || '未填'}`,
      `退餐：${record.meal || '未填'}`,
      `退餐判定：${record.mealReturned ? '可退' : '不可退'}`,
      `說明：${record.mealRuleMessage || '無'}`
    );
  });
  lines.push('', '回覆「確認請假」即可寫入出缺席，回覆「取消請假」可放棄。');
  return lines.join('\n');
}

function buildCounselingDraftSummary(record) {
  return [
    '輔導紀錄草稿：',
    `日期：${record.date || '未填'}`,
    `節次：${record.period || '未填'}`,
    `學生：${record.student || '未填'}`,
    `相關人員：${record.related || '無'}`,
    `事由：${record.reason || '未填'}`,
    '',
    'AI 輔導紀錄：',
    record.strategy || '（AI 未產生內容）',
    '',
    '回覆「確認輔導」即可寫入輔導紀錄，回覆「取消輔導」可放棄。'
  ].join('\n');
}

function resolvePeriodsValue(mode, startDate, endDate, inferredPeriods, days) {
  if (mode !== 'leave') return '';
  if (inferredPeriods) return inferredPeriods;
  const schoolDays = countWeekdaysInclusive(startDate, endDate);
  if (schoolDays > 1) return String(schoolDays * 7);
  if (String(days || '') === '0.5') return '4';
  return '7';
}

function evaluateMealRule(input) {
  const meal = String(input.meal || '');
  const leaveType = String(input.leaveType || '');
  const noticeDate = String(input.noticeDate || '');
  const startDate = String(input.startDate || '');
  const endDate = String(input.endDate || startDate);
  const schoolDays = countWeekdaysInclusive(startDate, endDate);
  const leadDays = calculateLeadDays(noticeDate, startDate);

  if (meal !== '是') return { eligible: false, message: '未申請退餐。' };

  if (leaveType === '病假') {
    if (schoolDays >= 3) {
      return { eligible: true, message: `符合病假退餐規則：連續 ${schoolDays} 個上課日以上，但通知當日不退費。` };
    }
    return { eligible: false, message: `不符合病假退餐規則：病假需連續 3 個上課日以上，目前僅 ${schoolDays} 個上課日。` };
  }

  if (leaveType === '事假') {
    if (schoolDays < 5) {
      return { eligible: false, message: `不符合事假退餐規則：事假需連續 5 個上課日以上，目前僅 ${schoolDays} 個上課日。` };
    }
    if (leadDays < 7) {
      return { eligible: false, message: `不符合事假退餐規則：事假需 7 天前通知，目前僅提前 ${leadDays} 天。` };
    }
    return { eligible: true, message: `符合事假退餐規則：連續 ${schoolDays} 個上課日，且提前 ${leadDays} 天通知。` };
  }

  if (leaveType === '公假') {
    if (leadDays < 7) {
      return { eligible: false, message: `不符合公假退餐規則：公假需 7 天前通知，目前僅提前 ${leadDays} 天。` };
    }
    return { eligible: true, message: `符合公假退餐規則：已提前 ${leadDays} 天通知。` };
  }

  return { eligible: false, message: `${leaveType || '此假別'} 不符合目前設定的退餐規則。` };
}

function calculateLeadDays(noticeDate, startDate) {
  const notice = parseIsoDate(noticeDate);
  const start = parseIsoDate(startDate);
  if (!notice || !start) return 0;
  const diffMs = start.getTime() - notice.getTime();
  return Math.max(0, Math.floor(diffMs / (24 * 60 * 60 * 1000)));
}

function todayIso() {
  const date = getTodayInTimeZone();
  return toIsoFromParts(date.getFullYear(), date.getMonth() + 1, date.getDate());
}

function getTodayInTimeZone() {
  const formatter = new Intl.DateTimeFormat('en-CA', {
    timeZone: APP_TIMEZONE,
    year: 'numeric',
    month: '2-digit',
    day: '2-digit'
  });
  const parts = formatter.formatToParts(new Date());
  const map = {};
  for (const part of parts) {
    if (part.type !== 'literal') map[part.type] = part.value;
  }
  return new Date(Number(map.year), Number(map.month) - 1, Number(map.day));
}

function addCalendarDays(dateObj, offsetDays) {
  const next = new Date(dateObj);
  next.setDate(next.getDate() + offsetDays);
  return next;
}

function formatMdDate(dateObj) {
  return `${dateObj.getMonth() + 1}/${dateObj.getDate()}`;
}

function formatTodoDueDateForLine(value) {
  if (value instanceof Date) return formatMdDate(value);

  const text = String(value || '').trim();
  const isoMatch = text.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})$/);
  if (isoMatch) return `${Number(isoMatch[2])}/${Number(isoMatch[3])}`;

  const mdMatch = text.match(/^(\d{1,2})[-\/](\d{1,2})$/);
  if (mdMatch) return `${Number(mdMatch[1])}/${Number(mdMatch[2])}`;

  return text;
}

function toIsoFromParts(year, month, day) {
  return [String(year), String(month).padStart(2, '0'), String(day).padStart(2, '0')].join('-');
}

function parseIsoDate(iso) {
  if (!iso) return null;
  const match = String(iso).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return null;
  const date = new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
  return Number.isNaN(date.getTime()) ? null : date;
}

function normalizeText(text) {
  return String(text || '')
    .replace(/\u3000/g, ' ')
    .replace(/[，、。；：]/g, ' ')
    .replace(/\s*([/.\-~～到至])\s*/g, '$1')
    .replace(/\s+/g, ' ')
    .trim();
}

function stripCounselingPrefix(text) {
  return String(text || '').replace(/^輔導[:：]?\s*/, '').trim();
}

function addWeekdays(startIso, extraDays) {
  const start = parseIsoDate(startIso);
  if (!start) return startIso;
  if (extraDays <= 0) return startIso;

  const cursor = new Date(start);
  let remaining = extraDays;
  while (remaining > 0) {
    cursor.setDate(cursor.getDate() + 1);
    if (cursor.getDay() !== 0 && cursor.getDay() !== 6) remaining -= 1;
  }
  return toIsoFromParts(cursor.getFullYear(), cursor.getMonth() + 1, cursor.getDate());
}

function countWeekdaysInclusive(startIso, endIso) {
  const start = parseIsoDate(startIso);
  const end = parseIsoDate(endIso);
  if (!start || !end || end < start) return 0;

  const cursor = new Date(start);
  let count = 0;
  while (cursor <= end) {
    if (cursor.getDay() !== 0 && cursor.getDay() !== 6) count += 1;
    cursor.setDate(cursor.getDate() + 1);
  }
  return count;
}

function calculateDays(startIso, endIso, periods) {
  if (!startIso) return '';
  if (!endIso || startIso === endIso) {
    const periodCount = parseFloat(periods) || 0;
    if (!periodCount) return '';
    return periodCount < 4 ? '0.5' : '1';
  }
  const weekdays = countWeekdaysInclusive(startIso, endIso);
  return weekdays ? String(weekdays) : '';
}

function parseDateTokens(text) {
  const value = String(text || '');
  const today = getTodayInTimeZone();
  const currentYear = today.getFullYear();
  const results = [];

  if (value.includes('今天') || value.includes('今日')) results.push(todayIso());
  if (value.includes('明天')) {
    const next = addCalendarDays(today, 1);
    results.push(toIsoFromParts(next.getFullYear(), next.getMonth() + 1, next.getDate()));
  }
  if (value.includes('後天')) {
    const next = addCalendarDays(today, 2);
    results.push(toIsoFromParts(next.getFullYear(), next.getMonth() + 1, next.getDate()));
  }
  if (value.includes('昨天')) {
    const prev = addCalendarDays(today, -1);
    results.push(toIsoFromParts(prev.getFullYear(), prev.getMonth() + 1, prev.getDate()));
  }

  for (const match of value.matchAll(/(\d{4})[\/.-](\d{1,2})[\/.-](\d{1,2})/g)) {
    results.push(toIsoFromParts(Number(match[1]), Number(match[2]), Number(match[3])));
  }

  for (const match of value.matchAll(/(^|[^\d])(\d{1,2})[\/.-](\d{1,2})(?!\d)/g)) {
    results.push(toIsoFromParts(currentYear, Number(match[2]), Number(match[3])));
  }

  for (const match of value.matchAll(/(下週|下礼拜|下禮拜|這週|這周|本週|本周|週|星期|禮拜)([一二三四五六日天])/g)) {
    const weekdayIndex = weekdayTextToIndex(match[2]);
    if (weekdayIndex === -1) continue;
    const useNextWeek = /^下/.test(match[1]);
    const resolved = findUpcomingWeekday(today, weekdayIndex, useNextWeek);
    results.push(toIsoFromParts(resolved.getFullYear(), resolved.getMonth() + 1, resolved.getDate()));
  }

  return Array.from(new Set(results)).filter(Boolean);
}

function stripDatePhrases(text) {
  return String(text || '')
    .replace(/今天|今日|明天|昨天|後天/g, ' ')
    .replace(/(下週|下礼拜|下禮拜|這週|這周|本週|本周|週|星期|禮拜)[一二三四五六日天]/g, ' ')
    .replace(/\d{4}[\/.-]\d{1,2}[\/.-]\d{1,2}\s*[-~～到至]\s*\d{4}[\/.-]\d{1,2}[\/.-]\d{1,2}/g, ' ')
    .replace(/\d{1,2}[\/.-]\d{1,2}\s*[-~～到至]\s*\d{1,2}[\/.-]\d{1,2}/g, ' ')
    .replace(/\d{4}[\/.-]\d{1,2}[\/.-]\d{1,2}/g, ' ')
    .replace(/\d{1,2}[\/.-]\d{1,2}/g, ' ');
}

function inferModeFromText(text) {
  if (String(text).includes('早退')) return 'early';
  if (String(text).includes('遲到')) return 'late';
  return 'leave';
}

function inferLeaveTypeFromText(text) {
  const explicit = LEAVE_TYPES.find((type) => String(text).includes(type));
  if (explicit) return explicit;
  if (/(身體不舒服|發燒|看醫生|不適|腸胃炎|感冒)/.test(text)) return '病假';
  if (/(家裡有事|臨時有事|私事|事假)/.test(text)) return '事假';
  if (/(比賽|活動|公假|校隊|代表隊)/.test(text)) return '公假';
  if (/(告別式|喪禮|治喪|奔喪)/.test(text)) return '喪假';
  return '事假';
}

function inferMealChoiceFromText(text) {
  if (/(不退餐|午餐照常|不用退餐)/.test(text)) return '否';
  if (/(退餐|停餐|不用餐|午餐退費|營養午餐.*退)/.test(text)) return '是';
  return null;
}

function inferPeriodsFromText(text) {
  const normalized = String(text || '')
    .replace(/節課/g, '節')
    .replace(/到/g, '-')
    .replace(/[～~至]/g, '-');

  const rangeMatch = normalized.match(/(\d+(?:\.\d+)?)\s*-\s*(\d+(?:\.\d+)?)\s*節/);
  if (rangeMatch) {
    const start = parseFloat(rangeMatch[1]);
    const end = parseFloat(rangeMatch[2]);
    if (!Number.isNaN(start) && !Number.isNaN(end) && end >= start) {
      return String(end - start + 1);
    }
  }

  const singleMatch = normalized.match(/(\d+(?:\.\d+)?)\s*節/);
  if (singleMatch) return String(singleMatch[1]);

  if (/(整天|一天|全天)/.test(text)) return '7';
  if (/(半天|上午|早上|中午接回)/.test(text)) return '4';
  if (/(下午)/.test(text)) return '3';
  return '';
}

function inferDurationDays(text) {
  const numericMatch = String(text || '').match(/(\d+)\s*天/);
  if (numericMatch) return Number(numericMatch[1]);

  const chineseDayMap = { 一: 1, 二: 2, 兩: 2, 三: 3, 四: 4, 五: 5, 六: 6, 七: 7 };
  const chineseMatch = String(text || '').match(/([一二兩三四五六七])\s*天/);
  if (chineseMatch) return chineseDayMap[chineseMatch[1]] || null;
  if (/一天/.test(text)) return 1;
  return null;
}

function inferStudentsFromText(text, students) {
  if (String(text).includes('全班')) return ['全班'];

  const matches = [];
  const textWithoutDates = stripDatePhrases(text);

  for (const match of textWithoutDates.matchAll(/(^|[^\d])(\d{1,2})\s*號(?!\d)/g)) {
    const student = findStudentBySeatNumber(match[2], students);
    if (student) matches.push(buildStudentLabel(student));
  }

  for (const student of students) {
    const label = buildStudentLabel(student);
    const studentName = String(student.name || '').trim();
    if (text.includes(label) || (studentName && text.includes(studentName))) {
      matches.push(label);
    }
  }

  return Array.from(new Set(matches));
}

function findStudentBySeatNumber(seatNumber, students) {
  const normalizedSeat = String(parseInt(seatNumber, 10));
  return students.find((student) => String(parseInt(student.num, 10)) === normalizedSeat) || null;
}

function buildStudentLabel(student) {
  return `${student.num} ${student.name}`;
}

function extractSeatNumber(label) {
  const match = String(label || '').match(/^\s*(\d{1,2})/);
  return match ? match[1] : '';
}

function buildSeatDisplay(selectedStudents) {
  const seatNumbers = Array.from(new Set((selectedStudents || [])
    .map((label) => extractSeatNumber(label))
    .filter(Boolean)));
  return seatNumbers.join('、');
}

function buildAbsenceStudentDisplay(selectedStudents) {
  const selected = Array.from(new Set((selectedStudents || []).filter(Boolean)));
  if (!selected.length) return '';
  if (selected[0] === '全班') return '全班';
  return buildSeatDisplay(selected);
}

function weekdayTextToIndex(text) {
  const map = { 一: 1, 二: 2, 三: 3, 四: 4, 五: 5, 六: 6, 日: 0, 天: 0 };
  return Object.prototype.hasOwnProperty.call(map, text) ? map[text] : -1;
}

function findUpcomingWeekday(baseDate, targetWeekday, forceNextWeek) {
  const base = new Date(baseDate);
  const currentWeekday = base.getDay();
  let delta = targetWeekday - currentWeekday;
  if (delta < 0) delta += 7;
  if (forceNextWeek || delta === 0) delta += 7;
  base.setDate(base.getDate() + delta);
  return base;
}

function isTodoListCommand(text) {
  const value = String(text || '').trim();
  return value === '待辦' || value === '待辦事項';
}

function parseTodoNumberList(str) {
  const nums = String(str || '').split(/[,，\s]+/)
    .map((s) => parseInt(s.trim(), 10))
    .filter((n) => !isNaN(n) && n > 0);
  return nums.length ? nums : null;
}

function parseCheckTodoCommand(text) {
  const value = String(text || '').trim();

  // "完成待辦 1,3,4" / "勾選待辦 1 3 4"
  let m = value.match(/^(?:完成待辦|勾選待辦|勾待辦)\s+([\d,，\s]+)$/);
  if (m) return parseTodoNumberList(m[1]);

  // "完成 1,3,4" / "完成 134"（有空格：按逗號/空格分隔）
  m = value.match(/^(?:完成|勾選)\s+([\d,，\s]+)$/);
  if (m) return parseTodoNumberList(m[1]);

  // "完成1,3,4"（無空格有逗號）或 "完成134"（無空格無逗號：逐位數）
  m = value.match(/^(?:完成|勾選)([\d,，]+)$/);
  if (m) {
    const part = m[1];
    if (part.includes(',') || part.includes('，')) return parseTodoNumberList(part);
    const digits = part.split('').map(Number).filter((n) => n > 0);
    return digits.length ? digits : null;
  }

  return null;
}

function isClearCompletedTodoCommand(text) {
  const value = String(text || '').trim();
  return value === '清除已完成' || value === '清空已完成' || value === '清除完成';
}

function isHelpCommand(text) {
  const value = String(text || '').trim();
  return value === '指令說明' || value === 'help' || value === '？' || value === '?' || value === '登記';
}

function parseDeleteTodoCommand(text) {
  const value = String(text || '').trim();

  // "刪除待辦 2,3" / "移除待辦 2 3"
  let m = value.match(/^(?:刪除待辦|移除待辦)\s+([\d,，\s]+)$/);
  if (m) return parseTodoNumberList(m[1]);

  // "刪除 2,3" / "刪除2,3" / "刪除23"（無逗號=單一編號）
  m = value.match(/^(?:刪除|移除)\s*([\d,，\s]+)$/);
  if (m && m[1].trim()) return parseTodoNumberList(m[1]);

  return null;
}

async function checkTodoItemsInGas(itemNos) {
  const gasUrl = new URL(GAS_BASE_URL);
  gasUrl.searchParams.set('action', 'checkTodoItem');
  gasUrl.searchParams.set('itemNos', itemNos.join(','));
  const response = await requestJson(gasUrl, { method: 'GET' });
  if (!response.ok) throw new Error(response.error || '勾選待辦事項失敗');
  return Array.isArray(response.items) ? response.items : [];
}

async function clearCompletedTodosFromGas() {
  const gasUrl = new URL(GAS_BASE_URL);
  gasUrl.searchParams.set('action', 'clearCompletedTodos');
  const response = await requestJson(gasUrl, { method: 'GET' });
  if (!response.ok) throw new Error(response.error || '清除已完成待辦失敗');
  return Array.isArray(response.items) ? response.items : [];
}

async function deleteTodoItemsFromGas(itemNos) {
  const gasUrl = new URL(GAS_BASE_URL);
  gasUrl.searchParams.set('action', 'deleteTodoItem');
  gasUrl.searchParams.set('itemNos', itemNos.join(','));
  const response = await requestJson(gasUrl, { method: 'GET' });
  if (!response.ok) throw new Error(response.error || '刪除待辦事項失敗');
  return Array.isArray(response.items) ? response.items : [];
}

function parseAddTodoCommand(text) {
  const value = String(text || '').trim();
  const match = value.match(/^(?:\+待辦|新增待辦)\s+(.+)$/);
  if (!match) return null;
  return extractTrailingDate(match[1].trim());
}

function extractTrailingDate(text) {
  const CN_MONTH = '(?:十[一二]?|[一二三四五六七八九])';
  const CN_DAY = '(?:三十一?|二十[一二三四五六七八九]?|十[一二三四五六七八九]?|[一二三四五六七八九十])';
  const DATE_PAT =
    '(?:\\d{4}[-\\/]\\d{1,2}[-\\/]\\d{1,2}' +
    '|\\d{1,2}[\\/\\-]\\d{1,2}' +
    '|\\d{1,2}月(?:' + CN_DAY + '|\\d{1,2})日?' +
    '|' + CN_MONTH + '月(?:' + CN_DAY + '|\\d{1,2})日?)';
  const KW = '(?:截止日期|日期|截止)';
  const SEP = '[，,、\\s]';
  const suffixRe = new RegExp(
    '(?:' + SEP + '+' + KW + '?' + SEP + '*|' + KW + SEP + '*)(' + DATE_PAT + ')$'
  );

  const m = text.match(suffixRe);
  if (m) {
    const parsed = parseDateString(m[1]);
    if (parsed) {
      const task = text.slice(0, m.index).replace(/[，,、\s]+$/, '').trim();
      if (task) return { task, dueDate: parsed };
    }
  }
  return { task: text, dueDate: '' };
}

function parseDateString(s) {
  const str = String(s || '').trim().replace(/日$/, '');
  if (!str) return null;

  const isoM = str.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})$/);
  if (isoM) return `${Number(isoM[2])}/${Number(isoM[3])}`;

  const mdM = str.match(/^(\d{1,2})[\/\-](\d{1,2})$/);
  if (mdM) return `${Number(mdM[1])}/${Number(mdM[2])}`;

  const aMonthM = str.match(/^(\d{1,2})月(.+)$/);
  if (aMonthM) {
    const d = chineseOrArabicToNum(aMonthM[2]);
    if (d) return `${Number(aMonthM[1])}/${d}`;
  }

  const cMonthM = str.match(/^(十[一二]?|[一二三四五六七八九十])月(.+)$/);
  if (cMonthM) {
    const mo = chineseMonthNum(cMonthM[1]);
    const d = chineseOrArabicToNum(cMonthM[2]);
    if (mo && d) return `${mo}/${d}`;
  }

  return null;
}

function chineseOrArabicToNum(s) {
  const str = String(s || '').trim().replace(/日$/, '');
  if (!str) return null;
  if (/^\d+$/.test(str)) return parseInt(str, 10);
  const singles = { '一': 1, '二': 2, '三': 3, '四': 4, '五': 5, '六': 6, '七': 7, '八': 8, '九': 9 };
  if (str === '十') return 10;
  const t1 = str.match(/^十([一二三四五六七八九]?)$/);
  if (t1) return 10 + (singles[t1[1]] || 0);
  const t2 = str.match(/^二十([一二三四五六七八九]?)$/);
  if (t2) return 20 + (singles[t2[1]] || 0);
  const t3 = str.match(/^三十([一]?)$/);
  if (t3) return 30 + (t3[1] ? 1 : 0);
  return singles[str] || null;
}

function chineseMonthNum(s) {
  const str = String(s || '').trim();
  if (str === '十一') return 11;
  if (str === '十二') return 12;
  const map = { '一': 1, '二': 2, '三': 3, '四': 4, '五': 5, '六': 6, '七': 7, '八': 8, '九': 9, '十': 10 };
  return map[str] || null;
}

async function fetchTodoItemsFromGas() {
  const gasUrl = new URL(GAS_BASE_URL);
  gasUrl.searchParams.set('action', 'getTodoItems');
  const response = await requestJson(gasUrl, { method: 'GET' });
  if (!response.ok) throw new Error(response.error || '無法取得待辦事項');
  return Array.isArray(response.items) ? response.items : [];
}

async function addTodoItemToGas(task, dueDate) {
  const gasUrl = new URL(GAS_BASE_URL);
  gasUrl.searchParams.set('action', 'addTodoItem');
  gasUrl.searchParams.set('task', task);
  if (dueDate) gasUrl.searchParams.set('dueDate', dueDate);
  const response = await requestJson(gasUrl, { method: 'GET' });
  if (!response.ok) throw new Error(response.error || '新增待辦事項失敗');
  return Array.isArray(response.items) ? response.items : [];
}

function buildTodoFlexMessage(items) {
  const active = (items || []).filter((item) => item.task);
  const rows = [];

  if (!active.length) {
    rows.push({
      type: 'box', layout: 'vertical', paddingAll: '12px',
      contents: [{ type: 'text', text: '（目前沒有待辦事項）', size: 'lg', color: '#aaaaaa', align: 'center' }]
    });
  } else {
    for (const item of active) {
      const mark = item.checked ? '✅' : '⬜';
      const textColor = item.checked ? '#aaaaaa' : '#243b63';
      const bg = item.checked ? '#f5f5f5' : '#ffffff';
      const label = (item.itemNo ? `${item.itemNo}. ` : '') + item.task;
      const rowContents = [
        { type: 'text', text: `${mark} ${label}`, size: 'lg', flex: 5, wrap: true, color: textColor }
      ];
      if (item.dueDate) {
        rowContents.push({ type: 'text', text: formatTodoDueDateForLine(item.dueDate), size: 'sm', flex: 2, align: 'end', color: '#888888', gravity: 'center' });
      }
      rows.push({ type: 'box', layout: 'horizontal', paddingAll: '8px', backgroundColor: bg, contents: rowContents });
      rows.push({ type: 'separator', color: '#e0e8e0' });
    }
  }

  return {
    type: 'flex',
    altText: '📝 待辦事項清單',
    quickReply: {
      items: [
        { type: 'action', action: { type: 'message', label: '待辦清單', text: '待辦' } },
        { type: 'action', action: { type: 'message', label: '清除已完成', text: '清除已完成' } },
        { type: 'action', action: { type: 'message', label: '今天進度', text: '今天' } }
      ]
    },
    contents: {
      type: 'bubble',
      size: 'mega',
      header: {
        type: 'box', layout: 'horizontal',
        backgroundColor: '#e6f4e6', paddingAll: '12px',
        contents: [{ type: 'text', text: '📝 待辦事項', weight: 'bold', size: 'xl', color: '#1a5c1a' }]
      },
      body: {
        type: 'box', layout: 'vertical',
        spacing: 'none',
        paddingTop: '0px', paddingBottom: '0px',
        paddingStart: 'lg', paddingEnd: 'lg',
        contents: rows
      },
      footer: {
        type: 'box', layout: 'vertical', paddingAll: '10px',
        backgroundColor: '#f0f7f0',
        contents: [{ type: 'text', text: '新增：+待辦 事項 5/20　勾選：完成待辦 1　刪除：刪除待辦 1', size: 'sm', color: '#558855', align: 'center', wrap: true }]
      }
    }
  };
}

function buildCommandHelpFlexMessage() {
  return {
    type: 'flex',
    altText: '📋 LINE 指令說明',
    contents: {
      type: 'carousel',
      contents: [
        {
          type: 'bubble', size: 'mega',
          header: { type: 'box', layout: 'vertical', backgroundColor: '#e6f4e6', paddingAll: '12px', contents: [{ type: 'text', text: '📝 待辦事項 指令說明', weight: 'bold', size: 'lg', color: '#1a5c1a' }] },
          body: {
            type: 'box', layout: 'vertical', spacing: 'none', paddingAll: '0px',
            contents: [
              { type: 'box', layout: 'horizontal', backgroundColor: '#c8e4c8', paddingAll: '6px', contents: [{ type: 'text', text: '指令', size: 'xs', weight: 'bold', flex: 5, color: '#1a5c1a' }, { type: 'text', text: '效果', size: 'xs', weight: 'bold', flex: 7, color: '#1a5c1a' }] },
              { type: 'box', layout: 'horizontal', paddingAll: '7px', backgroundColor: '#f9fff9', contents: [{ type: 'text', text: '待辦', size: 'sm', weight: 'bold', flex: 5, color: '#2d7a2d', wrap: true }, { type: 'text', text: '顯示待辦清單', size: 'sm', flex: 7, wrap: true, color: '#444444' }] },
              { type: 'separator', color: '#c8e4c8' },
              { type: 'box', layout: 'horizontal', paddingAll: '7px', contents: [{ type: 'text', text: '+待辦 訂便當', size: 'sm', weight: 'bold', flex: 5, color: '#2d7a2d', wrap: true }, { type: 'text', text: '新增（無截止日）', size: 'sm', flex: 7, wrap: true, color: '#444444' }] },
              { type: 'separator', color: '#c8e4c8' },
              { type: 'box', layout: 'horizontal', paddingAll: '7px', backgroundColor: '#f9fff9', contents: [{ type: 'text', text: '+待辦 訂便當 5/20', size: 'sm', weight: 'bold', flex: 5, color: '#2d7a2d', wrap: true }, { type: 'text', text: '新增＋截止日\n（支援中文日期）', size: 'sm', flex: 7, wrap: true, color: '#444444' }] },
              { type: 'separator', color: '#c8e4c8' },
              { type: 'box', layout: 'horizontal', paddingAll: '7px', contents: [{ type: 'text', text: '完成1', size: 'sm', weight: 'bold', flex: 5, color: '#2d7a2d', wrap: true }, { type: 'text', text: '切換第1項勾選狀態', size: 'sm', flex: 7, wrap: true, color: '#444444' }] },
              { type: 'separator', color: '#c8e4c8' },
              { type: 'box', layout: 'horizontal', paddingAll: '7px', backgroundColor: '#f9fff9', contents: [{ type: 'text', text: '完成1,3,4\n完成134', size: 'sm', weight: 'bold', flex: 5, color: '#2d7a2d', wrap: true }, { type: 'text', text: '批次切換勾選\n（逗號或逐位數）', size: 'sm', flex: 7, wrap: true, color: '#444444' }] },
              { type: 'separator', color: '#c8e4c8' },
              { type: 'box', layout: 'horizontal', paddingAll: '7px', contents: [{ type: 'text', text: '刪除待辦 2,3\n刪除2,3', size: 'sm', weight: 'bold', flex: 5, color: '#2d7a2d', wrap: true }, { type: 'text', text: '刪除並重新編號', size: 'sm', flex: 7, wrap: true, color: '#444444' }] },
              { type: 'separator', color: '#c8e4c8' },
              { type: 'box', layout: 'horizontal', paddingAll: '7px', backgroundColor: '#f9fff9', contents: [{ type: 'text', text: '清除已完成', size: 'sm', weight: 'bold', flex: 5, color: '#2d7a2d', wrap: true }, { type: 'text', text: '移除全部已完成項目', size: 'sm', flex: 7, wrap: true, color: '#444444' }] }
            ]
          },
          footer: { type: 'box', layout: 'vertical', backgroundColor: '#f0f7f0', paddingAll: '8px', contents: [{ type: 'text', text: '「新增待辦」可代替「+待辦」；截止日支援 5/20、五月二十日、截止日期… 等格式', size: 'xs', color: '#558855', wrap: true }] }
        },
        {
          type: 'bubble', size: 'mega',
          header: { type: 'box', layout: 'vertical', backgroundColor: '#e3eeff', paddingAll: '12px', contents: [{ type: 'text', text: '🏫 出缺席記錄 指令說明', weight: 'bold', size: 'lg', color: '#1a3a6a' }] },
          body: {
            type: 'box', layout: 'vertical', spacing: 'none', paddingAll: '0px',
            contents: [
              { type: 'box', layout: 'horizontal', backgroundColor: '#c0d4f0', paddingAll: '6px', contents: [{ type: 'text', text: '指令', size: 'xs', weight: 'bold', flex: 5, color: '#1a3a6a' }, { type: 'text', text: '效果', size: 'xs', weight: 'bold', flex: 7, color: '#1a3a6a' }] },
              { type: 'box', layout: 'horizontal', paddingAll: '7px', backgroundColor: '#f7faff', contents: [{ type: 'text', text: '7號宋珏 病假', size: 'sm', weight: 'bold', flex: 5, color: '#1a4a8a', wrap: true }, { type: 'text', text: '分析假別，產生請假草稿', size: 'sm', flex: 7, wrap: true, color: '#444444' }] },
              { type: 'separator', color: '#c0d4e8' },
              { type: 'box', layout: 'horizontal', paddingAll: '7px', contents: [{ type: 'text', text: '全班明天事假', size: 'sm', weight: 'bold', flex: 5, color: '#1a4a8a', wrap: true }, { type: 'text', text: '全班＋相對日期', size: 'sm', flex: 7, wrap: true, color: '#444444' }] },
              { type: 'separator', color: '#c0d4e8' },
              { type: 'box', layout: 'horizontal', paddingAll: '7px', backgroundColor: '#f7faff', contents: [{ type: 'text', text: '3號 遲到', size: 'sm', weight: 'bold', flex: 5, color: '#1a4a8a', wrap: true }, { type: 'text', text: '遲到 / 早退記錄', size: 'sm', flex: 7, wrap: true, color: '#444444' }] },
              { type: 'separator', color: '#c0d4e8' },
              { type: 'box', layout: 'horizontal', paddingAll: '7px', contents: [{ type: 'text', text: '7號 看醫生，退餐', size: 'sm', weight: 'bold', flex: 5, color: '#1a4a8a', wrap: true }, { type: 'text', text: '推斷假別＋退餐申請', size: 'sm', flex: 7, wrap: true, color: '#444444' }] },
              { type: 'separator', color: '#c0d4e8' },
              { type: 'box', layout: 'horizontal', paddingAll: '7px', backgroundColor: '#f7faff', contents: [{ type: 'text', text: '確認 / 確認請假\n送出請假', size: 'sm', weight: 'bold', flex: 5, color: '#1a4a8a', wrap: true }, { type: 'text', text: '草稿確認後寫入出缺席', size: 'sm', flex: 7, wrap: true, color: '#444444' }] },
              { type: 'separator', color: '#c0d4e8' },
              { type: 'box', layout: 'horizontal', paddingAll: '7px', contents: [{ type: 'text', text: '取消 / 取消請假', size: 'sm', weight: 'bold', flex: 5, color: '#1a4a8a', wrap: true }, { type: 'text', text: '放棄草稿，不寫入', size: 'sm', flex: 7, wrap: true, color: '#444444' }] }
            ]
          },
          footer: { type: 'box', layout: 'vertical', backgroundColor: '#eef2ff', paddingAll: '8px', contents: [{ type: 'text', text: '退餐規則：病假連續3日｜事假5日+7天前通知｜公假7天前通知', size: 'xs', color: '#3355aa', wrap: true }, { type: 'text', text: '觸發關鍵字：病假 事假 公假 喪假 曠課 遲到 早退 退餐 看醫生 發燒 家裡有事 比賽…', size: 'xs', color: '#778899', wrap: true, margin: 'sm' }] }
        },
        {
          type: 'bubble', size: 'mega',
          header: { type: 'box', layout: 'vertical', backgroundColor: '#f5eeff', paddingAll: '12px', contents: [{ type: 'text', text: '🧭 輔導記錄 指令說明', weight: 'bold', size: 'lg', color: '#4a1a7a' }] },
          body: {
            type: 'box', layout: 'vertical', spacing: 'none', paddingAll: '0px',
            contents: [
              { type: 'box', layout: 'horizontal', backgroundColor: '#d4c0ec', paddingAll: '6px', contents: [{ type: 'text', text: '指令', size: 'xs', weight: 'bold', flex: 5, color: '#4a1a7a' }, { type: 'text', text: '效果', size: 'xs', weight: 'bold', flex: 7, color: '#4a1a7a' }] },
              { type: 'box', layout: 'horizontal', paddingAll: '7px', backgroundColor: '#faf7ff', contents: [{ type: 'text', text: '輔導：7號宋珏 上課說話', size: 'sm', weight: 'bold', flex: 5, color: '#5a2a8a', wrap: true }, { type: 'text', text: '建立草稿＋AI 產生輔導建議', size: 'sm', flex: 7, wrap: true, color: '#444444' }] },
              { type: 'separator', color: '#d0c0e8' },
              { type: 'box', layout: 'horizontal', paddingAll: '7px', contents: [{ type: 'text', text: '輔導：5號 與8號衝突', size: 'sm', weight: 'bold', flex: 5, color: '#5a2a8a', wrap: true }, { type: 'text', text: '第二位學生自動設為相關人員', size: 'sm', flex: 7, wrap: true, color: '#444444' }] },
              { type: 'separator', color: '#d0c0e8' },
              { type: 'box', layout: 'horizontal', paddingAll: '7px', backgroundColor: '#faf7ff', contents: [{ type: 'text', text: '輔導：3號 第2節 情緒失控', size: 'sm', weight: 'bold', flex: 5, color: '#5a2a8a', wrap: true }, { type: 'text', text: '支援節次（第X節 / 晨光）', size: 'sm', flex: 7, wrap: true, color: '#444444' }] },
              { type: 'separator', color: '#d0c0e8' },
              { type: 'box', layout: 'horizontal', paddingAll: '7px', contents: [{ type: 'text', text: '確認輔導\n送出輔導', size: 'sm', weight: 'bold', flex: 5, color: '#5a2a8a', wrap: true }, { type: 'text', text: '草稿確認後寫入輔導紀錄', size: 'sm', flex: 7, wrap: true, color: '#444444' }] },
              { type: 'separator', color: '#d0c0e8' },
              { type: 'box', layout: 'horizontal', paddingAll: '7px', backgroundColor: '#faf7ff', contents: [{ type: 'text', text: '取消輔導', size: 'sm', weight: 'bold', flex: 5, color: '#5a2a8a', wrap: true }, { type: 'text', text: '放棄草稿，不寫入', size: 'sm', flex: 7, wrap: true, color: '#444444' }] }
            ]
          },
          footer: { type: 'box', layout: 'vertical', backgroundColor: '#f7f0ff', paddingAll: '8px', contents: [{ type: 'text', text: '必須以「輔導：」開頭才會觸發｜草稿有效期 30 分鐘', size: 'xs', color: '#6633aa', wrap: true }] }
        }
      ]
    }
  };
}

function buildFlexMessage(digest) {
  const schedule = digest.schedule || {};
  const periods = Array.isArray(schedule.periods) ? schedule.periods : [];
  const weeklyNotes = Array.isArray(digest.weeklyNotes) ? digest.weeklyNotes : [];
  const contactText = String(digest.contact || '').trim();
  const headerText = `📅 ${digest.dayLabel || digest.date || ''}${digest.todayKey ? `　${digest.todayKey}` : ''}`;

  const tableRows = [{
    type: 'box',
    layout: 'horizontal',
    backgroundColor: '#dce8fb',
    paddingAll: '8px',
    contents: [
      { type: 'text', text: '節', size: 'lg', flex: 1, align: 'center', weight: 'bold', color: '#2a4070' },
      { type: 'text', text: '進度', size: 'lg', flex: 5, weight: 'bold', color: '#2a4070' },
      { type: 'text', text: '備註', size: 'lg', flex: 4, weight: 'bold', color: '#2a4070' }
    ]
  }];

  if (periods.length) {
    for (const period of periods) {
      tableRows.push({
        type: 'box',
        layout: 'horizontal',
        backgroundColor: getRowColor(period),
        paddingAll: '8px',
        contents: [
          {
            type: 'text',
            text: period.period === '晨光' ? '晨' : String(period.period || ''),
            size: 'lg',
            flex: 1,
            align: 'center',
            color: '#5c6f92',
            weight: 'bold'
          },
          { type: 'text', text: String(period.content || '－'), size: 'lg', flex: 5, wrap: true, color: '#243b63' },
          { type: 'text', text: String(period.note || '　'), size: 'lg', flex: 4, wrap: true, color: '#50627f' }
        ]
      });
      tableRows.push({ type: 'separator', color: '#c8d5e8' });
    }
  } else {
    tableRows.push({
      type: 'box',
      layout: 'vertical',
      paddingAll: '8px',
      contents: [{ type: 'text', text: '（今日無課程進度）', size: 'lg', color: '#9aaabb', align: 'center' }]
    });
  }

  const footerItems = [
    { type: 'text', text: '📖 聯絡簿', weight: 'bold', size: 'xl', color: '#243b63' }
  ];

  const contactLines = contactText.split('\n').map((line) => line.trim()).filter(Boolean);
  if (contactLines.length) {
    for (const line of contactLines) {
      footerItems.push({ type: 'text', text: line, size: 'lg', color: '#50627f', wrap: true, margin: 'xs' });
    }
  } else {
    footerItems.push({ type: 'text', text: '（今日無聯絡簿內容）', size: 'lg', color: '#aaaaaa' });
  }

  if (weeklyNotes.length) {
    footerItems.push({ type: 'separator', color: '#c8d5e8', margin: 'md' });
    footerItems.push({ type: 'text', text: '📋 本週記事', weight: 'bold', size: 'xl', color: '#243b63', margin: 'md' });
    for (const note of weeklyNotes) {
      if (note) {
        footerItems.push({ type: 'text', text: `• ${note}`, size: 'lg', color: '#50627f', wrap: true, margin: 'xs' });
      }
    }
  }

  return {
    type: 'flex',
    altText: `${digest.dayLabel || digest.date || ''} 今日進度`,
    quickReply: {
      items: [
        { type: 'action', action: { type: 'message', label: '⬅ 前一天', text: `__prev__${digest.dayLabel || digest.date || ''}` } },
        { type: 'action', action: { type: 'message', label: '今天', text: '今天' } },
        { type: 'action', action: { type: 'message', label: '明天 ➡', text: `__next__${digest.dayLabel || digest.date || ''}` } }
      ]
    },
    contents: {
      type: 'bubble',
      size: 'mega',
      header: {
        type: 'box',
        layout: 'horizontal',
        backgroundColor: '#e3eeff',
        paddingAll: '12px',
        contents: [{ type: 'text', text: headerText, weight: 'bold', size: 'xl', color: '#1a3a6a', wrap: true }]
      },
      body: {
        type: 'box',
        layout: 'vertical',
        spacing: 'none',
        paddingTop: '0px',
        paddingBottom: '0px',
        paddingStart: 'lg',
        paddingEnd: 'lg',
        contents: tableRows
      },
      footer: {
        type: 'box',
        layout: 'vertical',
        spacing: 'sm',
        paddingTop: '10px',
        paddingBottom: '10px',
        paddingStart: 'lg',
        paddingEnd: 'lg',
        backgroundColor: '#f7faff',
        contents: footerItems
      }
    }
  };
}

function getRowColor(period) {
  const subject = String((period && period.subject) || '');
  const content = String((period && period.content) || '');
  const note = String((period && period.note) || '');
  const text = `${subject} ${content} ${note}`;

  if (/調課|借課|還課/.test(text)) return '#fff4a8';
  if (/(^|[^A-Za-z])X([^A-Za-z]|$)/i.test(content)) return '#d8dce4';
  if (LINE_HOMEROOM.some((keyword) => subject.includes(keyword) || content.includes(keyword))) return '#f5f0eb';
  if (LINE_SUBJECT.some((keyword) => subject.includes(keyword) || content.includes(keyword))) return '#dde6dc';
  return '#ffffff';
}

async function respondToLineEvent(replyToken, userId, message) {
  if (replyToken) return replyLineMessage(replyToken, message);
  if (userId) return pushLineMessage(userId, message);
  throw new Error('No replyToken or userId available for delivery');
}

async function replyLineMessage(replyToken, message) {
  const body = JSON.stringify({ replyToken, messages: [message] });
  return requestJson('https://api.line.me/v2/bot/message/reply', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      Authorization: `Bearer ${LINE_CHANNEL_ACCESS_TOKEN}`
    },
    body
  });
}

async function pushLineMessage(userId, message) {
  const body = JSON.stringify({ to: userId, messages: [message] });
  return requestJson('https://api.line.me/v2/bot/message/push', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      Authorization: `Bearer ${LINE_CHANNEL_ACCESS_TOKEN}`
    },
    body
  });
}

function verifyLineSignature(signature, rawBody, secret) {
  if (!signature || !secret) return false;
  const digest = crypto.createHmac('sha256', secret).update(rawBody).digest('base64');
  const actual = Buffer.from(String(signature));
  const expected = Buffer.from(digest);
  return actual.length === expected.length && crypto.timingSafeEqual(actual, expected);
}

function readRawBody(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    req.on('data', (chunk) => chunks.push(chunk));
    req.on('end', () => resolve(Buffer.concat(chunks).toString('utf8')));
    req.on('error', reject);
  });
}

function requestJson(targetUrl, options, redirectCount) {
  return new Promise((resolve, reject) => {
    const currentRedirectCount = Number(redirectCount || 0);
    const url = typeof targetUrl === 'string' ? new URL(targetUrl) : targetUrl;
    const requestOptions = Object.assign({}, options || {});
    const client = url.protocol === 'https:' ? https : http;

    const request = client.request(url, requestOptions, (response) => {
      const chunks = [];
      response.on('data', (chunk) => chunks.push(chunk));
      response.on('end', () => {
        const raw = Buffer.concat(chunks).toString('utf8');
        const statusCode = response.statusCode || 500;

        if ([301, 302, 303, 307, 308].includes(statusCode) && response.headers.location) {
          if (currentRedirectCount >= 5) {
            return reject(new Error(`HTTP ${statusCode}: too many redirects`));
          }
          const nextUrl = new URL(response.headers.location, url);
          const nextOptions = Object.assign({}, requestOptions);
          if (statusCode === 303) {
            nextOptions.method = 'GET';
            delete nextOptions.body;
            if (nextOptions.headers) {
              delete nextOptions.headers['Content-Length'];
              delete nextOptions.headers['content-length'];
              delete nextOptions.headers['Content-Type'];
              delete nextOptions.headers['content-type'];
            }
          }
          return resolve(requestJson(nextUrl, nextOptions, currentRedirectCount + 1));
        }

        if (!raw) {
          if (statusCode >= 200 && statusCode < 300) return resolve({ ok: true });
          return reject(new Error(`HTTP ${statusCode}`));
        }

        let parsed;
        try {
          parsed = JSON.parse(raw);
        } catch (_error) {
          if (statusCode >= 200 && statusCode < 300) return resolve({ ok: true, raw });
          const compactRaw = raw.replace(/\s+/g, ' ').replace(/<[^>]+>/g, ' ').trim().slice(0, 180);
          return reject(new Error(`HTTP ${statusCode}: ${compactRaw || 'Non-JSON response'}`));
        }

        if (statusCode >= 200 && statusCode < 300) return resolve(parsed);
        return reject(new Error(`HTTP ${statusCode}: ${parsed.error || raw}`));
      });
    });

    request.on('error', reject);
    if (requestOptions.body) request.write(requestOptions.body);
    request.end();
  });
}

function sendJson(res, statusCode, payload) {
  res.writeHead(statusCode, { 'Content-Type': 'application/json; charset=utf-8' });
  res.end(JSON.stringify(payload));
}

function sendText(res, statusCode, text) {
  res.writeHead(statusCode, { 'Content-Type': 'text/plain; charset=utf-8' });
  res.end(text);
}

function loadEnvFile() {
  const envPath = path.join(__dirname, '.env');
  if (!fs.existsSync(envPath)) return;

  const lines = fs.readFileSync(envPath, 'utf8').split(/\r?\n/);
  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith('#')) continue;
    const equalIndex = trimmed.indexOf('=');
    if (equalIndex === -1) continue;
    const key = trimmed.slice(0, equalIndex).trim();
    const value = trimmed.slice(equalIndex + 1).trim();
    if (!process.env[key]) process.env[key] = value;
  }
}

function escapeRegExp(text) {
  return String(text || '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
