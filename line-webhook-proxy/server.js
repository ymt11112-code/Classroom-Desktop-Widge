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
const GEMINI_API_KEY = String(process.env.GEMINI_API_KEY || '').trim();
const GEMINI_MODEL = String(process.env.GEMINI_MODEL || 'gemini-2.5-flash').trim();
const REVIEW_BASE_URL = String(process.env.REVIEW_BASE_URL || process.env.RENDER_EXTERNAL_URL || '').trim().replace(/\/+$/, '');

const LINE_HOMEROOM = ['國語', '數學', '社會', '健康', '樂活'];
const LINE_SUBJECT = ['自然', '藝專', '閩語', '視覺', '英語', '分部課', '樂理'];
const LEAVE_TYPES = ['事假', '病假', '公假', '喪假', '曠課'];
const PENDING_DRAFT_TTL_MS = 30 * 60 * 1000;
const STUDENT_CACHE_TTL_MS = 5 * 60 * 1000;
const TODO_REVIEW_TTL_MS = 7 * 24 * 60 * 60 * 1000;
const TODO_BATCH_DEBOUNCE_MS = Number(process.env.TODO_BATCH_DEBOUNCE_MS || 20000);
const MAX_INLINE_PDF_BYTES = 18 * 1024 * 1024;

const pendingAbsenceDrafts = new Map();
const pendingCounselingDrafts = new Map();
const pendingTodoReviews = new Map();
const pendingTodoSourceBatches = new Map();
const recentTodoWindowTexts = new Map();
const studentCache = { expiresAt: 0, students: [] };

if (!LINE_CHANNEL_SECRET) console.warn('Missing LINE_CHANNEL_SECRET');
if (!LINE_CHANNEL_ACCESS_TOKEN) console.warn('Missing LINE_CHANNEL_ACCESS_TOKEN');
if (!GAS_BASE_URL) console.warn('Missing GAS_BASE_URL');
if (!GEMINI_API_KEY) console.warn('Missing GEMINI_API_KEY');

const server = http.createServer(async (req, res) => {
  try {
    const url = new URL(req.url, `http://${req.headers.host || 'localhost'}`);

    if (req.method === 'GET' && url.pathname === '/') {
      return sendText(res, 200, 'LINE webhook proxy is running');
    }

    if (req.method === 'GET' && url.pathname === '/health') {
      return sendJson(res, 200, { ok: true, service: 'line-webhook-proxy' });
    }

    if (req.method === 'GET' && url.pathname === '/todo-review') {
      return serveTodoReviewPage(req, res, url);
    }

    if (req.method === 'POST' && url.pathname === '/api/todo-review/commit') {
      return commitTodoReview(req, res, url);
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
    if (!event || event.type !== 'message' || !event.message) {
      continue;
    }

    const replyToken = event.replyToken ? String(event.replyToken).trim() : '';
    const userId = event.source && event.source.userId ? String(event.source.userId).trim() : LINE_DEFAULT_USER_ID;

    try {
      if (event.message.type === 'file') {
        await queueTodoReviewFile(replyToken, userId, event.message);
        continue;
      }

      if (event.message.type !== 'text') continue;

      const text = String(event.message.text || '').trim();
      if (!text) continue;

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

      const todoAiText = parseTodoCandidateTextCommand(text);
      if (todoAiText) {
        await queueTodoReviewText(replyToken, userId, todoAiText);
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

      if (hasPendingTodoReviewBatch(userId)) {
        await queueTodoReviewText(replyToken, userId, text);
        continue;
      }

      rememberRecentTodoWindowText(userId, text);
      if (looksLikeTodoWindowCandidateText(text)) {
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
  cleanupExpiredDraftMap(pendingAbsenceDrafts, now, PENDING_DRAFT_TTL_MS);
  cleanupExpiredDraftMap(pendingCounselingDrafts, now, PENDING_DRAFT_TTL_MS);
  cleanupExpiredDraftMap(pendingTodoReviews, now, TODO_REVIEW_TTL_MS);
  cleanupRecentTodoWindowTexts(now);
}

function cleanupExpiredDraftMap(map, now, ttlMs) {
  for (const [userId, draft] of map.entries()) {
    if (!draft || !draft.createdAt || now - draft.createdAt > ttlMs) {
      map.delete(userId);
    }
  }
}

function cleanupRecentTodoWindowTexts(now) {
  for (const [userId, items] of recentTodoWindowTexts.entries()) {
    const recent = (items || []).filter((item) => item && now - item.createdAt <= TODO_BATCH_DEBOUNCE_MS);
    if (recent.length) {
      recentTodoWindowTexts.set(userId, recent);
    } else {
      recentTodoWindowTexts.delete(userId);
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

function parseTodoCandidateTextCommand(text) {
  const value = String(text || '').trim();
  const match = value.match(/^(?:整理待辦|AI待辦|ai待辦|待辦整理|待辦候選)\s*[:：]?\s*([\s\S]+)$/);
  if (match) {
    const body = String(match[1] || '').trim();
    return body || null;
  }
  if (looksLikeLineReportTodoOpening(value)) return value;
  if (looksLikeTodoBroadcastOpening(value)) return value;
  return null;
}

function looksLikeTodoBroadcastOpening(text) {
  const firstLine = String(text || '').trim().split(/\r?\n/)[0].trim();
  if (!firstLine) return false;
  const compact = firstLine.replace(/\s+/g, '');
  return /^@all(?:\b|[:：,，\s]|$)/i.test(firstLine) || /^@all/i.test(compact) || /^\u8acb\u5b78\u4e3b\u8f49\u50b3/.test(compact);
}

function hasPendingTodoReviewBatch(userId) {
  const batch = pendingTodoSourceBatches.get(userId);
  return !!(batch && Array.isArray(batch.sources) && batch.sources.length);
}

function rememberRecentTodoWindowText(userId, text) {
  if (!userId) return;
  const value = String(text || '').trim();
  if (!value) return;
  const now = Date.now();
  const existing = recentTodoWindowTexts.get(userId) || [];
  existing.push({ createdAt: now, sourceType: 'LINE', sourceLabel: 'LINE 訊息', sourceText: value });
  recentTodoWindowTexts.set(userId, existing.filter((item) => item && now - item.createdAt <= TODO_BATCH_DEBOUNCE_MS));
}

function takeRecentTodoWindowTexts(userId) {
  if (!userId) return [];
  const now = Date.now();
  const existing = recentTodoWindowTexts.get(userId) || [];
  const recent = existing.filter((item) => item && now - item.createdAt <= TODO_BATCH_DEBOUNCE_MS);
  recentTodoWindowTexts.delete(userId);
  return recent.map((item) => ({
    sourceType: item.sourceType,
    sourceLabel: item.sourceLabel,
    sourceText: item.sourceText
  }));
}

function looksLikeTodoWindowCandidateText(text) {
  const value = String(text || '').trim();
  if (!value) return false;
  if (!/https?:\/\//i.test(value)) return false;
  return /活動|比賽|競賽|報名|徵件|徵選|作品|檢附|參閱|辦法|截止|期限|轉知|公告|晨會|報告/.test(value);
}

function looksLikeLineReportTodoOpening(text) {
  const firstLine = String(text || '').trim().split(/\r?\n/)[0].trim();
  if (!firstLine) return false;

  const compact = firstLine.replace(/\s+/g, '');
  if (/^【[^】]{1,20}】報告(?:\(#?\d+\))?\d{6,8}/.test(compact)) return true;

  const rocDate = String.raw`\d{3}(?:[./-]?\d{1,2}[./-]?\d{1,2})`;
  const weekday = String.raw`(?:[（(][一二三四五六日天][）)])?`;
  const reportUnitOrMarker = String.raw`(?:報告|晨會|轉知|輔導組|研發組|學務處|訓育組|國際組|教務處|總務處|人事室|會計室|主任|組|處|室|#\d+|\(#?\d+\))`;
  return new RegExp(`^${rocDate}${weekday}.{0,40}${reportUnitOrMarker}`).test(compact);
}

async function queueTodoReviewText(replyToken, userId, text) {
  if (!userId) throw new Error('缺少 LINE userId，無法建立審核頁');
  await acknowledgeTodoBatch(replyToken, userId);
  queueTodoReviewSource(userId, {
    sourceType: 'LINE',
    sourceLabel: 'LINE 訊息',
    sourceText: text
  });
}

async function queueTodoReviewFile(replyToken, userId, message) {
  const fileName = String((message && message.fileName) || '').trim();
  const messageId = String((message && message.id) || '').trim();
  if (!messageId) throw new Error('LINE 檔案缺少 message id');
  if (!/\.pdf$/i.test(fileName)) {
    await respondToLineEvent(replyToken, userId, { type: 'text', text: '目前只會自動整理 PDF 檔案。' });
    return;
  }
  if (!userId) throw new Error('缺少 LINE userId，無法建立審核頁');

  await acknowledgeTodoBatch(replyToken, userId);
  const pdfBuffer = await downloadLineMessageContent(messageId);
  if (pdfBuffer.length > MAX_INLINE_PDF_BYTES) {
    throw new Error('PDF 超過 18MB，請先壓縮或拆成較小檔案再傳。');
  }
  takeRecentTodoWindowTexts(userId).forEach((source) => queueTodoReviewSource(userId, source));
  queueTodoReviewSource(userId, {
    sourceType: 'PDF',
    sourceLabel: fileName || 'LINE PDF',
    sourceText: fileName ? `檔名：${fileName}` : 'LINE PDF',
    pdfBuffer
  });
}

async function acknowledgeTodoBatch(replyToken, userId) {
  const batch = pendingTodoSourceBatches.get(userId);
  if (batch && batch.notified) return;
  await respondToLineEvent(replyToken, userId, {
    type: 'text',
    text: `收到，會把接下來 ${Math.round(TODO_BATCH_DEBOUNCE_MS / 1000)} 秒內傳來的報告訊息與 PDF 合併成同一個審核頁。`
  });
}

function queueTodoReviewSource(userId, source) {
  var batch = pendingTodoSourceBatches.get(userId);
  if (!batch) {
    batch = { createdAt: Date.now(), lineUserId: userId, sources: [], notified: true, timer: null };
    pendingTodoSourceBatches.set(userId, batch);
  }
  batch.sources.push(source);
  if (batch.timer) clearTimeout(batch.timer);
  batch.timer = setTimeout(function() {
    finalizeTodoReviewBatch(userId).catch(function(error) {
      console.error('[finalizeTodoReviewBatch]', error && error.stack ? error.stack : error);
      pushLineMessage(userId, {
        type: 'text',
        text: `整理待辦候選清單失敗：${error.message || error}`
      }).catch(function(deliveryError) {
        console.error('[todo batch delivery fallback]', deliveryError && deliveryError.stack ? deliveryError.stack : deliveryError);
      });
    });
  }, TODO_BATCH_DEBOUNCE_MS);
}

async function finalizeTodoReviewBatch(userId) {
  const batch = pendingTodoSourceBatches.get(userId);
  if (!batch || !batch.sources || !batch.sources.length) return;
  pendingTodoSourceBatches.delete(userId);

  const review = await createTodoReview({
    sourceType: batch.sources.some((source) => source.sourceType === 'PDF') ? 'MIXED' : 'LINE',
    sourceLabel: buildTodoBatchSourceLabel(batch.sources),
    sourceText: buildTodoBatchSourceText(batch.sources),
    reportDate: extractTodoReportDateFromSources(batch.sources),
    pdfBuffers: batch.sources.map((source) => source.pdfBuffer).filter(Boolean),
    lineUserId: userId
  });
  await pushLineMessage(userId, buildTodoReviewLinkFlexMessage(review));
}

function buildTodoBatchSourceLabel(sources) {
  const pdfCount = sources.filter((source) => source.sourceType === 'PDF').length;
  const lineCount = sources.filter((source) => source.sourceType === 'LINE').length;
  const parts = [];
  if (lineCount) parts.push(`${lineCount} 則 LINE 訊息`);
  if (pdfCount) parts.push(`${pdfCount} 份 PDF`);
  return parts.join(' + ') || 'LINE 批次資料';
}

function buildTodoBatchSourceText(sources) {
  return sources.map((source, index) => {
    const header = `【來源 ${index + 1}｜${source.sourceType || 'LINE'}｜${source.sourceLabel || ''}】`;
    return [header, String(source.sourceText || '').trim()].filter(Boolean).join('\n');
  }).join('\n\n');
}

function extractTodoReportDateFromSources(sources) {
  for (const source of sources || []) {
    const dateText = extractTodoReportDateText([
      source && source.sourceLabel,
      source && source.sourceText
    ].filter(Boolean).join('\n'));
    if (dateText) return dateText;
  }
  return '';
}

function extractTodoReportDateText(text) {
  const value = String(text || '').trim();
  if (!value) return '';

  let match = value.match(/(?:^|[^\d])(\d{3})[./-]?(\d{1,2})[./-]?(\d{1,2})(?:[^\d]|$)/);
  if (match) return `${Number(match[1])}/${Number(match[2])}/${Number(match[3])}`;

  match = value.match(/(?:^|[^\d])(\d{1,2})[./-](\d{1,2})(?:[^\d]|$)/);
  if (match) return `${Number(match[1])}/${Number(match[2])}`;

  return '';
}

async function createTodoReview(input) {
  cleanupExpiredDrafts();
  const candidates = await callGeminiForTodoCandidates(input);
  if (!candidates.length) throw new Error('Gemini 沒有整理出待辦候選項目');

  const id = crypto.randomBytes(16).toString('hex');
  const review = {
    id,
    createdAt: Date.now(),
    sourceType: input.sourceType,
    sourceLabel: input.sourceLabel,
    sourceText: input.sourceText,
    reportDate: input.reportDate || '',
    lineUserId: input.lineUserId,
    candidates,
    reviewUrl: buildReviewUrl(id)
  };
  pendingTodoReviews.set(id, review);
  await saveTodoReviewRecordToGas(review);
  return review;
}

function buildReviewUrl(id) {
  if (!REVIEW_BASE_URL) {
    throw new Error('請在 Render 環境變數設定 REVIEW_BASE_URL，例如 https://你的服務.onrender.com');
  }
  return `${REVIEW_BASE_URL}/todo-review?id=${encodeURIComponent(id)}`;
}

async function downloadLineMessageContent(messageId) {
  return requestBuffer(`https://api-data.line.me/v2/bot/message/${encodeURIComponent(messageId)}/content`, {
    method: 'GET',
    headers: { Authorization: `Bearer ${LINE_CHANNEL_ACCESS_TOKEN}` }
  });
}

async function callGeminiForTodoCandidates(input) {
  if (!GEMINI_API_KEY) throw new Error('請在 Render 環境變數設定 GEMINI_API_KEY');

  const parts = [];
  const pdfBuffers = input.pdfBuffers || (input.pdfBuffer ? [input.pdfBuffer] : []);
  for (const pdfBuffer of pdfBuffers) {
    parts.push({
      inline_data: {
        mime_type: 'application/pdf',
        data: pdfBuffer.toString('base64')
      }
    });
  }
  parts.push({ text: buildTodoGeminiPrompt(input) });

  const geminiUrl = new URL(`https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(GEMINI_MODEL)}:generateContent`);
  geminiUrl.searchParams.set('key', GEMINI_API_KEY);

  const response = await requestJson(geminiUrl, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      contents: [{ role: 'user', parts }],
      generationConfig: {
        temperature: 0.2,
        response_mime_type: 'application/json'
      }
    })
  });

  if (response.error) throw new Error(response.error.message || 'Gemini returned an error');
  const text = extractGeminiText(response);
  const parsed = parseGeminiJson(text);
  const rawItems = Array.isArray(parsed) ? parsed : (Array.isArray(parsed.items) ? parsed.items : []);
  return normalizeTodoCandidates(rawItems, input);
}

function buildTodoGeminiPrompt(input) {
  const sourceType = input.sourceType || 'LINE';
  const sourceText = String(input.sourceText || '').trim();
  return [
    '你是台灣國小導師的行政待辦整理助手。請從來源內容整理出需要老師後續處理的待辦候選清單。',
    '',
    '請嚴格回傳 JSON，格式如下：',
    '{"items":[{"title":"事件及日期時間清楚的標題","subtitle":"補充說明，可包含表單或連結摘要","dueDate":"YYYY-MM-DD 或空字串","details":"來源詳情完整整理","sourceType":"PDF、LINE 或 MIXED","sourceText":"原始來源文字；若來源含 LINE，一定要放完整訊息","teacherMessage":"給導師看的說明","parentMessage":"若需要轉傳家長群組，提供可直接轉傳文字，否則空字串","links":[{"label":"表單或連結名稱","url":"https://..."}]}]}',
    '',
    '規則：',
    '1. 標題要明確，列出清楚的事件及日期時間。',
    '2. subtitle 放標題下方小字，列出補充說明；若有表單或連結，links 必須抽出可點選 URL。',
    '3. details 要完整列出與標題相關的來源詳情，並標示來源是 PDF 還是 LINE。',
    '4. 若來源為 LINE，sourceText 必須是完整 LINE 訊息，不可摘要。',
    '4a. 若同一批包含多則 LINE 訊息與 PDF，請交叉整合重複事項，輸出一份去重後的候選清單，不要依來源機械分開。',
    '5. 若內容需要傳達到家長群組，parentMessage 要寫可直接轉傳到家長群組的版本；teacherMessage 寫給導師自己看的版本。',
    '6. 只保留需要行動或追蹤的事項；不要加入寒暄或不存在的資訊。',
    '7. LINE 訊息開頭常是民國日期、處室/組別與報告編號，例如「1150506學務處轉知」或「115.5.6(三)國際組(#713)晨會報告」，請優先把這些資訊整理進標題或 subtitle。',
    '8. 篩選優先順序：全校相關、四年級相關、四學年相關、需要導師班級宣導或提醒學生的內容。這些內容即使沒有明確表單或截止日期，也要列為候選。',
    '9. 不要漏掉活動、聚餐、謝師宴、畢業/期末相關、晨會要求、務必向學生宣導、請導師提醒、請轉知家長或學生等事項。',
    '10. 排除明顯只屬於其他年級、其他處室內部作業、純知會且不需老師行動的內容；若不確定是否與四年級或全校相關，請保留為候選並在 subtitle 說明不確定原因。',
    '',
    `來源類型：${sourceType}`,
    sourceText ? `來源文字：\n${sourceText}` : ''
  ].join('\n');
}

function extractGeminiText(response) {
  const candidates = Array.isArray(response && response.candidates) ? response.candidates : [];
  const parts = candidates[0] && candidates[0].content && Array.isArray(candidates[0].content.parts)
    ? candidates[0].content.parts
    : [];
  const text = parts.map((part) => String((part && part.text) || '')).join('').trim();
  if (!text) throw new Error('Gemini 沒有回傳文字內容');
  return text;
}

function parseGeminiJson(text) {
  const raw = String(text || '').trim();
  try {
    return JSON.parse(raw);
  } catch (_error) {
    const match = raw.match(/```(?:json)?\s*([\s\S]*?)```/) || raw.match(/(\{[\s\S]*\}|\[[\s\S]*\])/);
    if (match) return JSON.parse(match[1]);
    throw new Error('Gemini 回傳格式不是可解析的 JSON');
  }
}

function normalizeTodoCandidates(items, input) {
  return (items || []).map((item, index) => {
    const title = String(item && item.title || '').trim();
    if (!title) return null;
    const links = Array.isArray(item.links) ? item.links : [];
    return {
      id: `item-${index + 1}`,
      title,
      subtitle: String(item.subtitle || '').trim(),
      dueDate: normalizeTodoDueDate(item.dueDate),
      details: String(item.details || '').trim(),
      sourceType: String(item.sourceType || input.sourceType || '').trim(),
      sourceText: String(item.sourceText || input.sourceText || '').trim(),
      teacherMessage: String(item.teacherMessage || '').trim(),
      parentMessage: String(item.parentMessage || '').trim(),
      links: links.map((link) => ({
        label: String(link && (link.label || link.url) || '').trim(),
        url: String(link && link.url || '').trim()
      })).filter((link) => /^https?:\/\//i.test(link.url))
    };
  }).filter(Boolean).slice(0, 20);
}

function normalizeTodoDueDate(value) {
  const text = String(value || '').trim();
  if (!text) return '';
  const isoMatch = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (isoMatch) return `${isoMatch[1]}-${String(Number(isoMatch[2])).padStart(2, '0')}-${String(Number(isoMatch[3])).padStart(2, '0')}`;
  const mdMatch = text.match(/^(\d{1,2})[/-](\d{1,2})$/);
  if (mdMatch) {
    const today = getTodayInTimeZone();
    return `${today.getFullYear()}-${String(Number(mdMatch[1])).padStart(2, '0')}-${String(Number(mdMatch[2])).padStart(2, '0')}`;
  }
  return '';
}

function buildTodoReviewReportDateContents(review) {
  const reportDate = String(review && review.reportDate || '').trim();
  if (!reportDate) return [];
  return [
    { type: 'text', text: '\u6668\u6703\u5831\u544a\u65e5\u671f\uff1a' + reportDate, size: 'sm', color: '#365b42', wrap: true }
  ];
}

function buildTodoReviewLinkFlexMessage(review) {
  return {
    type: 'flex',
    altText: '待辦候選清單已建立',
    contents: {
      type: 'bubble',
      size: 'mega',
      header: {
        type: 'box',
        layout: 'vertical',
        backgroundColor: '#dff3df',
        paddingAll: '14px',
        contents: [{ type: 'text', text: '待辦候選清單', weight: 'bold', size: 'xl', color: '#145c2e' }]
      },
      body: {
        type: 'box',
        layout: 'vertical',
        spacing: 'sm',
        contents: [
          { type: 'text', text: `已整理 ${review.candidates.length} 個候選項目`, size: 'md', color: '#365b42', wrap: true },
          ...buildTodoReviewReportDateContents(review),
          { type: 'text', text: '請打開審核頁勾選要寫入 Google Sheet 的待辦事項。', size: 'sm', color: '#6b7f70', wrap: true }
        ]
      },
      footer: {
        type: 'box',
        layout: 'vertical',
        spacing: 'sm',
        contents: [
          { type: 'button', style: 'primary', color: '#1f8f4d', action: { type: 'uri', label: '開啟審核頁', uri: review.reviewUrl } }
        ]
      }
    }
  };
}

async function serveTodoReviewPage(_req, res, url) {
  cleanupExpiredDrafts();
  const id = String(url.searchParams.get('id') || '').trim();
  const review = await loadTodoReview(id);
  if (!review) {
    return sendHtml(res, 404, buildTodoReviewMissingHtml());
  }
  return sendHtml(res, 200, buildTodoReviewHtml(review));
}

async function commitTodoReview(req, res, url) {
  cleanupExpiredDrafts();
  const idFromQuery = String(url.searchParams.get('id') || '').trim();
  const raw = await readRawBody(req);
  let payload = {};
  try {
    payload = raw ? JSON.parse(raw) : {};
  } catch (_error) {
    return sendJson(res, 400, { ok: false, error: 'Invalid JSON body' });
  }

  const id = String(payload.id || idFromQuery || '').trim();
  const review = await loadTodoReview(id);
  if (!review) return sendJson(res, 404, { ok: false, error: '審核資料已過期或不存在' });
  if (review.committedAt) return sendJson(res, 409, { ok: false, error: '這份審核清單已經寫入過待辦事項' });

  const selectedIds = Array.isArray(payload.selectedIds) ? payload.selectedIds.map(String) : [];
  const selected = review.candidates.filter((item) => selectedIds.includes(item.id));
  if (!selected.length) return sendJson(res, 400, { ok: false, error: '請至少勾選一個待辦事項' });

  let latestItems = [];
  for (const item of selected) {
    latestItems = await addTodoItemToGas(item.title, item.dueDate);
  }

  review.committedAt = Date.now();
  review.selectedIds = selectedIds;
  await markTodoReviewCommittedInGas(id, selectedIds);
  pendingTodoReviews.delete(id);
  if (review.lineUserId) {
    pushLineMessage(review.lineUserId, {
      type: 'text',
      text: `已加入 ${selected.length} 個待辦事項到 Google Sheet。`
    }).catch((error) => console.error('[todo review notify]', error && error.stack ? error.stack : error));
  }
  return sendJson(res, 200, { ok: true, count: selected.length, items: latestItems });
}

function buildTodoReviewMissingHtml() {
  return '<!doctype html><meta charset="utf-8"><title>待辦審核</title><body style="font-family:system-ui;padding:24px"><h1>審核頁不存在或已過期</h1><p>請重新從 LINE 傳送 PDF 或用「整理待辦」建立候選清單。</p></body>';
}

async function loadTodoReview(id) {
  const reviewId = String(id || '').trim();
  if (!reviewId) return null;
  const memoryReview = pendingTodoReviews.get(reviewId);
  if (memoryReview) return memoryReview;
  try {
    const gasReview = await fetchTodoReviewRecordFromGas(reviewId);
    if (gasReview) pendingTodoReviews.set(reviewId, gasReview);
    return gasReview;
  } catch (error) {
    console.error('[loadTodoReview]', error && error.stack ? error.stack : error);
    return null;
  }
}

function buildTodoReviewHtml(review) {
  const payload = JSON.stringify({
    id: review.id,
    sourceLabel: review.sourceLabel,
    sourceType: review.sourceType,
    candidates: review.candidates,
    committedAt: review.committedAt || '',
    selectedIds: review.selectedIds || []
  }).replace(/</g, '\\u003c');

  return `<!doctype html>
<html lang="zh-Hant">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>待辦候選審核</title>
  <style>
    :root { color-scheme: light; --green:#1f8f4d; --green-dark:#145c2e; --mint:#e8f7e8; --line:#d7ead7; --text:#1f3328; }
    * { box-sizing: border-box; }
    body { margin: 0; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "Noto Sans TC", sans-serif; background: #f4f8f3; color: var(--text); }
    .app { max-width: 760px; margin: 0 auto; min-height: 100vh; background: #f9fcf8; }
    header { position: sticky; top: 0; z-index: 3; padding: 16px; background: #dff3df; border-bottom: 1px solid var(--line); }
    h1 { margin: 0; font-size: 22px; color: var(--green-dark); letter-spacing: 0; }
    .meta { margin-top: 4px; color: #5c7464; font-size: 13px; }
    main { padding: 14px; }
    .card { margin-bottom: 12px; border: 1px solid var(--line); border-radius: 8px; background: white; overflow: hidden; box-shadow: 0 1px 2px rgba(20, 92, 46, .06); }
    .card-top { display: grid; grid-template-columns: 34px 1fr; gap: 8px; padding: 12px; align-items: start; }
    input[type="checkbox"] { width: 22px; height: 22px; accent-color: var(--green); margin-top: 2px; }
    .title { font-size: 17px; font-weight: 700; color: #173d27; line-height: 1.35; }
    .subtitle { margin-top: 4px; font-size: 13px; color: #637568; line-height: 1.45; white-space: pre-wrap; }
    .due { display: inline-block; margin-top: 8px; padding: 3px 8px; border-radius: 999px; background: var(--mint); color: var(--green-dark); font-size: 12px; font-weight: 700; }
    .links { margin-top: 8px; display: flex; flex-wrap: wrap; gap: 6px; }
    .links a, .line-btn { border: 1px solid #b7dbb9; color: var(--green-dark); background: #f3fbf3; border-radius: 999px; padding: 6px 10px; font-size: 13px; text-decoration: none; font-weight: 700; }
    details { border-top: 1px solid #edf3ed; }
    summary { cursor: pointer; list-style: none; padding: 10px 12px; color: var(--green-dark); font-size: 14px; font-weight: 700; }
    summary::-webkit-details-marker { display: none; }
    .detail { padding: 0 12px 12px; color: #3d4b41; font-size: 14px; line-height: 1.6; }
    .block { margin-top: 10px; padding: 10px; background: #f6faf6; border: 1px solid #e2efe2; border-radius: 8px; white-space: pre-wrap; }
    .source { color: #6d7f71; font-size: 12px; margin-bottom: 6px; font-weight: 700; }
    .bar { position: sticky; bottom: 0; display: grid; grid-template-columns: 1fr auto; gap: 10px; padding: 12px 14px; background: rgba(249, 252, 248, .96); border-top: 1px solid var(--line); backdrop-filter: blur(8px); }
    button { border: 0; border-radius: 8px; padding: 12px 14px; background: var(--green); color: white; font-weight: 800; font-size: 15px; }
    button:disabled { opacity: .55; }
    .count { align-self: center; color: #5c7464; font-size: 14px; }
    .toast { min-height: 22px; padding: 0 14px 12px; color: var(--green-dark); font-weight: 700; }
    @media (max-width: 520px) {
      main { padding: 10px; }
      .card-top { grid-template-columns: 30px 1fr; padding: 11px; }
      .title { font-size: 16px; }
      .bar { grid-template-columns: 1fr; }
      button { width: 100%; }
    }
  </style>
</head>
<body>
  <div class="app">
    <header>
      <h1>待辦候選審核</h1>
      <div class="meta" id="meta"></div>
    </header>
    <main id="list"></main>
    <div class="bar">
      <div class="count" id="count">已勾選 0 項</div>
      <button id="submit">加入勾選項目</button>
    </div>
    <div class="toast" id="toast"></div>
  </div>
  <script>
    const review = ${payload};
    const list = document.getElementById('list');
    const count = document.getElementById('count');
    const submit = document.getElementById('submit');
    const toast = document.getElementById('toast');
    document.getElementById('meta').textContent = (review.sourceType || '') + ' - ' + (review.sourceLabel || '') + ' - ' + review.candidates.length + ' 個候選項目';
    const committed = Boolean(review.committedAt);
    const selectedSet = new Set((review.selectedIds || []).map(String));

    function esc(value) {
      return String(value || '').replace(/[&<>"']/g, ch => ({ '&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;' }[ch]));
    }
    function linkify(text) {
      return esc(text).replace(/(https?:\\/\\/[^\\s<]+)/g, '<a href="$1" target="_blank" rel="noopener">$1</a>');
    }
    function normalizeCompareText(text) {
      return String(text || '').replace(/\\s+/g, '').trim();
    }
    function shouldShowSourceText(details, sourceText) {
      const a = normalizeCompareText(details);
      const b = normalizeCompareText(sourceText);
      if (!b) return false;
      if (!a) return true;
      return a !== b && !a.includes(b) && !b.includes(a);
    }
    function updateCount() {
      const n = document.querySelectorAll('input[type="checkbox"]:checked').length;
      count.textContent = '已勾選 ' + n + ' 項';
    }
    function render() {
      list.innerHTML = review.candidates.map(item => {
        const checked = committed ? selectedSet.has(String(item.id)) : true;
        const links = (item.links || []).map(link => '<a href="' + esc(link.url) + '" target="_blank" rel="noopener">' + esc(link.label || '開啟連結') + '</a>').join('');
        const parentLink = item.parentMessage ? '<a class="line-btn" href="line://msg/text/' + encodeURIComponent(item.parentMessage) + '">傳送到 LINE</a>' : '';
        const detailsText = item.details || '';
        const sourceText = shouldShowSourceText(detailsText, item.sourceText) ? item.sourceText : '';
        return '<article class="card">'
          + '<div class="card-top"><input type="checkbox" data-id="' + esc(item.id) + '"' + (checked ? ' checked' : '') + (committed ? ' disabled' : '') + '><div>'
          + '<div class="title">' + esc(item.title) + '</div>'
          + (item.subtitle ? '<div class="subtitle">' + linkify(item.subtitle) + '</div>' : '')
          + (item.dueDate ? '<span class="due">' + esc(item.dueDate) + '</span>' : '')
          + ((links || parentLink) ? '<div class="links">' + links + parentLink + '</div>' : '')
          + '</div></div>'
          + '<details><summary>展開來源詳情</summary><div class="detail">'
          + '<div class="source">來源：' + esc(item.sourceType || review.sourceType || '') + '</div>'
          + (detailsText ? '<div class="block">' + linkify(detailsText) + '</div>' : '')
          + (sourceText ? '<div class="block">' + linkify(sourceText) + '</div>' : '')
          + (item.teacherMessage ? '<div class="block"><strong>給導師看</strong>\\n' + linkify(item.teacherMessage) + '</div>' : '')
          + (item.parentMessage ? '<div class="block"><strong>轉傳家長群組</strong>\\n' + linkify(item.parentMessage) + '</div>' : '')
          + '</div></details></article>';
      }).join('');
      document.querySelectorAll('input[type="checkbox"]').forEach(box => box.addEventListener('change', updateCount));
      if (committed) {
        submit.disabled = true;
        submit.textContent = '已加入待辦';
        toast.textContent = '這份審核清單已寫入過，頁面保留供回看。';
      }
      updateCount();
    }
    submit.addEventListener('click', async () => {
      if (committed) return;
      const selectedIds = Array.from(document.querySelectorAll('input[type="checkbox"]:checked')).map(box => box.dataset.id);
      if (!selectedIds.length) { toast.textContent = '請先勾選至少一個項目。'; return; }
      submit.disabled = true;
      toast.textContent = '正在寫入 Google Sheet...';
      try {
        const resp = await fetch('/api/todo-review/commit', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ id: review.id, selectedIds })
        });
        const data = await resp.json();
        if (!data.ok) throw new Error(data.error || '寫入失敗');
        toast.textContent = '已加入 ' + data.count + ' 個待辦事項。';
        submit.textContent = '已加入';
      } catch (error) {
        submit.disabled = false;
        toast.textContent = error.message || String(error);
      }
    });
    render();
  </script>
</body>
</html>`;
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

async function saveTodoReviewRecordToGas(review) {
  const payload = Buffer.from(JSON.stringify(review), 'utf8').toString('base64url');
  const chunkSize = 1500;
  const totalChunks = Math.max(1, Math.ceil(payload.length / chunkSize));
  let response = { ok: true };

  for (let i = 0; i < totalChunks; i++) {
    const gasUrl = new URL(GAS_BASE_URL);
    gasUrl.searchParams.set('action', 'saveTodoReviewRecordChunk');
    gasUrl.searchParams.set('reviewId', review.id);
    gasUrl.searchParams.set('chunkIndex', String(i));
    gasUrl.searchParams.set('totalChunks', String(totalChunks));
    gasUrl.searchParams.set('chunkEncoding', 'base64url');
    gasUrl.searchParams.set('chunk', payload.slice(i * chunkSize, (i + 1) * chunkSize));
    response = await requestJson(gasUrl, { method: 'GET' });
    if (!response.ok) throw new Error(response.error || '儲存待辦審核紀錄失敗');
  }

  return response;
}

async function fetchTodoReviewRecordFromGas(reviewId) {
  const gasUrl = new URL(GAS_BASE_URL);
  gasUrl.searchParams.set('action', 'getTodoReviewRecord');
  gasUrl.searchParams.set('reviewId', reviewId);
  const response = await requestJson(gasUrl, { method: 'GET' });
  if (!response.ok) throw new Error(response.error || '讀取待辦審核紀錄失敗');
  return response.record || null;
}

async function markTodoReviewCommittedInGas(reviewId, selectedIds) {
  const gasUrl = new URL(GAS_BASE_URL);
  gasUrl.searchParams.set('action', 'markTodoReviewCommitted');
  gasUrl.searchParams.set('reviewId', reviewId);
  gasUrl.searchParams.set('selectedIds', JSON.stringify(selectedIds || []));
  const response = await requestJson(gasUrl, { method: 'GET' });
  if (!response.ok) throw new Error(response.error || '更新待辦審核紀錄失敗');
  return response;
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
  if (replyToken) {
    try {
      return await replyLineMessage(replyToken, message);
    } catch (error) {
      if (!userId) throw error;
      console.warn('[reply fallback to push]', error && error.message ? error.message : error);
      return pushLineMessage(userId, message);
    }
  }
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
    const targetDescription = describeRequestTarget(url);
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
            return reject(new Error(`HTTP ${statusCode} ${targetDescription}: too many redirects`));
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
          return reject(new Error(`HTTP ${statusCode} ${targetDescription}`));
        }

        let parsed;
        try {
          parsed = JSON.parse(raw);
        } catch (_error) {
          if (statusCode >= 200 && statusCode < 300) return resolve({ ok: true, raw });
          const compactRaw = raw.replace(/\s+/g, ' ').replace(/<[^>]+>/g, ' ').trim().slice(0, 180);
          return reject(new Error(`HTTP ${statusCode} ${targetDescription}: ${compactRaw || 'Non-JSON response'}`));
        }

        if (statusCode >= 200 && statusCode < 300) return resolve(parsed);
        return reject(new Error(`HTTP ${statusCode} ${targetDescription}: ${formatHttpErrorBody(parsed, raw)}`));
      });
    });

    request.on('error', reject);
    if (requestOptions.body) request.write(requestOptions.body);
    request.end();
  });
}

function describeRequestTarget(url) {
  const action = url.searchParams ? url.searchParams.get('action') : '';
  const modelMatch = String(url.pathname || '').match(/\/models\/([^/:]+):/);
  const detail = action
    ? ` action=${action}`
    : (modelMatch ? ` model=${decodeURIComponent(modelMatch[1])}` : '');
  return `${url.hostname}${url.pathname}${detail}`;
}

function formatHttpErrorBody(parsed, raw) {
  if (parsed && parsed.error) {
    if (typeof parsed.error === 'string') return parsed.error;
    if (parsed.error.message) return parsed.error.message;
    try {
      return JSON.stringify(parsed.error).slice(0, 500);
    } catch (_error) {}
  }
  return String(raw || '').replace(/\s+/g, ' ').replace(/<[^>]+>/g, ' ').trim().slice(0, 500);
}

function requestBuffer(targetUrl, options, redirectCount) {
  return new Promise((resolve, reject) => {
    const currentRedirectCount = Number(redirectCount || 0);
    const url = typeof targetUrl === 'string' ? new URL(targetUrl) : targetUrl;
    const requestOptions = Object.assign({}, options || {});
    const client = url.protocol === 'https:' ? https : http;

    const request = client.request(url, requestOptions, (response) => {
      const chunks = [];
      response.on('data', (chunk) => chunks.push(chunk));
      response.on('end', () => {
        const statusCode = response.statusCode || 500;
        if ([301, 302, 303, 307, 308].includes(statusCode) && response.headers.location) {
          if (currentRedirectCount >= 5) {
            return reject(new Error(`HTTP ${statusCode}: too many redirects`));
          }
          const nextUrl = new URL(response.headers.location, url);
          return resolve(requestBuffer(nextUrl, requestOptions, currentRedirectCount + 1));
        }

        const body = Buffer.concat(chunks);
        if (statusCode >= 200 && statusCode < 300) return resolve(body);
        return reject(new Error(`HTTP ${statusCode}: ${body.toString('utf8').slice(0, 180)}`));
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

function sendHtml(res, statusCode, html) {
  res.writeHead(statusCode, { 'Content-Type': 'text/html; charset=utf-8' });
  res.end(html);
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
