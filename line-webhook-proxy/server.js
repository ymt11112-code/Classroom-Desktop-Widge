const http = require('http');
const https = require('https');
const crypto = require('crypto');
const fs = require('fs');
const path = require('path');
const { URL, URLSearchParams } = require('url');

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
    if (!res.headersSent) {
      return sendJson(res, 500, { ok: false, error: 'Internal server error' });
    }
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

      const dateStr = parseQueryDate(text);
      if (dateStr) {
        const digest = await fetchLineDigest(dateStr);
        await respondToLineEvent(replyToken, userId, buildFlexMessage(digest));
        console.log(`[event ${dateStr}] delivered`);
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
  if (mdMatch) {
    return `${Number(mdMatch[1])}/${Number(mdMatch[2])}`;
  }

  return null;
}

function looksLikeAbsenceMessage(text) {
  const value = normalizeAbsenceText(text);
  const keywords = [
    '請假', '病假', '事假', '公假', '喪假', '曠課', '遲到', '早退', '退餐',
    '身體不舒服', '不舒服', '發燒', '感冒', '咳嗽', '腸胃炎', '腹瀉',
    '看醫生', '就醫', '回診', '住院', '家裡有事', '臨時有事',
    '無法到校', '不能到校', '不克到校', '在家休息'
  ];
  return keywords.some((keyword) => value.includes(keyword));
}

function isConfirmAbsenceCommand(text) {
  const value = String(text || '').trim();
  return value === '確認' || value === '確認請假' || value === '送出請假';
}

function isCancelAbsenceCommand(text) {
  const value = String(text || '').trim();
  return value === '取消' || value === '取消請假';
}

async function prepareAbsenceDraft(replyToken, userId, rawText) {
  const students = await fetchStudentList();
  const records = buildParsedAbsenceRecords(rawText, students);
  if (!records.length) {
    throw new Error('找不到可辨識的請假學生，請補上座號或姓名');
  }
  if (!userId) {
    throw new Error('目前無法建立請假草稿，因為缺少使用者識別');
  }

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
  if (!userId) {
    throw new Error('目前無法確認請假，因為缺少使用者識別');
  }

  const pending = pendingAbsenceDrafts.get(userId);
  if (!pending || !Array.isArray(pending.records) || !pending.records.length) {
    await respondToLineEvent(replyToken, userId, {
      type: 'text',
      text: '目前沒有待確認的請假草稿。請先傳送請假訊息。'
    });
    return;
  }

  for (const record of pending.records) {
    await saveAbsenceRecordToGas(record);
  }

  pendingAbsenceDrafts.delete(userId);
  await respondToLineEvent(replyToken, userId, {
    type: 'text',
    text: `已記錄 ${pending.records.length} 筆出缺席資料。`
  });
}

async function cancelPendingAbsence(replyToken, userId) {
  if (userId) pendingAbsenceDrafts.delete(userId);
  await respondToLineEvent(replyToken, userId, {
    type: 'text',
    text: '已取消待確認的請假草稿。'
  });
}

function cleanupExpiredDrafts() {
  const now = Date.now();
  for (const [userId, draft] of pendingAbsenceDrafts.entries()) {
    if (!draft || !draft.createdAt || now - draft.createdAt > PENDING_DRAFT_TTL_MS) {
      pendingAbsenceDrafts.delete(userId);
    }
  }
}

async function fetchLineDigest(dateStr) {
  const gasUrl = new URL(GAS_BASE_URL);
  gasUrl.searchParams.set('action', 'getLineDigest');
  gasUrl.searchParams.set('date', dateStr);

  const response = await requestJson(gasUrl, { method: 'GET' });
  if (!response.ok) {
    throw new Error(response.error || 'GAS returned an error');
  }
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
  if (!response.ok) {
    throw new Error(response.error || '無法取得學生名單');
  }

  studentCache.students = Array.isArray(response.students) ? response.students : [];
  studentCache.expiresAt = now + STUDENT_CACHE_TTL_MS;
  return studentCache.students;
}

async function saveAbsenceRecordToGas(record) {
  const gasUrl = new URL(GAS_BASE_URL);
  const body = new URLSearchParams();
  body.set('action', 'saveAbsenceRecord');
  body.set('data', JSON.stringify(record));

  const response = await requestJson(gasUrl, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded; charset=utf-8'
    },
    body: body.toString()
  });

  if (!response.ok) {
    throw new Error(response.error || '寫入出缺席資料失敗');
  }
  return response;
}

function buildParsedAbsenceRecords(text, students) {
  const normalizedText = normalizeAbsenceText(text);
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

function resolvePeriodsValue(mode, startDate, endDate, inferredPeriods, days) {
  if (mode !== 'leave') return '';
  if (inferredPeriods) return inferredPeriods;

  const schoolDays = countWeekdaysInclusive(startDate, endDate);
  if (schoolDays > 1) {
    return String(schoolDays * 7);
  }

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

  if (meal !== '是') {
    return { eligible: false, message: '未申請退餐。' };
  }

  if (leaveType === '病假') {
    if (schoolDays >= 3) {
      return {
        eligible: true,
        message: `符合病假退餐規則：連續 ${schoolDays} 個上課日，通知當日不退費。`
      };
    }
    return {
      eligible: false,
      message: `不符合病假退餐規則：病假需連續 3 個上課日以上，目前僅 ${schoolDays} 個上課日。`
    };
  }

  if (leaveType === '事假') {
    if (schoolDays < 5) {
      return {
        eligible: false,
        message: `不符合事假退餐規則：事假需連續 5 個上課日以上，目前僅 ${schoolDays} 個上課日。`
      };
    }
    if (leadDays < 7) {
      return {
        eligible: false,
        message: `不符合事假退餐規則：事假需至少 7 天前通知，目前僅提前 ${leadDays} 天。`
      };
    }
    return {
      eligible: true,
      message: `符合事假退餐規則：連續 ${schoolDays} 個上課日，且已提前 ${leadDays} 天通知。`
    };
  }

  if (leaveType === '公假') {
    if (leadDays < 7) {
      return {
        eligible: false,
        message: `不符合公假退餐規則：公假需至少 7 天前通知，目前僅提前 ${leadDays} 天。`
      };
    }
    return {
      eligible: true,
      message: `符合公假退餐規則：已提前 ${leadDays} 天通知。`
    };
  }

  return {
    eligible: false,
    message: `目前 ${leaveType || '此假別'} 未符合可退餐規則。`
  };
}

function calculateLeadDays(noticeDate, startDate) {
  const notice = parseIsoDate(noticeDate);
  const start = parseIsoDate(startDate);
  if (!notice || !start) return 0;
  const diffMs = start.getTime() - notice.getTime();
  return Math.max(0, Math.floor(diffMs / (24 * 60 * 60 * 1000)));
}

function buildAbsenceDraftSummary(records) {
  const lines = ['請假分析結果：'];
  records.forEach((record, index) => {
    lines.push(
      `${index + 1}. ${record.selectedStudents[0]}`,
      `日期：${record.startDate}${record.endDate && record.endDate !== record.startDate ? ` ~ ${record.endDate}` : ''}`,
      `類型：${record.leaveType}`,
      `節數：${record.periods || '—'}`,
      `日數：${record.days || '—'}`,
      `退餐：${record.meal || '否'}`,
      `退餐判定：${record.mealReturned ? '可退' : '不可退'}`,
      `說明：${record.mealRuleMessage || '—'}`
    );
  });
  lines.push('', '回覆「確認請假」即可寫入出缺席，回覆「取消請假」可放棄。');
  return lines.join('\n');
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

function toIsoFromParts(year, month, day) {
  return [
    String(year),
    String(month).padStart(2, '0'),
    String(day).padStart(2, '0')
  ].join('-');
}

function parseIsoDate(iso) {
  if (!iso) return null;
  const date = new Date(`${iso}T00:00:00`);
  return Number.isNaN(date.getTime()) ? null : date;
}

function normalizeAbsenceText(text) {
  return String(text || '')
    .replace(/\u3000/g, ' ')
    .replace(/[，、；：]/g, ' ')
    .replace(/\s*([\/.\-~～])\s*/g, '$1')
    .replace(/\s+/g, ' ')
    .trim();
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
  const today = getTodayInTimeZone();
  const currentYear = today.getFullYear();
  const results = [];

  if (text.includes('今天') || text.includes('今日')) {
    results.push(toIsoFromParts(today.getFullYear(), today.getMonth() + 1, today.getDate()));
  }
  if (text.includes('明天')) {
    const next = addCalendarDays(today, 1);
    results.push(toIsoFromParts(next.getFullYear(), next.getMonth() + 1, next.getDate()));
  }
  if (text.includes('後天')) {
    const next = addCalendarDays(today, 2);
    results.push(toIsoFromParts(next.getFullYear(), next.getMonth() + 1, next.getDate()));
  }

  for (const match of text.matchAll(/(\d{4})[\/.-](\d{1,2})[\/.-](\d{1,2})/g)) {
    results.push(toIsoFromParts(parseInt(match[1], 10), parseInt(match[2], 10), parseInt(match[3], 10)));
  }

  for (const match of text.matchAll(/(^|[^\d])(\d{1,2})[\/.-](\d{1,2})(?!\d)/g)) {
    results.push(toIsoFromParts(currentYear, parseInt(match[2], 10), parseInt(match[3], 10)));
  }

  for (const match of text.matchAll(/(下週|下周|這週|這周|本週|本周|週|星期|禮拜)([一二三四五六日天])/g)) {
    const weekdayIndex = weekdayTextToIndex(match[2]);
    if (weekdayIndex === -1) continue;
    const useNextWeek = match[1] === '下週' || match[1] === '下周';
    const resolved = findUpcomingWeekday(today, weekdayIndex, useNextWeek);
    results.push(toIsoFromParts(resolved.getFullYear(), resolved.getMonth() + 1, resolved.getDate()));
  }

  return Array.from(new Set(results)).filter(Boolean);
}

function stripDatePhrases(text) {
  return String(text || '')
    .replace(/今天|今日|明天|後天/g, ' ')
    .replace(/(下週|下周|這週|這周|本週|本周|週|星期|禮拜)[一二三四五六日天]/g, ' ')
    .replace(/\d{1,2}[\/.-]\d{1,2}\s*[-~～]\s*\d{1,2}[\/.-]\d{1,2}/g, ' ')
    .replace(/\d{4}[\/.-]\d{1,2}[\/.-]\d{1,2}\s*[-~～]\s*\d{4}[\/.-]\d{1,2}[\/.-]\d{1,2}/g, ' ')
    .replace(/\d{4}[\/.-]\d{1,2}[\/.-]\d{1,2}/g, ' ')
    .replace(/\d{1,2}[\/.-]\d{1,2}/g, ' ');
}

function inferModeFromText(text) {
  if (text.includes('早退')) return 'early';
  if (text.includes('遲到')) return 'late';
  return 'leave';
}

function inferLeaveTypeFromText(text) {
  const explicit = LEAVE_TYPES.find((type) => text.includes(type));
  if (explicit) return explicit;

  if (/(身體不舒服|不舒服|發燒|感冒|咳嗽|腸胃炎|腹瀉|頭痛|看醫生|就醫|回診|住院|請病假)/.test(text)) {
    return '病假';
  }
  if (/(家裡有事|臨時有事|私人因素|家中有事|返鄉|掃墓|祭祖|旅遊|出國)/.test(text)) {
    return '事假';
  }
  if (/(比賽|演出|校外|代表隊|受訓|活動|參訪|公務)/.test(text)) {
    return '公假';
  }
  if (/(喪禮|告別式|治喪)/.test(text)) {
    return '喪假';
  }
  return '事假';
}

function inferMealChoiceFromText(text) {
  if (/(不退餐|免退餐|不用退餐|不用訂餐|午餐照常|有吃午餐)/.test(text)) return '否';
  if (/(退餐|停餐|不用餐|午餐不用|不吃午餐|取消午餐)/.test(text)) return '是';
  return null;
}

function inferPeriodsFromText(text) {
  const normalized = String(text || '')
    .replace(/第/g, '')
    .replace(/節/g, '')
    .replace(/[到至～~]/g, '-');

  const rangeMatch = normalized.match(/(\d+(?:\.\d+)?)\s*-\s*(\d+(?:\.\d+)?)/);
  if (rangeMatch) {
    const start = parseFloat(rangeMatch[1]);
    const end = parseFloat(rangeMatch[2]);
    if (!Number.isNaN(start) && !Number.isNaN(end) && end >= start) {
      return String(end - start + 1);
    }
  }

  const singleMatch = text.match(/(\d+(?:\.\d+)?)\s*節/);
  if (singleMatch) return String(singleMatch[1]);

  if (/(整天|全天|一天|全日)/.test(text)) return '7';
  if (/(半天|早上|上午|中午接回|午前|早退回家)/.test(text)) return '4';
  if (/(下午|午休後|午後)/.test(text)) return '3';
  return '';
}

function inferDurationDays(text) {
  const numericMatch = String(text || '').match(/(\d+)\s*天/);
  if (numericMatch) return parseInt(numericMatch[1], 10);

  const chineseDayMap = { 一: 1, 二: 2, 兩: 2, 三: 3, 四: 4, 五: 5, 六: 6, 七: 7 };
  const chineseMatch = String(text || '').match(/([一二兩三四五六七])\s*天/);
  if (chineseMatch) return chineseDayMap[chineseMatch[1]] || null;
  if (/半天/.test(text)) return 1;
  return null;
}

function inferStudentsFromText(text, students) {
  if (text.includes('全班')) return ['全班'];

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

function buildFlexMessage(digest) {
  const schedule = digest.schedule || {};
  const periods = Array.isArray(schedule.periods) ? schedule.periods : [];
  const weeklyNotes = Array.isArray(digest.weeklyNotes) ? digest.weeklyNotes : [];
  const contactText = String(digest.contact || '').trim();
  const headerText = `📅 ${digest.dayLabel || digest.date || ''}${digest.todayKey ? `　${digest.todayKey}` : ''}`;

  const tableRows = [
    {
      type: 'box',
      layout: 'horizontal',
      backgroundColor: '#dce8fb',
      paddingAll: '8px',
      contents: [
        { type: 'text', text: '節', size: 'lg', flex: 1, align: 'center', weight: 'bold', color: '#2a4070' },
        { type: 'text', text: '進度', size: 'lg', flex: 5, weight: 'bold', color: '#2a4070' },
        { type: 'text', text: '備註', size: 'lg', flex: 4, weight: 'bold', color: '#2a4070' }
      ]
    }
  ];

  if (periods.length) {
    for (const period of periods) {
      tableRows.push({
        type: 'box',
        layout: 'horizontal',
        backgroundColor: getRowColor(period),
        paddingAll: '8px',
        contents: [
          { type: 'text', text: period.period === '晨光' ? '晨' : String(period.period || ''), size: 'lg', flex: 1, align: 'center', color: '#5c6f92', weight: 'bold' },
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
    const client = url.protocol === 'https:' ? https : http;
    const request = client.request(url, options || {}, (response) => {
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
          const nextOptions = Object.assign({}, options || {});
          if (statusCode === 303) {
            nextOptions.method = 'GET';
            delete nextOptions.body;
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
        reject(new Error(`HTTP ${statusCode}: ${parsed.error || raw}`));
      });
    });

    request.on('error', reject);
    if (options && options.body) request.write(options.body);
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
