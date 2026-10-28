/**
 * INTERN BOT - Google Apps Script (NO CALENDAR VERSION)
 * Features: meeting invite by email, AI task-briefing chat,
 * post-meeting email, warning drafts with manual approval.
 *
 * Needs only: script property GEMINI_API_KEY (free from aistudio.google.com). No Services, no Calendar.
 */

const CONFIG = {
  BOT_NAME: 'Asha',
  COMPANY: 'Your Company',            // <-- change
  MENTOR_NAME: 'Your Mentor',         // <-- change
  MEETING_MINUTES: 30,
  // Open meet.google.com/new, click "Create a meeting for later", paste the link here.
  MEET_LINK: 'https://meet.google.com/xxx-xxxx-xxx',   // <-- change
  MODEL: 'gemini-2.5-flash-lite',   // free-tier model; see AI Studio for current names
  // Never put passwords here. Use SSO or a one-time access link instead.
  ACCESS_NOTE: 'Sign in with your company account. If you cannot get in, reply to this email and we will send you a secure one-time access link.'
};

/* ---------- Menu ---------- */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Intern Bot')
    .addItem('1. Set up sheets (first time)', 'setupSheets')
    .addItem('2. Send meeting invites', 'sendMeetingInvites')
    .addItem('3. Open task chatbot', 'showChat')
    .addItem('4. Send post-meeting emails now', 'sendPostMeetingEmails')
    .addItem('5. Draft warnings', 'draftWarnings')
    .addItem('6. Send approved warnings', 'sendApprovedWarnings')
    .addSeparator()
    .addItem('Turn on auto post-meeting emails', 'setupTriggers')
    .addToUi();
}

/* ---------- Helpers ---------- */

function readRows_(sheetName) {
  const sh = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sh) throw new Error('Sheet not found: ' + sheetName + '. Run "Set up sheets" first.');
  const values = sh.getDataRange().getValues();
  const headers = values.shift().map(function (h) { return String(h).trim(); });
  const rows = values.map(function (r, i) {
    const o = { _row: i + 2 };
    headers.forEach(function (h, j) { o[h] = r[j]; });
    return o;
  });
  return { sh: sh, headers: headers, rows: rows };
}

function setCell_(sh, headers, row, header, value) {
  sh.getRange(row, headers.indexOf(header) + 1).setValue(value);
}

function fmtDate_(d) {
  return Utilities.formatDate(new Date(d), Session.getScriptTimeZone(), 'EEE, d MMM yyyy');
}

function fmtDateTime_(d) {
  return Utilities.formatDate(new Date(d), Session.getScriptTimeZone(), 'EEE, d MMM yyyy, h:mm a');
}

function askClaude_(system, messages, maxTokens) {
  // (Name kept so the rest of the code works unchanged. This now calls FREE Google Gemini.)
  const key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!key) throw new Error('Add GEMINI_API_KEY in Project Settings > Script properties.');

  const contents = messages.map(function (m) {
    return { role: m.role === 'assistant' ? 'model' : 'user', parts: [{ text: m.content }] };
  });

  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + CONFIG.MODEL + ':generateContent';
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-goog-api-key': key },
    payload: JSON.stringify({
      system_instruction: { parts: [{ text: system }] },
      contents: contents,
      generationConfig: { maxOutputTokens: maxTokens || 600 }
    }),
    muteHttpExceptions: true
  });
  const data = JSON.parse(res.getContentText());
  if (res.getResponseCode() !== 200) {
    throw new Error('Gemini API error: ' + res.getContentText());
  }
  const parts = data.candidates[0].content.parts || [];
  return parts.map(function (p) { return p.text || ''; }).join('').trim();
}

/* ---------- 1. Sheet setup ---------- */

function setupSheets() {
  const ss = SpreadsheetApp.getActive();

  let interns = ss.getSheetByName('Interns');
  if (!interns) {
    interns = ss.insertSheet('Interns');
    interns.appendRow(['Name', 'Email', 'Task', 'TaskDescription', 'TaskURL',
                       'ClosureDate', 'MeetingTime', 'Status']);
    interns.appendRow(['Priya', 'priya@yourcompany.com', 'Build a login page',
      'Create a responsive login page with email and password validation. Deliver as a GitHub repo with a README.',
      'https://example.com/tasks/1', new Date('2026-11-10'), new Date('2026-10-30T11:00:00'), '']);
    interns.setFrozenRows(1);
  }

  let warn = ss.getSheetByName('Warnings');
  if (!warn) {
    warn = ss.insertSheet('Warnings');
    warn.appendRow(['Name', 'Mistake', 'Draft', 'Approved', 'Sent']);
    warn.getRange('D2:D200').insertCheckboxes();
    warn.setFrozenRows(1);
  }
  SpreadsheetApp.getUi().alert('Sheets ready. Fill the Interns sheet (one row per intern), then use the Intern Bot menu.');
}

/* ---------- 2. Meeting invites (by email, no calendar) ---------- */

function sendMeetingInvites() {
  const d = readRows_('Interns');
  let count = 0;

  d.rows.forEach(function (r) {
    if (!r.Email || !r.MeetingTime || r.Status) return;

    const body =
      'Hi ' + r.Name + ',\n\n' +
      'You are invited to a task briefing.\n\n' +
      'When: ' + fmtDateTime_(r.MeetingTime) + ' (' + CONFIG.MEETING_MINUTES + ' minutes)\n' +
      'Join here: ' + CONFIG.MEET_LINK + '\n' +
      'Your task: ' + r.Task + '\n' +
      'Closure date: ' + fmtDate_(r.ClosureDate) + '\n\n' +
      CONFIG.BOT_NAME + ', our AI assistant, will walk you through the task and answer your questions.\n\n' +
      'Please reply if this time does not work for you.\n\n' +
      CONFIG.MENTOR_NAME + '\n' + CONFIG.COMPANY;

    MailApp.sendEmail({
      to: r.Email,
      subject: 'Task briefing: ' + fmtDateTime_(r.MeetingTime),
      body: body
    });

    setCell_(d.sh, d.headers, r._row, 'Status', 'Invited');
    count++;
  });

  SpreadsheetApp.getActive().toast(count + ' invite(s) sent.', 'Intern Bot');
}

/* ---------- 3. Task chatbot (sidebar) ---------- */

function showChat() {
  const html = HtmlService.createHtmlOutputFromFile('Chat').setTitle(CONFIG.BOT_NAME + ' - Task Assistant');
  SpreadsheetApp.getUi().showSidebar(html);
}

function getInternNames() {
  return readRows_('Interns').rows
    .filter(function (r) { return r.Name; })
    .map(function (r) { return r.Name; });
}

function chatWithAsha(name, history) {
  const intern = readRows_('Interns').rows.filter(function (r) { return r.Name === name; })[0];
  if (!intern) throw new Error('Intern not found: ' + name);

  const system =
    'You are ' + CONFIG.BOT_NAME + ', the AI assistant for the ' + CONFIG.COMPANY + ' internship program. ' +
    'Always be upfront that you are an AI assistant. Speak in a warm, friendly, conversational tone ' +
    'with short sentences. Your replies may be read aloud, so do not use bullet points, markdown or emojis.\n\n' +
    'You are briefing ' + intern.Name + ' on their task.\n' +
    'Title: ' + intern.Task + '\n' +
    'Details: ' + intern.TaskDescription + '\n' +
    'Task URL: ' + intern.TaskURL + '\n' +
    'Closure date: ' + fmtDate_(intern.ClosureDate) + '\n\n' +
    'On the first message, greet them, introduce yourself as the team\'s AI assistant, and explain the task clearly. ' +
    'Then answer questions. Use only the details above. If you do not know something, say you will check with ' +
    CONFIG.MENTOR_NAME + '. Never invent requirements. Never share passwords or credentials.';

  return askClaude_(system, history, 500);
}

/* ---------- 4. Post-meeting email ---------- */

function sendPostMeetingEmails() {
  const d = readRows_('Interns');
  const now = new Date();
  let count = 0;

  d.rows.forEach(function (r) {
    if (r.Status !== 'Invited' || !r.MeetingTime) return;
    const meetingEnd = new Date(new Date(r.MeetingTime).getTime() + CONFIG.MEETING_MINUTES * 60000);
    if (meetingEnd > now) return; // meeting not over yet

    const body =
      'Hi ' + r.Name + ',\n\n' +
      'Thanks for attending the task briefing. Here is everything you need.\n\n' +
      'Task: ' + r.Task + '\n' +
      'Task link: ' + r.TaskURL + '\n' +
      'Closure date: ' + fmtDate_(r.ClosureDate) + '\n\n' +
      'Access: ' + CONFIG.ACCESS_NOTE + '\n\n' +
      'If anything is unclear, reply to this email or ask the task assistant.\n\n' +
      'Best of luck,\n' + CONFIG.MENTOR_NAME + '\n' + CONFIG.COMPANY;

    MailApp.sendEmail({
      to: r.Email,
      subject: 'Your task details: ' + r.Task,
      body: body
    });

    setCell_(d.sh, d.headers, r._row, 'Status', 'Email Sent');
    count++;
  });

  SpreadsheetApp.getActive().toast(count + ' email(s) sent.', 'Intern Bot');
}

function setupTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'sendPostMeetingEmails') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendPostMeetingEmails').timeBased().everyMinutes(15).create();
  SpreadsheetApp.getActive().toast('Post-meeting emails will now send automatically (checked every 15 min).', 'Intern Bot');
}

/* ---------- 5 & 6. Warnings with approval ---------- */

function draftWarnings() {
  const w = readRows_('Warnings');
  let count = 0;

  w.rows.forEach(function (r) {
    if (!r.Name || !r.Mistake || r.Draft) return;

    const system =
      'You write short, polite first-notice emails from an internship mentor. ' +
      'Be constructive, not harsh. State what went wrong, what to fix, and ask them to reply if they need help. ' +
      'Under 120 words. Plain text, no markdown. Sign as "' + CONFIG.MENTOR_NAME + '".';
    const draft = askClaude_(system, [{
      role: 'user',
      content: 'Intern name: ' + r.Name + '\nIssue: ' + r.Mistake + '\nWrite the email body only.'
    }], 400);

    setCell_(w.sh, w.headers, r._row, 'Draft', draft);
    count++;
  });

  SpreadsheetApp.getActive().toast(count + ' draft(s) created. Review them, then tick "Approved".', 'Intern Bot');
}

function sendApprovedWarnings() {
  const w = readRows_('Warnings');
  const interns = readRows_('Interns').rows;
  let count = 0;

  w.rows.forEach(function (r) {
    if (r.Approved !== true || !r.Draft || r.Sent) return;

    const intern = interns.filter(function (i) { return i.Name === r.Name; })[0];
    if (!intern) return;

    MailApp.sendEmail({
      to: intern.Email,
      subject: 'Notice regarding your task: ' + intern.Task,
      body: r.Draft
    });

    setCell_(w.sh, w.headers, r._row, 'Sent', new Date());
    count++;
  });

  SpreadsheetApp.getActive().toast(count + ' warning(s) sent.', 'Intern Bot');
}
