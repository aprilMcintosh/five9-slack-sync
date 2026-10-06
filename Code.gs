/**
 * Five9 -> Slack
 *
 *   syncFive9ToSlack()   every minute   mirrors agent state into Slack statuses
 *   refreshHoursCache()  every 15 min   caches talk/wait time from the Five9 report
 *   updateLeaderboard()  every 30 min   posts the leaderboard and locations boards
 *   doPost()             on demand      answers the /myhours and /myrank slash commands
 *
 * Setup order for /myhours is at the bottom of this file, under HOW TO DEPLOY.
 */

// ============================================================
// CONFIG
// ============================================================
const CFG = {
  FIVE9_HOST: 'https://api.five9.com',
  FIVE9_PATH: '/wssupervisor/v13/SupervisorWebService',
  FIVE9_ADMIN_PATH: '/wsadmin/v13/AdminWebService',

  TIMEZONE: 'America/Chicago',
  TIMEZONE_OFFSET_MS: -5 * 60 * 60 * 1000,  // -5 during CDT
  SHIFT_START_MS: 8 * 60 * 60 * 1000,       // 8:00 AM

  // Saved Five9 report used for hours. Must be in a SHARED folder.
  REPORT_FOLDER: 'Shared Reports',
  REPORT_NAME:   'Agent Login Stats',

  // Slack status buckets.
  BUCKETS: {
    'Working':    { emoji: ':fire:',                   text: 'Working' },
    'ACW':        { emoji: ':hourglass_flowing_sand:', text: 'Wrapping up' },
    'Not active': { emoji: ':chair:',                  text: 'Not active' }
  },
  DEFAULT_BUCKET: 'Not active',
  MISSING_BUCKET: 'Not active',
  FORCE_RESYNC_EVERY_MIN: 15,

  // Working window. ENABLED false = run 24/7.
  HOURS: { ENABLED: true, START: 7, END: 21, DAYS: [1, 2, 3, 4, 5] }, // 1=Mon ... 7=Sun

  // Pay week for /myhours. 0 = Sunday, 1 = Monday, etc.
  WEEK_STARTS_ON: 0,

  // FALLBACK ONLY. The real secret lives in Script Properties so that pasting
  // a new version of this file can't wipe it. Set it with saveWebSecret().
  WEB_SECRET: '',

  // ---- Leaderboard ----
  BOOKING_SHEET_ID: 'YOUR_BOOKING_SHEET_ID',
  BOOKING_ANCHOR_TAB: 'ALL Bookings',  // this week's tab is the one right after it
  BOOKING_CALLER_COL: 10,              // column J
  LEADERBOARD_CHANNEL: '#leaderboard',
  LEADERBOARD_TOP_N: 10,

  // ---- Active locations ----
  // The same sheet WordPress reads, so the board can never disagree with the form.
  // A = location, B = date, C = times, E = venue + address (first row of each location).
  WEBSITE_SHEET_ID: 'YOUR_WEBSITE_SHEET_ID',
  WEBSITE_TAB: '',                      // blank = use the first tab
  LOCATIONS_CHANNEL: '#locations',

  SHEET_AGENTS:   'Agents',
  SHEET_STATEMAP: 'StateMap',
  SHEET_HOURS:    'Hours',
  SHEET_NAMEMAP:  'NameMap',
  SHEET_DAYSTATE: 'DayState',
  SHEET_BOARD:    'Board'
};

// ============================================================
// SETUP / ADMIN
// ============================================================

/** Creates all the tabs. Safe to re-run. */
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss.getSheetByName(CFG.SHEET_AGENTS)) {
    const a = ss.insertSheet(CFG.SHEET_AGENTS);
    a.appendRow(['Five9 Username', 'Slack Member ID', 'Name (notes)', 'Last Bucket', 'Last Updated']);
    a.setFrozenRows(1);
  }
  if (!ss.getSheetByName(CFG.SHEET_STATEMAP)) {
    const s = ss.insertSheet(CFG.SHEET_STATEMAP);
    s.appendRow(['Five9 State', 'Bucket']);
    s.setFrozenRows(1);
  }
  if (!ss.getSheetByName(CFG.SHEET_HOURS)) {
    const h = ss.insertSheet(CFG.SHEET_HOURS);
    // Stored as plain seconds. Google Sheets turns "0:03:40" into an 1899 date
    // if you store it as text, which then lands in Slack looking insane.
    h.appendRow(['Five9 Username', 'Name',
                 'Today Talk (s)', 'Today Wait (s)', 'Today Paid (s)',
                 'Week Talk (s)', 'Week Wait (s)', 'Week Paid (s)',
                 'Week Not Ready (s)', 'Week ACW (s)', 'Updated']);
    h.setFrozenRows(1);
  }
  if (!ss.getSheetByName(CFG.SHEET_NAMEMAP)) {
    const n = ss.insertSheet(CFG.SHEET_NAMEMAP);
    n.appendRow(['Caller Name as typed', 'Show as  (leave blank to hide from board)']);
    n.setFrozenRows(1);
    n.appendRow(['awef', '']);   // the test row in the booking sheet
  }

  if (!ss.getSheetByName(CFG.SHEET_DAYSTATE)) {
    const d = ss.insertSheet(CFG.SHEET_DAYSTATE);
    d.appendRow(['Caller', 'Count at start of day', 'Date', 'Week tab']);
    d.setFrozenRows(1);
  }

  if (!ss.getSheetByName(CFG.SHEET_BOARD)) {
    const b = ss.insertSheet(CFG.SHEET_BOARD);
    b.appendRow(['Rank', 'Caller', 'This week', 'Today', 'Week tab', 'Updated']);
    b.setFrozenRows(1);
  }

  Logger.log('Tabs ready.');
}

/**
 * !!! DO NOT RUN unless you are deliberately changing credentials. !!!
 * Real values already live in Script Properties.
 */
function saveSecrets() {
  PropertiesService.getScriptProperties().setProperties({
    FIVE9_USER: 'your-five9-username',
    FIVE9_PASS: 'the-five9-password',
    SLACK_TOKEN: 'xoxp-your-user-token-here'
  });
  Logger.log('Saved. Now wipe the values out of this function and save the file again.');
}

/** Run once. Stores the /myhours URL secret where file updates can't wipe it. */
function saveWebSecret() {
  const SECRET = 'put-your-secret-here';   // must match ?k=... in the Slack command URL
  PropertiesService.getScriptProperties().setProperty('WEB_SECRET', SECRET);
  Logger.log('WEB_SECRET saved. Wipe the value above and save the file again.');
}

/** Run once. Stores the Slack BOT token (xoxb-) used to post the leaderboard. */
function saveBotToken() {
  const TOKEN = 'xoxb-put-the-bot-token-here';
  PropertiesService.getScriptProperties().setProperty('SLACK_BOT_TOKEN', TOKEN);
  Logger.log('Bot token saved. Wipe the value above and save the file again.');
}

/**
 * Installs all three scheduled jobs. Run once.
 * everyMinutes() only accepts 1, 5, 10, 15 or 30. Nothing else is valid.
 */
function installTriggers() {
  removeTriggers();
  ScriptApp.newTrigger('syncFive9ToSlack').timeBased().everyMinutes(1).create();
  ScriptApp.newTrigger('refreshHoursCache').timeBased().everyMinutes(15).create();
  ScriptApp.newTrigger('updateLeaderboard').timeBased().everyMinutes(30).create();
  Logger.log('Triggers installed: status sync 1 min, hours 15 min, leaderboard 30 min.');
}

function removeTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) { ScriptApp.deleteTrigger(t); });
}

// ============================================================
// HOURS CACHE
// ============================================================

/**
 * Pulls the Five9 report for the current pay week and writes per-agent totals
 * into the Hours tab. Runs on a timer so /myhours can answer instantly.
 */
function refreshHoursCache() {
  if (!withinHours_()) return;   // cached values stay readable outside these hours

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return;

  try {
    const now = new Date();
    const weekStart = weekStart_(now);
    const report = five9RunReport_(CFG.REPORT_FOLDER, CFG.REPORT_NAME, weekStart, now);
    if (!report) { Logger.log('Hours: report returned nothing.'); return; }

    const todayStr = Utilities.formatDate(now, CFG.TIMEZONE, 'yyyy/MM/dd');

    // username -> totals in seconds
    const acc = {};
    report.records.forEach(function (r) {
      const user = String(r['AGENT'] || '').trim();
      if (!user) return;

      if (!acc[user]) {
        acc[user] = {
          name: [r['AGENT FIRST NAME'], r['AGENT LAST NAME']].filter(String).join(' ').trim(),
          todayTalk: 0, todayWait: 0,
          weekTalk: 0, weekWait: 0, weekNotReady: 0, weekAcw: 0
        };
      }

      const talk     = hhmmssToSec_(r['TALK TIME']);
      const wait     = hhmmssToSec_(r['WAIT TIME']);
      const notReady = hhmmssToSec_(r['NOT READY TIME']);
      const acw      = hhmmssToSec_(r['ON ACW TIME']);

      acc[user].weekTalk     += talk;
      acc[user].weekWait     += wait;
      acc[user].weekNotReady += notReady;
      acc[user].weekAcw      += acw;

      // The report's DATE column follows the report's own timezone setting.
      // If you ever change that report to Central this stays correct either way.
      if (String(r['DATE'] || '').trim() === todayStr) {
        acc[user].todayTalk += talk;
        acc[user].todayWait += wait;
      }
    });

    // Numbers only. Never write "H:MM:SS" strings into a sheet.
    const rows = Object.keys(acc).sort().map(function (user) {
      const a = acc[user];
      return [
        user,
        a.name,
        a.todayTalk,
        a.todayWait,
        a.todayTalk + a.todayWait,
        a.weekTalk,
        a.weekWait,
        a.weekTalk + a.weekWait,
        a.weekNotReady,
        a.weekAcw,
        now
      ];
    });

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.SHEET_HOURS);

    // clear() not clearContent(): old cells carry a time-of-day number format,
    // and Sheets hands those back as Date objects no matter what you write.
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, 11).clear();
    }

    if (rows.length) {
      sheet.getRange(2, 1, rows.length, 11).setValues(rows);
      sheet.getRange(2, 3, rows.length, 8).setNumberFormat('0');                  // seconds
      sheet.getRange(2, 11, rows.length, 1).setNumberFormat('yyyy-mm-dd hh:mm');  // updated
    }

    Logger.log('Hours cache: ' + rows.length + ' agents, week from ' +
               Utilities.formatDate(weekStart, CFG.TIMEZONE, 'EEE MMM d'));

  } catch (e) {
    Logger.log('HOURS ERROR: ' + e);
  } finally {
    lock.releaseLock();
  }
}

/** Reads one agent's cached hours. Returns null if they aren't in the cache. */
function readCachedHours_(five9User) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.SHEET_HOURS);
  const last = sheet.getLastRow();
  if (last < 2) return null;

  const want = String(five9User).trim().toLowerCase();
  const rows = sheet.getRange(2, 1, last - 1, 11).getValues();
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).trim().toLowerCase() === want) {
      const r = rows[i];
      return {
        user: r[0], name: r[1],
        todayTalk: toSec_(r[2]), todayWait: toSec_(r[3]), todayPaid: toSec_(r[4]),
        weekTalk: toSec_(r[5]), weekWait: toSec_(r[6]), weekPaid: toSec_(r[7]),
        weekNotReady: toSec_(r[8]), weekAcw: toSec_(r[9]),
        updated: r[10]
      };
    }
  }
  return null;
}

/**
 * Coerce whatever is in a cell into seconds.
 * Handles plain numbers (current format), Date objects (Sheets' 1899 time
 * coercion from the old string format), and "H:MM:SS" text.
 */
function toSec_(v) {
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return Math.max(0, Math.round(v));
  if (v instanceof Date) return v.getHours() * 3600 + v.getMinutes() * 60 + v.getSeconds();
  return hhmmssToSec_(v);
}

// ============================================================
// SLACK SLASH COMMAND  ( /myhours )
// ============================================================

function doPost(e) {
  try {
    if (!e || !e.parameter) return jsonOut_({ text: 'No payload received.' });

    const secret = PropertiesService.getScriptProperties().getProperty('WEB_SECRET') || CFG.WEB_SECRET;
    if (secret && e.parameter.k !== secret) {
      return jsonOut_({ response_type: 'ephemeral', text: 'Not authorized.' });
    }

    const slackId = String(e.parameter.user_id || '').trim();
    if (!slackId) return jsonOut_({ response_type: 'ephemeral', text: 'No Slack user on that request.' });

    // /myrank -> personal leaderboard position
    if (String(e.parameter.command || '').indexOf('rank') !== -1) {
      return jsonOut_({ response_type: 'ephemeral', blocks: myRankBlocks_(slackId), text: 'Your rank' });
    }

    const five9User = five9UserForSlackId_(slackId);
    if (!five9User) {
      return jsonOut_({
        response_type: 'ephemeral',
        text: "I don't have you linked to a Five9 account yet. Ask your admin to add you to the Agents sheet."
      });
    }

    const h = readCachedHours_(five9User);
    if (!h) {
      return jsonOut_({
        response_type: 'ephemeral',
        text: "No hours recorded for you yet this week. If you've been logged in today, give it a few minutes and try again."
      });
    }

    return jsonOut_({ response_type: 'ephemeral', blocks: hoursBlocks_(h), text: 'Your hours' });

  } catch (err) {
    return jsonOut_({ response_type: 'ephemeral', text: 'Something broke on my end: ' + err });
  }
}

/** Visit the web app URL in a browser to confirm it deployed. */
function doGet(e) {
  return ContentService.createTextOutput(
    'Five9 hours endpoint is alive. Slack posts here; there is nothing to see in a browser.'
  );
}

function hoursBlocks_(h) {
  const f = secToHhmmss_;
  const weekLabel = Utilities.formatDate(weekStart_(new Date()), CFG.TIMEZONE, 'MMM d') +
                    ' through today';
  const stamp = h.updated instanceof Date
    ? Utilities.formatDate(h.updated, CFG.TIMEZONE, 'h:mm a')
    : '';

  return [
    { type: 'section', text: { type: 'mrkdwn',
      text: '*This week*   `' + f(h.weekPaid) + '`\n' +
            'talk ' + f(h.weekTalk) + '  +  wait ' + f(h.weekWait) } },
    { type: 'section', text: { type: 'mrkdwn',
      text: '*Today*   `' + f(h.todayPaid) + '`\n' +
            'talk ' + f(h.todayTalk) + '  +  wait ' + f(h.todayWait) } },
    { type: 'context', elements: [{ type: 'mrkdwn',
      text: 'Not paid: not ready ' + f(h.weekNotReady) + '  ·  wrap-up ' + f(h.weekAcw) +
            '\n' + weekLabel + (stamp ? '  ·  updated ' + stamp : '') }] }
  ];
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// LEADERBOARD
// ============================================================

/**
 * Reads this week's booking tab and returns the ranking.
 * "This week" = the tab immediately after the ALL Bookings tab.
 */
function readWeekBookings_(readOnly) {
  const ss = SpreadsheetApp.openById(CFG.BOOKING_SHEET_ID);
  const tabs = ss.getSheets();

  let idx = -1;
  const anchor = CFG.BOOKING_ANCHOR_TAB.trim().toLowerCase();
  for (let i = 0; i < tabs.length; i++) {
    if (tabs[i].getName().trim().toLowerCase() === anchor) { idx = i; break; }
  }
  if (idx === -1 || idx + 1 >= tabs.length) {
    Logger.log('Leaderboard: could not find the tab after "' + CFG.BOOKING_ANCHOR_TAB + '".');
    return null;
  }

  const week = tabs[idx + 1];
  const last = week.getLastRow();
  if (last < 2) return { tabName: week.getName(), ranked: [], total: 0 };

  const nameMap = loadNameMap_();
  const counts = {}, display = {};

  week.getRange(2, CFG.BOOKING_CALLER_COL, last - 1, 1).getValues().forEach(function (r) {
    const raw = String(r[0] || '').trim().replace(/\s+/g, ' ');
    if (!raw) return;
    const key = raw.toLowerCase();

    if (nameMap.hasOwnProperty(key)) {
      if (!nameMap[key]) return;              // blank in NameMap = hide from the board
      const shown = nameMap[key];
      const mkey = shown.toLowerCase();
      counts[mkey] = (counts[mkey] || 0) + 1;
      display[mkey] = shown;
      return;
    }

    counts[key] = (counts[key] || 0) + 1;
    if (!display[key]) display[key] = titleCase_(raw);
  });

  const ranked = Object.keys(counts)
    .map(function (k) { return { name: display[k], count: counts[k] }; })
    .sort(function (a, b) {
      if (b.count !== a.count) return b.count - a.count;
      return a.name.localeCompare(b.name);
    });

  let total = 0;
  ranked.forEach(function (r) { total += r.count; });

  const todayTotal = applyDayDeltas_(ranked, week.getName(), readOnly);

  return { tabName: week.getName(), ranked: ranked, total: total, todayTotal: todayTotal };
}

/**
 * The booking sheet has no submission timestamp, so "today" is derived:
 * snapshot everyone's count at the first run of each day, then today's
 * number is current minus that snapshot.
 *
 * Annotates each ranked entry with .today and returns the day's total.
 * Read-only when called from /myrank so a command can't corrupt the snapshot.
 */
function applyDayDeltas_(ranked, tabName, readOnly) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CFG.SHEET_DAYSTATE);
  if (!sheet) { ranked.forEach(function (r) { r.today = null; }); return null; }

  const todayStr = Utilities.formatDate(new Date(), CFG.TIMEZONE, 'yyyy-MM-dd');

  let storedDate = '', storedTab = '';
  const baseline = {};
  if (sheet.getLastRow() > 1) {
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
    rows.forEach(function (r) {
      const name = String(r[0] || '').trim().toLowerCase();
      if (name) baseline[name] = Number(r[1]) || 0;
      if (!storedDate) { storedDate = String(r[2] || '').trim(); storedTab = String(r[3] || '').trim(); }
    });
  }

  const staleDay = storedDate !== todayStr;
  const newWeek  = storedTab !== tabName;

  if ((staleDay || newWeek) && !readOnly) {
    // First run of a new day (or a new week tab): today's baseline is right now.
    const rows = ranked.map(function (r) { return [r.name, r.count, todayStr, tabName]; });
    if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).clear();
    if (rows.length) sheet.getRange(2, 1, rows.length, 4).setValues(rows);
    ranked.forEach(function (r) { r.today = 0; });
    Logger.log('DayState reset for ' + todayStr + ' on tab "' + tabName + '".');
    return 0;
  }

  if (staleDay || newWeek) { ranked.forEach(function (r) { r.today = null; }); return null; }

  let dayTotal = 0;
  ranked.forEach(function (r) {
    const was = baseline.hasOwnProperty(r.name.toLowerCase()) ? baseline[r.name.toLowerCase()] : 0;
    r.today = Math.max(0, r.count - was);
    dayTotal += r.today;
  });
  return dayTotal;
}

/** Scheduled every 30 min. Posts the board, or edits the existing one in place. */
function updateLeaderboard() {
  if (!withinHours_()) return;

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return;

  try {
    // Two independent boards. Each gets its own try so a bad week tab can't
    // take the locations board down with it, or the other way round.
    try {
      const data = readWeekBookings_();
      if (data) {
        writeBoardCache_(data);   // so /myrank answers inside Slack's 3-second limit
        slackPostOrUpdate_(CFG.LEADERBOARD_CHANNEL, leaderboardBlocks_(data),
                           'Leaderboard', 'LEADERBOARD_TS');
        Logger.log('Leaderboard: ' + data.total + ' bookings, ' + data.ranked.length + ' callers.');
      }
    } catch (e) {
      Logger.log('LEADERBOARD ERROR: ' + e);
    }

    try {
      updateLocationsBoard();
    } catch (e) {
      Logger.log('LOCATIONS ERROR: ' + e);
    }
  } finally {
    lock.releaseLock();
  }
}

function leaderboardBlocks_(data) {
  const medal = ['🥇', '🥈', '🥉'];
  const top = data.ranked.slice(0, CFG.LEADERBOARD_TOP_N);

  const lines = top.map(function (r, i) {
    const badge = i < 3 ? medal[i] : '`' + pad2_(i + 1) + '`';
    const today = (r.today) ? '   _+' + r.today + ' today_' : '';
    return badge + '  *' + r.name + '*  —  ' + r.count + today;
  });

  if (!lines.length) lines.push('_No bookings yet this week. Go get one._');

  const subtitle = data.tabName + '  ·  ' + data.total + ' this week' +
                   (data.todayTotal === null ? '' : '  ·  ' + data.todayTotal + ' today');

  return [
    { type: 'header', text: { type: 'plain_text', text: '🏆  Leaderboard', emoji: true } },
    { type: 'context', elements: [{ type: 'mrkdwn', text: subtitle }] },
    { type: 'section', text: { type: 'mrkdwn', text: lines.join('\n') } },
    { type: 'context', elements: [{ type: 'mrkdwn',
      text: 'Not on the board? Type `/myrank` to see where you stand  ·  updated ' +
            Utilities.formatDate(new Date(), CFG.TIMEZONE, 'h:mm a') }] }
  ];
}

/** Writes the ranking into the Board tab so /myrank never has to open the booking sheet. */
function writeBoardCache_(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.SHEET_BOARD);
  if (!sheet) return;

  const now = new Date();
  const rows = data.ranked.map(function (r, i) {
    return [i + 1, r.name, r.count, (r.today === null || r.today === undefined) ? '' : r.today,
            data.tabName, now];
  });

  if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).clear();
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, 6).setValues(rows);
    sheet.getRange(2, 6, rows.length, 1).setNumberFormat('yyyy-mm-dd hh:mm');
  }
}

/** Reads the Board tab. Fast, because it's in this spreadsheet. */
function readBoardCache_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.SHEET_BOARD);
  if (!sheet || sheet.getLastRow() < 2) return null;

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
  const ranked = rows.map(function (r) {
    return {
      rank: Number(r[0]) || 0,
      name: String(r[1] || ''),
      count: Number(r[2]) || 0,
      today: (r[3] === '' || r[3] === null) ? null : (Number(r[3]) || 0)
    };
  }).filter(function (r) { return r.name; });

  if (!ranked.length) return null;

  let total = 0;
  ranked.forEach(function (r) { total += r.count; });

  return { ranked: ranked, total: total, tabName: String(rows[0][4] || ''), updated: rows[0][5] };
}

/** Personal rank, ephemeral, for the /myrank command. Reads the cache only. */
function myRankBlocks_(slackId) {
  const data = readBoardCache_();
  if (!data) {
    return [{ type: 'section', text: { type: 'mrkdwn',
      text: "The board hasn't been built yet. It refreshes every half hour, try again shortly." } }];
  }

  const me = leaderboardNameForSlackId_(slackId);
  if (!me) {
    return [{ type: 'section', text: { type: 'mrkdwn',
      text: "I don't know which caller name is yours yet. Ask your admin to add it to the Agents sheet." } }];
  }

  let rank = -1, count = 0, today = null;
  for (let i = 0; i < data.ranked.length; i++) {
    if (data.ranked[i].name.toLowerCase() === me.toLowerCase()) {
      rank = i + 1; count = data.ranked[i].count; today = data.ranked[i].today; break;
    }
  }

  if (rank === -1) {
    return [{ type: 'section', text: { type: 'mrkdwn',
      text: '*' + me + '*\nNo bookings yet this week. ' + data.total + ' on the board so far.' } }];
  }

  const ahead = rank > 1 ? data.ranked[rank - 2] : null;
  const lines = [
    '*' + me + '*',
    'Rank *' + rank + '* of ' + data.ranked.length + '  ·  *' + count + '* this week' +
      (today === null ? '' : '  ·  *' + today + '* today')
  ];
  if (ahead) {
    lines.push('_' + ahead.name + ' is ahead of you with ' + ahead.count +
               ' (' + (ahead.count - count) + ' to catch)_');
  } else {
    lines.push('_You are in first. Keep it._');
  }

  return [
    { type: 'section', text: { type: 'mrkdwn', text: lines.join('\n') } },
    { type: 'context', elements: [{ type: 'mrkdwn', text: data.tabName }] }
  ];
}

/** Agents tab column C holds the booking-sheet caller name, if set. */
function leaderboardNameForSlackId_(slackId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.SHEET_AGENTS);
  const last = sheet.getLastRow();
  if (last < 2) return null;
  const rows = sheet.getRange(2, 1, last - 1, 3).getValues();
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][1]).trim() === slackId) {
      const c = String(rows[i][2] || '').trim();
      return c || null;
    }
  }
  return null;
}

function loadNameMap_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.SHEET_NAMEMAP);
  const map = {};
  if (!sheet || sheet.getLastRow() < 2) return map;
  sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues().forEach(function (r) {
    const from = String(r[0] || '').trim().toLowerCase().replace(/\s+/g, ' ');
    if (!from) return;
    map[from] = String(r[1] || '').trim();   // '' means hide
  });
  return map;
}

function titleCase_(s) {
  return String(s).split(' ').map(function (w) {
    return w ? w.charAt(0).toUpperCase() + w.slice(1).toLowerCase() : w;
  }).join(' ');
}

function pad2_(n) { return n < 10 ? ' ' + n : String(n); }


// ============================================================
// ACTIVE LOCATIONS
// ============================================================
//
// Source: the same Website sheet WordPress reads, so the board in Slack and
// the booking form can never disagree. Layout of the first tab:
//
//   A = Locations   "Austin, TX"
//   B = Date        "Wednesday, September 23rd"
//   C = Times       "2:45 PM, 6:45 PM"
//   D = (empty)
//   E = venue + address, on the FIRST row of each location only.
//       Line 1 is the venue, the remaining lines are the street address.
//
// Two rows per location, one per show day. A row only counts as a show row if
// A, B and C are all filled, which is what keeps the header row and any other
// block on the sheet from being read as a location.

/** Reads the Website sheet and groups the rows by location. */
function readLocations_() {
  let ss;
  try {
    ss = SpreadsheetApp.openById(CFG.WEBSITE_SHEET_ID);
  } catch (e) {
    Logger.log('Locations: cannot open the Website sheet. ' + e);
    return null;
  }

  const tab = CFG.WEBSITE_TAB
    ? ss.getSheetByName(CFG.WEBSITE_TAB)
    : ss.getSheets()[0];
  if (!tab) { Logger.log('Locations: tab not found.'); return null; }

  const last = tab.getLastRow();
  if (last < 2) return [];

  const rows = tab.getRange(1, 1, last, 5).getValues();

  const order = [];
  const byName = {};

  rows.forEach(function (r) {
    const name  = String(r[0] || '').trim().replace(/\s+/g, ' ');
    const date  = String(r[1] || '').trim();
    const times = String(r[2] || '').trim();
    const addr  = String(r[4] || '').trim();

    // A show row has all three. Anything else is a header or another block.
    if (!name || !date || !times) return;
    if (name.toLowerCase() === 'locations') return;

    const key = name.toLowerCase();
    if (!byName[key]) {
      byName[key] = { name: name, venue: '', address: '', days: [] };
      order.push(key);
    }

    // The address sits on the first row of the location, but take it from
    // whichever row actually carries it rather than assuming the first.
    if (addr && !byName[key].venue) {
      const parts = addr.split('\n')
        .map(function (s) { return s.trim(); })
        .filter(function (s) { return s; });
      byName[key].venue   = parts.shift() || '';
      byName[key].address = parts.join(', ').replace(/,\s*,/g, ',');
    }

    const d = shortDate_(date);
    byName[key].days.push({ dow: d.dow, md: d.md, times: times, isToday: isToday_(date) });
  });

  return order.map(function (k) { return byName[k]; });
}

const MONTH_NUM_ = {
  jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6,
  jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12
};

/**
 * "Wednesday, September 23rd" -> { dow: 'Wed', md: '9/23' }.
 * Anything it can't parse falls back to the raw text in md, so a weird cell
 * shows up on the board looking odd instead of vanishing silently.
 */
function shortDate_(s) {
  const raw = String(s).trim();
  const m = raw.match(/^\s*([A-Za-z]+)[,\s]+([A-Za-z]+)\s+(\d{1,2})/);
  if (!m) return { dow: '', md: raw };
  const mon = MONTH_NUM_[m[2].slice(0, 3).toLowerCase()];
  return {
    dow: m[1].slice(0, 3),
    md: mon ? (mon + '/' + Number(m[3])) : (m[2].slice(0, 3) + ' ' + m[3])
  };
}

/** Right-pads with spaces so the monospace columns line up. */
function padEnd_(s, n) {
  let out = String(s);
  while (out.length < n) out += ' ';
  return out;
}

/** True if the text names today's month and day, in the call center's timezone. */
function isToday_(s) {
  const m = String(s).match(/([A-Za-z]+)\s+(\d{1,2})/);
  if (!m) return false;
  const now = new Date();
  const month = Utilities.formatDate(now, CFG.TIMEZONE, 'MMMM').toLowerCase();
  const day   = Number(Utilities.formatDate(now, CFG.TIMEZONE, 'd'));
  return m[1].toLowerCase() === month && Number(m[2]) === day;
}

function locationsBlocks_(locs) {
  if (!locs || !locs.length) {
    return [
      { type: 'header', text: { type: 'plain_text', text: '📍  Active locations', emoji: true } },
      { type: 'section', text: { type: 'mrkdwn',
        text: "_Nothing on the Website sheet right now._" } }
    ];
  }

  // --- Column widths, measured off the real data so nothing ever wraps. ---
  let wCity = 'CITY'.length, wDay = 'DAY'.length, wTimes = 'TIMES'.length;
  locs.forEach(function (L) {
    if (L.name.length > wCity) wCity = L.name.length;
    L.days.forEach(function (d) {
      const day = (d.dow ? d.dow + ' ' : '') + d.md;
      if (day.length > wDay) wDay = day.length;
      if (d.times.length > wTimes) wTimes = d.times.length;
    });
  });

  const GAP = '  ';
  const width = 2 + wCity + GAP.length + wDay + GAP.length + wTimes;

  const table = [];
  table.push('  ' + padEnd_('CITY', wCity) + GAP + padEnd_('DAY', wDay) + GAP + 'TIMES');
  table.push(new Array(width + 1).join('-'));

  let anyToday = false;

  locs.forEach(function (L, i) {
    if (i > 0) table.push('');
    L.days.forEach(function (d, j) {
      if (d.isToday) anyToday = true;
      const mark = d.isToday ? '> ' : '  ';
      const city = j === 0 ? L.name : '';
      const day  = (d.dow ? d.dow + ' ' : '') + d.md;
      table.push(mark + padEnd_(city, wCity) + GAP + padEnd_(day, wDay) + GAP + d.times);
    });
  });

  const blocks = [
    { type: 'header', text: { type: 'plain_text', text: '📍  Active locations', emoji: true } },
    { type: 'section', text: { type: 'mrkdwn', text: '```\n' + table.join('\n') + '\n```' } }
  ];

  if (anyToday) {
    blocks.push({ type: 'context', elements: [{ type: 'mrkdwn', text: '`>` marks today' }] });
  }

  // --- Venues underneath, so the table stays narrow but nobody loses the address. ---
  const venueLines = [];
  locs.forEach(function (L) {
    if (!L.venue && !L.address) return;
    const parts = ['*' + L.name + '*'];
    if (L.venue) parts.push(L.venue);
    if (L.address) parts.push(mapsLink_(L.venue, L.address));
    venueLines.push(parts.join('\n'));
  });

  if (venueLines.length) {
    blocks.push({ type: 'divider' });
    blocks.push({ type: 'section', text: { type: 'mrkdwn',
      text: '*Venues*\n\n' + venueLines.join('\n\n') } });
  }

  blocks.push({ type: 'context', elements: [{ type: 'mrkdwn',
    text: 'From the booking sheet  ·  updated ' +
          Utilities.formatDate(new Date(), CFG.TIMEZONE, 'h:mm a') }] });

  return blocks;
}

/** Turns an address into a tappable Slack link that opens Google Maps. */
function mapsLink_(venue, address) {
  const q = encodeURIComponent(((venue ? venue + ' ' : '') + address).trim());
  return '<https://www.google.com/maps/search/?api=1&query=' + q + '|' + address + '>';
}

/**
 * Posts the locations board, or edits the existing one in place.
 * Called from updateLeaderboard() so it rides the same 30-minute trigger.
 */
function updateLocationsBoard() {
  const locs = readLocations_();
  if (locs === null) return;

  slackPostOrUpdate_(CFG.LOCATIONS_CHANNEL, locationsBlocks_(locs),
                     'Active locations', 'LOCATIONS_TS');
  Logger.log('Locations: ' + locs.length + ' cities posted.');
}

/**
 * Deletes the current locations message and forgets it, so the next run posts a
 * clean one. Needed once, because unfurl cards stay stuck to a message that was
 * already posted with unfurling on. Editing it does not remove them.
 */
function resetLocationsBoard() {
  slackDeleteRemembered_(CFG.LOCATIONS_CHANNEL, 'LOCATIONS_TS');
  updateLocationsBoard();
}

/** Deletes the message remembered under propKey, then clears the property. */
function slackDeleteRemembered_(channel, propKey) {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('SLACK_BOT_TOKEN');
  const ts = props.getProperty(propKey);
  if (!token || !ts) { Logger.log('Nothing to delete for ' + propKey + '.'); return; }

  const resp = UrlFetchApp.fetch('https://slack.com/api/chat.delete', {
    method: 'post',
    contentType: 'application/json; charset=utf-8',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify({ channel: channel, ts: ts }),
    muteHttpExceptions: true
  });

  let out = {};
  try { out = JSON.parse(resp.getContentText()); } catch (e) {}
  Logger.log('chat.delete ' + propKey + ': ' + (out.ok ? 'deleted' : out.error));

  // Forget it either way. If the delete failed the old message is orphaned, but
  // keeping the ts would mean editing a message we can no longer control.
  props.deleteProperty(propKey);
}

/** Prints the board to the log without touching Slack. */
function debugLocations() {
  const locs = readLocations_();
  if (locs === null) { Logger.log('Could not read the Website sheet.'); return; }
  if (!locs.length) { Logger.log('No location rows found.'); return; }
  // Print the real Slack text, not a summary, so what you read in the log is
  // what the channel will get.
  locationsBlocks_(locs).forEach(function (b) {
    if (b.type === 'header')  Logger.log(b.text.text);
    if (b.type === 'section') Logger.log(b.text.text);
    if (b.type === 'context') Logger.log(b.elements[0].text);
    if (b.type === 'divider') Logger.log('-----');
  });
}


/**
 * Posts a message the first time, then edits that same message forever after.
 * The message timestamp is remembered in Script Properties under propKey.
 */
function slackPostOrUpdate_(channel, blocks, fallbackText, propKey) {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('SLACK_BOT_TOKEN');
  if (!token) { Logger.log('SLACK_BOT_TOKEN missing. Run saveBotToken().'); return; }

  const call = function (method, payload) {
    const resp = UrlFetchApp.fetch('https://slack.com/api/' + method, {
      method: 'post',
      contentType: 'application/json; charset=utf-8',
      headers: { Authorization: 'Bearer ' + token },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    try { return JSON.parse(resp.getContentText()); }
    catch (e) { return { ok: false, error: 'bad_json' }; }
  };

  // unfurl_links/unfurl_media off, or Slack turns every Google Maps address into
  // a preview card and buries the board under them.
  const existing = props.getProperty(propKey);
  if (existing) {
    const upd = call('chat.update', {
      channel: channel, ts: existing, blocks: blocks, text: fallbackText,
      unfurl_links: false, unfurl_media: false
    });
    if (upd.ok) return;
    Logger.log('chat.update failed (' + upd.error + '), posting a fresh message.');
    props.deleteProperty(propKey);
  }

  const post = call('chat.postMessage', {
    channel: channel, blocks: blocks, text: fallbackText,
    unfurl_links: false, unfurl_media: false
  });
  if (post.ok && post.ts) {
    props.setProperty(propKey, post.ts);
    Logger.log('Posted a new message for ' + propKey + ', ts=' + post.ts);
  } else {
    Logger.log('chat.postMessage failed: ' + post.error);
  }
}

/** Prints the board to the log without touching Slack. */
function debugLeaderboard() {
  const data = readWeekBookings_();
  if (!data) return;
  Logger.log('Tab: ' + data.tabName + '   total: ' + data.total);
  data.ranked.slice(0, 15).forEach(function (r, i) {
    Logger.log((i + 1) + '. ' + r.name + '  ' + r.count);
  });
}

/** Agents tab: column B is the Slack member ID, column A is the Five9 username. */
function five9UserForSlackId_(slackId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.SHEET_AGENTS);
  const last = sheet.getLastRow();
  if (last < 2) return null;
  const rows = sheet.getRange(2, 1, last - 1, 2).getValues();
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][1]).trim() === slackId) return String(rows[i][0]).trim();
  }
  return null;
}

// ============================================================
// STATUS SYNC
// ============================================================

function syncFive9ToSlack() {
  if (!withinHours_()) return;

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const agentSheet = ss.getSheetByName(CFG.SHEET_AGENTS);
    const lastRow = agentSheet.getLastRow();
    if (lastRow < 2) { Logger.log('No agents in the Agents tab.'); return; }

    const agents = agentSheet.getRange(2, 1, lastRow - 1, 5).getValues();
    const stateMap = loadStateMap_();
    const live = five9GetAgentStates_();

    const liveBucket = {};
    live.rows.forEach(function (row) {
      const user = String(pickUsername_(row) || '').trim().toLowerCase();
      if (!user) return;
      const raw = String(pickState_(row) || '').trim();
      const mapped = stateMap[raw.toLowerCase()];
      liveBucket[user] = mapped || CFG.DEFAULT_BUCKET;
      if (!mapped && raw) {
        Logger.log('Unmapped Five9 state "' + raw + '" -> ' + CFG.DEFAULT_BUCKET);
      }
    });

    const now = new Date();
    const forceMs = CFG.FORCE_RESYNC_EVERY_MIN * 60 * 1000;
    let pushed = 0, failed = 0;

    for (let i = 0; i < agents.length; i++) {
      const five9User = String(agents[i][0] || '').trim().toLowerCase();
      const slackId   = String(agents[i][1] || '').trim();
      const lastB     = String(agents[i][3] || '').trim();
      const lastAt    = agents[i][4] instanceof Date ? agents[i][4].getTime() : 0;
      if (!five9User || !slackId) continue;

      const bucket = liveBucket.hasOwnProperty(five9User) ? liveBucket[five9User] : CFG.MISSING_BUCKET;
      const stale = forceMs > 0 && (now.getTime() - lastAt) > forceMs;
      if (bucket === lastB && !stale) continue;

      const look = CFG.BUCKETS[bucket] || CFG.BUCKETS[CFG.DEFAULT_BUCKET];
      if (slackSetStatus_(slackId, look.text, look.emoji)) {
        agents[i][3] = bucket;
        agents[i][4] = now;
        pushed++;
      } else {
        failed++;
      }
    }

    // Write ONLY columns D and E (Last Bucket, Last Updated).
    // Writing A-C back would clobber edits made to the sheet while this was
    // running, which is several seconds because of the Five9 round trip.
    agentSheet.getRange(2, 4, agents.length, 2).setValues(
      agents.map(function (r) { return [r[3], r[4]]; })
    );
    if (pushed || failed) Logger.log('Pushed ' + pushed + ', failed ' + failed);

  } catch (e) {
    Logger.log('SYNC ERROR: ' + e);
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// DIAGNOSTICS
// ============================================================

function debugDumpFive9() {
  const data = five9GetAgentStates_();
  Logger.log('COLUMNS: ' + JSON.stringify(data.columns));
  Logger.log('ROW COUNT: ' + data.rows.length);
  const seen = {};
  data.rows.forEach(function (r) { seen[pickState_(r)] = (seen[pickState_(r)] || 0) + 1; });
  Object.keys(seen).sort().forEach(function (s) { Logger.log('"' + s + '"  x' + seen[s]); });
}

function debugWhoIsOn() {
  const data = five9GetAgentStates_();
  const on = data.rows.filter(function (r) { return pickState_(r) !== 'Logged Out'; });
  Logger.log('Not logged out: ' + on.length);
  on.forEach(function (r) {
    Logger.log(pickUsername_(r) + '  |  ' + pickState_(r) + '  |  ' + (r['Full Name'] || ''));
  });
}

/** Runs the hours report and prints what it found, without touching the sheet. */
function debugHours() {
  const now = new Date();
  const start = weekStart_(now);
  Logger.log('Week: ' + five9Time_(start) + '  ->  ' + five9Time_(now));
  const report = five9RunReport_(CFG.REPORT_FOLDER, CFG.REPORT_NAME, start, now);
  if (!report) { Logger.log('Nothing came back.'); return; }
  Logger.log('HEADER: ' + JSON.stringify(report.header));
  Logger.log('RECORDS: ' + report.records.length);
  report.records.slice(0, 8).forEach(function (r) {
    Logger.log(r['AGENT'] + '  ' + r['DATE'] +
               '  talk ' + r['TALK TIME'] + '  wait ' + r['WAIT TIME'] +
               '  => paid ' + secToHhmmss_(hhmmssToSec_(r['TALK TIME']) + hhmmssToSec_(r['WAIT TIME'])));
  });
}

/** Pretends to be Slack calling /myhours. Put your own Slack member ID below. */
function debugMyHours() {
  const SLACK_ID = 'U00000000';   // <- change to a real member ID from the Agents tab
  const fake = { parameter: { k: CFG.WEB_SECRET, user_id: SLACK_ID } };
  Logger.log(doPost(fake).getContent());
}

/** Raw report XML, for when the parser looks wrong. */
function debugRunReport() {
  const now = new Date();
  const start = weekStart_(now);
  const xml = five9RunReportXml_(CFG.REPORT_FOLDER, CFG.REPORT_NAME, start, now);
  if (!xml) return;
  Logger.log('LENGTH: ' + xml.length);
  for (let i = 0; i < Math.min(xml.length, 6000); i += 1500) {
    Logger.log('CHUNK ' + (i / 1500) + ': ' + xml.substring(i, i + 1500));
  }
}

// ============================================================
// FIVE9 - SUPERVISOR (live state)
// ============================================================

function five9RawStatsXml_(statType) {
  const p = PropertiesService.getScriptProperties().getProperties();
  if (!p.FIVE9_USER || !p.FIVE9_PASS) throw new Error('Five9 credentials missing.');

  const url = CFG.FIVE9_HOST + CFG.FIVE9_PATH;
  const auth = 'Basic ' + Utilities.base64Encode(p.FIVE9_USER + ':' + p.FIVE9_PASS);

  const post = function (bodyXml, extraHeaders) {
    const headers = { Authorization: auth, SOAPAction: '' };
    if (extraHeaders) for (const k in extraHeaders) headers[k] = extraHeaders[k];
    return UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'text/xml; charset=utf-8',
      headers: headers,
      muteHttpExceptions: true,
      payload:
        '<?xml version="1.0" encoding="UTF-8"?>' +
        '<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" ' +
        'xmlns:ser="http://service.supervisor.ws.five9.com/">' +
          '<soapenv:Body>' + bodyXml + '</soapenv:Body>' +
        '</soapenv:Envelope>'
    });
  };

  const sess = post(
    '<ser:setSessionParameters><viewSettings>' +
      '<forceLogoutSession>true</forceLogoutSession>' +
      '<rollingPeriod>Minutes30</rollingPeriod>' +
      '<statisticsRange>CurrentDay</statisticsRange>' +
      '<shiftStart>' + CFG.SHIFT_START_MS + '</shiftStart>' +
      '<timeZone>' + CFG.TIMEZONE_OFFSET_MS + '</timeZone>' +
    '</viewSettings></ser:setSessionParameters>'
  );
  if (sess.getResponseCode() !== 200) {
    throw new Error('setSessionParameters failed (' + sess.getResponseCode() + '): ' +
                    sess.getContentText().slice(0, 600));
  }

  const cookie = extractCookie_(sess);
  const stats = post(
    '<ser:getStatistics><statisticType>' + statType + '</statisticType></ser:getStatistics>',
    cookie ? { Cookie: cookie } : null
  );
  if (stats.getResponseCode() !== 200) {
    throw new Error(statType + ' failed (' + stats.getResponseCode() + '): ' +
                    stats.getContentText().slice(0, 600));
  }
  return stats.getContentText();
}

function five9GetAgentStates_() {
  return parseStatistics_(five9RawStatsXml_('AgentState'));
}

// ============================================================
// FIVE9 - ADMIN (saved reports)
// ============================================================

function five9Admin_(bodyXml) {
  const p = PropertiesService.getScriptProperties().getProperties();
  const auth = 'Basic ' + Utilities.base64Encode(p.FIVE9_USER + ':' + p.FIVE9_PASS);
  return UrlFetchApp.fetch(CFG.FIVE9_HOST + CFG.FIVE9_ADMIN_PATH, {
    method: 'post',
    contentType: 'text/xml; charset=utf-8',
    headers: { Authorization: auth, SOAPAction: '' },
    muteHttpExceptions: true,
    payload:
      '<?xml version="1.0" encoding="UTF-8"?>' +
      '<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" ' +
      'xmlns:ser="http://service.admin.ws.five9.com/">' +
        '<soapenv:Body>' + bodyXml + '</soapenv:Body>' +
      '</soapenv:Envelope>'
  });
}

/** Kicks off a saved report, waits for it, returns the raw result XML. */
function five9RunReportXml_(folder, name, start, end) {
  const runResp = five9Admin_(
    '<ser:runReport>' +
      '<folderName>' + xmlEscape_(folder) + '</folderName>' +
      '<reportName>' + xmlEscape_(name) + '</reportName>' +
      '<criteria><time>' +
        '<end>' + five9Time_(end) + '</end>' +
        '<start>' + five9Time_(start) + '</start>' +
      '</time></criteria>' +
    '</ser:runReport>'
  );
  if (runResp.getResponseCode() !== 200) {
    Logger.log('runReport failed (' + runResp.getResponseCode() + '): ' +
               runResp.getContentText().slice(0, 500));
    return null;
  }

  const idEl = firstByName_(XmlService.parse(runResp.getContentText()).getRootElement(), 'return');
  if (!idEl) { Logger.log('runReport gave no identifier.'); return null; }
  const id = idEl.getText();

  let running = true;
  for (let i = 0; i < 15 && running; i++) {
    Utilities.sleep(1500);
    const r = five9Admin_('<ser:isReportRunning><identifier>' + id + '</identifier>' +
                          '<timeout>5</timeout></ser:isReportRunning>');
    const el = firstByName_(XmlService.parse(r.getContentText()).getRootElement(), 'return');
    running = el ? (el.getText() === 'true') : false;
  }
  if (running) { Logger.log('Report still running, gave up.'); return null; }

  const res = five9Admin_('<ser:getReportResult><identifier>' + id + '</identifier></ser:getReportResult>');
  if (res.getResponseCode() !== 200) {
    Logger.log('getReportResult failed (' + res.getResponseCode() + ')');
    return null;
  }
  return res.getContentText();
}

/** Runs a saved report and returns { header: [...], records: [{col: val}, ...] }. */
function five9RunReport_(folder, name, start, end) {
  const xml = five9RunReportXml_(folder, name, start, end);
  return xml ? parseReportResult_(xml) : null;
}

/**
 * Report shape:
 *   <return>
 *     <header><values><data>AGENT</data>...</values></header>
 *     <records><values><data>agent002</data>...</values></records>
 *     <records>...</records>
 *   </return>
 */
function parseReportResult_(xml) {
  const ret = firstByName_(XmlService.parse(xml).getRootElement(), 'return');
  if (!ret) return { header: [], records: [] };

  const headEl = firstByName_(ret, 'header');
  const header = headEl ? dataCells_(headEl) : [];

  const records = [];
  ret.getChildren().forEach(function (child) {
    if (child.getName() !== 'records') return;
    const vals = dataCells_(child);
    const obj = {};
    for (let i = 0; i < vals.length; i++) obj[header[i] || ('col' + i)] = vals[i];
    records.push(obj);
  });

  return { header: header, records: records };
}

function five9Time_(d) {
  return Utilities.formatDate(d, CFG.TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss.SSSXXX");
}

function xmlEscape_(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

// ============================================================
// XML PARSING (shared)
// ============================================================

function extractCookie_(resp) {
  const h = resp.getAllHeaders();
  let raw = h['Set-Cookie'] || h['set-cookie'];
  if (!raw) return '';
  if (!Array.isArray(raw)) raw = [raw];
  const parts = [];
  raw.forEach(function (c) { parts.push(String(c).split(';')[0]); });
  return parts.join('; ');
}

function parseStatistics_(xml) {
  const ret = firstByName_(XmlService.parse(xml).getRootElement(), 'return');
  if (!ret) throw new Error('No <return> in Five9 response: ' + xml.slice(0, 600));

  const colsEl = firstByName_(ret, 'columns');
  const columns = colsEl ? dataCells_(colsEl) : [];

  const rows = [];
  ret.getChildren().forEach(function (child) {
    if (child.getName() !== 'rows') return;
    const vals = dataCells_(child);
    const obj = {};
    for (let i = 0; i < vals.length; i++) obj[columns[i] || ('col' + i)] = vals[i];
    rows.push(obj);
  });

  return { columns: columns, rows: rows };
}

/** A <columns>/<rows>/<header>/<records> node -> its <values> child -> each <data>'s text. */
function dataCells_(el) {
  const valuesEl = firstByName_(el, 'values');
  if (!valuesEl) return [];
  return valuesEl.getChildren().map(function (d) { return d.getText(); });
}

function firstByName_(el, name) {
  const kids = el.getChildren();
  for (let i = 0; i < kids.length; i++) {
    if (kids[i].getName() === name) return kids[i];
    const deep = firstByName_(kids[i], name);
    if (deep) return deep;
  }
  return null;
}

function pickUsername_(row) {
  const keys = ['Username', 'User Name', 'Agent', 'Agent Name', 'Full Name', 'Name'];
  for (let i = 0; i < keys.length; i++) {
    if (row[keys[i]] !== undefined && row[keys[i]] !== '') return row[keys[i]];
  }
  return row['col0'];
}

function pickState_(row) {
  const keys = ['State', 'Agent State', 'Current State', 'Status'];
  for (let i = 0; i < keys.length; i++) {
    if (row[keys[i]] !== undefined && row[keys[i]] !== '') return row[keys[i]];
  }
  return '';
}

// ============================================================
// SLACK
// ============================================================

function slackSetStatus_(slackUserId, text, emoji) {
  const token = PropertiesService.getScriptProperties().getProperty('SLACK_TOKEN');
  if (!token) throw new Error('SLACK_TOKEN missing.');

  const resp = UrlFetchApp.fetch('https://slack.com/api/users.profile.set', {
    method: 'post',
    contentType: 'application/json; charset=utf-8',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify({
      user: slackUserId,
      profile: { status_text: text, status_emoji: emoji, status_expiration: 0 }
    }),
    muteHttpExceptions: true
  });

  let body;
  try { body = JSON.parse(resp.getContentText()); } catch (e) { body = { ok: false, error: 'bad_json' }; }
  if (!body.ok) {
    Logger.log('Slack failed for ' + slackUserId + ': ' + body.error);
    return false;
  }
  return true;
}

// ============================================================
// HELPERS
// ============================================================

function loadStateMap_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.SHEET_STATEMAP);
  const last = sheet.getLastRow();
  const map = {};
  if (last < 2) return map;
  sheet.getRange(2, 1, last - 1, 2).getValues().forEach(function (r) {
    const state = String(r[0] || '').trim().toLowerCase();
    const bucket = String(r[1] || '').trim();
    if (state && bucket) map[state] = bucket;
  });
  return map;
}

function withinHours_() {
  if (!CFG.HOURS.ENABLED) return true;
  const now = new Date();
  const day = now.getDay() === 0 ? 7 : now.getDay();
  if (CFG.HOURS.DAYS.indexOf(day) === -1) return false;
  const h = now.getHours();
  return h >= CFG.HOURS.START && h < CFG.HOURS.END;
}

/** Midnight at the start of the current pay week. */
function weekStart_(d) {
  const s = new Date(d);
  const diff = (s.getDay() - CFG.WEEK_STARTS_ON + 7) % 7;
  s.setDate(s.getDate() - diff);
  s.setHours(0, 0, 0, 0);
  return s;
}

/** "04:30:46" -> 16246 seconds. Blank, nil or junk -> 0. */
function hhmmssToSec_(v) {
  const s = String(v == null ? '' : v).trim();
  if (!s) return 0;
  const parts = s.split(':');
  if (parts.length !== 3) return 0;
  const h = parseInt(parts[0], 10), m = parseInt(parts[1], 10), sec = parseInt(parts[2], 10);
  if (isNaN(h) || isNaN(m) || isNaN(sec)) return 0;
  return h * 3600 + m * 60 + sec;
}

/** 16246 -> "4:30:46". No rounding, ever. Hours can exceed 24. */
function secToHhmmss_(total) {
  const n = Math.max(0, Math.floor(Number(total) || 0));
  const h = Math.floor(n / 3600);
  const m = Math.floor((n % 3600) / 60);
  const s = n % 60;
  return h + ':' + (m < 10 ? '0' : '') + m + ':' + (s < 10 ? '0' : '') + s;
}

// ============================================================
// HOW TO DEPLOY /myhours   (morning checklist)
//
// 1. In CFG above, change WEB_SECRET to any random string. Write it down.
//
// 2. Run setupSheet()  -> creates the new "Hours" tab.
//
// 3. Run debugHours()  -> should log real talk/wait per agent. Sanity check.
//
// 4. Deploy the web app:
// Deploy -> New deployment -> gear icon -> Web app
// Description:  myhours
// Execute as:   Me
// Who has access: Anyone
// Deploy, then COPY THE WEB APP URL (ends in /exec)
//
// 5. In Slack, api.slack.com/apps -> Five9 Status Sync -> Slash Commands
// Create New Command
// Command:      /myhours
// Request URL:  <the /exec URL>?k=<your WEB_SECRET>
// Short description: See your talk and wait time for the week
// Save
//
// 6. Slack will say the app needs reinstalling (adding a command adds the
// "commands" scope). Reinstall it.
// THEN go back to OAuth & Permissions and check whether the User OAuth
// Token changed. If it did, put the new one into saveSecrets(), run it
// once, and wipe the values again. If it's the same, do nothing.
//
// 7. Run installTriggers()  -> starts the 1-minute status sync, the
// 15-minute hours cache and the 30-minute boards.
//
// 8. In Slack, type /myhours.
//
// Troubleshooting:
// - "I don't have you linked" -> that Slack member ID isn't in the Agents tab
// - "No hours recorded"       -> refreshHoursCache() hasn't run yet, or that
// agent has no time this week
// - Nothing happens at all    -> the ?k= secret in the Slack URL doesn't match
// CFG.WEB_SECRET
// - debugMyHours() runs the whole thing without Slack. Put a real member ID
// in it and check the JSON it logs.
// ============================================================
