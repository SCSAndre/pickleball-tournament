// ============================================================================
// Emails.gs
// All email notification functions. Reads player data from Roster dynamically.
// Depends on: Config.gs
// ============================================================================

// ============================================================================
// STANDINGS EMAIL
// ============================================================================

/**
 * Sends current standings to all players with Notification = Yes.
 * Player emails are read dynamically from the Roster sheet.
 */
function emailStandingsToAll() {
  const ui = SpreadsheetApp.getUi();

  if (!CONFIG.ENABLE_EMAIL_NOTIFICATIONS) {
    ui.alert('Email Disabled', 'Email notifications are currently disabled in CONFIG.', ui.ButtonSet.OK);
    return;
  }

  const response = ui.alert(
    'Send Standings Email',
    'This will send current standings to all players with notifications enabled. Continue?',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  try {
    const ss         = SpreadsheetApp.getActive();
    const players    = getActivePlayers(ss);
    const leaderboard = ss.getSheetByName('Leaderboard');

    if (!leaderboard) {
      ui.alert('Error', 'Leaderboard sheet not found. Please run tournament setup first.', ui.ButtonSet.OK);
      return;
    }

    const N           = players.length;
    const standingsData = leaderboard.getRange(1, 1, N + 1, 10).getValues();
    const sheetUrl    = ss.getUrl();

    let sent = 0, failed = 0;

    players.forEach(player => {
      if (player.notification !== 'Yes' || !player.email.includes('@')) return;
      try {
        MailApp.sendEmail({
          to:       player.email,
          subject:  `🏆 ${CONFIG.TOURNAMENT_NAME} — Standings Update`,
          htmlBody: buildStandingsEmail(player.name, standingsData, sheetUrl)
        });
        sent++;
      } catch (e) {
        console.log(`Failed to send to ${player.name}: ${e.message}`);
        failed++;
      }
    });

    logActivity(`Standings emails: ${sent} sent, ${failed} failed`);
    ui.alert('📧 Done', `Sent ${sent} email(s). Failed: ${failed}`, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', 'Failed to send emails: ' + e.message, ui.ButtonSet.OK);
  }
}

/**
 * Builds an HTML standings email for a specific player.
 * @param {string} playerName
 * @param {Array<Array>} standingsData - Row 0 = header, rows 1+ = standings
 * @param {string} sheetUrl
 * @returns {string} HTML email body
 */
function buildStandingsEmail(playerName, standingsData, sheetUrl) {
  let tableRows = '';
  for (let i = 1; i < standingsData.length; i++) {
    const row       = standingsData[i];
    const isPlayer  = row[1] === playerName;
    const medal     = i === 1 ? '🥇' : i === 2 ? '🥈' : i === 3 ? '🥉' : '';
    const highlight = isPlayer ? `style="background:${CONFIG.COLORS.GOLD};font-weight:bold;"` : '';
    const winPct    = typeof row[9] === 'number' ? (row[9] * 100).toFixed(1) + '%' : row[9];
    const diff      = Number(row[7]) > 0 ? '+' + row[7] : row[7];

    tableRows += `
      <tr ${highlight}>
        <td>${row[0]} ${medal}</td>
        <td><strong>${row[1]}</strong></td>
        <td>${row[2]}</td>
        <td>${row[3]}-${row[4]}</td>
        <td>${winPct}</td>
        <td>${diff}</td>
      </tr>`;
  }

  return `<!DOCTYPE html>
<html>
<head><style>
  body{font-family:Arial,sans-serif;line-height:1.6;color:#333}
  .container{max-width:600px;margin:0 auto;padding:20px}
  .header{background:${CONFIG.COLORS.PRIMARY};color:white;padding:20px;text-align:center;border-radius:8px 8px 0 0}
  .content{background:#f9f9f9;padding:20px}
  table{width:100%;border-collapse:collapse;margin:20px 0;background:white}
  th{background:${CONFIG.COLORS.PRIMARY};color:white;padding:12px;text-align:left}
  td{padding:10px;border-bottom:1px solid #ddd}
  .button{display:inline-block;padding:12px 24px;background:${CONFIG.COLORS.SUCCESS};color:white;text-decoration:none;border-radius:5px;margin:10px 0}
  .footer{text-align:center;padding:20px;color:#666;font-size:12px}
</style></head>
<body>
  <div class="container">
    <div class="header"><h1>🏆 ${CONFIG.TOURNAMENT_NAME}</h1><p>Current Standings Update</p></div>
    <div class="content">
      <p>Hello <strong>${playerName}</strong>!</p>
      <p>Here are the current tournament standings as of ${getTimestamp()}:</p>
      <table>
        <thead><tr><th>Rank</th><th>Player</th><th>Points</th><th>W-L</th><th>Win%</th><th>Diff</th></tr></thead>
        <tbody>${tableRows}</tbody>
      </table>
      <p>Keep up the great work! Check the full dashboard for detailed statistics.</p>
      <center><a href="${sheetUrl}" class="button">View Full Dashboard →</a></center>
    </div>
    <div class="footer">
      <p>Automated message from the Tournament Management System</p>
      <p>To stop receiving emails, set Notification to "No" in the Roster sheet.</p>
    </div>
  </div>
</body></html>`;
}

// ============================================================================
// MATCH REMINDERS
// ============================================================================

/**
 * Scans the Schedule for pending matches and emails the involved players.
 * Fully implemented.
 */
function sendMatchReminders() {
  const ui = SpreadsheetApp.getUi();

  if (!CONFIG.ENABLE_EMAIL_NOTIFICATIONS) {
    ui.alert('Email Disabled', 'Email notifications are disabled in CONFIG.', ui.ButtonSet.OK);
    return;
  }

  const response = ui.alert(
    '🔔 Send Match Reminders',
    'This will email all players with pending matches. Continue?',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  try {
    const ss      = SpreadsheetApp.getActive();
    const sched   = ss.getSheetByName('Schedule');
    const players = getActivePlayers(ss);

    if (!sched) {
      ui.alert('Error', 'Schedule sheet not found.', ui.ButtonSet.OK);
      return;
    }

    const M        = sched.getLastRow() - 1;
    const emailMap = {};
    players.forEach(p => { if (p.email.includes('@')) emailMap[p.name] = p.email; });

    const schedData  = M > 0 ? sched.getRange(2, 1, M, 9).getValues() : [];
    const pendingMap = {}; // playerName -> [{matchNum, teamA, teamB}]

    schedData.forEach(row => {
      const status = String(row[7]).trim();
      if (status === '✓') return;
      const matchNum = row[0];
      const teamA    = `${row[1]} + ${row[2]}`;
      const teamB    = `${row[3]} + ${row[4]}`;
      [row[1], row[2], row[3], row[4]].forEach(name => {
        name = String(name).trim();
        if (!name) return;
        if (!pendingMap[name]) pendingMap[name] = [];
        pendingMap[name].push({ matchNum, teamA, teamB });
      });
    });

    let sent = 0, failed = 0;

    Object.keys(pendingMap).forEach(name => {
      const email = emailMap[name];
      if (!email) return;
      const matches = pendingMap[name];
      const player  = players.find(p => p.name === name);
      if (!player || player.notification !== 'Yes') return;

      const matchList = matches.map(m =>
        `<li>Match ${m.matchNum}: <strong>${m.teamA}</strong> vs <strong>${m.teamB}</strong></li>`
      ).join('');

      const body = `<!DOCTYPE html><html><body style="font-family:Arial,sans-serif;">
        <div style="max-width:500px;margin:0 auto;padding:20px;">
          <h2 style="color:${CONFIG.COLORS.PRIMARY};">🔔 Match Reminder</h2>
          <p>Hi <strong>${name}</strong>,</p>
          <p>You have <strong>${matches.length}</strong> pending match(es) in <em>${CONFIG.TOURNAMENT_NAME}</em>:</p>
          <ul>${matchList}</ul>
          <p>Enter results via the Tournament menu when complete. Good luck!</p>
          <p style="color:#666;font-size:12px;">Update Notification to "No" in the Roster to stop these emails.</p>
        </div></body></html>`;

      try {
        MailApp.sendEmail({
          to:       email,
          subject:  `🔔 ${CONFIG.TOURNAMENT_NAME} — Match Reminder`,
          htmlBody: body
        });
        sent++;
      } catch (e) {
        console.log(`Reminder failed for ${name}: ${e.message}`);
        failed++;
      }
    });

    logActivity(`Match reminders: ${sent} sent, ${failed} failed`);
    ui.alert('🔔 Done', `Sent ${sent} reminder(s). Failed: ${failed}`, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', 'Could not send reminders: ' + e.message, ui.ButtonSet.OK);
  }
}

// ============================================================================
// WINNER ANNOUNCEMENT
// ============================================================================

/**
 * Announces tournament winners to all subscribed players after all matches complete.
 */
function announceWinners() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();

  const sched = ss.getSheetByName('Schedule');
  if (!sched) {
    ui.alert('Error', 'Schedule sheet not found.', ui.ButtonSet.OK);
    return;
  }

  const M         = sched.getLastRow() - 1;
  const statusCol = M > 0 ? sched.getRange(2, 8, M, 1).getValues() : [];
  const completed = statusCol.filter(r => r[0] === '✓').length;

  if (completed < M) {
    ui.alert(
      'Tournament Incomplete',
      `Only ${completed} of ${M} matches completed. All matches must finish before announcing winners.`,
      ui.ButtonSet.OK
    );
    return;
  }

  const response = ui.alert(
    '🎉 Announce Winners',
    'Send championship email to all subscribed players?',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  try {
    const players   = getActivePlayers(ss);
    const leaderboard = ss.getSheetByName('Leaderboard');
    const winner    = leaderboard ? String(leaderboard.getRange('B2').getValue()) : 'Unknown';
    const points    = leaderboard ? leaderboard.getRange('C2').getValue() : 0;
    const sheetUrl  = ss.getUrl();

    let sent = 0;
    players.forEach(player => {
      if (player.notification !== 'Yes' || !player.email.includes('@')) return;
      try {
        MailApp.sendEmail({
          to:       player.email,
          subject:  `🏆 ${CONFIG.TOURNAMENT_NAME} — CHAMPIONS ANNOUNCED!`,
          htmlBody: buildWinnersEmail(winner, points, sheetUrl)
        });
        sent++;
      } catch (e) {
        console.log(`Winner email failed for ${player.name}: ${e.message}`);
      }
    });

    logActivity(`Winner announcement sent to ${sent} players`);
    ui.alert('🎉 Success', `Championship emails sent to ${sent} players!`, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', 'Failed to announce winners: ' + e.message, ui.ButtonSet.OK);
  }
}

/**
 * Builds the championship announcement HTML email.
 * @param {string} winner
 * @param {number} points
 * @param {string} sheetUrl
 * @returns {string}
 */
function buildWinnersEmail(winner, points, sheetUrl) {
  return `<!DOCTYPE html>
<html>
<head><style>
  body{font-family:Arial,sans-serif;background:#f5f5f5}
  .container{max-width:600px;margin:20px auto;background:white;border-radius:10px;overflow:hidden;box-shadow:0 4px 6px rgba(0,0,0,.1)}
  .header{background:linear-gradient(135deg,#ffd700,#ffed4e);padding:40px;text-align:center}
  .trophy{font-size:72px}
  .winner-box{background:#ffd700;padding:20px;border-radius:8px;text-align:center;margin:20px 0}
  .button{display:inline-block;padding:15px 30px;background:${CONFIG.COLORS.PRIMARY};color:white;text-decoration:none;border-radius:5px;font-weight:bold}
  .footer{background:#f5f5f5;padding:20px;text-align:center;color:#666}
</style></head>
<body>
  <div class="container">
    <div class="header">
      <div class="trophy">🏆</div>
      <h1 style="margin:0;color:#333">${CONFIG.TOURNAMENT_NAME}</h1>
      <p style="font-size:18px;color:#666">Tournament Complete!</p>
    </div>
    <div style="padding:30px">
      <h2 style="text-align:center;color:${CONFIG.COLORS.PRIMARY}">🎉 CONGRATULATIONS! 🎉</h2>
      <div class="winner-box">
        <p style="margin:0;font-size:14px;text-transform:uppercase;letter-spacing:2px">Champion</p>
        <h2 style="margin:10px 0;font-size:32px">${winner}</h2>
        <p style="font-size:22px;margin:0">${points} Points</p>
      </div>
      <p style="text-align:center;font-size:16px">Thank you to all participants for an amazing tournament!</p>
      <center><a href="${sheetUrl}" class="button">View Final Results →</a></center>
    </div>
    <div class="footer"><p>🏆 ${CONFIG.TOURNAMENT_NAME} — ${getTimestamp()}</p></div>
  </div>
</body></html>`;
}

// ============================================================================
// EMAIL SETTINGS DIALOG
// ============================================================================

/**
 * Shows the email configuration modal dialog.
 */
function configureEmailSettings() {
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html><html>
    <head><base target="_top">
    <style>
      body{font-family:Arial;padding:20px}
      .setting{margin:15px 0}
      label{display:block;margin-bottom:5px;font-weight:bold}
      input,select{width:100%;padding:8px;border:1px solid #ddd;border-radius:4px;box-sizing:border-box}
      button{background:${CONFIG.COLORS.PRIMARY};color:white;padding:10px 20px;border:none;border-radius:4px;cursor:pointer;margin-top:15px}
    </style></head>
    <body>
      <h2>⚙️ Email Notification Settings</h2>
      <div class="setting">
        <label>Enable Notifications:</label>
        <select id="enabled">
          <option value="true" ${CONFIG.ENABLE_EMAIL_NOTIFICATIONS ? 'selected' : ''}>Yes</option>
          <option value="false" ${!CONFIG.ENABLE_EMAIL_NOTIFICATIONS ? 'selected' : ''}>No</option>
        </select>
      </div>
      <div class="setting">
        <label>Send Frequency:</label>
        <select id="frequency">
          <option value="24">Daily</option>
          <option value="48">Every 2 days</option>
          <option value="168">Weekly</option>
        </select>
      </div>
      <div class="setting">
        <label>Admin Email:</label>
        <input type="email" id="adminEmail" value="${CONFIG.ADMIN_EMAIL}">
      </div>
      <p style="font-size:12px;color:#666">Note: To change defaults permanently, edit CONFIG in Config.gs</p>
      <button onclick="google.script.host.close()">Close</button>
    </body></html>
  `).setWidth(420).setHeight(360);
  SpreadsheetApp.getUi().showModalDialog(html, 'Email Settings');
}
