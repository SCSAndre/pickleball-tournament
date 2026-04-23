// ============================================================================
// UI.gs
// Custom menu, HTML dialogs, and score-entry CRUD functions.
// Depends on: Config.gs
// ============================================================================

// ============================================================================
// CUSTOM MENU
// ============================================================================

/**
 * Creates the 🏆 Tournament menu when the spreadsheet opens.
 * onOpen is an Apps Script trigger — must stay in the global scope.
 */
function onOpen() {
  const ui   = SpreadsheetApp.getUi();
  const menu = ui.createMenu('🏆 Tournament');

  menu.addSubMenu(ui.createMenu('📋 Setup')
    .addItem('🗂️ Create Roster',        'setupRoster')
    .addItem('🆕 New Tournament',       'setupTournament')
    .addItem('⚙️ Configure Settings',   'showConfigDialog')
    .addItem('🎨 Customize Theme',      'showThemeDialog'));

  menu.addSubMenu(ui.createMenu('🎯 Matches')
    .addItem('➕ Enter Score',           'showScoreEntryDialog')
    .addItem('🔴 Live Match Tracker',   'showLiveMatchTracker')
    .addItem('📅 Schedule Next Match',  'scheduleNextMatch')
    .addItem('🗑️ Clear All Scores',     'clearScoresOnly'));

  menu.addSubMenu(ui.createMenu('📊 Analytics')
    .addItem('📈 Statistics Dashboard',     'createEnhancedStatsDashboard')
    .addItem('🤝 Partnership Compatibility','createPartnershipCompatibility')
    .addItem('🎮 Player Profiles',          'createPlayerProfiles')
    .addItem('🏅 Achievements',             'createAchievementsSheet')
    .addItem('✅ Check Achievements Now',   'checkAndAwardAchievements')
    .addItem('🔮 Predictions & Insights',   'createPredictionsSheet')
    .addItem('📊 Generate All Charts',      'createAllCharts'));

  menu.addSubMenu(ui.createMenu('📧 Notifications')
    .addItem('📧 Email Current Standings',  'emailStandingsToAll')
    .addItem('🔔 Send Match Reminders',     'sendMatchReminders')
    .addItem('🎉 Announce Winners',         'announceWinners')
    .addItem('⚙️ Configure Email Settings', 'configureEmailSettings'));

  menu.addSubMenu(ui.createMenu('💾 Export')
    .addItem('📄 Generate PDF Report',  'generateCompletePDFReport')
    .addItem('📧 Email PDF to All',     'emailPDFToAll')
    .addItem('💾 Backup Tournament',   'backupTournament')
    .addItem('📦 Archive & Start New', 'archiveAndStartNew')
    .addItem('📊 Export to Excel',     'exportToExcel'));

  menu.addSeparator();
  menu.addItem('🔒 Protect Formulas',         'protectFormulaCells');
  menu.addItem('🔄 Refresh All Data',         'refreshAllData');
  menu.addItem('🎲 Generate Playoff Bracket', 'generatePlayoffBracket');
  menu.addItem('❓ Help & Documentation',     'showHelp');

  menu.addToUi();
}

// ============================================================================
// HELP DIALOG
// ============================================================================

/**
 * Shows the quick-start help dialog.
 */
function showHelp() {
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html><html>
    <head><base target="_top"><style>
      body{font-family:Arial,sans-serif;padding:20px;line-height:1.6}
      h2{color:${CONFIG.COLORS.PRIMARY}}
      .feature{margin:15px 0;padding:10px;background:#f5f5f5;border-radius:5px}
      .feature-title{font-weight:bold;color:#333}
      code{background:#e8e8e8;padding:2px 6px;border-radius:3px}
    </style></head>
    <body>
      <h2>🏆 Tournament System — Quick Start</h2>
      <div class="feature">
        <div class="feature-title">📋 Getting Started</div>
        1. Click <code>Setup → Create Roster</code> and add your players<br>
        2. Run <code>Setup → New Tournament</code> to generate all sheets<br>
        3. Enter scores via <code>Matches → Enter Score</code>
      </div>
      <div class="feature">
        <div class="feature-title">🗂️ Roster Sheet</div>
        - The Roster is the single source of truth for players<br>
        - Set Status = "Active" to include a player<br>
        - Works with any number of players (min 4)
      </div>
      <div class="feature">
        <div class="feature-title">📧 Email Notifications</div>
        - Set Notification = "Yes" in the Roster<br>
        - Configure via <code>Notifications → Configure Email Settings</code>
      </div>
      <div class="feature">
        <div class="feature-title">🏅 Achievements</div>
        - Run <code>Analytics → Check Achievements Now</code> after entering scores<br>
        - 15 achievements across 4 rarity tiers
      </div>
      <div class="feature">
        <div class="feature-title">📄 PDF Export</div>
        - Use <code>Export → Generate PDF Report</code> for a summary sheet<br>
        - Then File → Download → PDF from Google Sheets
      </div>
      <p><strong>Admin:</strong> ${CONFIG.ADMIN_EMAIL}</p>
    </body></html>
  `).setWidth(600).setHeight(520);
  SpreadsheetApp.getUi().showModalDialog(html, '📖 Help & Documentation');
}

// ============================================================================
// SCORE ENTRY DIALOG
// ============================================================================

/**
 * Shows the enhanced score entry modal dialog.
 */
function showScoreEntryDialog() {
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html><html>
    <head><base target="_top"><style>
      body{font-family:'Segoe UI',sans-serif;margin:0;padding:0;background:linear-gradient(135deg,#667eea,#764ba2)}
      .container{background:white;margin:20px;padding:30px;border-radius:12px;box-shadow:0 10px 30px rgba(0,0,0,.2)}
      h2{color:${CONFIG.COLORS.PRIMARY};margin-top:0;text-align:center;font-size:22px}
      .form-group{margin-bottom:18px}
      label{display:block;margin-bottom:6px;font-weight:600;color:#333}
      select,input{width:100%;padding:11px;border:2px solid #e0e0e0;border-radius:6px;box-sizing:border-box;font-size:14px;transition:border-color .3s}
      select:focus,input:focus{outline:none;border-color:${CONFIG.COLORS.PRIMARY}}
      .matchup{background:linear-gradient(135deg,#667eea,#764ba2);color:white;padding:18px;border-radius:8px;margin:18px 0;text-align:center;font-size:15px;font-weight:bold;display:none}
      .vs{display:inline-block;background:rgba(255,255,255,.3);padding:4px 14px;border-radius:20px;margin:8px 0}
      .score-inputs{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin:18px 0}
      .team-section{background:#f5f5f5;padding:14px;border-radius:8px}
      .team-a{border-left:4px solid ${CONFIG.COLORS.PRIMARY}}
      .team-b{border-left:4px solid ${CONFIG.COLORS.DANGER}}
      button{background:linear-gradient(135deg,#667eea,#764ba2);color:white;border:none;padding:14px;border-radius:8px;cursor:pointer;font-size:15px;font-weight:bold;width:100%;margin-top:8px;transition:transform .2s}
      button:hover{transform:translateY(-2px);box-shadow:0 5px 15px rgba(102,126,234,.4)}
      .error{color:#d93025;background:#fce8e6;padding:12px;border-radius:6px;margin-top:10px;border-left:4px solid #d93025}
      .success{color:#0f9d58;background:#e6f4ea;padding:12px;border-radius:6px;margin-top:10px;border-left:4px solid #0f9d58}
      .loading{text-align:center;padding:18px;color:#666}
    </style></head>
    <body>
      <div class="container">
        <h2>🎯 Enter Match Score</h2>
        <div class="form-group">
          <label for="match">Select Match:</label>
          <select id="match" onchange="updateMatchup()">
            <option value="">-- Select Match --</option>
          </select>
        </div>
        <div id="matchup" class="matchup"><div id="matchupText"></div></div>
        <div class="score-inputs">
          <div class="team-section team-a">
            <label for="scoreA">Team A Score:</label>
            <input type="number" id="scoreA" min="0" step="1" placeholder="0">
          </div>
          <div class="team-section team-b">
            <label for="scoreB">Team B Score:</label>
            <input type="number" id="scoreB" min="0" step="1" placeholder="0">
          </div>
        </div>
        <button onclick="submitScore()">💾 Save Score</button>
        <div id="message"></div>
      </div>
      <script>
        window.onload = function() {
          google.script.run.withSuccessHandler(populateMatches).getMatchList();
        };
        function populateMatches(matches) {
          const sel = document.getElementById('match');
          matches.forEach(m => {
            const opt = document.createElement('option');
            opt.value = m.number;
            opt.text  = 'Match ' + m.number + ': ' + m.matchup + ' ' + m.status;
            sel.appendChild(opt);
          });
        }
        function updateMatchup() {
          const sel     = document.getElementById('match');
          const box     = document.getElementById('matchup');
          const txt     = document.getElementById('matchupText');
          if (!sel.value) { box.style.display='none'; return; }
          google.script.run.withSuccessHandler(function(d) {
            txt.innerHTML = '<strong>' + d.teamA + '</strong><div class="vs">VS</div><strong>' + d.teamB + '</strong>';
            box.style.display = 'block';
            if (d.scoreA !== '') document.getElementById('scoreA').value = d.scoreA;
            if (d.scoreB !== '') document.getElementById('scoreB').value = d.scoreB;
          }).getMatchDetails(sel.value);
        }
        function submitScore() {
          const matchNum = document.getElementById('match').value;
          const scoreA   = document.getElementById('scoreA').value;
          const scoreB   = document.getElementById('scoreB').value;
          const msg      = document.getElementById('message');
          if (!matchNum)              { msg.innerHTML='<div class="error">❌ Please select a match</div>'; return; }
          if (scoreA===''||scoreB==='') { msg.innerHTML='<div class="error">❌ Enter both scores</div>'; return; }
          if (scoreA<0||scoreB<0)     { msg.innerHTML='<div class="error">❌ Scores must be ≥ 0</div>'; return; }
          msg.innerHTML = '<div class="loading">⏳ Saving…</div>';
          google.script.run.withSuccessHandler(function(r) {
            if (r.success) {
              msg.innerHTML = '<div class="success">✅ Score saved!</div>';
              setTimeout(function(){ google.script.host.close(); }, 1400);
            } else {
              msg.innerHTML = '<div class="error">❌ ' + r.message + '</div>';
            }
          }).saveMatchScore(matchNum, scoreA, scoreB);
        }
      </script>
    </body></html>
  `).setWidth(550).setHeight(540);
  SpreadsheetApp.getUi().showModalDialog(html, 'Enter Match Score');
}

// ============================================================================
// SCORE CRUD (called from HTML dialog via google.script.run)
// ============================================================================

/**
 * Returns a list of all matches from the Schedule sheet for the dropdown.
 * @returns {Array<{number:number, matchup:string, status:string}>}
 */
function getMatchList() {
  const ss    = SpreadsheetApp.getActive();
  const sched = ss.getSheetByName('Schedule');
  if (!sched) return [];

  const lastRow = sched.getLastRow();
  if (lastRow < 2) return [];

  const data    = sched.getRange(2, 1, lastRow - 1, 8).getValues();
  return data
    .filter(row => row[0] !== '')
    .map(row => ({
      number:  row[0],
      matchup: `${row[1]} + ${row[2]} vs ${row[3]} + ${row[4]}`,
      status:  row[7] === '✓' ? '(✓)' : '(⏳)'
    }));
}

/**
 * Returns team names and existing scores for the selected match.
 * @param {number} matchNum
 * @returns {{teamA:string, teamB:string, scoreA:number|string, scoreB:number|string}}
 */
function getMatchDetails(matchNum) {
  const ss    = SpreadsheetApp.getActive();
  const sched = ss.getSheetByName('Schedule');
  // Match N is at row N+1 (header in row 1, data from row 2)
  const row   = parseInt(matchNum) + 1;

  const data = sched.getRange(row, 1, 1, 7).getValues()[0];
  return {
    teamA:  `${data[1]} + ${data[2]}`,
    teamB:  `${data[3]} + ${data[4]}`,
    scoreA: data[5] !== '' ? data[5] : '',
    scoreB: data[6] !== '' ? data[6] : ''
  };
}

/**
 * Writes a score pair to the Schedule sheet.
 * @param {number} matchNum
 * @param {number|string} scoreA
 * @param {number|string} scoreB
 * @returns {{success:boolean, message?:string}}
 */
function saveMatchScore(matchNum, scoreA, scoreB) {
  try {
    const ss    = SpreadsheetApp.getActive();
    const sched = ss.getSheetByName('Schedule');
    const row   = parseInt(matchNum) + 1;

    if (!isValidScore(scoreA) || !isValidScore(scoreB)) {
      return { success: false, message: 'Invalid score values' };
    }

    sched.getRange(row, 6).setValue(parseInt(scoreA));
    sched.getRange(row, 7).setValue(parseInt(scoreB));
    sched.getRange(row, 9).setValue(getTimestamp());

    logActivity(`Match ${matchNum} score: ${scoreA}-${scoreB}`);

    // Check tournament completion
    const lastRow   = sched.getLastRow();
    const statusCol = lastRow > 1 ? sched.getRange(2, 8, lastRow - 1, 1).getValues() : [];
    const completed = statusCol.filter(r => r[0] === '✓').length;
    const total     = lastRow - 1;
    if (completed === total && total > 0) {
      logActivity('Tournament completed — all matches finished!');
    }

    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ============================================================================
// CONFIG & THEME DIALOGS
// ============================================================================

/**
 * Shows the tournament configuration dialog.
 */
function showConfigDialog() {
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html><html>
    <head><base target="_top"><style>
      body{font-family:Arial;padding:20px}
      h2{color:${CONFIG.COLORS.PRIMARY}}
      .section{margin:20px 0;padding:15px;background:#f5f5f5;border-radius:5px}
      label{display:block;margin:10px 0 5px;font-weight:bold}
      input,select{width:100%;padding:8px;border:1px solid #ddd;border-radius:4px;box-sizing:border-box}
      button{background:${CONFIG.COLORS.PRIMARY};color:white;padding:12px 24px;border:none;border-radius:4px;cursor:pointer;margin-top:15px}
      small{color:#666;font-size:11px}
    </style></head>
    <body>
      <h2>⚙️ Tournament Configuration</h2>
      <div class="section">
        <h3>General Settings</h3>
        <label>Tournament Name:</label>
        <input type="text" id="tournamentName" value="${CONFIG.TOURNAMENT_NAME}">
        <label>Admin Email:</label>
        <input type="email" id="adminEmail" value="${CONFIG.ADMIN_EMAIL}">
      </div>
      <div class="section">
        <h3>Feature Toggles</h3>
        <label><input type="checkbox" id="enableEmail" ${CONFIG.ENABLE_EMAIL_NOTIFICATIONS ? 'checked' : ''}> Enable Email Notifications</label>
        <label><input type="checkbox" id="enableAch"   ${CONFIG.ENABLE_ACHIEVEMENTS ? 'checked' : ''}> Enable Achievement System</label>
        <label><input type="checkbox" id="enablePred"  ${CONFIG.ENABLE_PREDICTIONS ? 'checked' : ''}> Enable AI Predictions</label>
      </div>
      <small>⚠️ To persist changes, update the CONFIG object in Config.gs directly.</small><br>
      <button onclick="google.script.host.close()">Close</button>
    </body></html>
  `).setWidth(500).setHeight(480);
  SpreadsheetApp.getUi().showModalDialog(html, 'Configuration');
}

/**
 * Shows the theme color customization dialog.
 */
function showThemeDialog() {
  const c = CONFIG.COLORS;
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html><html>
    <head><base target="_top"><style>
      body{font-family:Arial;padding:20px}
      h2{color:${c.PRIMARY}}
      .row{display:flex;align-items:center;margin:12px 0}
      .row label{width:160px;font-weight:bold}
      .row input{width:80px;height:38px;border:none;cursor:pointer;border-radius:4px}
      button{background:${c.PRIMARY};color:white;padding:12px 24px;border:none;border-radius:4px;cursor:pointer;margin-top:15px}
      small{color:#666;font-size:11px}
    </style></head>
    <body>
      <h2>🎨 Customize Theme</h2>
      <p>Preview colors — update CONFIG.COLORS in Config.gs to persist changes.</p>
      <div class="row"><label>Primary:</label>   <input type="color" value="${c.PRIMARY}"></div>
      <div class="row"><label>Success:</label>   <input type="color" value="${c.SUCCESS}"></div>
      <div class="row"><label>Warning:</label>   <input type="color" value="${c.WARNING}"></div>
      <div class="row"><label>Danger:</label>    <input type="color" value="${c.DANGER}"></div>
      <div class="row"><label>Gold (1st):</label><input type="color" value="${c.GOLD}"></div>
      <div class="row"><label>Win color:</label> <input type="color" value="${c.WIN}"></div>
      <div class="row"><label>Loss color:</label><input type="color" value="${c.LOSS}"></div>
      <small>Copy the hex values above into CONFIG.COLORS in Config.gs, then re-run Setup.</small><br>
      <button onclick="google.script.host.close()">Close</button>
    </body></html>
  `).setWidth(420).setHeight(480);
  SpreadsheetApp.getUi().showModalDialog(html, 'Theme Customization');
}
