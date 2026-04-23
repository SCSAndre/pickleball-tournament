// ============================================================================
// Setup.gs
// Tournament initialization: Roster sheet creation and full tournament setup.
// Depends on: Config.gs, CoreEngine.gs, Dashboard.gs, Analytics.gs, Advanced.gs
// ============================================================================

/**
 * Creates or resets the Roster sheet — the source of truth for all players.
 * Columns: Player Name | Email | Skill Rating | Notification | Status
 */
function setupRoster() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();

  const roster = getOrCreateSheet(ss, 'Roster');
  roster.clear();

  // Header row
  const headers = ['Player Name', 'Email', 'Skill Rating', 'Notification', 'Status'];
  roster.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader(roster.getRange(1, 1, 1, headers.length));
  roster.setFrozenRows(1);

  // Example players (users replace these)
  const examplePlayers = [
    ['Player One',   'player1@email.com', 3, 'Yes', 'Active'],
    ['Player Two',   'player2@email.com', 3, 'Yes', 'Active'],
    ['Player Three', 'player3@email.com', 3, 'Yes', 'Active'],
    ['Player Four',  'player4@email.com', 3, 'Yes', 'Active'],
    ['Player Five',  'player5@email.com', 3, 'Yes', 'Active'],
    ['Player Six',   'player6@email.com', 3, 'Yes', 'Active'],
    ['Player Seven', 'player7@email.com', 3, 'Yes', 'Active'],
    ['Player Eight', 'player8@email.com', 3, 'Yes', 'Active']
  ];

  roster.getRange(2, 1, examplePlayers.length, 5).setValues(examplePlayers);

  // Data validation: Notification (D column)
  const notifValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Yes', 'No'])
    .setAllowInvalid(false)
    .setHelpText('Yes = receive email notifications')
    .build();
  roster.getRange('D2:D200').setDataValidation(notifValidation);

  // Data validation: Status (E column)
  const statusValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Active', 'Inactive', 'Injured'])
    .setAllowInvalid(false)
    .setHelpText('Only Active players are included in the tournament')
    .build();
  roster.getRange('E2:E200').setDataValidation(statusValidation);

  // Data validation: Skill Rating (C column) — 1 to 5
  const ratingValidation = SpreadsheetApp.newDataValidation()
    .requireNumberBetween(1, 5)
    .setAllowInvalid(false)
    .setHelpText('Enter a skill rating from 1 (beginner) to 5 (expert)')
    .build();
  roster.getRange('C2:C200').setDataValidation(ratingValidation);

  // Instructions
  roster.getRange('A' + (examplePlayers.length + 3)).setValue('📝 INSTRUCTIONS:').setFontWeight('bold').setFontSize(12);
  roster.getRange('A' + (examplePlayers.length + 4)).setValue('• Replace example players with real names and emails');
  roster.getRange('A' + (examplePlayers.length + 5)).setValue('• Set Status to "Active" to include a player in the tournament');
  roster.getRange('A' + (examplePlayers.length + 6)).setValue('• Set Notification to "Yes" to send email updates');
  roster.getRange('A' + (examplePlayers.length + 7)).setValue('• Minimum 4 Active players required (divisible by 4 recommended)');
  roster.getRange('A' + (examplePlayers.length + 8)).setValue('• After editing, run Tournament → Setup → New Tournament to regenerate all sheets');

  roster.autoResizeColumns(1, 5);
  roster.setColumnWidth(2, 220);

  ss.setActiveSheet(roster);
  logActivity('Roster sheet initialized');

  ui.alert(
    '✅ Roster Created',
    '📋 The Roster sheet has been set up!\n\n' +
    'Next steps:\n' +
    '1. Replace the example players with real names and emails\n' +
    '2. Set each player\'s Status to "Active"\n' +
    '3. Run "Tournament → Setup → New Tournament" to generate all sheets',
    ui.ButtonSet.OK
  );
}

// ============================================================================
// MAIN SETUP ORCHESTRATOR
// ============================================================================

/**
 * Creates the complete tournament system based on active players from the Roster.
 * This is the entry point called from the menu.
 */
function setupTournament() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();

  // Ensure Roster exists
  if (!ss.getSheetByName('Roster')) {
    const rosterResp = ui.alert(
      '📋 Roster Required',
      'No Roster sheet found. Would you like to create one now?\n\n' +
      'You must add your players to the Roster before setting up the tournament.',
      ui.ButtonSet.YES_NO
    );
    if (rosterResp === ui.Button.YES) setupRoster();
    return;
  }

  const players = getActivePlayers(ss);

  if (players.length < 4) {
    ui.alert(
      '⚠️ Not Enough Players',
      `Only ${players.length} active player(s) found in the Roster.\n\n` +
      'You need at least 4 active players to create a tournament.\n' +
      'Please update the Roster sheet and try again.',
      ui.ButtonSet.OK
    );
    return;
  }

  const response = ui.alert(
    '🆕 Setup New Tournament',
    `Found ${players.length} active players:\n` +
    players.map(p => '• ' + p.name).join('\n') + '\n\n' +
    'This will create a complete tournament system. Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  ui.alert('⏳ Setting Up...', 'Creating tournament system. This may take up to 60 seconds...', ui.ButtonSet.OK);

  // Build all sheets in order
  createScheduleSheet(ss);
  createStandingsSheet(ss);
  createLeaderboardSheet(ss);
  createPartnershipsSheet(ss);
  createHeadToHeadSheet(ss);
  createPlayerManagementSheet(ss);
  createProgressSheet(ss);
  createEnhancedStatsDashboard(ss);
  createPartnershipCompatibility(ss);
  createPlayerProfiles(ss);
  createAchievementsSheet(ss);
  createPredictionsSheet(ss);
  createLiveMatchTracker(ss);

  // Charts & formatting
  createAllCharts();
  applyAdvancedFormatting();
  protectFormulaCells();

  logActivity(`Tournament setup: ${players.length} players, ${generateRoundRobinSchedule(players.map(p => p.name)).length} matches`);

  ui.alert(
    '✅ Tournament Ready!',
    `🎉 Tournament system created for ${players.length} players!\n\n` +
    '📊 Features enabled:\n' +
    '• Dynamic Schedule & Standings\n' +
    '• Statistics Dashboard & Charts\n' +
    '• Partnership Compatibility\n' +
    '• Player Profiles\n' +
    '• Achievements System\n' +
    '• AI Predictions\n' +
    '• Live Match Tracker\n' +
    '• Email Notifications\n' +
    '• PDF Export\n\n' +
    'Check the 🏆 Tournament menu for all options!',
    ui.ButtonSet.OK
  );
}

/**
 * Creates a legacy "Players" sheet mirroring the Roster for backward compatibility.
 * This is referenced by some email functions.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function createPlayerManagementSheet(ss) {
  const pm      = getOrCreateSheet(ss, 'Players');
  const players = getActivePlayers(ss);
  pm.clear();

  pm.getRange('A1:E1').setValues([['Player Name', 'Email', 'Phone', 'Notification', 'Status']]);
  styleHeader(pm.getRange('A1:E1'));
  pm.setFrozenRows(1);

  if (players.length > 0) {
    const rows = players.map(p => [p.name, p.email, '', p.notification, p.status]);
    pm.getRange(2, 1, rows.length, 5).setValues(rows);
  }

  pm.autoResizeColumns(1, 5);

  pm.getRange('A' + (players.length + 3)).setValue('📝 Edit players in the Roster sheet, then re-run Setup.').setFontStyle('italic').setFontColor('#666666');
}
