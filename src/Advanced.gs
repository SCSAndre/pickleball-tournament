// ============================================================================
// Advanced.gs
// PDF export, backup/archive, playoff bracket, protection, conditional
// formatting helpers, and miscellaneous utility actions.
// Depends on: Config.gs, CoreEngine.gs
// ============================================================================

// ============================================================================
// PDF EXPORT
// ============================================================================

/**
 * Shows PDF export instructions and then creates the PDF Summary sheet.
 */
function generateCompletePDFReport() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    '📄 Generate PDF Report',
    'Creating a formatted PDF Summary sheet…\n\n' +
    'After creation:\n' +
    '1. Select the "PDF Export Summary" sheet\n' +
    '2. Go to File → Download → PDF Document\n' +
    '3. Choose "Fit to Width" and click Export',
    ui.ButtonSet.OK
  );
  createPDFSummarySheet();
}

/**
 * Creates a print-ready PDF Export Summary sheet with dynamic ranges.
 */
function createPDFSummarySheet() {
  const ss      = SpreadsheetApp.getActive();
  const players = getActivePlayerNames(ss);
  const N       = players.length;
  const M       = generateRoundRobinSchedule(players).length;
  const nEnd    = N + 1;

  const summary = getOrCreateSheet(ss, 'PDF Export Summary');
  summary.clear();

  // Title page
  summary.getRange('A1:H1').merge()
    .setValue(`🏆 ${CONFIG.TOURNAMENT_NAME}`)
    .setFontSize(28).setFontWeight('bold').setHorizontalAlignment('center');
  summary.getRange('A2:H2').merge()
    .setValue('FINAL TOURNAMENT REPORT')
    .setFontSize(18).setHorizontalAlignment('center');
  summary.getRange('A3:H3').merge()
    .setValue(getTimestamp())
    .setFontSize(12).setHorizontalAlignment('center').setFontColor('#666666');

  // Executive Summary
  summary.getRange('A5:H5').merge().setValue('📊 EXECUTIVE SUMMARY')
    .setFontSize(16).setFontWeight('bold')
    .setBackground(CONFIG.COLORS.PRIMARY).setFontColor('#ffffff');

  summary.getRange('A7:B7').merge().setValue('Tournament Status:');
  summary.getRange('C7:D7').merge().setFormula(
    `=IF('Stats Dashboard'!C6=${M},"✅ COMPLETE","🔄 IN PROGRESS")`
  ).setFontWeight('bold');

  summary.getRange('A8:B8').merge().setValue('Total Matches:');
  summary.getRange('C8').setValue(M);

  summary.getRange('A9:B9').merge().setValue('Total Players:');
  summary.getRange('C9').setValue(N);

  summary.getRange('A10:B10').merge().setValue('Champion:');
  summary.getRange('C10:D10').merge()
    .setFormula('=Leaderboard!B2').setFontWeight('bold').setFontSize(14);

  // Final Standings
  summary.getRange('A13:H13').merge().setValue('🏆 FINAL STANDINGS')
    .setFontSize(16).setFontWeight('bold').setBackground(CONFIG.COLORS.GOLD);
  summary.getRange('A14').setFormula(`=Leaderboard!A1:J${nEnd}`);

  // Key Statistics
  summary.getRange('A' + (15 + N) + ':H' + (15 + N)).merge().setValue('📈 KEY STATISTICS')
    .setFontSize(16).setFontWeight('bold')
    .setBackground(CONFIG.COLORS.PRIMARY).setFontColor('#ffffff');

  const kRow = 16 + N;
  summary.getRange(`A${kRow}:B${kRow}`).merge().setValue('Most Wins:');
  summary.getRange(`C${kRow}:D${kRow}`).merge()
    .setFormula(`=INDEX(SORT(Standings!A2:C${nEnd},3,FALSE),1,1)`);

  summary.getRange(`A${kRow+1}:B${kRow+1}`).merge().setValue('Highest Win %:');
  summary.getRange(`C${kRow+1}:D${kRow+1}`).merge()
    .setFormula(`=INDEX(SORT(Standings!A2:I${nEnd},9,FALSE),1,1)`);

  summary.getRange(`A${kRow+2}:B${kRow+2}`).merge().setValue('Best Point Diff:');
  summary.getRange(`C${kRow+2}:D${kRow+2}`).merge()
    .setFormula(`=INDEX(SORT(Standings!A2:G${nEnd},7,FALSE),1,1)`);

  summary.setColumnWidths(1, 8, 120);
  summary.getRange('A1:H' + (kRow + 5)).setFontFamily('Arial');

  SpreadsheetApp.getUi().alert(
    '✅ PDF Summary Created',
    'The "PDF Export Summary" sheet is ready.\n\n' +
    'File → Download → PDF Document to export.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  ss.setActiveSheet(summary);
}

/**
 * Shows instructions for emailing the PDF (full Drive API required for automation).
 */
function emailPDFToAll() {
  SpreadsheetApp.getUi().alert(
    '📧 Email PDF Report',
    'To share the PDF:\n\n' +
    '1. Generate the PDF via "Export → Generate PDF Report"\n' +
    '2. File → Download → PDF Document\n' +
    '3. Upload to Google Drive and share the link\n\n' +
    'Automated attachment requires the Drive API (advanced setup).',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ============================================================================
// BACKUP & ARCHIVE
// ============================================================================

/**
 * Copies the Schedule sheet as a timestamped backup tab.
 */
function backupTournament() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  try {
    const timestamp  = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HHmmss');
    const backupName = 'Backup_' + timestamp;
    const sched      = ss.getSheetByName('Schedule');
    if (!sched) throw new Error('Schedule sheet not found');

    const backup = sched.copyTo(ss);
    backup.setName(backupName);
    backup.getRange('K1').setValue('Backup Date:');
    backup.getRange('L1').setValue(getTimestamp());

    logActivity('Backup created: ' + backupName);
    ui.alert('✅ Backup Created', 'Data backed up to sheet: ' + backupName, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('❌ Backup Failed', 'Error: ' + e.message, ui.ButtonSet.OK);
  }
}

/**
 * Archives the current tournament winner and starts a fresh tournament.
 */
function archiveAndStartNew() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Archive & Start New',
    'This will save the current winner to the Archive sheet and start a fresh tournament. Continue?',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  backupTournament();

  const ss      = SpreadsheetApp.getActive();
  const archive = getOrCreateSheet(ss, 'Archive');
  const timestamp = getTimestamp();

  if (archive.getLastRow() === 0) {
    archive.getRange('A1:D1').setValues([['Date','Winner','Points','Backup Sheet']]);
    styleHeader(archive.getRange('A1:D1'));
  }

  const lb = ss.getSheetByName('Leaderboard');
  if (lb) {
    const nextRow = archive.getLastRow() + 1;
    const winner  = lb.getRange('B2').getValue();
    const points  = lb.getRange('C2').getValue();
    archive.getRange(nextRow, 1).setValue(timestamp);
    archive.getRange(nextRow, 2).setValue(winner);
    archive.getRange(nextRow, 3).setValue(points);
    archive.getRange(nextRow, 4).setValue('Backup_' + timestamp.substring(0, 19).replace(' ', '_'));
  }

  setupTournament();
  ui.alert('✅ Complete', 'Tournament archived and new tournament created!', ui.ButtonSet.OK);
}

// ============================================================================
// SCORE MANAGEMENT
// ============================================================================

/**
 * Clears all entered scores from the Schedule (keeps structure).
 */
function clearScoresOnly() {
  const ui       = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Clear All Scores',
    'This clears all match scores but keeps the tournament structure. Continue?',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  const ss    = SpreadsheetApp.getActive();
  const sched = ss.getSheetByName('Schedule');
  if (!sched) return;

  const lastRow = sched.getLastRow();
  if (lastRow > 1) {
    sched.getRange(2, 6, lastRow - 1, 2).clearContent(); // ScoreA, ScoreB
    sched.getRange(2, 9, lastRow - 1, 1).clearContent(); // Date
  }

  logActivity('All scores cleared');
  ui.alert('✅ Scores Cleared', 'All match scores have been reset.', ui.ButtonSet.OK);
}

// ============================================================================
// FORMULA PROTECTION
// ============================================================================

/**
 * Applies warning-only protection to formula cells in key sheets.
 * Ranges are computed dynamically from active player/match counts.
 */
function protectFormulaCells() {
  const ss      = SpreadsheetApp.getActive();
  const ui      = SpreadsheetApp.getUi();
  const players = getActivePlayerNames(ss);
  const N       = players.length;
  const M       = generateRoundRobinSchedule(players).length;

  const protections = [
    { sheet: 'Standings',               range: `B2:K${N + 1}`,     desc: 'Standings calculations'   },
    { sheet: 'Leaderboard',             range: `A2:L${N + 1}`,     desc: 'Leaderboard auto-sort'     },
    { sheet: 'Partnerships',            range: `A2:G${M * 2 + 1}`, desc: 'Partnership stats'         },
    { sheet: 'Head-to-Head',            range: 'B8:B22',           desc: 'H2H calculations'          },
    { sheet: 'Stats Dashboard',         range: 'B5:H30',           desc: 'Statistics'                },
    { sheet: 'Partnership Compatibility', range: 'B5:G500',        desc: 'Compatibility matrix'      },
    { sheet: 'Player Profiles',         range: 'B1:F500',          desc: 'Player stats'              },
    { sheet: 'Achievements',            range: 'A1:H500',          desc: 'Achievement tracking'      },
    { sheet: 'Predictions',             range: 'A1:K500',          desc: 'Predictions & insights'    },
    { sheet: 'Schedule',                range: `H2:I${M + 1}`,     desc: 'Match status'              }
  ];

  let applied = 0;
  protections.forEach(p => {
    const sheet = ss.getSheetByName(p.sheet);
    if (!sheet) return;
    try {
      const prot = sheet.getRange(p.range).protect();
      prot.setDescription(p.desc + ' — Auto-calculated');
      prot.setWarningOnly(true);
      applied++;
    } catch (e) {
      console.log(`Could not protect ${p.sheet}: ${e.message}`);
    }
  });

  logActivity('Formula cells protected: ' + applied + ' ranges');
  ui.alert('✅ Protection Applied', `${applied} formula ranges are now protected from accidental edits.`, ui.ButtonSet.OK);
}

// ============================================================================
// CONDITIONAL FORMATTING HELPERS (called by Setup)
// ============================================================================

/**
 * Applies win/loss/status conditional formatting to the Schedule sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sched
 * @param {number} M - Number of matches
 */
function applyScheduleFormatting(sched, M) {
  if (M < 1) return;
  const dataRange = `2:${M + 1}`;
  const teamA     = sched.getRange(`B${dataRange}`);  // reused below
  const rules = [
    // Team A wins
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$F2>$G2').setBackground(CONFIG.COLORS.WIN).setBold(true)
      .setRanges([sched.getRange(`B2:C${M + 1}`)]).build(),
    // Team B wins
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$G2>$F2').setBackground(CONFIG.COLORS.WIN).setBold(true)
      .setRanges([sched.getRange(`D2:E${M + 1}`)]).build(),
    // Team A loses
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$F2<$G2').setBackground(CONFIG.COLORS.LOSS)
      .setRanges([sched.getRange(`B2:C${M + 1}`)]).build(),
    // Team B loses
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$G2<$F2').setBackground(CONFIG.COLORS.LOSS)
      .setRanges([sched.getRange(`D2:E${M + 1}`)]).build(),
    // Status: complete
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('✓').setBackground(CONFIG.COLORS.WIN).setBold(true)
      .setRanges([sched.getRange(`H2:H${M + 1}`)]).build(),
    // Status: pending
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('⏳').setBackground(CONFIG.COLORS.PENDING)
      .setRanges([sched.getRange(`H2:H${M + 1}`)]).build(),
    // Score gradient
    SpreadsheetApp.newConditionalFormatRule()
      .setGradientMaxpointWithValue('#4caf50', SpreadsheetApp.InterpolationType.NUMBER, '30')
      .setGradientMinpointWithValue('#ffffff', SpreadsheetApp.InterpolationType.NUMBER, '0')
      .setRanges([sched.getRange(`F2:G${M + 1}`)]).build()
  ];
  sched.setConditionalFormatRules(rules);
}

/**
 * Applies alternating row shading to the Schedule sheet.
 */
function applyScheduleAdvancedFormatting() {
  const ss    = SpreadsheetApp.getActive();
  const sched = ss.getSheetByName('Schedule');
  if (!sched) return;
  const lastRow = sched.getLastRow();
  for (let r = 2; r <= lastRow; r++) {
    if (r % 2 === 0) sched.getRange(r, 1, 1, 9).setBackground('#f8f9fa');
  }
}

/**
 * Applies conditional formatting to the Standings sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} std
 * @param {number} N - Number of players
 */
function applyStandingsFormatting(std, N) {
  if (N < 1) return;
  const dEnd = N + 1;
  const rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0).setBackground(CONFIG.COLORS.WIN).setBold(true)
      .setRanges([std.getRange(`G2:G${dEnd}`)]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0).setBackground(CONFIG.COLORS.LOSS)
      .setRanges([std.getRange(`G2:G${dEnd}`)]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .setGradientMaxpointWithValue(CONFIG.COLORS.SUCCESS, SpreadsheetApp.InterpolationType.NUMBER, '1')
      .setGradientMidpointWithValue('#ffffff', SpreadsheetApp.InterpolationType.NUMBER, '0.5')
      .setGradientMinpointWithValue(CONFIG.COLORS.DANGER, SpreadsheetApp.InterpolationType.NUMBER, '0')
      .setRanges([std.getRange(`I2:I${dEnd}`)]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberEqualTo(1).setBackground(CONFIG.COLORS.GOLD).setBold(true)
      .setRanges([std.getRange(`H2:H${dEnd}`)]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberEqualTo(2).setBackground(CONFIG.COLORS.SILVER).setBold(true)
      .setRanges([std.getRange(`H2:H${dEnd}`)]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberEqualTo(3).setBackground(CONFIG.COLORS.BRONZE).setBold(true)
      .setRanges([std.getRange(`H2:H${dEnd}`)]).build()
  ];
  std.setConditionalFormatRules(rules);
}

/**
 * Highlights the top 3 rows in the Standings sheet.
 */
function applyStandingsAdvancedFormatting() {
  const ss  = SpreadsheetApp.getActive();
  const std = ss.getSheetByName('Standings');
  if (!std) return;
  const N = getActivePlayerNames(ss).length;
  if (N >= 1) std.getRange(`A2:K2`).setBackground('#fff9c4').setFontWeight('bold');
  if (N >= 2) std.getRange(`A3:K3`).setBackground('#f5f5f5').setFontWeight('bold');
  if (N >= 3) std.getRange(`A4:K4`).setBackground('#ffe0b2');
}

/**
 * Applies WIN/LOSS/Team conditional formatting to the Partnerships sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} part
 * @param {number} totalRows - Total data rows (M * 2)
 */
function applyPartnershipsFormatting(part, totalRows) {
  if (totalRows < 1) return;
  const dEnd = totalRows + 1;
  const rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('WIN').setBackground(CONFIG.COLORS.WIN).setBold(true)
      .setRanges([part.getRange(`E2:E${dEnd}`)]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('LOSS').setBackground(CONFIG.COLORS.LOSS)
      .setRanges([part.getRange(`E2:E${dEnd}`)]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Team A').setBackground('#e3f2fd')
      .setRanges([part.getRange(`B2:B${dEnd}`)]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Team B').setBackground('#fce4ec')
      .setRanges([part.getRange(`B2:B${dEnd}`)]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .setGradientMaxpointWithValue(CONFIG.COLORS.SUCCESS, SpreadsheetApp.InterpolationType.NUMBER, '15')
      .setGradientMidpointWithValue('#ffffff', SpreadsheetApp.InterpolationType.NUMBER, '0')
      .setGradientMinpointWithValue(CONFIG.COLORS.DANGER, SpreadsheetApp.InterpolationType.NUMBER, '-15')
      .setRanges([part.getRange(`F2:F${dEnd}`)]).build()
  ];
  part.setConditionalFormatRules(rules);
}

/**
 * Orchestrates all advanced formatting passes.
 */
function applyAdvancedFormatting() {
  applyScheduleAdvancedFormatting();
  applyStandingsAdvancedFormatting();
}

// ============================================================================
// PLAYOFF BRACKET
// ============================================================================

/**
 * Generates a single-elimination playoff bracket for the top 4 players.
 * Requires all regular-season matches to be complete.
 */
function generatePlayoffBracket() {
  const ui    = SpreadsheetApp.getUi();
  const ss    = SpreadsheetApp.getActive();
  const sched = ss.getSheetByName('Schedule');

  if (!sched) { ui.alert('Error', 'Schedule not found.', ui.ButtonSet.OK); return; }

  const M         = sched.getLastRow() - 1;
  const statusCol = M > 0 ? sched.getRange(2, 8, M, 1).getValues() : [];
  const completed = statusCol.filter(r => r[0] === '✓').length;

  if (completed < M) {
    ui.alert('Season Incomplete',
      `Only ${completed} of ${M} matches completed. Finish the regular season first!`,
      ui.ButtonSet.OK);
    return;
  }

  const response = ui.alert(
    '🎲 Generate Playoff Bracket',
    'Create a 4-player single-elimination bracket with the top 4 players?',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  const lb   = ss.getSheetByName('Leaderboard');
  const top4 = lb ? [2,3,4,5].map(r => String(lb.getRange(r, 2).getValue())) : [];

  const bracket = getOrCreateSheet(ss, 'Playoff Bracket');
  bracket.clear();

  bracket.getRange('A1:H1').merge().setValue('🏆 PLAYOFF BRACKET')
    .setFontSize(24).setFontWeight('bold').setBackground(CONFIG.COLORS.GOLD).setHorizontalAlignment('center');
  bracket.getRange('A2:H2').merge().setValue('Single Elimination — Top 4 Players')
    .setFontSize(12).setHorizontalAlignment('center').setFontColor('#666666');

  // Semifinal 1 (Seed 1 vs Seed 4)
  bracket.getRange('A5:C5').merge().setValue('🔹 SEMIFINAL 1')
    .setFontSize(14).setFontWeight('bold').setBackground('#e3f2fd').setHorizontalAlignment('center');
  bracket.getRange('A6').setValue('Seed 1:'); bracket.getRange('B6').setValue(top4[0] || '--').setFontWeight('bold');
  bracket.getRange('A7').setValue('Seed 4:'); bracket.getRange('B7').setValue(top4[3] || '--');
  bracket.getRange('A8').setValue('Score:');  bracket.getRange('B8:C8').merge();
  bracket.getRange('A9').setValue('Winner:'); bracket.getRange('B9:C9').merge().setBackground('#e8f5e9');

  // Semifinal 2 (Seed 2 vs Seed 3)
  bracket.getRange('E5:G5').merge().setValue('🔹 SEMIFINAL 2')
    .setFontSize(14).setFontWeight('bold').setBackground('#e3f2fd').setHorizontalAlignment('center');
  bracket.getRange('E6').setValue('Seed 2:'); bracket.getRange('F6').setValue(top4[1] || '--').setFontWeight('bold');
  bracket.getRange('E7').setValue('Seed 3:'); bracket.getRange('F7').setValue(top4[2] || '--');
  bracket.getRange('E8').setValue('Score:');  bracket.getRange('F8:G8').merge();
  bracket.getRange('E9').setValue('Winner:'); bracket.getRange('F9:G9').merge().setBackground('#e8f5e9');

  // Championship
  bracket.getRange('B12:F12').merge().setValue('🏆 CHAMPIONSHIP MATCH')
    .setFontSize(16).setFontWeight('bold').setBackground(CONFIG.COLORS.GOLD).setHorizontalAlignment('center');
  bracket.getRange('B14').setValue('SF1 Winner:'); bracket.getRange('C14:D14').merge().setFormula('=B9');
  bracket.getRange('B15').setValue('SF2 Winner:'); bracket.getRange('C15:D15').merge().setFormula('=F9');
  bracket.getRange('B16').setValue('Score:');      bracket.getRange('C16:D16').merge();
  bracket.getRange('B17').setValue('CHAMPION:');
  bracket.getRange('C17:D17').merge()
    .setBackground(CONFIG.COLORS.GOLD).setFontWeight('bold').setFontSize(14);

  bracket.setColumnWidths(1, 8, 120);

  logActivity('Playoff bracket generated');
  ui.alert('✅ Bracket Created', 'Playoff bracket ready! Enter results in the bracket cells.', ui.ButtonSet.OK);
  ss.setActiveSheet(bracket);
}

// ============================================================================
// UTILITY ACTIONS
// ============================================================================

/**
 * Forces a spreadsheet recalculation flush.
 */
function refreshAllData() {
  const ui = SpreadsheetApp.getUi();
  try {
    SpreadsheetApp.flush();
    logActivity('Data refreshed');
    ui.alert('✅ Refreshed', 'All data has been recalculated!', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', 'Could not refresh data: ' + e.message, ui.ButtonSet.OK);
  }
}

/**
 * Shows match scheduling instructions (stub for calendar integration).
 */
function scheduleNextMatch() {
  SpreadsheetApp.getUi().alert(
    '📅 Schedule Match',
    'Match scheduling:\n\n' +
    '• Set the Date column (column I) in the Schedule sheet\n' +
    '• Full calendar invite integration requires Google Calendar API\n\n' +
    'For now, coordinate matches manually and enter scores when complete.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Shows Excel export instructions.
 */
function exportToExcel() {
  SpreadsheetApp.getUi().alert(
    '📊 Export to Excel',
    'To export:\n\n' +
    'File → Download → Microsoft Excel (.xlsx)\n\n' +
    'All sheets and formatting will be preserved.\n' +
    'Note: Charts may need to be regenerated in Excel.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}
