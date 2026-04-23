// ============================================================================
// Dashboard.gs
// Stats Dashboard, Player Profiles, Live Tracker, and Chart generation.
// Depends on: Config.gs, CoreEngine.gs
// ============================================================================

// ============================================================================
// STATS DASHBOARD
// ============================================================================

/**
 * Creates the enhanced Statistics Dashboard with dynamic player/match counts.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss]
 */
function createEnhancedStatsDashboard(ss) {
  if (!ss) ss = SpreadsheetApp.getActive();

  const players = getActivePlayerNames(ss);
  const N       = players.length;
  const M       = generateRoundRobinSchedule(players).length;
  const mEnd    = M + 1;
  const nEnd    = N + 1;

  const stats = getOrCreateSheet(ss, 'Stats Dashboard');
  stats.clear();

  // Title
  stats.getRange('A1:H1').merge().setValue('📊 TOURNAMENT STATISTICS DASHBOARD')
    .setFontSize(20).setFontWeight('bold');
  stats.getRange('A2:H2').merge().setValue('Last Updated: ' + getTimestamp())
    .setFontSize(10).setFontColor('#666666');

  // Tournament Overview
  stats.getRange('A4:H4').merge().setValue('🏁 TOURNAMENT OVERVIEW')
    .setFontSize(14).setFontWeight('bold')
    .setBackground(CONFIG.COLORS.PRIMARY).setFontColor('#ffffff').setHorizontalAlignment('center');

  stats.getRange('A5:B5').merge().setValue('Total Matches:').setFontWeight('bold');
  stats.getRange('C5').setValue(M);

  stats.getRange('E5:F5').merge().setValue('Total Players:').setFontWeight('bold');
  stats.getRange('G5').setValue(N);

  stats.getRange('A6:B6').merge().setValue('Completed:').setFontWeight('bold');
  stats.getRange('C6').setFormula(`=COUNTIF(Schedule!H2:H${mEnd},"✓")`);

  stats.getRange('E6:F6').merge().setValue('Remaining:').setFontWeight('bold');
  stats.getRange('G6').setFormula(`=${M}-C6`);

  stats.getRange('A7:B7').merge().setValue('Progress:').setFontWeight('bold');
  stats.getRange('C7').setFormula(`=IF(${M}=0,0,C6/${M})`).setNumberFormat('0.0%');
  stats.getRange('D7:H7').merge()
    .setFormula(`=REPT("█",ROUND(C7*20,0))&REPT("░",20-ROUND(C7*20,0))`)
    .setFontFamily('Courier New').setFontSize(12).setBackground('#f5f5f5');

  // Current Leaders
  stats.getRange('A9:H9').merge().setValue('🏆 CURRENT LEADERS')
    .setFontSize(14).setFontWeight('bold')
    .setBackground(CONFIG.COLORS.GOLD).setHorizontalAlignment('center');

  stats.getRange('A10:B10').merge().setValue('1st Place:').setFontWeight('bold').setBackground(CONFIG.COLORS.GOLD);
  stats.getRange('C10:D10').merge().setFormula('=INDEX(Leaderboard!B:B,2)').setFontWeight('bold');
  stats.getRange('E10').setFormula('=INDEX(Leaderboard!C:C,2)&" pts"');

  stats.getRange('A11:B11').merge().setValue('2nd Place:').setFontWeight('bold').setBackground(CONFIG.COLORS.SILVER);
  stats.getRange('C11:D11').merge().setFormula('=INDEX(Leaderboard!B:B,3)');
  stats.getRange('E11').setFormula('=INDEX(Leaderboard!C:C,3)&" pts"');

  stats.getRange('A12:B12').merge().setValue('3rd Place:').setFontWeight('bold').setBackground(CONFIG.COLORS.BRONZE);
  stats.getRange('C12:D12').merge().setFormula('=INDEX(Leaderboard!B:B,4)');
  stats.getRange('E12').setFormula('=INDEX(Leaderboard!C:C,4)&" pts"');

  // Top Performers
  stats.getRange('A14:H14').merge().setValue('📈 TOP PERFORMERS')
    .setFontSize(14).setFontWeight('bold')
    .setBackground(CONFIG.COLORS.SUCCESS).setFontColor('#ffffff').setHorizontalAlignment('center');

  stats.getRange('A15:B15').merge().setValue('Most Wins:').setFontWeight('bold');
  stats.getRange('C15:D15').merge().setFormula(`=INDEX(SORT(Standings!A2:C${nEnd},3,FALSE),1,1)`);
  stats.getRange('E15').setFormula(`=MAX(Standings!C2:C${nEnd})&" wins"`);

  stats.getRange('A16:B16').merge().setValue('Best Win %:').setFontWeight('bold');
  stats.getRange('C16:D16').merge().setFormula(`=INDEX(SORT(Standings!A2:I${nEnd},9,FALSE),1,1)`);
  stats.getRange('E16').setFormula(`=TEXT(MAX(Standings!I2:I${nEnd}),"0.0%")`);

  stats.getRange('A17:B17').merge().setValue('Most Points Scored:').setFontWeight('bold');
  stats.getRange('C17:D17').merge().setFormula(`=INDEX(SORT(Standings!A2:E${nEnd},5,FALSE),1,1)`);
  stats.getRange('E17').setFormula(`=MAX(Standings!E2:E${nEnd})&" pts"`);

  stats.getRange('A18:B18').merge().setValue('Best Point Diff:').setFontWeight('bold');
  stats.getRange('C18:D18').merge().setFormula(`=INDEX(SORT(Standings!A2:G${nEnd},7,FALSE),1,1)`);
  stats.getRange('E18').setFormula(`=TEXT(MAX(Standings!G2:G${nEnd}),"+0;-0;0")`);

  stats.getRange('A19:B19').merge().setValue('Stingiest Defense:').setFontWeight('bold');
  stats.getRange('C19:D19').merge().setFormula(`=INDEX(SORT(Standings!A2:F${nEnd},6,TRUE),1,1)`);
  stats.getRange('E19').setFormula(`=MIN(Standings!F2:F${nEnd})&" allowed"`);

  // Tournament Totals
  stats.getRange('A21:H21').merge().setValue('🎯 TOURNAMENT STATISTICS')
    .setFontSize(14).setFontWeight('bold')
    .setBackground(CONFIG.COLORS.PRIMARY).setFontColor('#ffffff').setHorizontalAlignment('center');

  stats.getRange('A22:B22').merge().setValue('Total Points Scored:').setFontWeight('bold');
  stats.getRange('C22').setFormula(`=SUM(Schedule!F2:G${mEnd})`);

  stats.getRange('E22:F22').merge().setValue('Avg Per Match:').setFontWeight('bold');
  stats.getRange('G22').setFormula('=IF(C6>0,C22/(C6*2),0)').setNumberFormat('0.0');

  stats.getRange('A23:B23').merge().setValue('Highest Score:').setFontWeight('bold');
  stats.getRange('C23').setFormula(`=MAX(Schedule!F2:G${mEnd})`);

  stats.getRange('E23:F23').merge().setValue('Lowest Score:').setFontWeight('bold');
  stats.getRange('G23').setFormula(
    `=IF(C6>0,MIN(IF(Schedule!F2:F${mEnd}<>"",Schedule!F2:F${mEnd}),IF(Schedule!G2:G${mEnd}<>"",Schedule!G2:G${mEnd})),0)`
  );

  stats.getRange('A24:B24').merge().setValue('Avg Win Margin:').setFontWeight('bold');
  stats.getRange('C24').setFormula(
    `=IF(C6>0,AVERAGE(ABS(Schedule!F2:F${mEnd}-Schedule!G2:G${mEnd})),0)`
  ).setNumberFormat('0.0');

  stats.getRange('E24:F24').merge().setValue('Closest Match:').setFontWeight('bold');
  stats.getRange('G24').setFormula(
    `=IF(C6>0,MIN(ABS(Schedule!F2:F${mEnd}-Schedule!G2:G${mEnd})),0)`
  );

  stats.getRange('A25:B25').merge().setValue('Biggest Blowout:').setFontWeight('bold');
  stats.getRange('C25').setFormula(
    `=IF(C6>0,MAX(ABS(Schedule!F2:F${mEnd}-Schedule!G2:G${mEnd})),0)`
  );

  // Recent Results header
  stats.getRange('A27:H27').merge().setValue('🕒 RECENT RESULTS (Last 5)')
    .setFontSize(14).setFontWeight('bold')
    .setBackground(CONFIG.COLORS.WARNING).setHorizontalAlignment('center');
  stats.getRange('A28:E28').setValues([['Match','Matchup','Score','Winner','Margin']]);
  stats.getRange('A28:E28').setFontWeight('bold').setBackground('#f5f5f5');

  // Column widths
  stats.setColumnWidths(1, 2, 120);
  stats.setColumnWidths(3, 2, 150);
  stats.setColumnWidth(5, 150);
  stats.setColumnWidths(6, 3, 100);

  return stats;
}

// ============================================================================
// PLAYER PROFILES
// ============================================================================

/**
 * Creates individual player profile sections for every active player.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss]
 */
function createPlayerProfiles(ss) {
  if (!ss) ss = SpreadsheetApp.getActive();

  const players = getActivePlayers(ss);
  const N       = players.length;
  const nEnd    = N + 1;

  const profiles = getOrCreateSheet(ss, 'Player Profiles');
  profiles.clear();

  profiles.getRange('A1').setValue('🎮 INDIVIDUAL PLAYER PROFILES').setFontSize(18).setFontWeight('bold');

  let currentRow = 3;

  players.forEach((player, idx) => {
    const name = player.name;

    // Player header
    profiles.getRange(currentRow, 1, 1, 6).merge()
      .setValue(`${idx + 1}. ${name}`)
      .setFontSize(14).setFontWeight('bold')
      .setBackground(CONFIG.COLORS.PRIMARY).setFontColor('#ffffff');
    currentRow++;

    // Basic Stats
    profiles.getRange(currentRow, 1, 1, 6).merge()
      .setValue('📊 BASIC STATISTICS').setFontWeight('bold').setBackground('#e8f0fe');
    currentRow++;

    profiles.getRange(currentRow, 1).setValue('Current Rank:');
    profiles.getRange(currentRow, 2).setFormula(
      `=IFERROR(VLOOKUP("${name}",Standings!A2:H${nEnd},8,FALSE),"N/A")`
    );
    profiles.getRange(currentRow, 3).setValue('Total Points:');
    profiles.getRange(currentRow, 4).setFormula(
      `=IFERROR(VLOOKUP("${name}",Standings!A2:B${nEnd},2,FALSE),0)`
    );
    currentRow++;

    profiles.getRange(currentRow, 1).setValue('Record (W-L):');
    profiles.getRange(currentRow, 2).setFormula(
      `=IFERROR(VLOOKUP("${name}",Standings!A2:C${nEnd},3,FALSE)&"-"&VLOOKUP("${name}",Standings!A2:D${nEnd},4,FALSE),"0-0")`
    );
    profiles.getRange(currentRow, 3).setValue('Win %:');
    profiles.getRange(currentRow, 4).setFormula(
      `=IFERROR(VLOOKUP("${name}",Standings!A2:I${nEnd},9,FALSE),0)`
    ).setNumberFormat('0.0%');
    currentRow++;

    profiles.getRange(currentRow, 1).setValue('Points For:');
    profiles.getRange(currentRow, 2).setFormula(
      `=IFERROR(VLOOKUP("${name}",Standings!A2:E${nEnd},5,FALSE),0)`
    );
    profiles.getRange(currentRow, 3).setValue('Points Against:');
    profiles.getRange(currentRow, 4).setFormula(
      `=IFERROR(VLOOKUP("${name}",Standings!A2:F${nEnd},6,FALSE),0)`
    );
    currentRow++;

    profiles.getRange(currentRow, 1).setValue('Point Diff:');
    profiles.getRange(currentRow, 2).setFormula(
      `=IFERROR(VLOOKUP("${name}",Standings!A2:G${nEnd},7,FALSE),0)`
    );
    profiles.getRange(currentRow, 3).setValue('Avg Score / Match:');
    profiles.getRange(currentRow, 4).setFormula(
      `=IFERROR(VLOOKUP("${name}",Standings!A2:E${nEnd},5,FALSE)/MAX(1,VLOOKUP("${name}",Standings!A2:C${nEnd},3,FALSE)+VLOOKUP("${name}",Standings!A2:D${nEnd},4,FALSE)),0)`
    ).setNumberFormat('0.0');
    currentRow++;

    // Performance Metrics
    currentRow++;
    profiles.getRange(currentRow, 1, 1, 6).merge()
      .setValue('📈 PERFORMANCE METRICS').setFontWeight('bold').setBackground('#fff2cc');
    currentRow++;

    profiles.getRange(currentRow, 1).setValue('Status:');
    profiles.getRange(currentRow, 2).setFormula(
      `=IFERROR(IF(VLOOKUP("${name}",Standings!A2:I${nEnd},9,FALSE)>0.6,"🔥 HOT",IF(VLOOKUP("${name}",Standings!A2:I${nEnd},9,FALSE)<0.4,"🧊 COLD","➖ Neutral")),"N/A")`
    );
    profiles.getRange(currentRow, 3).setValue('Skill Rating:');
    profiles.getRange(currentRow, 4).setValue(player.skillRating || '--');
    currentRow++;

    // Spacer
    currentRow += 2;
  });

  profiles.setColumnWidth(1, 180);
  profiles.setColumnWidth(2, 120);
  profiles.setColumnWidth(3, 180);
  profiles.setColumnWidth(4, 120);
  profiles.setColumnWidth(5, 120);
  profiles.setColumnWidth(6, 120);

  return profiles;
}

// ============================================================================
// LIVE MATCH TRACKER
// ============================================================================

/**
 * Creates the Live Match Tracker display sheet.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss]
 */
function createLiveMatchTracker(ss) {
  if (!ss) ss = SpreadsheetApp.getActive();

  const live = getOrCreateSheet(ss, 'Live Tracker');
  live.clear();

  live.getRange('A1:F1').merge().setValue('🔴 LIVE MATCH TRACKER')
    .setFontSize(20).setFontWeight('bold')
    .setBackground('#ff0000').setFontColor('#ffffff').setHorizontalAlignment('center');
  live.getRange('A2:F2').merge().setValue('Real-time match monitoring and updates')
    .setFontSize(10).setFontColor('#666666').setHorizontalAlignment('center');

  live.getRange('A4:F4').merge().setValue('NOW PLAYING')
    .setFontSize(16).setFontWeight('bold').setBackground('#ffeb3b').setHorizontalAlignment('center');

  live.getRange('A6').setValue('Match #:');
  live.getRange('B6').setValue('--');
  live.getRange('D6').setValue('Status:');
  live.getRange('E6').setValue('⏳ Pending');

  live.getRange('A8:C8').merge().setValue('TEAM A')
    .setFontSize(14).setFontWeight('bold').setBackground('#e3f2fd').setHorizontalAlignment('center');
  live.getRange('D8:F8').merge().setValue('TEAM B')
    .setFontSize(14).setFontWeight('bold').setBackground('#fce4ec').setHorizontalAlignment('center');

  live.getRange('A10:C10').merge().setFontSize(48).setHorizontalAlignment('center')
    .setBorder(true, true, true, true, true, true).setValue('0');
  live.getRange('D10:F10').merge().setFontSize(48).setHorizontalAlignment('center')
    .setBorder(true, true, true, true, true, true).setValue('0');

  live.getRange('A12').setValue('Players:');
  live.getRange('B12:C12').merge().setValue('--');
  live.getRange('D12').setValue('Players:');
  live.getRange('E12:F12').merge().setValue('--');

  live.getRange('A15:F15').merge().setValue('📝 INSTRUCTIONS')
    .setFontSize(14).setFontWeight('bold')
    .setBackground(CONFIG.COLORS.PRIMARY).setFontColor('#ffffff').setHorizontalAlignment('center');
  live.getRange('A16:F16').merge().setValue('Use "Matches → Enter Score" to update match results');
  live.getRange('A17:F17').merge().setValue('This tracker shows the current / most recent match');

  live.getRange('A20:F20').merge().setValue('📊 RECENT ACTIVITY')
    .setFontSize(14).setFontWeight('bold')
    .setBackground(CONFIG.COLORS.SUCCESS).setFontColor('#ffffff').setHorizontalAlignment('center');
  live.getRange('A21:F21').setValues([['Time','Match','Winner','Score','Margin','Status']])
    .setFontWeight('bold').setBackground('#f5f5f5');

  live.setColumnWidths(1, 6, 120);
  live.setRowHeight(10, 80);

  return live;
}

/**
 * Navigates to the Live Tracker sheet, creating it if necessary.
 */
function showLiveMatchTracker() {
  const ss   = SpreadsheetApp.getActive();
  const live = ss.getSheetByName('Live Tracker');
  if (live) ss.setActiveSheet(live);
  else createLiveMatchTracker(ss);
}

// ============================================================================
// CHART GENERATION
// ============================================================================

/**
 * Generates all charts across Stats Dashboard and Partnership Compatibility sheets.
 */
function createAllCharts() {
  const ss = SpreadsheetApp.getActive();
  createStandingsCharts(ss);
  createStatsDashboardCharts(ss);
  createPartnershipCharts(ss);
  SpreadsheetApp.getUi().alert('✅ Charts Created', 'All visualization charts have been generated!', SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Creates the four core standings charts on the Stats Dashboard sheet.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function createStandingsCharts(ss) {
  const stats = ss.getSheetByName('Stats Dashboard');
  if (!stats) return;

  const players = getActivePlayerNames(ss);
  const N       = players.length;
  const nEnd    = N + 1; // e.g. 9 for 8 players

  // Clear existing charts
  stats.getCharts().forEach(c => stats.removeChart(c));

  const lb = ss.getSheetByName('Leaderboard');
  if (!lb) return;

  // Chart 1: Points Distribution (Donut)
  stats.insertChart(stats.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(lb.getRange(`B2:B${nEnd}`))
    .addRange(lb.getRange(`C2:C${nEnd}`))
    .setPosition(29, 1, 0, 0)
    .setOption('title', 'Points Distribution')
    .setOption('width', 450).setOption('height', 300)
    .setOption('pieHole', 0.4)
    .setOption('legend', { position: 'right' })
    .setOption('chartArea', { width: '90%', height: '80%' })
    .build());

  // Chart 2: Wins vs Losses (Column)
  stats.insertChart(stats.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(lb.getRange(`B1:B${nEnd}`))
    .addRange(lb.getRange(`D1:E${nEnd}`))
    .setPosition(29, 6, 0, 0)
    .setOption('title', 'Wins vs Losses')
    .setOption('width', 500).setOption('height', 300)
    .setOption('legend', { position: 'bottom' })
    .setOption('colors', [CONFIG.COLORS.SUCCESS, CONFIG.COLORS.DANGER])
    .setOption('chartArea', { width: '85%', height: '70%' })
    .build());

  // Chart 3: Point Differential (Bar)
  stats.insertChart(stats.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(lb.getRange(`B1:B${nEnd}`))
    .addRange(lb.getRange(`H1:H${nEnd}`))
    .setPosition(44, 1, 0, 0)
    .setOption('title', 'Point Differential')
    .setOption('width', 500).setOption('height', 350)
    .setOption('legend', { position: 'none' })
    .setOption('colors', [CONFIG.COLORS.PRIMARY])
    .setOption('chartArea', { width: '75%', height: '80%' })
    .build());

  // Chart 4: Win Percentage (Bar)
  stats.insertChart(stats.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(lb.getRange(`B1:B${nEnd}`))
    .addRange(lb.getRange(`J1:J${nEnd}`))
    .setPosition(44, 6, 0, 0)
    .setOption('title', 'Win Percentage')
    .setOption('width', 500).setOption('height', 350)
    .setOption('legend', { position: 'none' })
    .setOption('colors', [CONFIG.COLORS.SUCCESS])
    .setOption('hAxis', { format: '0%' })
    .setOption('chartArea', { width: '70%', height: '80%' })
    .build());
}

/**
 * Additional stat charts (extensible stub — add more as needed).
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function createStatsDashboardCharts(ss) {
  // Reserved for additional analytics charts
}

/**
 * Creates the partnership win-rate bar chart on the Partnership Compatibility sheet.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function createPartnershipCharts(ss) {
  const compat = ss.getSheetByName('Partnership Compatibility');
  if (!compat) return;

  compat.getCharts().forEach(c => compat.removeChart(c));

  const lastRow = compat.getLastRow();
  if (lastRow < 6) return;

  compat.insertChart(compat.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(compat.getRange(`A5:A${lastRow}`))
    .addRange(compat.getRange(`E5:E${lastRow}`))
    .setPosition(5, 9, 0, 0)
    .setOption('title', 'Partnership Win Rates')
    .setOption('width', 620).setOption('height', 420)
    .setOption('legend', { position: 'none' })
    .setOption('hAxis', { format: '0%', title: 'Win Rate' })
    .setOption('chartArea', { width: '70%', height: '85%' })
    .build());
}
