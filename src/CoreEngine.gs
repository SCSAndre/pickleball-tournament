// ============================================================================
// CoreEngine.gs
// Dynamic schedule generator, standings, leaderboard, partnerships, H2H, progress.
// Depends on: Config.gs
// ============================================================================

// ============================================================================
// ROUND-ROBIN SCHEDULE GENERATOR
// ============================================================================

/**
 * Generates all unique doubles matches via a greedy pair-matching algorithm.
 * For N players produces C(N,2)/2 = N*(N-1)/4 matches (exact for N divisible by 4).
 * Leftover pairs (odd groups) are skipped gracefully.
 * @param {string[]} players - Array of player names.
 * @returns {Array<[string,string,string,string]>} Array of [TeamA_P1, TeamA_P2, TeamB_P1, TeamB_P2]
 */
function generateRoundRobinSchedule(players) {
  const n = players.length;
  const pairs = [];

  // Generate all unique pairs C(n,2)
  for (let i = 0; i < n; i++) {
    for (let j = i + 1; j < n; j++) {
      pairs.push([players[i], players[j]]);
    }
  }

  const used    = new Array(pairs.length).fill(false);
  const matches = [];

  for (let i = 0; i < pairs.length; i++) {
    if (used[i]) continue;
    for (let j = i + 1; j < pairs.length; j++) {
      if (used[j]) continue;
      const a = pairs[i];
      const b = pairs[j];
      // No shared players between the two pairs
      if (a[0] !== b[0] && a[0] !== b[1] && a[1] !== b[0] && a[1] !== b[1]) {
        matches.push([a[0], a[1], b[0], b[1]]);
        used[i] = true;
        used[j] = true;
        break;
      }
    }
  }

  return matches;
}

// ============================================================================
// SCHEDULE SHEET
// ============================================================================

/**
 * Creates the Schedule sheet dynamically based on active players.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function createScheduleSheet(ss) {
  const players = getActivePlayerNames(ss);
  const matches = generateRoundRobinSchedule(players);
  const M       = matches.length;

  const sched = getOrCreateSheet(ss, 'Schedule');
  sched.clear();

  // Header
  sched.getRange('A1:I1').setValues([['Match','TeamA_P1','TeamA_P2','TeamB_P1','TeamB_P2','ScoreA','ScoreB','Status','Date']]);
  styleHeader(sched.getRange('A1:I1'));
  sched.setFrozenRows(1);

  // Match data rows
  const rows = matches.map((m, idx) => [idx + 1, m[0], m[1], m[2], m[3], '', '', '⏳', '']);
  if (rows.length > 0) {
    sched.getRange(2, 1, rows.length, 9).setValues(rows);
  }

  // Status formula: auto-marks ✓ when both scores are numbers
  for (let r = 2; r <= M + 1; r++) {
    sched.getRange(r, 8).setFormula(
      `=IF(AND(ISNUMBER(F${r}),ISNUMBER(G${r})),"✓","⏳")`
    );
  }

  // Score validation
  const scoreVal = SpreadsheetApp.newDataValidation()
    .requireNumberGreaterThanOrEqualTo(0)
    .setAllowInvalid(false)
    .setHelpText('Enter a non-negative integer score')
    .build();
  if (M > 0) sched.getRange(2, 6, M, 2).setDataValidation(scoreVal);

  sched.autoResizeColumns(1, 9);
  sched.getRange(`F2:G${M + 1}`).setNumberFormat('0');
  sched.setColumnWidth(8, 65);

  applyScheduleFormatting(sched, M);

  logActivity(`Schedule created: ${M} matches for ${players.length} players`);
}

// ============================================================================
// STANDINGS SHEET
// ============================================================================

/**
 * Creates the Standings sheet with fully dynamic SUMPRODUCT formulas.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function createStandingsSheet(ss) {
  const players = getActivePlayerNames(ss);
  const N       = players.length;
  const M       = generateRoundRobinSchedule(players).length;
  const mEnd    = M + 1; // last data row in Schedule

  const std = getOrCreateSheet(ss, 'Standings');
  std.clear();

  std.getRange('A1:K1').setValues([['Player','Points','Wins','Losses','PF','PA','Diff','Rank','Win%','Form','Streak']]);
  styleHeader(std.getRange('A1:K1'));
  std.setFrozenRows(1);

  // Player names
  std.getRange(2, 1, N, 1).setValues(players.map(p => [p]));

  for (let r = 2; r <= N + 1; r++) {
    const sB = `Schedule!B$2:B$${mEnd}`;
    const sC = `Schedule!C$2:C$${mEnd}`;
    const sD = `Schedule!D$2:D$${mEnd}`;
    const sE = `Schedule!E$2:E$${mEnd}`;
    const sF = `Schedule!F$2:F$${mEnd}`;
    const sG = `Schedule!G$2:G$${mEnd}`;
    const A  = `A${r}`;
    const pRange = `$B$2:$B$${N + 1}`;
    const gRange = `$G$2:$G$${N + 1}`;

    // Points (2 per win)
    std.getRange(r, 2).setFormula(
      `=2*(SUMPRODUCT(((${sB}=${A})+(${sC}=${A}))*(${sF}>${sG}))+SUMPRODUCT(((${sD}=${A})+(${sE}=${A}))*(${sG}>${sF})))`
    );
    // Wins
    std.getRange(r, 3).setFormula(
      `=SUMPRODUCT(((${sB}=${A})+(${sC}=${A}))*(${sF}>${sG}))+SUMPRODUCT(((${sD}=${A})+(${sE}=${A}))*(${sG}>${sF}))`
    );
    // Losses
    std.getRange(r, 4).setFormula(
      `=SUMPRODUCT(((${sB}=${A})+(${sC}=${A}))*(${sF}<${sG}))+SUMPRODUCT(((${sD}=${A})+(${sE}=${A}))*(${sG}<${sF}))`
    );
    // Points For
    std.getRange(r, 5).setFormula(
      `=SUMPRODUCT(((${sB}=${A})+(${sC}=${A}))*(${sF}))+SUMPRODUCT(((${sD}=${A})+(${sE}=${A}))*(${sG}))`
    );
    // Points Against
    std.getRange(r, 6).setFormula(
      `=SUMPRODUCT(((${sB}=${A})+(${sC}=${A}))*(${sG}))+SUMPRODUCT(((${sD}=${A})+(${sE}=${A}))*(${sF}))`
    );
    // Diff
    std.getRange(r, 7).setFormula(`=E${r}-F${r}`);
    // Rank (tie-broken by point diff)
    std.getRange(r, 8).setFormula(
      `=RANK(B${r},${pRange},0)+COUNTIFS(${pRange},B${r},${gRange},">"&G${r})`
    );
    // Win %
    std.getRange(r, 9).setFormula(
      `=IF(C${r}+D${r}=0,0,C${r}/(C${r}+D${r}))`
    );
    // Form & Streak (script-calculated, placeholder)
    std.getRange(r, 10).setValue('--');
    std.getRange(r, 11).setValue('--');
  }

  std.autoResizeColumns(1, 11);
  if (N > 0) {
    std.getRange(`B2:H${N + 1}`).setNumberFormat('0');
    std.getRange(`I2:I${N + 1}`).setNumberFormat('0.0%');
  }

  applyStandingsFormatting(std, N);
}

// ============================================================================
// LEADERBOARD SHEET
// ============================================================================

/**
 * Creates the Leaderboard sheet — auto-sorted standings with podium medals.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function createLeaderboardSheet(ss) {
  const players = getActivePlayerNames(ss);
  const N       = players.length;
  const stdEnd  = N + 1; // last data row in Standings

  const lb = getOrCreateSheet(ss, 'Leaderboard');
  lb.clear();

  lb.getRange('A1:L1').setValues([['Pos','Player','Points','Wins','Losses','PF','PA','Diff','Rank','Win%','Form','Medal']]);
  styleHeader(lb.getRange('A1:L1'));
  lb.setFrozenRows(1);

  // Position numbers
  for (let i = 1; i <= N; i++) lb.getRange(i + 1, 1).setValue(i);

  // Dynamic SORT formula — picks up however many rows Standings has
  lb.getRange('B2').setFormula(
    `=SORT(Standings!A2:K${stdEnd},Standings!B2:B${stdEnd},FALSE,Standings!G2:G${stdEnd},FALSE)`
  );

  // Medal emojis for top 3
  if (N >= 1) lb.getRange('L2').setValue('🥇').setFontSize(18);
  if (N >= 2) lb.getRange('L3').setValue('🥈').setFontSize(18);
  if (N >= 3) lb.getRange('L4').setValue('🥉').setFontSize(18);

  // Podium row colors
  if (N >= 1) lb.getRange(`A2:L2`).setBackground(CONFIG.COLORS.GOLD).setFontWeight('bold');
  if (N >= 2) lb.getRange(`A3:L3`).setBackground(CONFIG.COLORS.SILVER).setFontWeight('bold');
  if (N >= 3) lb.getRange(`A4:L4`).setBackground(CONFIG.COLORS.BRONZE).setFontWeight('bold');

  lb.autoResizeColumns(1, 12);
  lb.setColumnWidth(12, 55);
}

// ============================================================================
// PARTNERSHIPS SHEET
// ============================================================================

/**
 * Creates the Partnerships sheet with one row per team per match.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function createPartnershipsSheet(ss) {
  const players = getActivePlayerNames(ss);
  const M       = generateRoundRobinSchedule(players).length;

  const part = getOrCreateSheet(ss, 'Partnerships');
  part.clear();

  part.getRange('A1:G1').setValues([['Match','Team','Partnership','Scored','Result','Margin','Quality']]);
  styleHeader(part.getRange('A1:G1'));
  part.setFrozenRows(1);

  let row = 2;
  for (let r = 2; r <= M + 1; r++) {
    // Team A row
    part.getRange(row, 1).setFormula(`=Schedule!A${r}`);
    part.getRange(row, 2).setValue('Team A');
    part.getRange(row, 3).setFormula(`=Schedule!B${r}&" + "&Schedule!C${r}`);
    part.getRange(row, 4).setFormula(`=Schedule!F${r}`);
    part.getRange(row, 5).setFormula(
      `=IF(Schedule!F${r}>Schedule!G${r},"WIN",IF(Schedule!F${r}<Schedule!G${r},"LOSS",IF(AND(ISNUMBER(Schedule!F${r}),ISNUMBER(Schedule!G${r})),"TIE","")))`
    );
    part.getRange(row, 6).setFormula(`=Schedule!F${r}-Schedule!G${r}`);
    part.getRange(row, 7).setFormula(
      `=IF(E${row}="WIN",IF(F${row}>10,"⭐⭐⭐",IF(F${row}>5,"⭐⭐","⭐")),"")`
    );
    row++;

    // Team B row
    part.getRange(row, 1).setFormula(`=Schedule!A${r}`);
    part.getRange(row, 2).setValue('Team B');
    part.getRange(row, 3).setFormula(`=Schedule!D${r}&" + "&Schedule!E${r}`);
    part.getRange(row, 4).setFormula(`=Schedule!G${r}`);
    part.getRange(row, 5).setFormula(
      `=IF(Schedule!G${r}>Schedule!F${r},"WIN",IF(Schedule!G${r}<Schedule!F${r},"LOSS",IF(AND(ISNUMBER(Schedule!F${r}),ISNUMBER(Schedule!G${r})),"TIE","")))`
    );
    part.getRange(row, 6).setFormula(`=Schedule!G${r}-Schedule!F${r}`);
    part.getRange(row, 7).setFormula(
      `=IF(E${row}="WIN",IF(F${row}>10,"⭐⭐⭐",IF(F${row}>5,"⭐⭐","⭐")),"")`
    );
    row++;
  }

  part.autoResizeColumns(1, 7);
  applyPartnershipsFormatting(part, M * 2);
}

// ============================================================================
// HEAD-TO-HEAD SHEET
// ============================================================================

/**
 * Creates the Head-to-Head analysis sheet with dynamic player dropdowns.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function createHeadToHeadSheet(ss) {
  const players = getActivePlayerNames(ss);
  const M       = generateRoundRobinSchedule(players).length;
  const mEnd    = M + 1;

  const h2h = getOrCreateSheet(ss, 'Head-to-Head');
  h2h.clear();

  h2h.getRange('A1').setValue('🔍 HEAD-TO-HEAD ANALYSIS').setFontSize(18).setFontWeight('bold');
  h2h.getRange('A3').setValue('Player 1:').setFontWeight('bold');
  h2h.getRange('A4').setValue('Player 2:').setFontWeight('bold');

  const validation = SpreadsheetApp.newDataValidation().requireValueInList(players).build();
  h2h.getRange('B3').setDataValidation(validation);
  h2h.getRange('B4').setDataValidation(validation);

  // As Partners section
  h2h.getRange('A6').setValue('🤝 AS PARTNERS').setFontSize(14).setFontWeight('bold').setBackground('#e8f0fe');
  h2h.getRange('A7:B7').setValues([['Metric','Value']]).setFontWeight('bold').setBackground('#d0e1f9');
  h2h.getRange('A8:A13').setValues([
    ['Matches Together'],['Record (W-L)'],['Win Rate'],['Points Scored'],['Points Allowed'],['Avg Margin']
  ]);

  // As Opponents section
  h2h.getRange('A15').setValue('⚔️ AS OPPONENTS').setFontSize(14).setFontWeight('bold').setBackground('#fce8e6');
  h2h.getRange('A16:B16').setValues([['Metric','Value']]).setFontWeight('bold').setBackground('#f4cccc');
  h2h.getRange('A17:A22').setValues([
    ['Matches Faced'],['Player 1 Wins'],['Player 2 Wins'],
    ['Player 1 Avg Score'],['Player 2 Avg Score'],['Head-to-Head Leader']
  ]);

  // Dynamic schedule ranges for partner formulas
  const sB = `Schedule!B$2:B$${mEnd}`;
  const sC = `Schedule!C$2:C$${mEnd}`;
  const sD = `Schedule!D$2:D$${mEnd}`;
  const sE = `Schedule!E$2:E$${mEnd}`;
  const sF = `Schedule!F$2:F$${mEnd}`;
  const sG = `Schedule!G$2:G$${mEnd}`;

  // Partner formulas
  h2h.getRange('B8').setFormula(
    `=IF(OR(ISBLANK(B3),ISBLANK(B4)),"",SUMPRODUCT(((${sB}=B3)*(${sC}=B4))+((${sB}=B4)*(${sC}=B3))+((${sD}=B3)*(${sE}=B4))+((${sD}=B4)*(${sE}=B3))))`
  );
  h2h.getRange('B9').setFormula(
    `=IF(OR(ISBLANK(B3),ISBLANK(B4)),"",SUMPRODUCT(((${sB}=B3)*(${sC}=B4)+(${sB}=B4)*(${sC}=B3))*(${sF}>${sG}))+SUMPRODUCT(((${sD}=B3)*(${sE}=B4)+(${sD}=B4)*(${sE}=B3))*(${sG}>${sF}))&"-"&SUMPRODUCT(((${sB}=B3)*(${sC}=B4)+(${sB}=B4)*(${sC}=B3))*(${sF}<${sG}))+SUMPRODUCT(((${sD}=B3)*(${sE}=B4)+(${sD}=B4)*(${sE}=B3))*(${sG}<${sF})))`
  );
  h2h.getRange('B10').setFormula(
    `=IF(OR(ISBLANK(B3),ISBLANK(B4),B8=0),"",TEXT((SUMPRODUCT(((${sB}=B3)*(${sC}=B4)+(${sB}=B4)*(${sC}=B3))*(${sF}>${sG}))+SUMPRODUCT(((${sD}=B3)*(${sE}=B4)+(${sD}=B4)*(${sE}=B3))*(${sG}>${sF})))/B8,"0.0%"))`
  );
  h2h.getRange('B11').setFormula(
    `=IF(OR(ISBLANK(B3),ISBLANK(B4)),"",SUMPRODUCT(((${sB}=B3)*(${sC}=B4)+(${sB}=B4)*(${sC}=B3))*(${sF}))+SUMPRODUCT(((${sD}=B3)*(${sE}=B4)+(${sD}=B4)*(${sE}=B3))*(${sG})))`
  );
  h2h.getRange('B12').setFormula(
    `=IF(OR(ISBLANK(B3),ISBLANK(B4)),"",SUMPRODUCT(((${sB}=B3)*(${sC}=B4)+(${sB}=B4)*(${sC}=B3))*(${sG}))+SUMPRODUCT(((${sD}=B3)*(${sE}=B4)+(${sD}=B4)*(${sE}=B3))*(${sF})))`
  );
  h2h.getRange('B13').setFormula(`=IF(OR(ISBLANK(B3),ISBLANK(B4),B8=0),"",(B11-B12)/B8)`);

  // Opponent formulas
  h2h.getRange('B17').setFormula(
    `=IF(OR(ISBLANK(B3),ISBLANK(B4)),"",SUMPRODUCT(((${sB}=B3)+(${sC}=B3))*((${sD}=B4)+(${sE}=B4)))+SUMPRODUCT(((${sD}=B3)+(${sE}=B3))*((${sB}=B4)+(${sC}=B4))))`
  );
  h2h.getRange('B18').setFormula(
    `=IF(OR(ISBLANK(B3),ISBLANK(B4)),"",SUMPRODUCT(((${sB}=B3)+(${sC}=B3))*((${sD}=B4)+(${sE}=B4))*(${sF}>${sG}))+SUMPRODUCT(((${sD}=B3)+(${sE}=B3))*((${sB}=B4)+(${sC}=B4))*(${sG}>${sF})))`
  );
  h2h.getRange('B19').setFormula(
    `=IF(OR(ISBLANK(B3),ISBLANK(B4)),"",SUMPRODUCT(((${sB}=B4)+(${sC}=B4))*((${sD}=B3)+(${sE}=B3))*(${sF}>${sG}))+SUMPRODUCT(((${sD}=B4)+(${sE}=B4))*((${sB}=B3)+(${sC}=B3))*(${sG}>${sF})))`
  );
  h2h.getRange('B20').setFormula(
    `=IF(OR(ISBLANK(B3),ISBLANK(B4),B17=0),"",ROUND((SUMPRODUCT(((${sB}=B3)+(${sC}=B3))*((${sD}=B4)+(${sE}=B4))*(${sF}))+SUMPRODUCT(((${sD}=B3)+(${sE}=B3))*((${sB}=B4)+(${sC}=B4))*(${sG})))/B17,1))`
  );
  h2h.getRange('B21').setFormula(
    `=IF(OR(ISBLANK(B3),ISBLANK(B4),B17=0),"",ROUND((SUMPRODUCT(((${sB}=B4)+(${sC}=B4))*((${sD}=B3)+(${sE}=B3))*(${sF}))+SUMPRODUCT(((${sD}=B4)+(${sE}=B4))*((${sB}=B3)+(${sC}=B3))*(${sG})))/B17,1))`
  );
  h2h.getRange('B22').setFormula(
    `=IF(OR(ISBLANK(B3),ISBLANK(B4)),"",IF(B18>B19,B3&" leads "&B18&"-"&B19,IF(B19>B18,B4&" leads "&B19&"-"&B18,"Tied "&B18&"-"&B19)))`
  );

  h2h.setColumnWidth(1, 200);
  h2h.setColumnWidth(2, 180);
}

// ============================================================================
// PROGRESS TRACKER SHEET
// ============================================================================

/**
 * Creates the Progress Tracker sheet — dynamic row count based on match count.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function createProgressSheet(ss) {
  const players = getActivePlayerNames(ss);
  const M       = generateRoundRobinSchedule(players).length;

  const prog = getOrCreateSheet(ss, 'Progress Tracker');
  prog.clear();

  prog.getRange('A1').setValue('⏱️ MATCH PROGRESS TRACKER').setFontSize(16).setFontWeight('bold');
  prog.getRange('A3:F3').setValues([['Match','Matchup','Status','Completion','Date','Notes']]);
  styleHeader(prog.getRange('A3:F3'));
  prog.setFrozenRows(3);

  for (let r = 2; r <= M + 1; r++) {
    const row = r + 2;
    prog.getRange(row, 1).setFormula(`=Schedule!A${r}`);
    prog.getRange(row, 2).setFormula(
      `=Schedule!B${r}&" + "&Schedule!C${r}&" vs "&Schedule!D${r}&" + "&Schedule!E${r}`
    );
    prog.getRange(row, 3).setFormula(`=Schedule!H${r}`);
    prog.getRange(row, 4).setFormula(`=IF(Schedule!H${r}="✓","100%","0%")`);
  }

  prog.autoResizeColumns(1, 6);
  prog.setColumnWidth(2, 350);
  prog.setColumnWidth(6, 200);

  if (M > 0) {
    const completeRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('✓').setBackground(CONFIG.COLORS.WIN)
      .setRanges([prog.getRange(`C4:C${M + 3}`)]).build();
    const pendingRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('⏳').setBackground(CONFIG.COLORS.PENDING)
      .setRanges([prog.getRange(`C4:C${M + 3}`)]).build();
    prog.setConditionalFormatRules([completeRule, pendingRule]);
  }
}
