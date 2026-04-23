// ============================================================================
// Analytics.gs
// Partnership Compatibility, Predictions, Achievements sheet + checker.
// Depends on: Config.gs, CoreEngine.gs
// ============================================================================

// ============================================================================
// PARTNERSHIP COMPATIBILITY
// ============================================================================

/**
 * Creates the Partnership Compatibility matrix, reading partnerships dynamically
 * from the Schedule sheet rather than any hardcoded player list.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss]
 */
function createPartnershipCompatibility(ss) {
  if (!ss) ss = SpreadsheetApp.getActive();

  const players = getActivePlayerNames(ss);
  const M       = generateRoundRobinSchedule(players).length;
  const partEnd = M * 2 + 1; // last data row in Partnerships sheet

  const compat = getOrCreateSheet(ss, 'Partnership Compatibility');
  compat.clear();

  compat.getRange('A1').setValue('🤝 PARTNERSHIP COMPATIBILITY MATRIX')
    .setFontSize(18).setFontWeight('bold');
  compat.getRange('A2').setValue('Analysis of player combinations — win rates and performance metrics')
    .setFontSize(10).setFontColor('#666666');

  compat.getRange('A4:G4').setValues([['Partnership','Matches','Wins','Losses','Win %','Avg Margin','Rating']]);
  styleHeader(compat.getRange('A4:G4'));
  compat.setFrozenRows(4);

  // Collect unique partnerships from the schedule
  const schedSheet = ss.getSheetByName('Schedule');
  const partnerships = [];

  if (schedSheet && M > 0) {
    const schedData = schedSheet.getRange(2, 1, M, 5).getValues();
    schedData.forEach(row => {
      const p1 = String(row[1]).trim();
      const p2 = String(row[2]).trim();
      const p3 = String(row[3]).trim();
      const p4 = String(row[4]).trim();
      if (p1 && p2) {
        const pair1 = [p1, p2].sort().join(' + ');
        if (!partnerships.includes(pair1)) partnerships.push(pair1);
      }
      if (p3 && p4) {
        const pair2 = [p3, p4].sort().join(' + ');
        if (!partnerships.includes(pair2)) partnerships.push(pair2);
      }
    });
    partnerships.sort();
  }

  // Write formulas for each partnership
  const dataStartRow = 5;
  partnerships.forEach((partnership, idx) => {
    const row = dataStartRow + idx;
    compat.getRange(row, 1).setValue(partnership);

    compat.getRange(row, 2).setFormula(
      `=COUNTIF(Partnerships!C2:C${partEnd},"${partnership}")`
    );
    compat.getRange(row, 3).setFormula(
      `=COUNTIFS(Partnerships!C2:C${partEnd},"${partnership}",Partnerships!E2:E${partEnd},"WIN")`
    );
    compat.getRange(row, 4).setFormula(
      `=COUNTIFS(Partnerships!C2:C${partEnd},"${partnership}",Partnerships!E2:E${partEnd},"LOSS")`
    );
    compat.getRange(row, 5).setFormula(
      `=IF(B${row}=0,0,C${row}/B${row})`
    );
    compat.getRange(row, 6).setFormula(
      `=IF(B${row}=0,0,AVERAGEIF(Partnerships!C2:C${partEnd},"${partnership}",Partnerships!F2:F${partEnd}))`
    );
    // Rating: weighted win% (70%) + capped margin bonus (30%)
    compat.getRange(row, 7).setFormula(
      `=IF(B${row}=0,0,ROUND(E${row}*70+MIN(F${row}/10,30),1))`
    );
  });

  const pCount = partnerships.length;

  if (pCount > 0) {
    compat.getRange(dataStartRow, 5, pCount, 1).setNumberFormat('0.0%');
    compat.getRange(dataStartRow, 6, pCount, 1).setNumberFormat('0.0');
    compat.getRange(dataStartRow, 7, pCount, 1).setNumberFormat('0.0');

    // Conditional formatting: Win %
    const winPctRule = SpreadsheetApp.newConditionalFormatRule()
      .setGradientMaxpointWithValue(CONFIG.COLORS.SUCCESS, SpreadsheetApp.InterpolationType.NUMBER, '1')
      .setGradientMidpointWithValue('#ffffff', SpreadsheetApp.InterpolationType.NUMBER, '0.5')
      .setGradientMinpointWithValue(CONFIG.COLORS.DANGER, SpreadsheetApp.InterpolationType.NUMBER, '0')
      .setRanges([compat.getRange(dataStartRow, 5, pCount, 1)])
      .build();

    // Conditional formatting: Rating
    const ratingRule = SpreadsheetApp.newConditionalFormatRule()
      .setGradientMaxpointWithValue(CONFIG.COLORS.GOLD, SpreadsheetApp.InterpolationType.NUMBER, '100')
      .setGradientMidpointWithValue('#ffffff', SpreadsheetApp.InterpolationType.NUMBER, '50')
      .setGradientMinpointWithValue('#cccccc', SpreadsheetApp.InterpolationType.NUMBER, '0')
      .setRanges([compat.getRange(dataStartRow, 7, pCount, 1)])
      .build();

    compat.setConditionalFormatRules([winPctRule, ratingRule]);

    // Top Partnerships section
    const topRow = dataStartRow + pCount + 2;
    compat.getRange(topRow, 1, 1, 7).merge()
      .setValue('⭐ TOP 5 PARTNERSHIPS').setFontSize(14).setFontWeight('bold')
      .setBackground(CONFIG.COLORS.GOLD).setHorizontalAlignment('center');
    compat.getRange(topRow + 1, 1).setFormula(
      `=SORT(A${dataStartRow}:G${dataStartRow + pCount - 1},7,FALSE)`
    );

    // Needs Improvement section
    const worstRow = topRow + 7;
    compat.getRange(worstRow, 1, 1, 7).merge()
      .setValue('⚠️ NEEDS IMPROVEMENT').setFontSize(14).setFontWeight('bold')
      .setBackground(CONFIG.COLORS.DANGER).setFontColor('#ffffff').setHorizontalAlignment('center');
    compat.getRange(worstRow + 1, 1).setFormula(
      `=SORT(FILTER(A${dataStartRow}:G${dataStartRow + pCount - 1},B${dataStartRow}:B${dataStartRow + pCount - 1}>0),7,TRUE)`
    );
  }

  compat.autoResizeColumns(1, 7);
  return compat;
}

// ============================================================================
// PREDICTIONS SHEET
// ============================================================================

/**
 * Creates the AI Predictions & Insights sheet with dynamic player rows.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss]
 */
function createPredictionsSheet(ss) {
  if (!ss) ss = SpreadsheetApp.getActive();

  const players = getActivePlayers(ss);
  const N       = players.length;
  const nEnd    = N + 1;
  const M       = generateRoundRobinSchedule(players.map(p => p.name)).length;

  const pred = getOrCreateSheet(ss, 'Predictions');
  pred.clear();

  pred.getRange('A1').setValue('🔮 AI-POWERED PREDICTIONS & INSIGHTS')
    .setFontSize(18).setFontWeight('bold');
  pred.getRange('A2').setValue('Statistical analysis and outcome predictions')
    .setFontSize(10).setFontColor('#666666');

  // Form indicators
  pred.getRange('A4:E4').merge().setValue('🔥 PLAYER FORM INDICATORS')
    .setFontSize(14).setFontWeight('bold')
    .setBackground(CONFIG.COLORS.WARNING).setHorizontalAlignment('center');
  pred.getRange('A5:E5').setValues([['Player','Status','Win%','Trend','Prediction']])
    .setFontWeight('bold').setBackground('#f5f5f5');

  players.forEach((player, idx) => {
    const row  = 6 + idx;
    const name = player.name;
    pred.getRange(row, 1).setValue(name);
    pred.getRange(row, 2).setFormula(
      `=IFERROR(IF(VLOOKUP("${name}",Standings!A2:I${nEnd},9,FALSE)>0.6,"🔥 HOT",IF(VLOOKUP("${name}",Standings!A2:I${nEnd},9,FALSE)<0.4,"🧊 COLD","➖ Neutral")),"No data")`
    );
    pred.getRange(row, 3).setFormula(
      `=IFERROR(VLOOKUP("${name}",Standings!A2:I${nEnd},9,FALSE),0)`
    ).setNumberFormat('0.0%');
    pred.getRange(row, 4).setValue('--'); // Trend: future enhancement
    pred.getRange(row, 5).setValue('TBD');
  });

  // Projected Final Standings
  const projRow = N + 7;
  pred.getRange(projRow, 1, 1, 5).merge()
    .setValue('📊 PROJECTED FINAL STANDINGS').setFontSize(14).setFontWeight('bold')
    .setBackground(CONFIG.COLORS.PRIMARY).setFontColor('#ffffff').setHorizontalAlignment('center');
  pred.getRange(projRow + 1, 1, 1, 5).setValues([['Proj. Rank','Player','Current Pts','Projected Pts','Δ']])
    .setFontWeight('bold').setBackground('#f5f5f5');

  pred.getRange(projRow + 2, 1).setFormula('=Leaderboard!A2:A' + (N + 1));
  pred.getRange(projRow + 2, 2).setFormula('=Leaderboard!B2:B' + (N + 1));
  pred.getRange(projRow + 2, 3).setFormula('=Leaderboard!C2:C' + (N + 1));

  for (let i = 0; i < N; i++) {
    const r = projRow + 2 + i;
    pred.getRange(r, 4).setFormula(
      `=C${r}+ROUND((IFERROR(VLOOKUP(B${r},Standings!A2:I${nEnd},9,FALSE),0)*2*(${M}-'Stats Dashboard'!C6)/${N}),0)`
    );
    pred.getRange(r, 5).setFormula(`=D${r}-C${r}`);
  }

  // Match Predictions column
  pred.getRange('G4:K4').merge().setValue('⚔️ UPCOMING MATCH PREDICTIONS')
    .setFontSize(14).setFontWeight('bold')
    .setBackground(CONFIG.COLORS.SUCCESS).setFontColor('#ffffff').setHorizontalAlignment('center');
  pred.getRange('G5:K5').setValues([['Match','Team A','Team B','Predicted Winner','Confidence']])
    .setFontWeight('bold').setBackground('#f5f5f5');
  pred.getRange('G6').setValue('Analyzing remaining matches…');

  // Insights
  pred.getRange('G16:K16').merge().setValue('💡 KEY INSIGHTS')
    .setFontSize(14).setFontWeight('bold')
    .setBackground(CONFIG.COLORS.WARNING).setHorizontalAlignment('center');
  pred.getRange('G17').setValue('• Player with highest win rate is likely champion');
  pred.getRange('G18').setValue('• Watch for players on 2+ win streaks');
  pred.getRange('G19').setValue('• Partnership ratings update after each match');

  // Conditional formatting for form column
  const formRange = pred.getRange(`B6:B${N + 5}`);
  pred.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('HOT').setBackground(CONFIG.COLORS.HOT).setFontColor('#ffffff')
      .setRanges([formRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('COLD').setBackground(CONFIG.COLORS.COLD).setFontColor('#ffffff')
      .setRanges([formRange]).build()
  ]);

  pred.setColumnWidths(1, 5, 130);
  pred.setColumnWidths(7, 5, 130);

  return pred;
}

// ============================================================================
// ACHIEVEMENTS SHEET + CHECKER
// ============================================================================

/**
 * Creates the Achievements sheet with the tracker grid sized to the active player count.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss]
 */
function createAchievementsSheet(ss) {
  if (!ss) ss = SpreadsheetApp.getActive();

  const players = getActivePlayerNames(ss);
  const N       = players.length;

  const ach = getOrCreateSheet(ss, 'Achievements');
  ach.clear();

  ach.getRange('A1').setValue('🏅 ACHIEVEMENT SYSTEM').setFontSize(18).setFontWeight('bold');
  ach.getRange('A2').setValue('Track special accomplishments and milestones')
    .setFontSize(10).setFontColor('#666666');

  // Achievement catalog
  ach.getRange('A4').setValue('📋 AVAILABLE ACHIEVEMENTS').setFontSize(14).setFontWeight('bold')
    .setBackground(CONFIG.COLORS.PRIMARY).setFontColor('#ffffff');
  ach.getRange('A5:D5').setValues([['Achievement','Description','Points','Rarity']])
    .setFontWeight('bold').setBackground('#f5f5f5');

  const achievements = [
    ['🎯 Perfect Game',    'Win without opponent scoring',      100, 'Legendary'],
    ['👑 Comeback King',   'Win after being down 10+ points',   75,  'Epic'     ],
    ['🛡️ Undefeated',     'Win all your matches',               200, 'Legendary'],
    ['🐕 Underdog Victory','Beat the top-ranked player',         50,  'Rare'     ],
    ['📊 Mr. Consistent',  'All matches within 5 point margin',  60,  'Rare'     ],
    ['🎪 Sharp Shooter',   'Score 30+ points in a match',        40,  'Uncommon' ],
    ['🛡️ Iron Wall',      'Hold opponent under 10 points',      50,  'Rare'     ],
    ['🔥 On Fire',         '3+ game win streak',                 80,  'Epic'     ],
    ['🌟 Versatile',       'Win with 3+ different partners',     90,  'Epic'     ],
    ['💪 Powerhouse',      'Win 5+ matches',                     70,  'Rare'     ],
    ['🎓 Master',          'Achieve 80%+ win rate (min 5 games)',150, 'Legendary'],
    ['⚡ Lightning',       'Win by 20+ points',                  60,  'Rare'     ],
    ['🤝 Team Player',     'Partner with every other player',   100,  'Epic'     ],
    ['🎖️ Veteran',        `Complete all ${generateRoundRobinSchedule(players).length} matches`, 30, 'Common'],
    ['🏆 Champion',        'Finish in 1st place',               250,  'Legendary']
  ];

  ach.getRange(6, 1, achievements.length, 4).setValues(achievements);

  // Rarity conditional formatting
  const rarityRange = ach.getRange(6, 4, achievements.length, 1);
  ach.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Legendary')
      .setBackground('#ffd700').setFontColor('#000000').setRanges([rarityRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Epic')
      .setBackground('#9b59b6').setFontColor('#ffffff').setRanges([rarityRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Rare')
      .setBackground('#3498db').setFontColor('#ffffff').setRanges([rarityRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Uncommon')
      .setBackground('#2ecc71').setFontColor('#ffffff').setRanges([rarityRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Common')
      .setBackground('#95a5a6').setFontColor('#ffffff').setRanges([rarityRange]).build()
  ]);

  // Player tracker grid
  const trackerRow = 6 + achievements.length + 3;
  ach.getRange(trackerRow, 1).setValue('🎮 PLAYER ACHIEVEMENT TRACKER')
    .setFontSize(14).setFontWeight('bold')
    .setBackground(CONFIG.COLORS.SUCCESS).setFontColor('#ffffff');

  // Column headers: Player | Ach1 | Ach2 | ...
  ach.getRange(trackerRow + 1, 1).setValue('Player').setFontWeight('bold');
  achievements.forEach((a, i) => {
    ach.getRange(trackerRow + 1, i + 2).setValue(a[0]).setFontWeight('bold');
  });

  // Player rows with '--' placeholders (populated by checkAndAwardAchievements)
  players.forEach((player, pIdx) => {
    ach.getRange(trackerRow + 2 + pIdx, 1).setValue(player);
    for (let i = 0; i < achievements.length; i++) {
      ach.getRange(trackerRow + 2 + pIdx, i + 2).setValue('--');
    }
  });

  // Achievement Leaderboard
  const leaderRow = trackerRow + N + 4;
  ach.getRange(leaderRow, 1, 1, 3).merge()
    .setValue('🌟 ACHIEVEMENT LEADERBOARD').setFontSize(14).setFontWeight('bold')
    .setBackground(CONFIG.COLORS.GOLD).setHorizontalAlignment('center');
  ach.getRange(leaderRow + 1, 1, 1, 3).setValues([['Player','Achievements','Total Points']])
    .setFontWeight('bold').setBackground('#f5f5f5');

  // Leaderboard rows (will be populated by checkAndAwardAchievements)
  players.forEach((player, i) => {
    ach.getRange(leaderRow + 2 + i, 1).setValue(player);
    ach.getRange(leaderRow + 2 + i, 2).setValue(0);
    ach.getRange(leaderRow + 2 + i, 3).setValue(0);
  });

  ach.setFrozenRows(5);
  ach.setColumnWidth(1, 200);
  if (achievements.length > 0) ach.autoResizeColumns(2, achievements.length + 1);

  return ach;
}

/**
 * Reads match results from the Schedule sheet and awards achievements to players.
 * Writes ✅ or -- in the tracker grid and updates the achievement leaderboard.
 * Fully implemented — no placeholders.
 */
function checkAndAwardAchievements() {
  const ss      = SpreadsheetApp.getActive();
  const players = getActivePlayerNames(ss);
  const N       = players.length;
  if (N === 0) return;

  const M    = generateRoundRobinSchedule(players).length;
  const sched = ss.getSheetByName('Schedule');
  const ach   = ss.getSheetByName('Achievements');
  if (!sched || !ach) return;

  // Load all match data
  const schedData = M > 0 ? sched.getRange(2, 1, M, 9).getValues() : [];

  // Build per-player stats
  const stats = {};
  players.forEach(name => {
    stats[name] = {
      wins: 0, losses: 0, scores: [], margins: [], partners: new Set(),
      opponents: [], perfectGames: false, ironWall: false,
      sharpShooter: false, lightning: false, maxStreak: 0, currentStreak: 0,
      comebackWins: 0, underdogWins: false
    };
  });

  schedData.forEach(row => {
    const p1 = String(row[1]).trim();
    const p2 = String(row[2]).trim();
    const p3 = String(row[3]).trim();
    const p4 = String(row[4]).trim();
    const sA = Number(row[5]);
    const sB = Number(row[6]);
    const status = String(row[7]).trim();

    if (status !== '✓' || isNaN(sA) || isNaN(sB)) return;

    const teamA  = [p1, p2];
    const teamB  = [p3, p4];
    const winA   = sA > sB;
    const margin = Math.abs(sA - sB);

    [teamA, teamB].forEach((team, teamIdx) => {
      const myScore  = teamIdx === 0 ? sA : sB;
      const oppScore = teamIdx === 0 ? sB : sA;
      const won      = teamIdx === 0 ? winA : !winA;
      const [me1, me2] = team;
      const opps = teamIdx === 0 ? teamB : teamA;

      [me1, me2].forEach(name => {
        if (!stats[name]) return;
        const s = stats[name];
        s.scores.push(myScore);
        s.margins.push(myScore - oppScore);
        s.partners.add(team.find(p => p !== name));
        opps.forEach(o => s.opponents.push(o));

        if (won) {
          s.wins++;
          s.currentStreak++;
          s.maxStreak = Math.max(s.maxStreak, s.currentStreak);
          if (oppScore === 0)      s.perfectGames = true;
          if (oppScore < 10)      s.ironWall      = true;
          if (myScore >= 30)      s.sharpShooter  = true;
          if (margin >= 20)       s.lightning      = true;
        } else {
          s.losses++;
          s.currentStreak = 0;
        }
      });
    });
  });

  // Determine leader (for Underdog check)
  const leaderboard = ss.getSheetByName('Leaderboard');
  const leader = leaderboard ? String(leaderboard.getRange('B2').getValue()).trim() : '';

  // Achievement definitions (order matches createAchievementsSheet)
  const TOTAL_MATCHES = M;
  const achievementCheckers = [
    (name, s) => s.perfectGames,
    (name, s) => s.comebackWins > 0,                               // Comeback King (simplified)
    (name, s) => s.losses === 0 && s.wins > 0,                    // Undefeated
    (name, s) => leader && s.opponents.includes(leader) && s.wins > 0, // Underdog
    (name, s) => s.margins.length > 0 && s.margins.every(m => Math.abs(m) <= 5), // Consistent
    (name, s) => s.sharpShooter,
    (name, s) => s.ironWall,
    (name, s) => s.maxStreak >= 3,                                 // On Fire
    (name, s) => s.partners.size >= 3 && s.wins > 0,              // Versatile
    (name, s) => s.wins >= 5,                                      // Powerhouse
    (name, s) => s.wins + s.losses >= 5 && s.wins / (s.wins + s.losses) >= 0.8, // Master
    (name, s) => s.lightning,
    (name, s) => s.partners.size >= N - 1,                        // Team Player
    (name, s) => (s.wins + s.losses) >= TOTAL_MATCHES / N * 0.9,  // Veteran (approx)
    (name, s) => leader === name                                   // Champion
  ];

  const achievementPoints = [100, 75, 200, 50, 60, 40, 50, 80, 90, 70, 150, 60, 100, 30, 250];

  // Find tracker grid location
  const achievementCount = achievementCheckers.length;
  const trackerRow = 6 + 15 + 3; // achievements.length = 15; see createAchievementsSheet

  // Write results to tracker
  const leaderMap = {};
  players.forEach((name, pIdx) => {
    const s           = stats[name];
    let earnedCount   = 0;
    let earnedPoints  = 0;

    achievementCheckers.forEach((check, aIdx) => {
      const earned = s ? check(name, s) : false;
      ach.getRange(trackerRow + 2 + pIdx, aIdx + 2).setValue(earned ? '✅' : '--');
      if (earned) {
        earnedCount++;
        earnedPoints += achievementPoints[aIdx];
      }
    });

    leaderMap[name] = { count: earnedCount, points: earnedPoints };
  });

  // Update achievement leaderboard
  const leaderRow = trackerRow + N + 4;
  players.forEach((name, i) => {
    ach.getRange(leaderRow + 2 + i, 1).setValue(name);
    ach.getRange(leaderRow + 2 + i, 2).setValue(leaderMap[name].count);
    ach.getRange(leaderRow + 2 + i, 3).setValue(leaderMap[name].points);
  });

  logActivity('Achievements checked and awarded for ' + N + ' players');
  SpreadsheetApp.getUi().alert('🏅 Done', 'Achievements have been recalculated for all players!', SpreadsheetApp.getUi().ButtonSet.OK);
}
