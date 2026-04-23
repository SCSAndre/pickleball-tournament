// ============================================================================
// Config.gs
// Global configuration, player data access, and shared utility functions.
// All other modules depend on this file.
// ============================================================================

const CONFIG = {
  TOURNAMENT_NAME: 'Epic Doubles Tournament',
  ENABLE_EMAIL_NOTIFICATIONS: true,
  ENABLE_ACHIEVEMENTS: true,
  ENABLE_PREDICTIONS: true,
  EMAIL_SEND_INTERVAL_HOURS: 24,
  ADMIN_EMAIL: Session.getActiveUser().getEmail(),

  COLORS: {
    PRIMARY:  '#4285f4',
    SUCCESS:  '#34a853',
    WARNING:  '#fbbc04',
    DANGER:   '#ea4335',
    GOLD:     '#ffd700',
    SILVER:   '#c0c0c0',
    BRONZE:   '#cd7f32',
    WIN:      '#d9ead3',
    LOSS:     '#f4cccc',
    PENDING:  '#fff2cc',
    HOT:      '#ff6b6b',
    COLD:     '#4ecdc4'
  },

  ACHIEVEMENTS: {
    PERFECT_GAME:  { name: '🎯 Perfect Game',   desc: 'Win without opponent scoring',       points: 100 },
    COMEBACK_KING: { name: '👑 Comeback King',   desc: 'Win after being down 10+ points',    points: 75  },
    UNDEFEATED:    { name: '🛡️ Undefeated',      desc: 'Win all matches',                    points: 200 },
    UNDERDOG:      { name: '🐕 Underdog',         desc: 'Beat top-ranked player',             points: 50  },
    CONSISTENT:    { name: '📊 Mr. Consistent',   desc: 'All matches within 5 points',        points: 60  },
    SHARP_SHOOTER: { name: '🎪 Sharp Shooter',    desc: 'Score 30+ points in a match',        points: 40  },
    IRON_WALL:     { name: '🛡️ Iron Wall',        desc: 'Hold opponent under 10 points',      points: 50  },
    STREAK:        { name: '🔥 On Fire',           desc: '3+ win streak',                      points: 80  },
    VERSATILE:     { name: '🌟 Versatile',         desc: 'Win with 3+ different partners',     points: 90  }
  }
};

// ============================================================================
// DYNAMIC PLAYER ACCESS — The single source of truth for active players.
// ============================================================================

/**
 * Reads the Roster sheet and returns all active players.
 * Columns: A=Player Name, B=Email, C=Skill Rating, D=Notification, E=Status
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss] - Optional spreadsheet reference.
 * @returns {Array<{name:string, email:string, skillRating:number, notification:string, status:string}>}
 */
function getActivePlayers(ss) {
  if (!ss) ss = SpreadsheetApp.getActive();
  const roster = ss.getSheetByName('Roster');
  if (!roster) return [];

  const lastRow = roster.getLastRow();
  if (lastRow < 2) return [];

  const data = roster.getRange(2, 1, lastRow - 1, 5).getValues();
  return data
    .filter(row => row[0] && String(row[4]).toLowerCase() === 'active')
    .map(row => ({
      name:         String(row[0]).trim(),
      email:        String(row[1]).trim(),
      skillRating:  Number(row[2]) || 0,
      notification: String(row[3]).trim(),
      status:       String(row[4]).trim()
    }));
}

/**
 * Returns just the player name strings for formula/list building.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss]
 * @returns {string[]}
 */
function getActivePlayerNames(ss) {
  return getActivePlayers(ss).map(p => p.name);
}

// ============================================================================
// SHARED UTILITY FUNCTIONS
// ============================================================================

/**
 * Gets a sheet by name, creating it if it doesn't exist.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {string} name
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getOrCreateSheet(ss, name) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

/**
 * Returns the current timestamp formatted for the script's timezone.
 * @returns {string}
 */
function getTimestamp() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

/**
 * Validates that a score value is a non-negative integer.
 * @param {*} score
 * @returns {boolean}
 */
function isValidScore(score) {
  return !isNaN(score) && score >= 0 && Number.isInteger(Number(score));
}

/**
 * Logs an activity entry to the Activity Log sheet (max 100 rows, rolling).
 * @param {string} message
 */
function logActivity(message) {
  try {
    const ss  = SpreadsheetApp.getActive();
    const log = getOrCreateSheet(ss, 'Activity Log');

    if (log.getLastRow() === 0) {
      log.getRange('A1:C1').setValues([['Timestamp', 'Action', 'User']]);
      log.getRange('A1:C1')
        .setFontWeight('bold')
        .setBackground(CONFIG.COLORS.PRIMARY)
        .setFontColor('#ffffff');
      log.setFrozenRows(1);
    }

    const nextRow = log.getLastRow() + 1;
    const user    = Session.getActiveUser().getEmail() || 'Unknown';
    log.getRange(nextRow, 1, 1, 3).setValues([[getTimestamp(), message, user]]);

    if (nextRow > 101) log.deleteRow(2);
  } catch (e) {
    console.log('Logging error: ' + e.message);
  }
}

/**
 * Returns a standard header style applicator (fluent helper).
 * @param {GoogleAppsScript.Spreadsheet.Range} range
 * @returns {GoogleAppsScript.Spreadsheet.Range}
 */
function styleHeader(range) {
  return range
    .setFontWeight('bold')
    .setBackground(CONFIG.COLORS.PRIMARY)
    .setFontColor('#ffffff');
}
