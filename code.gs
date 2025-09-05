/**
 * Chess.com API Data Fetcher for Google Sheets
 * Fetches monthly game archives with all possible Chess.com API and PGN headers
 */

// Configuration - Edit these values
const USERNAME = "hikaru";
const YEAR = 2025;
const MONTH = 6;

/**
 * Complete list of all possible Chess.com API JSON fields
 */
const CHESS_COM_API_HEADERS = [
  'Game URL', 'Game FEN', 'Start Time', 'End Time', 'Start Time Formatted', 'End Time Formatted',
  'Time Control', 'Base Time', 'Increment', 'Correspondence Length (days)', 'Time Class', 'Rules',
  'Format', 'ECO Opening', 'Rated', 'Tournament URL', 'Match URL',
  'White Username', 'White Rating', 'White Result', 'White Profile URL', 'White UUID', 'White Country',
  'Black Username', 'Black Rating', 'Black Result', 'Black Profile URL', 'Black UUID', 'Black Country',
  'White Accuracy', 'Black Accuracy', 'Initial Setup', 'PGN'
];

/**
 * Complete list of all possible PGN headers from the official specification
 */
const ALL_PGN_HEADERS = [
  // Seven Tag Roster (STR) - Required headers
  'Event', 'Site', 'Date', 'Round', 'White', 'Black', 'Result',
  
  // Player Information
  'WhiteTitle', 'BlackTitle', 'WhiteElo', 'BlackElo', 'WhiteUSCF', 'BlackUSCF',
  'WhiteNA', 'BlackNA', 'WhiteType', 'BlackType', 'WhiteRatingDiff', 'BlackRatingDiff',
  
  // Event Information
  'EventDate', 'EventSponsor', 'Section', 'Stage', 'Board',
  
  // Opening Information
  'Opening', 'Variation', 'SubVariation', 'ECO', 'NIC',
  
  // Time and Date
  'Time', 'UTCTime', 'UTCDate', 'TimeControl', 'SetUp', 'FEN',
  
  // Result Information
  'Termination', 'PlyCount', 'WhiteClock', 'BlackClock',
  
  // Annotation and Analysis
  'Annotator', 'Mode', 'WhiteTeam', 'BlackTeam',
  
  // Chess.com specific PGN headers
  'Link', 'CurrentPosition', 'Timezone', 'EndTime', 'StartTime',
  
  // Additional common headers found in online games
  'Variant', 'WhiteRD', 'BlackRD', 'WhiteIsComp', 'BlackIsComp',
  'TimeIncrement', 'WhiteTimeLeft', 'BlackTimeLeft'
];

// Opening family mapping to Chess.com URLs
const OPENING_FAMILY_MAP = {
  'Alekhine': 'https://www.chess.com/openings/Alekhine',
  'Benko Gambit': 'https://www.chess.com/openings/Benko',
  'Benoni': 'https://www.chess.com/openings/Benoni',
  'Bird': 'https://www.chess.com/openings/Birds-Opening',
  'Bishop': 'https://www.chess.com/openings/Bishops-Opening',
  'Bogo-Indian': 'https://www.chess.com/openings/Bogo-Indian-Defense',
  'Caro-Kann': 'https://www.chess.com/openings/Caro-Kann',
  'Catalan': 'https://www.chess.com/openings/Catalan',
  'Danish Gambit': 'https://www.chess.com/openings/Danish-Gambit',
  'Dutch': 'https://www.chess.com/openings/Dutch',
  'English': 'https://www.chess.com/openings/English',
  'Four Knights Game': 'https://www.chess.com/openings/Four-Knights-Game',
  'French': 'https://www.chess.com/openings/French',
  'Grünfeld': 'https://www.chess.com/openings/Grunfeld-Defense',
  "King's Fianchetto": 'https://www.chess.com/openings/Kings-Fianchetto-Opening',
  'Budapest Gambit': 'https://www.chess.com/openings/Budapest-Gambit',
  'Indian Game': 'https://www.chess.com/openings/Indian-Game',
  "King's Indian Defense": 'https://www.chess.com/openings/Kings-Indian-Defense',
  'Italian Game': 'https://www.chess.com/openings/Italian-Game',
  "King's Gambit": 'https://www.chess.com/openings/Kings-Gambit',
  "King's Indian Attack": 'https://www.chess.com/openings/Kings-Indian-Attack',
  'Modern w/1.e4': 'https://www.chess.com/openings/Modern-Defense',
  'Nimzo-Indian': 'https://www.chess.com/openings/Nimzo-Indian-Defense',
  'Nimzowitsch-Larsen': 'https://www.chess.com/openings/Nimzowitsch-Larsen-Attack',
  'Nimzowitsch': 'https://www.chess.com/openings/Nimzowitsch-Defense',
  'Old Indian': 'https://www.chess.com/openings/Old-Indian-Defense',
  'Owen': 'https://www.chess.com/openings/Owens-Defense',
  'Philidor': 'https://www.chess.com/openings/Philidor-Defense',
  'Polish': 'https://www.chess.com/openings/Polish-Opening',
  'Ponziani': 'https://www.chess.com/openings/Ponziani-Opening',
  'Colle': 'https://www.chess.com/openings/Colle-System',
  'London': 'https://www.chess.com/openings/London-System',
  "Queen's Pawn": 'https://www.chess.com/openings/Queens-Pawn-Opening',
  "Queen's Gambit": 'https://www.chess.com/openings/Queens-Gambit',
  "Queen's Indian": 'https://www.chess.com/openings/Queens-Indian',
  'Petrov': 'https://www.chess.com/openings/Petrov-Defense',
  'Scandinavian': 'https://www.chess.com/openings/Scandinavian-Defense',
  'Scotch Game': 'https://www.chess.com/openings/Scotch-Game',
  'Semi-Slav': 'https://www.chess.com/openings/Semi-Slav-Defense',
  'Alapin Sicilian': 'https://www.chess.com/openings/Alapin-Sicilian-Defense',
  'Closed Sicilian': 'https://www.chess.com/openings/Closed-Sicilian-Defense',
  'Sicilian': 'https://www.chess.com/openings/Sicilian-Defense',
  'Slav': 'https://www.chess.com/openings/Slav-Defense',
  'Ruy López': 'https://www.chess.com/openings/Ruy-Lopez-Opening',
  'Tarrasch': 'https://www.chess.com/openings/Tarrasch-Defense',
  'Three Knights': 'https://www.chess.com/openings/Three-Knights-Opening',
  'Trompowsky': 'https://www.chess.com/openings/Trompowsky-Attack',
  'Vienna Game': 'https://www.chess.com/openings/Vienna-Game',
  'Réti': 'https://www.chess.com/openings/Reti-Opening',
  "King's Pawn": 'https://www.chess.com/openings/Kings-Pawn-Opening',
  'Pirc': 'https://www.chess.com/openings/Pirc-Defense',
  'Center Game': 'https://www.chess.com/openings/Center-Game',
  'Richter-Versov Attack': 'https://www.chess.com/openings/Richter-Veresov-Attack',
  'Mieses Opening': 'https://www.chess.com/openings/Mieses-Opening',
  'Giuoco-Piano': 'https://www.chess.com/openings/Giuoco-Piano-Game',
  'Neo Grunfeld': 'https://www.chess.com/openings/Neo-Grunfeld-Defense',
  'Old Benoni': 'https://www.chess.com/openings/Old-Benoni-Defense',
  'Englund Gambit': 'https://www.chess.com/openings/Englund-Gambit',
  "Van 't Kruijs Opening": "https://www.chess.com/openings/Van-t-Kruijs-Opening",
  'Grob': 'https://www.chess.com/openings/Grob-Opening',
  "Petrov's Defense": 'https://www.chess.com/openings/Petrovs-Defense',
  'Lion Defense': 'https://www.chess.com/openings/Lion-Defense',
  'Clemez Opening': 'https://www.chess.com/openings/Clemenz-Opening',
  'Saragossa Opening': 'https://www.chess.com/openings/Saragossa-Opening',
  'Van Geet Opening': 'https://www.chess.com/openings/Van-Geet',
  "Anderssen's Opening": 'https://www.chess.com/openings/Anderssen-Opening'
};

/**
 * Main function to fetch Chess.com data and populate sheet
 */
function fetchChessComData() {
  try {
    const formattedMonth = MONTH.toString().padStart(2, '0');
    const apiUrl = `https://api.chess.com/pub/player/${USERNAME}/games/${YEAR}/${formattedMonth}`;
    
    Logger.log(`Fetching data from: ${apiUrl}`);
    
    // Fetch games from API
    const response = UrlFetchApp.fetch(apiUrl);
    if (response.getResponseCode() !== 200) {
      throw new Error(`API request failed with status: ${response.getResponseCode()}`);
    }
    
    const data = JSON.parse(response.getContentText());
    if (!data.games || data.games.length === 0) {
      Browser.msgBox('Info', 'No games found for the specified period', Browser.Buttons.OK);
      return;
    }
    
    Logger.log(`Found ${data.games.length} games`);
    
    // Setup sheet and process data
    const sheet = setupSheet();
    const { headers, gameRows } = processAllGames(data.games);
    
    // Write to sheet
    writeDataToSheet(sheet, headers, gameRows);
    
    Browser.msgBox('Success', `Successfully loaded ${data.games.length} games with all headers!`, Browser.Buttons.OK);
    Logger.log(`Successfully fetched data for ${USERNAME} - ${YEAR}/${formattedMonth}`);
    
  } catch (error) {
    Logger.log(`Error: ${error.toString()}`);
    Browser.msgBox('Error', `Failed to fetch data: ${error.toString()}`, Browser.Buttons.OK);
  }
}

/**
 * Setup or clear the RawData sheet
 */
function setupSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName('RawData');
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet('RawData');
    Logger.log('Created new RawData sheet');
  } else {
    sheet.clear();
    Logger.log('Cleared existing RawData sheet');
  }
  
  return sheet;
}

/**
 * Process all games with complete header structure
 */
function processAllGames(games) {
  // Create complete header list with all possible PGN headers
  const pgnHeaders = ALL_PGN_HEADERS.map(h => `PGN_${h}`);
  const pgnDataHeaders = [
    'PGN_Moves',
    'PGN_Moves_With_Clocks',
    'PGN_Ply_Count',
    'PGN_Move_Count',
    'PGN_Game_Outcome',
    'Opening_Family',
    'Opening_Family_URL',
    'Material_Imbalance',
    'Avg_Move_Time_Sec_per_Ply',
    'White_Castled',
    'Black_Castled',
    'Time_Scramble_Flag'
  ];
  const allHeaders = [...CHESS_COM_API_HEADERS, ...pgnHeaders, ...pgnDataHeaders];
  
  Logger.log(`Created comprehensive header set with ${allHeaders.length} columns`);
  Logger.log(`Chess.com API headers: ${CHESS_COM_API_HEADERS.length}`);
  Logger.log(`PGN headers: ${pgnHeaders.length}`);
  Logger.log(`Additional PGN data headers: ${pgnDataHeaders.length}`);
  
  // Process game rows with all parsed PGN data
  const gameRows = games.map(game => {
    const pgnData = parsePGN(game.pgn);
    return processGameRow(game, pgnData, allHeaders);
  });
  
  return { headers: allHeaders, gameRows };
}

/**
 * Process a single game into a comprehensive row array
 */
function processGameRow(game, pgnData, headers) {
  const safeGet = (obj, path, defaultValue = '') => {
    try {
      const value = path.split('.').reduce((current, key) => current && current[key], obj);
      return value == null ? defaultValue : value;
    } catch {
      return defaultValue;
    }
  };

  const startTimeFormatted = game.start_time ? new Date(game.start_time * 1000).toLocaleString() : '';
  const endTimeFormatted = game.end_time ? new Date(game.end_time * 1000).toLocaleString() : '';

  // Scale accuracies to 0–1 for percentage formatting
  const whiteAccRaw = safeGet(game, 'accuracies.white', '');
  const blackAccRaw = safeGet(game, 'accuracies.black', '');
  const whiteAccuracy = whiteAccRaw === '' ? '' : Number(whiteAccRaw) / 100;
  const blackAccuracy = blackAccRaw === '' ? '' : Number(blackAccRaw) / 100;

  // ECO/Opening fallback
  const ecoOpening = safeGet(game, 'eco', '') || (pgnData.headers.Opening || pgnData.headers.ECO || '');

  // Complete Chess.com API data extraction
  // Compute Base Time (minutes), Increment (seconds), Correspondence Length (days), and Format
  const timeControl = safeGet(game, 'time_control', '');
  const timeClass = safeGet(game, 'time_class', '');
  const rules = safeGet(game, 'rules', '');

  let baseTimeMinutes = '';
  let incrementSeconds = '';
  let correspondenceDays = '';

  if (typeof timeControl === 'string' && timeControl.includes('/')) {
    // Daily/correspondence format like "1/86400" => 1 move per 86400 seconds
    const parts = timeControl.split('/');
    const secondsPerMove = parts.length > 1 ? Number(parts[1]) : NaN;
    if (!isNaN(secondsPerMove) && secondsPerMove > 0) {
      correspondenceDays = secondsPerMove / 86400;
    }
  } else if (typeof timeControl === 'string' && timeControl.includes('+')) {
    const [baseStr, incStr] = timeControl.split('+');
    const baseSeconds = Number(baseStr);
    const incSeconds = Number(incStr);
    if (!isNaN(baseSeconds)) baseTimeMinutes = baseSeconds / 60;
    if (!isNaN(incSeconds)) incrementSeconds = incSeconds;
    if (baseTimeMinutes !== '' && incrementSeconds === '') incrementSeconds = 0;
  } else {
    const baseSeconds = Number(timeControl);
    if (!isNaN(baseSeconds) && baseSeconds > 0) {
      baseTimeMinutes = baseSeconds / 60;
      incrementSeconds = 0;
    }
  }

  const toTitle = s => (s && typeof s === 'string' && s.length) ? (s.charAt(0).toUpperCase() + s.slice(1)) : '';
  let formatDisplay = '';
  if (rules === 'chess' || rules === '' || rules == null) {
    // Same as time class for standard chess
    formatDisplay = toTitle(timeClass);
  } else if (rules === 'chess960') {
    formatDisplay = (timeClass === 'daily') ? 'Daily 960' : 'Live 960';
  } else {
    // Other variants are live-only per spec; show variant name
    formatDisplay = toTitle(rules);
  }

  const apiData = [
    safeGet(game, 'url'), safeGet(game, 'fen'), safeGet(game, 'start_time'), safeGet(game, 'end_time'),
    startTimeFormatted, endTimeFormatted, timeControl, baseTimeMinutes, incrementSeconds, correspondenceDays,
    timeClass, rules, formatDisplay, ecoOpening, safeGet(game, 'rated'), safeGet(game, 'tournament'),
    safeGet(game, 'match'), safeGet(game, 'white.username'), safeGet(game, 'white.rating'),
    safeGet(game, 'white.result'), safeGet(game, 'white.@id'), safeGet(game, 'white.uuid'),
    safeGet(game, 'white.country'), safeGet(game, 'black.username'), safeGet(game, 'black.rating'),
    safeGet(game, 'black.result'), safeGet(game, 'black.@id'), safeGet(game, 'black.uuid'),
    safeGet(game, 'black.country'), whiteAccuracy, blackAccuracy,
    safeGet(game, 'initial_setup'), safeGet(game, 'pgn')
  ];

  // Complete PGN header data extraction with date/time coercion
  const pgnHeaderData = ALL_PGN_HEADERS.map(header => {
    const val = pgnData.headers[header] || '';
    if (header === 'Date' || header === 'EventDate' || header === 'UTCDate') {
      return parsePgnDate(val);
    }
    if (header === 'Time' || header === 'UTCTime' || header === 'StartTime' || header === 'EndTime') {
      return parsePgnTime(val);
    }
    return val;
  });

  // Derive opening family using provided mapping
  const openingText = (pgnData.headers.Opening || pgnData.headers.ECO || ecoOpening || '').toString();
  const { familyName, familyUrl } = getOpeningFamily(openingText);

  // Parse moves and clocks; recompute ply/move counts
  const parsedMoves = parseMovesWithClocks(pgnData.moves);
  const plyCount = parsedMoves.plyCount;
  const moveCount = Math.ceil(plyCount / 2);

  // Duration and averages
  const durationSec = (game.end_time && game.start_time) ? Math.max(0, Number(game.end_time) - Number(game.start_time)) : '';
  const avgPerPly = (typeof durationSec === 'number' && plyCount) ? (durationSec / plyCount) : '';

  // Castling detection
  const whiteCastled = parsedMoves.whiteCastled;
  const blackCastled = parsedMoves.blackCastled;

  // Time scramble detection: any clock under 5 seconds
  const timeScrambleFlag = parsedMoves.anyClockUnderFiveSec;

  // Material imbalance from final FEN
  const materialImbalance = computeMaterialImbalance(safeGet(game, 'fen', ''));

  // PGN parsed data
  const pgnDataArray = [
    pgnData.moves,
    JSON.stringify(parsedMoves.moves),
    plyCount,
    moveCount,
    pgnData.gameOutcome,
    familyName,
    familyUrl,
    materialImbalance,
    avgPerPly,
    whiteCastled,
    blackCastled,
    timeScrambleFlag
  ];

  return [...apiData, ...pgnHeaderData, ...pgnDataArray];
}

/**
 * Write headers and data to sheet with enhanced formatting
 */
function writeDataToSheet(sheet, headers, gameRows) {
  // Write headers with color coding
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight('bold');
  headerRange.setFontColor('white');
  
  // Color-code different header sections
  const apiHeaderCount = CHESS_COM_API_HEADERS.length;
  const pgnHeaderCount = ALL_PGN_HEADERS.length;
  
  // Chess.com API headers - Blue
  const apiRange = sheet.getRange(1, 1, 1, apiHeaderCount);
  apiRange.setBackground('#4285f4');
  
  // PGN headers - Green  
  const pgnRange = sheet.getRange(1, apiHeaderCount + 1, 1, pgnHeaderCount);
  pgnRange.setBackground('#34a853');
  
  // PGN data headers - Orange
  const pgnDataCount = headers.length - (apiHeaderCount + pgnHeaderCount);
  const pgnDataRange = sheet.getRange(1, apiHeaderCount + pgnHeaderCount + 1, 1, pgnDataCount);
  pgnDataRange.setBackground('#ff9800');
  
  // Write game data
  if (gameRows.length > 0) {
    const dataRange = sheet.getRange(2, 1, gameRows.length, gameRows[0].length);
    dataRange.setValues(gameRows);
    
    // Apply formatting
    formatColumns(sheet, headers, gameRows.length);
  }
  
  // Auto-resize columns
  sheet.autoResizeColumns(1, headers.length);
  
  Logger.log(`Sheet updated with ${headers.length} columns and ${gameRows.length} rows`);
  Logger.log('Header sections: API (Blue), PGN (Green), PGN Data (Orange)');
}

/**
 * Apply comprehensive formatting to columns
 */
function formatColumns(sheet, headers, numRows) {
  const formatColumn = (headerName, format) => {
    const colIndex = headers.indexOf(headerName) + 1;
    if (colIndex > 0) {
      const range = sheet.getRange(2, colIndex, numRows, 1);
      range.setNumberFormat(format);
    }
  };
  
  // Format timestamp columns
  formatColumn('Start Time Formatted', 'yyyy-mm-dd hh:mm:ss');
  formatColumn('End Time Formatted', 'yyyy-mm-dd hh:mm:ss');
  
  // Format rating columns
  formatColumn('White Rating', '#,##0');
  formatColumn('Black Rating', '#,##0');
  formatColumn('PGN_WhiteElo', '#,##0');
  formatColumn('PGN_BlackElo', '#,##0');
  formatColumn('PGN_WhiteUSCF', '#,##0');
  formatColumn('PGN_BlackUSCF', '#,##0');
  
  // Format accuracy columns
  formatColumn('White Accuracy', '0.0%');
  formatColumn('Black Accuracy', '0.0%');

  // Format new time control derived columns
  formatColumn('Base Time', '#,##0.00');
  formatColumn('Increment', '#,##0');
  formatColumn('Correspondence Length (days)', '#,##0.00');
  
  // Format move count
  formatColumn('PGN_Move_Count', '#,##0');
  formatColumn('PGN_Ply_Count', '#,##0');
  formatColumn('Avg_Move_Time_Sec_per_Ply', '#,##0.00');
  
  // Format boolean columns
  const booleanHeaders = ['Rated', 'PGN_SetUp', 'PGN_WhiteIsComp', 'PGN_BlackIsComp', 'White_Castled', 'Black_Castled', 'Time_Scramble_Flag'];
  booleanHeaders.forEach(header => {
    const colIndex = headers.indexOf(header) + 1;
    if (colIndex > 0) {
      const range = sheet.getRange(2, colIndex, numRows, 1);
      range.setHorizontalAlignment('center');
    }
  });
  
  // Format date columns
  const dateHeaders = ['PGN_Date', 'PGN_EventDate', 'PGN_UTCDate'];
  dateHeaders.forEach(header => {
    const colIndex = headers.indexOf(header) + 1;
    if (colIndex > 0) {
      const range = sheet.getRange(2, colIndex, numRows, 1);
      range.setNumberFormat('yyyy-mm-dd');
    }
  });
  
  // Format time columns
  const timeHeaders = ['PGN_Time', 'PGN_UTCTime', 'PGN_StartTime', 'PGN_EndTime'];
  timeHeaders.forEach(header => {
    const colIndex = headers.indexOf(header) + 1;
    if (colIndex > 0) {
      const range = sheet.getRange(2, colIndex, numRows, 1);
      range.setNumberFormat('hh:mm:ss');
    }
  });

  // Other numeric columns
  formatColumn('Material_Imbalance', '#,##0');
  
  Logger.log('Comprehensive column formatting applied');
}

/**
 * Enhanced PGN parser with better header extraction
 */
function parsePGN(pgnString) {
  if (!pgnString || typeof pgnString !== 'string') {
    return { headers: {}, moves: '', moveCount: 0, gameOutcome: '' };
  }
  
  const result = { headers: {}, moves: '', moveCount: 0, gameOutcome: '' };
  
  try {
    const lines = pgnString.split('\n');
    let movesText = '';
    
    for (let line of lines) {
      line = line.trim();
      if (!line) continue;
      
      // Parse headers with improved regex
      if (line.startsWith('[') && line.endsWith(']')) {
        const headerMatch = line.match(/^\[(\w+)\s+"([^"]*)"\]$/);
        if (headerMatch) {
          const [, key, value] = headerMatch;
          result.headers[key] = value;
        }
      } else {
        // Collect moves text
        movesText += line + ' ';
      }
    }
    
    // Process moves with enhanced parsing
    movesText = movesText.trim();
    if (movesText) {
      // Extract game outcome with more robust regex first
      const outcomeMatch = movesText.match(/(1-0|0-1|1\/2-1\/2|\*)(?:\s*$)/);
      if (outcomeMatch) {
        result.gameOutcome = outcomeMatch[1];
      }

      // Remove result token from moves text for cleaner parsing
      const cleaned = movesText.replace(/\s*(1-0|0-1|1\/2-1\/2|\*)\s*$/, '').trim();
      result.moves = cleaned;

      // Count full moves more robustly: numbers followed by dot or ellipsis
      const moveMatches = cleaned.match(/\b\d+\.(?:\.\.)?/g);
      result.moveCount = moveMatches ? moveMatches.length : 0;
    }
    
    // Post-process some headers for better data quality
    if (result.headers.Date && result.headers.Date.includes('.')) {
      // Ensure date format consistency
      const dateParts = result.headers.Date.split('.');
      if (dateParts.length === 3) {
        result.headers.Date = `${dateParts[0]}.${dateParts[1].padStart(2, '0')}.${dateParts[2].padStart(2, '0')}`;
      }
    }
    
  } catch (error) {
    Logger.log(`Error parsing PGN: ${error.toString()}`);
  }
  
  return result;
}

/**
 * Helpers to convert PGN date/time strings into Date objects for formatting
 */
function parsePgnDate(value) {
  if (!value || value === '????.??.??') return '';
  const parts = (value || '').split('.');
  const year = Number(parts[0]);
  const month = Math.max(1, Number(parts[1] || '1')) - 1;
  const day = Math.max(1, Number(parts[2] || '1'));
  if (!year || isNaN(year)) return '';
  return new Date(year, month, day);
}

function parsePgnTime(value) {
  if (!value || value === '??:??:??') return '';
  const parts = (value || '').split(':');
  const hours = Number(parts[0] || '0');
  const minutes = Number(parts[1] || '0');
  const seconds = Number(parts[2] || '0');
  return new Date(Date.UTC(1970, 0, 1, hours, minutes, seconds));
}

/**
 * Parse SAN moves with embedded clock annotations into a structured array
 * Returns an object with:
 * - moves: [{ ply, moveNumber, side, san, clock, clockSeconds }]
 * - plyCount, whiteCastled, blackCastled, anyClockUnderFiveSec
 */
function parseMovesWithClocks(movesText) {
  const result = { moves: [], plyCount: 0, whiteCastled: false, blackCastled: false, anyClockUnderFiveSec: false };
  if (!movesText) return result;

  // Remove comments other than clock annotations and NAGs; keep SAN + clock tags
  // Normalize multiple spaces
  const tokens = movesText
    .replace(/\{\s*\[%eval[^}]*\]\s*\}/g, ' ') // drop eval annotations if present
    .replace(/\s+/g, ' ')
    .trim()
    .split(' ');

  let moveNumber = 0;
  let side = 'w';
  let ply = 0;

  const isMoveNumber = t => /^(\d+)\.(\.{2})?$/.test(t);
  const isClockTag = t => /^\{\[%clk\s+[^}]+\]\}$/.test(t);

  for (let i = 0; i < tokens.length; i++) {
    const tok = tokens[i];
    if (!tok) continue;
    if (isMoveNumber(tok)) {
      const m = tok.match(/^(\d+)\.(\.{2})?$/);
      moveNumber = Number(m[1]);
      side = m[2] ? 'b' : 'w';
      continue;
    }
    if (/^(1-0|0-1|1\/2-1\/2|\*)$/.test(tok)) {
      break;
    }

    // tok should be SAN
    const san = tok;
    // Next token may be a clock tag
    let clock = '';
    let clockSeconds = '';
    if (i + 1 < tokens.length && isClockTag(tokens[i + 1])) {
      clock = tokens[i + 1].slice(1, -1); // remove braces
      // Extract time like 0:03:00.9 or 3:00
      const clkMatch = clock.match(/%clk\s+([0-9]+:[0-9]{2}:[0-9]{2}(?:\.[0-9])?|[0-9]+:[0-9]{2}(?:\.[0-9])?)/);
      if (clkMatch) {
        clockSeconds = parseClockToSeconds(clkMatch[1]);
        if (typeof clockSeconds === 'number' && clockSeconds < 5) {
          result.anyClockUnderFiveSec = true;
        }
      }
      i++;
    }

    ply += 1;
    const entry = { ply, moveNumber, side, san, clock, clockSeconds };
    result.moves.push(entry);

    if (san === 'O-O' || san === 'O-O-O') {
      if (side === 'w') result.whiteCastled = true;
      else result.blackCastled = true;
    }

    // Toggle side and increment move number when black finishes
    if (side === 'w') {
      side = 'b';
    } else {
      side = 'w';
    }
  }

  result.plyCount = ply;
  return result;
}

function parseClockToSeconds(text) {
  // Accept H:MM:SS(.d) or M:SS(.d)
  if (!text) return '';
  const parts = text.split(':');
  let hours = 0, minutes = 0, seconds = 0;
  if (parts.length === 3) {
    hours = Number(parts[0]);
    minutes = Number(parts[1]);
    seconds = Number(parts[2]);
  } else if (parts.length === 2) {
    minutes = Number(parts[0]);
    seconds = Number(parts[1]);
  } else {
    return '';
  }
  if (isNaN(hours) || isNaN(minutes) || isNaN(seconds)) return '';
  return hours * 3600 + minutes * 60 + seconds;
}

/**
 * Compute material imbalance from FEN using simple piece values
 * Positive means White is ahead, negative means Black is ahead
 */
function computeMaterialImbalance(fen) {
  if (!fen) return '';
  const pieceValues = { p:1, n:3, b:3, r:5, q:9, k:0 };
  const board = fen.split(' ')[0];
  let score = 0;
  for (const ch of board) {
    if (/[1-8/]/.test(ch)) continue;
    const lower = ch.toLowerCase();
    const val = pieceValues[lower] || 0;
    score += (ch === lower) ? -val : val;
  }
  return score;
}

/**
 * Map opening text to a Chess.com family name and URL
 */
function getOpeningFamily(openingText) {
  const map = OPENING_FAMILY_MAP;
  const normalized = (openingText || '').toLowerCase();
  for (const key in map) {
    if (!Object.prototype.hasOwnProperty.call(map, key)) continue;
    const keyNorm = key.toLowerCase();
    if (normalized.includes(keyNorm)) {
      return { familyName: key, familyUrl: map[key] };
    }
  }
  return { familyName: '', familyUrl: '' };
}

/**
 * Get available archives for the configured player
 */
function getAvailableArchives() {
  const apiUrl = `https://api.chess.com/pub/player/${USERNAME}/games/archives`;
  
  try {
    const response = UrlFetchApp.fetch(apiUrl);
    const data = JSON.parse(response.getContentText());
    
    Logger.log(`Found ${data.archives.length} available archives for ${USERNAME}:`);
    data.archives.forEach((archive, index) => {
      Logger.log(`${index + 1}: ${archive}`);
    });
    
    return data.archives;
    
  } catch (error) {
    Logger.log(`Error fetching archives: ${error.toString()}`);
    return [];
  }
}

/**
 * Fetch all archives for the configured player and populate the sheet
 */
function fetchAllArchivesForUser() {
  try {
    const archives = getAvailableArchives();
    if (!archives.length) {
      Browser.msgBox('Info', `No archives found for ${USERNAME}`, Browser.Buttons.OK);
      return;
    }

    const allGames = [];
    archives.forEach(url => {
      try {
        const res = UrlFetchApp.fetch(url);
        if (res.getResponseCode() !== 200) return;
        const data = JSON.parse(res.getContentText());
        if (data.games && data.games.length) {
          allGames.push(...data.games);
        }
        Utilities.sleep(150);
      } catch (innerErr) {
        Logger.log(`Failed to fetch archive ${url}: ${innerErr}`);
      }
    });

    if (!allGames.length) {
      Browser.msgBox('Info', 'No games found across all archives', Browser.Buttons.OK);
      return;
    }

    const sheet = setupSheet();
    const { headers, gameRows } = processAllGames(allGames);
    writeDataToSheet(sheet, headers, gameRows);

    Browser.msgBox('Success', `Loaded ${allGames.length} games from ${archives.length} archives`, Browser.Buttons.OK);
    Logger.log(`Successfully fetched ${allGames.length} games across ${archives.length} archives for ${USERNAME}`);
  } catch (error) {
    Logger.log(`Error fetching all archives: ${error.toString()}`);
    Browser.msgBox('Error', `Failed to fetch all archives: ${error.toString()}`, Browser.Buttons.OK);
  }
}

/**
 * Utility function to display header information
 */
function showHeaderInfo() {
  Logger.log('=== CHESS.COM API HEADERS ===');
  CHESS_COM_API_HEADERS.forEach((header, index) => {
    Logger.log(`${index + 1}: ${header}`);
  });
  
  Logger.log('\n=== ALL POSSIBLE PGN HEADERS ===');
  ALL_PGN_HEADERS.forEach((header, index) => {
    Logger.log(`${index + 1}: ${header}`);
  });
  
  Logger.log(`\nTotal headers: ${CHESS_COM_API_HEADERS.length + ALL_PGN_HEADERS.length + 3}`);
  Logger.log('Note: +3 for PGN_Moves, PGN_Move_Count, PGN_Game_Outcome');
}
