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

// Perspective headers appended at the end of the dataset
const PERSPECTIVE_HEADERS = [
  'My_Color',
  'My_Result',
  'Termination',
  'My_Rating_Change'
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

// Synonyms -> Canonical family name (normalized keys)
const OPENING_FAMILY_SYNONYMS = {
  // Modern Defense synonyms
  'modern defense': 'Modern w/1.e4',
  'modern': 'Modern w/1.e4',
  // Reti
  'reti': 'Réti',
  // Petrov/Petrovs spelling
  "petrov's defense": "Petrov's Defense",
  'petrov defense': 'Petrov',
  // Giuoco Piano without hyphen
  'giuoco piano': 'Giuoco-Piano',
  // Ruy Lopez without accent
  'ruy lopez': 'Ruy López',
  // Grunfeld without umlaut
  'grunfeld': 'Grünfeld',
  'neo grunfeld': 'Neo Grunfeld',
  // Kings Pawn variations
  "king's pawn opening": "King's Pawn",
  // Van 't Kruijs variants
  "van t kruijs": "Van 't Kruijs Opening",
  "van't kruijs": "Van 't Kruijs Opening",
  // Four Knights variations
  'four knights': 'Four Knights Game',
  // Scotch without Game
  'scotch': 'Scotch Game',
  // Queen's Pawn/Gambit common variants
  "queens pawn": "Queen's Pawn",
  "queens gambit": "Queen's Gambit",
  // Kings Gambit
  "kings gambit": "King's Gambit",
  // Indian Game variants
  'old indian defense': 'Old Indian',
  'bogo indian': 'Bogo-Indian',
  // Sicilian variants
  'alapin': 'Alapin Sicilian',
  'closed sicilian defense': 'Closed Sicilian',
  // Benko/Benoni spellings
  'benko': 'Benko Gambit',
  'englund': 'Englund Gambit'
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
    'PGN_Moves_String',
    'PGN_Ply_Count',
    'PGN_Move_Count',
    'PGN_Game_Outcome',
    'Opening_Family',
    'Opening_Family_URL',
    'Material_Imbalance',
    'Avg_Move_Time_Sec_per_Ply',
    'White_Castled',
    'Black_Castled',
    'Time_Scramble_Flag',
    'Sub5s_Count',
    'Sub3s_Count',
    'White_Clock_Min_Sec',
    'Black_Clock_Min_Sec',
    'First_Sub5s_Ply'
  ];

  // Build base rows and collect metadata for rating computation
  const rowObjs = games.map(game => {
    const pgnData = parsePGN(game.pgn);
    const baseRow = processGameRow(game, pgnData, []);

    // Derive format display (same logic as in processGameRow)
    const timeControl = (game && game.time_control) || '';
    const timeClass = (game && game.time_class) || '';
    const rules = (game && game.rules) || '';
    const toTitle = s => (s && typeof s === 'string' && s.length) ? (s.charAt(0).toUpperCase() + s.slice(1)) : '';
    let formatDisplay = '';
    if (rules === 'chess' || rules === '' || rules == null) {
      formatDisplay = toTitle(timeClass);
    } else if (rules === 'chess960') {
      formatDisplay = (timeClass === 'daily') ? 'Daily 960' : 'Live 960';
    } else {
      formatDisplay = toTitle(rules);
    }

    // Determine my color and pre-rating
    const me = (USERNAME || '').toLowerCase();
    const whiteUser = ((game && game.white && game.white.username) || '').toLowerCase();
    const blackUser = ((game && game.black && game.black.username) || '').toLowerCase();
    const myColor = me && whiteUser && me === whiteUser ? 'White' : (me && blackUser && me === blackUser ? 'Black' : '');
    const preRating = myColor === 'White' ? Number((game && game.white && game.white.rating) || '')
                     : myColor === 'Black' ? Number((game && game.black && game.black.rating) || '')
                     : '';

    // Rating diff from PGN headers
    let ratingDiff = '';
    if (myColor === 'White') {
      const diffRaw = pgnData.headers['WhiteRatingDiff'];
      if (diffRaw != null && diffRaw !== '') ratingDiff = Number(String(diffRaw).replace('+', ''));
    } else if (myColor === 'Black') {
      const diffRaw = pgnData.headers['BlackRatingDiff'];
      if (diffRaw != null && diffRaw !== '') ratingDiff = Number(String(diffRaw).replace('+', ''));
    }

    const afterRating = (typeof preRating === 'number' && !isNaN(preRating) && typeof ratingDiff === 'number' && !isNaN(ratingDiff))
      ? (preRating + ratingDiff)
      : (typeof preRating === 'number' && !isNaN(preRating) ? preRating : '');

    const endTime = Number((game && game.end_time) || 0);

    return { row: baseRow, endTime, formatDisplay, myColor, preRating, ratingDiff, afterRating };
  });

  // Sort by end time ascending to enable forward-fill by time
  rowObjs.sort((a, b) => (a.endTime || 0) - (b.endTime || 0));

  // Collect unique formats encountered and ensure default formats are present
  const DEFAULT_FORMATS = ['Bullet', 'Blitz', 'Rapid', 'Daily', 'Live 960', 'Daily 960'];
  const encountered = [];
  rowObjs.forEach(o => { if (o.formatDisplay && !encountered.includes(o.formatDisplay)) encountered.push(o.formatDisplay); });
  const formats = Array.from(new Set([...DEFAULT_FORMATS, ...encountered]));

  // Build dynamic rating headers for each format
  const toHeaderSafe = f => `My_Rating_${String(f).replace(/\s+/g, '_')}`;
  const ratingHeaders = formats.map(toHeaderSafe);

  // Rebuild rows to insert rating columns (forward-filled) and rating change in perspective
  const rebuiltRows = [];
  const lastKnownByFormat = {};

  const apiHeaderCount = CHESS_COM_API_HEADERS.length;
  const pgnHeaderCount = ALL_PGN_HEADERS.length;
  const baseCount = apiHeaderCount + pgnHeaderCount + pgnDataHeaders.length;

  for (const obj of rowObjs) {
    const basePart = obj.row.slice(0, baseCount);
    const perspectivePart = obj.row.slice(baseCount); // [My_Color, My_Result, Termination]

    // Values before applying this game's update (most recent before time-wise)
    const ratingValuesBefore = formats.map(fmt => {
      const prev = lastKnownByFormat[fmt];
      return (typeof prev === 'number' && !isNaN(prev)) ? prev : '';
    });

    // Compute this game's format overrides and rating change
    let ratingChange = '';
    if (obj.formatDisplay) {
      const fmtIndex = formats.indexOf(obj.formatDisplay);
      if (fmtIndex >= 0) {
        // Rating change equals PGN rating diff when available; else derive from last known
        if (typeof obj.ratingDiff === 'number' && !isNaN(obj.ratingDiff)) {
          ratingChange = obj.ratingDiff;
        } else {
          const prev = lastKnownByFormat[obj.formatDisplay];
          if ((typeof obj.afterRating === 'number' && !isNaN(obj.afterRating)) && (typeof prev === 'number' && !isNaN(prev))) {
            ratingChange = obj.afterRating - prev;
          }
        }

        // Override the current format with after-game rating
        if (typeof obj.afterRating === 'number' && !isNaN(obj.afterRating)) {
          ratingValuesBefore[fmtIndex] = obj.afterRating;
          lastKnownByFormat[obj.formatDisplay] = obj.afterRating;
        }
      }
    }

    const newPerspective = [...perspectivePart, ratingChange];
    rebuiltRows.push([...basePart, ...ratingValuesBefore, ...newPerspective]);
  }

  // Do not backfill earlier rows with later ratings; leave blanks until first occurrence per format

  const allHeaders = [...CHESS_COM_API_HEADERS, ...pgnHeaders, ...pgnDataHeaders, ...ratingHeaders, ...PERSPECTIVE_HEADERS];

  Logger.log(`Created comprehensive header set with ${allHeaders.length} columns`);
  Logger.log(`Chess.com API headers: ${CHESS_COM_API_HEADERS.length}`);
  Logger.log(`PGN headers: ${pgnHeaders.length}`);
  Logger.log(`Additional PGN data headers: ${pgnDataHeaders.length}`);
  Logger.log(`Dynamic rating headers: ${ratingHeaders.length}`);

  return { headers: allHeaders, gameRows: rebuiltRows };
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

  // Derive start time from end time and PGN Start/End time difference when start_time is missing
  const pgnStartStr = (pgnData && pgnData.headers && pgnData.headers['StartTime']) ? String(pgnData.headers['StartTime']) : '';
  const pgnEndStr = (pgnData && pgnData.headers && pgnData.headers['EndTime']) ? String(pgnData.headers['EndTime']) : '';
  let pgnDeltaSec = '';
  if (pgnStartStr && pgnEndStr) {
    const startSec = parseClockToSeconds(pgnStartStr);
    const endSec = parseClockToSeconds(pgnEndStr);
    if (typeof startSec === 'number' && typeof endSec === 'number') {
      let delta = endSec - startSec;
      if (delta < 0) delta += 24 * 3600; // handle crossing midnight
      pgnDeltaSec = delta;
    }
  }

  const endTimeEpoch = safeGet(game, 'end_time', '');
  let startTimeEpoch = safeGet(game, 'start_time', '');
  if ((startTimeEpoch === '' || startTimeEpoch == null) && endTimeEpoch !== '' && typeof pgnDeltaSec === 'number') {
    startTimeEpoch = Number(endTimeEpoch) - pgnDeltaSec;
  }

  const startTimeFormatted = (startTimeEpoch !== '' && startTimeEpoch != null) ? new Date(Number(startTimeEpoch) * 1000) : '';
  const endTimeFormatted = (endTimeEpoch !== '' && endTimeEpoch != null) ? new Date(Number(endTimeEpoch) * 1000) : '';

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
    safeGet(game, 'url'), safeGet(game, 'fen'), startTimeEpoch, endTimeEpoch,
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

  // Derive opening family using Chess.com ECO Opening text
  const openingText = (ecoOpening || '').toString();
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

  // Build human-readable moves string
  const humanMoves = buildHumanReadableMoves(parsedMoves.moves);

  // Perspective fields
  const me = (USERNAME || '').toLowerCase();
  const whiteUser = (safeGet(game, 'white.username', '') || '').toLowerCase();
  const blackUser = (safeGet(game, 'black.username', '') || '').toLowerCase();
  const myColor = me && whiteUser && me === whiteUser ? 'White' : (me && blackUser && me === blackUser ? 'Black' : '');
  const whiteResult = safeGet(game, 'white.result', '');
  const blackResult = safeGet(game, 'black.result', '');
  let myResult = '';
  if (myColor === 'White') myResult = resultToScore(whiteResult);
  else if (myColor === 'Black') myResult = resultToScore(blackResult);
  // Termination: if one side is 'win', take the other side's result; if draw, both same -> take that
  let termination = '';
  if (whiteResult === 'win' && blackResult) termination = blackResult;
  else if (blackResult === 'win' && whiteResult) termination = whiteResult;
  else if (whiteResult && blackResult && whiteResult === blackResult) termination = whiteResult;

  // PGN parsed data
  const compactMoves = parsedMoves.moves.map(m => {
    const timeStr = (m.clockSeconds !== '' && m.clockSeconds != null) ? formatSecondsToHmsTenths(m.clockSeconds) : '';
    return [m.ply, m.moveNumber, m.side, m.san, timeStr];
  });
  const pgnDataArray = [
    pgnData.moves,
    JSON.stringify(compactMoves),
    humanMoves,
    plyCount,
    moveCount,
    pgnData.gameOutcome,
    familyName,
    familyUrl,
    materialImbalance,
    avgPerPly,
    whiteCastled,
    blackCastled,
    timeScrambleFlag,
    parsedMoves.sub5Count,
    parsedMoves.sub3Count,
    parsedMoves.whiteClockMinSec,
    parsedMoves.blackClockMinSec,
    parsedMoves.firstSub5Ply
  ];

  const perspectiveArray = [myColor, myResult, termination];

  return [...apiData, ...pgnHeaderData, ...pgnDataArray, ...perspectiveArray];
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
  const ratingCount = headers.filter(h => String(h).indexOf('My_Rating_') === 0).length;
  const pgnDataCount = headers.length - (apiHeaderCount + pgnHeaderCount + ratingCount + PERSPECTIVE_HEADERS.length);
  const pgnDataRange = sheet.getRange(1, apiHeaderCount + pgnHeaderCount + 1, 1, pgnDataCount);
  pgnDataRange.setBackground('#ff9800');

  // Rating-by-format headers - Teal
  const ratingStart = apiHeaderCount + pgnHeaderCount + pgnDataCount + 1;
  if (ratingCount > 0) {
    const ratingRange = sheet.getRange(1, ratingStart, 1, ratingCount);
    ratingRange.setBackground('#00bcd4');
  }

  // Perspective headers - Purple
  const perspectiveStart = apiHeaderCount + pgnHeaderCount + pgnDataCount + ratingCount + 1;
  const perspectiveRange = sheet.getRange(1, perspectiveStart, 1, PERSPECTIVE_HEADERS.length);
  perspectiveRange.setBackground('#9c27b0');
  
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
  Logger.log('Header sections: API (Blue), PGN (Green), PGN Data (Orange), Perspective (Purple)');
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
  formatColumn('Sub5s_Count', '#,##0');
  formatColumn('Sub3s_Count', '#,##0');
  formatColumn('White_Clock_Min_Sec', '#,##0.0');
  formatColumn('Black_Clock_Min_Sec', '#,##0.0');
  formatColumn('First_Sub5s_Ply', '#,##0');

  // Perspective
  formatColumn('My_Result', '0.0');
  formatColumn('My_Rating_Change', '#,##0');

  // Rating-by-format columns
  headers.forEach(h => {
    if (String(h).indexOf('My_Rating_') === 0) {
      formatColumn(h, '#,##0');
    }
  });
  
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
  const result = { moves: [], plyCount: 0, whiteCastled: false, blackCastled: false, anyClockUnderFiveSec: false, sub5Count: 0, sub3Count: 0, whiteClockMinSec: '', blackClockMinSec: '', firstSub5Ply: '' };
  if (!movesText) return result;

  // 1) Keep only clock comments; remove all other braces comments
  // Accept patterns like {[%clk H:MM:SS(.d)]} or {[%clk M:SS(.d)]}
  const commentsCleaned = movesText.replace(/\{[^}]*\}/g, (m) => (/\[%\s*clk\b/.test(m) ? m : ' '));
  // 2) Remove Numeric Annotation Glyphs like $1, $12, etc.
  const nagsCleaned = commentsCleaned.replace(/\$\d+/g, ' ');
  // 3) Remove simple parentheses used for variations (shallow)
  const parenCleaned = nagsCleaned.replace(/[()]/g, ' ');
  // 4) Normalize spaces
  const tokens = parenCleaned.replace(/\s+/g, ' ').trim().split(' ');

  // SAN validation regex (covers pieces, disambiguation, captures, promotions, checks/mates, castling)
  const SAN_REGEX = /^(?:O-O(?:-O)?[+#]?|[KQRBN]?[a-h]?[1-8]?x?[a-h][1-8](?:=[QRBN])?[+#]?|[a-h]x[a-h][1-8](?:=[QRBN])?[+#]?|[a-h][1-8](?:=[QRBN])?[+#]?)(?:[!?]{1,2})?$/;

  let moveNumber = 0;
  let side = 'w';
  let ply = 0;
  let maxFullMove = 0;

  const isMoveNumber = t => /^(\d+)\.(\.{2})?$/.test(t);
  const isClockTag = t => /^\{\[%\s*clk\s+[^}]+\]\}$/.test(t);
  const isResultTok = t => /^(1-0|0-1|1\/2-1\/2|\*)$/.test(t);
  const isSAN = t => SAN_REGEX.test(t);

  for (let i = 0; i < tokens.length; i++) {
    const tok = tokens[i];
    if (!tok) continue;

    if (isMoveNumber(tok)) {
      const m = tok.match(/^(\d+)\.(\.{2})?$/);
      moveNumber = Number(m[1]);
      if (moveNumber > maxFullMove) maxFullMove = moveNumber;
      side = m[2] ? 'b' : 'w';
      continue;
    }

    if (isResultTok(tok)) break;

    // Sanity: only accept valid SAN tokens as moves
    if (!isSAN(tok)) continue;

    // Next token may be a clock tag
    let clock = '';
    let clockSeconds = '';
    if (i + 1 < tokens.length && isClockTag(tokens[i + 1])) {
      clock = tokens[i + 1].slice(1, -1); // remove braces
      // Capture H:MM:SS(.d) or M:SS(.d)
      const clkMatch = clock.match(/%\s*clk\s+([0-9]+:[0-9]{2}:[0-9]{2}(?:\.[0-9])?|[0-9]+:[0-9]{2}(?:\.[0-9])?)/);
      if (clkMatch) {
        clockSeconds = parseClockToSeconds(clkMatch[1]);
        if (typeof clockSeconds === 'number') {
          if (clockSeconds < 5) {
            result.anyClockUnderFiveSec = true;
            result.sub5Count += 1;
            if (result.firstSub5Ply === '') result.firstSub5Ply = (ply + 1);
          }
          if (clockSeconds < 3) result.sub3Count += 1;
          if (side === 'w') {
            if (result.whiteClockMinSec === '' || clockSeconds < result.whiteClockMinSec) result.whiteClockMinSec = clockSeconds;
          } else {
            if (result.blackClockMinSec === '' || clockSeconds < result.blackClockMinSec) result.blackClockMinSec = clockSeconds;
          }
        }
      }
      i++;
    }

    // Default first move number if omitted
    if (moveNumber === 0 && side === 'w') moveNumber = 1;

    ply += 1;
    const san = tok;
    const entry = { ply, moveNumber, side, san, clock, clockSeconds };
    result.moves.push(entry);

    if (san === 'O-O' || san === 'O-O-O') {
      if (side === 'w') result.whiteCastled = true; else result.blackCastled = true;
    }

    // If black just moved, the next white move is the next full-move number (when not explicitly provided)
    if (side === 'b') {
      moveNumber += 1;
      if (moveNumber > maxFullMove) maxFullMove = moveNumber;
    }

    // Toggle side after processing the move
    side = (side === 'w') ? 'b' : 'w';
  }

  result.plyCount = ply;
  return result;
}

function buildHumanReadableMoves(moves) {
  if (!moves || !moves.length) return '';
  const parts = [];
  let currentMoveNumber = 0;
  let buffer = '';
  for (const m of moves) {
    if (m.moveNumber !== currentMoveNumber) {
      // flush previous
      if (buffer) {
        parts.push(buffer.trim());
        buffer = '';
      }
      currentMoveNumber = m.moveNumber;
      buffer += `${m.moveNumber}.`;
      if (m.side === 'b') buffer += '..';
      buffer += ` ${m.san}`;
    } else {
      buffer += ` ${m.san}`;
    }
  }
  if (buffer) parts.push(buffer.trim());
  return parts.join(' ');
}

function resultToScore(result) {
  if (!result) return '';
  // Chess.com results: win, checkmated, resigned, timeout, stalemate, agreed, repetition, timevsinsufficient, abandoned, other
  if (result === 'win') return 1;
  if (result === 'stalemate' || result === 'agreed' || result === 'repetition' || result === 'timevsinsufficient') return 0.5;
  return 0;
}

function parseClockToSeconds(text) {
  // Accept H:MM:SS(.d) or M:SS(.d)
  if (!text) return '';
  const cleaned = String(text).trim();
  const frac = cleaned.includes('.') ? Number(cleaned.slice(cleaned.lastIndexOf('.') + 1)) / 10 : 0;
  const base = cleaned.replace(/\.[0-9]$/, '');
  const parts = base.split(':');
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
  return hours * 3600 + minutes * 60 + seconds + frac;
}

// Format seconds to H:MM:SS.d where tenths shown if fractional
function formatSecondsToHmsTenths(totalSeconds) {
  if (totalSeconds === '' || totalSeconds == null || isNaN(totalSeconds)) return '';
  const hasFraction = Math.abs(totalSeconds - Math.trunc(totalSeconds)) > 0.0001;
  const tenths = Math.round((totalSeconds % 1) * 10);
  const whole = Math.floor(totalSeconds);
  const hours = Math.floor(whole / 3600);
  const minutes = Math.floor((whole % 3600) / 60);
  const seconds = whole % 60;
  const h = String(hours);
  const mm = String(minutes).padStart(2, '0');
  const ss = String(seconds).padStart(2, '0');
  if (hasFraction && tenths > 0) {
    return `${h}:${mm}:${ss}.${tenths}`;
  }
  return `${h}:${mm}:${ss}`;
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

  // Normalize helper: lower, strip diacritics, remove punctuation, collapse spaces
  const normalize = (s) => {
    return String(s || '')
      .toLowerCase()
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .replace(/[_'`’‑–—-]+/g, ' ')
      .replace(/[^a-z0-9\s]/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  };

  const textNorm = normalize(openingText);

  // 1) Try synonyms first
  const synonymEntries = Object.keys(OPENING_FAMILY_SYNONYMS || {}).map(k => ({
    keyNorm: normalize(k),
    canonical: OPENING_FAMILY_SYNONYMS[k]
  }));
  for (const entry of synonymEntries) {
    if (entry.keyNorm && textNorm.includes(entry.keyNorm)) {
      const canonical = entry.canonical;
      if (canonical && map[canonical]) {
        return { familyName: canonical, familyUrl: map[canonical] };
      }
    }
  }

  // 2) Direct map keys with normalization
  for (const key in map) {
    if (!Object.prototype.hasOwnProperty.call(map, key)) continue;
    const keyNorm = normalize(key);
    if (keyNorm && textNorm.includes(keyNorm)) {
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
