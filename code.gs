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
  'Time Control', 'Time Class', 'Rules', 'ECO Opening', 'Rated', 'Tournament URL', 'Match URL',
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
  const pgnDataHeaders = ['PGN_Moves', 'PGN_Move_Count', 'PGN_Game_Outcome'];
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
      return path.split('.').reduce((current, key) => current && current[key], obj) || defaultValue;
    } catch {
      return defaultValue;
    }
  };
  
  const startTimeFormatted = game.start_time ? new Date(game.start_time * 1000).toLocaleString() : '';
  const endTimeFormatted = game.end_time ? new Date(game.end_time * 1000).toLocaleString() : '';
  
  // Complete Chess.com API data extraction
  const apiData = [
    safeGet(game, 'url'), safeGet(game, 'fen'), safeGet(game, 'start_time'), safeGet(game, 'end_time'),
    startTimeFormatted, endTimeFormatted, safeGet(game, 'time_control'), safeGet(game, 'time_class'),
    safeGet(game, 'rules'), safeGet(game, 'eco'), safeGet(game, 'rated'), safeGet(game, 'tournament'),
    safeGet(game, 'match'), safeGet(game, 'white.username'), safeGet(game, 'white.rating'),
    safeGet(game, 'white.result'), safeGet(game, 'white.@id'), safeGet(game, 'white.uuid'),
    safeGet(game, 'white.country'), safeGet(game, 'black.username'), safeGet(game, 'black.rating'),
    safeGet(game, 'black.result'), safeGet(game, 'black.@id'), safeGet(game, 'black.uuid'),
    safeGet(game, 'black.country'), safeGet(game, 'accuracies.white'), safeGet(game, 'accuracies.black'),
    safeGet(game, 'initial_setup'), safeGet(game, 'pgn')
  ];
  
  // Complete PGN header data extraction
  const pgnHeaderData = ALL_PGN_HEADERS.map(header => pgnData.headers[header] || '');
  
  // PGN parsed data
  const pgnDataArray = [pgnData.moves, pgnData.moveCount, pgnData.gameOutcome];
  
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
  const pgnDataRange = sheet.getRange(1, apiHeaderCount + pgnHeaderCount + 1, 1, 3);
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
  
  // Format move count
  formatColumn('PGN_Move_Count', '#,##0');
  formatColumn('PGN_PlyCount', '#,##0');
  
  // Format boolean columns
  const booleanHeaders = ['Rated', 'PGN_SetUp', 'PGN_WhiteIsComp', 'PGN_BlackIsComp'];
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
      result.moves = movesText;
      
      // Count full moves (each number represents one full move)
      const moveMatches = movesText.match(/\d+\./g);
      result.moveCount = moveMatches ? moveMatches.length : 0;
      
      // Extract game outcome with more robust regex
      const outcomeMatch = movesText.match(/(1-0|0-1|1\/2-1\/2|\*)(?:\s*$)/);
      if (outcomeMatch) {
        result.gameOutcome = outcomeMatch[1];
      }
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
