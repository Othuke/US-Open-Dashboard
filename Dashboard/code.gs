/**
 * @OnlyCurrentDoc
 */

const SPREADSHEET_NAME = "usopen";
const SHEET_NAME = "usopen";

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('US Open Dashboard')
    .addItem('Open Dashboard', 'showDashboard')
    .addToUi();
}

function showDashboard() {
  const html = HtmlService.createHtmlOutputFromFile('index')
    .setWidth(1600)
    .setHeight(1024);
  SpreadsheetApp.getUi().showModalDialog(html, 'US Open Dashboard');
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setWidth(1600)
    .setHeight(1024);
}

/**
 * Helper function to fetch all raw data from the spreadsheet.
 */
function fetchRawData() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    const headers = values.shift();

    const allData = values.map(row => {
      const obj = {};
      headers.forEach((header, i) => {
        obj[header.trim()] = row[i];
      });
      return obj;
    });
    return allData;
  } catch (e) {
    Logger.log("Error fetching raw data: " + e.message);
    return [];
  }
}

/**
 * Fetches all data from the spreadsheet and processes it for the dashboard.
 */
function getAllDashboardData() {
  const allData = fetchRawData();
  const processedData = processDashboardData(allData);
  return processedData;
}

/**
 * Processes data for all dashboard charts and cards.
 * @param {Array<Object>} filteredData The data filtered by the user's selection.
 * @param {Array<Object>} unfilteredData The full, original dataset.
 */
/**
 * Processes data for all dashboard charts and cards.
 * @param {Array<Object>} allData The full, original dataset.
 */
function processDashboardData(allData) {
  // We still need the list of years for the Sinner vs Alcaraz chart,
  // so we'll generate it here.
  const years = [...new Set(allData.map(row => row.Date.getFullYear().toString()))].sort();
  
  return {
    cardData: getCardData(allData),
    sinnerAlcarazData: getSinnerAlcarazData(allData),
    top10WinsData: getTop10WinsData(allData),
    matchOutcomeData: getMatchOutcomeData(allData),
    dominantWinsData: getDominantWinsData(allData),
    winRateData: getWinRateByRankData(allData),
    semiFinalData: getSemiFinalData(allData),
    years: years, // This is needed to populate chart labels
  };
}

// === Card Data Processing ===

function getCardData(data) {
  try {
    const totalMatches = data.length;
    const allPlayers = data.flatMap(row => [row.Winner, row.Loser]);
    const uniquePlayers = new Set(allPlayers.filter(p => p)).size;
    const totalSets = data.reduce((sum, row) => sum + (row.Wsets || 0) + (row.Lsets || 0), 0);
    const avgMatchLength = totalMatches > 0 ? (totalSets / totalMatches).toFixed(1) : 0;
    const fiveSetters = data.filter(row => (row.Wsets + row.Lsets) === 5).length;
    const upsets = data.filter(row => (row.WRank && row.LRank) && (parseInt(row.WRank) > parseInt(row.LRank))).length;
    
    return { totalMatches, uniquePlayers, avgMatchLength, fiveSetters, upsets };
  } catch (e) {
    Logger.log("Error in getCardData: " + e.message);
    return {};
  }
}

// === Chart Data Processing ===

function getSinnerAlcarazData(data) {
  try {
    const yearWins = {};
    const yearStages = {};
    const stageToNumeric = {
      '1st Round': 1, '2nd Round': 2, '3rd Round': 3, '4th Round': 4,
      'Quarterfinals': 5, 'Semifinals': 6, 'Final': 7, 'Winner': 8
    };

    data.forEach(row => {
      const year = row.Date instanceof Date ? row.Date.getFullYear() : null;
      if (!year) return;
      
      const winner = row.Winner ? row.Winner.trim() : null;
      const loser = row.Loser ? row.Loser.trim() : null;
      const round = row.Round ? row.Round.trim() : null;
      
      if (winner && (winner === "Sinner J." || winner === "Alcaraz C.")) {
        if (!yearWins[year]) yearWins[year] = { "Sinner J.": 0, "Alcaraz C.": 0 };
        yearWins[year][winner]++;
      }

      const processPlayerStage = (player, currentYear, currentRound, isWinner) => {
        if (player && (player === "Sinner J." || player === "Alcaraz C.")) {
          if (!yearStages[currentYear]) yearStages[currentYear] = { "Sinner J.": 0, "Alcaraz C.": 0 };
          const numericStage = isWinner ? 8 : (stageToNumeric[currentRound] || 0);
          if (numericStage > yearStages[currentYear][player]) {
            yearStages[currentYear][player] = numericStage;
          }
        }
      };
      
      processPlayerStage(winner, year, round, round === 'Final' && winner);
      processPlayerStage(loser, year, round, false);
    });
    
    const years = Object.keys(yearWins).sort();
    const sinnerWins = years.map(y => yearWins[y]["Sinner J."] || 0);
    const alcarazWins = years.map(y => yearWins[y]["Alcaraz C."] || 0);
    const sinnerStages = years.map(y => yearStages[y]["Sinner J."] || 0);
    const alcarazStages = years.map(y => yearStages[y]["Alcaraz C."] || 0);
    
    return { years, sinnerWins, alcarazWins, sinnerStages, alcarazStages };
  } catch (e) {
    Logger.log("Error in getSinnerAlcarazData: " + e.message);
    return { years: [], sinnerWins: [], alcarazWins: [], sinnerStages: [], alcarazStages: [] };
  }
}

function getTop10WinsData(data) {
  try {
    const playerStats = {};

    data.forEach(row => {
      const winner = row.Winner ? row.Winner.trim() : null;
      const wsets = row.Wsets || 0;
      const loser = row.Loser ? row.Loser.trim() : null;
      const lsets = row.Lsets || 0;

      if (winner) {
        if (!playerStats[winner]) {
          playerStats[winner] = { matches: 0, sets: 0 };
        }
        playerStats[winner].matches++;
        playerStats[winner].sets += wsets;
      }

      if (loser) {
        if (!playerStats[loser]) {
          playerStats[loser] = { matches: 0, sets: 0 };
        }
        playerStats[loser].sets += lsets;
      }
    });

    const sortedPlayers = Object.entries(playerStats)
      .sort(([, a], [, b]) => b.matches - a.matches)
      .slice(0, 10);

    const players = sortedPlayers.map(p => p[0]);
    const matchesWon = sortedPlayers.map(p => p[1].matches);
    const setsWon = sortedPlayers.map(p => p[1].sets);
    
    return { players, matchesWon, setsWon };
  } catch (e) {
    Logger.log("Error in getTop10WinsData: " + e.message);
    return { players: [], matchesWon: [], setsWon: [] };
  }
}

function getMatchOutcomeData(data) {
  try {
    const outcomes = {};
    data.forEach(row => {
      const outcome = row.Comment ? row.Comment.trim() : 'Unknown';
      outcomes[outcome] = (outcomes[outcome] || 0) + 1;
    });
    return outcomes;
  } catch (e) {
    Logger.log("Error in getMatchOutcomeData: " + e.message);
    return {};
  }
}

function getDominantWinsData(data) {
  try {
    const dominantWins = {};

    data.forEach(row => {
      const winner = row.Winner ? row.Winner.trim() : null;
      if (!winner) return;

      const wsets = row.Wsets || 0;
      const lsets = row.Lsets || 0;

      if (lsets === 0) { // Straight sets win
        let gameMargin = 0;
        gameMargin += (row.W1 || 0) - (row.L1 || 0);
        gameMargin += (row.W2 || 0) - (row.L2 || 0);
        gameMargin += (row.W3 || 0) - (row.L3 || 0);
        gameMargin += (row.W4 || 0) - (row.L4 || 0);
        gameMargin += (row.W5 || 0) - (row.L5 || 0);

        if (gameMargin >= 6) {
          dominantWins[winner] = (dominantWins[winner] || 0) + 1;
        }
      }
    });

    const sortedPlayers = Object.entries(dominantWins)
      .sort(([, a], [, b]) => b - a)
      .slice(0, 10);

    const players = sortedPlayers.map(p => p[0]);
    const wins = sortedPlayers.map(p => p[1]);
    return { players, wins };
  } catch (e) {
    Logger.log("Error in getDominantWinsData: " + e.message);
    return { players: [], wins: [] };
  }
}

function getWinRateByRankData(data) {
  try {
    const ranks = {
      '1-10': { wins: 0, losses: 0 },
      '10-50': { wins: 0, losses: 0 },
      '50-100': { wins: 0, losses: 0 },
      'Outside top 100': { wins: 0, losses: 0 }
    };

    const getRankBin = (rank) => {
      const r = parseInt(rank);
      if (r >= 1 && r <= 10) return '1-10';
      if (r > 10 && r <= 50) return '10-50';
      if (r > 50 && r <= 100) return '50-100';
      if (r > 100) return 'Outside top 100';
      return null;
    };
    
    data.forEach(row => {
      const wRank = row.WRank;
      const lRank = row.LRank;

      const winnerBin = getRankBin(wRank);
      if (winnerBin) ranks[winnerBin].wins++;
      
      const loserBin = getRankBin(lRank);
      if (loserBin) ranks[loserBin].losses++;
    });

    const winRates = Object.entries(ranks).map(([tier, stats]) => {
      const totalMatches = stats.wins + stats.losses;
      const winRate = totalMatches > 0 ? (stats.wins / totalMatches) * 100 : 0;
      return { tier, winRate: parseFloat(winRate.toFixed(1)) };
    });

    return winRates;
  } catch (e) {
    Logger.log("Error in getWinRateByRankData: " + e.message);
    return [];
  }
}

function getSemiFinalData(data) {
  try {
    const semiFinalPlayers = {};

    data.forEach(row => {
      if (row.Round && row.Round.trim() === "Semifinals") {
        const winner = row.Winner ? row.Winner.trim() : null;
        const loser = row.Loser ? row.Loser.trim() : null;
        if (winner) semiFinalPlayers[winner] = (semiFinalPlayers[winner] || 0) + 1;
        if (loser) semiFinalPlayers[loser] = (semiFinalPlayers[loser] || 0) + 1;
      }
    });

    const sortedPlayers = Object.entries(semiFinalPlayers)
      .sort(([, a], [, b]) => b - a)
      .slice(0, 5);

    const players = sortedPlayers.map(p => p[0]);
    const appearances = sortedPlayers.map(p => p[1]);
    return { players, appearances };
  } catch (e) {
    Logger.log("Error in getSemiFinalData: " + e.message);
    return { players: [], appearances: [] };
  }
}
