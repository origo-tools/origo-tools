/**
 * Origo Bakery — Production Forecasting Engine (Apps Script)
 *
 * Deploy as Web App (Execute as: Me, Access: Anyone)
 *
 * Required Advanced Services:
 *   - BigQuery API (v2)
 *
 * SETUP:
 * 1. Create a new Apps Script project at script.google.com
 * 2. Paste this code
 * 3. Enable BigQuery Advanced Service: Services → BigQuery API
 * 4. Deploy → New deployment → Web app
 * 5. Copy the deployment URL into production-estimates.html
 */

// ============================================================
// CONFIGURATION
// ============================================================

const CONFIG = {
  BQ_PROJECT: 'origo-data-center',
  BQ_DATASET: 'Origo_Database',
  BQ_TABLE: 'Daily_Sales',

  SPREADSHEETS: {
    'Milà':     '17WvIllUHdvKm0LJw_ezupZdo1dhzl1Ic_pu-KSxGDJk',
    'Sant Pau':  '1vbn14j5kX5ZBZDS6meqE1Vc67VXU_Q2aLBllyRfi1tQ',
    'Sant Joan': '1nRavOoyMVrO8LlJUx6WknoQ3xT_hYp91eEKPpGOMOu8',
  },

  // Day tab names in the production sheets (0=Monday)
  DAY_TABS: ['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado', 'Domingo'],

  // Weighted average: most recent week gets highest weight
  WEEK_WEIGHTS: [0.50, 0.25, 0.15, 0.10],  // W1 (last week) to W4

  // Weather API (Open-Meteo, free, no key needed)
  WEATHER_URL: 'https://api.open-meteo.com/v1/forecast',
  BARCELONA_LAT: 41.39,
  BARCELONA_LON: 2.16,

  // Seasonal indices (1.0 = average, derived from 2025 data)
  SEASONAL_INDEX: {
    1: 0.95,  // Jan
    2: 0.97,  // Feb
    3: 1.00,  // Mar (baseline)
    4: 0.98,  // Apr
    5: 0.96,  // May
    6: 0.88,  // Jun
    7: 0.77,  // Jul
    8: 0.78,  // Aug
    9: 0.90,  // Sep
    10: 0.97, // Oct
    11: 0.98, // Nov
    12: 0.93, // Dec
  },

  // 2026 Catalan bank holidays (month-day format)
  HOLIDAYS_2026: [
    '01-01', '01-06', '04-03', '04-06', '05-01', '06-24',
    '08-15', '09-11', '09-24', '10-12', '11-01', '12-06',
    '12-08', '12-25', '12-26'
  ],

  // ── Sobras Integration (v2) ──────────────────────────────
  // URL of the Sobras API web app (Apps Script deployment)
  // Set this after deploying apps-script-sobras-api.js
  SOBRAS_API_URL: 'https://script.google.com/macros/s/AKfycbz1z5Yz-qyJSPfxINXMFnqBJwR-nNQ3NlTrACj04mwOgs_Xk0U_tdmjsLYI7eM7rV0/exec',  // e.g. 'https://script.google.com/macros/s/XXXXXX/exec'

  // Store ID mapping (sobras API uses lowercase IDs)
  SOBRAS_STORE_IDS: {
    'Milà': 'mila',
    'Sant Pau': 'santpau',
    'Sant Joan': 'santjoan',
  },

  // Minimum weeks of sobras data needed before applying adjustments
  MIN_SOBRAS_WEEKS: 2,

  // Number of past weeks of sobras data to fetch
  SOBRAS_HISTORY_DAYS: 28,  // 4 weeks

  // Step 6: Sold-out demand correction factors
  // When a product batch sells out early, actual demand was higher than recorded sales
  SOLDOUT_FACTORS: {
    earlyMorning: 1.30,   // B1 sold out before 11:00 → +30%
    midday: 1.20,         // B2 sold out before 13:00 → +20%
    afternoon: 1.10,      // B3 sold out before 15:00 → +10%
    lateAfternoon: 1.05,  // B3 sold out after 15:00 → +5%
  },

  // Step 7: Overproduction dampening thresholds
  // sobrasRatio = avg_sobras / avg_production_estimate
  SOBRAS_THRESHOLDS: {
    high: 0.25,     // ratio > 25% → significant overproduction
    medium: 0.15,   // ratio > 15% → moderate overproduction
    low: 0.05,      // ratio > 5%  → slight overproduction
  },
  SOBRAS_REDUCTIONS: {
    high: 0.85,     // -15%
    medium: 0.90,   // -10%
    low: 0.95,      // -5%
  },
};


// ============================================================
// PRODUCT ↔ SQUARE NAME ↔ CELL MAPPING
// ============================================================

/**
 * Maps production sheet product names to:
 *   - squareNames: array of matching item names in BigQuery/Square
 *   - cell: the cell in column D (TOTAL) for that product
 *   - type: 'bread' or 'pastry'
 *   - isFocaccia: if true, convert servings to whole focaccias (÷15, round up)
 *   - isNiu: if true, match all "NIU ..." seasonal variants (same product, different flavors)
 *   - lataSize: for pastry items, how many units per lata (tray)
 */
const PRODUCT_MAP_COMMON = [
  // === BREAD ===
  // squareNames: current Catalan names + older Spanish names for historical data
  // Half-bread suffixes (- 1/2, - Mitad) are stripped and counted as 0.5 units by the matching logic
  // Size suffixes (- Sencer, - Normal, - Porció, etc.) are stripped before matching
  { sheetName: 'ORIGO',           cell: 'D8',  type: 'bread', squareNames: ['Origo'] },
  { sheetName: 'ORIGO MOLDE',     cell: 'D9',  type: 'bread', squareNames: ['Motlle', 'Origo Molde', 'Molde'] },
  { sheetName: 'ORIGO INTEGRAL',  cell: 'D10', type: 'bread', squareNames: ['Origo Integral'] },
  { sheetName: 'ESPECIAL',        cell: 'D12', type: 'bread', squareNames: ['Pan Especial', 'Especial'] },
  { sheetName: 'LLAVORS',         cell: 'D13', type: 'bread', squareNames: ['Llavors', 'Pan de Semillas', 'Semillas'] },
  { sheetName: 'SÈGOL INTEGRAL',  cell: 'D16', type: 'bread', squareNames: ['Sègol integral', 'Pan de Centeno', 'Centeno integral'] },
  { sheetName: 'COPENHAGUE',      cell: 'D17', type: 'bread', squareNames: ['Copenhague', 'Copenhagen'] },
  { sheetName: 'BAGUETTE',        cell: 'D18', type: 'bread', squareNames: ['Baguette', 'Baguette Origo'] },
  { sheetName: 'FOCACCIA',        cell: 'D19', type: 'bread', squareNames: ['Focaccia'], isFocaccia: true },
  { sheetName: 'ESPELTA PETITA',  cell: 'D20', type: 'bread', squareNames: ['Espelta petita', 'Espelta'] },

  // === PASTRY (all 3 stores) ===
  { sheetName: 'Canela Twist',              cell: 'D29', type: 'pastry', lataSize: 12, squareNames: ['Canela Twist', 'Cinnamon Twist'] },
  { sheetName: 'Cardamomo twist',           cell: 'D30', type: 'pastry', lataSize: 12, squareNames: ['Cardamomo twist', 'Cardamom Twist', 'Cardamomo Twist'] },
  { sheetName: 'Vainilla Roll',             cell: 'D31', type: 'pastry', lataSize: 12, squareNames: ['Vainilla Roll', 'Vanilla Roll'] },
  { sheetName: 'Brioix clàssic',            cell: 'D32', type: 'pastry', lataSize: 12, squareNames: ['Brioix clàssic', 'Brioix clàssic (hamburguesa)', 'Brioche Clásico', 'Brioche Classic', 'Brioche clasico'] },
  { sheetName: 'Croissant',                 cell: 'D34', type: 'pastry', lataSize: 12, squareNames: ['Croissant', 'Croissant Mantequilla'] },
  { sheetName: 'Pain au chocolat',          cell: 'D35', type: 'pastry', lataSize: 12, squareNames: ['Pain au chocolat'] },
  { sheetName: 'Napolitana pernil i formatge', cell: 'D36', type: 'pastry', lataSize: 12, squareNames: ['Napolitana pernil i formatge', 'Napolitana jamón y queso', 'Napolitana Jamón y Queso', 'Napo Jamon y queso', 'Napolitana'] },
  { sheetName: 'Mico',                      cell: 'D38', type: 'pastry', lataSize: 6,  squareNames: ['Mico', 'Mico / Babka', 'Babka Chocolate'] },
  { sheetName: 'Niu de temporada',          cell: 'D39', type: 'pastry', lataSize: 12, squareNames: ['NIU de temporada', 'Niu de temporada', 'Niu Temporada', 'Niu', 'NIU'], isNiu: true },

  // === PASTRY (previously St Joan only, now all stores) ===
  { sheetName: "Croissant d'ametlla",       cell: 'D41', type: 'pastry', lataSize: 12, squareNames: ["Croissant d'ametlla", "Croissant Almendra", "Croissant d'Ametlla", "Croissant de almendras"] },
  { sheetName: 'Cookie chocolate',          cell: 'D43', type: 'pastry', lataSize: 0,  squareNames: ['Cookie chocolate 70% Venezuela', 'Cookie chocolate 70% Nicaragua', 'Cookie chocolate', 'Cookie Chocolate', 'Cookie'] },
  { sheetName: 'Pastry del mes',            cell: 'D44', type: 'pastry', lataSize: 0,  squareNames: [], isPastryDelMes: true },
  { sheetName: 'Pastís llimona i ametlla',  cell: 'D46', type: 'pastry', lataSize: 0,  squareNames: ['Pastís llimona i ametlla', 'Pastís Llimona', 'Cake Limón y Almendra'] },
];

/**
 * Build the product map, injecting the current "Pastry del mes" Square name.
 * The pastryDelMes name can come from:
 *   1. A parameter passed by the frontend (highest priority)
 *   2. Script Properties (set via the 'setPastryDelMes' action)
 *   3. Falls back to empty (no sales matched)
 */
function getProductMap(store, pastryDelMesName) {
  const map = PRODUCT_MAP_COMMON.map(p => ({ ...p }));

  // Resolve the pastry del mes name
  const name = pastryDelMesName || getPastryDelMesSetting();
  if (name) {
    const pdm = map.find(p => p.isPastryDelMes);
    if (pdm) {
      pdm.squareNames = [name];
    }
  }

  return map;
}

/**
 * Get the stored "Pastry del mes" Square name from Script Properties.
 */
function getPastryDelMesSetting() {
  try {
    const props = PropertiesService.getScriptProperties();
    return props.getProperty('pastryDelMes') || '';
  } catch (e) {
    return '';
  }
}

/**
 * Save the "Pastry del mes" Square name to Script Properties.
 */
function setPastryDelMesSetting(name) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('pastryDelMes', name || '');
  return { success: true, pastryDelMes: name };
}


// ============================================================
// WEB APP ENDPOINTS
// ============================================================

function doGet(e) {
  const action = (e.parameter.action || '').toLowerCase();
  const store = e.parameter.store || '';
  const dateStr = e.parameter.date || '';  // YYYY-MM-DD
  const pastryDelMes = e.parameter.pastryDelMes || '';  // Current pastry of the month

  let result;
  try {
    switch (action) {
      case 'estimates':
        result = getEstimates(store, dateStr, pastryDelMes);
        break;
      case 'week':
        result = getWeekEstimates(store, dateStr, pastryDelMes);
        break;
      case 'weather':
        result = getWeatherForecast(dateStr);
        break;
      case 'stores':
        result = { stores: Object.keys(CONFIG.SPREADSHEETS) };
        break;
      case 'getpastrydelmes':
        result = { pastryDelMes: getPastryDelMesSetting() };
        break;
      case 'generateddates':
        result = getGeneratedDates();
        break;
      default:
        result = { error: 'Unknown action. Use: estimates, week, weather, stores, getPastryDelMes, generatedDates' };
    }
  } catch (err) {
    result = { error: err.message, stack: err.stack };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: 'Invalid JSON body' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const action = (body.action || '').toLowerCase();
  let result;

  try {
    switch (action) {
      case 'sync':
        result = syncToSheet(body.store, body.dayOfWeek, body.estimates);
        break;
      case 'setpastrydelmes':
        result = setPastryDelMesSetting(body.name);
        break;
      case 'recordapproval':
        result = recordRafaApproval(body.store, body.dayOfWeek, body.estimates);
        break;
      default:
        result = { error: 'Unknown action. Use: sync, setPastryDelMes, recordApproval' };
    }
  } catch (err) {
    result = { error: err.message, stack: err.stack };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}


// ============================================================
// FORECASTING ENGINE
// ============================================================

/**
 * Get production estimates for a single day
 */
function getEstimates(store, dateStr, pastryDelMes) {
  if (!store || !dateStr) throw new Error('Missing store or date parameter');
  if (!CONFIG.SPREADSHEETS[store]) throw new Error('Unknown store: ' + store);

  const targetDate = new Date(dateStr + 'T00:00:00');
  const dayOfWeek = targetDate.getDay(); // 0=Sun, 1=Mon, ... 6=Sat
  const month = targetDate.getMonth() + 1;

  // 1. Get historical sales data from BigQuery (last 4 weeks, same day of week)
  const historicalData = queryHistoricalSales(store, targetDate, dayOfWeek);

  // 2. Get weather forecast
  const weather = getWeatherForDate(dateStr);

  // 3. Check if it's a holiday
  const isHoliday = checkHoliday(dateStr);

  // 4. Get sobras history (v2 — for steps 6 and 7)
  const sobrasData = querySobrasHistory(store, targetDate, dayOfWeek);
  const sobrasAvailable = sobrasData && sobrasData.weeksAvailable >= CONFIG.MIN_SOBRAS_WEEKS;

  // 5. Calculate estimates for each product
  const productMap = getProductMap(store, pastryDelMes);
  const estimates = [];

  for (const product of productMap) {
    // Find matching sobras entries for this product
    let productSobrasEntries = null;
    if (sobrasAvailable) {
      // Match by product sheetName or squareNames (case-insensitive)
      const namesToCheck = [product.sheetName, ...product.squareNames].map(n => n.toLowerCase());
      for (const [sobrasProduct, entries] of Object.entries(sobrasData.productSobras)) {
        if (namesToCheck.includes(sobrasProduct.toLowerCase())) {
          productSobrasEntries = entries;
          break;
        }
      }
    }

    const estimate = calculateProductEstimate(product, historicalData, dayOfWeek, month, weather, isHoliday, productSobrasEntries);

    // 6. Query hourly batch distribution for pastry products
    if (product.type === 'pastry') {
      const batchDist = queryBatchDistribution(store, targetDate, dayOfWeek, product);
      if (batchDist) {
        estimate.batchSplit = batchDist;
      }
      // If null, frontend will use a sensible default (33/33/34)
    }

    estimates.push(estimate);
  }

  return {
    store,
    date: dateStr,
    dayOfWeek: CONFIG.DAY_TABS[dayOfWeek === 0 ? 6 : dayOfWeek - 1],
    dayIndex: dayOfWeek,
    isHoliday,
    weather,
    seasonalIndex: CONFIG.SEASONAL_INDEX[month],
    sobrasWeeksAvailable: sobrasData ? sobrasData.weeksAvailable : 0,
    estimates,
    generatedAt: new Date().toISOString(),
  };
}

/**
 * Get estimates for an entire week (Monday to Sunday)
 */
function getWeekEstimates(store, startDateStr, pastryDelMes) {
  if (!store) throw new Error('Missing store parameter');
  if (!CONFIG.SPREADSHEETS[store]) throw new Error('Unknown store: ' + store);

  // Find next Monday if no date provided
  let startDate;
  if (startDateStr) {
    startDate = new Date(startDateStr + 'T00:00:00');
  } else {
    startDate = getNextMonday();
  }

  const weekEstimates = {};
  for (let i = 0; i < 7; i++) {
    const d = new Date(startDate);
    d.setDate(d.getDate() + i);
    const ds = formatDate(d);
    const dayName = CONFIG.DAY_TABS[i];
    try {
      weekEstimates[dayName] = getEstimates(store, ds, pastryDelMes);
    } catch (err) {
      weekEstimates[dayName] = { error: err.message, date: ds };
    }
  }

  return {
    store,
    weekStart: formatDate(startDate),
    days: weekEstimates,
    generatedAt: new Date().toISOString(),
  };
}


/**
 * Parse a BigQuery item name: strip known suffixes and detect half-bread.
 *
 * Examples:
 *   "Origo - Sencer"          → { baseName: "origo", isHalf: false }
 *   "Origo - 1/2"             → { baseName: "origo", isHalf: true }
 *   "Motlle - Mitad"          → { baseName: "motlle", isHalf: true }
 *   "Canela Twist - Normal"   → { baseName: "canela twist", isHalf: false }
 *   "Cookie chocolate 70% Venezuela - Normal" → { baseName: "cookie chocolate 70% venezuela", isHalf: false }
 *   "Brioix clàssic (hamburguesa)" → { baseName: "brioix clàssic", isHalf: false }
 *   "Pastís llimona i ametlla - Porció" → { baseName: "pastís llimona i ametlla", isHalf: false }
 *   "Pain au chocolat d´ametlla" → { baseName: "pain au chocolat d´ametlla", isHalf: false }
 *   "Focaccia - Porció"       → { baseName: "focaccia", isHalf: false }
 */
function parseItemName(rawName) {
  let name = rawName.trim();
  let isHalf = false;

  // Check for half-bread suffixes (must be checked before general suffix stripping)
  const halfSuffixes = [' - 1/2', ' - Mitad', ' - mitad'];
  for (const suffix of halfSuffixes) {
    if (name.endsWith(suffix)) {
      name = name.slice(0, -suffix.length);
      isHalf = true;
      break;
    }
  }

  // Strip known size/format suffixes (case-insensitive check, preserve original for slicing)
  if (!isHalf) {
    const sizeSuffixes = [' - Sencer', ' - sencer', ' - Normal', ' - normal',
                          ' - Porció', ' - porció', ' - Regular', ' - regular',
                          ' - Ración', ' - ración', ' - Entero', ' - entero'];
    for (const suffix of sizeSuffixes) {
      if (name.endsWith(suffix)) {
        name = name.slice(0, -suffix.length);
        break;
      }
    }
  }

  // Strip parenthetical suffixes like "(hamburguesa)"
  name = name.replace(/\s*\([^)]*\)\s*$/, '');

  // Strip "Entrepà - Bocadillo - " prefix for bocadillos
  name = name.replace(/^Entrepà\s*-\s*Bocadillo\s*-\s*/i, '');

  return {
    baseName: name.trim().toLowerCase(),
    isHalf,
  };
}


/**
 * Core forecast calculation for a single product (v2 with sobras feedback)
 *
 * Pipeline:
 *   Step 1: Weighted average of 4 weeks sales
 *   Step 2: Trend adjustment (±15%)
 *   Step 3: Seasonal adjustment
 *   Step 4: Holiday adjustment
 *   Step 5: Weather adjustment
 *   Step 6: Sold-out demand correction (v2 — from sobras data)
 *   Step 7: Overproduction dampening (v2 — from sobras data)
 */
function calculateProductEstimate(product, historicalData, dayOfWeek, month, weather, isHoliday, sobrasEntries) {
  // Collect sales data for this product across the 4 weeks
  const weekSales = [null, null, null, null]; // W1 (most recent) to W4

  // Build a set of lowercase squareNames for exact matching after stripping
  const squareNamesLower = product.squareNames.map(n => n.toLowerCase());

  for (let w = 0; w < 4; w++) {
    const weekData = historicalData[w] || {};
    for (const [itemName, qty] of Object.entries(weekData)) {
      const parsed = parseItemName(itemName);

      // Check for exact match against any configured squareName
      // For focaccia and niu: match all variants starting with the base name
      // (e.g. "Focaccia - Temporada", "NIU de figues i festuc" are seasonal variants)
      const isMatch = squareNamesLower.some(sn => parsed.baseName === sn) ||
        (product.isFocaccia && parsed.baseName.startsWith('focaccia')) ||
        (product.isNiu && parsed.baseName.startsWith('niu'));

      if (isMatch) {
        // Apply 0.5x multiplier for half-bread sales
        const effectiveQty = parsed.isHalf ? qty * 0.5 : qty;
        weekSales[w] = (weekSales[w] || 0) + effectiveQty;
      }
    }
  }

  // Step 1: Weighted average
  let weightedSum = 0;
  let weightSum = 0;
  for (let w = 0; w < 4; w++) {
    if (weekSales[w] !== null && weekSales[w] !== undefined) {
      weightedSum += weekSales[w] * CONFIG.WEEK_WEIGHTS[w];
      weightSum += CONFIG.WEEK_WEIGHTS[w];
    }
  }

  let baseEstimate = weightSum > 0 ? weightedSum / weightSum : 0;

  // Step 2: Trend adjustment (compare W1 vs W3-W4 average)
  let trendMultiplier = 1.0;
  if (weekSales[0] !== null && (weekSales[2] !== null || weekSales[3] !== null)) {
    const recentWeek = weekSales[0];
    const olderWeeks = [];
    if (weekSales[2] !== null) olderWeeks.push(weekSales[2]);
    if (weekSales[3] !== null) olderWeeks.push(weekSales[3]);
    const olderAvg = olderWeeks.reduce((a, b) => a + b, 0) / olderWeeks.length;

    if (olderAvg > 0) {
      const trendRatio = recentWeek / olderAvg;
      // Cap trend adjustment to ±15%
      trendMultiplier = Math.max(0.85, Math.min(1.15, trendRatio));
    }
  }

  // Step 3: Seasonal adjustment
  const seasonalIndex = CONFIG.SEASONAL_INDEX[month] || 1.0;
  // We apply seasonal relative to March (1.0 baseline)
  // Only apply if we're outside the base period
  const seasonalMultiplier = seasonalIndex;

  // Step 4: Holiday adjustment
  let holidayMultiplier = 1.0;
  if (isHoliday) {
    // If it's a weekday holiday, treat like Saturday (+15% avg uplift from data)
    if (dayOfWeek >= 1 && dayOfWeek <= 5) {
      holidayMultiplier = 1.15;
    }
  }

  // Step 5: Weather adjustment
  let weatherMultiplier = 1.0;
  if (weather && weather.precipitation_mm !== undefined) {
    if (weather.precipitation_mm > 10) {
      weatherMultiplier = 0.90; // Heavy rain: -10%
    } else if (weather.precipitation_mm > 3) {
      weatherMultiplier = 0.95; // Light/moderate rain: -5%
    }
  }

  // Combine steps 1-5 multipliers
  let estimateAfterStep5 = baseEstimate * trendMultiplier * seasonalMultiplier * holidayMultiplier * weatherMultiplier;

  // Step 6: Sold-out demand correction (v2)
  let soldOutMultiplier = 1.0;
  if (sobrasEntries && sobrasEntries.length > 0) {
    soldOutMultiplier = calcSoldOutCorrection(product, sobrasEntries);
  }

  let estimateAfterStep6 = estimateAfterStep5 * soldOutMultiplier;

  // Step 7: Overproduction dampening (v2)
  // Uses estimateAfterStep6 as the "production baseline" for sobras ratio calculation
  let overproductionMultiplier = 1.0;
  if (sobrasEntries && sobrasEntries.length > 0) {
    overproductionMultiplier = calcOverproductionDampening(product, sobrasEntries, estimateAfterStep6);
  }

  let rawEstimate = estimateAfterStep6 * overproductionMultiplier;

  // For focaccia: convert servings to whole focaccias
  let finalEstimate;
  if (product.isFocaccia) {
    finalEstimate = Math.ceil(rawEstimate / 15);
  } else if (product.lataSize && product.lataSize > 0) {
    // Round pastry to nearest lata multiple
    finalEstimate = Math.ceil(rawEstimate / product.lataSize) * product.lataSize;
  } else {
    finalEstimate = Math.round(rawEstimate);
  }

  // Minimum of 0
  finalEstimate = Math.max(0, finalEstimate);

  return {
    product: product.sheetName,
    cell: product.cell,
    type: product.type,
    estimate: finalEstimate,
    rawEstimate: Math.round(rawEstimate * 10) / 10,
    weekSales: weekSales,
    factors: {
      trend: Math.round(trendMultiplier * 100) / 100,
      seasonal: seasonalMultiplier,
      holiday: holidayMultiplier,
      weather: weatherMultiplier,
      soldOut: soldOutMultiplier,
      overproduction: overproductionMultiplier,
    },
    sobrasData: sobrasEntries ? {
      weeksWithData: sobrasEntries.length,
      hasSoldOuts: sobrasEntries.some(e => e.b1SoldOut || e.b2SoldOut || e.b3SoldOut),
      avgSobras: sobrasEntries.filter(e => e.sobras !== null && e.sobras !== undefined && e.sobras !== '')
        .length > 0
        ? Math.round(
            sobrasEntries.filter(e => e.sobras !== null && e.sobras !== undefined && e.sobras !== '')
              .map(e => parseFloat(e.sobras) || 0)
              .reduce((a, b) => a + b, 0) /
            sobrasEntries.filter(e => e.sobras !== null && e.sobras !== undefined && e.sobras !== '').length * 10
          ) / 10
        : null,
    } : null,
    isFocaccia: product.isFocaccia || false,
    lataSize: product.lataSize || null,
  };
}


// ============================================================
// BIGQUERY DATA ACCESS
// ============================================================

/**
 * Query BigQuery for historical sales, returning data for each of the 4 prior weeks
 * for the same day of the week.
 */
function queryHistoricalSales(store, targetDate, dayOfWeek) {
  const results = [];

  // Map store name to BigQuery store name
  const bqStoreName = store; // Already matches: 'Milà', 'Sant Pau', 'Sant Joan'

  for (let weekOffset = 1; weekOffset <= 4; weekOffset++) {
    const refDate = new Date(targetDate);
    refDate.setDate(refDate.getDate() - (weekOffset * 7));
    const dateStr = formatDate(refDate);

    const sql = `
      SELECT Item_Name, SUM(Quantity) as total_qty
      FROM \`${CONFIG.BQ_PROJECT}.${CONFIG.BQ_DATASET}.${CONFIG.BQ_TABLE}\`
      WHERE Store = '${bqStoreName}'
        AND DATE(Date) = '${dateStr}'
        AND Quantity > 0
      GROUP BY Item_Name
    `;

    try {
      const request = { query: sql, useLegacySql: false };
      const queryResults = BigQuery.Jobs.query(request, CONFIG.BQ_PROJECT);

      const weekData = {};
      if (queryResults.rows) {
        for (const row of queryResults.rows) {
          const itemName = row.f[0].v;
          const qty = parseFloat(row.f[1].v) || 0;
          weekData[itemName] = qty;
        }
      }
      results.push(weekData);
    } catch (err) {
      Logger.log('BigQuery error for week ' + weekOffset + ': ' + err.message);
      results.push({});
    }
  }

  return results;
}


// ============================================================
// BATCH DISTRIBUTION (hourly sales patterns from BigQuery)
// ============================================================

/**
 * Query BigQuery for hourly sales distribution of a specific product.
 * Returns the percentage of sales in each batch window:
 *   B1: 08:00–11:00
 *   B2: 11:00–13:00
 *   B3: 13:00–20:30 (weekdays) or 13:00–19:00 (weekends)
 *
 * Uses the last 4 occurrences of the same weekday.
 *
 * @param {string} store - Store name
 * @param {Date} targetDate - The date we're forecasting for
 * @param {number} dayOfWeek - JS day of week (0=Sun, 1=Mon, ..., 6=Sat)
 * @param {object} product - Product definition from PRODUCT_MAP
 * @returns {object} { b1Pct, b2Pct, b3Pct } or null if insufficient data
 */
function queryBatchDistribution(store, targetDate, dayOfWeek, product) {
  if (product.type !== 'pastry') return null;

  // Build the list of names to match (lowercase for LOWER() in SQL)
  const names = product.squareNames.map(n => n.toLowerCase());
  if (names.length === 0 && !product.isFocaccia && !product.isNiu) return null;

  // Get the 4 reference dates (same weekday, previous 4 weeks)
  const refDates = [];
  for (let w = 1; w <= 4; w++) {
    const d = new Date(targetDate);
    d.setDate(d.getDate() - (w * 7));
    refDates.push(formatDate(d));
  }

  const dateList = refDates.map(d => `'${d}'`).join(',');

  // Build name matching clause
  let nameClause;
  if (product.isFocaccia) {
    nameClause = "LOWER(Item_Name) LIKE 'focaccia%'";
  } else if (product.isNiu) {
    nameClause = "LOWER(Item_Name) LIKE 'niu%'";
  } else {
    const escapedNames = names.map(n => `'${n.replace(/'/g, "\\'")}'`).join(',');
    nameClause = `LOWER(Item_Name) IN (${escapedNames})`;
  }

  // Determine batch 3 end hour based on weekday/weekend
  const isWeekend = (dayOfWeek === 0 || dayOfWeek === 6);
  const b3EndHour = isWeekend ? 19 : 21; // 20:30 rounded up to 21 for HOUR() comparison

  const sql = `
    SELECT
      SUM(CASE WHEN EXTRACT(HOUR FROM Date) >= 8 AND EXTRACT(HOUR FROM Date) < 11 THEN Quantity ELSE 0 END) as b1_sales,
      SUM(CASE WHEN EXTRACT(HOUR FROM Date) >= 11 AND EXTRACT(HOUR FROM Date) < 13 THEN Quantity ELSE 0 END) as b2_sales,
      SUM(CASE WHEN EXTRACT(HOUR FROM Date) >= 13 AND EXTRACT(HOUR FROM Date) < ${b3EndHour} THEN Quantity ELSE 0 END) as b3_sales,
      SUM(Quantity) as total_sales
    FROM \`${CONFIG.BQ_PROJECT}.${CONFIG.BQ_DATASET}.${CONFIG.BQ_TABLE}\`
    WHERE Store = '${store}'
      AND DATE(Date) IN (${dateList})
      AND ${nameClause}
      AND Quantity > 0
  `;

  try {
    const request = { query: sql, useLegacySql: false };
    const queryResults = BigQuery.Jobs.query(request, CONFIG.BQ_PROJECT);

    if (queryResults.rows && queryResults.rows.length > 0) {
      const row = queryResults.rows[0];
      const b1 = parseFloat(row.f[0].v) || 0;
      const b2 = parseFloat(row.f[1].v) || 0;
      const b3 = parseFloat(row.f[2].v) || 0;
      const total = parseFloat(row.f[3].v) || 0;

      if (total < 4) {
        // Not enough data to be meaningful — fall back to default
        return null;
      }

      return {
        b1Pct: b1 / total,
        b2Pct: b2 / total,
        b3Pct: b3 / total,
        totalSampled: total,
        weeksUsed: refDates.length,
      };
    }
    return null;
  } catch (err) {
    Logger.log('Batch distribution query error for ' + product.sheetName + ': ' + err.message);
    return null;
  }
}


// ============================================================
// SOBRAS DATA ACCESS (v2)
// ============================================================

/**
 * Query the Sobras API for historical leftover/sold-out data.
 *
 * @param {string} store - Store name (Milà, Sant Pau, Sant Joan)
 * @param {Date} targetDate - The date we're forecasting for
 * @param {number} dayOfWeek - JS day of week (0=Sun, 1=Mon, ..., 6=Sat)
 * @returns {object} Sobras history grouped by product, filtered to same day of week
 */
function querySobrasHistory(store, targetDate, dayOfWeek) {
  if (!CONFIG.SOBRAS_API_URL) {
    Logger.log('Sobras API URL not configured — skipping sobras integration.');
    return null;
  }

  const storeId = CONFIG.SOBRAS_STORE_IDS[store];
  if (!storeId) {
    Logger.log('Unknown store for sobras: ' + store);
    return null;
  }

  try {
    const url = CONFIG.SOBRAS_API_URL +
      '?action=getHistory&store=' + storeId +
      '&days=' + CONFIG.SOBRAS_HISTORY_DAYS;

    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const data = JSON.parse(response.getContentText());

    if (data.error) {
      Logger.log('Sobras API error: ' + data.error);
      return null;
    }

    // Filter history to only include the same day of week as target
    const dayNames = ['Domingo', 'Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado'];
    const targetDayName = dayNames[dayOfWeek];
    const history = data.history || {};

    // Group sobras by product across matching days
    const productSobras = {};  // { productName: [{ date, sobras, b1, b2, b3 }] }

    for (const [dateKey, dayData] of Object.entries(history)) {
      // Check if this date is the same day of week
      const d = new Date(dateKey + 'T00:00:00');
      if (d.getDay() !== dayOfWeek) continue;

      const products = dayData.products || {};
      for (const [productName, pData] of Object.entries(products)) {
        if (!productSobras[productName]) {
          productSobras[productName] = [];
        }
        productSobras[productName].push({
          date: dateKey,
          sobras: pData.sobras,
          b1SoldOut: pData.b1SoldOut || null,
          b2SoldOut: pData.b2SoldOut || null,
          b3SoldOut: pData.b3SoldOut || null,
        });
      }
    }

    // Count how many weeks of data we have for the target day
    const uniqueDates = new Set();
    for (const [dateKey, dayData] of Object.entries(history)) {
      const d = new Date(dateKey + 'T00:00:00');
      if (d.getDay() === dayOfWeek) uniqueDates.add(dateKey);
    }

    return {
      productSobras,
      weeksAvailable: uniqueDates.size,
      storeId,
    };

  } catch (err) {
    Logger.log('Sobras query error: ' + err.message);
    return null;
  }
}

/**
 * Step 6: Calculate the sold-out demand correction factor for a product.
 *
 * If batches consistently sell out early, actual demand > recorded sales.
 * We look at the last N weeks of sold-out times for the same weekday
 * and apply an uplift factor.
 *
 * @param {object} product - Product definition from PRODUCT_MAP
 * @param {Array} sobrasEntries - Array of { b1SoldOut, b2SoldOut, b3SoldOut } for this product
 * @returns {number} Multiplier (>= 1.0, e.g. 1.15 means +15%)
 */
function calcSoldOutCorrection(product, sobrasEntries) {
  if (!sobrasEntries || sobrasEntries.length === 0) return 1.0;

  // Only apply to pastry products (bread doesn't have batch tracking)
  if (product.type !== 'pastry') return 1.0;

  let totalFactor = 0;
  let countWithSoldOuts = 0;

  for (const entry of sobrasEntries) {
    let maxFactor = 1.0;

    // Check each batch sold-out time and pick the highest factor
    // (if B1 sold out at 10:30, that's worse than B3 at 15:30)
    if (entry.b1SoldOut) {
      const hour = parseInt(entry.b1SoldOut.split(':')[0]);
      if (hour < 11) maxFactor = Math.max(maxFactor, CONFIG.SOLDOUT_FACTORS.earlyMorning);
      else maxFactor = Math.max(maxFactor, CONFIG.SOLDOUT_FACTORS.midday);
    }
    if (entry.b2SoldOut) {
      const hour = parseInt(entry.b2SoldOut.split(':')[0]);
      if (hour < 13) maxFactor = Math.max(maxFactor, CONFIG.SOLDOUT_FACTORS.midday);
      else maxFactor = Math.max(maxFactor, CONFIG.SOLDOUT_FACTORS.afternoon);
    }
    if (entry.b3SoldOut) {
      const hour = parseInt(entry.b3SoldOut.split(':')[0]);
      if (hour < 15) maxFactor = Math.max(maxFactor, CONFIG.SOLDOUT_FACTORS.afternoon);
      else maxFactor = Math.max(maxFactor, CONFIG.SOLDOUT_FACTORS.lateAfternoon);
    }

    if (maxFactor > 1.0) {
      totalFactor += maxFactor;
      countWithSoldOuts++;
    }
  }

  if (countWithSoldOuts === 0) return 1.0;

  // Average factor across weeks that had sold-outs, weighted by frequency
  // If 3 of 4 weeks had sold-outs, apply more correction than if only 1 of 4
  const avgFactor = totalFactor / countWithSoldOuts;
  const frequency = countWithSoldOuts / sobrasEntries.length;

  // Blend: full correction if sold-out every week, partial if occasional
  // Factor = 1 + (avgFactor - 1) * frequency
  const blendedFactor = 1.0 + (avgFactor - 1.0) * frequency;

  return Math.round(blendedFactor * 100) / 100;
}

/**
 * Step 7: Calculate the overproduction dampening factor for a product.
 *
 * If a product consistently has high sobras (leftovers), we reduce production.
 * Uses the ratio: avg_sobras / base_estimate
 *
 * @param {object} product - Product definition from PRODUCT_MAP
 * @param {Array} sobrasEntries - Array of { sobras } for this product
 * @param {number} baseEstimate - The estimate after steps 1-6 (before dampening)
 * @returns {number} Multiplier (<= 1.0, e.g. 0.90 means -10%)
 */
function calcOverproductionDampening(product, sobrasEntries, baseEstimate) {
  if (!sobrasEntries || sobrasEntries.length === 0) return 1.0;
  if (baseEstimate <= 0) return 1.0;

  // Collect valid sobras counts (ignore entries where sobras was not recorded)
  const validSobras = sobrasEntries
    .filter(e => e.sobras !== null && e.sobras !== undefined && e.sobras !== '')
    .map(e => parseFloat(e.sobras) || 0);

  if (validSobras.length === 0) return 1.0;

  const avgSobras = validSobras.reduce((a, b) => a + b, 0) / validSobras.length;
  const sobrasRatio = avgSobras / baseEstimate;

  if (sobrasRatio > CONFIG.SOBRAS_THRESHOLDS.high) {
    return CONFIG.SOBRAS_REDUCTIONS.high;   // -15%
  } else if (sobrasRatio > CONFIG.SOBRAS_THRESHOLDS.medium) {
    return CONFIG.SOBRAS_REDUCTIONS.medium;  // -10%
  } else if (sobrasRatio > CONFIG.SOBRAS_THRESHOLDS.low) {
    return CONFIG.SOBRAS_REDUCTIONS.low;     // -5%
  }

  return 1.0;
}


// ============================================================
// WEATHER
// ============================================================

function getWeatherForecast(dateStr) {
  try {
    const url = `${CONFIG.WEATHER_URL}?latitude=${CONFIG.BARCELONA_LAT}&longitude=${CONFIG.BARCELONA_LON}&daily=precipitation_sum,temperature_2m_max,temperature_2m_min,weathercode&timezone=Europe/Madrid&start_date=${dateStr}&end_date=${dateStr}`;
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const data = JSON.parse(response.getContentText());

    if (data.daily && data.daily.time && data.daily.time.length > 0) {
      return {
        date: data.daily.time[0],
        precipitation_mm: data.daily.precipitation_sum[0] || 0,
        temp_max: data.daily.temperature_2m_max[0],
        temp_min: data.daily.temperature_2m_min[0],
        weathercode: data.daily.weathercode[0],
        description: weatherCodeToText(data.daily.weathercode[0]),
      };
    }
  } catch (err) {
    Logger.log('Weather fetch error: ' + err.message);
  }
  return { precipitation_mm: 0, description: 'Unknown' };
}

function getWeatherForDate(dateStr) {
  // For dates within the next 7 days, use the forecast API
  // For past dates, weather won't affect the estimate anyway
  const today = new Date();
  const target = new Date(dateStr + 'T00:00:00');
  const diffDays = Math.floor((target - today) / (1000 * 60 * 60 * 24));

  if (diffDays >= 0 && diffDays <= 14) {
    return getWeatherForecast(dateStr);
  }
  return { precipitation_mm: 0, description: 'No forecast available' };
}

function weatherCodeToText(code) {
  const codes = {
    0: 'Clear sky', 1: 'Mainly clear', 2: 'Partly cloudy', 3: 'Overcast',
    45: 'Fog', 48: 'Rime fog',
    51: 'Light drizzle', 53: 'Moderate drizzle', 55: 'Dense drizzle',
    61: 'Slight rain', 63: 'Moderate rain', 65: 'Heavy rain',
    71: 'Slight snow', 73: 'Moderate snow', 75: 'Heavy snow',
    80: 'Slight showers', 81: 'Moderate showers', 82: 'Violent showers',
    95: 'Thunderstorm', 96: 'Thunderstorm with hail', 99: 'Thunderstorm with heavy hail',
  };
  return codes[code] || 'Unknown (' + code + ')';
}


// ============================================================
// HOLIDAY CHECK
// ============================================================

function checkHoliday(dateStr) {
  const monthDay = dateStr.substring(5); // "MM-DD"
  return CONFIG.HOLIDAYS_2026.includes(monthDay);
}


// ============================================================
// SYNC TO PRODUCTION SHEETS
// ============================================================

/**
 * Write estimates into the correct production spreadsheet.
 *
 * @param {string} store - Store name (Milà, Sant Pau, Sant Joan)
 * @param {string} dayOfWeek - Day name in Spanish (Lunes, Martes, etc.)
 * @param {Array} estimates - Array of { cell, value } objects
 * @returns {object} Result with status
 */
function syncToSheet(store, dayOfWeek, estimates) {
  if (!store || !dayOfWeek || !estimates) {
    throw new Error('Missing required fields: store, dayOfWeek, estimates');
  }

  const sheetId = CONFIG.SPREADSHEETS[store];
  if (!sheetId) throw new Error('Unknown store: ' + store);

  if (!CONFIG.DAY_TABS.includes(dayOfWeek)) {
    throw new Error('Invalid day: ' + dayOfWeek + '. Must be one of: ' + CONFIG.DAY_TABS.join(', '));
  }

  const spreadsheet = SpreadsheetApp.openById(sheetId);
  const sheet = spreadsheet.getSheetByName(dayOfWeek);
  if (!sheet) throw new Error('Tab "' + dayOfWeek + '" not found in ' + store + ' spreadsheet');

  const updates = [];
  for (const est of estimates) {
    if (est.cell && est.value !== undefined && est.value !== null) {
      const cell = sheet.getRange(est.cell);
      cell.setValue(est.value);
      updates.push({ cell: est.cell, value: est.value });
    }
  }

  SpreadsheetApp.flush(); // Ensure all writes are committed

  return {
    success: true,
    store,
    dayOfWeek,
    updatedCells: updates.length,
    updates,
    syncedAt: new Date().toISOString(),
  };
}


// ============================================================
// SCHEDULED TRIGGER (call from time-driven trigger)
// ============================================================

/**
 * Generate estimates for yesterday's weekday NEXT WEEK.
 * Call this from a time-driven trigger set to 6:00 AM Europe/Madrid.
 *
 * Logic: At 6am each day, yesterday's sobras data is now available.
 * We use that to generate next week's estimate for the same weekday.
 *
 * Examples (triggered at 6am):
 * - Tuesday  6am → Monday's sobras in    → generate next Monday    (today + 6 days)
 * - Wednesday 6am → Tuesday's sobras in  → generate next Tuesday   (today + 6 days)
 * - Thursday 6am → Wednesday's sobras in → generate next Wednesday (today + 6 days)
 * - Friday   6am → Thursday's sobras in  → generate next Thursday  (today + 6 days)
 * - Saturday 6am → Friday's sobras in    → generate next Friday    (today + 6 days)
 * - Sunday   6am → Saturday's sobras in  → generate next Saturday  (today + 6 days)
 * - Monday   6am → Sunday's sobras in    → skip (bakery closed Sunday) OR generate next Sunday
 *
 * The pastry team works weekly, so estimates build up progressively Mon→Sat
 * and the full week is ready by Saturday morning.
 */
function generateDailyEstimates() {
  const today = new Date();
  const dayOfWeek = today.getDay(); // 0=Sun, 1=Mon, ..., 6=Sat

  // Yesterday's day of week (the day whose sobras just came in)
  // On Monday (1), yesterday was Sunday (0) — bakery may be closed, skip
  if (dayOfWeek === 1) {
    Logger.log('Monday 6am: Sunday sobras — skipping (bakery closed Sunday).');
    return { skipped: true, reason: 'Sunday - bakery closed' };
  }

  // Target date = yesterday's weekday, but next week
  // That's always today + 6 days (yesterday + 7 days = today + 6 days)
  const targetDate = new Date(today);
  targetDate.setDate(today.getDate() + 6);

  const dateStr = formatDate(targetDate);
  const dayNames = ['Domingo', 'Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado'];
  const targetDayName = dayNames[targetDate.getDay()];

  Logger.log(`Generating estimates for ${targetDayName} ${dateStr} (yesterday's weekday, next week)`);

  const stores = Object.keys(CONFIG.SPREADSHEETS);
  const allEstimates = {};

  for (const store of stores) {
    try {
      allEstimates[store] = getEstimates(store, dateStr);
      Logger.log('Generated estimates for ' + store + ' on ' + dateStr);
    } catch (err) {
      Logger.log('Error generating estimates for ' + store + ': ' + err.message);
      allEstimates[store] = { error: err.message };
    }
  }

  // Store in Script Properties — keyed by target date so web app can retrieve any day
  const props = PropertiesService.getScriptProperties();

  // Save this day's estimate
  props.setProperty('estimates_' + dateStr, JSON.stringify({
    date: dateStr,
    dayOfWeek: targetDayName,
    generatedAt: new Date().toISOString(),
    estimates: allEstimates,
  }));

  // Also update latest pointer
  props.setProperty('latestEstimates', JSON.stringify({
    date: dateStr,
    dayOfWeek: targetDayName,
    generatedAt: new Date().toISOString(),
    estimates: allEstimates,
  }));

  // Save forecast to BigQuery for historical tracking
  try {
    saveForecastToBigQuery(dateStr, targetDayName, allEstimates, new Date().toISOString());
  } catch (e) {
    Logger.log('BigQuery save error (non-fatal): ' + e.message);
  }

  // Send email notification to François
  try {
    const firstStore = stores[0];
    const est = allEstimates[firstStore];
    if (est && !est.error) {
      const nonZero = est.estimates ? est.estimates.filter(e => e.estimate > 0).length : 0;
      const subject = `🥐 Origo - Estimaciones listas: ${targetDayName} ${dateStr}`;
      const body = `Estimaciones de producción generadas para ${targetDayName} ${dateStr} (semana que viene).\n\n` +
        `${nonZero} productos con estimaciones > 0.\n` +
        `Basadas en los datos de sobras de ayer.\n\n` +
        `Revisa y sincroniza en la app de estimaciones.\n\n` +
        `Tiempo previsto: ${est.weather ? est.weather.description : 'N/A'}\n` +
        `Festivo: ${est.isHoliday ? 'Sí' : 'No'}`;
      MailApp.sendEmail('francois@origobakery.com', subject, body);
    }
  } catch (e) {
    Logger.log('Email notification error: ' + e.message);
  }

  return allEstimates;
}

/**
 * Get the latest generated estimates (from the scheduled trigger)
 */
function getLatestEstimates() {
  const props = PropertiesService.getScriptProperties();
  const data = props.getProperty('latestEstimates');
  return data ? JSON.parse(data) : null;
}

/**
 * Get the list of dates for which the trigger has generated estimates.
 * Used by the frontend to know which days to show as "next week" vs "this week".
 */
function getGeneratedDates() {
  const props = PropertiesService.getScriptProperties();
  const all = props.getProperties();
  const dates = [];
  for (const key of Object.keys(all)) {
    if (key.startsWith('estimates_')) {
      const dateStr = key.replace('estimates_', '');
      try {
        const data = JSON.parse(all[key]);
        dates.push({ date: dateStr, generatedAt: data.generatedAt, dayOfWeek: data.dayOfWeek });
      } catch(e) {}
    }
  }
  return { dates };
}


// ============================================================
// FORECAST HISTORY — BigQuery Storage
// ============================================================

const BQ_FORECAST_TABLE = 'Forecast_History';

/**
 * One-time setup: create the Forecast_History table.
 * Run this manually once from the Apps Script editor.
 */
function setupForecastTable() {
  const tableRef = {
    projectId: CONFIG.BQ_PROJECT,
    datasetId: CONFIG.BQ_DATASET,
    tableId: BQ_FORECAST_TABLE,
  };

  const schema = {
    fields: [
      { name: 'forecast_id',              type: 'STRING',    mode: 'REQUIRED' },
      { name: 'generated_at',             type: 'TIMESTAMP', mode: 'REQUIRED' },
      { name: 'target_date',              type: 'DATE',      mode: 'REQUIRED' },
      { name: 'day_of_week',              type: 'STRING',    mode: 'REQUIRED' },
      { name: 'store',                    type: 'STRING',    mode: 'REQUIRED' },
      { name: 'product',                  type: 'STRING',    mode: 'REQUIRED' },
      { name: 'cell',                     type: 'STRING' },
      { name: 'product_type',             type: 'STRING' },
      { name: 'estimate',                 type: 'INT64' },
      { name: 'raw_estimate',             type: 'FLOAT64' },
      { name: 'week1_sales',              type: 'FLOAT64' },
      { name: 'week2_sales',              type: 'FLOAT64' },
      { name: 'week3_sales',              type: 'FLOAT64' },
      { name: 'week4_sales',              type: 'FLOAT64' },
      { name: 'factor_trend',             type: 'FLOAT64' },
      { name: 'factor_seasonal',          type: 'FLOAT64' },
      { name: 'factor_holiday',           type: 'FLOAT64' },
      { name: 'factor_weather',           type: 'FLOAT64' },
      { name: 'factor_sold_out',          type: 'FLOAT64' },
      { name: 'factor_overproduction',    type: 'FLOAT64' },
      { name: 'sobras_weeks_available',   type: 'INT64' },
      { name: 'sobras_avg',               type: 'FLOAT64' },
      { name: 'weather_description',      type: 'STRING' },
      { name: 'weather_precipitation_mm', type: 'FLOAT64' },
      { name: 'is_holiday',               type: 'BOOL' },
      { name: 'rafa_approved_value',      type: 'INT64' },
      { name: 'rafa_synced_at',           type: 'TIMESTAMP' },
    ]
  };

  const table = {
    tableReference: tableRef,
    schema: schema,
    timePartitioning: { type: 'DAY', field: 'target_date' },
    clustering: { fields: ['store', 'product'] },
    description: 'Production forecast history — AI suggestions + Rafa-approved final values',
  };

  try {
    BigQuery.Tables.insert(table, CONFIG.BQ_PROJECT, CONFIG.BQ_DATASET);
    Logger.log('✓ Forecast_History table created successfully.');
    return { success: true, message: 'Table created' };
  } catch (e) {
    if (e.message && e.message.includes('Already Exists')) {
      Logger.log('Table already exists — no action needed.');
      return { success: true, message: 'Table already exists' };
    }
    Logger.log('Error creating table: ' + e.message);
    return { error: e.message };
  }
}

/**
 * Save generated forecast estimates to BigQuery.
 * Called automatically by generateDailyEstimates().
 */
function saveForecastToBigQuery(dateStr, dayOfWeek, allEstimates, generatedAt) {
  const rows = [];

  for (const [store, storeData] of Object.entries(allEstimates)) {
    if (storeData.error || !storeData.estimates) continue;

    for (const est of storeData.estimates) {
      rows.push({
        json: {
          forecast_id: `${dateStr}_${store}_${est.product}`,
          generated_at: generatedAt,
          target_date: dateStr,
          day_of_week: dayOfWeek,
          store: store,
          product: est.product,
          cell: est.cell,
          product_type: est.type,
          estimate: est.estimate,
          raw_estimate: est.rawEstimate,
          week1_sales: est.weekSales[0],
          week2_sales: est.weekSales[1],
          week3_sales: est.weekSales[2],
          week4_sales: est.weekSales[3],
          factor_trend: est.factors.trend,
          factor_seasonal: est.factors.seasonal,
          factor_holiday: est.factors.holiday,
          factor_weather: est.factors.weather,
          factor_sold_out: est.factors.soldOut,
          factor_overproduction: est.factors.overproduction,
          sobras_weeks_available: storeData.sobrasWeeksAvailable || 0,
          sobras_avg: est.sobrasData ? est.sobrasData.avgSobras : null,
          weather_description: storeData.weather ? storeData.weather.description : null,
          weather_precipitation_mm: storeData.weather ? storeData.weather.precipitation_mm : null,
          is_holiday: storeData.isHoliday || false,
          rafa_approved_value: null,
          rafa_synced_at: null,
        }
      });
    }
  }

  if (rows.length === 0) {
    Logger.log('No forecast rows to save.');
    return;
  }

  try {
    // Delete existing rows for this date (in case of re-generation)
    const deleteSql = `DELETE FROM \`${CONFIG.BQ_PROJECT}.${CONFIG.BQ_DATASET}.${BQ_FORECAST_TABLE}\` WHERE target_date = '${dateStr}'`;
    BigQuery.Jobs.query({ query: deleteSql, useLegacySql: false }, CONFIG.BQ_PROJECT);

    // Insert new rows
    const insertRequest = { rows: rows };
    BigQuery.Tabledata.insertAll(insertRequest, CONFIG.BQ_PROJECT, CONFIG.BQ_DATASET, BQ_FORECAST_TABLE);
    Logger.log('Saved ' + rows.length + ' forecast rows to BigQuery for ' + dateStr);
  } catch (e) {
    Logger.log('BigQuery forecast save error: ' + e.message);
  }
}

/**
 * Record Rafa's approved values after she syncs.
 * Called via POST action 'recordApproval'.
 */
function recordRafaApproval(store, dayOfWeek, approvedEstimates) {
  if (!store || !dayOfWeek || !approvedEstimates) {
    throw new Error('Missing required fields: store, dayOfWeek, approvedEstimates');
  }

  const syncedAt = new Date().toISOString();

  // Find the target date for this day (next occurrence of dayOfWeek)
  const dayIndex = CONFIG.DAY_TABS.indexOf(dayOfWeek);
  if (dayIndex === -1) throw new Error('Invalid day: ' + dayOfWeek);

  // Build UPDATE statements for each approved product
  let updatedCount = 0;
  for (const est of approvedEstimates) {
    if (!est.cell || est.value === undefined) continue;

    // Find product name from cell
    const product = PRODUCT_MAP_COMMON.find(p => p.cell === est.cell);
    if (!product) continue;

    try {
      const sql = `UPDATE \`${CONFIG.BQ_PROJECT}.${CONFIG.BQ_DATASET}.${BQ_FORECAST_TABLE}\`
        SET rafa_approved_value = ${est.value}, rafa_synced_at = TIMESTAMP('${syncedAt}')
        WHERE store = '${store}' AND product = '${product.sheetName}' AND day_of_week = '${dayOfWeek}'
          AND rafa_synced_at IS NULL
        ORDER BY generated_at DESC LIMIT 1`;
      BigQuery.Jobs.query({ query: sql, useLegacySql: false }, CONFIG.BQ_PROJECT);
      updatedCount++;
    } catch (e) {
      Logger.log('Approval update error for ' + product.sheetName + ': ' + e.message);
    }
  }

  return { success: true, updatedProducts: updatedCount, syncedAt };
}


// ============================================================
// UTILITY FUNCTIONS
// ============================================================

function formatDate(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}

function getNextMonday() {
  const today = new Date();
  const day = today.getDay(); // 0=Sun
  const diff = day === 0 ? 1 : (8 - day); // Days until next Monday
  const nextMon = new Date(today);
  nextMon.setDate(today.getDate() + diff);
  return nextMon;
}


// ============================================================
// TEST / DEBUG
// ============================================================

/**
 * Test function — run manually to verify the system works
 */
function testEstimates() {
  // Test for next Monday
  const nextMon = getNextMonday();
  const dateStr = formatDate(nextMon);

  Logger.log('Testing estimates for: ' + dateStr);

  // Test weather
  const weather = getWeatherForecast(dateStr);
  Logger.log('Weather: ' + JSON.stringify(weather));

  // Test holiday check
  Logger.log('Is holiday: ' + checkHoliday(dateStr));

  // Test estimates for each store
  for (const store of Object.keys(CONFIG.SPREADSHEETS)) {
    try {
      const est = getEstimates(store, dateStr);
      Logger.log('\n=== ' + store + ' ===');
      for (const e of est.estimates) {
        Logger.log(e.product + ' (' + e.cell + '): ' + e.estimate +
          ' [raw: ' + e.rawEstimate + ', trend: ' + e.factors.trend +
          ', seasonal: ' + e.factors.seasonal + ']');
      }
    } catch (err) {
      Logger.log('Error for ' + store + ': ' + err.message);
    }
  }
}

/**
 * Test sync - writes test values to a specific day tab
 * WARNING: This will overwrite real data! Use with caution.
 */
function testSync() {
  const testEstimates = [
    { cell: 'D8', value: 50 },  // ORIGO
    { cell: 'D29', value: 72 }, // Canela Twist
  ];

  // const result = syncToSheet('Sant Pau', 'Lunes', testEstimates);
  // Logger.log('Sync result: ' + JSON.stringify(result));
  Logger.log('Test sync is commented out for safety. Uncomment to test.');
}
