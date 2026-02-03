/***********************
 * TRISOR LIVE SYNC - PRODUCTION READY
 * LiveData Dashboard + Export Control Center
 * Keine Duplikate | Vollständig | Skalierbar
 ***********************/

const CONFIG = {
  SHEETS: {
    LIVEDATA: 'LiveData',
    ZIELE: 'Ziele',
    EMAIL_RECIPIENTS: 'Emailreport'
  },
  THRESHOLDS: {
    green: 0.85,
    yellow: 0.75
  },
  BRAND: {
    purple: '#6E2DA0',
    purpleDark: '#5A2480',
    gold: '#B89D73',
    goldLight: '#F9F7F4',
    goldAccent: '#D4C4A8',
    white: '#ffffff',
    stripeGray: '#FBFBFB',
    border: '#E0E0E0',
    textDark: '#222222',
    textLight: '#777777',
    bg: '#F2F2F2',
    success: '#10B981',
    warning: '#F59E0B',
    danger: '#E05252'
  },
  CITY_ORDER: [
    'Berlin',
    'München',
    'Hamburg',
    'Köln',
    'Düsseldorf',
    'Nürnberg',
    'Stuttgart'
  ],
  ICON_URLS: {
    check: "https://img.icons8.com/ios-filled/100/6E2DA0/ok--v1.png",
    chart: "https://img.icons8.com/?size=100&id=wmLdIN2s5OPt&format=png&color=6E2DA0",
    target: "https://img.icons8.com/ios-filled/100/6E2DA0/goal--v1.png",
    table: "https://img.icons8.com/ios-filled/50/6E2DA0/table-1.png"
  }
};


/***********************
 * CORE DATA ENGINE
 ***********************/

/**
 * VOLLSTÄNDIGES BACKEND: Inklusive Wochen-Details (Anzahl + Name)
 * + Monatsdetails für Team-Analytics (Tarif, Größe, Promo pro Monat)
 * + Matrix: planBySize (Tarif x Größe)
 */
function getHeatmapForPeriod(period) {
  const frontend = loadDataForFrontend();
  const heatmap = frontend.heatmap || {};
  const periods = frontend.periods || [];

  // Wenn period leer/ungültig → fallback: latest
  const selected = (period && periods.includes(period)) ? period : (periods[0] || null);
  if (!selected) return { period: null, cities: {}, totals: { S:0, M:0, L:0, U:0 } };

  // Ergebnis: pro Stadt immer S/M/L/U vorhanden
  const cities = {};
  const totals = { S:0, M:0, L:0, U:0 };

  Object.keys(heatmap).forEach(city => {
    const m = heatmap[city]?.[selected] || {};
    const obj = {
      S: Number(m.S || 0),
      M: Number(m.M || 0),
      L: Number(m.L || 0),
      U: Number(m.U || 0),
    };
    cities[city] = obj;
    totals.S += obj.S; totals.M += obj.M; totals.L += obj.L; totals.U += obj.U;
  });

  return { period: selected, cities, totals, availablePeriods: periods };
}
function getLiveData() {
  try {
    const t0 = new Date().getTime(); // Performance Messung Start
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const liveSheet = ss.getSheetByName(CONFIG.SHEETS.LIVEDATA);
    const blacklistSheet = ss.getSheetByName('Blacklist');

    const empty = () => ({
      data: {}, heatmap: {}, tariffHeatmap: {}, planBySizeHeatmap: {},
      activators: {}, periods: [], totalEvents: 0,
      lastUpdate: new Date().toISOString()
    });

    if (!liveSheet || liveSheet.getLastRow() < 2) return empty();

    // 1) Blacklist einlesen
    let blacklist = [];
    if (blacklistSheet && blacklistSheet.getLastRow() >= 2) {
      blacklist = blacklistSheet
        .getRange(2, 1, blacklistSheet.getLastRow() - 1, 1)
        .getValues()
        .map(r => String(r[0]).trim().toLowerCase())
        .filter(e => e !== "");
    }

    // 🚀 PERFORMANCE-BOOST: Settings EINMAL laden, statt 14.000 mal
    const settingsMap = getCitySettingsMap();

    const data = {};
    const heatmap = {};
    const tariffHeatmap = {};
    const planBySizeHeatmap = {};
    const activators = {};
    const periods = new Set();
    const WEEKDAY_NAMES = ['Sonntag', 'Montag', 'Dienstag', 'Mittwoch', 'Donnerstag', 'Freitag', 'Samstag'];

    // Daten in einem Rutsch holen (sehr schnell)
    const rows = liveSheet.getRange(2, 1, liveSheet.getLastRow() - 1, 28).getValues();

    rows.forEach(row => {
      const location_name = row[1];
      const activated_date = row[25];
      const activatorEmail = String(row[27] || '').trim().toLowerCase();

      if (!location_name || !activated_date || !activatorEmail) return;
      if (blacklist.indexOf(activatorEmail) > -1) return;

      // 🚀 HIER IST DER FIX: Wir übergeben die 'settingsMap'
      // Kein Cache-Zugriff mehr innerhalb der Schleife!
      const standort = normalizeCityNameFromSettings(location_name, false, settingsMap);
      
      const monat = normalizeDateToMonth(activated_date);
      if (!monat) return;

      const boxSize = normalizeBoxSize(String(row[6] || '').trim());
      const plan = String(row[7] || '').trim() || 'Unbekannt';
      const promo = String(row[14] || '').trim();

      periods.add(monat);

      // Daten aggregieren (Standard Logik)
      data[standort] = data[standort] || {};
      data[standort][monat] = (data[standort][monat] || 0) + 1;

      // Heatmap
      heatmap[standort] = heatmap[standort] || {};
      heatmap[standort][monat] = heatmap[standort][monat] || { S: 0, M: 0, L: 0, U: 0 };
      heatmap[standort][monat][boxSize] = (heatmap[standort][monat][boxSize] || 0) + 1;

      // Tarif Heatmap
      tariffHeatmap[standort] = tariffHeatmap[standort] || {};
      tariffHeatmap[standort][monat] = tariffHeatmap[standort][monat] || {};
      tariffHeatmap[standort][monat][plan] = (tariffHeatmap[standort][monat][plan] || 0) + 1;

      // Plan by Size
      planBySizeHeatmap[standort] = planBySizeHeatmap[standort] || {};
      planBySizeHeatmap[standort][monat] = planBySizeHeatmap[standort][monat] || {};
      planBySizeHeatmap[standort][monat][plan] = planBySizeHeatmap[standort][monat][plan] || { S: 0, M: 0, L: 0, U: 0 };
      planBySizeHeatmap[standort][monat][plan][boxSize] = (planBySizeHeatmap[standort][monat][plan][boxSize] || 0) + 1;

      // Activator Logik
      if (!activators[activatorEmail]) {
        activators[activatorEmail] = {
          total: 0, cities: {}, boxSizes: { S: 0, M: 0, L: 0, U: 0 },
          pricePlans: {}, promoCodes: {}, weekdays: {}, monthlyData: {},
          monthlyDetails: {}, weeklyData: {}, totalRevenue: 0, bestWeekCount: 0, bestWeekLabel: '-', avgRevenuePerYear: 0
        };
      }

      const emp = activators[activatorEmail];
      emp.total++;
      emp.monthlyData[monat] = (emp.monthlyData[monat] || 0) + 1;
      emp.cities[standort] = (emp.cities[standort] || 0) + 1;
      emp.boxSizes[boxSize] = (emp.boxSizes[boxSize] || 0) + 1;
      emp.pricePlans[plan] = (emp.pricePlans[plan] || 0) + 1;
      if (promo) emp.promoCodes[promo] = (emp.promoCodes[promo] || 0) + 1;

      // Monthly Details
      emp.monthlyDetails[monat] = emp.monthlyDetails[monat] || {
        boxSizes: { S: 0, M: 0, L: 0, U: 0 }, pricePlans: {}, promoCodes: {}, cities: {}, planBySize: {}
      };
      const md = emp.monthlyDetails[monat];
      md.boxSizes[boxSize] = (md.boxSizes[boxSize] || 0) + 1;
      md.pricePlans[plan] = (md.pricePlans[plan] || 0) + 1;
      if (promo) md.promoCodes[promo] = (md.promoCodes[promo] || 0) + 1;
      md.cities[standort] = (md.cities[standort] || 0) + 1;
      md.planBySize[plan] = md.planBySize[plan] || { S: 0, M: 0, L: 0, U: 0 };
      md.planBySize[plan][boxSize] = (md.planBySize[plan][boxSize] || 0) + 1;

      // Umsatz
      const basePrice = parseFloat(String(row[13] || '0').replace(',', '.')) || 0;
      const promoMonths = parseInt(row[15] || 0, 10) || 0;
      const promoPrice = parseFloat(String(row[16] || '0').replace(',', '.')) || 0;
      const freeMonths = parseInt(row[17] || 0, 10) || 0;

      let yr = (promoMonths > 0 && promoPrice > 0)
        ? (promoPrice * Math.min(promoMonths, 12)) + (basePrice * Math.max(0, 12 - promoMonths))
        : basePrice * 12;

      if (freeMonths > 0) {
        yr = Math.max(0, yr - (basePrice * Math.min(freeMonths, 12)));
      }
      emp.totalRevenue += yr;

      // Wochen-Check
      const d = new Date(activated_date);
      if (!isNaN(d.getTime())) {
        emp.weekdays[WEEKDAY_NAMES[d.getDay()]] = (emp.weekdays[WEEKDAY_NAMES[d.getDay()]] || 0) + 1;
        const weekKey = getWeekKey_(d);
        emp.weeklyData[weekKey] = (emp.weeklyData[weekKey] || 0) + 1;
      }
    });

    // Nachbearbeitung Activators
    Object.keys(activators).forEach(email => {
      const e = activators[email];
      e.avgRevenuePerYear = e.total > 0 ? (e.totalRevenue / e.total) : 0;
      let maxVal = 0;
      let bestLabel = '-';
      Object.keys(e.weeklyData || {}).forEach(key => {
        if ((e.weeklyData[key] || 0) > maxVal) {
          maxVal = e.weeklyData[key];
          const parts = key.split('-W');
          bestLabel = "KW " + parts[1] + " (" + parts[0] + ")";
        }
      });
      e.bestWeekCount = maxVal;
      e.bestWeekLabel = bestLabel;
      delete e.weeklyData;
    });

    const t1 = new Date().getTime();
    Logger.log(`⚡ getLiveData fertig in ${t1 - t0}ms für ${rows.length} Zeilen.`);

    return {
      data, heatmap, tariffHeatmap, planBySizeHeatmap, activators,
      periods: Array.from(periods).sort().reverse(),
      totalEvents: rows.length,
      lastUpdate: new Date().toISOString()
    };

  } catch (error) {
    Logger.log('❌ getLiveData error: ' + error.toString());
    return { data: {}, heatmap: {}, tariffHeatmap: {}, planBySizeHeatmap: {}, activators: {}, periods: [], totalEvents: 0, lastUpdate: new Date().toISOString() };
  }
}

/**
 * Kalenderwoche Helper
 */
function getWeekKey_(date) {
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  const weekNo = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
  return d.getUTCFullYear() + "-W" + (weekNo < 10 ? "0" + weekNo : weekNo);
}


function getGoals() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.ZIELE);
    
    if (!sheet || sheet.getLastRow() < 2) {
      Logger.log('⚠️ Ziele-Sheet leer');
      return {};
    }
    
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    const headers = data[0];
    const goals = {};
    
    for (let i = 1; i < data.length; i++) {
      const standort = String(data[i][0]).trim();
      if (!standort) continue;
      
      goals[standort] = {};
      
      for (let j = 1; j < headers.length; j++) {
        let monatKey = headers[j];
        
        if (monatKey instanceof Date) {
          const year = monatKey.getFullYear();
          const month = String(monatKey.getMonth() + 1).padStart(2, '0');
          monatKey = `${year}-${month}`;
        } else {
          monatKey = String(monatKey).trim();
        }
        
        const value = Number(data[i][j]);
        
        if (monatKey.match(/^\d{4}-\d{2}$/) && !isNaN(value) && value > 0) {
          goals[standort][monatKey] = value;
        }
      }
    }
    
    Logger.log('✅ getGoals:', Object.keys(goals).length, 'Standorte');
    return goals;
  } catch (error) {
    Logger.log('❌ getGoals error:', error.toString());
    return {};
  }
}

function setGoal(standort, month, year, ziel) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(CONFIG.SHEETS.ZIELE);
    
    const monatKey = `${year}-${String(month).padStart(2, '0')}`;
    const zielValue = Number(ziel) || 0;
    
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.SHEETS.ZIELE);
      sheet.getRange(1, 1).setValue('Standort');
      sheet.getRange(1, 1).setNumberFormat('@');
    }
    
    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
    
    let colIndex = -1;
    for (let i = 0; i < headers.length; i++) {
      if (String(headers[i]).trim() === monatKey) {
        colIndex = i + 1;
        break;
      }
    }
    
    if (colIndex === -1) {
      colIndex = headers.length + 1;
      const cell = sheet.getRange(1, colIndex);
      cell.setNumberFormat('@');
      cell.setValue(monatKey);
    }
    
    const lastRow = sheet.getLastRow();
    let rowIndex = -1;
    
    if (lastRow > 1) {
      const standorte = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < standorte.length; i++) {
        if (String(standorte[i][0]).trim() === standort) {
          rowIndex = i + 2;
          break;
        }
      }
    }
    
    if (rowIndex === -1) {
      const newRow = Array(colIndex).fill(0);
      newRow[0] = standort;
      newRow[colIndex - 1] = zielValue;
      sheet.appendRow(newRow);
    } else {
      sheet.getRange(rowIndex, colIndex).setValue(zielValue);
    }
    
    Logger.log(`✅ Ziel gespeichert: ${standort} | ${monatKey} = ${zielValue}`);
    return `Gespeichert: ${standort} ${monatKey} = ${zielValue}`;
  } catch (error) {
    Logger.log('❌ setGoal error:', error.toString());
    throw new Error('Fehler: ' + error.message);
  }
}

function setBulkGoal(standort, year, ziel) {
  for (let month = 1; month <= 12; month++) {
    setGoal(standort, month, year, ziel);
  }
  return getGoals();
}

function getPricePlanData() {
  try {
    const cache = CacheService.getScriptCache();
    const cached = cache.get('PRICE_PLAN_DATA_V1');
    if (cached) return JSON.parse(cached);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.LIVEDATA);

    if (!sheet || sheet.getLastRow() < 2) {
      const empty = { total: {}, byCity: {} };
      cache.put('PRICE_PLAN_DATA_V1', JSON.stringify(empty), 300);
      return empty;
    }

    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
    const total = {};
    const byCity = {};

    rows.forEach(row => {
      const location_name = String(row[1]).trim();
      const price_plan = String(row[7]).trim();

      if (!location_name || !price_plan) return;

      total[price_plan] = (total[price_plan] || 0) + 1;
      byCity[location_name] = byCity[location_name] || {};
      byCity[location_name][price_plan] = (byCity[location_name][price_plan] || 0) + 1;
    });

    const result = { total, byCity };
    cache.put('PRICE_PLAN_DATA_V1', JSON.stringify(result), 300); // 5 Minuten
    return result;

  } catch (error) {
    Logger.log('❌ getPricePlanData error: ' + error);
    return { total: {}, byCity: {} };
  }
}

/***********************
 * HELPER FUNCTIONS
 ***********************/

function normalizeDateToMonth(value) {
  try {
    if (!value) return null;

    // ✅ FALL 1: echtes Date-Objekt (DER WICHTIGSTE FIX)
    if (value instanceof Date && !isNaN(value.getTime())) {
      const y = value.getFullYear();
      const m = String(value.getMonth() + 1).padStart(2, '0');
      return `${y}-${m}`;
    }

    // ab hier: String-Fälle
    const str = String(value).trim();

    // yyyy-mm (bereits normalisiert)
    if (/^\d{4}-\d{2}$/.test(str)) return str;

    let date;

    // dd.MM.yyyy
    const m = str.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
    if (m) {
      date = new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
    } else {
      // ISO / sonstige Formate
      date = new Date(str);
    }

    if (isNaN(date.getTime())) return null;

    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    return `${year}-${month}`;

  } catch (e) {
    return null;
  }
}

function normalizeBoxSize(boxStr) {
  const str = String(boxStr).toUpperCase().trim();
  if (str.includes('S') || str.includes('KLEIN')) return 'S';
  if (str.includes('M') || str.includes('MITTEL')) return 'M';
if (str.includes('L') || str.includes('GROSS') || str.includes('GROß') || str.includes('GROẞ')) return 'L';
  return 'U';
}

function extractNameFromEmail(email) {
  if (!email || !email.includes('@')) return 'Unbekannt';
  const localPart = email.split('@')[0];
  const parts = localPart.split('.');
  if (parts.length >= 2) {
    const firstName = parts[0].charAt(0).toUpperCase() + parts[0].slice(1).toLowerCase();
    const lastName = parts[1].charAt(0).toUpperCase() + parts[1].slice(1).toLowerCase();
    return `${firstName} ${lastName}`;
  }
  return localPart.charAt(0).toUpperCase() + localPart.slice(1).toLowerCase();
}


/***********************
 * KI ANALYSEN
 ***********************/

/**
 * GENERIERT DIE KI-DIAGNOSE IM ELITE-LOOK
 * Executive Summary oben, Details einklappbar.
 */

function getAiDashboardSummary() {
  try {
    const frontend = loadDataForFrontend();
    const { data, goals, forecasts, periods } = frontend;

    if (!periods?.length) return { summaryText: '<div class="ai-main-container">Keine Daten vorhanden.</div>' };

    const latestPeriod = periods[0];
    const [year, month] = latestPeriod.split('-').map(Number);
    const monthName = new Intl.DateTimeFormat('de-DE', { month: 'long' }).format(new Date(year, month - 1));

    let totalIst = 0, totalZiel = 0, totalForecast = 0;
    Object.keys(data).forEach(city => {
      const val = Array.isArray(data[city][latestPeriod]) ? data[city][latestPeriod].reduce((a, b) => a + (b || 0), 0) : (data[city][latestPeriod] || 0);
      totalIst += val;
      totalZiel += (goals[city] && goals[city][latestPeriod]) || 0;
      if (forecasts[city]) totalForecast += forecasts[city].forecast;
    });

    const istRatio = totalZiel > 0 ? Math.round((totalIst / totalZiel) * 100) : 0;
    const forecastRatio = totalZiel > 0 ? Math.round((totalForecast / totalZiel) * 100) : 0;
    const statusColor = forecastRatio >= 95 ? '#00E6A7' : forecastRatio >= 75 ? '#FFB938' : '#FF5C5C';

    const insights = _generateLocationInsights(data, goals, latestPeriod, 
      { istMonth: totalIst, zielMonth: totalZiel, ratioMonth: istRatio, ratioYtd: 68 }, // Beispielwert 68%
      { ratio: forecastRatio }, true
    );

    return {
      summaryText: `
      <div class="ai-main-container" style="padding: 10px;">
        <div class="ai-header" style="display:flex; justify-content:space-between; align-items:center; margin-bottom:15px;">
          <div style="display:flex; align-items:center; gap:8px;">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#B89D73" stroke-width="2"><path d="M21 16V8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16z"></path><polyline points="3.27 6.96 12 12.01 20.73 6.96"></polyline></svg>
            <span style="font-weight:800; font-size:10px; letter-spacing:0.1em; color:#B89D73;">EXECUTIVE SUMMARY</span>
          </div>
          <button onclick="toggleKiDetails()" style="background:rgba(255,255,255,0.05); border:1px solid rgba(255,255,255,0.1); color:#717786; padding:5px 12px; border-radius:6px; font-size:10px; cursor:pointer; font-weight:700;">
            <span id="kiToggleText">DETAILS ANZEIGEN</span>
          </button>
        </div>

        <div class="ai-exec-summary" style="font-size:13px; line-height:1.7; color:rgba(255,255,255,0.9); margin-bottom: 10px; padding: 0 5px;">
          ${insights.summaryText}
        </div>

        <div id="kiDetailsContainer" style="display:none; margin-top:20px; padding-top:15px; border-top:1px solid rgba(255,255,255,0.1);">
          <div style="display:flex; justify-content:space-between; align-items: flex-start; margin-bottom: 20px;">
            <div>
              <div style="font-size:10px; color:#717786; font-weight:700; text-transform:uppercase;">STAND ${monthName.toUpperCase()}</div>
              <div style="display:flex; align-items:baseline; gap:6px; margin-top:5px;">
                <span style="font-size:24px; font-weight:800; color:#FFF;">${totalIst}</span>
                <span style="font-size:14px; color:#717786;">/ ${totalZiel} Ziel</span>
              </div>
              <div style="font-size:11px; color:var(--color-purple-light); font-weight:600; margin-top:5px;">Aktueller Fortschritt: ${istRatio}%</div>
            </div>
            <div style="text-align:right;">
              <div style="font-size:10px; color:#717786; font-weight:700; text-transform:uppercase; margin-bottom:6px;">ZIEL-PROGNOSE</div>
              <div style="padding:6px 14px; background:${statusColor}15; color:${statusColor}; border-radius:10px; font-size:16px; font-weight:800; border:1px solid ${statusColor}30;">
                ${forecastRatio}%
              </div>
            </div>
          </div>

          <div style="margin-bottom:10px; padding: 0 5px;">
            <div style="display:flex; justify-content:space-between; margin-bottom:8px;">
              <span style="font-size:10px; font-weight:700; color:#717786; text-transform:uppercase;">JAHRES-PERFORMANCE (YTD)</span>
              <span style="font-size:11px; font-weight:700; color:#FF5C5C;">-1.620 zum Jahresziel</span>
            </div>
            <div style="height:8px; background:rgba(255,255,255,0.05); border-radius:4px; overflow:hidden;">
              <div style="background:linear-gradient(90deg, #6E2DA0, #9B6DC8); width:68%; height:100%;"></div>
            </div>
          </div>
        </div>
      </div>`
    };
  } catch (error) {
    return { summaryText: `Fehler: ${error.message}` };
  }
}

/** --- LOGIK-HELFER --- **/

function _generateLocationInsights(data, goals, period, stats, forecast, isCurrentMonth) {
  const perf = [];
  Object.keys(data).forEach(city => {
    const ist = Array.isArray(data[city][period]) ? data[city][period].reduce((a, b) => a + (b || 0), 0) : (data[city][period] || 0);
    const ziel = goals[city]?.[period] || 0;
    if (ziel > 0) perf.push({ name: city, ratio: (ist / ziel) * 100, gap: ziel - ist });
  });

  if (!perf.length) return { summaryText: "Datenbasis für Analyse unzureichend." };
  
  perf.sort((a, b) => b.ratio - a.ratio);
  const top = perf[0], bottom = perf[perf.length - 1];

  const diffToYear = stats.ratioMonth - stats.ratioYtd;
  const trendLabel = diffToYear > 0 ? `übersteigt mit ${Math.abs(diffToYear)}%` : `liegt mit ${Math.abs(diffToYear)}% unter`;
  const gapAbs = stats.zielMonth - stats.istMonth;
  
  let text = `Der aktuelle Monat <strong>${trendLabel}</strong> den bisherigen Jahresschnitt (${stats.ratioYtd}%). `;
  text += `Spitzenreiter bei den Aktivierungen ist aktuell <strong>${top.name}</strong> (${Math.round(top.ratio)}%). `;
  
  if (isCurrentMonth) {
    if (gapAbs > 0) {
      text += `Systemweit fehlen noch <strong>${gapAbs.toLocaleString('de-DE')} Aktivierungen</strong> zum Monatsziel. `;
      if (forecast) {
        text += `Bei aktuellem Tempo wird ein Abschluss von <strong>${forecast.ratio}%</strong> prognostiziert. `;
      }
    }
  }

  if (bottom.ratio < 85) {
    text += `In <strong>${bottom.name}</strong> ist die Erreichungsquote mit ${Math.round(bottom.ratio)}% derzeit am geringsten ausgeprägt.`;
  }

  return { summaryText: text };
}

function _calculateSummaryStats(data, goals, year, latestPeriod) {
  let istYtd = 0, zielYtd = 0, istMonth = 0, zielMonth = 0;
  const yearStr = String(year);
  Object.keys(data).forEach(city => {
    Object.entries(data[city]).forEach(([p, val]) => {
      if (p.startsWith(yearStr)) {
        const sum = Array.isArray(val) ? val.reduce((a, b) => a + (b || 0), 0) : (val || 0);
        istYtd += sum;
        if (p === latestPeriod) istMonth += sum;
      }
    });
  });
  Object.keys(goals).forEach(city => {
    Object.entries(goals[city]).forEach(([p, val]) => {
      if (p.startsWith(yearStr)) {
        zielYtd += (val || 0);
        if (p === latestPeriod) zielMonth += (val || 0);
      }
    });
  });
  return { istYtd, zielYtd, istMonth, zielMonth, ratioYtd: zielYtd > 0 ? Math.round((istYtd / zielYtd) * 100) : 0, ratioMonth: zielMonth > 0 ? Math.round((istMonth / zielMonth) * 100) : 0 };
}

function _calculateForecast(data, year, month, latestPeriod, zielMonth) {
  let forecastTotal = 0;
  Object.keys(data).forEach(city => {
    const cityData = data[city][latestPeriod];
    if (cityData === undefined) return;
    const val = Array.isArray(cityData) ? cityData.reduce((a, b) => a + (b || 0), 0) : (cityData || 0);
    const elapsed = getElapsedWorkingDays(city, year, month);
    const remain = getRemainingWorkingDays(city, year, month);
    const pace = elapsed > 0 ? val / elapsed : 0;
    forecastTotal += Math.max(val, Math.round(val + (pace * Math.max(0, remain))));
  });
  return { total: forecastTotal, ratio: zielMonth > 0 ? Math.round((forecastTotal / zielMonth) * 100) : 0 };
}

function _calculateAvgWorkdays(data, year, month) {
  let t = 0, e = 0, r = 0, c = 0;
  Object.keys(data).forEach(city => {
    try {
      t += getWorkingDaysInMonth(city, year, month); 
      e += getElapsedWorkingDays(city, year, month); 
      r += getRemainingWorkingDays(city, year, month); 
      c++;
    } catch(err) {}
  });
  return c > 0 ? { total: Math.round(t/c), elapsed: Math.round(e/c), remaining: Math.round(r/c) } : null;
}

function callGemini(prompt) {
  try {
    const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINIAPIKEY');
    if (!apiKey) return null;
    
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash-001:generateContent?key=${apiKey}`;
    
    const payload = {
      contents: [{ parts: [{ text: prompt }] }],
      generationConfig: { 
        temperature: 0.1, 
        maxOutputTokens: 3000  // ✅ VON 300 AUF 3000 ERHÖHEN!
      }
    };
    
    const response = UrlFetchApp.fetch(url, {
      method: 'POST',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() !== 200) return null;
    
    const json = JSON.parse(response.getContentText());
    return json.candidates?.[0]?.content?.parts?.[0]?.text || null;
  } catch (error) {
    Logger.log('❌ Gemini error:', error.toString());
    return null;
  }
}


/***********************
 * EMAIL RECIPIENTS MANAGEMENT
 ***********************/

function getEmailRecipients() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const mailSheet = ss.getSheetByName(CONFIG.SHEETS.EMAIL_RECIPIENTS);
    
    if (mailSheet && mailSheet.getLastRow() >= 1) {
      const lastRow = mailSheet.getLastRow();
      const allData = mailSheet.getRange(1, 1, lastRow, 1).getValues();
      
      const emails = [];
      
      allData.forEach((row, index) => {
        const value = String(row[0]).trim();
        
        if (!value) return;
        
        if (index === 0 && (
          value.toLowerCase() === 'email' ||
          value.toLowerCase() === 'e-mail' ||
          value.toLowerCase() === 'empfänger' ||
          value.toLowerCase() === 'recipient'
        )) {
          return;
        }
        
        if (value.includes('@')) {
          emails.push(value);
        }
      });
      
      if (emails.length > 0) {
        Logger.log('✅ Email-Empfänger aus Sheet:', emails.length, '→', emails);
        return emails;
      }
    }
    
    const props = PropertiesService.getScriptProperties();
    const recipientsStr = props.getProperty('EMAIL_RECIPIENTS') || '';
    const propsEmails = recipientsStr.split(',').map(e => e.trim()).filter(e => e);
    
    Logger.log('✅ Email-Empfänger aus Properties:', propsEmails.length);
    return propsEmails;
  } catch (error) {
    Logger.log('❌ getEmailRecipients error:', error.toString());
    return [];
  }
}

function addEmailRecipient(email) {
  try {
    if (!email || !email.includes('@')) {
      return { success: false, message: 'Ungültige Email-Adresse.' };
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let mailSheet = ss.getSheetByName(CONFIG.SHEETS.EMAIL_RECIPIENTS);
    
    if (!mailSheet) {
      mailSheet = ss.insertSheet(CONFIG.SHEETS.EMAIL_RECIPIENTS);
      mailSheet.getRange(1, 1).setValue('Email');
    }
    
    const lastRow = mailSheet.getLastRow();
    const existingEmails = lastRow >= 2 
      ? mailSheet.getRange(2, 1, lastRow - 1, 1).getValues().map(r => String(r[0]).trim())
      : [];
    
    if (existingEmails.includes(email)) {
      return { success: false, message: 'Email-Adresse bereits vorhanden.' };
    }
    
    mailSheet.appendRow([email]);
    Logger.log('✅ Email hinzugefügt:', email);
    return { success: true, message: email + ' hinzugefügt.' };
  } catch (error) {
    Logger.log('❌ addEmailRecipient error:', error.toString());
    return { success: false, message: 'Fehler: ' + error.message };
  }
}

function removeEmailRecipient(email) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const mailSheet = ss.getSheetByName(CONFIG.SHEETS.EMAIL_RECIPIENTS);
    
    if (!mailSheet) {
      return { success: false, message: 'Kein Email-Sheet gefunden.' };
    }
    
    const lastRow = mailSheet.getLastRow();
    if (lastRow < 2) {
      return { success: false, message: 'Keine Empfänger vorhanden.' };
    }
    
    const data = mailSheet.getRange(2, 1, lastRow - 1, 1).getValues();
    
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim() === email) {
        mailSheet.deleteRow(i + 2);
        Logger.log('✅ Email entfernt:', email);
        return { success: true, message: email + ' entfernt.' };
      }
    }
    
    return { success: false, message: 'Email nicht gefunden.' };
  } catch (error) {
    Logger.log('❌ removeEmailRecipient error:', error.toString());
    return { success: false, message: 'Fehler: ' + error.message };
  }
}

/***********************
 * SCHEDULE SETTINGS
 ***********************/

function getScheduleSettings() {
  try {
    const props = PropertiesService.getScriptProperties();
    
    return {
      enabled: props.getProperty('SCHEDULE_ENABLED') === 'true',
      time: props.getProperty('SCHEDULE_TIME') || '08:00',
      timezone: props.getProperty('SCHEDULE_TIMEZONE') || 'Europe/Berlin'
    };
  } catch (error) {
    Logger.log('❌ getScheduleSettings error:', error.toString());
    return { enabled: false, time: '08:00', timezone: 'Europe/Berlin' };
  }
}

function saveScheduleSettings(settings) {
  try {
    const props = PropertiesService.getScriptProperties();
    
    props.setProperty('SCHEDULE_ENABLED', String(settings.enabled));
    props.setProperty('SCHEDULE_TIME', settings.time);
    props.setProperty('SCHEDULE_TIMEZONE', settings.timezone);
    
    if (settings.enabled) {
      setupDailyTrigger(settings.time);
      Logger.log('✅ Schedule aktiviert:', settings.time);
      return { success: true, message: 'Zeitplan aktiviert für ' + settings.time + ' Uhr.' };
    } else {
      removeDailyTrigger();
      Logger.log('✅ Schedule deaktiviert');
      return { success: true, message: 'Zeitplan deaktiviert.' };
    }
  } catch (error) {
    Logger.log('❌ saveScheduleSettings error:', error.toString());
    return { success: false, message: 'Fehler: ' + error.message };
  }
}

function setupDailyTrigger(time) {
  try {
    removeDailyTrigger();
    
    const [hour, minute] = time.split(':').map(Number);
    
    ScriptApp.newTrigger('sendDailyEmailSummary')
      .timeBased()
      .atHour(hour)
      .nearMinute(minute || 0)
      .everyDays(1)
      .create();
    
    Logger.log('✅ Daily Trigger erstellt für', time);
  } catch (error) {
    Logger.log('❌ setupDailyTrigger error:', error.toString());
    throw error;
  }
}

function removeDailyTrigger() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'sendDailyEmailSummary') {
        ScriptApp.deleteTrigger(trigger);
      }
    });
    
    Logger.log('✅ Daily Trigger gelöscht');
  } catch (error) {
    Logger.log('❌ removeDailyTrigger error:', error.toString());
  }
}

/***********************
 * SEND REPORT FUNCTIONS
 ***********************/

function sendReportToAll() {
  try {
    const recipients = getEmailRecipients();
    
    if (!recipients || recipients.length === 0) {
      return { success: false, message: 'Keine Empfänger konfiguriert.' };
    }
    
    sendDailyEmailSummary();
    
    Logger.log('✅ Report an', recipients.length, 'Empfänger gesendet');
    return {
      success: true,
      message: 'Report an ' + recipients.length + ' Empfänger versendet.'
    };
  } catch (error) {
    Logger.log('❌ sendReportToAll error:', error.toString());
    return { success: false, message: 'Fehler: ' + error.message };
  }
}

function sendEmailReportToAll() {
  return sendReportToAll();
}

function sendDailyEmailSummary() {
  try {
    const frontend = loadDataForFrontend();
    const data = frontend.data || {};
    const goals = frontend.goals || {};
    const forecast = frontend.forecasts || {};
    const periods = frontend.periods || [];
    const sollTD = frontend.sollToDate || {}; // ✅ muss aus loadDataForFrontend() kommen
    
    if (!periods.length) {
      Logger.log('⚠️ Keine Perioden für Email-Report');
      return;
    }

    const latestPeriod = periods[0];
    const [year, month] = latestPeriod.split('-');
    const keyMonth = latestPeriod;

    const tz = Session.getScriptTimeZone();
    const today = new Date();
    const todayStr = Utilities.formatDate(today, tz, 'dd.MM.yyyy'); // ✅ Stand HEUTE

    let cityRows = [];
    let totalIstMonth = 0;
    let totalZielMonth = 0;
    let totalSollMonth = 0;
    let totalForecast = 0;

    for (const city in data) {
      const istMonth = (data[city] && data[city][keyMonth]) ? Number(data[city][keyMonth]) : 0;
      const zielMonth = (goals[city] && goals[city][keyMonth]) ? Number(goals[city][keyMonth]) : 0;

      // ✅ Soll-to-Date bis HEUTE (kommt aus Backend)
      const citySollTD = (sollTD[city] && typeof sollTD[city].month !== 'undefined') ? Number(sollTD[city].month) : 0;

      let fcMonth = 0;
      if (forecast[city] && typeof forecast[city].forecast === 'number') {
        fcMonth = Number(forecast[city].forecast) || 0;
      }

      const ratioIstZiel = zielMonth > 0 ? istMonth / zielMonth : 0;
      const ratioIstSollTD = citySollTD > 0 ? istMonth / citySollTD : 0;

      cityRows.push({
        city,
        istMonth,
        zielMonth,
        sollMonth: citySollTD,
        forecastMonth: fcMonth,
        ratioIstZiel,
        ratioIstSollTD
      });

      totalIstMonth += istMonth;
      totalZielMonth += zielMonth;
      totalSollMonth += citySollTD;
      totalForecast += fcMonth;
    }

    cityRows.sort((a, b) => {
      let indexA = CONFIG.CITY_ORDER.indexOf(a.city);
      let indexB = CONFIG.CITY_ORDER.indexOf(b.city);

      if (indexA === -1) indexA = 999;
      if (indexB === -1) indexB = 999;

      if (indexA === 999 && indexB === 999) {
        return a.city.localeCompare(b.city);
      }
      return indexA - indexB;
    });

    const ratioIstZielTotal = totalZielMonth > 0 ? totalIstMonth / totalZielMonth : 0;
    const ratioIstSollTDTotal = totalSollMonth > 0 ? totalIstMonth / totalSollMonth : 0;

    const summary = {
      latestPeriod,
      year,
      month,
      dateLabel: todayStr, // ✅ HIER: HEUTE statt gestern
      cityRows,
      total: {
        istMonth: totalIstMonth,
        zielMonth: totalZielMonth,
        sollMonth: totalSollMonth,
        forecastMonth: totalForecast,
        ratioIstZiel: ratioIstZielTotal,
        ratioIstSollTD: ratioIstSollTDTotal
      }
    };

    const recipients = getEmailRecipients();
    if (!recipients || recipients.length === 0) {
      Logger.log('⚠️ Keine Email-Empfänger konfiguriert');
      return;
    }

    const htmlBody = buildEmailHtml_(summary);
    const subject = `Aktivierungsreport ${month}/${year} – Stand ${todayStr}`; // ✅ HEUTE

    MailApp.sendEmail({
      to: recipients[0],
      bcc: recipients.slice(1).join(','),
      subject,
      htmlBody
    });

    Logger.log(`✅ Email-Report an ${recipients.length} Empfänger gesendet`);
  } catch (error) {
    Logger.log('❌ sendDailyEmailSummary error: ' + error.toString());
  }
}

function buildEmailHtml_(summary) {
  const fmtInt = n => Utilities.formatString('%s', Math.round(n || 0)).replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  const fmtPct = r => (isFinite(r) && r > 0 ? Utilities.formatString('%.0f%%', r * 100) : '–');
  
  const monthNameDe = {
    '01': 'Januar', '02': 'Februar', '03': 'März', '04': 'April',
    '05': 'Mai', '06': 'Juni', '07': 'Juli', '08': 'August',
    '09': 'September', '10': 'Oktober', '11': 'November', '12': 'Dezember'
  }[summary.month] || summary.month;

  const total = summary.total;
  const deltaFcIst = total.forecastMonth - total.istMonth;
  const deltaSign = deltaFcIst >= 0 ? '+' : '';

  let trendColor = CONFIG.BRAND.textLight;
  let trendIcon = '▶';
  
  if (total.ratioIstSollTD >= CONFIG.THRESHOLDS.green) {
    trendColor = CONFIG.BRAND.success;
    trendIcon = '▲';
  } else if (total.ratioIstSollTD >= CONFIG.THRESHOLDS.yellow) {
    trendColor = CONFIG.BRAND.warning;
    trendIcon = '▶';
  } else {
    trendColor = CONFIG.BRAND.danger;
    trendIcon = '▼';
  }

  const EMAIL_STYLES = {
    wrapper: `background-color: ${CONFIG.BRAND.bg}; font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; padding: 40px 0;`,
    main: `margin: 0 auto; background-color: ${CONFIG.BRAND.white}; border-radius: 12px; overflow: hidden; max-width: 720px; box-shadow: 0 10px 30px rgba(0,0,0,0.08); border: 1px solid ${CONFIG.BRAND.border};`,
    header: `background: linear-gradient(135deg, ${CONFIG.BRAND.purple} 0%, ${CONFIG.BRAND.purpleDark} 100%); padding: 48px;`,
    headerTitle: `margin: 0 0 6px 0; font-size: 32px; font-weight: 300; color: ${CONFIG.BRAND.white}; letter-spacing: -0.5px;`,
    headerSubtitle: `margin: 0; font-size: 14px; font-weight: 500; color: rgba(255,255,255,0.8); text-transform: uppercase;`,
    headerMeta: `margin: 0; font-size: 11px; font-weight: 600; color: rgba(255,255,255,0.7); text-transform: uppercase; text-align: right;`,
    dashboard: `padding: 40px 48px 20px 48px;`,
    card: `background-color: ${CONFIG.BRAND.goldLight}; border: 1px solid ${CONFIG.BRAND.goldAccent}; border-radius: 10px; padding: 24px 15px; text-align: center; min-height: 240px;`,
    iconCircle: `display: inline-block; width: 48px; height: 48px; border-radius: 50%; background-color: ${CONFIG.BRAND.white}; margin-bottom: 16px; box-shadow: 0 4px 10px rgba(0,0,0,0.05);`,
    kpiLabel: `font-size: 11px; font-weight: 700; color: ${CONFIG.BRAND.textLight}; text-transform: uppercase; margin-bottom: 8px; display: block;`,
    kpiValue: `font-size: 32px; font-weight: 300; color: ${CONFIG.BRAND.purple}; margin: 0; line-height: 1;`,
    kpiSub: `font-size: 12px; color: ${CONFIG.BRAND.textDark}; margin: 8px 0 0 0;`,
    sectionDivider: `padding: 20px 48px 0 48px;`,
    sectionTitle: `font-size: 14px; font-weight: 700; color: ${CONFIG.BRAND.textDark}; text-transform: uppercase; display: flex; align-items: center;`,
    sectionIconBox: `margin-right: 12px; width: 28px; height: 28px; border-radius: 6px; background: ${CONFIG.BRAND.goldLight}; text-align: center;`,
    tableSection: `padding: 10px 48px 48px 48px;`,
    tableWrapper: `border: 1px solid ${CONFIG.BRAND.border}; border-radius: 8px; overflow: hidden;`,
    table: `width: 100%; border-collapse: collapse; font-size: 13px;`,
    th: `padding: 16px 14px; text-align: right; color: ${CONFIG.BRAND.purple}; font-size: 10px; font-weight: 700; text-transform: uppercase; border-bottom: 2px solid ${CONFIG.BRAND.gold}; background-color: ${CONFIG.BRAND.goldLight};`,
    thLeft: `text-align: left; padding-left: 20px;`,
    td: `padding: 16px 14px; border-bottom: 1px solid ${CONFIG.BRAND.border}; color: ${CONFIG.BRAND.textDark}; text-align: right;`,
    tdLeft: `text-align: left; font-weight: 600; padding-left: 20px;`,
    trTotal: `background-color: ${CONFIG.BRAND.purple};`,
    tdTotal: `padding: 18px 14px; color: ${CONFIG.BRAND.white}; font-weight: 700; border-top: 3px solid ${CONFIG.BRAND.gold};`,
    tdTotalLeft: `padding: 18px 20px; color: ${CONFIG.BRAND.white}; text-align: left; font-weight: 700; border-top: 3px solid ${CONFIG.BRAND.gold};`,
    footer: `padding: 32px 48px; background-color: ${CONFIG.BRAND.goldLight}; border-top: 1px solid ${CONFIG.BRAND.border};`,
    footerLogo: `font-size: 18px; font-weight: 700; color: ${CONFIG.BRAND.purple}; margin: 0 0 8px 0;`,
    footerText: `font-size: 11px; color: ${CONFIG.BRAND.textLight}; line-height: 1.7; margin: 0;`
  };

  const renderIconInCircle = (url) => `
    <div style="${EMAIL_STYLES.iconCircle}">
      <img src="${url}" width="24" height="24" style="margin-top:12px; border:0;">
    </div>
  `;

  const headerHtml = `
    <tr>
      <td style="${EMAIL_STYLES.header}">
        <table width="100%" cellpadding="0" cellspacing="0" border="0">
          <tr>
            <td valign="bottom" align="left" style="width: 65%;">
              <h1 style="${EMAIL_STYLES.headerTitle}">Vertragsaktivierung</h1>
              <p style="${EMAIL_STYLES.headerSubtitle}">Monatliches Performance Dashboard</p>
            </td>
            <td valign="bottom" align="right" style="width: 35%;">
              <p style="${EMAIL_STYLES.headerMeta}">
                ${monthNameDe.toUpperCase()} ${summary.year}<br>
                <span style="opacity: 0.6; font-size: 10px;">REPORT-DATUM: ${summary.dateLabel}</span>
              </p>
            </td>
          </tr>
        </table>
      </td>
    </tr>
  `;

  const dashboardHtml = `
    <tr>
      <td style="${EMAIL_STYLES.dashboard}">
        <table width="100%" cellpadding="0" cellspacing="0" border="0">
          <tr>
            <td width="32%" valign="top" style="padding-right:2%;">
              <div style="${EMAIL_STYLES.card}">
                ${renderIconInCircle(CONFIG.ICON_URLS.check)}
                <span style="${EMAIL_STYLES.kpiLabel}">Ist-Stand Aktuell</span>
                <h3 style="${EMAIL_STYLES.kpiValue}">${fmtInt(total.istMonth)}</h3>
                <div style="${EMAIL_STYLES.kpiSub}">
                   von ${fmtInt(total.zielMonth)} Monatsziel<br>
                   <span style="color:${CONFIG.BRAND.textLight}; opacity:0.8;">(${fmtPct(total.ratioIstZiel)} erreicht)</span>
                </div>
              </div>
            </td>
            
            <td width="32%" valign="top" style="padding-right:2%;">
              <div style="${EMAIL_STYLES.card}">
                ${renderIconInCircle(CONFIG.ICON_URLS.chart)}
                <span style="${EMAIL_STYLES.kpiLabel}">Forecast Monatsende</span>
                <h3 style="${EMAIL_STYLES.kpiValue}">${fmtInt(total.forecastMonth)}</h3>
                <div style="${EMAIL_STYLES.kpiSub}">
                   Erwartete Aktivierungen<br>
                   <span style="color:${CONFIG.BRAND.textLight};">Delta: ${deltaSign}${fmtInt(deltaFcIst)}</span>
                </div>
              </div>
            </td>

            <td width="32%" valign="top">
              <div style="${EMAIL_STYLES.card}">
                ${renderIconInCircle(CONFIG.ICON_URLS.target)}
                <span style="${EMAIL_STYLES.kpiLabel}">Soll vs. Ist</span>
                <h3 style="${EMAIL_STYLES.kpiValue}">${fmtPct(total.ratioIstSollTD)}</h3>
                <div style="${EMAIL_STYLES.kpiSub}">
                   Time-Distributed Target<br>
                   <span style="color:${trendColor}; font-weight:700; font-size:16px;">${trendIcon}</span>
                   <span style="color:${CONFIG.BRAND.textLight}; font-size:10px; display:block; margin-top:2px;">Soll kumuliert: ${fmtInt(total.sollMonth)}</span>
                </div>
              </div>
            </td>
          </tr>
        </table>
      </td>
    </tr>
  `;

  const sectionHtml = `
    <tr>
      <td style="${EMAIL_STYLES.sectionDivider}">
        <div style="${EMAIL_STYLES.sectionTitle}">
           <div style="${EMAIL_STYLES.sectionIconBox}">
             <img src="${CONFIG.ICON_URLS.table}" width="14" height="14" style="margin-top:7px;">
           </div>
           Standort-Performance Details
        </div>
      </td>
    </tr>
  `;

  let rowsHtml = '';
  summary.cityRows.forEach((row, index) => {
    const rowBg = (index % 2 !== 0) ? `background-color:${CONFIG.BRAND.stripeGray};` : `background-color:${CONFIG.BRAND.white};`;
    
    let perfStyle = '';
    if (row.ratioIstSollTD >= CONFIG.THRESHOLDS.green) perfStyle = `color:${CONFIG.BRAND.success}; font-weight:bold;`;
    else if (row.ratioIstSollTD < CONFIG.THRESHOLDS.yellow) perfStyle = `color:${CONFIG.BRAND.danger};`;

    rowsHtml += `
      <tr style="${rowBg}">
        <td style="${EMAIL_STYLES.td} ${EMAIL_STYLES.tdLeft}">${row.city}</td>
        <td style="${EMAIL_STYLES.td}"><strong>${fmtInt(row.istMonth)}</strong></td>
        <td style="${EMAIL_STYLES.td}">${fmtInt(row.sollMonth)}</td>
        <td style="${EMAIL_STYLES.td}">${fmtInt(row.zielMonth)}</td>
        <td style="${EMAIL_STYLES.td}">${fmtInt(row.forecastMonth)}</td>
        <td style="${EMAIL_STYLES.td}">${fmtPct(row.ratioIstZiel)}</td>
        <td style="${EMAIL_STYLES.td} ${perfStyle}">${fmtPct(row.ratioIstSollTD)}</td>
      </tr>
    `;
  });

  const totalRowHtml = `
    <tr style="${EMAIL_STYLES.trTotal}">
      <td style="${EMAIL_STYLES.tdTotalLeft}">GESAMT</td>
      <td style="${EMAIL_STYLES.tdTotal}">${fmtInt(total.istMonth)}</td>
      <td style="${EMAIL_STYLES.tdTotal}">${fmtInt(total.sollMonth)}</td>
      <td style="${EMAIL_STYLES.tdTotal}">${fmtInt(total.zielMonth)}</td>
      <td style="${EMAIL_STYLES.tdTotal}">${fmtInt(total.forecastMonth)}</td>
      <td style="${EMAIL_STYLES.tdTotal}">${fmtPct(total.ratioIstZiel)}</td>
      <td style="${EMAIL_STYLES.tdTotal}">${fmtPct(total.ratioIstSollTD)}</td>
    </tr>
  `;

  const tableHtml = `
    <tr>
      <td style="${EMAIL_STYLES.tableSection}">
        <div style="${EMAIL_STYLES.tableWrapper}">
          <table cellpadding="0" cellspacing="0" border="0" style="${EMAIL_STYLES.table}">
            <thead>
              <tr>
                <th style="${EMAIL_STYLES.th} ${EMAIL_STYLES.thLeft}">Standort</th>
                <th style="${EMAIL_STYLES.th}">Ist</th>
                <th style="${EMAIL_STYLES.th}">Soll (TD)</th>
                <th style="${EMAIL_STYLES.th}">Ziel</th>
                <th style="${EMAIL_STYLES.th}">Fcst</th>
                <th style="${EMAIL_STYLES.th}">% Ziel</th>
                <th style="${EMAIL_STYLES.th}">% Soll</th>
              </tr>
            </thead>
            <tbody>
              ${rowsHtml}
              ${totalRowHtml}
            </tbody>
          </table>
        </div>
      </td>
    </tr>
  `;

  const footerHtml = `
    <tr>
      <td style="${EMAIL_STYLES.footer}">
        <table width="100%" cellpadding="0" cellspacing="0" border="0">
          <tr>
            <td valign="top" align="left" style="width: 60%;">
              <p style="${EMAIL_STYLES.footerLogo}">TRISOR GROUP</p>
              <p style="${EMAIL_STYLES.footerText}">
                <strong>Performance Reporting System</strong><br>
                Automatisch generiert · Live-Daten aus System
              </p>
              <p style="margin: 12px 0 0 0; font-size: 11px;">
                <a href="https://mysync.short.gy/forecast" style="color: ${CONFIG.BRAND.purple}; text-decoration: none; font-weight: bold; padding: 6px 12px; background-color: #fff; border-radius: 4px; border: 1px solid ${CONFIG.BRAND.purple}; display:inline-block;">
                  Dashboard öffnen →
                </a>
              </p>
            </td>
            <td valign="top" align="right" style="width: 40%;">
              <p style="font-size: 10px; color: ${CONFIG.BRAND.textLight}; text-align: right; line-height: 1.6; margin: 0;">
                <strong>METADATEN</strong><br>
                Datum: ${summary.dateLabel}<br>
                Periode: ${monthNameDe} ${summary.year}<br>
                <span style="opacity: 0.7;">© ${summary.year} Trisor Group</span>
              </p>
            </td>
          </tr>
        </table>
      </td>
    </tr>
  `;

  return `
    <div style="${EMAIL_STYLES.wrapper}">
      <table align="center" border="0" cellpadding="0" cellspacing="0" width="100%" style="${EMAIL_STYLES.main}">
        ${headerHtml}
        ${dashboardHtml}
        ${sectionHtml}
        ${tableHtml}
        ${footerHtml}
      </table>
    </div>
  `;
}
function sendPdfReportViaEmail(email, from, to) {
  try {
    if (!email) throw new Error('Email fehlt');
    if (!from || !to) throw new Error('from/to fehlen');
    if (from > to) throw new Error('"Von" darf nicht nach "Bis" liegen');

    const config = {
      format: 'pdf',
      startMonth: from,
      endMonth: to,
      sections: {
        executiveSummary: true,
        summary: true,
        standorte: true,
        ranking: true,
        sollVergleich: true,
        vorjahr: true
      }
    };

    const gen = generateConfigurableReport(config);
    if (!gen || !gen.success) throw new Error(gen?.error || 'PDF-Generierung fehlgeschlagen');

    const bytes = Utilities.base64Decode(gen.data);
    const blob = Utilities.newBlob(bytes, 'application/pdf', gen.filename);

    MailApp.sendEmail({
      to: email,
      subject: `TRISOR Report ${from} – ${to}`,
      htmlBody: `Hi,<br><br>anbei der Report für den Zeitraum <b>${from}</b> bis <b>${to}</b>.<br><br>Beste Grüße`,
      attachments: [blob]
    });

    return { success: true, message: `PDF wurde an ${email} versendet (${from} – ${to}).` };

  } catch (e) {
    Logger.log('❌ sendPdfReportViaEmail Error: ' + e);
    return { success: false, message: e.toString() };
  }
}



/***********************
 * KONFIGURIERBARER REPORT (PREMIUM PDF SYSTEM) – FINAL BLOCK
 * - No Drive needed (returns base64 for browser download)
 * - Zeitraum korrekt in allen Bereichen
 * - Top 10 Promo + Top 10 Mitarbeiter (Zeitraum-aggregiert)
 * - Deckblatt-Zeitraum in Langform (August 2022 - Dezember 2025)
 * - KI-Insights im PDF
 ***********************/

/***********************
 * 1) PUBLIC ENTRYPOINTS
 ***********************/

function generatePdfReport(from, to) {
  try {
    Logger.log('📄 generatePdfReport(from,to) START: ' + from + ' -> ' + to);

    if (!from || !to) throw new Error('from/to fehlen');
    if (from > to) throw new Error('"Von" darf nicht nach "Bis" liegen');

    const config = {
      format: 'pdf',
      startMonth: from,
      endMonth: to,
      cities: null, // alle
      sections: {
        executiveSummary: true,
        summary: true,
        standorte: true,
        ranking: true,
        sollVergleich: true,
        vorjahr: true
      }
    };

    const gen = generateConfigurableReport(config);
    if (!gen || !gen.success) throw new Error(gen?.error || 'PDF-Generierung fehlgeschlagen');

    // ✅ Kein Drive: Base64 + Filename
    return {
      success: true,
      data: gen.data,
      filename: gen.filename
    };

  } catch (e) {
    Logger.log('❌ generatePdfReport Error: ' + e);
    return { success: false, error: e.toString() };
  }
}

function generateConfigurableReport(config) {
  try {
    Logger.log('📊 generateConfigurableReport START');
    Logger.log('   Config: ' + JSON.stringify(config));

    const reportData = buildReportData(config);

    if (config.format === 'pdf') {
      const html = buildPremiumReportHtml(reportData);
      const pdfBytes = Utilities.newBlob(html, 'text/html').getAs('application/pdf').getBytes();

      return {
        success: true,
        data: Utilities.base64Encode(pdfBytes),
        filename: `TRISOR_Report_${config.startMonth}_${config.endMonth}_${formatDateForFilename_(new Date())}.pdf`
      };
    } else {
      const csv = buildReportCsv(reportData);
      return {
        success: true,
        data: csv,
        filename: `TRISOR_Report_${config.startMonth}_${config.endMonth}_${formatDateForFilename_(new Date())}.csv`
      };
    }

  } catch (error) {
    Logger.log('❌ generateConfigurableReport Error: ' + error);
    return { success: false, error: error.toString() };
  }
}


/***********************
 * 2) REPORT DATA COMPOSITION
 ***********************/

function buildReportData(config) {
  const allData = loadDataForFrontend();

  const filteredPeriods = _getFilteredPeriods_(allData.periods || [], config.startMonth, config.endMonth);
  const cities = config.cities || Object.keys(allData.data || {});

  const reportData = {
    timestamp: new Date(),
    // ✅ Langformat:
    periodLabel: _formatRangeLongDE_(config.startMonth, config.endMonth),
    periodFrom: config.startMonth,
    periodTo: config.endMonth,
    cities,
    periods: filteredPeriods
  };

  // Standard Sections
  if (config.sections?.executiveSummary) reportData.executiveSummary = buildExecutiveSummary(allData, cities, filteredPeriods);
  if (config.sections?.summary)          reportData.summary          = buildSummary(allData, cities, filteredPeriods);
  if (config.sections?.standorte)        reportData.standorte        = buildStandorteData(allData, cities, filteredPeriods);
  if (config.sections?.ranking)          reportData.ranking          = buildRankingData(allData, cities, filteredPeriods);
  if (config.sections?.sollVergleich)    reportData.sollVergleich    = buildSollVergleich(allData, cities, filteredPeriods);
  if (config.sections?.vorjahr)          reportData.vorjahr          = buildVorjahrVergleich(allData, cities, filteredPeriods);

  if (config.format === 'pdf') {
    reportData.zeitreihe = buildZeitreihe(allData, cities, filteredPeriods);
  }

  // Premium Sections
  reportData.forecastMetrics = buildForecastMetrics(allData, cities, filteredPeriods);

  // ✅ Top10 Promo / Mitarbeiter
  reportData.promoCodes = buildPromoCodeAnalysis(allData, filteredPeriods).slice(0, 10);
  reportData.mitarbeiter = buildMitarbeiterAnalyse(allData, filteredPeriods).slice(0, 10);

  // ✅ KI-Text
  reportData.aiInsights = buildAiNarrative(reportData);

  return reportData;
}


/***********************
 * 3) STANDARD BUILDERS
 ***********************/

function buildExecutiveSummary(allData, cities, periods) {
  let totalIst = 0;
  let totalZiel = 0;

  cities.forEach(city => {
    periods.forEach(period => {
      totalIst += _sumMaybeArray_(allData.data?.[city]?.[period]);
      totalZiel += (allData.goals?.[city]?.[period] || 0);
    });
  });

  const fortschritt = totalZiel > 0 ? Math.round((totalIst / totalZiel) * 100) : 0;
  const differenz = totalIst - totalZiel;

  return {
    highlights: [
      { label: 'Gesamt Aktivierungen', value: totalIst, color: '#6E2DA0' },
      { label: 'Zielvorgabe', value: totalZiel, color: '#B89D73' },
      { label: 'Zielerreichung', value: fortschritt + '%', color: fortschritt >= 100 ? '#10B981' : '#F59E0B' }
    ],
    text:
      `Im Berichtszeitraum wurden ${totalIst} Aktivierungen erreicht bei einem Ziel von ${totalZiel}. ` +
      `Dies entspricht einer Zielerreichung von ${fortschritt}%. ` +
      `Die Abweichung beträgt ${differenz >= 0 ? '+' : ''}${differenz} Aktivierungen.`
  };
}

function buildSummary(allData, cities, periods) {
  let totalIst = 0;
  let totalZiel = 0;

  cities.forEach(city => {
    periods.forEach(period => {
      totalIst += _sumMaybeArray_(allData.data?.[city]?.[period]);
      totalZiel += (allData.goals?.[city]?.[period] || 0);
    });
  });

  const fortschritt = totalZiel > 0 ? Math.round((totalIst / totalZiel) * 100) : 0;
  const differenz = totalIst - totalZiel;

  // ✅ Einheitliche Statuslogik wie bei Standorten
  const cls = classifyRisk_(totalIst, totalZiel, {});

  return {
    totalIst,
    totalZiel,
    fortschritt,
    differenz,
    status: cls.risk,          // z.B. "RISIKO"
    statusClass: cls.riskClass // z.B. "badge-yellow"
  };
}

function buildStandorteData(allData, cities, periods) {
  const lastPeriod = getLastPeriod_(periods);

  return cities.map(city => {
    let ist = 0, ziel = 0;
    let context = {};

    if (REPORT_THRESHOLDS.mode === 'currentMonth') {
      ist = lastPeriod ? _sumMaybeArray_(allData.data?.[city]?.[lastPeriod]) : 0;
      ziel = lastPeriod ? (allData.goals?.[city]?.[lastPeriod] || 0) : 0;
    } else {
      // ✅ range ist default
      const sum = sumIstZiel_(allData, city, periods);
      ist = sum.ist; ziel = sum.ziel;
    }

    if (REPORT_THRESHOLDS.mode === 'sollToDate') {
      // sollToDate ist in loadDataForFrontend bereits berechnet (für aktuellen Monat)
      // Wenn lastPeriod nicht aktueller Monat ist, wird das ohnehin leer/0 sein → fallback greift
      context.sollToDate = allData.sollToDate?.[city]?.month || null;
    }

    const fortschritt = ziel > 0 ? Math.round((ist / ziel) * 100) : 0;
    const cls = classifyRisk_(ist, ziel, context);

    return {
      city,
      label: city,
      ist,
      ziel,
      fortschritt,
      trend: 0,
      risk: cls.risk,
      riskClass: cls.riskClass
    };
  });
}

function buildRankingData(allData, cities, periods) {
  const data = cities.map(city => {
    let ist = 0;
    let ziel = 0;

    periods.forEach(period => {
      ist += _sumMaybeArray_(allData.data?.[city]?.[period]);
      ziel += (allData.goals?.[city]?.[period] || 0);
    });

    const fortschritt = ziel > 0 ? Math.round((ist / ziel) * 100) : 0;
    return { city, ist, ziel, fortschritt, label: city };
  });

  data.sort((a, b) => b.fortschritt - a.fortschritt);

  return data.map((item, index) => ({
    ...item,
    platz: index + 1
  }));
}

function buildSollVergleich(allData, cities, periods) {
  // ✅ Referenz-Monat = letzter Monat im Zeitraum
  const refPeriod = periods?.length ? periods[periods.length - 1] : null;

  return (cities || []).map(city => {
    let istMonat = 0;
    let monatSoll = 0;

    let istJahr = 0;
    let jahrSoll = 0;

    (periods || []).forEach(period => {
      const ist = _sumMaybeArray_(allData.data?.[city]?.[period]);
      const ziel = (allData.goals?.[city]?.[period] || 0);

      // YTD über gesamten Zeitraum (nicht Kalenderjahr, sondern Zeitraum-Summe)
      istJahr += ist;
      jahrSoll += ziel;

      // ✅ Monatsteil = refPeriod (Endmonat)
      if (refPeriod && period === refPeriod) {
        istMonat = ist;
        monatSoll = ziel;
      }
    });

    const diffMonat = istMonat - monatSoll;
    const diffYear = istJahr - jahrSoll;
    const fortschrittJahr = jahrSoll > 0 ? Math.round((istJahr / jahrSoll) * 100) : 0;

    return {
      city,
      label: city,
      refPeriod,       // optional: falls du es im PDF mal anzeigen willst
      istMonat,
      monatSoll,
      diffMonat,
      istJahr,
      jahrSoll,
      diffYear,
      fortschrittJahr,
      trendPct: 0
    };
  });
}

function buildVorjahrVergleich(allData, cities, periods) {
  const currentYear = periods?.[0] ? periods[0].substring(0, 4) : String(new Date().getFullYear());
  const previousYear = String(parseInt(currentYear, 10) - 1);

  return cities.map(city => {
    let current = 0;
    let previous = 0;

    periods.forEach(period => {
      current += _sumMaybeArray_(allData.data?.[city]?.[period]);

      const prevPeriod = period.replace(currentYear, previousYear);
      previous += _sumMaybeArray_(allData.data?.[city]?.[prevPeriod]);
    });

    const diffPct = previous > 0 ? Math.round(((current - previous) / previous) * 100) : 0;

    return {
      city,
      label: city,
      current,
      previous,
      diffPct,
      compareLabel: previousYear
    };
  });
}

function buildZeitreihe(allData, cities, periods) {
  const monthNames = ['Jan', 'Feb', 'Mär', 'Apr', 'Mai', 'Jun', 'Jul', 'Aug', 'Sep', 'Okt', 'Nov', 'Dez'];

  const monthMap = periods.map(p => {
    const [y, m] = p.split('-');
    return {
      key: p,
      month: monthNames[parseInt(m, 10) - 1] + ' ' + y
    };
  });

  return monthMap.map(({ key, month }) => {
    const row = { month };

    cities.forEach(city => {
      row[city + '_ist'] = _sumMaybeArray_(allData.data?.[city]?.[key]);
      row[city + '_ziel'] = (allData.goals?.[city]?.[key] || 0);
    });

    return row;
  });
}


/***********************
 * 4) PREMIUM BUILDERS (FORECAST / PROMO / MITARBEITER / KI)
 ***********************/

function buildForecastMetrics(allData, cities, periods) {
  // Endmonat des gewählten Zeitraums
  const period = periods?.length ? periods[periods.length - 1] : null;
  if (!period) return [];

  // Forecast nur sinnvoll, wenn Endmonat == aktueller Monat (weil forecasts in loadDataForFrontend() "heute" berechnet werden)
  const now = new Date();
  const currentPeriod = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
  if (period !== currentPeriod) return [];

  return (cities || []).map(city => {
    const ist = _sumMaybeArray_(allData.data?.[city]?.[period]);
    const ziel = (allData.goals?.[city]?.[period] || 0);

    const fc = allData.forecasts?.[city] || {};
    const avgDayNum = _toNumberOrNull_(fc.activationsPerWorkday);
    const reqDayNum = _toNumberOrNull_(fc.requiredPerDay);

    const zielErreicht = ziel > 0 && ist >= ziel;

    // Anzeige-Regeln SOLL/Tag
    let sollTagDisplay = '—';
    if (ziel <= 0) sollTagDisplay = '—';
    else if (zielErreicht) sollTagDisplay = '-';
    else if (reqDayNum === null) sollTagDisplay = '—';
    else if (reqDayNum <= 0) sollTagDisplay = '-';
    else sollTagDisplay = _fmt_(reqDayNum, 1);

    return {
      city,
      period,
      ist,
      ziel,
      // forecast kann fehlen → fallback auf ist
      forecast: (fc.forecast !== undefined && fc.forecast !== null) ? fc.forecast : ist,
      // Ø IST/Tag kann fehlen
      avgDay: (avgDayNum === null) ? '—' : _fmt_(avgDayNum, 1),
      sollTag: sollTagDisplay,
      zielErreicht
    };
  });
}

function buildPromoCodeAnalysis(allData, periods) {
  const promoMap = {};
  const activators = allData.activators || {};
  if (!periods?.length || !Object.keys(activators).length) return [];

  Object.values(activators).forEach(emp => {
    periods.forEach(p => {
      const md = emp?.monthlyDetails?.[p];
      if (!md || !md.promoCodes) return;

      Object.entries(md.promoCodes).forEach(([code, cnt]) => {
        const c = Number(cnt) || 0;
        if (!code) return;
        promoMap[code] = (promoMap[code] || 0) + c;
      });
    });
  });

  return Object.entries(promoMap)
    .map(([code, count]) => ({ code, count }))
    .sort((a, b) => b.count - a.count);
}

function buildMitarbeiterAnalyse(allData, periods) {
  // ✅ Zeitraum-aggregiert (nicht nur periods[0])
  const activators = allData.activators || {};
  if (!periods?.length || !Object.keys(activators).length) return [];

  const rows = Object.entries(activators).map(([key, emp]) => {
    let istTotal = 0;
    let promoTotal = 0;
    const citySet = new Set();

    periods.forEach(p => {
      istTotal += _sumMaybeArray_(emp?.monthlyData?.[p]);

      const md = emp?.monthlyDetails?.[p] || {};
      if (md?.promoCodes) {
        promoTotal += Object.values(md.promoCodes).reduce((a, b) => a + (Number(b) || 0), 0);
      }
      if (md?.cities) {
        Object.keys(md.cities).forEach(c => citySet.add(c));
      }
    });

    return {
      id: key,
      name: emp?.name || emp?.displayName || key,
      ist: istTotal,
      promos: promoTotal,
      cityCount: citySet.size
    };
  });

  rows.sort((a, b) => b.ist - a.ist);
  return rows;
}

function buildAiNarrative(reportData) {
  const s = reportData.summary;
  if (!s) return '';

  const top3 = (reportData.ranking || []).slice(0, 3);
  const bottom3 = (reportData.ranking || []).slice(-3);

  const topTxt = top3.map(x => `${x.label} (${x.fortschritt}%)`).join(', ');
  const botTxt = bottom3.map(x => `${x.label} (${x.fortschritt}%)`).join(', ');

  let tone = 'kritisch';
  if (s.fortschritt >= 100) tone = 'sehr stark';
  else if (s.fortschritt >= 90) tone = 'solide';
  else if (s.fortschritt >= 75) tone = 'ausbaufähig';

  const diffSign = s.differenz >= 0 ? '+' : '';

  return [
    `Im Berichtszeitraum ergibt sich insgesamt ein ${tone}es Bild: ${s.fortschritt}% Zielerreichung (${s.totalIst} Ist vs. ${s.totalZiel} Ziel, Abweichung ${diffSign}${s.differenz}).`,
    top3.length ? `Stärkste Standorte nach Zielerreichung: ${topTxt}.` : '',
    bottom3.length ? `Handlungsbedarf besteht insbesondere bei: ${botTxt}.` : '',
    `Empfehlung: Best Practices der Top-Standorte standardisieren (Ablauf, Argumentation, Nachfassen) und bei den unteren Standorten 1–2 Hebel priorisieren (z.B. Terminqualität & Nachfassquote, Promo-Disziplin).`
  ].filter(Boolean).join(' ');
}


/***********************
 * 5) CSV EXPORT
 ***********************/

function buildReportCsv(reportData) {
  let csv = 'Standort;Ist;Ziel;Fortschritt\n';
  if (reportData.standorte) {
    reportData.standorte.forEach(s => {
      csv += `${s.label};${s.ist};${s.ziel};${s.fortschritt}%\n`;
    });
  }
  return csv;
}


/***********************
 * 6) HELPERS (REPORT)
 ***********************/
function _formatPeriodShortDE_(ym) {
  if (!ym) return '—';
  const [y, m] = ym.split('-').map(Number);
  const d = new Date(y, m - 1, 1);
  return new Intl.DateTimeFormat('de-DE', { month: 'long', year: 'numeric' }).format(d);
}
function _asNumber_(v) {
  const n = Number(v);
  return Number.isFinite(n) ? n : 0;
}

function _sumMaybeArray_(v) {
  if (Array.isArray(v)) return v.reduce((s, a) => s + _asNumber_(a), 0);
  return _asNumber_(v);
}

function _toNumberOrNull_(v) {
  if (v === null || v === undefined || v === '') return null;
  const n = typeof v === 'string' ? parseFloat(v.replace(',', '.')) : Number(v);
  return Number.isFinite(n) ? n : null;
}

function _fmt_(n, decimals) {
  const num = _toNumberOrNull_(n);
  if (num === null) return '—';
  return num.toLocaleString('de-DE', { minimumFractionDigits: decimals, maximumFractionDigits: decimals });
}

function _getFilteredPeriods_(allPeriods, startMonth, endMonth) {
  const filtered = (allPeriods || []).filter(p => p >= startMonth && p <= endMonth);
  filtered.sort(); // asc
  return filtered;
}

function _formatPeriodLongDE_(ym) {
  if (!ym) return '';
  const [y, m] = ym.split('-').map(Number);
  const d = new Date(y, m - 1, 1);
  const monthName = new Intl.DateTimeFormat('de-DE', { month: 'long' }).format(d);
  return `${monthName.charAt(0).toUpperCase() + monthName.slice(1)} ${y}`;
}

function _formatRangeLongDE_(fromYm, toYm) {
  return `${_formatPeriodLongDE_(fromYm)} - ${_formatPeriodLongDE_(toYm)}`;
}
const REPORT_THRESHOLDS = {
  mode: 'range', // ✅ 'range' | 'currentMonth' | 'sollToDate'

  // Neue Abstufungen (für RANGE / CURRENT MONTH)
  progress: {
    veryGoodMin: 90,  // >= 90  -> SEHR GUT
    goodMin: 80,      // >= 80  -> GUT
    okMin: 70,        // >= 70  -> OK
    riskMin: 50       // >= 50  -> RISIKO
    // < 50            -> KRITISCH
  },

  sollToDateTolerance: {
    okSlackPct: 5,
    riskSlackPct: 12
  },

  noGoal: { label: 'Kein Ziel', class: 'badge-gray' }
};
function sumIstZiel_(allData, city, periods) {
  let ist = 0, ziel = 0;
  periods.forEach(p => {
    ist += _sumMaybeArray_(allData.data?.[city]?.[p]);
    ziel += (allData.goals?.[city]?.[p] || 0);
  });
  return { ist, ziel };
}

function getLastPeriod_(periods) {
  return periods?.length ? periods[periods.length - 1] : null;
}
function classifyRisk_(ist, ziel, context) {
  // Kein Ziel definiert
  if (!ziel || ziel <= 0) {
    return {
      risk: REPORT_THRESHOLDS.noGoal.label,
      riskClass: REPORT_THRESHOLDS.noGoal.class
    };
  }

  const progressPct = (ist / ziel) * 100;

  /************************************************
   * VARIANTE A: RANGE / CURRENT MONTH
   ************************************************/
  if (REPORT_THRESHOLDS.mode !== 'sollToDate') {
    if (progressPct >= REPORT_THRESHOLDS.progress.veryGoodMin) {
      return { risk: 'SEHR GUT', riskClass: 'badge-green' };
    }
    if (progressPct >= REPORT_THRESHOLDS.progress.goodMin) {
      return { risk: 'GUT', riskClass: 'badge-lightgreen' };
    }
    if (progressPct >= REPORT_THRESHOLDS.progress.okMin) {
      return { risk: 'OK', riskClass: 'badge-blue' };
    }
    if (progressPct >= REPORT_THRESHOLDS.progress.riskMin) {
      return { risk: 'RISIKO', riskClass: 'badge-yellow' };
    }
    return { risk: 'KRITISCH', riskClass: 'badge-red' };
  }

  /************************************************
   * VARIANTE B: SOLL-TO-DATE (nur sinnvoll im aktiven Monat)
   ************************************************/
  const sollToDate = context?.sollToDate;

  // Fallback, falls keine Soll-to-date-Basis existiert
  if (!sollToDate || sollToDate <= 0) {
    if (progressPct >= REPORT_THRESHOLDS.progress.veryGoodMin) {
      return { risk: 'SEHR GUT', riskClass: 'badge-green' };
    }
    if (progressPct >= REPORT_THRESHOLDS.progress.goodMin) {
      return { risk: 'GUT', riskClass: 'badge-lightgreen' };
    }
    if (progressPct >= REPORT_THRESHOLDS.progress.okMin) {
      return { risk: 'OK', riskClass: 'badge-blue' };
    }
    if (progressPct >= REPORT_THRESHOLDS.progress.riskMin) {
      return { risk: 'RISIKO', riskClass: 'badge-yellow' };
    }
    return { risk: 'KRITISCH', riskClass: 'badge-red' };
  }

  // Vergleich Ist vs Soll-to-date
  const gapToDate = sollToDate - ist;

  // Soll-to-date erreicht oder übererfüllt => mindestens OK
  if (gapToDate <= 0) {
    if (progressPct >= REPORT_THRESHOLDS.progress.veryGoodMin) return { risk: 'SEHR GUT', riskClass: 'badge-green' };
    if (progressPct >= REPORT_THRESHOLDS.progress.goodMin)     return { risk: 'GUT', riskClass: 'badge-lightgreen' };
    return { risk: 'OK', riskClass: 'badge-blue' };
  }

  const gapPct = (gapToDate / sollToDate) * 100;

  if (gapPct <= REPORT_THRESHOLDS.sollToDateTolerance.okSlackPct) {
    return { risk: 'OK', riskClass: 'badge-blue' };
  }
  if (gapPct <= REPORT_THRESHOLDS.sollToDateTolerance.riskSlackPct) {
    return { risk: 'RISIKO', riskClass: 'badge-yellow' };
  }

  return { risk: 'KRITISCH', riskClass: 'badge-red' };
}
/***********************
 * 7) PREMIUM PDF HTML
 ***********************/

function buildPremiumReportHtml(reportData) {
  const ts = formatDateForFilename_(reportData.timestamp).replace('_', ' ');

  const colors = {
    primary: '#b57cf5',
    secondary: '#0f172a',
    accent: '#3b82f6',
    bgAlt: '#f8fafc',
    border: '#e2e8f0',
    success: '#10b981',
    warning: '#f59e0b',
    danger: '#ef4444',
    muted: '#64748b'
  };

const getProgressColor = (pct) => {
  if (pct >= REPORT_THRESHOLDS.progress.veryGoodMin) return '#14532d'; // dunkelgrün
  if (pct >= REPORT_THRESHOLDS.progress.goodMin)     return '#22c55e'; // hellgrün
  if (pct >= REPORT_THRESHOLDS.progress.okMin)       return '#3b82f6'; // blau
  if (pct >= REPORT_THRESHOLDS.progress.riskMin)     return '#f59e0b'; // orange
  return '#ef4444';                                  // rot
};

  const escapeHtml = (s) => {
    if (s === null || s === undefined) return '';
    return String(s)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  };

  const toNumberOrNull = (v) => {
    if (v === null || v === undefined || v === '') return null;
    const n = typeof v === 'string' ? parseFloat(v.replace(',', '.')) : Number(v);
    return Number.isFinite(n) ? n : null;
  };

  const badgeForRisk = (zielErreicht, sollTagDisplay) => {
    if (zielErreicht) return { cls: 'badge-green', text: 'Erreicht' };
    if (sollTagDisplay === '—') return { cls: 'badge-gray', text: 'N/A' };
    return { cls: 'badge-yellow', text: 'Offen' };
  };

  // Dynamische Seitennummern
  let pageNo = 0;
  const footer = () => `
    <div class="footer">
      <div>Vertragsaktivierungen Report</div>
      <div>Seite ${pageNo}</div>
    </div>
  `;

  let html = `<!DOCTYPE html>
<html lang="de">
<head>
  <meta charset="UTF-8">
  <title>Report</title>
  <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;700&display=swap');
    * { margin: 0; padding: 0; box-sizing: border-box; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
    body { font-family: 'Inter', sans-serif; color: ${colors.secondary}; line-height: 1.5; font-size: 11px; background: white; }
    @page { size: A4; margin: 0; }
    .sheet { width: 210mm; min-height: 297mm; margin: 0 auto; position: relative; padding: 40px 50px; page-break-after: always; background: white; }
    h1 { font-size: 26px; font-weight: 700; color: ${colors.secondary}; margin-bottom: 4px; letter-spacing: -0.5px; }
    h2 { font-size: 14px; font-weight: 700; text-transform: uppercase; color: ${colors.primary}; letter-spacing: 0.5px; margin-top: 35px; margin-bottom: 15px; border-left: 3px solid ${colors.primary}; padding-left: 12px; }
    .subtitle { color: ${colors.muted}; font-size: 12px; margin-bottom: 25px; }
    .grid-2 { display: grid; grid-template-columns: 1fr 1fr; gap: 30px; }
    .grid-3 { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 20px; }
    .card { background: ${colors.bgAlt}; padding: 16px; border-radius: 6px; border: 1px solid ${colors.border}; }
    .card-label { font-size: 10px; text-transform: uppercase; color: ${colors.muted}; font-weight: 600; margin-bottom: 6px; }
    .card-value { font-size: 24px; font-weight: 700; color: ${colors.secondary}; }
    table { width: 100%; border-collapse: collapse; margin-bottom: 20px; font-size: 11px; }
    th { text-align: left; padding: 10px 8px; border-bottom: 2px solid ${colors.border}; color: ${colors.muted}; font-size: 9px; text-transform: uppercase; font-weight: 600; letter-spacing: 0.5px; }
    td { padding: 10px 8px; border-bottom: 1px solid ${colors.border}; vertical-align: middle; }
    tr:last-child td { border-bottom: none; }
    .num { text-align: right; font-variant-numeric: tabular-nums; }
    .font-bold { font-weight: 600; }
    .badge { padding: 3px 8px; border-radius: 100px; font-size: 9px; font-weight: 700; text-transform: uppercase; display: inline-block; }
    .badge-green { background: #dcfce7; color: #14532d; }      /* sehr gut */
    .badge-lightgreen { background: #ecfdf5; color: #065f46; }/* gut */
    .badge-blue { background: #dbeafe; color: #1e3a8a; }      /* ok */
    .badge-yellow { background: #fef3c7; color: #78350f; }    /* risiko */
    .badge-red { background: #fee2e2; color: #7f1d1d; }       /* kritisch */
    .badge-gray { background: #f1f5f9; color: #475569; }
    .badge-indigo { background: #e0e7ff; color: #3730a3; }
    .prog-track { width: 80px; height: 6px; background: #e2e8f0; border-radius: 3px; display: inline-block; vertical-align: middle; margin-right: 8px; overflow: hidden; }
    .prog-fill { height: 100%; border-radius: 3px; }
    .comment { background: #eff6ff; border-left: 3px solid ${colors.accent}; padding: 14px; font-size: 11px; color: ${colors.secondary}; margin: 15px 0; border-radius: 0 4px 4px 0; line-height: 1.6; }
    .sub-comment { background: ${colors.bgAlt}; padding: 10px; font-size: 10px; color: ${colors.muted}; margin-top: 5px; border-radius: 4px; }
    .footer { position: absolute; bottom: 35px; left: 50px; right: 50px; border-top: 1px solid ${colors.border}; padding-top: 12px; display: flex; justify-content: space-between; font-size: 9px; color: ${colors.muted}; }
    .cover-page { background: ${colors.secondary}; color: white; display: flex; flex-direction: column; justify-content: center; }
    .cover-header { border-left: 4px solid ${colors.primary}; padding-left: 25px; margin-bottom: 40px; }
    .cover-page h1 { color: white; font-size: 48px; line-height: 1.1; margin-bottom: 20px; }
    .cover-meta { border-top: 1px solid rgba(255,255,255,0.15); padding-top: 25px; margin-top: 50px; display: flex; justify-content: space-between; font-size: 11px; color: #94a3b8; }
  </style>
</head>
<body>`;

  // DECKBLATT
  pageNo += 1;
  html += `
  <div class="sheet cover-page">
    <div style="margin-top: auto; margin-bottom: auto;">
      <div style="text-transform:uppercase; letter-spacing:2px; font-size:12px; color:${colors.accent}; margin-bottom:30px; font-weight:600;">Trisor Management Report</div>
      <div class="cover-header">
        <h1>Vertragsaktivierungen<br>Performance Report</h1>
        <p style="font-size: 18px; color: #cbd5e1; font-weight: 300;">Berichtszeitraum: ${escapeHtml(reportData.periodLabel || 'Aktuell')}</p>
      </div>
    </div>
    <div class="cover-meta">
      <div><strong>Erstellt am:</strong><br>${escapeHtml(ts)}</div>
      <div style="text-align:right;"><strong>Umfang:</strong><br>${(reportData.cities || []).length} Standorte analysiert</div>
    </div>
  </div>`;

  // EXECUTIVE SUMMARY
  pageNo += 1;
  html += `
  <div class="sheet">
    <div style="display:flex; justify-content:space-between; align-items:flex-end; border-bottom: 2px solid ${colors.secondary}; padding-bottom: 20px; margin-bottom: 30px;">
      <div>
        <h1>Executive Summary</h1>
        <div class="subtitle" style="margin-bottom:0;">Management Übersicht & KPIs</div>
      </div>
      <div class="badge badge-gray">${escapeHtml(reportData.periodLabel || '')}</div>
    </div>

    ${reportData.executiveSummary && (reportData.executiveSummary.highlights || []).length > 0 ? `
    <div class="grid-3" style="margin-bottom: 30px;">
      ${reportData.executiveSummary.highlights.map(h => `
        <div class="card">
          <div class="card-label">${escapeHtml(h.label)}</div>
          <div class="card-value" style="color:${escapeHtml(h.color || colors.secondary)}">${escapeHtml(h.value)}</div>
        </div>
      `).join('')}
    </div>` : ''}

    <div class="grid-2">
      <div>
        <h2>Management Summary</h2>
        <div style="line-height: 1.7; text-align: justify; color:${colors.secondary};">
          ${reportData.executiveSummary ? escapeHtml(reportData.executiveSummary.text || '').replace(/\n/g, '<br>') : 'Keine Zusammenfassung verfügbar.'}
        </div>

        ${reportData.aiInsights ? `
          <div class="comment" style="margin-top:18px;">
            <div style="font-weight:800; margin-bottom:6px;">KI-Insights</div>
            ${escapeHtml(reportData.aiInsights).replace(/\n/g, '<br>')}
          </div>
        ` : ''}
      </div>

      <div>
        <h2>Gesamt Performance</h2>
        ${reportData.summary ? `
        <div class="card" style="background:white;">
          <table style="margin:0;">
            <tr><td>Gesamt Ist</td><td class="num font-bold">${escapeHtml(reportData.summary.totalIst)}</td></tr>
            <tr><td>Gesamt Ziel</td><td class="num">${escapeHtml(reportData.summary.totalZiel)}</td></tr>
            <tr><td>Erreichung</td><td class="num" style="color:${colors.primary}; font-weight:700;">${escapeHtml(reportData.summary.fortschritt)}%</td></tr>
            <tr><td>Differenz</td><td class="num" style="color:${reportData.summary.differenz >= 0 ? colors.success : colors.danger}">${reportData.summary.differenz > 0 ? '+' : ''}${escapeHtml(reportData.summary.differenz)}</td></tr>
          </table>
          <div style="text-align:center; margin-top:15px;">
          <span class="badge ${escapeHtml(reportData.summary.statusClass || 'badge-gray')}">
  Status: ${escapeHtml(reportData.summary.status || '—')}
</span>
          </div>
        </div>` : ''}
      </div>
    </div>

    ${footer()}
  </div>`;

  // STANDORT DETAILS
  if (reportData.standorte && reportData.standorte.length > 0) {
    pageNo += 1;
    html += `
    <div class="sheet">
      <h1>Standort Details</h1>
      <div class="subtitle">Detaillierte Performance nach Region</div>

      <table style="margin-top:25px;">
        <thead>
          <tr>
            <th style="width:25%">Standort</th>
            <th class="num">Ist</th>
            <th class="num">Ziel</th>
            <th style="width:25%">Fortschritt</th>
            <th class="num">Status</th>
          </tr>
        </thead>
        <tbody>
          ${reportData.standorte.map(s => {
            const pColor = getProgressColor(s.fortschritt);
            return `
            <tr style="background:${colors.bgAlt}; border-bottom:2px solid white;">
              <td class="font-bold">${escapeHtml(s.label)}</td>
              <td class="num font-bold">${escapeHtml(s.ist)}</td>
              <td class="num">${escapeHtml(s.ziel)}</td>
              <td>
                <div class="prog-track"><div class="prog-fill" style="width:${Math.min(s.fortschritt,100)}%; background:${pColor};"></div></div>
                <span style="font-size:10px; font-weight:700; color:${pColor}">${escapeHtml(s.fortschritt)}%</span>
              </td>
              <td class="num"><span class="badge ${escapeHtml(s.riskClass)}">${escapeHtml(s.risk)}</span></td>
            </tr>`;
          }).join('')}
        </tbody>
      </table>

      ${footer()}
    </div>`;
  }

  // RANKING
  if (reportData.ranking && reportData.ranking.length > 0) {
    pageNo += 1;
    html += `
    <div class="sheet">
      <h1>Ranking</h1>
      <div class="subtitle">Kompetitiver Vergleich der Standorte</div>

      <div class="grid-2">
        <div style="display:flex; align-items:flex-end; justify-content:center; gap:15px; height:220px; padding-bottom:20px; background:#fafafa; border-radius:8px;">
          ${(reportData.ranking || []).slice(0,3).sort((a,b)=> (a.platz===1?1:a.platz===3?-1:0)).map(r => {
            const height = r.platz === 1 ? '140px' : r.platz === 2 ? '100px' : '70px';
            const bg = r.platz === 1 ? colors.primary : '#94a3b8';
            return `
            <div style="display:flex; flex-direction:column; align-items:center; width:80px;">
              <div style="font-weight:bold; font-size:10px; margin-bottom:8px; text-align:center;">${escapeHtml(r.label)}</div>
              <div style="width:100%; height:${height}; background:${bg}; color:white; display:flex; flex-direction:column; justify-content:flex-end; align-items:center; padding-bottom:10px; border-radius:6px 6px 0 0; box-shadow:0 4px 6px rgba(0,0,0,0.1);">
                <div style="font-weight:bold; font-size:18px;">${escapeHtml(r.platz)}.</div>
                <div style="font-size:10px; opacity:0.9;">${escapeHtml(r.fortschritt)}%</div>
              </div>
            </div>`;
          }).join('')}
        </div>

        <div>
          <table>
            <thead>
              <tr>
                <th style="width:40px;">#</th>
                <th>Standort</th>
                <th class="num">Zielerreichung</th>
                <th class="num">Ist / Ziel</th>
              </tr>
            </thead>
            <tbody>
              ${reportData.ranking.map(r => `
                <tr>
                  <td style="font-weight:700; color:${r.platz <= 3 ? colors.primary : colors.muted};">${escapeHtml(r.platz)}.</td>
                  <td>${escapeHtml(r.label)}</td>
                  <td class="num" style="font-weight:700;">${escapeHtml(r.fortschritt)}%</td>
                  <td class="num" style="color:${colors.muted}; font-size:10px;">${escapeHtml(r.ist)} / ${escapeHtml(r.ziel)}</td>
                </tr>
              `).join('')}
            </tbody>
          </table>
        </div>
      </div>

      ${footer()}
    </div>`;
  }

  // FORECAST PAGE (Top 10)
  if (reportData.forecastMetrics && reportData.forecastMetrics.length > 0) {
    const rows = (reportData.forecastMetrics || []).slice().sort((a, b) => {
      const aEr = !!a.zielErreicht, bEr = !!b.zielErreicht;
      if (aEr !== bEr) return aEr ? 1 : -1;

      const aSoll = toNumberOrNull(a.sollTag);
      const bSoll = toNumberOrNull(b.sollTag);
      if (aSoll !== null && bSoll !== null) return bSoll - aSoll;
      if (aSoll !== null && bSoll === null) return -1;
      if (aSoll === null && bSoll !== null) return 1;

      const aDiff = (Number(a.ziel) || 0) - (Number(a.ist) || 0);
      const bDiff = (Number(b.ziel) || 0) - (Number(b.ist) || 0);
      return bDiff - aDiff;
    });

    const top10 = rows.slice(0, 10);

    pageNo += 1;
    html += `
    <div class="sheet">
      <h1>Forecast & Tagessteuerung</h1>
      <div class="subtitle">
Ø IST/Tag, SOLL/Tag und Prognose (Top 10) – Referenzmonat: ${escapeHtml(_formatPeriodLongDE_(reportData.periods?.[reportData.periods.length - 1]))}
</div>

      <div class="comment">
  <div style="font-weight:700; margin-bottom:6px;">Interpretation (Referenzmonat)</div>
  Die Kennzahlen in dieser Tabelle beziehen sich ausschließlich auf den
  <b>Referenzmonat</b> – also den <b>letzten Monat des ausgewählten Berichtszeitraums</b>.
  <br><br>
  <b>SOLL/Tag</b> zeigt, wie viele Aktivierungen pro verbleibendem Arbeitstag notwendig wären,
  um das Monatsziel noch zu erreichen.
  <ul style="margin:6px 0 0 14px;">
    <li><b>„-“</b> → Ziel im Referenzmonat bereits erreicht</li>
    <li><b>„—“</b> → kein Ziel definiert oder Berechnung nicht sinnvoll</li>
  </ul>
</div>

      <table style="margin-top:20px;">
        <thead>
          <tr>
            <th style="width:22%">Standort</th>
            <th class="num">Ist</th>
            <th class="num">Ziel</th>
            <th class="num">Prognose</th>
            <th class="num">Ø IST/Tag</th>
            <th class="num">SOLL/Tag</th>
            <th style="width:13%">Status</th>
          </tr>
        </thead>
        <tbody>
          ${top10.map(r => {
            const pct = (Number(r.ziel) || 0) > 0 ? Math.round(((Number(r.ist) || 0) / (Number(r.ziel) || 0)) * 100) : 0;
            const pColor = getProgressColor(pct);
            const badge = badgeForRisk(r.zielErreicht, r.sollTag);
            return `
            <tr>
              <td class="font-bold">${escapeHtml(r.city)}</td>
              <td class="num font-bold">${escapeHtml(r.ist)}</td>
              <td class="num">${escapeHtml(r.ziel)}</td>
              <td class="num" style="color:${colors.primary}; font-weight:700;">${escapeHtml(r.forecast)}</td>
              <td class="num" style="color:${colors.accent}; font-weight:700;">${escapeHtml(r.avgDay)}</td>
              <td class="num" style="font-weight:800; color:${r.sollTag === '-' ? colors.muted : pColor};">${escapeHtml(r.sollTag)}</td>
              <td><span class="badge ${badge.cls}">${escapeHtml(badge.text)}</span></td>
            </tr>`;
          }).join('')}
        </tbody>
      </table>

      <div class="sub-comment">
  Sortierung nach Priorität:
  Standorte mit <b>nicht erreichtem Ziel</b> im Referenzmonat
  und hohem erforderlichen <b>SOLL/Tag</b> werden zuerst angezeigt,
  da hier der höchste operative Handlungsbedarf besteht.
</div>

      ${footer()}
    </div>`;
  }

  // SOLL-VERGLEICH
  if (reportData.sollVergleich && reportData.sollVergleich.length > 0) {
    pageNo += 1;
    html += `
    <div class="sheet">
      <h1>Soll-Analyse</h1>
      <div class="subtitle">Monatsergebnis vs. Year-to-Date (YTD)</div>

      <h2>Monatsergebnis (Referenzmonat:${escapeHtml(_formatPeriodShortDE_(reportData.periods?.[reportData.periods.length - 1]))})
</h2>
      <table>
        <thead>
          <tr>
            <th>Standort</th>
            <th class="num">Ist (Monat)</th>
            <th class="num">Soll (Monat)</th>
            <th class="num">Abweichung</th>
          </tr>
        </thead>
        <tbody>
        ${reportData.sollVergleich.map(s => {
          const diff = s.diffMonat || 0;
          return `
          <tr>
            <td>${escapeHtml(s.label)}</td>
            <td class="num font-bold">${escapeHtml(s.istMonat)}</td>
            <td class="num">${escapeHtml(s.monatSoll)}</td>
            <td class="num" style="color:${diff >= 0 ? colors.success : colors.danger}; font-weight:700;">
              ${diff > 0 ? '+' : ''}${escapeHtml(diff)}
            </td>
          </tr>`;
        }).join('')}
        </tbody>
      </table>

      <h2>2. Jahresergebnis (YTD)</h2>
      <table>
        <thead>
          <tr>
            <th>Standort</th>
            <th class="num">Ist (YTD)</th>
            <th class="num">Soll (YTD)</th>
            <th class="num">Abweichung</th>
            <th class="num">Erfüllung</th>
          </tr>
        </thead>
        <tbody>
        ${reportData.sollVergleich.map(s => {
          const diffY = s.diffYear || 0;
          return `
          <tr>
            <td>${escapeHtml(s.label)}</td>
            <td class="num font-bold">${escapeHtml(s.istJahr)}</td>
            <td class="num">${escapeHtml(s.jahrSoll)}</td>
            <td class="num" style="color:${diffY >= 0 ? colors.success : colors.danger}; font-weight:700;">
              ${diffY > 0 ? '+' : ''}${escapeHtml(diffY)}
            </td>
            <td class="num"><span class="badge ${s.fortschrittJahr >= 100 ? 'badge-green' : s.fortschrittJahr >= 90 ? 'badge-blue' : 'badge-yellow'}">${escapeHtml(s.fortschrittJahr)}%</span></td>
          </tr>`;
        }).join('')}
        </tbody>
      </table>

      ${footer()}
    </div>`;
  }

  // VORJAHR
  if (reportData.vorjahr && reportData.vorjahr.length > 0) {
    pageNo += 1;
    html += `
    <div class="sheet">
      <h1>Historischer Vergleich</h1>
      <div class="subtitle">Entwicklung gegenüber der Vergleichsperiode</div>

      <table>
        <thead>
          <tr>
            <th>Standort</th>
            <th class="num">Aktuell</th>
            <th class="num">Vergleichsperiode</th>
            <th class="num">Wachstum</th>
            <th>Periode</th>
          </tr>
        </thead>
        <tbody>
          ${reportData.vorjahr.map(v => {
            const isPos = v.diffPct >= 0;
            return `
            <tr>
              <td class="font-bold">${escapeHtml(v.label)}</td>
              <td class="num font-bold">${escapeHtml(v.current)}</td>
              <td class="num" style="color:${colors.muted}">${escapeHtml(v.previous)}</td>
              <td class="num"><span class="badge ${isPos ? 'badge-green' : 'badge-red'}">${isPos ? '▲' : '▼'} ${escapeHtml(v.diffPct)}%</span></td>
              <td style="font-size:10px; color:${colors.muted}">${escapeHtml(v.compareLabel || '-')}</td>
            </tr>`;
          }).join('')}
        </tbody>
      </table>

      ${footer()}
    </div>`;
  }

  // PROMO + MITARBEITER (Top 10 jeweils)
  {
    const promo = (reportData.promoCodes || []).slice(0, 10);
    const staff = (reportData.mitarbeiter || []).slice(0, 10);

    if (promo.length > 0 || staff.length > 0) {
      pageNo += 1;
      html += `
      <div class="sheet">
        <h1>Promo & Mitarbeiter</h1>
        <div class="subtitle">Top 10 Promo Codes und Top 10 Mitarbeiter (Gesamter Berichtszeitraum)</div>

        <div class="grid-2" style="margin-top:15px;">
          <div>
            <h2>Top Promo Codes</h2>
            ${promo.length === 0 ? `
              <div class="card" style="background:white;"><div style="color:${colors.muted}; font-size:11px;">Keine Promo-Code Daten verfügbar.</div></div>
            ` : `
              <table style="margin-top:10px;">
                <thead><tr><th style="width:50px;">#</th><th>Code</th><th class="num">Nutzung</th></tr></thead>
                <tbody>
                  ${promo.map((p, idx) => `
                    <tr>
                      <td style="font-weight:700; color:${idx < 3 ? colors.primary : colors.muted};">${idx + 1}.</td>
                      <td class="font-bold">${escapeHtml(p.code)}</td>
                      <td class="num font-bold">${escapeHtml(p.count)}</td>
                    </tr>`).join('')}
                </tbody>
              </table>
            `}
          </div>

          <div>
            <h2>Top Mitarbeiter</h2>
            ${staff.length === 0 ? `
              <div class="card" style="background:white;"><div style="color:${colors.muted}; font-size:11px;">Keine Mitarbeiter-Daten verfügbar.</div></div>
            ` : `
              <table style="margin-top:10px;">
                <thead><tr><th style="width:50px;">#</th><th>Mitarbeiter</th><th class="num">Aktivierungen</th><th class="num">Promo</th><th class="num">Städte</th></tr></thead>
                <tbody>
                  ${staff.map((s, idx) => `
                    <tr>
                      <td style="font-weight:700; color:${idx < 3 ? colors.primary : colors.muted};">${idx + 1}.</td>
                      <td class="font-bold">${escapeHtml(s.name)}</td>
                      <td class="num font-bold">${escapeHtml(s.ist)}</td>
                      <td class="num">${escapeHtml(s.promos)}</td>
                      <td class="num" style="color:${colors.muted};">${escapeHtml(s.cityCount)}</td>
                    </tr>`).join('')}
                </tbody>
              </table>
            `}
          </div>
        </div>

        ${footer()}
      </div>`;
    }
  }

  // ZEITREIHE
  if (reportData.zeitreihe && reportData.zeitreihe.length > 0 && reportData.standorte && reportData.standorte.length > 0) {
    pageNo += 1;
    html += `
    <div class="sheet">
      <h1>Zeitreihen Analyse</h1>
      <div class="subtitle">Monatlicher Verlauf im Berichtszeitraum</div>

      <div style="overflow-x:auto;">
        <table style="font-size:10px;">
          <thead style="background:${colors.bgAlt};">
            <tr>
              <th style="padding:12px 8px;">Monat</th>
              ${reportData.standorte.map(s => `<th colspan="2" style="text-align:center; border-left:1px solid ${colors.border}">${escapeHtml(s.label)}</th>`).join('')}
            </tr>
            <tr>
              <th></th>
              ${reportData.standorte.map(s => `
                <th style="text-align:center; border-left:1px solid ${colors.border}; color:${colors.muted}; font-size:8px; padding:4px;">IST</th>
                <th style="text-align:center; color:${colors.muted}; font-size:8px; padding:4px;">ZIEL</th>
              `).join('')}
            </tr>
          </thead>
          <tbody>
            ${reportData.zeitreihe.map((row, idx) => `
              <tr style="${idx % 2 !== 0 ? 'background:#f9f9f9;' : ''}">
                <td style="font-weight:600;">${escapeHtml(row.month)}</td>
                ${reportData.standorte.map(s => {
                  const ist = row[s.city + '_ist'];
                  const ziel = row[s.city + '_ziel'];
                  const met = (Number(ist) || 0) >= (Number(ziel) || 0) && (Number(ziel) || 0) > 0;
                  return `
                    <td style="text-align:center; border-left:1px solid ${colors.border}; font-weight:${met ? '700' : '400'}; color:${met ? colors.success : colors.secondary};">${escapeHtml(ist)}</td>
                    <td style="text-align:center; color:${colors.muted};">${escapeHtml(ziel)}</td>
                  `;
                }).join('')}
              </tr>
            `).join('')}
          </tbody>
        </table>
      </div>

      ${footer()}
    </div>`;
  }

  html += `</body></html>`;
  return html;
}
/**
 * Mitarbeiter-Details für Team Analytics (DEBUG-VERSION)
 */
function getEmployeeDetails(employeeEmail) {
  try {
    Logger.log('🔍 getEmployeeDetails aufgerufen mit: "' + employeeEmail + '"');
    Logger.log('   Typ: ' + typeof employeeEmail);
    Logger.log('   Länge: ' + employeeEmail.length);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const liveSheet = ss.getSheetByName('LiveData');
    
    if (!liveSheet) {
      Logger.log('❌ LiveData Sheet nicht gefunden!');
      return { total: 0, currentMonth: 0, average: 0, bestWeek: 0, cities: {}, boxSizes: {}, pricePlans: {}, promoCodes: {}, weekdays: {}, monthlyData: {}, totalRevenue: 0, avgRevenuePerYear: 0, trend: null };
    }
    
    const data = liveSheet.getDataRange().getValues();
    const headers = data[0];
    
    // Spalte 27 ist contract_activated_by
    const activatorCol = 27;
    
    Logger.log('📊 Gesamt Zeilen: ' + (data.length - 1));
    
    // Alle E-Mails ausgeben
    const allEmails = [];
    for (let i = 1; i < Math.min(data.length, 20); i++) {
      const email = data[i][activatorCol];
      if (email && allEmails.indexOf(email) === -1) {
        allEmails.push(email);
      }
    }
    Logger.log('📧 Gefundene E-Mails (erste 20 Zeilen): ' + JSON.stringify(allEmails));
    
    // Filter mit TRIM!
    const rows = [];
    for (let i = 1; i < data.length; i++) {
      const rowEmail = String(data[i][activatorCol] || '').trim();
      const searchEmail = String(employeeEmail || '').trim();
      
      if (rowEmail === searchEmail) {
        rows.push(data[i]);
      }
    }
    
    Logger.log('✅ Gefilterte Zeilen: ' + rows.length);
    
    if (rows.length === 0) {
      Logger.log('⚠️ KEINE ZEILEN GEFUNDEN für: "' + employeeEmail + '"');
      return { total: 0, currentMonth: 0, average: 0, bestWeek: 0, cities: {}, boxSizes: {}, pricePlans: {}, promoCodes: {}, weekdays: {}, monthlyData: {}, totalRevenue: 0, avgRevenuePerYear: 0, trend: null };
    }
    
    // Spalten-Indizes
    const cols = {
      city: 2,
      boxSize: 6,
      pricePlan: 7,
      promoCode: 14,
      basePrice: 13,
      promoMonths: 15,
      promoPrice: 16,
      freeMonths: 17,
      activatedDate: 25
    };
    
    // Zähler
    const cities = {};
    const boxSizes = { S: 0, M: 0, L: 0 };
    const pricePlans = {};
    const promoCodes = {};
    const weekdays = {};
    const monthlyData = {};
    const weeklyData = {};
    
    let totalRevenue = 0;
    let totalContracts = rows.length;
    
    const now = new Date();
    const currentYear = now.getFullYear();
    const currentMonth = now.getMonth() + 1;
    const currentPeriod = currentYear + '-' + String(currentMonth).padStart(2, '0');
    
    let currentMonthCount = 0;
    
    const WEEKDAY_NAMES = ['Sonntag', 'Montag', 'Dienstag', 'Mittwoch', 'Donnerstag', 'Freitag', 'Samstag'];
    
    // Daten aggregieren
    rows.forEach(function(row) {
      // Stadt
      const city = row[cols.city] || 'Unbekannt';
      cities[city] = (cities[city] || 0) + 1;
      
      // Box-Größe
      const boxSize = String(row[cols.boxSize] || '').toUpperCase();
      if (boxSize === 'S' || boxSize === 'M' || boxSize === 'L') {
        boxSizes[boxSize]++;
      }
      
      // Tarif
      const plan = row[cols.pricePlan] || 'Unbekannt';
      pricePlans[plan] = (pricePlans[plan] || 0) + 1;
      
      // Promo-Code
      const promo = row[cols.promoCode];
      if (promo && promo !== '') {
        promoCodes[promo] = (promoCodes[promo] || 0) + 1;
      }
      
      // Datum parsen
      const dateStr = row[cols.activatedDate];
      if (dateStr) {
        let date = null;
        
        if (dateStr instanceof Date) {
          date = dateStr;
        } else if (typeof dateStr === 'string') {
          const parts = dateStr.split('.');
          if (parts.length === 3) {
            const day = parseInt(parts[0]);
            const month = parseInt(parts[1]) - 1;
            const year = parseInt(parts[2]);
            date = new Date(year, month, day);
          }
        }
        
        if (date && !isNaN(date.getTime())) {
          // Monat
          const year = date.getFullYear();
          const month = date.getMonth() + 1;
          const period = year + '-' + String(month).padStart(2, '0');
          monthlyData[period] = (monthlyData[period] || 0) + 1;
          
          if (period === currentPeriod) {
            currentMonthCount++;
          }
          
          // Woche
          const weekNum = getWeekNumber(date);
          const weekKey = year + '-W' + String(weekNum).padStart(2, '0');
          weeklyData[weekKey] = (weeklyData[weekKey] || 0) + 1;
          
          // Wochentag
          const weekday = WEEKDAY_NAMES[date.getDay()];
          weekdays[weekday] = (weekdays[weekday] || 0) + 1;
        }
      }
      
      // Umsatz
      const basePrice = parseFloat(String(row[cols.basePrice] || '0').replace(',', '.')) || 0;
      const promoMonths = parseInt(row[cols.promoMonths] || 0) || 0;
      const promoPrice = parseFloat(String(row[cols.promoPrice] || '0').replace(',', '.')) || 0;
      const freeMonths = parseInt(row[cols.freeMonths] || 0) || 0;
      
      let yearlyRevenue = 0;
      
      if (promoMonths > 0 && promoPrice > 0) {
        const reducedMonths = Math.min(promoMonths, 12);
        const normalMonths = 12 - reducedMonths;
        yearlyRevenue = (promoPrice * reducedMonths) + (basePrice * normalMonths);
      } else {
        yearlyRevenue = basePrice * 12;
      }
      
      if (freeMonths > 0) {
        const freeReduction = basePrice * Math.min(freeMonths, 12);
        yearlyRevenue -= freeReduction;
        if (yearlyRevenue < 0) yearlyRevenue = 0;
      }
      
      totalRevenue += yearlyRevenue;
    });
    
    // Beste Woche
    let bestWeek = 0;
    Object.values(weeklyData).forEach(function(count) {
      if (count > bestWeek) bestWeek = count;
    });
    
    // Durchschnitt
    const uniqueMonths = Object.keys(monthlyData).length;
    const average = uniqueMonths > 0 ? Math.round(totalContracts / uniqueMonths) : 0;
    
    // Ø Umsatz
    const avgRevenuePerYear = totalContracts > 0 ? totalRevenue / totalContracts : 0;
    
    // Trend
    let trend = null;
    const sortedPeriods = Object.keys(monthlyData).sort();
    if (sortedPeriods.length >= 2) {
      const currentIdx = sortedPeriods.indexOf(currentPeriod);
      if (currentIdx > 0) {
        const lastPeriod = sortedPeriods[currentIdx - 1];
        const currentCount = monthlyData[currentPeriod] || 0;
        const lastCount = monthlyData[lastPeriod] || 0;
        
        if (lastCount > 0) {
          const change = Math.round(((currentCount - lastCount) / lastCount) * 100);
          trend = {
            change: change,
            lastMonth: lastCount
          };
        }
      }
    }
    
    const result = {
      total: totalContracts,
      currentMonth: currentMonthCount,
      average: average,
      bestWeek: bestWeek,
      cities: cities,
      boxSizes: boxSizes,
      pricePlans: pricePlans,
      promoCodes: promoCodes,
      weekdays: weekdays,
      monthlyData: monthlyData,
      totalRevenue: totalRevenue,
      avgRevenuePerYear: avgRevenuePerYear,
      trend: trend
    };
    
    Logger.log('📊 RESULT: ' + JSON.stringify(result));
    
    return result;
    
  } catch (error) {
    Logger.log('❌ FEHLER: ' + error.toString());
    return { total: 0, currentMonth: 0, average: 0, bestWeek: 0, cities: {}, boxSizes: {}, pricePlans: {}, promoCodes: {}, weekdays: {}, monthlyData: {}, totalRevenue: 0, avgRevenuePerYear: 0, trend: null };
  }
}

function getWeekNumber(date) {
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  return Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
}

/***********************
 * HELPER FUNCTIONS
 ***********************/

function formatDateForFilename_(date) {
  if (!date) date = new Date();
  const d = new Date(date);
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  const hours = String(d.getHours()).padStart(2, '0');
  const mins = String(d.getMinutes()).padStart(2, '0');
  return year + '-' + month + '-' + day + '_' + hours + '-' + mins;
}

/***********************
 * MAIN FRONTEND DATA
 ***********************/
 function _cacheKey_(prefix, city, year, month) {
  return `${prefix}:${city}:${year}-${String(month).padStart(2,'0')}`;
}

function getWorkingDaysInMonthCached(city, year, month) {
  const cache = CacheService.getScriptCache();
  const key = _cacheKey_('TW', city, year, month);
  const c = cache.get(key);
  if (c !== null) return Number(c);

  const val = getWorkingDaysInMonth(city, year, month);
  cache.put(key, String(val), 60 * 60 * 6); // 6h
  return val;
}

function getElapsedWorkingDaysInclTodayCached(city, year, month) {
  const cache = CacheService.getScriptCache();
  // Tagesabhängig → Datum in Key
  const todayKey = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const key = `${_cacheKey_('EW', city, year, month)}:${todayKey}`;
  const c = cache.get(key);
  if (c !== null) return Number(c);

  const val = getElapsedWorkingDaysInclToday(city, year, month);
  cache.put(key, String(val), 60 * 15); // 15min reicht
  return val;
}
function getElapsedWorkingDaysInclToday(city, year, month) {
  try {
    const now = new Date();

    if (now.getFullYear() !== year || (now.getMonth() + 1) !== month) {
      return getElapsedWorkingDays(city, year, month);
    }

    let ew = getElapsedWorkingDays(city, year, month);

    const dow = now.getDay();
    const isWorkday = (dow >= 1 && dow <= 6);

    if (isWorkday && !isHoliday(now, city)) {
      const todayDate = now.getDate();
      const exactInclToday = _countWorkingDaysRange_(city, year, month, 1, todayDate);

      if (ew < exactInclToday) ew = exactInclToday;
    }

    return ew;

  } catch (err) {
    Logger.log('❌ getElapsedWorkingDaysInclToday error: ' + err.toString());
    return getElapsedWorkingDays(city, year, month);
  }
}

/**
 * Interner Zähler: Werktage (Mo–Sa) innerhalb eines Monatsbereichs, Feiertage via isHoliday()
 * startDay/endDay: 1..31
 */
function _countWorkingDaysRange_(city, year, month, startDay, endDay) {
  let count = 0;
  for (let d = startDay; d <= endDay; d++) {
    const date = new Date(year, month - 1, d);
    if (isNaN(date.getTime())) continue;

    // Werktage Mo–Sa
    const dow = date.getDay();
    const isWorkday = (dow >= 1 && dow <= 6);
    if (!isWorkday) continue;

    if (isHoliday(date, city)) continue;
    count++;
  }
  return count;
}
/**
 * ============================================================
 * TRISOR ENTERPRISE BACKEND v3.2 - ENSEMBLE FORECAST ENGINE
 * ============================================================
 */

// ⚙️ KONFIGURATION - ANGEPASST AN DEIN SHEET!
const CONFIG_ENGINE = {
  sheetName: 'LiveData',
  colCity: 1,                // Spalte B = Index 1
  colDate: 25,               // Spalte Z = Index 25
  cacheKeyData: 'DASH_DATA_V3_FINAL',
  cacheKeyPatterns: 'DASH_PATTERNS_V1'
};

/**
 * 1. HAUPTFUNKTION FÜR DAS FRONTEND
 */
function loadDataForFrontend(forceRefresh) {
  // A) Cache Check
  if (!forceRefresh) {
    const cached = getFromCache(CONFIG_ENGINE.cacheKeyData);
    if (cached) {
      Logger.log('🚀 Serving from CACHE');
      return cached;
    }
  }

  try {
    Logger.log(`🔄 CALCULATING FRESH DATA (Force: ${forceRefresh})`);

    // 1. Pacing-Muster laden (Historie)
    let patterns = getFromCache(CONFIG_ENGINE.cacheKeyPatterns);
    if (!patterns) {
      // Wenn noch keine Muster da sind, einmalig berechnen
      patterns = generatePacingPatterns(); 
    }

    // 2. Live-Daten & Settings laden
    // Wir nutzen hier deine existierenden Funktionen, um Konsistenz zu wahren
    const settingsMap = getCitySettingsMap();
    const liveData = getLiveData(); 
    const goals = getGoals();

    // 3. Zeit-Konstanten
    const today = new Date();
    const currentYear = today.getFullYear();
    const currentMonth = today.getMonth() + 1;
    const currentPeriod = `${currentYear}-${String(currentMonth).padStart(2, '0')}`;

    // 4. Alle Städte sammeln
    const allCities = Array.from(new Set([
      ...Object.keys(liveData?.data || {}),
      ...Object.keys(goals || {})
    ])).sort();

    // 5. Berechnung
    const forecasts = {};
    const sollToDateMap = {};

    allCities.forEach(city => {
      // Clean Data (Schutz vor "Urlaub" Strings)
      const ist = cleanNum(liveData.data[city]?.[currentPeriod]);
      const ziel = cleanNum(goals[city]?.[currentPeriod]);

      // Arbeitstage (High Precision aus feiertage.gs)
      const totalWD = getWorkingDaysInMonth(city, currentYear, currentMonth);
      const elapsedWD = getElapsedWorkingDays(city, currentYear, currentMonth);
      const remainingWD = Math.max(0, totalWD - elapsedWD);

      // --- THE ENSEMBLE PREDICTOR ---
      // Wir holen das historische Muster
      const cityPattern = patterns ? (patterns[city] || patterns['_GLOBAL_']) : null;
      
      // Hier passiert die Magie: Mix aus Linear, Pattern und Ziel
      const prediction = calculateEnsembleForecast(ist, ziel, elapsedWD, totalWD, cityPattern);

      // --- Required Pace (Run Rate) ---
      let reqPace = 0;
      const gap = Math.max(0, ziel - ist);
      if (ziel > 0 && gap > 0) {
         reqPace = remainingWD > 0 ? (gap / remainingWD) : gap; 
      }

      // --- Soll bis Heute (Historisch korrekt) ---
      // Wir nutzen das Pattern, um zu sagen: "Wo müssten wir heute eigentlich stehen?"
      let sollBisHeute = 0;
      if (ziel > 0 && cityPattern) {
         const expectedPct = getPatternPercent(elapsedWD, totalWD, cityPattern);
         sollBisHeute = Math.round(ziel * expectedPct);
      } else {
         sollBisHeute = totalWD > 0 ? Math.round((ziel / totalWD) * elapsedWD) : 0;
      }

      forecasts[city] = {
        forecast: Math.round(prediction.forecast),
        activationsPerWorkday: prediction.currentPace.toFixed(2), // Aktueller linearer Schnitt
        requiredPerDay: reqPace.toFixed(2),
        remainingWorkdays: remainingWD,
        totalWorkdays: totalWD,
        elapsedWorkdays: elapsedWD
      };

      sollToDateMap[city] = { month: sollBisHeute };
    });

    const result = {
      data: liveData.data,
      goals: goals,
      settingsMap: settingsMap,
      forecasts: forecasts,
      sollToDate: sollToDateMap,
      heatmap: liveData.heatmap,
      tariffHeatmap: liveData.tariffHeatmap,
      activators: liveData.activators,
      periods: liveData.periods,
      latestTimestamp: liveData.lastUpdate,
      cacheTimestamp: Date.now(),
      sortedCities: allCities
    };

    saveToCache(CONFIG_ENGINE.cacheKeyData, result, 900); // 15 Min Cache
    return result;

  } catch (e) {
    Logger.log('❌ CRITICAL ERROR in loadDataForFrontend: ' + e);
    return { error: e.toString(), data: {}, goals: {} };
  }
}

/**
 * 2. PATTERN ENGINE (TRAINING)
 * Analysiert alte Monate aus LiveData (Spalte Z) und erstellt Kurven.
 * Sollte als Nightly-Trigger laufen!
 */
function generatePacingPatterns() {
  Logger.log('🧠 Starte Pattern Recognition...');
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_ENGINE.sheetName);
  if (!sheet) return {};

  // Wir lesen nur die relevanten Spalten, um RAM zu sparen
  // B (Index 1) bis Z (Index 25)
  // Wir lesen alles ab Zeile 2
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};

  const range = sheet.getRange(2, 1, lastRow - 1, 26); // Bis Spalte Z
  const data = range.getValues();
  
  const rawStats = {}; 

  // Header überspringen, Daten aggregieren
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const city = String(row[CONFIG_ENGINE.colCity]).trim(); // Spalte B
    const dateVal = row[CONFIG_ENGINE.colDate];             // Spalte Z
    
    if (!city || !dateVal) continue;

    let dateObj = parseDate(dateVal);
    if (!dateObj) continue;

    const today = new Date();
    // Nur Vergangenheit analysieren (abgeschlossene Monate)
    const isPastMonth = (dateObj.getFullYear() < today.getFullYear()) || 
                        (dateObj.getFullYear() === today.getFullYear() && dateObj.getMonth() < today.getMonth());
    
    if (!isPastMonth) continue; 

    const key = `${city}|${dateObj.getFullYear()}-${dateObj.getMonth()}`;
    if (!rawStats[key]) rawStats[key] = { city: city, year: dateObj.getFullYear(), month: dateObj.getMonth(), days: [] };
    rawStats[key].days.push(dateObj.getDate());
  }

  const finalPatterns = {};
  const cityCurves = {};

  // Kurven berechnen
  Object.values(rawStats).forEach(stat => {
    // Einfache Kalendertage-Logik für Performance
    const daysInM = new Date(stat.year, stat.month + 1, 0).getDate();
    const salesByDay = {};
    stat.days.forEach(d => salesByDay[d] = (salesByDay[d] || 0) + 1);
    
    if (stat.days.length < 5) return; // Zu wenig Daten für Muster

    let running = 0;
    const monthCurve = [];
    
    for (let d = 1; d <= daysInM; d++) {
      if (salesByDay[d]) running += salesByDay[d];
      monthCurve.push({ 
        t: d / daysInM,           // Zeit %
        s: running / stat.days.length // Sales %
      });
    }
    
    if (!cityCurves[stat.city]) cityCurves[stat.city] = [];
    cityCurves[stat.city].push(monthCurve);
  });

  // Durchschnittskurven bilden (in 10%-Schritten)
  Object.keys(cityCurves).forEach(city => {
    const allMonths = cityCurves[city];
    finalPatterns[city] = {};
    
    for (let i = 1; i <= 10; i++) {
      const targetT = i / 10;
      let sumS = 0;
      allMonths.forEach(curve => {
        // Finde den Punkt, der am nächsten an targetT ist
        const pt = curve.find(p => p.t >= targetT) || { s: 1 };
        sumS += pt.s;
      });
      finalPatterns[city][i * 10] = sumS / allMonths.length;
    }
  });

  // Globalen Durchschnitt berechnen (für neue Städte ohne Historie)
  const globalSamples = {};
  let globalCount = 0;
  Object.values(cityCurves).forEach(months => {
     months.forEach(curve => {
        for (let i = 1; i <= 10; i++) {
           const pt = curve.find(p => p.t >= (i/10)) || {s:1};
           globalSamples[i] = (globalSamples[i]||0) + pt.s;
        }
        globalCount++;
     });
  });
  finalPatterns['_GLOBAL_'] = {};
  for(let i=1; i<=10; i++) {
     finalPatterns['_GLOBAL_'][i*10] = globalSamples[i] / globalCount;
  }

  saveToCache(CONFIG_ENGINE.cacheKeyPatterns, finalPatterns, 21600); // 6h Cache
  Logger.log('🧠 Patterns gelernt: ' + Object.keys(finalPatterns).join(', '));
  return finalPatterns;
}

/**
 * 3. ENSEMBLE FORECAST (LOGIC)
 * Mischt Linear, Pattern und Ziel basierend auf Fortschritt.
 */
function calculateEnsembleForecast(ist, ziel, elapsedWD, totalWD, pattern) {
  // Edge Case: Monat noch nicht gestartet
  if (elapsedWD <= 0 || totalWD <= 0) return { forecast: (ziel > 0 ? ziel : 0), currentPace: 0 };

  const progress = elapsedWD / totalWD; // 0.0 bis 1.0
  const linearPace = ist / elapsedWD;
  
  // 1. Linear Forecast (Naive Extrapolation)
  const forecastLinear = ist + (linearPace * (totalWD - elapsedWD));

  // 2. Pattern Forecast (Seasonality)
  let forecastPattern = forecastLinear; // Fallback
  if (pattern) {
     // Finde erwarteten %-Satz für aktuellen Fortschritt
     // Wir runden auf den nächsten 10er Schritt auf (konservativ)
     const step = Math.ceil(progress * 10) * 10;
     const expectedShare = pattern[step] || progress;
     
     // Wenn wir erst 2% des Umsatzes erwarten, ist eine Hochrechnung gefährlich (Teilen durch fast Null).
     // Wir nutzen Pattern erst ab >5% erwartetem Umsatz.
     if (expectedShare > 0.05) {
        forecastPattern = ist / expectedShare;
     }
  }

  // 3. Goal Anchor (Stability)
  const forecastGoal = ziel > 0 ? ziel : forecastLinear;

  // --- GEWICHTUNG (THE MIX) ---
  let wLinear = 0;
  let wPattern = 0;
  let wGoal = 0;

  if (progress < 0.15) { 
    // START (Tag 1-3): Wir wissen nichts. Vertrauen auf Ziel & Historie.
    wGoal = 0.60;    // 60% Ziel
    wPattern = 0.30; // 30% Muster
    wLinear = 0.10;  // 10% Ist (zu volatil)
  } else if (progress < 0.80) {
    // MITTE: Das Muster ist jetzt König. Linear holt auf.
    wGoal = 0.10;
    wPattern = 0.50;
    wLinear = 0.40;
  } else {
    // ENDE: Realität zählt.
    wGoal = 0.0;
    wPattern = 0.20; // Muster hilft noch beim Endspurt-Faktor
    wLinear = 0.80;
  }

  // Fallback: Wenn kein Pattern da ist, verteile Pattern-Gewicht auf Linear
  if (!pattern) {
     wLinear += wPattern;
     wPattern = 0;
  }

  let finalForecast = (forecastLinear * wLinear) + 
                      (forecastPattern * wPattern) + 
                      (forecastGoal * wGoal);

  // Sanity Check: Forecast sollte nicht kleiner sein als das, was wir schon haben
  finalForecast = Math.max(ist, finalForecast);

  return { 
    forecast: finalForecast, 
    currentPace: linearPace 
  };
}

/**
 * HELPER: Holt Pattern-% für Soll-To-Date
 */
function getPatternPercent(elapsed, total, pattern) {
  if (total <= 0) return 0;
  const pct = elapsed / total;
  if (!pattern) return pct; // Linear fallback
  const step = Math.ceil(pct * 10) * 10;
  return pattern[step] || pct;
}

// --- UTIL HELPER ---
function getCache() { return CacheService.getScriptCache(); }
function getFromCache(key) {
  try { return JSON.parse(getCache().get(key)); } catch(e) { return null; }
}
function saveToCache(key, data, time) {
  try { getCache().put(key, JSON.stringify(data), time); } catch(e) {}
}
function cleanNum(val) {
  if (!val) return 0;
  if (typeof val === 'number') return Number.isFinite(val) ? val : 0;
  const s = String(val).replace(',', '.').replace(/[^0-9.-]/g, '');
  return parseFloat(s) || 0;
}
function parseDate(dateInput) {
  if (dateInput instanceof Date) return dateInput;
  if (typeof dateInput === 'string') {
    // Format: "30.12.2025" -> Split Datum
    const parts = dateInput.split('.');
    if (parts.length === 3) {
       // Monat ist 0-basiert in JS Date (0=Jan)
       return new Date(parts[2], parts[1]-1, parts[0]);
    }
  }
  return null;
}


function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
function doGet(e) {
  // ✅ AUTO-INIT: Feiertage-Sheets prüfen & erstellen
  try {
    initializeHolidaySheets();
  } catch (error) {
    Logger.log('⚠️ Feiertage-Init Fehler:', error.toString());
  }
  
  // Dashboard laden
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Trisor Live Sync')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
/**
 * HELPER: Gibt nur Aktivierungsdaten zurück
 */
function getActivationsData() {
  try {
    const frontend = loadDataForFrontend();
    return frontend.data || {};
  } catch (error) {
    Logger.log('Fehler getActivationsData: ' + error.toString());
    return {};
  }
}

/**
 * HELPER: Gibt nur Zieldaten zurück
 */
function getGoalsData() {
  try {
    const frontend = loadDataForFrontend();
    return frontend.goals || {};
  } catch (error) {
    Logger.log('Fehler getGoalsData: ' + error.toString());
    return {};
  }
}

/**
 * Speichert alle 12 Monatsziele eines Jahres für einen Standort in einem Rutsch.
 * Ersetzt die Notwendigkeit für 12 Einzelaufrufe.
 */
function setGoalsBatch(standort, year, goalsObject) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(CONFIG.SHEETS.ZIELE);
    
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.SHEETS.ZIELE);
      sheet.getRange(1, 1).setValue('Standort').setNumberFormat('@');
    }
    
    // 1. Header und Spalten-Mapping vorbereiten
    let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn() || 1).getDisplayValues()[0];
    const monthColumns = {}; // Map: "YYYY-MM" -> SpaltenIndex
    
    for (let m = 1; m <= 12; m++) {
      const monatKey = `${year}-${String(m).padStart(2, '0')}`;
      let colIndex = headers.indexOf(monatKey) + 1;
      
      if (colIndex === 0) { // Spalte existiert noch nicht
        colIndex = sheet.getLastColumn() + 1;
        sheet.getRange(1, colIndex).setNumberFormat('@').setValue(monatKey);
        headers.push(monatKey); // Header-Cache aktualisieren
      }
      monthColumns[m] = colIndex;
    }
    
    // 2. Zeile für den Standort finden oder erstellen
    const lastRow = sheet.getLastRow();
    let rowIndex = -1;
    if (lastRow > 1) {
      const standorte = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < standorte.length; i++) {
        if (String(standorte[i][0]).trim() === standort) {
          rowIndex = i + 2;
          break;
        }
      }
    }
    
    if (rowIndex === -1) {
      // Neuer Standort: Zeile mit Namen anlegen
      sheet.appendRow([standort]);
      rowIndex = sheet.getLastRow();
    }
    
    // 3. Alle 12 Werte schreiben
    // Wir nutzen hier Einzel-SetValues nur für die Ziel-Spalten, um andere Daten nicht zu überschreiben
    for (let m = 1; m <= 12; m++) {
      const targetCol = monthColumns[m];
      const val = Number(goalsObject[m]) || 0;
      sheet.getRange(rowIndex, targetCol).setValue(val);
    }
    
    Logger.log(`✅ Batch-Update erfolgreich: ${standort} (${year})`);
    return { success: true, message: `Ziele für ${standort} (${year}) gespeichert.` };
    
  } catch (error) {
    Logger.log('❌ setGoalsBatch Fehler: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}
