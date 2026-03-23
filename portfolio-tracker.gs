// ── Sheet names ────────────────────────────────────────────────
var SHEETS = {
  PORTFOLIO  : 'Portfolio',
  DASHBOARD  : 'Dashboard',
  HISTORY    : 'Daily History',
  WEEKLY     : 'Weekly Summary',
  MONTHLY    : 'Monthly Summary',
  WATCHLIST  : 'Watchlist',
  CHARTS     : 'Charts',
  SECTOR_MAP : '_SectorData'
};

// ── Column index constants (1 = col A) ─────────────────────────
var COL = {
  NAME      : 1,   // A
  SYMBOL    : 2,   // B
  SHARES    : 3,   // C
  BUY_PRICE : 4,   // D
  BUY_DATE  : 5,   // E
  CUR_PRICE : 6,   // F
  CUR_VALUE : 7,   // G
  INVESTED  : 8,   // H
  PNL       : 9,   // I
  PNL_PCT   : 10,  // J
  CAGR      : 11,  // K
  DAY_CHG   : 12,  // L
  DAY_CHG_R : 13,  // M
  W52_HIGH  : 14,  // N
  W52_LOW   : 15,  // O
  FROM_HIGH : 16,  // P
  SECTOR    : 17,  // Q
  TARGET    : 18,  // R
  STOP_LOSS : 19,  // S
  TO_TARGET : 20,  // T
  STATUS    : 21,  // U
};

// ── Design constants ───────────────────────────────────────────
var COLORS = {
  HEADER_BG   : '#1565C0',
  HEADER_FG   : '#FFFFFF',
  PROFIT      : '#2E7D32',
  LOSS        : '#C62828',
  WARN        : '#E65100',
  POS_BG      : '#E8F5E9',
  NEG_BG      : '#FFEBEE',
  WARN_BG     : '#FFF8E1',
  EVEN_ROW    : '#E3F2FD',
  CARD_BORDER : '#BBDEFB',
  CARD_LABEL  : '#F5F5F5',
  MUTED       : '#9E9E9E',
  SUBHEADING  : '#757575',
};

var FORMATS = {
  CURRENCY : '₹#,##0.00',
  PERCENT  : '0.00"%"',
  DATE     : 'DD MMM YYYY',
  DATE_IN  : 'DD/MM/YYYY',
  LARGE_NUM: '#,##0.00',
};

// Alert type → display style mapping
var ALERT_STYLES = {
  'TARGET HIT'    : { bg: COLORS.POS_BG,  fc: COLORS.PROFIT },
  'STOP LOSS HIT' : { bg: COLORS.NEG_BG,  fc: COLORS.LOSS   },
  'NEAR STOP LOSS': { bg: COLORS.WARN_BG, fc: COLORS.WARN   },
};

// ── Sample stocks ──────────────────────────────────────────────
var SAMPLE_STOCKS = [
  ['Tata Steel v2 - MO', 'TATASTEEL', 130,  120.13,,  '01/01/2025'],
  // Add your stocks as above mentioned format
];


// ================================================================
//  SETUP — Run once only
// ================================================================
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone('Asia/Kolkata');

  _buildSectorMapSheet(ss);
  _buildPortfolioSheet(ss);
  _buildDashboardSheet(ss);
  _buildHistorySheet(ss);
  _buildWeeklySheet(ss);
  _buildMonthlySheet(ss);
  _buildWatchlistSheet(ss);
  _buildChartsSheet(ss);

  _clearTriggers();
  ScriptApp.newTrigger('runDailySnapshot').timeBased().atHour(16).everyDays(1).create();
  ScriptApp.newTrigger('checkAlerts').timeBased().atHour(15).everyDays(1).create();

  Utilities.sleep(7000);
  runDailySnapshot();

  try {
    SpreadsheetApp.getUi().alert(
      'Setup complete!\n\n' +
      'Go to the "Portfolio" sheet and fill only columns A to E:\n' +
      '  A = Stock Name\n' +
      '  B = NSE Symbol  (e.g.  INFY, TATAMOTORS, HDFCBANK)\n' +
      '  C = Shares owned\n' +
      '  D = Avg Buy Price in ₹\n' +
      '  E = Buy Date in DD/MM/YYYY format\n\n' +
      'Everything else fills automatically.\n\n' +
      'Alerts run daily at 3:35 PM IST.\n' +
      'Snapshots run daily at 4:00 PM IST.'
    );
  } catch (e) {
    // getUi() unavailable when called from a trigger context — skip alert
    Logger.log('Setup complete (UI alert skipped: not an interactive session).');
  }
}


// ================================================================
//  DAILY SNAPSHOT — auto-triggered at 4 PM IST
// ================================================================
function runDailySnapshot() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var now = new Date();
  if (now.getDay() === 0 || now.getDay() === 6) { Logger.log('Weekend — skipping.'); return; }

  SpreadsheetApp.flush();
  Utilities.sleep(5000);

  _saveHistoryRow(ss, now);
  _rebuildPeriodSummary(ss, SHEETS.WEEKLY,  _weekKey);
  _rebuildPeriodSummary(ss, SHEETS.MONTHLY, _monthKey);
  _updateHealthScore(ss);
  _refreshCharts(ss);
  _applyDashboardColors(ss);

  Logger.log('Snapshot: ' + Utilities.formatDate(now, 'Asia/Kolkata', 'dd MMM yyyy HH:mm'));
}


// ================================================================
//  EMAIL ALERTS — auto at 3:35 PM IST
// ================================================================
function checkAlerts() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.PORTFOLIO);
  if (!sheet) return;

  var day = new Date().getDay();
  if (day === 0 || day === 6) return;

  SpreadsheetApp.flush();
  Utilities.sleep(4000);

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var data      = sheet.getRange(2, 1, lastRow - 1, COL.STATUS).getValues();
  var props     = PropertiesService.getScriptProperties();
  var allProps  = props.getProperties(); // load all at once — avoids N+1 lookups
  var todayKey  = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyyMMdd');
  var alerts    = [];
  var newProps  = {};

  data.forEach(function(row) {
    var name     = row[COL.NAME      - 1] || '';
    var symbol   = row[COL.SYMBOL    - 1] || '';
    var curPrice = parseFloat(row[COL.CUR_PRICE - 1]) || 0;
    var target   = parseFloat(row[COL.TARGET    - 1]) || 0;
    var sl       = parseFloat(row[COL.STOP_LOSS - 1]) || 0;
    var pnlPct   = parseFloat(row[COL.PNL_PCT   - 1]) || 0;
    var dayChg   = parseFloat(row[COL.DAY_CHG   - 1]) || 0;
    if (!symbol || curPrice === 0) return;

    var tKey = 'ALERT_' + symbol + '_T_'  + todayKey;
    var sKey = 'ALERT_' + symbol + '_SL_' + todayKey;
    var nKey = 'ALERT_' + symbol + '_NS_' + todayKey;

    if (target > 0 && curPrice >= target && !allProps[tKey]) {
      alerts.push({ name:name, symbol:symbol, type:'TARGET HIT',     curPrice:curPrice, level:target, pnlPct:pnlPct, dayChg:dayChg });
      newProps[tKey] = 'sent';
    }
    if (sl > 0 && curPrice <= sl && !allProps[sKey]) {
      alerts.push({ name:name, symbol:symbol, type:'STOP LOSS HIT',  curPrice:curPrice, level:sl,     pnlPct:pnlPct, dayChg:dayChg });
      newProps[sKey] = 'sent';
    }
    if (sl > 0 && curPrice > sl && (curPrice - sl) / curPrice < 0.03 && !allProps[nKey]) {
      alerts.push({ name:name, symbol:symbol, type:'NEAR STOP LOSS', curPrice:curPrice, level:sl,     pnlPct:pnlPct, dayChg:dayChg });
      newProps[nKey] = 'sent';
    }
  });

  if (Object.keys(newProps).length) props.setProperties(newProps); // one batch write

  if (alerts.length === 0) { Logger.log('No alerts.'); return; }

  var subject = 'Portfolio Alert — ' + alerts.length + ' notification(s) — ' +
    Utilities.formatDate(new Date(), 'Asia/Kolkata', 'dd MMM yyyy');

  var rows = alerts.map(function(a) {
    var style = ALERT_STYLES[a.type] || { bg: '#FFFFFF', fc: '#000000' };
    return '<tr style="background:' + style.bg + '">' +
      '<td>' + a.name + '</td><td><b>' + a.symbol + '</b></td>' +
      '<td style="color:' + style.fc + ';font-weight:bold">' + a.type + '</td>' +
      '<td>₹' + a.curPrice.toFixed(2) + '</td><td>₹' + a.level.toFixed(2) + '</td>' +
      '<td style="color:' + (a.pnlPct >= 0 ? COLORS.PROFIT : COLORS.LOSS) + '">' + a.pnlPct.toFixed(2) + '%</td>' +
      '<td style="color:' + (a.dayChg >= 0 ? COLORS.PROFIT : COLORS.LOSS) + '">' + a.dayChg.toFixed(2) + '%</td>' +
      '</tr>';
  }).join('');

  var body = '<div style="font-family:Arial,sans-serif">' +
    '<h2 style="color:' + COLORS.HEADER_BG + ';margin-bottom:4px">Portfolio Alerts</h2>' +
    '<p style="color:' + COLORS.MUTED + '">Triggered at ' +
    Utilities.formatDate(new Date(), 'Asia/Kolkata', 'hh:mm a') + ' IST</p>' +
    '<table border="1" cellpadding="8" cellspacing="0" style="border-collapse:collapse;font-size:13px">' +
    '<tr style="background:' + COLORS.HEADER_BG + ';color:' + COLORS.HEADER_FG + '">' +
    '<th>Stock</th><th>Symbol</th><th>Alert</th><th>Current ₹</th><th>Level ₹</th><th>P&L %</th><th>Day %</th></tr>' +
    rows + '</table>' +
    '<br><p style="font-size:11px;color:' + COLORS.MUTED + '">Sent by your Portfolio Tracker</p></div>';

  MailApp.sendEmail({ to: Session.getActiveUser().getEmail(), subject: subject, htmlBody: body });
  Logger.log('Alert email sent: ' + alerts.length + ' alert(s).');
}


// ================================================================
//  PORTFOLIO SHEET  (Columns A–U, 21 columns)
// ================================================================
function _buildPortfolioSheet(ss) {
  var sheet = ss.getSheetByName(SHEETS.PORTFOLIO);

  // Preserve existing user data before rebuild
  var saved = [];
  if (sheet && sheet.getLastRow() > 1) {
    var lr  = sheet.getLastRow();
    var lc  = sheet.getLastColumn();
    var raw = sheet.getRange(2, 1, lr - 1, Math.max(lc, COL.STOP_LOSS)).getValues();
    raw.forEach(function(r) {
      if (r[COL.SYMBOL - 1]) saved.push({
        name     : r[COL.NAME      - 1],
        symbol   : r[COL.SYMBOL    - 1],
        shares   : r[COL.SHARES    - 1],
        buyPrice : r[COL.BUY_PRICE - 1],
        buyDate  : r[COL.BUY_DATE  - 1] || new Date(),
        target   : lc >= COL.TARGET    ? r[COL.TARGET    - 1] : 0,
        sl       : lc >= COL.STOP_LOSS ? r[COL.STOP_LOSS - 1] : 0,
      });
    });
    sheet.clear();
    sheet.clearFormats();
  } else {
    sheet = sheet || ss.insertSheet(SHEETS.PORTFOLIO, 0);
    sheet.clear();
    sheet.clearFormats();
  }

  var headers = [
    'Stock Name','NSE Symbol','Shares','Avg Buy Price ₹','Buy Date',
    'Current Price ₹','Current Value ₹','Invested ₹',
    'P&L ₹','P&L %','CAGR %',
    'Day Change %','Day Change ₹',
    '52W High ₹','52W Low ₹','% from 52W High',
    'Sector','Target Price ₹','Stop Loss ₹','To Target %','Status'
  ];
  _applyHeaderStyle(sheet, headers,
    [175,100,65,130,110,130,130,120,100,80,80,110,120,110,110,120,110,120,110,100,140]);

  var fill = saved.length > 0 ? saved : SAMPLE_STOCKS.map(function(s) {
    return { name:s[0], symbol:s[1], shares:s[2], buyPrice:s[3], buyDate:s[4], target:0, sl:0 };
  });

  // Batch user-input columns (A–E) in one setValues call
  var userRows  = [];
  var tslRows   = [];  // target + stop loss (R, S)
  var formulaRows = [];

  fill.forEach(function(s, i) {
    var r      = i + 2;
    var price  = parseFloat(s.buyPrice) || 0;
    var tgt    = parseFloat(s.target) > 0 ? s.target : (price * 1.20);
    var sl     = parseFloat(s.sl)     > 0 ? s.sl     : (price * 0.90);
    var buyDt  = _parseBuyDate(s.buyDate);

    userRows.push([s.name, s.symbol, s.shares, price, buyDt]);
    tslRows.push([tgt, sl]);
    formulaRows.push(_buildPortfolioFormulaRow(r));
  });

  var n = fill.length;
  sheet.getRange(2, COL.NAME,      n, 5).setValues(userRows);
  sheet.getRange(2, COL.TARGET,    n, 2).setValues(tslRows);
  // Batch all 15 formula columns in one call (cols F–T, skipping R+S which are values)
  // Cols F–Q (6–17) = 12 cols, then T–U (20–21) = 2 cols; set each contiguous range
  var formulaColsLeft  = formulaRows.map(function(r) { return r.slice(0, 12); }); // F–Q
  var formulaColsRight = formulaRows.map(function(r) { return r.slice(12);    }); // T–U
  sheet.getRange(2, COL.CUR_PRICE, n, 12).setFormulas(formulaColsLeft);
  sheet.getRange(2, COL.TO_TARGET, n,  2).setFormulas(formulaColsRight);

  // Formats
  sheet.getRange('D2:D100').setNumberFormat(FORMATS.CURRENCY);
  sheet.getRange('E2:E100').setNumberFormat(FORMATS.DATE_IN);
  sheet.getRange('F2:I100').setNumberFormat(FORMATS.CURRENCY);
  sheet.getRange('J2:L100').setNumberFormat(FORMATS.PERCENT);
  sheet.getRange('M2:M100').setNumberFormat(FORMATS.CURRENCY);
  sheet.getRange('N2:O100').setNumberFormat(FORMATS.CURRENCY);
  sheet.getRange('P2:P100').setNumberFormat(FORMATS.PERCENT);
  sheet.getRange('R2:S100').setNumberFormat(FORMATS.CURRENCY);
  sheet.getRange('T2:T100').setNumberFormat(FORMATS.PERCENT);

  sheet.setFrozenColumns(2);

  // Alternating rows — batch via getRangeList instead of 99 individual setBackground calls
  var evenRanges = []; var oddRanges = [];
  for (var j = 2; j <= 100; j++) {
    (j % 2 === 0 ? evenRanges : oddRanges).push('A' + j + ':U' + j);
  }
  sheet.getRangeList(evenRanges).setBackground(COLORS.EVEN_ROW);
  sheet.getRangeList(oddRanges).setBackground('#FFFFFF');

  sheet.getRange('A1').setNote(
    'Fill ONLY columns A to E with your stock info.\n' +
    'Columns F onwards are all automatic.\n\n' +
    'Col R = Target Price (default: buy price +20%)\n' +
    'Col S = Stop Loss   (default: buy price -10%)\n' +
    'You can type any value in R or S to override.\n\n' +
    'Status (col U): TARGET HIT / SL HIT / NEAR SL / NEAR TARGET / HOLD'
  );

  Logger.log('Portfolio sheet ready — ' + n + ' stocks.');
}

// Returns the 14-element formula array for a single portfolio row (cols F–Q, T–U)
function _buildPortfolioFormulaRow(r) {
  var b = 'B'+r; var c = 'C'+r; var d = 'D'+r; var e = 'E'+r;
  var f = 'F'+r; var g = 'G'+r; var h = 'H'+r;
  var ri = 'R'+r; var s = 'S'+r;
  return [
    // F–Q (12 formulas)
    '=IFERROR(GOOGLEFINANCE("NSE:"&' + b + ',"price"),0)',
    '=IFERROR(' + c + '*' + f + ',0)',
    '=IF(' + c + '="",0,' + c + '*' + d + ')',
    '=IFERROR(' + g + '-' + h + ',0)',
    '=IFERROR((' + g + '/' + h + '-1)*100,0)',
    '=IFERROR(((' + f + '/' + d + ')^(365/MAX(TODAY()-' + e + ',1))-1)*100,0)',
    '=IFERROR(GOOGLEFINANCE("NSE:"&' + b + ',"changepct"),0)',
    '=IFERROR(' + f + '*(L' + r + '/100)*' + c + ',0)',
    '=IFERROR(GOOGLEFINANCE("NSE:"&' + b + ',"high52"),0)',
    '=IFERROR(GOOGLEFINANCE("NSE:"&' + b + ',"low52"),0)',
    '=IFERROR((' + f + '/N' + r + '-1)*100,0)',
    '=IFERROR(VLOOKUP(UPPER(' + b + '),_SectorData!A:B,2,0),"Other")',
    // T–U (2 formulas)
    '=IFERROR((' + ri + '/' + f + '-1)*100,0)',
    '=IF(' + f + '=0,"",IF(' + f + '>=' + ri + ',"TARGET HIT",IF(' + f + '<=' + s + ',"SL HIT",' +
      'IF((' + f + '-' + s + ')/' + f + '<0.03,"NEAR SL",' +
      'IF((' + ri + '-' + f + ')/' + f + '<0.05,"NEAR TARGET","HOLD")))))',
  ];
}

function _parseBuyDate(val) {
  if (val instanceof Date) return val;
  try {
    var p = String(val).split('/');
    return new Date(parseInt(p[2]), parseInt(p[1]) - 1, parseInt(p[0]));
  } catch(e) { return new Date(); }
}


// ================================================================
//  SECTOR MAP — hidden helper sheet
// ================================================================
function _buildSectorMapSheet(ss) {
  if (ss.getSheetByName(SHEETS.SECTOR_MAP)) return;
  var sheet = ss.insertSheet(SHEETS.SECTOR_MAP);
  var map = [
    ['TCS','IT'],['INFY','IT'],['WIPRO','IT'],['HCLTECH','IT'],['TECHM','IT'],
    ['LTIM','IT'],['MPHASIS','IT'],['COFORGE','IT'],['PERSISTENT','IT'],['OFSS','IT'],
    ['NAUKRI','IT Services'],['ZOMATO','Consumer Tech'],['PAYTM','Fintech'],
    ['HDFCBANK','Banking'],['ICICIBANK','Banking'],['SBIN','Banking'],
    ['KOTAKBANK','Banking'],['AXISBANK','Banking'],['INDUSINDBK','Banking'],
    ['BANKBARODA','Banking'],['PNB','Banking'],['CANBK','Banking'],
    ['UNIONBANK','Banking'],['IDFCFIRSTB','Banking'],['FEDERALBNK','Banking'],
    ['YESBANK','Banking'],['RBLBANK','Banking'],
    ['BAJFINANCE','NBFC'],['BAJAJFINSV','NBFC'],['MUTHOOTFIN','NBFC'],
    ['CHOLAFIN','NBFC'],['PFC','Finance'],['RECLTD','Finance'],
    ['IRFC','Finance'],['M&MFIN','NBFC'],
    ['HDFCLIFE','Insurance'],['SBILIFE','Insurance'],['ICICIPRULI','Insurance'],
    ['LIC','Insurance'],['POLICYBZR','Insurance'],
    ['RELIANCE','Energy'],['ONGC','Energy'],['NTPC','Energy'],
    ['POWERGRID','Energy'],['GAIL','Energy'],['BPCL','Energy'],['IOC','Energy'],
    ['HINDPETRO','Energy'],['TATAPOWER','Energy'],['ADANIGREEN','Energy'],
    ['TORNTPOWER','Energy'],['ADANIPOWER','Energy'],['CESC','Energy'],
    ['HAL','Defence'],['BEL','Defence'],['BHEL','Defence'],['COCHINSHIP','Defence'],
    ['GRSE','Defence'],['MAZAGON','Defence'],['BEML','Defence'],
    ['APOLLOHOSP','Healthcare'],['FORTIS','Healthcare'],['MAXHEALTH','Healthcare'],
    ['SUNPHARMA','Pharma'],['DRREDDY','Pharma'],['CIPLA','Pharma'],
    ['DIVISLAB','Pharma'],['AUROPHARMA','Pharma'],['LUPIN','Pharma'],
    ['BIOCON','Pharma'],['TORNTPHARM','Pharma'],['ALKEM','Pharma'],
    ['MARUTI','Auto'],['TATAMOTORS','Auto'],['M&M','Auto'],
    ['BAJAJ-AUTO','Auto'],['HEROMOTOCO','Auto'],['EICHERMOT','Auto'],
    ['ASHOKLEY','Auto'],['TVSMOTORS','Auto'],
    ['APOLLOTYRE','Auto Ancillary'],['MRF','Auto Ancillary'],
    ['BOSCHLTD','Auto Ancillary'],['MOTHERSON','Auto Ancillary'],
    ['TATASTEEL','Metals'],['JSWSTEEL','Metals'],['HINDALCO','Metals'],
    ['VEDL','Metals'],['COALINDIA','Mining'],['NMDC','Mining'],['SAIL','Metals'],
    ['ULTRACEMCO','Cement'],['SHREECEM','Cement'],['ACC','Cement'],
    ['AMBUJACEM','Cement'],['RAMCOCEM','Cement'],
    ['HINDUNILVR','FMCG'],['ITC','FMCG'],['NESTLEIND','FMCG'],
    ['BRITANNIA','FMCG'],['DABUR','FMCG'],['MARICO','FMCG'],
    ['GODREJCP','FMCG'],['COLPAL','FMCG'],['EMAMILTD','FMCG'],
    ['TITAN','Consumer'],['DMART','Retail'],['TRENT','Retail'],
    ['ASIANPAINT','Consumer'],['BERGER','Consumer'],
    ['CERA','Consumer Goods'],['KAJARIACER','Consumer Goods'],
    ['PIDILITIND','Chemicals'],['SRF','Chemicals'],['DEEPAKNTR','Chemicals'],
    ['UPL','Agro Chemicals'],['PI','Agro Chemicals'],
    ['LT','Infrastructure'],['ABB','Capital Goods'],['SIEMENS','Capital Goods'],
    ['HAVELLS','Capital Goods'],['VOLTAS','Capital Goods'],
    ['ADANIENT','Conglomerate'],['ADANIPORTS','Logistics'],['CONCOR','Logistics'],
    ['DLF','Real Estate'],['GODREJPROP','Real Estate'],['OBEROIRLTY','Real Estate'],
    ['PRESTIGE','Real Estate'],['PHOENIXLTD','Real Estate'],
    ['BHARTIARTL','Telecom'],['IDEA','Telecom'],
    ['IRCTC','Tourism'],['WHIRLPOOL','Consumer Durables'],['NYKAA','Consumer Tech'],
  ];
  sheet.getRange(1, 1, map.length, 2).setValues(map);
  sheet.hideSheet();
  Logger.log('Sector map: ' + map.length + ' stocks.');
}


// ================================================================
//  DASHBOARD SHEET
// ================================================================
function _buildDashboardSheet(ss) {
  var sheet = ss.getSheetByName(SHEETS.DASHBOARD) || ss.insertSheet(SHEETS.DASHBOARD, 1);
  sheet.clear();
  sheet.clearFormats();
  sheet.setColumnWidth(1, 20);

  sheet.getRange('B1').setValue('PORTFOLIO DASHBOARD')
    .setFontSize(22).setFontWeight('bold').setFontColor(COLORS.HEADER_BG);
  sheet.getRange('B2')
    .setFormula('="Last updated: "&TEXT(NOW(),"DD MMM YYYY HH:MM")&" IST"')
    .setFontColor(COLORS.MUTED).setFontSize(10);

  sheet.getRange('B3').setValue('MARKET OVERVIEW')
    .setFontSize(11).setFontWeight('bold').setFontColor(COLORS.SUBHEADING);

  _marketCard(sheet, 'B', 'NIFTY 50',
    '=IFERROR(GOOGLEFINANCE("INDEXNSE:NIFTY_50","price"),"-")',
    '=IFERROR(TEXT(GOOGLEFINANCE("INDEXNSE:NIFTY_50","changepct"),"0.00")&"%","-")');
  _marketCard(sheet, 'D', 'SENSEX',
    '=IFERROR(GOOGLEFINANCE("INDEXBOM:SENSEX","price"),"-")',
    '=IFERROR(TEXT(GOOGLEFINANCE("INDEXBOM:SENSEX","changepct"),"0.00")&"%","-")');
  _marketCard(sheet, 'F', 'NIFTY BANK',
    '=IFERROR(GOOGLEFINANCE("INDEXNSE:NIFTY_BANK","price"),"-")',
    '=IFERROR(TEXT(GOOGLEFINANCE("INDEXNSE:NIFTY_BANK","changepct"),"0.00")&"%","-")');
  _marketCard(sheet, 'H', 'NIFTY IT',
    '=IFERROR(GOOGLEFINANCE("INDEXNSE:NIFTY_IT","price"),"-")',
    '=IFERROR(TEXT(GOOGLEFINANCE("INDEXNSE:NIFTY_IT","changepct"),"0.00")&"%","-")');

  sheet.getRange('B7').setValue('PORTFOLIO OVERVIEW')
    .setFontSize(11).setFontWeight('bold').setFontColor(COLORS.SUBHEADING);

  _statCard(sheet, 8, 'B', 'NET WORTH ₹',
    '=SUMPRODUCT(IFERROR(Portfolio!C2:C100,0)*IFERROR(IF(IFERROR(Portfolio!F2:F100,0)=0,Portfolio!D2:D100,Portfolio!F2:F100),0))');
  _statCard(sheet, 8, 'E', 'TOTAL INVESTED ₹',
    '=SUMPRODUCT(IFERROR(Portfolio!C2:C100,0)*IFERROR(Portfolio!D2:D100,0))');
  _statCard(sheet, 8, 'H', 'TOTAL P&L ₹', '=B9-E9');
  _statCard(sheet, 8, 'K', 'P&L %', '=IFERROR((H9/E9)*100,0)', FORMATS.PERCENT);

  sheet.getRange('B11').setValue('TODAY')
    .setFontSize(11).setFontWeight('bold').setFontColor(COLORS.SUBHEADING);

  _statCard(sheet, 12, 'B', "TODAY'S CHANGE ₹",
    '=SUMPRODUCT(IFERROR(Portfolio!M2:M100,0))');
  _statCard(sheet, 12, 'E', "TODAY'S CHANGE %",
    '=IFERROR((B13/E9)*100,0)', FORMATS.PERCENT);
  _statCard(sheet, 12, 'H', 'BEST TODAY',
    '=IFERROR(INDEX(Portfolio!A2:A100,MATCH(MAX(Portfolio!L2:L100),Portfolio!L2:L100,0))&" +"&TEXT(MAX(Portfolio!L2:L100),"0.00")&"%","-")', '@');
  _statCard(sheet, 12, 'K', 'WORST TODAY',
    '=IFERROR(INDEX(Portfolio!A2:A100,MATCH(MIN(Portfolio!L2:L100),Portfolio!L2:L100,0))&" "&TEXT(MIN(Portfolio!L2:L100),"0.00")&"%","-")', '@');

  sheet.getRange('M7').setValue('HEALTH SCORE')
    .setFontSize(10).setFontWeight('bold').setFontColor(COLORS.SUBHEADING)
    .setBackground(COLORS.CARD_LABEL).setHorizontalAlignment('center');
  sheet.getRange('M8').setValue('Run snapshot')
    .setFontSize(18).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange('M9').setValue('to calculate')
    .setFontSize(9).setFontColor(COLORS.MUTED).setHorizontalAlignment('center');
  sheet.getRange('M7:M9').setBorder(true,true,true,true,false,false,
    COLORS.CARD_BORDER, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.setColumnWidth(13, 160);

  sheet.getRange('B15').setValue('HOLDINGS BREAKDOWN')
    .setFontSize(13).setFontWeight('bold').setFontColor(COLORS.HEADER_BG);

  var tH = ['Stock','Shares','Buy ₹','Current ₹','Value ₹','P&L ₹','P&L %',
            'CAGR %','Day %','52W High','From High','Sector','Target ₹','SL ₹','Status'];
  _applyHeaderStyle(
    sheet, tH,
    [165,120,120,120,120,120,120,120,120,120,120,120,120,120,120],
    16, 2 // startRow=16, startCol=2 (col B)
  );

  // Portfolio column letters for the holdings table
  var portCols = ['A','C','D','F','G','I','J','K','L','N','P','Q','R','S','U'];
  var formulaGrid = [];
  for (var i = 0; i < 3000; i++) {
    var pr = 2 + i;
    formulaGrid.push(portCols.map(function(c) {
      return '=IF(Portfolio!' + c + pr + '="","",Portfolio!' + c + pr + ')';
    }));
    sheet.getRange(17 + i, 2, 1, tH.length)
      .setBackground(i % 2 === 0 ? COLORS.EVEN_ROW : '#FFFFFF');
  }
  sheet.getRange(17, 2, 3000, portCols.length).setFormulas(formulaGrid);

  sheet.getRange('D17:H46').setNumberFormat(FORMATS.CURRENCY);
  sheet.getRange('I17:K46').setNumberFormat(FORMATS.PERCENT);
  sheet.getRange('M17:N46').setNumberFormat(FORMATS.CURRENCY);

  sheet.setFrozenRows(2);
  Logger.log('Dashboard sheet ready.');
}

function _marketCard(sheet, col, label, priceFormula, changeFormula) {
  sheet.getRange(col + '4').setValue(label)
    .setFontSize(9).setFontColor(COLORS.SUBHEADING).setFontWeight('bold')
    .setBackground(COLORS.CARD_LABEL).setHorizontalAlignment('center');
  sheet.getRange(col + '5').setFormula(priceFormula)
    .setFontSize(14).setFontWeight('bold')
    .setNumberFormat(FORMATS.LARGE_NUM).setBackground('#FFFFFF').setHorizontalAlignment('center');
  sheet.getRange(col + '6').setFormula(changeFormula)
    .setFontSize(11).setFontWeight('bold').setBackground('#FFFFFF').setHorizontalAlignment('center');
  sheet.getRange(col + '4:' + col + '6')
    .setBorder(true,true,true,true,false,false, COLORS.CARD_BORDER, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

function _statCard(sheet, startRow, col, label, formula, fmt) {
  sheet.getRange(col + startRow).setValue(label)
    .setFontSize(10).setFontColor(COLORS.SUBHEADING).setFontWeight('bold')
    .setBackground(COLORS.CARD_LABEL).setHorizontalAlignment('center');
  sheet.getRange(col + (startRow + 1)).setFormula(formula)
    .setFontSize(20).setFontWeight('bold')
    .setNumberFormat(fmt || FORMATS.CURRENCY).setBackground('#FFFFFF').setHorizontalAlignment('center');
  sheet.getRange(col + startRow + ':' + col + (startRow + 1))
    .setBorder(true,true,true,true,false,false, COLORS.CARD_BORDER, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

// Shared header styling helper used by all data sheets
function _applyHeaderStyle(sheet, headers, widths, startRow, startCol) {
  var r = startRow  || 1;
  var c = startCol  || 1;
  sheet.getRange(r, c, 1, headers.length).setValues([headers])
    .setBackground(COLORS.HEADER_BG).setFontColor(COLORS.HEADER_FG)
    .setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center');
  widths.forEach(function(w, i) { sheet.setColumnWidth(c + i, w); });
  sheet.setFrozenRows(r);
}


// ================================================================
//  PORTFOLIO HEALTH SCORE
// ================================================================
function _updateHealthScore(ss) {
  var portSheet = ss.getSheetByName(SHEETS.PORTFOLIO);
  var dash      = ss.getSheetByName(SHEETS.DASHBOARD);
  if (!portSheet || !dash) return;

  var lr = portSheet.getLastRow();
  if (lr < 2) return;

  var data = portSheet.getRange(2, 1, lr - 1, COL.STATUS).getValues();
  var totalStocks = 0; var inProfit = 0;
  var sectors = {}; var totalVal = 0; var maxVal = 0;

  data.forEach(function(row) {
    if (!row[COL.SYMBOL - 1]) return;
    totalStocks++;
    var pnl    = parseFloat(row[COL.PNL      - 1]) || 0;
    var val    = parseFloat(row[COL.CUR_VALUE - 1]) || 0;
    var sector = row[COL.SECTOR - 1] || 'Other';
    if (pnl >= 0) inProfit++;
    totalVal += val;
    if (val > maxVal) maxVal = val;
    sectors[sector] = (sectors[sector] || 0) + val;
  });

  if (totalStocks === 0) return;

  var numSectors    = Object.keys(sectors).length;
  var concentration = totalVal > 0 ? maxVal / totalVal : 1;
  var score = Math.round((inProfit / totalStocks) * 40) +
              Math.min(30, numSectors * 6) +
              Math.round((1 - Math.min(concentration, 1)) * 30);

  var label = score >= 80 ? 'Excellent' : score >= 60 ? 'Good' : score >= 40 ? 'Fair' : 'Needs Review';
  var color = score >= 80 ? COLORS.PROFIT : score >= 60 ? COLORS.HEADER_BG : score >= 40 ? COLORS.WARN : COLORS.LOSS;

  dash.getRange('M8').setValue(score + ' / 100').setFontColor(color).setFontSize(20);
  dash.getRange('M9').setValue(label + ' — ' + numSectors + ' sectors, ' + inProfit + '/' + totalStocks + ' profitable');
  Logger.log('Health: ' + score + '/100 (' + label + ')');
}


// ================================================================
//  DAILY HISTORY SHEET
// ================================================================
function _buildHistorySheet(ss) {
  var sheet = ss.getSheetByName(SHEETS.HISTORY);
  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.HISTORY, 2);
    _applyHeaderStyle(sheet,
      ['Date','Net Worth ₹','Invested ₹','Total P&L ₹','P&L %','Day Change ₹','Day Change %'],
      [120,150,130,130,80,150,120]);
    sheet.getRange('A2:A1000').setNumberFormat(FORMATS.DATE);
    sheet.getRange('B2:D1000').setNumberFormat(FORMATS.CURRENCY);
    sheet.getRange('E2:E1000').setNumberFormat(FORMATS.PERCENT);
    sheet.getRange('F2:F1000').setNumberFormat(FORMATS.CURRENCY);
    sheet.getRange('G2:G1000').setNumberFormat(FORMATS.PERCENT);
  }
  Logger.log('History sheet ready.');
}

function _saveHistoryRow(ss, now) {
  var portSheet = ss.getSheetByName(SHEETS.PORTFOLIO);
  var histSheet = ss.getSheetByName(SHEETS.HISTORY);
  if (!portSheet || !histSheet) return;

  var lr = portSheet.getLastRow();
  if (lr < 2) return;

  var data = portSheet.getRange(2, 1, lr - 1, COL.STATUS).getValues();
  var totalValue = 0; var totalInvested = 0;

  data.forEach(function(row) {
    var shares   = parseFloat(row[COL.SHARES    - 1]) || 0;
    var curPrice = parseFloat(row[COL.CUR_PRICE - 1]) || 0;
    var buyPrice = parseFloat(row[COL.BUY_PRICE - 1]) || 0;
    if (shares > 0) {
      totalValue    += shares * (curPrice > 0 ? curPrice : buyPrice);
      totalInvested += shares * buyPrice;
    }
  });

  var totalPnL  = totalValue - totalInvested;
  var pnlPct    = totalInvested > 0 ? (totalPnL / totalInvested) * 100 : 0;
  var histLR    = histSheet.getLastRow();
  var prevVal   = histLR > 1 ? (parseFloat(histSheet.getRange(histLR, 2).getValue()) || 0) : 0;
  var dayChange = totalValue - prevVal;
  var dayChgPct = prevVal > 0 ? (dayChange / prevVal) * 100 : 0;

  var todayStr = Utilities.formatDate(now, 'Asia/Kolkata', 'ddMMyyyy');
  if (histLR > 1) {
    var lastDV = histSheet.getRange(histLR, 1).getValue();
    if (lastDV instanceof Date && Utilities.formatDate(lastDV, 'Asia/Kolkata', 'ddMMyyyy') === todayStr) {
      histSheet.getRange(histLR, 2, 1, 6).setValues(
        [[totalValue, totalInvested, totalPnL, pnlPct, dayChange, dayChgPct]]);
      Logger.log('History updated (same day).');
      return;
    }
  }
  histSheet.appendRow([now, totalValue, totalInvested, totalPnL, pnlPct, dayChange, dayChgPct]);
  Logger.log('History row appended.');
}


// ================================================================
//  WEEKLY + MONTHLY SUMMARY — unified via _rebuildPeriodSummary
// ================================================================
function _weekKey(date) {
  var d = new Date(date);
  var day = d.getDay() || 7;
  d.setDate(d.getDate() - day + 1);
  return Utilities.formatDate(d, 'Asia/Kolkata', 'dd MMM yyyy');
}

function _monthKey(date) {
  return Utilities.formatDate(date, 'Asia/Kolkata', 'MMM yyyy');
}

function _buildWeeklySheet(ss) {
  if (ss.getSheetByName(SHEETS.WEEKLY)) { _applyPeriodFormats(ss.getSheetByName(SHEETS.WEEKLY)); return; }
  var sheet = ss.insertSheet(SHEETS.WEEKLY, 3);
  _applyHeaderStyle(sheet,
    ['Week Of (Mon)','Opening ₹','Closing ₹','Change ₹','Change %','Weekly High ₹','Weekly Low ₹'],
    [140,155,155,135,90,155,145]);
  _applyPeriodFormats(sheet);
  Logger.log('Weekly sheet ready.');
}

function _buildMonthlySheet(ss) {
  if (ss.getSheetByName(SHEETS.MONTHLY)) { _applyPeriodFormats(ss.getSheetByName(SHEETS.MONTHLY)); return; }
  var sheet = ss.insertSheet(SHEETS.MONTHLY, 4);
  _applyHeaderStyle(sheet,
    ['Month','Opening ₹','Closing ₹','Change ₹','Change %','Monthly High ₹','Monthly Low ₹'],
    [120,160,160,140,90,155,145]);
  _applyPeriodFormats(sheet);
  Logger.log('Monthly sheet ready.');
}

function _applyPeriodFormats(sheet) {
  sheet.getRange('B2:D1000').setNumberFormat(FORMATS.CURRENCY);
  sheet.getRange('E2:E1000').setNumberFormat(FORMATS.PERCENT);
  sheet.getRange('F2:G1000').setNumberFormat(FORMATS.CURRENCY);
}

function _rebuildPeriodSummary(ss, targetSheetName, keyFn) {
  var histSheet = ss.getSheetByName(SHEETS.HISTORY);
  var outSheet  = ss.getSheetByName(targetSheetName);
  if (!histSheet || !outSheet) return;

  var lastRow = histSheet.getLastRow();
  if (lastRow < 2) return;

  var histData = histSheet.getRange(2, 1, lastRow - 1, 2).getValues();
  var periods = {}; var order = [];

  histData.forEach(function(row) {
    var date = row[0]; var val = parseFloat(row[1]) || 0;
    if (!(date instanceof Date) || val === 0) return;
    var key = keyFn(date);
    if (!periods[key]) { periods[key] = { values:[] }; order.push(key); }
    periods[key].values.push(val);
  });

  if (!order.length) return;

  var rows = order.map(function(k) {
    var v = periods[k].values;
    var open = v[0]; var close = v[v.length - 1];
    var high = Math.max.apply(null, v); var low = Math.min.apply(null, v);
    var chg  = close - open; var pct = open > 0 ? (chg / open) * 100 : 0;
    return [k, open, close, chg, pct, high, low];
  });

  outSheet.getRange(2, 1, rows.length, 7).setValues(rows);
  var sl = outSheet.getLastRow();
  if (sl > rows.length + 1)
    outSheet.getRange(rows.length + 2, 1, sl - rows.length - 1, 7).clearContent();

  // Batch color application via getRangeList — avoids N*4 individual setBackground/setFontColor calls
  var posBgRanges = []; var negBgRanges = [];
  var posDRanges  = []; var negDRanges  = [];
  var posERanges  = []; var negERanges  = [];

  rows.forEach(function(row, i) {
    var rowAddr = (i + 2) + ':' + (i + 2);
    var dAddr   = 'D' + (i + 2); var eAddr = 'E' + (i + 2);
    (row[3] >= 0 ? posBgRanges : negBgRanges).push(rowAddr);
    (row[3] >= 0 ? posDRanges  : negDRanges ).push(dAddr);
    (row[4] >= 0 ? posERanges  : negERanges ).push(eAddr);
  });

  if (posBgRanges.length) outSheet.getRangeList(posBgRanges).setBackground(COLORS.POS_BG);
  if (negBgRanges.length) outSheet.getRangeList(negBgRanges).setBackground(COLORS.NEG_BG);
  if (posDRanges.length)  outSheet.getRangeList(posDRanges).setFontColor(COLORS.PROFIT);
  if (negDRanges.length)  outSheet.getRangeList(negDRanges).setFontColor(COLORS.LOSS);
  if (posERanges.length)  outSheet.getRangeList(posERanges).setFontColor(COLORS.PROFIT);
  if (negERanges.length)  outSheet.getRangeList(negERanges).setFontColor(COLORS.LOSS);

  Logger.log(targetSheetName + ': ' + rows.length + ' periods.');
}


// ================================================================
//  WATCHLIST SHEET
// ================================================================
function _buildWatchlistSheet(ss) {
  if (ss.getSheetByName(SHEETS.WATCHLIST)) return;
  var sheet = ss.insertSheet(SHEETS.WATCHLIST, 5);
  _applyHeaderStyle(sheet,
    ['Stock Name','NSE Symbol','Current Price ₹','Day Change %','Day Change ₹',
     '52W High ₹','52W Low ₹','% from 52W High','Sector','Your Notes'],
    [175,110,130,120,120,120,110,130,110,220]);

  var samples = [['Infosys','INFY'],['HDFC Bank','HDFCBANK'],
                 ['Bajaj Finance','BAJFINANCE'],['Maruti Suzuki','MARUTI']];
  var names = []; var formulas = [];
  samples.forEach(function(s, i) {
    names.push([s[0], s[1]]);
    formulas.push(_buildWatchlistFormulaRow(i + 2));
  });
  sheet.getRange(2, 1, samples.length, 2).setValues(names);
  sheet.getRange(2, 3, samples.length, 7).setFormulas(formulas);

  sheet.getRange('C2:C100').setNumberFormat(FORMATS.CURRENCY);
  sheet.getRange('D2:D100').setNumberFormat(FORMATS.PERCENT);
  sheet.getRange('E2:E100').setNumberFormat(FORMATS.CURRENCY);
  sheet.getRange('F2:G100').setNumberFormat(FORMATS.CURRENCY);
  sheet.getRange('H2:H100').setNumberFormat(FORMATS.PERCENT);

  sheet.getRange('A1').setNote(
    'Fill only col A (Name) and B (Symbol).\n' +
    'All price and data columns fill automatically.');
  Logger.log('Watchlist sheet ready.');
}

function _buildWatchlistFormulaRow(r) {
  var t = '"NSE:"&B' + r;
  return [
    '=IFERROR(GOOGLEFINANCE(' + t + ',"price"),0)',
    '=IFERROR(GOOGLEFINANCE(' + t + ',"changepct"),0)',
    '=IFERROR(C' + r + '*(D' + r + '/100),0)',
    '=IFERROR(GOOGLEFINANCE(' + t + ',"high52"),0)',
    '=IFERROR(GOOGLEFINANCE(' + t + ',"low52"),0)',
    '=IFERROR((C' + r + '/F' + r + '-1)*100,0)',
    '=IFERROR(VLOOKUP(UPPER(B' + r + '),_SectorData!A:B,2,0),"Other")',
  ];
}


// ================================================================
//  CHARTS SHEET (6 charts)
// ================================================================
function _buildChartsSheet(ss) {
  var sheet = ss.getSheetByName(SHEETS.CHARTS) || ss.insertSheet(SHEETS.CHARTS, 6);
  sheet.clear();
  sheet.getRange('A1').setValue('PORTFOLIO CHARTS')
    .setFontSize(18).setFontWeight('bold').setFontColor(COLORS.HEADER_BG);
  sheet.getRange('A2').setValue('Auto-updated every trading day at 4 PM IST.')
    .setFontColor(COLORS.MUTED).setFontSize(10);
}

function _refreshCharts(ss) {
  var chartsSheet = ss.getSheetByName(SHEETS.CHARTS);
  var histSheet   = ss.getSheetByName(SHEETS.HISTORY);
  var portSheet   = ss.getSheetByName(SHEETS.PORTFOLIO);
  var monthSheet  = ss.getSheetByName(SHEETS.MONTHLY);
  var weekSheet   = ss.getSheetByName(SHEETS.WEEKLY);
  if (!chartsSheet) return;

  chartsSheet.getCharts().forEach(function(c) { chartsSheet.removeChart(c); });

  var hR = histSheet  ? histSheet.getLastRow()  : 0;
  var pR = portSheet  ? portSheet.getLastRow()  : 0;
  var mR = monthSheet ? monthSheet.getLastRow() : 0;
  var wR = weekSheet  ? weekSheet.getLastRow()  : 0;

  if (hR > 2) chartsSheet.insertChart(
    chartsSheet.newChart().setChartType(Charts.ChartType.LINE)
      .addRange(histSheet.getRange(1, 1, hR, 2)).setPosition(4, 1, 0, 0)
      .setOption('title','Net Worth Over Time').setOption('width',560).setOption('height',320)
      .setOption('legend',{position:'none'}).setOption('colors',['#1565C0']).setOption('lineWidth',2)
      .setOption('curveType','function').setOption('vAxis',{title:'Value (₹)',format:'₹#,###'}).build());

  if (pR > 2) chartsSheet.insertChart(
    chartsSheet.newChart().setChartType(Charts.ChartType.PIE)
      .addRange(portSheet.getRange(1, 1, pR, 1))
      .addRange(portSheet.getRange(1, COL.CUR_VALUE, pR, 1))
      .setPosition(4, 9, 0, 0)
      .setOption('title','Portfolio Allocation').setOption('width',460).setOption('height',320)
      .setOption('pieHole',0.35).build());

  if (mR > 2) chartsSheet.insertChart(
    chartsSheet.newChart().setChartType(Charts.ChartType.COLUMN)
      .addRange(monthSheet.getRange(1, 1, mR, 1)).addRange(monthSheet.getRange(1, 4, mR, 1))
      .setPosition(24, 1, 0, 0)
      .setOption('title','Monthly P&L ₹').setOption('width',560).setOption('height',300)
      .setOption('legend',{position:'none'}).setOption('colors',['#42A5F5'])
      .setOption('vAxis',{format:'₹#,###'}).build());

  if (hR > 2) chartsSheet.insertChart(
    chartsSheet.newChart().setChartType(Charts.ChartType.AREA)
      .addRange(histSheet.getRange(1, 1, hR, 1)).addRange(histSheet.getRange(1, 6, hR, 1))
      .setPosition(24, 9, 0, 0)
      .setOption('title','Daily Change ₹').setOption('width',460).setOption('height',300)
      .setOption('legend',{position:'none'}).setOption('colors',['#66BB6A'])
      .setOption('vAxis',{format:'₹#,###'}).build());

  if (pR > 2) chartsSheet.insertChart(
    chartsSheet.newChart().setChartType(Charts.ChartType.BAR)
      .addRange(portSheet.getRange(1, 1, pR, 1))
      .addRange(portSheet.getRange(1, COL.PNL_PCT, pR, 1))
      .setPosition(44, 1, 0, 0)
      .setOption('title','Overall P&L % per Stock').setOption('width',560).setOption('height',320)
      .setOption('legend',{position:'none'}).setOption('colors',['#1565C0'])
      .setOption('hAxis',{format:'0.00"%"'}).build());

  if (wR > 2) chartsSheet.insertChart(
    chartsSheet.newChart().setChartType(Charts.ChartType.COLUMN)
      .addRange(weekSheet.getRange(1, 1, wR, 1)).addRange(weekSheet.getRange(1, 4, wR, 1))
      .setPosition(44, 9, 0, 0)
      .setOption('title','Weekly P&L ₹').setOption('width',460).setOption('height',320)
      .setOption('legend',{position:'none'}).setOption('colors',['#26A69A'])
      .setOption('vAxis',{format:'₹#,###'}).build());

  Logger.log('Charts: ' + chartsSheet.getCharts().length);
}


// ================================================================
//  DASHBOARD COLOR CODING
// ================================================================
function _applyDashboardColors(ss) {
  var sheet = ss.getSheetByName(SHEETS.DASHBOARD);
  if (!sheet) return;

  // Top stat cards
  [['H9', sheet.getRange('H9').getValue()],
   ['B13',sheet.getRange('B13').getValue()],
   ['E13',sheet.getRange('E13').getValue()]
  ].forEach(function(pair) {
    sheet.getRange(pair[0]).setFontColor(parseFloat(pair[1]) >= 0 ? COLORS.PROFIT : COLORS.LOSS);
  });

  // Holdings table: batch read all 3000 rows × 3 columns, then apply via getRangeList
  var vals = sheet.getRange(17, 7, 3000, 3).getValues();
  var profitRanges = []; var lossRanges = [];

  vals.forEach(function(row, i) {
    var dr = 17 + i;
    [0, 1, 2].forEach(function(ci) {
      var addr = String.fromCharCode(71 + ci) + dr; // G, H, I → cols 7,8,9
      (parseFloat(row[ci]) >= 0 ? profitRanges : lossRanges).push(addr);
    });
  });

  if (profitRanges.length) sheet.getRangeList(profitRanges).setFontColor(COLORS.PROFIT);
  if (lossRanges.length)   sheet.getRangeList(lossRanges).setFontColor(COLORS.LOSS);

  Logger.log('Dashboard colors applied.');
}


// ================================================================
//  UTILITIES
// ================================================================
function _clearTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    var fn = t.getHandlerFunction();
    if (fn === 'runDailySnapshot' || fn === 'checkAlerts') ScriptApp.deleteTrigger(t);
  });
}

/**
 * fixFormats — fixes #VALUE! display on ₹ cells without touching data.
 * Apps Script → select "fixFormats" → Run.
 */
function fixFormats() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var p = ss.getSheetByName(SHEETS.PORTFOLIO);
  if (p) {
    p.getRange('D2:I100').setNumberFormat(FORMATS.CURRENCY);
    p.getRange('J2:L100').setNumberFormat(FORMATS.PERCENT);
    p.getRange('M2:O100').setNumberFormat(FORMATS.CURRENCY);
    p.getRange('P2:P100').setNumberFormat(FORMATS.PERCENT);
    p.getRange('R2:S100').setNumberFormat(FORMATS.CURRENCY);
    p.getRange('T2:T100').setNumberFormat(FORMATS.PERCENT);
  }
  var h = ss.getSheetByName(SHEETS.HISTORY);
  if (h) {
    h.getRange('A2:A1000').setNumberFormat(FORMATS.DATE);
    h.getRange('B2:D1000').setNumberFormat(FORMATS.CURRENCY);
    h.getRange('E2:E1000').setNumberFormat(FORMATS.PERCENT);
    h.getRange('F2:F1000').setNumberFormat(FORMATS.CURRENCY);
    h.getRange('G2:G1000').setNumberFormat(FORMATS.PERCENT);
  }
  [SHEETS.WEEKLY, SHEETS.MONTHLY].forEach(function(name) {
    var s = ss.getSheetByName(name);
    if (s) _applyPeriodFormats(s);
  });
  SpreadsheetApp.getUi().alert('Formats fixed!\n\nRun "forceSnapshotNow" to also refresh your data.');
}

/**
 * forceSnapshotNow — take a snapshot immediately.
 * Use after adding stocks or if data looks stale.
 */
function forceSnapshotNow() {
  SpreadsheetApp.flush();
  Utilities.sleep(5000);
  runDailySnapshot();
  SpreadsheetApp.getUi().alert('Snapshot complete!');
}

/**
 * fillNewStockFormulas — run after adding new rows to Portfolio sheet.
 */
function fillNewStockFormulas() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.PORTFOLIO);
  if (!sheet) return;
  var count = 0;
  for (var r = 2; r <= sheet.getLastRow(); r++) {
    if (!sheet.getRange(r, COL.SYMBOL).getValue()) continue;
    if (!sheet.getRange(r, COL.CUR_PRICE).getFormula()) {
      var formulas = _buildPortfolioFormulaRow(r);
      sheet.getRange(r, COL.CUR_PRICE, 1, 12).setFormulas([formulas.slice(0, 12)]);
      sheet.getRange(r, COL.TO_TARGET, 1,  2).setFormulas([formulas.slice(12)]);
      count++;
    }
  }
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert('Formulas added to ' + count + ' new row(s).');
}

/**
 * addToWatchlist — dialog to add a stock to the Watchlist sheet.
 */
function addToWatchlist() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.WATCHLIST);
  if (!sheet) { SpreadsheetApp.getUi().alert('Run setup first.'); return; }
  var ui   = SpreadsheetApp.getUi();
  var name = ui.prompt('Stock Name (e.g. Infosys)').getResponseText().trim();
  var sym  = ui.prompt('NSE Symbol (e.g. INFY)').getResponseText().trim().toUpperCase();
  if (!name || !sym) return;
  var r = sheet.getLastRow() + 1;
  sheet.getRange(r, 1, 1, 2).setValues([[name, sym]]);
  sheet.getRange(r, 3, 1, 7).setFormulas([_buildWatchlistFormulaRow(r)]);
  ui.alert(name + ' (' + sym + ') added to Watchlist!');
}
