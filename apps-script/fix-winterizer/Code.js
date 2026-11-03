/* ════════════════════════════════════════════════════════════
   ONE-TIME FIX (2026-11-03): Rd 7 winterizer was logging the wrong product.

   The "Lawn Care - Rd 7 (Winterizer)" service said Spring Valley 18-2-3 at
   4 lb / 1000 sq ft. What actually goes down is Lebanon Pro High Bio
   Winterizer 24-0-5 at 3 lb / 1000 sq ft (Devon, 2026-11-03).

   What it changes:
     1. The Rd 7 service (Services tab): product, rate, method, notes.
     2. The Products tab: "High Bio Winterizer Fertilizer (high K)" becomes
        Lebanon Pro High Bio Winterizer 24-0-5 at 3 lb.
     3. Every Rd 7 (Winterizer) row in the 2026 monthly tabs that still says
        Spring Valley 18-2-3: Product, Rate, Method, Notes, and Total Used
        (the logged pounds x 3/4, e.g. 40.00 lbs -> 30.00 lbs).
     4. Deletes one Rd 7 row from Nov 2 that had been logged twice (found by
        its Entry ID), only if that row is still the same customer.
   Nothing else is touched. Rows already fixed are skipped, so running it
   twice does no harm.

   Its own standalone Apps Script project ("LPP winterizer fix 2026-11-03"),
   so the live LPP Log backend is not touched. Run fixWinterizerPreview first
   (changes nothing, logs the list), then fixWinterizerApply, which first makes
   a dated copy of the whole sheet in the "LPP App Backups" folder.

   getServices() and saveServicesJson() below are copied from the LPP Log
   backend (script v22 layout: JSON in column B of the DATA rows, pieces
   after the first start with '~').
   ════════════════════════════════════════════════════════════ */

var FIXW_SHEET_ID = '1ZLuqBz61IXa5CEQieNQbhTLxH6EJYGmxJiXjeyaswa0';
var FIXW_SERVICE  = 'Lawn Care - Rd 7 (Winterizer)';
var FIXW_OLD_PROD = 'Spring Valley 18-2-3';
var FIXW_NEW_PROD = 'Lebanon Pro High Bio Winterizer 24-0-5';
var FIXW_NEW_RATE = '3 lbs / 1000 sqft granular';
var FIXW_OLD_NOTE = 'Winterizer: Spring Valley 18-2-3 (50% ReGain, 30% bio, 1% iron) - the same bag as Rd 6.';
var FIXW_NEW_NOTE = 'Winterizer: Lebanon Pro High Bio Winterizer 24-0-5 at 3 lb / 1000 sq ft = 0.72 lb N / 1000 (PA limit 0.9 total, 0.7 available per application).';
// the customer's name is left out of this public copy; the copy that ran had it
var FIXW_DUP      = { entryId:'E1793654115635', customer:'(customer name)', tab:'Nov_2026' };
var FIXW_OLD_PRODUCTS_NAME = 'High Bio Winterizer Fertilizer (high K)';

function fixWinterizerPreview() { fixWinterizer_(false); }
function fixWinterizerApply()   { fixWinterizer_(true); }

function fixWinterizer_(apply) {
  var ss = SpreadsheetApp.openById(FIXW_SHEET_ID);
  var say = function(s) { Logger.log(s); };
  say(apply ? '=== APPLYING ===' : '=== PREVIEW ONLY - nothing is changed ===');

  if (apply) {
    var name = 'LPP App Log - backup ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HHmm') + ' before winterizer fix';
    var it = DriveApp.getFoldersByName('LPP App Backups');
    var copy = it.hasNext() ? DriveApp.getFileById(FIXW_SHEET_ID).makeCopy(name, it.next()) : DriveApp.getFileById(FIXW_SHEET_ID).makeCopy(name);
    say('Backup made: ' + name + ' (' + copy.getUrl() + ')');
  }

  // 1. the service
  var sv = getServices(ss);
  if (!sv || !sv[FIXW_SERVICE]) {
    say('!! Service "' + FIXW_SERVICE + '" not found - service not changed.');
  } else {
    var s = sv[FIXW_SERVICE], p = (s.products || [])[0] || {};
    if (p.product === FIXW_NEW_PROD) say('Service: already Lebanon Pro - skipped.');
    else {
      say('Service: "' + p.product + '" @ ' + p.rate + '  ->  "' + FIXW_NEW_PROD + '" @ ' + FIXW_NEW_RATE);
      if (apply) {
        p.product = FIXW_NEW_PROD;
        p.rate = FIXW_NEW_RATE;
        p.method = String(p.method || '').replace(/@\s*4\s*lbs?\s*\/\s*1000/i, '@ 3 lbs/1000');
        s.notes = FIXW_NEW_NOTE;
        s.products[0] = p;
        saveServicesJson(ss, JSON.stringify(sv));
      }
    }
  }

  // 2. the Products tab
  var ps = ss.getSheetByName('Products');
  if (ps) {
    var pv = ps.getDataRange().getValues(), found = false;
    for (var i = 1; i < pv.length; i++) {
      var nm = String(pv[i][0]).trim();
      if (nm === FIXW_NEW_PROD) { found = true; say('Products: Lebanon Pro already listed - skipped.'); break; }
      if (nm === FIXW_OLD_PRODUCTS_NAME) {
        found = true;
        say('Products row ' + (i + 1) + ': "' + nm + '" @ ' + pv[i][3] + '  ->  "' + FIXW_NEW_PROD + '" @ ' + FIXW_NEW_RATE);
        if (apply) {
          ps.getRange(i + 1, 1).setValue(FIXW_NEW_PROD);
          ps.getRange(i + 1, 4).setValue(FIXW_NEW_RATE);
          ps.getRange(i + 1, 5).setValue(String(pv[i][4]).replace(/@\s*4\s*lbs?\s*\/\s*1000/i, '@ 3 lbs/1000'));
        }
        break;
      }
    }
    if (!found) say('Products: no winterizer row found - not changed.');
  }

  // 3. the logged rows
  var fixed = 0;
  ss.getSheets().forEach(function(sh) {
    var tab = sh.getName();
    if (!/^[A-Z][a-z]{2}_2026$/.test(tab) || sh.getLastRow() < 2) return;
    var v = sh.getDataRange().getValues(), h = v[0].map(function(x) { return String(x).trim(); });
    var c = function(n) { return h.indexOf(n); };
    var cSvc = c('Service'), cProd = c('Product'), cRate = c('Rate'), cMeth = c('Method'), cNote = c('Notes'), cUsed = c('Total Used'), cCust = c('Customer'), cDate = c('Date');
    if (cSvc < 0 || cProd < 0) return;
    for (var r = 1; r < v.length; r++) {
      if (String(v[r][cSvc]).trim() !== FIXW_SERVICE || String(v[r][cProd]).trim() !== FIXW_OLD_PROD) continue;
      var used = String(v[r][cUsed]), m = /([\d.]+)\s*lbs?/i.exec(used);
      var newUsed = m ? (parseFloat(m[1]) * 3 / 4).toFixed(2) + ' lbs' : used;
      var d = v[r][cDate] instanceof Date ? Utilities.formatDate(v[r][cDate], Session.getScriptTimeZone(), 'MM/dd/yyyy') : v[r][cDate];
      say(tab + ' row ' + (r + 1) + ': ' + d + ' ' + v[r][cCust] + ' | ' + used + ' -> ' + newUsed + (m ? '' : '  (!! could not read pounds - left as is)'));
      if (apply) {
        sh.getRange(r + 1, cProd + 1).setValue(FIXW_NEW_PROD);
        if (cRate >= 0) sh.getRange(r + 1, cRate + 1).setValue(FIXW_NEW_RATE);
        if (cMeth >= 0) sh.getRange(r + 1, cMeth + 1).setValue(String(v[r][cMeth]).replace(/@\s*4\s*lbs?\s*\/\s*1000/i, '@ 3 lbs/1000'));
        if (cNote >= 0) sh.getRange(r + 1, cNote + 1).setValue(String(v[r][cNote]).replace(FIXW_OLD_NOTE, FIXW_NEW_NOTE));
        if (cUsed >= 0 && m) sh.getRange(r + 1, cUsed + 1).setValue(newUsed);
      }
      fixed++;
    }
  });
  say(fixed + ' winterizer row(s) ' + (apply ? 'corrected.' : 'would be corrected.'));

  // 4. the duplicate (done last, so the row numbers above stay right)
  var dsh = ss.getSheetByName(FIXW_DUP.tab);
  if (dsh) {
    var dv = dsh.getDataRange().getValues(), dh = dv[0].map(function(x) { return String(x).trim(); });
    var cId = dh.indexOf('Entry ID'), cC = dh.indexOf('Customer'), cS = dh.indexOf('Service'), hit = -1;
    for (var k = 1; k < dv.length; k++) if (String(dv[k][cId]).trim() === FIXW_DUP.entryId) { hit = k; break; }
    if (hit < 0) say('Duplicate: Entry ID ' + FIXW_DUP.entryId + ' not found (already removed?) - nothing deleted.');
    else if (String(dv[hit][cC]).trim() !== FIXW_DUP.customer || String(dv[hit][cS]).trim() !== FIXW_SERVICE) say('!! Duplicate: row ' + (hit + 1) + ' is "' + dv[hit][cC] + '" / "' + dv[hit][cS] + '", not ' + FIXW_DUP.customer + ' Rd 7 - nothing deleted.');
    else {
      say('Duplicate: ' + FIXW_DUP.tab + ' row ' + (hit + 1) + ' ' + FIXW_DUP.customer + ' Rd 7 (Entry ID ' + FIXW_DUP.entryId + ') ' + (apply ? 'deleted.' : 'would be deleted.'));
      if (apply) dsh.deleteRow(hit + 1);
    }
  }
  say(apply ? '=== DONE ===' : '=== END OF PREVIEW ===');
}

// ── copied from the LPP Log backend ──
var SERVICES_CHUNK = 40000;
var SERVICES_MIN_ROWS = 12;

function getServices(ss) {
  var sheet = ss.getSheetByName('Services');
  if(!sheet || sheet.getLastRow()<1) return null;
  var rows = sheet.getDataRange().getValues();
  var text = '';
  for(var i=0;i<rows.length;i++){
    if(String(rows[i][0]).trim().toUpperCase()==='DATA' && rows[i][1]){
      var piece = String(rows[i][1]);
      text += piece.charAt(0)==='~' ? piece.slice(1) : piece;
    }
  }
  if(!text) return null;
  try{ return JSON.parse(text); }catch(e){ return null; }
}

function saveServicesJson(ss, text) {
  var sheet = ss.getSheetByName('Services');
  var rows = [['Key','Value']];
  var i = 0;
  while (i < text.length) {
    var end = Math.min(i + SERVICES_CHUNK, text.length);
    var code = text.charCodeAt(end - 1);
    if (end < text.length && code >= 0xD800 && code <= 0xDBFF) end--;
    rows.push(['DATA', (i ? '~' : '') + text.slice(i, end)]);
    i = end;
  }
  var n = Math.max(rows.length, sheet.getLastRow(), SERVICES_MIN_ROWS);
  while (rows.length < n) rows.push(['','']);
  sheet.getRange(1, 1, n, 2).setValues(rows);
}
