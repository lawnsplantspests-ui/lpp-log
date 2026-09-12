// ================================================================
// LPP App Log — Google Apps Script API v9
// Sheet: LPP App Log (1ZLuqBz61IXa5CEQieNQbhTLxH6EJYGmxJiXjeyaswa0)
//
// v9 changes:
//   - addCustomer / updateCustomer no longer destroy the Address column.
//     They wrote 7 columns with UpdatedAt in position 7, but column 7 is
//     Address; UpdatedAt is column 13. They now write columns 1-6 and set
//     column 13 separately, leaving Address/City/State/Zip/Phone/Email alone.
//   - ensureCustomerHeader writes the real 13-column layout.
//   - getLog returns Site Address.
//   - updateLogColumn / updateLogCell / deleteLogRow for corrections.
// ================================================================

var SHEET_ID = '1ZLuqBz61IXa5CEQieNQbhTLxH6EJYGmxJiXjeyaswa0';

// Tabs that are NOT log tabs
var SYSTEM_TABS = ['Customers','Services','Products','SyncQueue',
  'FP_LawnVisits','FP_PhcVisits','FP_SoilSamples','FP_BaitStations',
  'FP_Ponds','FP_PondVisits','FP_Programs','FP_CustPrograms',
  'FP_Routes','FP_Products'];

function doGet(e) {
  return json({ status:'ok', message:'LPP API v9 running' });
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss   = SpreadsheetApp.openById(SHEET_ID);
    var action = data.action || '';

    // ── GET ALL (startup) ──────────────────────────────────────
    if (action === 'getAll') {
      return json({
        status:    'ok',
        customers: getCustomers(ss),
        services:  getServices(ss),
        products:  getProducts(ss)
      });
    }

    // ── GET LOG TABS ───────────────────────────────────────────
    if (action === 'getLogTabs') {
      var tabs = ss.getSheets()
        .map(function(s){ return s.getName(); })
        .filter(function(n){ return SYSTEM_TABS.indexOf(n) < 0; })
        .sort().reverse(); // newest first
      return json({ status:'ok', tabs:tabs });
    }

    // ── GET LOG (fetch one monthly tab) ────────────────────────
    if (action === 'getLog') {
      var tab   = data.tab || '';
      var sheet = ss.getSheetByName(tab);
      if (!sheet || sheet.getLastRow() < 2)
        return json({ status:'ok', entries:[] });

      var rows = sheet.getDataRange().getValues();
      var h    = rows[0].map(function(c){ return String(c).trim(); });
      var entries = [];

      for (var i = 1; i < rows.length; i++) {
        var r = rows[i];
        var entry = {};
        h.forEach(function(col, j){ entry[col] = cellStr(r[j]); });
        if (!entry['Customer']) continue; // skip blank rows
        entries.push({
          id:          'sh_' + tab + '_' + i,
          date:        entry['Date']              || '',
          time:        entry['Time']              || '',
          customer:    entry['Customer']          || '',
          sqFt:        entry['Sq Ft']             || '',
          service:     entry['Service']           || '',
          product:     entry['Product']           || '',
          activeIng:   entry['Active Ingredient'] || '',
          epa:         entry['EPA Reg #']         || '',
          targetPest:  entry['Target Pest']       || '',
          appSite:     entry['App Site']          || '',
          method:      entry['Method']            || '',
          rate:        entry['Rate']              || '',
          totalMixed:  entry['Total Mixed']       || '',
          totalUsed:   entry['Total Used']        || '',
          spotSqFt:    entry['Spot Spray Sq Ft']  || '',
          weather:     entry['Weather']           || '',
          wind:        entry['Wind']              || '',
          temp:        entry['Temp']              || '',
          notes:       entry['Notes']             || '',
          siteAddress: entry['Site Address']      || '',
          dbhIn:       entry['DBH (in)']          || '',
          synced:      true
        });
      }
      return json({ status:'ok', entries:entries });
    }

    // ── CUSTOMERS ──────────────────────────────────────────────
    // Customers columns:
    //  1 ID | 2 Name | 3 SqFt_Lawn | 4 SqFt_PHC | 5 SqFt_Pest | 6 SqFt_Aquatic
    //  7 Address | 8 City | 9 State | 10 Zip | 11 Phone | 12 Email | 13 UpdatedAt
    // Only 1-6 and 13 are ever written here. Never touch 7-12.
    if (action === 'addCustomer') {
      var sheet = getOrCreate(ss, 'Customers');
      ensureCustomerHeader(sheet);
      var id = String(Date.now());
      var newRow = sheet.getLastRow() + 1;
      sheet.getRange(newRow,1,1,6).setValues([[
        id, data.name,
        data.sqFtLawn||data.sqFt||'', data.sqFtPHC||'',
        data.sqFtPest||'', data.sqFtAquatic||''
      ]]);
      sheet.getRange(newRow,13).setValue(new Date().toISOString());
      return json({ status:'ok', id:id });
    }

    if (action === 'updateCustomer') {
      var sheet = getOrCreate(ss, 'Customers');
      var rows  = sheet.getDataRange().getValues();
      var h2    = rows[0].map(function(c){return String(c).trim().toLowerCase().replace(/_/g,'');});
      var idCol = h2.indexOf('id'); if(idCol<0) idCol=0;
      for (var i = 1; i < rows.length; i++) {
        if (String(rows[i][idCol]) === String(data.id)) {
          sheet.getRange(i+1,1,1,6).setValues([[
            data.id, data.name,
            data.sqFtLawn||data.sqFt||'', data.sqFtPHC||'',
            data.sqFtPest||'', data.sqFtAquatic||''
          ]]);
          sheet.getRange(i+1,13).setValue(new Date().toISOString());
          return json({ status:'ok' });
        }
      }
      var newRow2 = sheet.getLastRow() + 1;
      sheet.getRange(newRow2,1,1,6).setValues([[
        data.id, data.name,
        data.sqFtLawn||data.sqFt||'', data.sqFtPHC||'',
        data.sqFtPest||'', data.sqFtAquatic||''
      ]]);
      sheet.getRange(newRow2,13).setValue(new Date().toISOString());
      return json({ status:'ok' });
    }

    // Write address fields without disturbing the sq ft columns.
    if (action === 'updateCustomerAddress') {
      var sheet = getOrCreate(ss, 'Customers');
      var rows  = sheet.getDataRange().getValues();
      var h3    = rows[0].map(function(c){return String(c).trim().toLowerCase().replace(/_/g,'');});
      var idCol3 = h3.indexOf('id'); if(idCol3<0) idCol3=0;
      for (var k = 1; k < rows.length; k++) {
        if (String(rows[k][idCol3]) === String(data.id)) {
          sheet.getRange(k+1,7,1,4).setValues([[
            data.address||'', data.city||'', data.state||'', data.zip||''
          ]]);
          sheet.getRange(k+1,13).setValue(new Date().toISOString());
          return json({ status:'ok' });
        }
      }
      return json({ status:'error', message:'Customer id not found: '+data.id });
    }

    if (action === 'deleteCustomer') {
      var sheet = getOrCreate(ss, 'Customers');
      var rows  = sheet.getDataRange().getValues();
      var h2    = rows[0].map(function(c){return String(c).trim().toLowerCase().replace(/_/g,'');});
      var idCol = h2.indexOf('id'); if(idCol<0) idCol=0;
      for (var i = rows.length-1; i >= 1; i--) {
        if (String(rows[i][idCol]) === String(data.id)) {
          sheet.deleteRow(i+1);
          return json({ status:'ok' });
        }
      }
      return json({ status:'ok' });
    }

    // ── SERVICES ───────────────────────────────────────────────
    if (action === 'saveServices') {
      var sheet = getOrCreate(ss, 'Services');
      sheet.clearContents();
      sheet.appendRow(['Key','Value']);
      sheet.appendRow(['DATA', data.services]);
      return json({ status:'ok' });
    }

    // ── PRODUCTS ───────────────────────────────────────────────
    if (action === 'getProducts') {
      return json({ status:'ok', data:getProducts(ss) });
    }
    if (action === 'addProduct') {
      var sheet = getOrCreate(ss, 'Products');
      ensureProductHeader(sheet);
      sheet.appendRow([data.product, data.activeIng||'', data.epa||'',
        data.rate||'', data.method||'', data.targetPest||'',
        data.notes||'', new Date().toLocaleDateString()]);
      return json({ status:'ok' });
    }
    if (action === 'updateProduct') {
      var sheet = getOrCreate(ss, 'Products');
      var rows  = sheet.getDataRange().getValues();
      for (var i = 1; i < rows.length; i++) {
        if (String(rows[i][0]).trim().toLowerCase() === String(data.originalName).trim().toLowerCase()) {
          sheet.getRange(i+1,1,1,7).setValues([[data.product,
            data.activeIng||'', data.epa||'', data.rate||'',
            data.method||'', data.targetPest||'', data.notes||'']]);
          return json({ status:'ok' });
        }
      }
      return json({ status:'ok' });
    }
    if (action === 'deleteProduct') {
      var sheet = getOrCreate(ss, 'Products');
      var rows  = sheet.getDataRange().getValues();
      for (var i = rows.length-1; i >= 1; i--) {
        if (String(rows[i][0]).trim().toLowerCase() === String(data.product).trim().toLowerCase()) {
          sheet.deleteRow(i+1);
          return json({ status:'ok' });
        }
      }
      return json({ status:'ok' });
    }

    // ── APPEND LOG ENTRY ───────────────────────────────────────
    if (action === 'appendLog') {
      var tab   = data.tab || 'Log';
      var sheet = getOrCreate(ss, tab);
      if (sheet.getLastRow() === 0) sheet.appendRow(data.headers);
      // Keep the header row as wide as the rows being written, so a column
      // added later (Site Address) does not sit under a blank heading.
      if (data.headers && data.headers.length > sheet.getLastColumn()) {
        sheet.getRange(1,1,1,data.headers.length).setValues([data.headers]);
      }
      sheet.appendRow(data.row);
      return json({ status:'ok' });
    }

    // ── SYNC QUEUE ─────────────────────────────────────────────
    if (action === 'syncQueue') {
      var sheet = getOrCreate(ss, 'SyncQueue');
      if (sheet.getLastRow() === 0) sheet.appendRow(['Timestamp','Type','Payload']);
      sheet.appendRow([new Date().toISOString(), data.type||'entry',
        JSON.stringify(data.payload||{})]);
      return json({ status:'ok' });
    }

    // ── UPDATE LOG (corrections) ───────────────────────────────
    if (action === 'updateLogColumn') {
      var sheet = ss.getSheetByName(data.tab);
      if (!sheet) return json({ status:'error', message:'Tab not found: '+data.tab });
      var rows = sheet.getLastRow() - 1;
      if (rows !== (data.values||[]).length)
        return json({ status:'error', message:'Row mismatch on '+data.tab+': sheet has '+rows+', got '+(data.values||[]).length });
      var vals = data.values.map(function(v){ return [v]; });
      sheet.getRange(2, data.col, vals.length, 1).setValues(vals);
      return json({ status:'ok', rows:vals.length });
    }

    if (action === 'updateLogCell') {
      var sheet = ss.getSheetByName(data.tab);
      if (!sheet) return json({ status:'error', message:'Tab not found: '+data.tab });
      sheet.getRange(data.row, data.col).setValue(data.value);
      return json({ status:'ok' });
    }

    // Delete one log row. Requires the customer name to match, so a stale
    // row number cannot quietly delete somebody else's record.
    if (action === 'deleteLogRow') {
      var sheet = ss.getSheetByName(data.tab);
      if (!sheet) return json({ status:'error', message:'Tab not found: '+data.tab });
      var got = String(sheet.getRange(data.row, 3).getValue()).trim();
      if (got !== String(data.confirmCustomer||'').trim())
        return json({ status:'error', message:'Row '+data.row+' is "'+got+'", not "'+data.confirmCustomer+'". Nothing deleted.' });
      sheet.deleteRow(data.row);
      return json({ status:'ok', deleted:got });
    }

    return json({ status:'error', message:'Unknown action: '+action });

  } catch(err) {
    return json({ status:'error', message:err.toString() });
  }
}

// ── CELL VALUE HELPER ──────────────────────────────────────────────
function cellStr(val) {
  if (val === null || val === undefined || val === '') return '';
  if (val instanceof Date) {
    var yr = val.getFullYear();
    // 1899/1900 = time-only value stored as date (Google Sheets quirk)
    if (yr === 1899 || yr === 1900) {
      var h = val.getUTCHours(), m = val.getUTCMinutes();
      var ap = h >= 12 ? 'PM' : 'AM';
      h = h % 12 || 12;
      return h + ':' + (m < 10 ? '0' : '') + m + ' ' + ap;
    }
    // Use LOCAL date (not UTC) to avoid timezone shifting the date
    var mo = String(val.getMonth() + 1).padStart(2, '0');
    var dy = String(val.getDate()).padStart(2, '0');
    return mo + '/' + dy + '/' + yr;
  }
  var s = String(val).trim();
  return s;
}

// Format a Date as MM/DD/YYYY using local timezone
function fmtDate(d) {
  return String(d.getMonth()+1).padStart(2,'0') + '/'
       + String(d.getDate()).padStart(2,'0') + '/'
       + d.getFullYear();
}

// ── READ HELPERS ───────────────────────────────────────────────────
function getCustomers(ss) {
  var sheet = ss.getSheetByName('Customers');
  if (!sheet || sheet.getLastRow() < 2) return [];
  var rows = sheet.getDataRange().getValues();
  var h = rows[0].map(function(c){
    return String(c).trim().toLowerCase().replace(/_/g,'').replace(/\s/g,'');
  });
  var idCol=h.indexOf('id'), nameCol=h.indexOf('name');
  var sqLawnCol=h.indexOf('sqftlawn'); if(sqLawnCol<0) sqLawnCol=h.indexOf('sqft');
  var sqPHCCol=h.indexOf('sqftphc'), sqPestCol=h.indexOf('sqftpest'),
      sqAquaticCol=h.indexOf('sqftaquatic');
  var addrCol=h.indexOf('address'), cityCol=h.indexOf('city'),
      stateCol=h.indexOf('state'), zipCol=h.indexOf('zip');
  if(nameCol<0){nameCol=1;idCol=0;sqLawnCol=2;sqPHCCol=3;sqPestCol=4;sqAquaticCol=5;}
  var out=[];
  for(var i=1;i<rows.length;i++){
    var name=String(rows[i][nameCol]||'').trim();
    if(!name) continue;
    out.push({
      id:          idCol>=0?String(rows[i][idCol]):'imp_'+i,
      name:        name,
      sqFt:        sqLawnCol>=0?String(rows[i][sqLawnCol]||''):'',
      sqFtLawn:    sqLawnCol>=0?String(rows[i][sqLawnCol]||''):'',
      sqFtPHC:     sqPHCCol>=0?String(rows[i][sqPHCCol]||''):'',
      sqFtPest:    sqPestCol>=0?String(rows[i][sqPestCol]||''):'',
      sqFtAquatic: sqAquaticCol>=0?String(rows[i][sqAquaticCol]||''):'',
      address:     addrCol>=0?String(rows[i][addrCol]||''):'',
      city:        cityCol>=0?String(rows[i][cityCol]||''):'',
      state:       stateCol>=0?String(rows[i][stateCol]||''):'',
      zip:         zipCol>=0?String(rows[i][zipCol]||''):''
    });
  }
  return out;
}

function getServices(ss) {
  var sheet = ss.getSheetByName('Services');
  if(!sheet || sheet.getLastRow()<1) return null;
  var rows = sheet.getDataRange().getValues();
  for(var i=0;i<rows.length;i++){
    if(String(rows[i][0]).trim().toUpperCase()==='DATA' && rows[i][1]){
      try{ return JSON.parse(String(rows[i][1])); }catch(e){ return null; }
    }
  }
  return null;
}

function getProducts(ss) {
  var sheet = ss.getSheetByName('Products');
  if(!sheet || sheet.getLastRow()<2) return [];
  var rows = sheet.getDataRange().getValues();
  var h = rows[0].map(function(c){
    return String(c).trim().toLowerCase().replace(/\s/g,'').replace(/#/g,'');
  });
  var pCol=h.indexOf('productname'); if(pCol<0) pCol=0;
  var aiCol=h.indexOf('activeingredient'); if(aiCol<0) aiCol=1;
  var eCol=h.indexOf('epareg'); if(eCol<0) eCol=2;
  var rCol=h.indexOf('rate'); if(rCol<0) rCol=3;
  var mCol=h.indexOf('method'); if(mCol<0) mCol=4;
  var tCol=h.indexOf('targetpest'); if(tCol<0) tCol=5;
  var nCol=h.indexOf('notes'); if(nCol<0) nCol=6;
  var out=[];
  for(var i=1;i<rows.length;i++){
    var name=String(rows[i][pCol]||'').trim();
    if(!name) continue;
    out.push({
      product:    name,
      activeIng:  String(rows[i][aiCol]||''),
      epa:        String(rows[i][eCol]||''),
      rate:       String(rows[i][rCol]||''),
      method:     String(rows[i][mCol]||''),
      targetPest: String(rows[i][tCol]||''),
      notes:      String(rows[i][nCol]||'')
    });
  }
  return out;
}

function getOrCreate(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}
function ensureCustomerHeader(sheet) {
  if(sheet.getLastRow()===0)
    sheet.appendRow(['ID','Name','SqFt_Lawn','SqFt_PHC','SqFt_Pest','SqFt_Aquatic',
                     'Address','City','State','Zip','Phone','Email','UpdatedAt']);
}
function ensureProductHeader(sheet) {
  if(sheet.getLastRow()===0)
    sheet.appendRow(['Product Name','Active Ingredient','EPA Reg #','Rate','Method','Target Pest','Notes','Added']);
}
function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── TESTS ──────────────────────────────────────────────────────────
function testGetLog() {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Mar_2026');
  if(!sheet){ Logger.log('Mar_2026 NOT FOUND'); return; }
  var rows  = sheet.getDataRange().getValues();
  Logger.log('Total rows: ' + (rows.length-1));
  for(var i=1;i<=Math.min(5,rows.length-1);i++){
    Logger.log('Row '+i+': date='+cellStr(rows[i][0])+' | customer='+rows[i][2]);
  }
}
function testAPI() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  Logger.log('Customers: ' + getCustomers(ss).length);
  Logger.log('Services: ' + (getServices(ss)?'found':'none'));
  Logger.log('Products: ' + getProducts(ss).length);
}
