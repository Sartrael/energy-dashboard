/**
 * Coenergy Dashboard - Google Apps Script Backend
 * Cole este código no App Script da sua Google Sheet.
 */

var SHEET_NAME = "Dados";

// Prepara a planilha caso esteja vazia
function setupSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(SHEET_NAME);
  }
  
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["ID", "Nome do Cliente", "Data de Inclusão", "Média de kW", "Status", "Data de Saída", "Inadimplente"]);
  }
}

// Lida com requisições GET (Buscar dados)
function doGet(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return ContentService.createTextOutput(JSON.stringify([])).setMimeType(ContentService.MimeType.JSON);
  
  var data = sheet.getDataRange().getValues();
  var json = [];
  
  if (data.length > 1) {
    var headers = data[0];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var obj = {
        id: row[0],
        clientName: row[1],
        dateInclusion: row[2],
        avgKw: row[3],
        status: row[4],
        dateSaida: row[5] || "",
        inadimplente: row[6] === true || String(row[6]).toLowerCase() === 'verdadeiro' || String(row[6]).toLowerCase() === 'true'
      };
      
      // Parse dateInclusion
      if (row[2] instanceof Date) {
        var d = row[2];
        var month = '' + (d.getMonth() + 1);
        var day = '' + d.getDate();
        var year = d.getFullYear();
        if (month.length < 2) month = '0' + month;
        if (day.length < 2) day = '0' + day;
        obj.dateInclusion = [year, month, day].join('-');
      } else if (typeof row[2] === 'string') {
        var dp = row[2].split('/');
        if (dp.length === 3) {
            obj.dateInclusion = [dp[2], dp[1].padStart(2, '0'), dp[0].padStart(2, '0')].join('-');
        } else {
            obj.dateInclusion = row[2];
        }
      }

      // Parse dateSaida
      if (row[5] instanceof Date) {
        var d2 = row[5];
        var m2 = '' + (d2.getMonth() + 1);
        var d2_day = '' + d2.getDate();
        var y2 = d2.getFullYear();
        if (m2.length < 2) m2 = '0' + m2;
        if (d2_day.length < 2) d2_day = '0' + d2_day;
        obj.dateSaida = [y2, m2, d2_day].join('-');
      } else if (typeof row[5] === 'string' && row[5].includes('/')) {
        var dp2 = row[5].split('/');
        if (dp2.length === 3) {
            obj.dateSaida = [dp2[2], dp2[1].padStart(2, '0'), dp2[0].padStart(2, '0')].join('-');
        }
      }
      
      json.push(obj);
    }
  }
  
  return ContentService.createTextOutput(JSON.stringify(json))
    .setMimeType(ContentService.MimeType.JSON)
    .setHeader("Access-Control-Allow-Origin", "*");
}

// Lida com requisições POST (Adicionar e Deletar)
function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    var body = JSON.parse(e.postData.contents);
    var action = body.action;
    
    if (action === 'add') {
      var data = sheet.getDataRange().getValues();
      var maxId = 0;
      for (var i = 1; i < data.length; i++) {
        var cid = parseInt(data[i][0]);
        if (!isNaN(cid) && cid > maxId) {
          maxId = cid;
        }
      }
      var newId = maxId + 1;
      
      var isInadimplente = body.inadimplente === true;
      sheet.appendRow([newId, body.clientName, body.dateInclusion, body.avgKw, body.status, body.dateSaida || "", isInadimplente]);
      return ContentService.createTextOutput(JSON.stringify({success: true, id: newId})).setMimeType(ContentService.MimeType.JSON);
    }
    
    if (action === 'add_lote') {
      var data = sheet.getDataRange().getValues();
      var maxId = 0;
      for (var i = 1; i < data.length; i++) {
        var cid = parseInt(data[i][0]);
        if (!isNaN(cid) && cid > maxId) {
          maxId = cid;
        }
      }
      
      var clients = body.clients || [];
      var nextId = maxId + 1;
      
      for (var j = 0; j < clients.length; j++) {
        var client = clients[j];
        var isInadimplente = client.inadimplente === true;
        sheet.appendRow([nextId, client.clientName, client.dateInclusion, client.avgKw, client.status, client.dateSaida || "", isInadimplente]);
        nextId++;
      }
      
      return ContentService.createTextOutput(JSON.stringify({success: true, count: clients.length})).setMimeType(ContentService.MimeType.JSON);
    }
    
    if (action === 'toggle_inadimplente') {
      var data = sheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        if (data[i][0] == body.id) {
          var rowIndex = i + 1;
          var currentVal = sheet.getRange(rowIndex, 7).getValue();
          var newVal = !(currentVal === true || String(currentVal).toLowerCase() === 'verdadeiro' || String(currentVal).toLowerCase() === 'true');
          sheet.getRange(rowIndex, 7).setValue(newVal);
          return ContentService.createTextOutput(JSON.stringify({success: true, inadimplente: newVal})).setMimeType(ContentService.MimeType.JSON);
        }
      }
      return ContentService.createTextOutput(JSON.stringify({success: false, message: "ID não encontrado"})).setMimeType(ContentService.MimeType.JSON);
    }
    
    if (action === 'delete') {
      var data = sheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        if (data[i][0] == body.id) {
          sheet.deleteRow(i + 1);
          return ContentService.createTextOutput(JSON.stringify({success: true})).setMimeType(ContentService.MimeType.JSON);
        }
      }
    }
    
    if (action === 'delete_month') {
      var data = sheet.getDataRange().getValues();
      var rowsToDelete = [];
      for (var i = 1; i < data.length; i++) {
        var dateVal = data[i][2];
        var dYear, dMonth;
        if (dateVal instanceof Date) {
            dYear = dateVal.getFullYear();
            dMonth = dateVal.getMonth();
        } else if (typeof dateVal === 'string') {
            var dp = dateVal.split(/[-\/]/);
            if (dateVal.includes('/')) {
                dYear = parseInt(dp[2], 10);
                dMonth = parseInt(dp[1], 10) - 1;
            } else if (dateVal.includes('-')) {
                dYear = parseInt(dp[0], 10);
                dMonth = parseInt(dp[1], 10) - 1;
            }
        }
        if (dYear == body.year && dMonth == body.monthIdx) {
            rowsToDelete.push(i + 1);
        }
      }
      
      for (var j = rowsToDelete.length - 1; j >= 0; j--) {
        sheet.deleteRow(rowsToDelete[j]);
      }
      return ContentService.createTextOutput(JSON.stringify({success: true, deleted: rowsToDelete.length})).setMimeType(ContentService.MimeType.JSON);
    }
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({error: err.toString()})).setMimeType(ContentService.MimeType.JSON);
  }
  
  return ContentService.createTextOutput(JSON.stringify({success: false, message: "Ação não reconhecida ou ID não encontrado"})).setMimeType(ContentService.MimeType.JSON);
}
