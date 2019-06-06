var url = 'https://nzzz.meraki.com/api/v0/networks/L_xxxxxxxxxxxx/clients/provision';
var apiKey = 'yyyyyyyyyyyyyyyyyyyyyy';

function importclient(url, payload){
  var headers = {
    'x-cisco-meraki-api-key': apiKey
    };
  var options =
      {
      'headers': headers,
      'method': 'post',
      'followRedirects': false, 
      'validateHttpsCertificates': false, 
      'content-type': 'application/json',
      'payload': payload,
      'muteHttpExceptions' : true,
      };     
  var response = UrlFetchApp.fetch(url, options);
  //Logger.log( payload);
  Logger.log( response);
  
}

var IMPORTED = 'IMPORTED';
function importClient() {
  var sheet = SpreadsheetApp.getActive().getSheetByName('List')
  var rows = SpreadsheetApp.getActiveSheet().getLastRow();
  var lastRow = rows-1;
  var startRow = 2;

  // Fetch data range
  var dataRange = sheet.getRange(startRow, 1, lastRow, 5);
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  for (var i = 0; i < data.length; ++i) {
    var row = data[i];
    var macAddress = row[0];
    var clientName = row[1]; 
//    var groupPolicy = row[2]; 
 var groupPolicy;
    if (row[2] == "") {
          groupPolicy == "Normal"
    } 
    else{ 
       groupPolicy = row[2] ;
      }
    var groupPolicyId= row[3]; 
    var request_payload = {
     "mac": macAddress,
     "name": clientName,
     "devicePolicy": groupPolicy,
     "groupPolicyId": groupPolicyId
   };
    
    var check = row[4]; 
    if (check != IMPORTED) { // Prevents duplicates
      importclient(url,request_payload);
      sheet.getRange(startRow + i, 5).setValue(IMPORTED);
      // Make sure the cell is updated right away in case the script is interrupted
      SpreadsheetApp.flush();
    }
  }
}

