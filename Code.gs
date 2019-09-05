//NEW FEATURE 8-16-2019: CREATE FILE OF MARC RECORDS
var rand = Math.floor((Math.random() * 1000) + 1);
var fileToReturn = '<marc:collection xmlns:marc="http://www.loc.gov/MARC21/slim">';
//END NEW FEATURE

function onOpen() {
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .createMenu('Custom Menu')
      .addItem('Launch Match MARC OCLC Search', 'showSidebar')
      .addToUi();
}

function onInstall() {
  onOpen();
}


function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('sidebar')
      .setTitle('OCLC Lookup:')
      .setWidth(500);
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .showSidebar(html);    
}



//THIS FUNCTION IS LAUNCHED WHEN THE 'START SEARCH' BUTTON
//ON THE SIDEBAR IS CLICKED
//'form' REPRESENTS THE FORM ON THE SIDEBAR
//THIS METHOD IS THE HEART OF THE FUNCTIONALITY
//IT BOILS DOWN TO THREE LOOPS
//OUTER LOOP IS FOR EACH LOOKUP TO BE PERFORMED
//FOR EACH LOOKUP, LOOK AT THE SEARCH CRITERIA
//IN EACH OF THE ROWS - WHICH IS MADE UP OF MULTIPLE
//CELLS
function startLookup(form) {
   
   var ui = SpreadsheetApp.getUi();
   
   //MAKE SURE THE OCLC API KEY HAS BEEN ENTERED
   var apiKey = form.apiKey;
   if (apiKey == null || apiKey == "") {
     ui.alert("OCLC API Key is Required");
     return;
   }
   PropertiesService.getUserProperties().setProperty('apiKey', apiKey);
   
      
   var emailAddress = form.emailAddress;
   PropertiesService.getUserProperties().setProperty('emailAddress', emailAddress);
   
   //SETUP SHEETS/TABS/RANGES TO READ FROM/WRITE TO
   var settingsTabName = form.tabSelection;
   var dataTabName = form.searchForTab;
 
   var spreadsheet = SpreadsheetApp.getActive();
   var dataSheet = spreadsheet.getSheetByName(dataTabName);
   SpreadsheetApp.setActiveSheet(dataSheet);
   var settingsSheet = spreadsheet.getSheetByName(settingsTabName);
   var settingsRange = settingsSheet.getDataRange();
   var outputRange = settingsSheet.getRange(12,1,8,2);
   var checkLocalHoldings = settingsRange.getCell(2, 1).getDisplayValue();
   var checkLocal = false;
   var checkLocalCode = "";
   
   
   //DOES THE SEARCH CRITERIA INDICATE SEARCH FOR LOCAL HOLDINGS:
   if (checkLocalHoldings.indexOf('holdings') > -1) {
     checkLocal = true;
     var x = checkLocalHoldings.indexOf("=");
     checkLocalCode = checkLocalHoldings.substring(x+1,checkLocalHoldings.length);
   }
   
   
   //FOR EACH ITEM TO BE LOOKED UP IN THE DATA SPREADSHEET:
   var lastRow = dataSheet.getLastRow();
   var lastCol = 100;
   var dataRange = dataSheet.getRange(2, 1, lastRow , 100)
   var numRows = dataRange.getNumRows();
   for (var x = 1; x <= numRows; x++) {
        var isbnCell = dataRange.getCell(x,1);
        var lccnCell = dataRange.getCell(x,2)
        var searchCriteria = null;
        //IF ISBN COL IS BLANK, SEARCH WILL USE THE
        //VALUE IN THE LCCN COL.
        //IF THE ROW CONTAINS NEITHER IT MOVES TO THE NEXT ROW
        if (!isbnCell.isBlank()) {
          var isbn = isbnCell.getValue();
          if (isbn.length < 10) isbn = pad(10,isbn,0);
          searchCriteria = "srw.bn=" + isbn;
        }
        else if (!lccnCell.isBlank()) {
          searchCriteria = "srw.dn=" + lccnCell.getValue();
        }
        else {
          continue;
        }
        if (searchCriteria == null) continue;
        
        
        //IF SEARCH FOR LOCAL HOLDINGS IS REQUIRED, CALL THE API INLCUDING THE
        //OCLC SYMBOL
        if (checkLocal) {
          try {
            var foundLocalRecord = findLocalRecord(x,dataRange,searchCriteria,checkLocalCode,outputRange);
          }
          catch(err) {
            ui.alert("Communication with API failed.  Please check your API key.");
            ui.alert(err);
            return;
          }
         
         //IF IT FOUND A MATCH USING THE LOCAL SEARCH, MOVE ONTO THE NEXT ROW:
         if (foundLocalRecord) {
           dataRange.getCell(x,3).setValue("local record found");
           continue;
         }
       }
       
       
       //OTHERWISE, CALL THE API SORTED BY LIBRARY HOLDINGS COUNT
       try {
         var searchResults = findRecord(searchCriteria);
       }
       catch(err) {
         ui.alert("API call failed.  Please check your API key.");
         return;
       }
       
       
       var nsp = XmlService.getNamespace('http://www.loc.gov/zing/srw/');
       var slimNsp = XmlService.getNamespace('http://www.loc.gov/MARC21/slim'); 
     
       var root = searchResults.getRootElement();
       var test = root.getChild("numberOfRecords",nsp).getValue();
       var records = root.getChild("records",nsp);
       if (records == null) continue; //CONTINUE ON TO THE NEXT RECORD...API DIDN'T FIND ANYTHING
       
       

       var collectionOfSettingsRange = settingsSheet.getRange(3,1,8,6);
       
       
       //FOR EACH RECORD FOUND FOR THIS ISBN/LCCN:
       var listOfRecords = records.getChildren();
       var found = false;
       
       //LOOK AT THE MATCH CRITERIA:
       //NOTE: THE RESULTS ARE SORTED BY LIBRARY HOLDINGS COUNT (DESCENDING)
       //THE FIRST RECORD IT FINDS WITH ALL MATCH CRITERIA IS THE RECORD IT SELECTS
       //THAT MEANS THE RESULTS ARE A RECORD THAT MATCHED WHICH HAS THE LARGEST NUMBER
       //OF HOLDINGS
       for (var y = 0; y < listOfRecords.length; y++) {      
           var controlFieldsInLocalRecord = listOfRecords[y].getChild("recordData",nsp).getChild("record",slimNsp).getChildren("controlfield",slimNsp);
           var oclcNumber = getControlFieldValue(controlFieldsInLocalRecord,"001");
       
           //LOOPS THROUGH THE SETTINGS
           //ROWS OF MATCH CRITERIA
           for (var i = 1; i < 8; i++) {

               if (collectionOfSettingsRange.getCell(i, 1).isBlank()) continue //skip this row
               var matchedTheCriteria = 0;
               //LOOP THROUGH COLUMNS IN THE ROW
               for (var e = 1; e < 6; e++) {
                 // ui.alert("looking at row " + i + " col. " + e);
                 var v = collectionOfSettingsRange.getCell(i, e).getValue();
                 // ui.alert("this is the value " + v);
                 if (v==null || v == "") continue;
                 var indexOfValue = v.indexOf('=');
                 if (v.indexOf('$') > -1) {
                    var indexOfSubField = v.indexOf('$');
                    var subField = v.substring(indexOfSubField+1,indexOfValue);
                    //ui.alert('a subfield exists' + v);
                    var field = v.substring(0,indexOfSubField);
                  }
                  else {
                    var field = v.substring(0,indexOfValue);
                    var subField = "";
                  }
                   
                  var l = v.length;
                  var desiredValue = v.substring(indexOfValue+1,l);
                   
                   
                  var dataFields = listOfRecords[y].getChild("recordData",nsp).getChild("record",slimNsp).getChildren("datafield",slimNsp);
                  var controlFields = listOfRecords[y].getChild("recordData",nsp).getChild("record",slimNsp).getChildren("controlfield",slimNsp);  
                  //ui.alert("GETTING THE DATA FIELDS FOR " + field);
                  var dataField = getDataField(dataFields,field); //040
                  //ui.alert(dataField);
                  if (dataField == null) continue;
                  var subfields = dataField.getChildren("subfield",slimNsp);
                  //ui.alert(subfields);
                  var valueExists = 0;
                  if (subField != null && subField != "") {
                    //RETURNS A '1' IF THERE IS NO MATCH
                    //RETURNS A '0' IF THERE IS A MATCH
                    var valueExists = doesSubFieldContainWithin(subfields,desiredValue,subField); 
                  }
                  else {
                     //RETURNS A '1' IF THERE IS NO MATCH
                     //RETURNS A '0' IF THERE IS A MATCH
                     var valueExists = doesSubFieldContain(subfields,desiredValue);
                  }
                  
                   matchedTheCriteria = matchedTheCriteria + valueExists;
                  //ui.alert("matched the criteria: " + matchedTheCriteria + " / " + desiredValue + "/" + subField);
                 
                 //IF THE MATCH CRITERIA IS GREATER THAN ZERO IT MEANS
                 //ONE OF THE COLS IN THE ROW EVALUATED TO FALSE - GO ON TO THE NEXT ROW OF CRITERIA
                 if (matchedTheCriteria > 0) break; //stop looking in this row
               }
              
             //IF THE MATCHED CRITERIA HAS REMAINED AT ZERO ALL OF THE COLS
             //IN THE ROW EVALUATED TO TRUE -> A MATCH WAS FOUND, STOP LOOKING
             if (matchedTheCriteria == 0) break;

           }

           //IF AN EXACT MATCH TO THE CRITERIA WAS FOUND, STOP LOOKING, PUT THE DATA IN THE SHEET & MOVE
           //TO THE NEXT LOOKUP
           if (matchedTheCriteria == 0) { 
               found = true;
               matchFoundWriteResults(outputRange,dataFields,controlFields,dataRange,x);
               //NEW FEATURE 8-16-2019: CREATE FILE OF MARC RECORDS
               if (emailAddress != null && emailAddress != "") {
                 var xmlText = XmlService.getPrettyFormat().format(listOfRecords[y].getChild("recordData",nsp).getChild("record",slimNsp));
                 fileToReturn = fileToReturn + xmlText;
               }
               //END NEW FEATURE
               break;
           }

       }
       //ADDED 8/22/2019
       //IF FOUND == FALSE & THE SEARCH FOUND AT LEAST ONE RECORD - USE THE TOP RECORD
       //IT WILL BE THE ONE WITH THE LARGEST NUMBER OF HOLDINGS
       if (listOfRecords.length > 0 && found == false) {
         var dataFields = listOfRecords[0].getChild("recordData",nsp).getChild("record",slimNsp).getChildren("datafield",slimNsp);
         var controlFields = listOfRecords[0].getChild("recordData",nsp).getChild("record",slimNsp).getChildren("controlfield",slimNsp);
         matchFoundWriteResults(outputRange,dataFields,controlFields,dataRange,x);
         //NEW FEATURE 8-16-2019: CREATE FILE OF MARC RECORDS
         if (emailAddress != null && emailAddress != "") {
            var xmlText = XmlService.getPrettyFormat().format(listOfRecords[0].getChild("recordData",nsp).getChild("record",slimNsp));
            fileToReturn = fileToReturn + xmlText;
         }
       }
       
   }
   
   //NEW FEATURE 8-16-2019: CREATE FILE OF MARC RECORDS
   if (emailAddress != null && emailAddress != "") {
     fileToReturn = fileToReturn + '</marc:collection>';
     //var mydoc = DriveApp.getRootFolder().createFile('marc-' + rand + '.xml', fileToReturn);
     var blob = Utilities.newBlob(fileToReturn, 'text/xml', 'marc.xml');
     MailApp.sendEmail(emailAddress, 'MARC File Attached', '', {
      name: 'Automatic Emailer Script',
      attachments: [blob]
     });
   }
   //END NEW FEATURE
   ui.alert("done");

 }

  //https://stackoverflow.com/questions/2686855/is-there-a-javascript-function-that-can-pad-a-string-to-get-to-a-determined-leng
  function pad(width, string, padding) { 
     return (width <= string.length) ? string : pad(width, padding + string, padding)
  }



  function matchFoundWriteResults(outputRange,dataFields,controlFields,dataRange,rowNumber) {
  

               var outPutSettingsRows = outputRange.getNumRows();
               //ui.alert(outPutSettingsRows);
               for (var b = 1; b <= outPutSettingsRows; b++) {
                 var field = outputRange.getCell(b, 1).getValue();
                 var outputCol = outputRange.getCell(b, 2).getValue();
                 if (field == null || field == "") continue;
                 if (outputCol == null || outputCol == "") continue;
                 //SPLIT BY : - IN CASE OF MULTIPLE CHOICES OF FIELDS
                 var fieldArray = field.split(":");
                 //ui.alert(fieldArray);
                 for (var d = 0; d <= fieldArray.length; d ++) {
                   //for each field in the cell separated by :
                   var valueToPrint = null;
                   //ui.alert("LOOKING FOR FIELD..." + fieldArray[d]);
                   var fieldIndicator = fieldArray[d];
                   if (fieldIndicator == null || fieldIndicator == "") continue;
                   //check for the existance of a subfield (e.g. 040$b)
                   var fieldSubFieldArray = fieldIndicator.split("$");
                   if (fieldSubFieldArray.length > 1) {
                      valueToPrint = getValueForFieldSubField(dataFields,fieldSubFieldArray[0],fieldSubFieldArray[1]);
                   }
                   else {
                      valueToPrint = getValueForField(dataFields,controlFields,fieldSubFieldArray[0]);
                   }
                   if (valueToPrint != null) {
                     var valueToPrint = valueToPrint.replace(/\n/g,"");// replace all \n with ''
                     //ui.alert("value to print: " + valueToPrint);
                     dataRange.getCell(rowNumber, outputCol).setValue(valueToPrint);
                     break;
                   }

                 }
               }

   }



  

  function getTabs() {
    var ui = SpreadsheetApp.getUi();
    var out = new Array();
    var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    for (var i=0 ; i<sheets.length ; i++) {
       out.push( [ sheets[i].getName() ] );
    }
    return out;
  }
  
  
  
  function getStoredAPIKey() {
     return PropertiesService.getUserProperties().getProperty('apiKey')
  }
  
  function getStoredEmailAddress() {
    return PropertiesService.getUserProperties().getProperty('emailAddress')
  }
  
  
  
  //CALL THE WORLDCAT API SPECIFICALLY LOOKING FOR LOCAL HOLDINGS
  function findLocalRecord(x,dataRange,searchCriteria,localCode,outputRange) {
       var ui = SpreadsheetApp.getUi();
       var apiKey = PropertiesService.getUserProperties().getProperty('apiKey');
       var emailAddress = PropertiesService.getUserProperties().getProperty('emailAddress');
       var url = "http://worldcat.org/webservices/catalog/search/sru?query=" + searchCriteria + " AND srw.li=" + localCode + "&wskey=" + apiKey + "&recordSchema=info:srw/schema/1/marcxml-v1.1&frbrGrouping=off&servicelevel=full";
       var options = {
         "method" : "GET",
         "headers" : {
           "x-api-key" : apiKey
         }
       };
       var xml = UrlFetchApp.fetch(url,options).getContentText();
       var document = XmlService.parse(xml);
       var nsp = XmlService.getNamespace('http://www.loc.gov/zing/srw/');
       var slimNsp = XmlService.getNamespace('http://www.loc.gov/MARC21/slim'); 
    
       var root = document.getRootElement();
       var test = root.getChild("numberOfRecords",nsp).getValue();
       if (test == "1") {
         //FOUND A LOCAL RECORD, WRITE RESULTS TO THE SPREADSHEET
         var records = root.getChild("records",nsp);
         var listOfRecords = records.getChildren();
         var dataFields = listOfRecords[0].getChild("recordData",nsp).getChild("record",slimNsp).getChildren("datafield",slimNsp);
         var controlFields = listOfRecords[0].getChild("recordData",nsp).getChild("record",slimNsp).getChildren("controlfield",slimNsp);  
         matchFoundWriteResults(outputRange,dataFields,controlFields,dataRange,x);
         //NEW FEATURE 8-16-2019: CREATE FILE OF MARC RECORDS
         //IF THEY'VE INCLUDED AN EMAIL, INCLUDE THE RECORD TO EMAIL
         if (emailAddress != null && emailAddress != "") {
           var xmlText = XmlService.getPrettyFormat().format(listOfRecords[0].getChild("recordData",nsp).getChild("record",slimNsp));
           fileToReturn = fileToReturn + xmlText;
         }
         //END NEW FEATURE
         return true;
       }
       return false;
  }
  
  //CALL THE WORLDCAT API 
  //RESULTS WILL BE SORTED BY NUMBER
  //OF HOLDINGS LIBRARIES
  //SEARCH CRITERIA IS EITHER BY ISSN OR LCCN
  function findRecord(searchCriteria) {
     var ui = SpreadsheetApp.getUi();
      var apiKey = PropertiesService.getUserProperties().getProperty('apiKey');
      var url = "http://worldcat.org/webservices/catalog/search/sru?query=" + searchCriteria + "&wskey=" + apiKey + "&recordSchema=info:srw/schema/1/marcxml-v1.1&frbrGrouping=off&servicelevel=full&sortKeys=LibraryCount,,0&frbrGrouping=off";
      var options = {
         "method" : "GET",
         "headers" : {
           "x-api-key" : apiKey
         }
      }
      //ui.alert(url);
      var xml = UrlFetchApp.fetch(url,options).getContentText();
      var document = XmlService.parse(xml);
      return document;
  }
  
  

									
   
