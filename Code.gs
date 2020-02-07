
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
   
      
   
   
   var startingRow = form.rowNumber;
   
   //SETUP SHEETS/TABS/RANGES TO READ FROM/WRITE TO
   var settingsTabName = form.tabSelection;
   var dataTabName = form.searchForTab;
 
   var spreadsheet = SpreadsheetApp.getActive();
   var dataSheet = spreadsheet.getSheetByName(dataTabName);
   SpreadsheetApp.setActiveSheet(dataSheet);
   var settingsSheet = spreadsheet.getSheetByName(settingsTabName);
   var settingsRange = settingsSheet.getDataRange();
   var outputItemsRequested = settingsSheet.getLastRow() - 11;
   Logger.log("output items: "  + outputItemsRequested);
   var outputRange = settingsSheet.getRange(12,1,outputItemsRequested,2);
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
   var x = 1;
   if (startingRow != null && startingRow != "") x = startingRow -1;
   for (x; x <= numRows; x++) {
        var isbnCell = dataRange.getCell(x,1);
        var lccnCell = dataRange.getCell(x,2)
        var searchCriteria = null;
        //IF ISBN COL IS BLANK, SEARCH WILL USE THE
        //VALUE IN THE LCCN COL.
        //IF THE ROW CONTAINS NEITHER IT MOVES TO THE NEXT ROW
        if (!isbnCell.isBlank()) {
          var isbn = isbnCell.getValue();
          if (isbn.length < 10) isbn = pad(10,isbn,0);
          searchCriteria = "srw.bn=" + "%22" + isbn + "%22";
        }
        else if (!lccnCell.isBlank()) {
          searchCriteria = "srw.dn=" + "%22" + lccnCell.getValue() + "%22";
        }
        else {
          continue;
        }
        if (searchCriteria == null) continue;
        
        
        //IF SEARCH FOR LOCAL HOLDINGS IS REQUIRED, CALL THE API INLCUDING THE
        //OCLC SYMBOL
        if (checkLocal) {
          try {
            var foundLocalRecord = findLocalRecord(x,dataRange,searchCriteria,checkLocalCode,outputRange,dataSheet);
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
         ui.alert(err);
         return;
       }
       
       
       var nsp = XmlService.getNamespace('http://www.loc.gov/zing/srw/');
       var slimNsp = XmlService.getNamespace('http://www.loc.gov/MARC21/slim'); 
     
       var root = searchResults.getRootElement();
       //var test = root.getChild("numberOfRecords",nsp).getValue();
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
               matchFoundWriteResults(outputRange,dataFields,controlFields,dataRange,x,dataSheet);
               break;
           }

       }
       //ADDED 8/22/2019
       //IF FOUND == FALSE & THE SEARCH FOUND AT LEAST ONE RECORD - USE THE TOP RECORD
       //IT WILL BE THE ONE WITH THE LARGEST NUMBER OF HOLDINGS
       if (listOfRecords.length > 0 && found == false) {
         var dataFields = listOfRecords[0].getChild("recordData",nsp).getChild("record",slimNsp).getChildren("datafield",slimNsp);
         var controlFields = listOfRecords[0].getChild("recordData",nsp).getChild("record",slimNsp).getChildren("controlfield",slimNsp);
         matchFoundWriteResults(outputRange,dataFields,controlFields,dataRange,x,dataSheet);
       }
       
   }

   spreadsheet.toast("complete");

 }

  //https://stackoverflow.com/questions/2686855/is-there-a-javascript-function-that-can-pad-a-string-to-get-to-a-determined-leng
  function pad(width, string, padding) { 
     return (width <= string.length) ? string : pad(width, padding + string, padding)
  }



  function matchFoundWriteResults(outputRange,dataFields,controlFields,dataRange,rowNumber,dataSheet) {
  
               var ui = SpreadsheetApp.getUi();
               var colors = new Array(1);
               colors[0] = new Array(outputRange.getNumRows());
               var outPutSettingsRows = outputRange.getNumRows();
               Logger.log("***> " + lastRowInRange(outputRange));
               Logger.log("--->" + outPutSettingsRows);
               var xx = 0;
               var yy = 0;
               var outputColStart = outputRange.getCell(1, 2).getValue();
               for (var b = 1; b <= outPutSettingsRows; b++) {
                 var field = outputRange.getCell(b, 1).getValue();
                 var outputCol = outputRange.getCell(b, 2).getValue();
                 if (field == null || field == "") continue;
                 //if (outputCol == null || outputCol == "") continue;
                 //SPLIT BY : - IN CASE OF MULTIPLE CHOICES OF FIELDS
                 var fieldArray = field.split(":");
                 for (var d = 0; d <= fieldArray.length; d ++) {
                   //for each field in the cell separated by :
                   var valueToPrint = null;
                   var fieldIndicator = fieldArray[d];
                   if (fieldIndicator == null || fieldIndicator == "") continue;
                   var fieldSubFieldArray = fieldIndicator.split("$");
                   if (fieldSubFieldArray.length > 1) {
                      valueToPrint = getValueForFieldSubField(dataFields,fieldSubFieldArray[0],fieldSubFieldArray[1]);
                   }
                   else {
                      valueToPrint = getValueForField(dataFields,controlFields,fieldSubFieldArray[0]);
                   }
                   if (valueToPrint != null) {
                     var valueToPrint = valueToPrint.replace(/\n/g,"");// replace all \n with ''
                     //dataRange.getCell(rowNumber, outputCol).setValue(valueToPrint);
                     colors[xx][yy] = valueToPrint;
                     break;
                   }
                   else {
                     colors[xx][yy] = "";
                   }

                 }
                 yy++;
               }
               xx++;
               //WRITE RESULTS
               var oneRowDataRange = dataSheet.getRange(rowNumber+1,outputColStart,1,outputRange.getNumRows());
               oneRowDataRange.setValues(colors);
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
  
  function getStoredOclcColNumber() {
    return PropertiesService.getUserProperties().getProperty('oclcColNumber')
  }
  
  
  
  //CALL THE WORLDCAT API SPECIFICALLY LOOKING FOR LOCAL HOLDINGS
  function findLocalRecord(x,dataRange,searchCriteria,localCode,outputRange,dataSheet) {
       var ui = SpreadsheetApp.getUi();
       var apiKey = PropertiesService.getUserProperties().getProperty('apiKey');
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
         matchFoundWriteResults(outputRange,dataFields,controlFields,dataRange,x,dataSheet);
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
  
  function lastRowInRange(range) {
    var lastrow = range.getLastRow() - 1;
    var values = range.getValue();
    while (lastrow > -1 && values[lastrow]) {
      lastRow--;
    }
    if (lastrow == -1) {
      return "Empty Column";
    } else {
      return lastrow + 1;
    }
    
  }
  
  
  function sendMARCRecordByEmail(form) {
     
     var fileToReturn = '<marc:collection xmlns:marc="http://www.loc.gov/MARC21/slim">';
     var listOfRecordIds = [];
     
     var ui = SpreadsheetApp.getUi();
     var emailAddress = form.emailAddress;
     PropertiesService.getUserProperties().setProperty('emailAddress', emailAddress);

     
     //MAKE SURE THE OCLC API KEY HAS BEEN ENTERED
     var apiKey = form.apiKey;
     if (apiKey == null || apiKey == "") {
       ui.alert("OCLC API Key is Required");
       return;
     }
     
     //MAKE SURE THE EMAIL ADDRESS HAS BEEN ENTERED
     var emailAddress = form.emailAddress;
     if (emailAddress == null || emailAddress == "") {
        ui.alert("email is required to send MARC record");
        return;
     }
     
     //MAKE SURE THE EMAIL ADDRESS HAS BEEN ENTERED
     var oclcColNumber = form.oclcNumber;
     if (oclcColNumber == null || oclcColNumber == "") {
        ui.alert("Indicate which column number the OCLC number is in.");
        return;
     }
     
    
     PropertiesService.getUserProperties().setProperty('emailAddress', emailAddress);
     PropertiesService.getUserProperties().setProperty('oclcColNumber', oclcColNumber);
     
     //OPTIONAL - START ROW
     var startingRow = form.rowNumberForEmail;
     
     //SETUP SHEETS/TABS/RANGES TO READ FROM/WRITE TO:
     
     var settingsTabName = form.tabSelection;
     var dataTabName = form.searchForTab;
   
     var spreadsheet = SpreadsheetApp.getActive();
     spreadsheet.toast("starting...");
     var dataSheet = spreadsheet.getSheetByName(dataTabName);
     SpreadsheetApp.setActiveSheet(dataSheet);
     var lastRow = dataSheet.getLastRow();
     var dataRange = dataSheet.getRange(2, 1, lastRow , 100)
     var numRows = dataRange.getNumRows();
     var x = 1;
     if (startingRow != null && startingRow != "") x = startingRow -1;
     for (x; x <= numRows; x++) {
        //GET MARC RECORD USING THE OCLC NUMBER IN COL indciated in the oclcColNumber field
        var oclcNumberCell = dataRange.getCell(x,oclcColNumber);
        if (oclcNumberCell.isBlank()) continue;
        //IF THE OCLC NUMBER HAS ALREADY BEEN LOOKED UP, IT'S A DUPLICATE - SKIP IT:
        if (listOfRecordIds.indexOf(oclcNumberCell.getValue()) >= 0) continue;
        listOfRecordIds.push(oclcNumberCell.getValue());
        oclcNumberCell.setBackground('#ffffcc');
        //GET THE MARC RECORD BY OCLC NUMBER
        var url = "http://www.worldcat.org/webservices/catalog/content/" + oclcNumberCell.getValue() + "?wskey=" + apiKey + "&recordSchema=info:srw/schema/1/marcxml-v1.1&frbrGrouping=off&servicelevel=full&sortKeys=LibraryCount,,0&frbrGrouping=off";
        var options = {
           "method" : "GET",
           "headers" : {
             "x-api-key" : apiKey
           }
        }
        try {
         var xml = UrlFetchApp.fetch(url,options).getContentText();
        }
        catch(err) {
            ui.alert("Communication with API failed.  Please check your API key.");
            ui.alert(err);
            return;
       }
       var document = XmlService.parse(xml);
       var root = document.getRootElement();
       if (root == null) continue;
       
       //*****ADD FIELDS TO THE MARC RECORD*******************************
       //LOOP THROUGH THE FIRST ROW - ALL COLUMNS LOOKING FOR FIELDS TO ADD
       var headerColumn = dataSheet.getRange(1, 1, 1, 100);
       var listOfFieldsAddedToRecord = [];
       
       
       for (var i = 5, l = 100; i < l; i += 1) {
            Logger.log(headerColumn.getCell(1,i).getValue());
            var fieldSubfield = headerColumn.getCell(1,i).getValue();
            if (fieldSubfield.indexOf('$') > -1) {
                    var indexOfSubField = fieldSubfield.indexOf('$');
                    var subField = fieldSubfield.substring(indexOfSubField+1,10);
                    //ui.alert('a subfield exists' + v);
                    var field = fieldSubfield.substring(0,indexOfSubField);
                    //Logger.log("field - " + field + "--" + subField);
                    //GET THE VALUE FOR THE NEW FIELD
                    var fieldValue = dataRange.getCell(x,i).getValue();
                    if (dataRange.getCell(x,i).isBlank()) continue;
                    Logger.log("value - for row/col" + i +" is" + fieldValue);

                    //CREATE DATAFIELD ELEMENT 
                    var datafieldElement = XmlService.createElement("datafield");
                    datafieldElement.setAttribute("tag",field);
                    datafieldElement.setAttribute("ind1","");
                    datafieldElement.setAttribute("ind2","");
                    
                    //CREATE SUBFIELD 
                    var subfieldElement = XmlService.createElement("subfield");
                    subfieldElement.setAttribute("code",subField);
                    subfieldElement.setText(fieldValue);
                    
                    datafieldElement.addContent(subfieldElement);
                    //Logger.log(datafieldElement);
                    listOfFieldsAddedToRecord = addNewElement(listOfFieldsAddedToRecord,datafieldElement,subfieldElement,field);
             }
       }
       //ADD THE NEW ELEMENTS TO THE MARC RECORD
       
       
       //ADD ALL OF THE NEW FIELDS/SUBFIELDS FROM THIS ROW TO THE RECORD:
       for (var i = 0, l = listOfFieldsAddedToRecord.length; i < l; i += 1) {
         root.addContent(listOfFieldsAddedToRecord[i]);
       }
       
       
       
       //CHECK FOR MORE FIELDS TO ADD IN ROWS BELOW:
       //THIS IS A LITTLE UGLY
       var anotherRecord = true;
       var tempx = x + 1;
       while (anotherRecord && tempx <=dataRange.getNumRows()) {
         var nextOclcNumber = dataRange.getCell(tempx,oclcColNumber);
         var issn = dataRange.getCell(tempx,1);
         var lccn = dataRange.getCell(tempx,2);
         //MAKE SURE THEY WERE NOT LOOKING FOR ANOTHER RECORD THAT WAS NOT FOUND
         if (nextOclcNumber.isBlank() && issn.isBlank() && lccn.isBlank()) { 
             //Logger.log("^^^^^^^^^^^^^^^^found another row");
             var listOfFieldsAddedToRecord = [];
       
             //FOR EACH COLUMN IN THIS ROW
             for (var i = 5, l = 100; i < l; i += 1) {
                  Logger.log(headerColumn.getCell(1,i).getValue());
                  var fieldSubfield = headerColumn.getCell(1,i).getValue();
                  if (fieldSubfield.indexOf('$') > -1) {
                          var indexOfSubField = fieldSubfield.indexOf('$');
                          var subField = fieldSubfield.substring(indexOfSubField+1,10);
                          var field = fieldSubfield.substring(0,indexOfSubField);
                          //GET THE VALUE FOR THE NEW FIELD
                          var fieldValue = dataRange.getCell(tempx,i).getValue();
                          if (dataRange.getCell(tempx,i).isBlank()) continue;
                          
                          var datafieldElement = XmlService.createElement("datafield");
                          datafieldElement.setAttribute("tag",field);
                          datafieldElement.setAttribute("ind1","");
                          datafieldElement.setAttribute("ind2","");
                          
                          var subfieldElement = XmlService.createElement("subfield");
                          subfieldElement.setAttribute("code",subField);
                          subfieldElement.setText(fieldValue);
                          
                          datafieldElement.addContent(subfieldElement);
                          listOfFieldsAddedToRecord = addNewElement(listOfFieldsAddedToRecord,datafieldElement,subfieldElement,field);
                   }
             }
             //ADD ALL OF THE NEW FIELDS/SUBFIELDS FROM THIS ROW TO THE RECORD:          
             for (var i = 0, l = listOfFieldsAddedToRecord.length; i < l; i += 1) {
               root.addContent(listOfFieldsAddedToRecord[i]);
             }
             
             //CHECK FOR MORE FIELDS *FOR THIS RECORD*
             tempx = tempx + 1;
         }
         else {
           //MOVING ON TO THE NEXT OCLC NUMBER
           anotherRecord = false;
           //Logger.log("..........moving on...........");
           var listOfFieldsAddedToRecord = [];
         
         }
       }
       
       var xmlText = XmlService.getPrettyFormat().format(root);
       fileToReturn = fileToReturn + xmlText;
         
    }
    fileToReturn = fileToReturn + '</marc:collection>';
    var blob = Utilities.newBlob(fileToReturn, 'text/xml', 'marc.xml');
    MailApp.sendEmail(emailAddress, 'MARC File Attached', '', {
        name: 'Automatic Emailer Script',
        attachments: [blob]
    });
    spreadsheet.toast("done! email sent to: " + emailAddress);
 }  
 
 
 
 function addNewElement(collectionOfNewFields,newElement,subfieldElement,field) {
   var ui = SpreadsheetApp.getUi();
   var existingField = getDataField(collectionOfNewFields,field);
   if (existingField == null) {
     collectionOfNewFields.push(newElement);
   }
   
   //ELEMENT EXISTS SO ADD SUBFIELD (e.g b)
   else {
   
     for (var z = 0; z < collectionOfNewFields.length; z++) {
             var tagAttribute = collectionOfNewFields[z].getAttribute("tag");
             if (tagAttribute != null && tagAttribute.getValue() == field) {  //e.g. 040
                 
                 var code = subfieldElement.getAttribute("code").getValue();
                 var v = subfieldElement.getText();
                 //NOTE: WOULDN'T LET ME ADD THE SUBFIELD IF IT WAS PASSED IN AS AN ARG
                 //ONLY LET ME IF I CREATED THE ELEMENT IN THIS FUNCTION?
                 var subfield = XmlService.createElement('subfield');
                 subfield.setAttribute("code", code);
                 subfield.setText(v);
                 
                 collectionOfNewFields[z].addContent(subfield);
             }
          }

   }
   
   //COLLECTION OF FIELDS W/THE NEW FIELD ADDED
   return collectionOfNewFields;
 }
  

