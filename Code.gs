function onOpen() {
  SpreadsheetApp.getUi() 
      .createMenu('Custom Menu')
      .addItem('Launch', 'showSidebar')
      .addToUi();
}


function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('sidebar')
      .setTitle('MARC Search & Select:')
      .setWidth(500);
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .showSidebar(html); 
  //TODO - Add code that will set up a tab that will contain
  //sample configurations (if one does not already exist).
  //Maybe add a button that will allow the user to indicate
  //they want sample config data inserted into a tab?
}


//GET ALL BUT THE FIRST TAB
function getTabs() {
  var ui = SpreadsheetApp.getUi();
  var out = new Array();
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (var i=1 ; i<sheets.length ; i++) {
    out.push( [ sheets[i].getName() ] );
  }
  return out;
}
  
  
  
//THIS METHOD
//LOOPS THROUGH EACH RECORD 
//CHECKS TO SEE IF WE HAVE A LOCAL RECORD (IF CONFIGURED TO CHECK)
//OTHERWISE LOOKS UP THE IDENTIFIER (ISBN/LCCN)
//MATCHES RECORDS FOUND AGAINST THE ROWS AND COLUMNS IN THE SEARCH CRITERIA
//IF IT MATCHES ALL OF THE SEARCH CRITERIA IN ONE ROW - IT IS CONSIDERED A MATCH
//THE RECORD DATA IS WRITTEN TO THE SPREADSHEET (USING THE FIELDS & COLUMNS DEFINED IN THE CONFIG)
  
  
function startLookup(form) {
   var ui = SpreadsheetApp.getUi();
   ui.alert('start...');
   var apiKey = form.apiKey;
   PropertiesService.getScriptProperties().setProperty('apiKey', apiKey); 
   
   var settingsTabName = form.tabSelection;
  
   //AS OF NOW - THE 'ACTIVE' TAB
   //IS EXPECTED TO CONTAIN THE LOOKUP DATA & WILL
   //RECEIVE THE RESULTS
   var spreadsheet = SpreadsheetApp.getActive();
   var dataSheet = spreadsheet.getActiveSheet();
   
   //THE USER SELECTS WHICH 'TAB' TO USE FOR THE CONFIGURATION OPTIONS
   var settingsSheet = spreadsheet.getSheetByName(settingsTabName);
   var settingsRange = settingsSheet.getDataRange();
   var outputRange = settingsSheet.getRange(12,1,8,2);
   
   //DO THEY WANT TO CHECK FOR LOCAL HOLDINGS?
   //THIS WILL BE LOCATED IN THE SECOND ROW, FIRST COLUMN
   var checkLocalHoldings = settingsRange.getCell(2, 1).getDisplayValue();
   
   //INIT VARIABLES
   var checkLocal = false;
   var checkLocalCode = "";
   
   //IF THE CONFIGURATION TAB IS SET TO SEARCH FOR LOCAL HOLDINGS
   //GET THE OCLC CODE
   //VALUE OF THE CELL IS EXPECTED TO CONTAIN holdings=OCLCCODE
   if (checkLocalHoldings.indexOf('holdings') > -1) {
     checkLocal = true;
     var x = checkLocalHoldings.indexOf("=");
     checkLocalCode = checkLocalHoldings.substring(x+1,checkLocalHoldings.length);
   }
   
   
   //FOR EACH ROW IN THE DATA SPREADSHEET
   var lastRow = dataSheet.getLastRow();
   var lastCol = 100; //PLACEHOLDER TODO?
   //ON THE 'DATA TAB' - START WITH THE SECOND ROW, FIRST COLUMN
   var dataRange = dataSheet.getRange(2, 1, lastRow , lastCol)
   var numRows = dataRange.getNumRows();

   for (var x = 1; x <= numRows; x++) {
        //THE FIRST COLUMN OF THE 'DATA' TAB WILL CONTAIN AN ISBN
        //THE SECOND COLUMN WILL CONTAIN AN LCCN (IN CASE ISBN IS NOT AVAILABLE)
        var isbnCell = dataRange.getCell(x,1);
        var lccnCell = dataRange.getCell(x,2);
        
        
       //INIT
       var searchCriteria = null;
       //MAKE SURE THE ROW CONTAINS EITHER ISBN OR LCCN
       //SET UP THE API CALL STRING FOR EITHER & SETUP SEARCH CRITERIA
       //BASED ON ISBN OR LCCN
       if (!isbnCell.isBlank()) {
         var isbn = isbnCell.getValue();
         //TESTING SHOWS THE ISBN MUST BE 10 CHARS...SO THE VALUE IS PADDED
         if (isbn.length < 10) isbn = pad(10,isbn,0);
            searchCriteria = "srw.bn=" + isbn;
       }
       else if (!lccnCell.isBlank()) {
         searchCriteria = "srw.dn=" + lccnCell.getValue();
       }
       //OTHERWISE SKIP THIS ROW - NOTHING TO LOOKUP
       else {
         continue;
       }
       if (searchCriteria == null) continue;

       
       
       if (checkLocal) {
           //x is current row number
           //FIND THE 'LOCAL' RECORD AND WRITES THE RESULTS TO THE SHEET
           var foundLocalRecord = findLocalRecord(x,dataRange,searchCriteria,checkLocalCode,outputRange);
           //IF IT FOUND A MATCH MOVE ONTO THE NEXT ROW:
           if (foundLocalRecord) {
             dataRange.getCell(x,3).setValue("local record found");
             ui.alert("local record found"); //remove
             continue;
           }
       }
       
       
       //LOCAL RECORD NOT FOUND - 
       //DO 2ND SEARCH
       var searchResults = findRecord(searchCriteria);
       
       
       var nsp = XmlService.getNamespace('http://www.loc.gov/zing/srw/');
       var slimNsp = XmlService.getNamespace('http://www.loc.gov/MARC21/slim'); 
     
       var root = searchResults.getRootElement();
       var test = root.getChild("numberOfRecords",nsp).getValue();
       var records = root.getChild("records",nsp);
       if (records == null) continue; //CONTINUE ON TO THE NEXT RECORD...API DIDN'T FIND ANYTHING
       
     
       var collectionOfSettingsRange = settingsSheet.getRange(3,1,8,6);
       
       
       var listOfRecords = records.getChildren();
       //INIT FOUND TO FALSE
       var found = false;
       
       //FIND THE BEST MATCH BASED ON THE SEARCH CRITERIA IN THE TAB:
       
       //*****************************************
       //FOR EACH RECORD FOUND FOR THIS ISBN/LCCN:
       for (var y = 0; y < listOfRecords.length; y++) {
       
           //LOOP THROUGH THE SETTINGS
           //HARDCODING 8 ROWS OF SETTINGS & 6 COL OF SETTINGS
           //B3-->G10
           //ROWS OF MATCH CRITERIA
           //***********************************
           //FOR EACH ROW OF SETTINGS (8)
           for (var i = 1; i < 8; i++) {
           
               if (collectionOfSettingsRange.getCell(i, 1).isBlank()) continue //skip this row
               var matchedTheCriteria = 0;
               //******************************
               //FOR EACH COLUMNS OF SETTINGS IN THE ROW
               //TO BE A MATCH - EVERY COLUMN IN A ROW MUST BE TRUE
               for (var e = 1; e < 6; e++) {
                 
                 var v = collectionOfSettingsRange.getCell(i, e).getValue();
                 //IF THE CELL IS EMPTY - SKIP TO THE NEXT ONE
                 if (v==null || v == "") continue; 
                 
                 //SETTING IN EACH CELL LOOKS LIKE THIS: 040$b=eng   (if looking in a subfield)  OR      040=dlc (if not looking in a subfield)
                 var indexOfValue = v.indexOf('=');
                 if (v.indexOf('$') > -1) {
                    var indexOfSubField = v.indexOf('$');
                    var subField = v.substring(indexOfSubField+1,indexOfValue);
                    var field = v.substring(0,indexOfSubField);
                 }
                 else {
                   var field = v.substring(0,indexOfValue);
                   var subField = "";
                 }
                   
                 var l = v.length;

                 var desiredValue = v.substring(indexOfValue+1,l);
                   
                 //ui.alert("LOOKING AT: " + oclcNumber + "for" + subField + "(subfield) / " + field + "/" + "should equal" + desiredValue);
                   
                 //GET THE FIELDS IN THIS RECORD  
                 var dataFields = listOfRecords[y].getChild("recordData",nsp).getChild("record",slimNsp).getChildren("datafield",slimNsp);
                 var controlFields = listOfRecords[y].getChild("recordData",nsp).getChild("record",slimNsp).getChildren("controlfield",slimNsp);  
                 var dataField = getDataField(dataFields,field); 
                 //IF THE DATAFIELD IS NOT IN THE RECORD - CONTINUE TO NEXT RECORD
                 if (dataField == null) continue;
                 var subfields = dataField.getChildren("subfield",slimNsp);

                 //BECAUSE *ALL* COLUMNS IN THIS ROW MUST EVALUATE TO TRUE
                 var valueExists = 0;
                 if (subField != null && subField != "") {
                    var valueExists = doesSubFieldContainWithin(subfields,desiredValue,subField);
                 }
                 else {
                     var valueExists = doesSubFieldContain(subfields,desiredValue);
                 }
                 matchedTheCriteria = matchedTheCriteria + valueExists;

                  
                 if (matchedTheCriteria > 0) break; //stop looking in this row because at least one column evaluated to false

               }
               
               if (matchedTheCriteria == 0) break;  
               //LOOKED IN ALL THE COLUMNS OF THE CRITERIA ROW - THEY ALL MATCHED -- STOP LOOKING
               
               
       
         }
         
         
         //IF AN EXACT MATCH TO THE CRITERIA WAS FOUND, STOP LOOKING, PUT THE DATA IN THE SHEET & MOVE
         //TO THE NEXT LOOKUP
         if (matchedTheCriteria == 0) {
               matchFoundWriteResults(outputRange,dataFields,controlFields,dataRange,x);
               break;
         }
       
       } //END LOOP CHECKING FOR LIST OF RECORDS FOUND IN SEARCH
       
   } //END LOOP OF RECORDS TO LOOKUP
   ui.alert("done");
  } //END startLookup method
  
  
  
  
  
  
  
  
  
  
  
   
   
   ///MOVE HELPER FUNCTIONS TO DIFFERENT FILE?
   
   function findRecord(searchCriteria) {
       var ui = SpreadsheetApp.getUi();
       var apiKey = PropertiesService.getScriptProperties().getProperty('apiKey');
       var url = "http://worldcat.org/webservices/catalog/search/sru?query=" + searchCriteria + "&wskey=" + apiKey + "&recordSchema=info:srw/schema/1/marcxml-v1.1&frbrGrouping=off&servicelevel=full&sortKeys=LibraryCount,,0&frbrGrouping=off";
       var options = {
       //"async": true,
       //"crossDomain": true,
         "method" : "GET",
         "headers" : {
          "x-api-key" : apiKey,
        //"cache-control": "no-cache"
       }
      }
      var xml = UrlFetchApp.fetch(url,options).getContentText();
      var document = XmlService.parse(xml);
      return document;      
  
  }
  
  
  
  
  
  
  function findLocalRecord(x,dataRange,searchCriteria,localCode,outputRange) {
      var ui = SpreadsheetApp.getUi();
      var apiKey = PropertiesService.getScriptProperties().getProperty('apiKey');
      var url = "http://worldcat.org/webservices/catalog/search/sru?query=" + searchCriteria + " AND srw.li=LYU&wskey=" + apiKey + "&recordSchema=info:srw/schema/1/marcxml-v1.1&frbrGrouping=off&servicelevel=full";
    
      var options = {
        //"async": true,
        //"crossDomain": true,
        "method" : "GET",
        "headers" : {
          "x-api-key" : apiKey,
          //"cache-control": "no-cache"
        }
      };
      var xml = UrlFetchApp.fetch(url,options).getContentText();
      var document = XmlService.parse(xml);
      var nsp = XmlService.getNamespace('http://www.loc.gov/zing/srw/');
      var slimNsp = XmlService.getNamespace('http://www.loc.gov/MARC21/slim'); 
    
      var root = document.getRootElement();
      var test = root.getChild("numberOfRecords",nsp).getValue();
      if (test == "1") {
        //FOUND A LOCAL RECORD
        var records = root.getChild("records",nsp);
        var listOfRecords = records.getChildren();
        var dataFields = listOfRecords[0].getChild("recordData",nsp).getChild("record",slimNsp).getChildren("datafield",slimNsp);
        var controlFields = listOfRecords[0].getChild("recordData",nsp).getChild("record",slimNsp).getChildren("controlfield",slimNsp);  
        
        matchFoundWriteResults(outputRange,dataFields,controlFields,dataRange,x);
     
        return true;
    }
    return false;
  }
  
  
  //https://stackoverflow.com/questions/2686855/is-there-a-javascript-function-that-can-pad-a-string-to-get-to-a-determined-leng
  function pad(width, string, padding) { 
    return (width <= string.length) ? string : pad(width, padding + string, padding)
  }



  function matchFoundWriteResults(outputRange,dataFields,controlFields,dataRange,rowNumber) {
  
      var outPutSettingsRows = outputRange.getNumRows();

      for (var b = 1; b <= outPutSettingsRows; b++) {
        var field = outputRange.getCell(b, 1).getValue();
        var outputCol = outputRange.getCell(b, 2).getValue();
        if (field == null || field == "") continue;
        if (outputCol == null || outputCol == "") continue;
        //SPLIT BY : - IN CASE OF MULTIPLE CHOICES OF FIELDS
        var fieldArray = field.split(":");
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
            dataRange.getCell(rowNumber, outputCol).setValue(valueToPrint);
            break;
          }
        }
      }
  
}


  function getValueForFieldSubField(allTheFields,field,subfield) {
    var slimNsp = XmlService.getNamespace('http://www.loc.gov/MARC21/slim'); 
    var ui = SpreadsheetApp.getUi();
    //ui.alert(field + "/" + subfield);
    var dataFields = getDataField(allTheFields,field); //040
    //ui.alert(dataField);
    if (dataFields == null) return null;
    var subfields = dataFields.getChildren("subfield",slimNsp);
    //FOR EACH SUBFIELD - IF THE 'CODE' OF THE SUBFIELD MATCHES THE SUBFIELD WE'RE
    //LOOKING FOR RETURN THE VALUE (245$a)
    for (x = 0; x <= subfields.length; x++) {
      if (subfields[x] == null) continue;
      var subfieldCode = subfields[x].getAttribute('code').getValue();
      if (subfieldCode.toLowerCase() == subfield.toLowerCase()) {
         //ui.alert("returning: " + subfields[x].getValue());
         return subfields[x].getValue();
      }
    }
  }



function getValueForField(allTheFields,controlFields,field) {
     var test = getControlField(controlFields,field);
     if (test == null) test = getDataField(allTheFields,field);
     if (test == null) return "";
     return test.getValue();
}


  function getDataField(dataFields,lookForField) {
      var ui = SpreadsheetApp.getUi();
      
      for (var z = 0; z < dataFields.length; z++) {
        var tagAttribute = dataFields[z].getAttribute("tag");
        if (tagAttribute != null && tagAttribute.getValue() == lookForField) {  //e.g. 040
          return dataFields[z];
        }
      }
      return null;
    
  }
  
  
  function doesSubFieldContain(subfields,lookForValue) {
     //var ui = SpreadsheetApp.getUi();
     //ui.alert("LOOKING FOR: " + lookForValue + " IN " + subfields);
     for (a = 0; a < subfields.length; a++) {
          var subfieldValue = subfields[a].getValue();
          if (subfieldValue.toLowerCase() == lookForValue.toLowerCase()) {
             //ui.alert("kNEW THIS WAS A MATCH");
             return 0;
          }
          else {
             //ui.alert("DIND'T THINK THIS WAS A MATCH");
          }
        }
        return 1;
  }
  
  
 function doesSubFieldContainWithin(subfields,desiredValue,subField) {
     
        for (a = 0; a < subfields.length; a++) {
            var subfieldValue = subfields[a].getValue();
            //getAttribute('term').getValue()
            var subfieldCode = subfields[a].getAttribute('code').getValue();
            //ui.alert(subfieldValue);
            //ui.alert("LOOKING FOR: " + desiredValue + " against " + subfieldValue + " and " + subField + " against" + subfieldCode);
            if (subfieldValue.toLowerCase() == desiredValue.toLowerCase() && subfieldCode.toLowerCase() == subField.toLowerCase()) {
               //ui.alert("found a match returning 0");
               return 0;
            }
        }
        //ui.alert("NO MATCH RETURNING 1");
        return 1;
  }
  
    
  function getControlField(controlFields,tagId) {
      for (var x = 0; x < controlFields.length; x++) {
         var cField = controlFields[x];
         var tagValue = cField.getAttribute("tag");
         if (tagValue != null && tagValue.getValue() == tagId) {
           return controlFields[x];
         }
      }
    return null;
  }
   
