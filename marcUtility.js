    function getOclcControlNumber(controlFields) {
             for (var x = 0; x < controlFields.length; x++) {
                 var cField = controlFields[x];
                 var oclcNumber = cField.getAttribute("tag");
                 if (oclcNumber != null && oclcNumber.getValue() == '001') {
                   return cField.getValue();
                 } 
             }
      
    }
    
    
    
   function getControlFieldValue(controlFields,fieldId) {
             for (var x = 0; x < controlFields.length; x++) {
                 var cField = controlFields[x];
                 var tagAtt = cField.getAttribute("tag");
                 if (tagAtt != null && tagAtt.getValue() == fieldId) {
                   return cField.getValue();
                 } 
             } 
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
  
  
  //RETURNS '0' IF IT FINDS A MATCH
  //RETURNS '1' IF NO MATCH FOUND
  function doesSubFieldContain(subfields,lookForValue) {
        var ui = SpreadsheetApp.getUi();
        for (a = 0; a < subfields.length; a++) {
          var subfieldValue = subfields[a].getValue();
          if (subfieldValue.toLowerCase() == lookForValue.toLowerCase()) {
             return 0;
          }
          else {
             //ui.alert("DIND'T THINK THIS WAS A MATCH");
          }
          
        }
        return 1;
  }
  
  
  //RETURNS '0' IF IT FINDS A MATCH
  //RETURNS '1' IF NO MATCH FOUND
  //ATTEMPTS TO MATCH ON FIELD/SUBFIELD AND VALUE
  function doesSubFieldContainWithin(subfields,desiredValue,subField) {
        var ui = SpreadsheetApp.getUi();
        for (a = 0; a < subfields.length; a++) {
          var subfieldValue = subfields[a].getValue();
          var subfieldCode = subfields[a].getAttribute('code').getValue();
          if (subfieldValue.toLowerCase() == desiredValue.toLowerCase() && subfieldCode.toLowerCase() == subField.toLowerCase()) {
             return 0;
          }
          
        }
        return 1;
  }
  
  
  function getValueForFieldSubField(allTheFields,field,subfield) {
    var slimNsp = XmlService.getNamespace('http://www.loc.gov/MARC21/slim'); 
    var ui = SpreadsheetApp.getUi();
    var dataFields = getDataField(allTheFields,field); //040
    if (dataFields == null) return null;
    var subfields = dataFields.getChildren("subfield",slimNsp);
    //FOR EACH SUBFIELD - IF THE 'CODE' OF THE SUBFIELD MATCHES THE SUBFIELD WE'RE
    //LOOKING FOR RETURN THE VALUE (245$a)
    for (x = 0; x <= subfields.length; x++) {
      if (subfields[x] == null) continue;
      var subfieldCode = subfields[x].getAttribute('code').getValue();
      if (subfieldCode.toLowerCase() == subfield.toLowerCase()) {
         return subfields[x].getValue();
      }
    }
  }
  
  

  function getControlField(controlFields,tagId) {
      var ui = SpreadsheetApp.getUi();
      for (var x = 0; x < controlFields.length; x++) {
         var cField = controlFields[x];
         var tagValue = cField.getAttribute("tag");
         if (tagValue != null && tagValue.getValue() == tagId) {
           return controlFields[x];
         }
      }
      return null;
  }
  

  function getValueForField(allTheFields,controlFields,field) {
       var ui = SpreadsheetApp.getUi();
       var test = getControlField(controlFields,field);
       if (test == null) test = getDataField(allTheFields,field);
       if (test == null) return "";
       return test.getValue();
  }


  
  function getCellValue(dataRange,row,col) {
    try {
      return dataRange.getCell(row,1).getValue();
    }
    catch(e) {
      //NOT A PROBLEM...JUST RETURN NULL
      return null;
    }
  }
  
