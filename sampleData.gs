//WHEN THE "Click to initialize sample tabs" BUTTON IS CLICKED
//THIS FUNCTION INSERTS TABS THAT CONTAIN SAMPLE SEARCH DATA &
//SAMPLE SEARCH CRITERIA
function initSampleData(form) {
    var rand = Math.floor((Math.random() * 1000) + 1);
    var ui = SpreadsheetApp.getUi();
    var spreadsheet = SpreadsheetApp.getActive();
    settingsSheet = spreadsheet.insertSheet();
    try {
      settingsSheet.setName("Sample Search Criteria");
    }
    catch(err) {
      //IN CASE TAB WITH THIS NAME ALREADY EXISTS
      settingsSheet.setName("Search Criteria Sample" + rand);
    }
    var outputRange = settingsSheet.getRange(1, 1, 37, 8);
    
    
    //ADD THE SEARCH CRITERIA
    outputRange.getCell(1,1).setValue("search:").setFontWeight("bold");
    outputRange.getCell(2,1).setValue("holdings=LYU");
    outputRange.getCell(2,2).setValue("<-- set your OCLC Symbol here or clear this cell if you don't want it to look for your own holdings");
    
    
    outputRange.getCell(3,1).setValue("040=dlc").setBackground("#D6EAF8");
    outputRange.getCell(3,2).setValue("040$b=eng").setBackground("#D6EAF8");
    outputRange.getCell(3,3).setValue("336$b=txt").setBackground("#D6EAF8");
    outputRange.getCell(3,4).setValue("337$b=n").setBackground("#D6EAF8");
    outputRange.getCell(3,5).setValue("338$b=nc").setBackground("#D6EAF8");
    outputRange.getCell(3,6).setBackground("#D6EAF8");
    outputRange.getCell(3,7).setBackground("#D6EAF8");
    outputRange.getCell(3,8).setValue("<-- all must be true for a 'match'");
    
    outputRange.getCell(4,1).setValue("042=pcc").setBackground("#D6EAF8");
    outputRange.getCell(4,2).setValue("040$b=eng").setBackground("#D6EAF8");
    outputRange.getCell(4,3).setValue("336$b=txt").setBackground("#D6EAF8");
    outputRange.getCell(4,4).setValue("337$b=n").setBackground("#D6EAF8");
    outputRange.getCell(4,5).setValue("338$b=nc").setBackground("#D6EAF8");
    outputRange.getCell(4,6).setBackground("#D6EAF8");
    outputRange.getCell(4,7).setBackground("#D6EAF8");
    outputRange.getCell(4,8).setValue("<-- all must be true for a 'match'");
    
    outputRange.getCell(5,1).setValue("336$b=txt").setBackground("#D6EAF8");
    outputRange.getCell(5,2).setValue("337$b=c").setBackground("#D6EAF8");
    outputRange.getCell(5,3).setValue("338$b=cr").setBackground("#D6EAF8");
    outputRange.getCell(5,4).setBackground("#D6EAF8");
    outputRange.getCell(5,5).setBackground("#D6EAF8");
    outputRange.getCell(5,6).setBackground("#D6EAF8");
    outputRange.getCell(5,7).setBackground("#D6EAF8");
    outputRange.getCell(5,8).setValue("<-- all must be true for a 'match'");
    
    outputRange.getCell(6,1).setBackground("#D6EAF8");
    outputRange.getCell(6,2).setBackground("#D6EAF8");
    outputRange.getCell(6,3).setBackground("#D6EAF8");
    outputRange.getCell(6,4).setBackground("#D6EAF8");
    outputRange.getCell(6,5).setBackground("#D6EAF8");
    outputRange.getCell(6,6).setBackground("#D6EAF8");
    outputRange.getCell(6,7).setBackground("#D6EAF8");

    
    outputRange.getCell(7,1).setBackground("#D6EAF8");
    outputRange.getCell(7,2).setBackground("#D6EAF8");
    outputRange.getCell(7,3).setBackground("#D6EAF8");
    outputRange.getCell(7,4).setBackground("#D6EAF8");
    outputRange.getCell(7,5).setBackground("#D6EAF8");
    outputRange.getCell(7,6).setBackground("#D6EAF8");
    outputRange.getCell(7,7).setBackground("#D6EAF8");

    
    
    outputRange.getCell(8,1).setBackground("#D6EAF8");
    outputRange.getCell(8,2).setBackground("#D6EAF8");
    outputRange.getCell(8,3).setBackground("#D6EAF8");
    outputRange.getCell(8,4).setBackground("#D6EAF8");
    outputRange.getCell(8,5).setBackground("#D6EAF8");
    outputRange.getCell(8,6).setBackground("#D6EAF8");
    outputRange.getCell(8,7).setBackground("#D6EAF8");

    
    //ADD THE RETURN CRITERIA
    outputRange.getCell(11,1).setValue("fields:").setFontWeight("bold");
    outputRange.getCell(11,2).setValue("starting column:").setFontWeight("bold");
    
    outputRange.getCell(12,1).setNumberFormat("@").setValue("001").setBackground("#D6EAF8");
    outputRange.getCell(12,2).setNumberFormat("@").setBackground("#D6EAF8").setValue("4");
    outputRange.getCell(12,3).setValue("Cols 1,2,3 are used for ISBN, LCCN, 'local record found indicator'.  082:092:050 notation will print first field found.");
    
    outputRange.getCell(13,1).setValue("245$a").setBackground("#D6EAF8");
    //outputRange.getCell(13,2).setNumberFormat("@").setBackground("#D6EAF8").setValue("5");
    
    outputRange.getCell(14,1).setValue("245$b").setBackground("#D6EAF8");
    //outputRange.getCell(14,2).setNumberFormat("@").setBackground("#D6EAF8").setValue("6");
    
    outputRange.getCell(15,1).setValue("245$c").setBackground("#D6EAF8");
    //outputRange.getCell(15,2).setNumberFormat("@").setBackground("#D6EAF8").setValue("7");
    
    outputRange.getCell(16,1).setNumberFormat("@").setBackground("#D6EAF8").setValue("082:092:050");
    //outputRange.getCell(16,2).setNumberFormat("@").setBackground("#D6EAF8").setValue("8");
    //outputRange.getCell(16,3).setValue("<-- will print first field found");
    
    outputRange.getCell(17,1).setNumberFormat("@").setBackground("#D6EAF8").setValue("050:090");
    //outputRange.getCell(17,2).setNumberFormat("@").setBackground("#D6EAF8").setValue("9");
    
    outputRange.getCell(18,1).setNumberFormat("@").setBackground("#D6EAF8");
    //outputRange.getCell(18,2).setNumberFormat("@").setBackground("#D6EAF8");
    
    outputRange.getCell(19,1).setNumberFormat("@").setBackground("#D6EAF8");
    //outputRange.getCell(19,2).setNumberFormat("@").setBackground("#D6EAF8");
    
    outputRange.getCell(20,1).setNumberFormat("@").setBackground("#D6EAF8");
    //outputRange.getCell(20,2).setNumberFormat("@").setBackground("#D6EAF8");
    
    outputRange.getCell(21,1).setNumberFormat("@").setBackground("#D6EAF8");
    //outputRange.getCell(21,2).setNumberFormat("@").setBackground("#D6EAF8");
    
    outputRange.getCell(22,1).setNumberFormat("@").setBackground("#D6EAF8");
    //outputRange.getCell(22,2).setNumberFormat("@").setBackground("#D6EAF8");
    
    outputRange.getCell(23,1).setNumberFormat("@").setBackground("#D6EAF8");
    //outputRange.getCell(23,2).setNumberFormat("@").setBackground("#D6EAF8");
    
    outputRange.getCell(24,1).setNumberFormat("@").setBackground("#D6EAF8");
    //outputRange.getCell(24,2).setNumberFormat("@").setBackground("#D6EAF8");
    
    outputRange.getCell(25,1).setNumberFormat("@").setBackground("#D6EAF8");
    //outputRange.getCell(25,2).setNumberFormat("@").setBackground("#D6EAF8");
    
    outputRange.getCell(26,1).setNumberFormat("@").setBackground("#D6EAF8");
    //outputRange.getCell(26,2).setNumberFormat("@").setBackground("#D6EAF8");
    
    outputRange.getCell(27,1).setNumberFormat("@").setBackground("#D6EAF8");
    //outputRange.getCell(27,2).setNumberFormat("@").setBackground("#D6EAF8");
    
    outputRange.getCell(28,1).setNumberFormat("@").setBackground("#D6EAF8");
    //outputRange.getCell(28,2).setNumberFormat("@").setBackground("#D6EAF8");
    
    outputRange.getCell(29,1).setNumberFormat("@").setBackground("#D6EAF8");
    //outputRange.getCell(29,2).setNumberFormat("@").setBackground("#D6EAF8");
    
    outputRange.getCell(30,1).setNumberFormat("@").setBackground("#D6EAF8");
    //outputRange.getCell(30,2).setNumberFormat("@").setBackground("#D6EAF8");
    
    outputRange.getCell(31,1).setNumberFormat("@").setBackground("#D6EAF8");
    //outputRange.getCell(31,2).setNumberFormat("@").setBackground("#D6EAF8");
    
    outputRange.getCell(32,1).setNumberFormat("@").setBackground("#D6EAF8");
    //outputRange.getCell(32,2).setNumberFormat("@").setBackground("#D6EAF8");
    
    outputRange.getCell(33,1).setNumberFormat("@").setBackground("#D6EAF8");
    //outputRange.getCell(33,2).setNumberFormat("@").setBackground("#D6EAF8");
    
    outputRange.getCell(34,1).setNumberFormat("@").setBackground("#D6EAF8");
    //outputRange.getCell(34,2).setNumberFormat("@").setBackground("#D6EAF8");
    
    outputRange.getCell(35,1).setNumberFormat("@").setBackground("#D6EAF8");
    //outputRange.getCell(35,2).setNumberFormat("@").setBackground("#D6EAF8");
    
    outputRange.getCell(36,1).setNumberFormat("@").setBackground("#D6EAF8");
    //outputRange.getCell(36,2).setNumberFormat("@").setBackground("#D6EAF8");
    
    
        
    //CREATE SHEET WITH SAMPLE SEARCHES
    settingsSheet = spreadsheet.insertSheet();
    try {
      settingsSheet.setName("Sample Searches");
    }
    catch(err) {
      //IN CASE TAB WITH THIS NAME ALREADY EXISTS
      settingsSheet.setName("Sample Searches" + rand);
    }
    var outputRange = settingsSheet.getRange(1, 1, 4, 4);
    
    outputRange.getCell(1,1).setValue("ISBN").setFontWeight("bold");
    outputRange.getCell(1,2).setValue("LCCN").setFontWeight("bold")
    outputRange.getCell(1,3).setValue("local record indicator").setFontWeight("bold");

    outputRange.getCell(1,4).setValue("<--script will populate this column");
    outputRange.getCell(2,1).setNumberFormat("@").setValue("9781849763721");
    outputRange.getCell(3,1).setNumberFormat("@").setValue("9780297178545");
    outputRange.getCell(4,2).setNumberFormat("@").setValue("66-25377");
    
    
    
    return getTabs();
    

  }
