// **********************************************
// function fcnUpdateLinksIDs()
//
// This function updates all sheets Links and IDs  
// in the Config File
//
// **********************************************

function fcnUpdateLinksIDs(){
  
  // Main Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Config Spreadsheet
  var shtConfig = ss.getSheetByName('Config');
  
  // Copy Log Spreadsheet
  var shtCopyLogID = shtConfig.getRange(46, 2).getValue();
  var LinksStatus = shtConfig.getRange(46, 6).getValue();
  
  //  Update Links and ID if Status is Null
  if (shtCopyLogID != '' && LinksStatus =='') {
    var shtCopyLog = SpreadsheetApp.openById(shtCopyLogID).getSheets()[0];
  
    var CopyLogNbFiles = shtCopyLog.getRange(2, 6).getValue();
    var StartRowCopyLog = 5;
    var StartRowConfigId = 30
    var StartRowConfigLink = 17;
    
    var CopyLogVal = shtCopyLog.getRange(StartRowCopyLog, 2, CopyLogNbFiles, 3).getValues();
    
    var FileName;
    var Link;
    var Formula;
    var ConfigRowID = 'Not Found';
    var ConfigRowLk = 'Not Found';
    
    // Clear Configuration File
    shtConfig.getRange(17,2,10,1).clearContent();
    shtConfig.getRange(30,2,13,1).clearContent();
    
    // Loop through all Copied Sheets and get their Link and ID
    for (var row = 0; row < CopyLogNbFiles; row++){
      // Get File Name
      FileName = CopyLogVal[row][0];
      
      switch(FileName){
        case 'Master TCG Booster League' :
          ConfigRowID = StartRowConfigId + 0;
          ConfigRowLk = 'Not Found'; break;
        case 'Master TCG Booster League Card DB' :
          ConfigRowID = StartRowConfigId + 1; 
          ConfigRowLk = 'Not Found'; break;
        case 'Master TCG Booster League Card Pool EN' :
          ConfigRowID = StartRowConfigId + 2; 
          ConfigRowLk = StartRowConfigLink + 1; break;
        case 'Master TCG Booster League Card Pool FR' :
          ConfigRowID = StartRowConfigId + 3; 
          ConfigRowLk = StartRowConfigLink + 4; break;
        case 'Master TCG Booster League Standings EN' :
          ConfigRowID = StartRowConfigId + 4; 
          ConfigRowLk = StartRowConfigLink + 0; break;
        case 'Master TCG Booster League Standings FR' :
          ConfigRowID = StartRowConfigId + 5; 
          ConfigRowLk = StartRowConfigLink + 3; break;
        case 'Master TCG Booster League Match Reporter EN' :
          ConfigRowID = 'Not Found';
          ConfigRowLk = 'Not Found'; break;
        case 'Master TCG Booster League Match Reporter FR' :
          ConfigRowID = 'Not Found';
          ConfigRowLk = 'Not Found'; break;	
        case 'Master TCG Booster League Registration EN' :
          ConfigRowID = 'Not Found';
          ConfigRowLk = 'Not Found'; break;
        case 'Master TCG Booster League Registration FR' :
          ConfigRowID = 'Not Found';
          ConfigRowLk = 'Not Found'; break;	
        case 'Master TCG Booster League Players List & Booster' :
          ConfigRowID = StartRowConfigId + 10; 
          ConfigRowLk = StartRowConfigLink + 8; break;
        case 'Master TCG Booster League Starting Pool' :
          ConfigRowID = StartRowConfigId + 11; 
          ConfigRowLk = StartRowConfigLink + 9; break;
        default : 
          ConfigRowID = 'Not Found'; 
          ConfigRowLk = 'Not Found'; break;
      }
      
      // Set the Appropriate Sheet ID Value in the Config File
      if (ConfigRowID != 'Not Found') {
        shtConfig.getRange(ConfigRowID, 2).setValue(CopyLogVal[row][2]);
      }
      // Set tthe Appropriate Sheet ID Value in the Config File
      if (ConfigRowLk != 'Not Found') {
        // Opens Spreadsheet by ID
        Link = SpreadsheetApp.openById(CopyLogVal[row][2]).getUrl();
        Logger.log(Link); 
        
        shtConfig.getRange(ConfigRowLk, 2).setValue(Link);
      }
    }
    // Clear Standing Spreadsheets Initialization Values
    shtConfig.getRange(34,5,2,1).clearContent();
    
    // Set Links Updated when Complete
    shtConfig.getRange(46, 6).setValue('Links Updated')
  }
}

// **********************************************
// function fcnInitLeague()
//
// This function clears all data from sheets  
// to start a new league
//
// **********************************************

function fcnInitLeague(){

  // Main Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Insert ui to confirm
  var ui = SpreadsheetApp.getUi();
  var title = "RESET LEAGUE MATCH CONFIRMATION";
  var msg = "Do you really want to reset all League Match Results";
  var uiResponse = ui.alert(title, msg, ui.ButtonSet.OK_CANCEL);
        
  // If Confirmed (OK), Initialize all League Data
  if(uiResponse == "OK"){
    // Open Spreadsheets
    var shtConfig = ss.getSheetByName('Config');
    var shtStandings   = ss.getSheetByName('Standings');
    var shtMatchRslt   = ss.getSheetByName('Match Results');
    var shtWeek;
    var shtResponses   = ss.getSheetByName('Responses');
    var shtResponsesEN = ss.getSheetByName('Responses EN');
    var shtResponsesFR = ss.getSheetByName('Responses FR');
    var ssWeekBstrID = shtConfig.getRange(40, 2).getValue();
    
    var MaxRowRslt = shtMatchRslt.getMaxRows();
    var MaxColRslt = shtMatchRslt.getMaxColumns();
    var MaxRowRspn = shtResponses.getMaxRows();
    var MaxColRspn = shtResponses.getMaxColumns();
    var MaxRowRspnEN = shtResponsesEN.getMaxRows();
    var MaxColRspnEN = shtResponsesEN.getMaxColumns();
    var MaxRowRspnFR = shtResponsesFR.getMaxRows();
    var MaxColRspnFR = shtResponsesFR.getMaxColumns();
    
    var ConfigData = shtConfig.getRange(56,2,32,1).getValues();
    var OptWeekRound = ConfigData[10][0]; 
    var ColMatchID = ConfigData[14][0];
    var ColMatchIDLastVal = ConfigData[19][0];
    
    // Clear Data
    shtStandings.getRange(6,2,32,6).clearContent();
    shtMatchRslt.getRange(6,2,MaxRowRslt-5,MaxColRslt-2).clearContent();
    shtResponses.getRange(2,1,MaxRowRspn-1,MaxColRspn).clearContent();
    shtResponses.getRange(1,ColMatchIDLastVal).setValue(0);
    shtResponsesEN.getRange(2,1,MaxRowRspnEN-1,MaxColRspnEN).clearContent();
    shtResponsesFR.getRange(2,1,MaxRowRspnFR-1,MaxColRspnFR).clearContent();
    
    // Clear Standing Spreadsheets Initialization Values
    shtConfig.getRange(34,5,2,1).clearContent();
    
    // Week Results
    for (var WeekNum = 1; WeekNum <= 8; WeekNum++){
      // Select Week or Round prefix (League or Tournament)
      if(OptWeekRound == 'Week') shtWeek = ss.getSheetByName('Week'+WeekNum);
      if(OptWeekRound == 'Round') shtWeek = ss.getSheetByName('Round'+WeekNum);
      // Clear Wins/Loss
      shtWeek.getRange(5,5,32,2).clearContent();
      // Clear Rest of Information
      shtWeek.getRange(5,8,32,107-8).clearContent();
    }
    
    Logger.log('League Data Cleared');
    
//    // Clear Weekly Booster Sheet
//    if(ssWeekBstrID != ''){
//      var ssWeekBstr = SpreadsheetApp.openById(ssWeekBstrID);
//      var WeekBstrSheets = ssWeekBstr.getSheets();
//      var WeekBstrNumSheets = ssWeekBstr.getNumSheets();
//      var shtWeekBstr = WeekBstrSheets[0];
//      var MaxCols;
//      
//      for(var sheet = 0; sheet < WeekBstrNumSheets; sheet++){
//        shtWeekBstr = WeekBstrSheets[sheet];
//        MaxCols = shtWeekBstr.getMaxColumns();
//        shtWeekBstr.getRange(4,2,18,MaxCols-1).clearContent();
//      }
//    }
    
    // Update Standings Copies
    fcnCopyStandingsSheets(ss, shtConfig, 0, 1);
    Logger.log('Standings Updated');
    
    // Clear Players DB and Card Pools
    fcnDelPlayerCardDB();
    fcnDelPlayerCardList();
    fcnDelPlayerStartPool();
    Logger.log('Card DB and Card Pool Cleared');
    
  // Generate Players DB and Card Pools
    fcnGenPlayerCardDB();
    fcnGenPlayerCardList();
    fcnGenPlayerStartPoolMain();
    Logger.log('Card DB and Card Pool Generated');
  }
}

// **********************************************
// function fcnResetLeagueMatch()
//
// This function clears all data from sheets  
// to start a new league
//
// **********************************************

function fcnResetLeagueMatch(){

  // Main Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Insert ui to confirm
  var ui = SpreadsheetApp.getUi();
  var title = "RESET LEAGUE MATCH CONFIRMATION";
  var msg = "Do you really want to reset all League Match Results";
  var uiResponse = ui.alert(title, msg, ui.ButtonSet.OK_CANCEL);
    
  // If Confirmed (OK), Reset all League Match Data, Responses and Card DB are unchanged
  if(uiResponse == "OK"){
    // Open Spreadsheets
    var shtConfig = ss.getSheetByName('Config');
    var shtStandings   = ss.getSheetByName('Standings');
    var shtMatchRslt   = ss.getSheetByName('Match Results');
    var shtWeek;
    var shtResponses   = ss.getSheetByName('Responses');
    var shtResponsesEN = ss.getSheetByName('Responses EN');
    var shtResponsesFR = ss.getSheetByName('Responses FR');
    
    var MaxRowRslt = shtMatchRslt.getMaxRows();
    var MaxColRslt = shtMatchRslt.getMaxColumns();
    var MaxRowRspn = shtResponses.getMaxRows();
    var MaxColRspn = shtResponses.getMaxColumns();
    var MaxRowRspnEN = shtResponsesEN.getMaxRows();
    var MaxColRspnEN = shtResponsesEN.getMaxColumns();
    var MaxRowRspnFR = shtResponsesFR.getMaxRows();
    var MaxColRspnFR = shtResponsesFR.getMaxColumns();
    
    var ConfigData = shtConfig.getRange(56,2,32,1).getValues();
    var OptWeekRound = ConfigData[10][0]; 
    var ColMatchID = ConfigData[14][0];
    var ColMatchIDLastVal = ConfigData[19][0];
    
    // Clear Data
    shtStandings.getRange(6,2,32,6).clearContent();
    shtMatchRslt.getRange(6,2,MaxRowRslt-5,MaxColRslt-2).clearContent();
    shtResponses.getRange(2,1,MaxRowRspn-1,MaxColRspn).clearContent();
    shtResponses.getRange(1,ColMatchIDLastVal).setValue(0);
    shtResponsesEN.getRange(2,ColMatchID,MaxRowRspnEN-1,7).clearContent();
    shtResponsesFR.getRange(2,ColMatchID,MaxRowRspnFR-1,7).clearContent();
    
    // Week Results
    for (var WeekNum = 1; WeekNum <= 8; WeekNum++){
      // Select Week or Round prefix (League or Tournament)
      if(OptWeekRound == 'Week') shtWeek = ss.getSheetByName('Week'+WeekNum);
      if(OptWeekRound == 'Round') shtWeek = ss.getSheetByName('Round'+WeekNum);
      // Clear Wins/Loss
      shtWeek.getRange(5,5,32,2).clearContent();
      // Clear Rest of Information
      shtWeek.getRange(5,8,32,107-8).clearContent();
    }
    
    Logger.log('League Data Cleared');
    
    // Update Standings Copies
    fcnCopyStandingsSheets(ss, shtConfig, 0, 1);
    Logger.log('Standings Updated');
  }
}




// **********************************************
// function fcnGenPlayerCardDB()
//
// This function generates all Card DB for all 
// players from the Config File
//
// **********************************************

function fcnGenPlayerCardDB(){
  
  // Main Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Config Spreadsheet
  var shtConfig = ss.getSheetByName('Config');
  
  // Card DB Spreadsheet
  var CardDBShtID = shtConfig.getRange(31, 2).getValue();
  var ssCardDB = SpreadsheetApp.openById(CardDBShtID);
  var NumSheet = ssCardDB.getNumSheets();
  var SheetsCardDB = ssCardDB.getSheets();
  var shtCardDB = ssCardDB.getSheetByName('Template');
  var shtCardDBNum;
  var SheetName;
  var CardDBHeader = shtCardDB.getRange(4,1,4,48).getValues();

  // Players 
  var shtPlayers = ss.getSheetByName('Players'); 
  var NbPlayers = shtPlayers.getRange(2,1).getValue();
  var PlayerFound = 0;
  
  var shtPlyrCardDB;
  var shtPlyrName;
  var SetNum;
  var PlyrRow;
  var CardDBNumSht;
  
  // Gets the Card Set Data from Config File to Populate the Header
  for (var col = 0; col < 48; col++){
    SetNum = CardDBHeader[0][col];
    switch (SetNum){
      case 1: 
        CardDBHeader[1][col] = shtConfig.getRange(7, 5).getValue();
        if (col < 32) CardDBHeader[2][col] = shtConfig.getRange(7, 6).getValue();
        if (col > 32) CardDBHeader[2][col] = shtConfig.getRange(7, 7).getValue();
        break;
      case 2: 
        CardDBHeader[1][col] = shtConfig.getRange(8, 5).getValue();
        if (col < 32) CardDBHeader[2][col] = shtConfig.getRange(8, 6).getValue();
        if (col > 32) CardDBHeader[2][col] = shtConfig.getRange(8, 7).getValue();
        break;
      case 3: 
        CardDBHeader[1][col] = shtConfig.getRange(9, 5).getValue();
        if (col < 32) CardDBHeader[2][col] = shtConfig.getRange(9, 6).getValue();
        if (col > 32) CardDBHeader[2][col] = shtConfig.getRange(9, 7).getValue();
        break;
      case 4: 
        CardDBHeader[1][col] = shtConfig.getRange(10, 5).getValue();
        if (col < 32) CardDBHeader[2][col] = shtConfig.getRange(10, 6).getValue();
        if (col > 32) CardDBHeader[2][col] = shtConfig.getRange(10, 7).getValue();
        break;
      case 5: 
        CardDBHeader[1][col] = shtConfig.getRange(11, 5).getValue();
        if (col < 32) CardDBHeader[2][col] = shtConfig.getRange(11, 6).getValue();
        if (col > 32) CardDBHeader[2][col] = shtConfig.getRange(11, 7).getValue();
        break;
      case 6: 
        CardDBHeader[1][col] = shtConfig.getRange(12, 5).getValue();
        if (col < 32) CardDBHeader[2][col] = shtConfig.getRange(12, 6).getValue();
        if (col > 32) CardDBHeader[2][col] = shtConfig.getRange(12, 7).getValue();
        break;
      case 7: 
        CardDBHeader[1][col] = shtConfig.getRange(13, 5).getValue();
        if (col < 32) CardDBHeader[2][col] = shtConfig.getRange(13, 6).getValue();
        if (col > 32) CardDBHeader[2][col] = shtConfig.getRange(13, 7).getValue();
        break;
      case 8: 
        CardDBHeader[1][col] = shtConfig.getRange(14, 5).getValue();
        if (col < 32) CardDBHeader[2][col] = shtConfig.getRange(14, 6).getValue();
        if (col > 32) CardDBHeader[2][col] = shtConfig.getRange(14, 7).getValue();
        break;
    }    
  }
  // Set Card Set Names and Codes
  shtCardDB.getRange(4,1,4,48).setValues(CardDBHeader);
  
  // Loops through each player starting from the First
  for (var plyr = 1; plyr <= NbPlayers; plyr++){
    
    // Update the Player Row and Get Player Name from Config File
    PlyrRow = plyr + 2; // 2 is the row where the player list starts
    shtPlyrName = shtPlayers.getRange(PlyrRow, 2).getValue();
    
    // Resets the Player Found flag before searching
    PlayerFound = 0;
        
    // Look if player exists, if yes, skip, if not, create player
    for(var sheet = NumSheet - 1; sheet >= 0; sheet --){
      SheetName = SheetsCardDB[sheet].getSheetName();
      if (SheetName == shtPlyrName) PlayerFound = 1;
    }
    
    // If Player is not found, add a tab
    if (PlayerFound == 0){
      // Get the Template sheet index
      CardDBNumSht = ssCardDB.getNumSheets();
      // INSERTS TAB BEFORE "Card DB" TAB
      ssCardDB.insertSheet(shtPlyrName, CardDBNumSht-2, {template: shtCardDB});
      shtPlyrCardDB = ssCardDB.getSheetByName(shtPlyrName);
      
      // Opens the new sheet and modify appropriate data (Player Name, Header)
      shtPlyrCardDB.getRange(3,3).setValue(shtPlyrName);
      shtPlyrCardDB.getRange(4,1,4,48).setValues(CardDBHeader);
    }
  }
  shtPlyrCardDB = ssCardDB.getSheets()[0];
  ssCardDB.setActiveSheet(shtPlyrCardDB);
}


// **********************************************
// function fcnGenPlayerCardList()
//
// This function generates all Card List for all 
// players from the Config File
//
// **********************************************

function fcnGenPlayerCardList(){
    
  // Main Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Config Spreadsheet
  var shtConfig = ss.getSheetByName('Config');
  
  // Card List Spreadsheet
  var CardListShtEnID = shtConfig.getRange(32, 2).getValue();
  var CardListShtFrID = shtConfig.getRange(33, 2).getValue();
  var ssCardListEn = SpreadsheetApp.openById(CardListShtEnID);
  var ssCardListFr = SpreadsheetApp.openById(CardListShtFrID);
  var shtCardListEn = ssCardListEn.getSheetByName('Template');
  var shtCardListFr = ssCardListFr.getSheetByName('Template');
  var shtCardListNum;
  var NumSheet = ssCardListEn.getNumSheets();
  var SheetsCardList = ssCardListEn.getSheets();
  var SheetName;
  var PlayerFound = 0;
  
  // Players 
  var shtPlayers = ss.getSheetByName('Players'); 
  var NbPlayers = shtPlayers.getRange(2,1).getValue();
  
  var shtPlyrCardListEn;
  var shtPlyrCardListFr;
  var shtPlyrName;
  var PlyrRow;
  var CardListNumSht;
  var shtName;
  
  // Loops through each player starting from the first
  for (var plyr = 1; plyr <= NbPlayers; plyr++){
  
    // Update the Player Row and Get Player Name from Config File
    PlyrRow = plyr + 2; // 2 is the row where the player list starts
    shtPlyrName = shtPlayers.getRange(PlyrRow, 2).getValue();
    
    // Resets the Player Found flag before searching
    PlayerFound = 0;
    
    // Look if player exists, if yes, skip, if not, create player
    for(var sheet = NumSheet - 1; sheet >= 0; sheet --){
      SheetName = SheetsCardList[sheet].getSheetName();
      
      if (SheetName == shtPlyrName) PlayerFound = 1;
    }
    
    if (PlayerFound == 0){
      // Get the Template sheet index
      CardListNumSht = ssCardListEn.getNumSheets();
      // INSERTS TAB BEFORE "Card DB" TAB
      // English Version
      ssCardListEn.insertSheet(shtPlyrName, CardListNumSht-1, {template: shtCardListEn});
      shtPlyrCardListEn = ssCardListEn.getSheetByName(shtPlyrName).showSheet();
      
      // Opens the new sheet and modify appropriate data (Player Name, Header)
      shtPlyrCardListEn.getRange(2,1).setValue(shtPlyrName);
      
      // French Version
      ssCardListFr.insertSheet(shtPlyrName, CardListNumSht-1, {template: shtCardListFr});
      shtPlyrCardListFr = ssCardListFr.getSheetByName(shtPlyrName).showSheet();
      
      // Opens the new sheet and modify appropriate data (Player Name, Header)
      shtPlyrCardListFr.getRange(2,1).setValue(shtPlyrName);    
    }
  }
  
  // English Version
  shtPlyrCardListEn = ssCardListEn.getSheets()[0];
  shtName = shtPlyrCardListEn.getName();
  ssCardListEn.setActiveSheet(shtPlyrCardListEn);
  if (shtName != 'Template') ssCardListEn.getSheetByName('Template').hideSheet();
    
  // French Version
  shtPlyrCardListFr = ssCardListFr.getSheets()[0];
  shtName = shtPlyrCardListFr.getName();
  ssCardListFr.setActiveSheet(shtPlyrCardListFr);
   if (shtName != 'Template') ssCardListFr.getSheetByName('Template').hideSheet();
}


// **********************************************
// function fcnGenPlayerStartPoolMain()
//
// This function generates Starting Pool for all 
// players from the Config File
//
// **********************************************

function fcnGenPlayerStartPoolMain(){
    
  // Main Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ssID = ss.getId();
  
  // Config Spreadsheet
  var shtConfig = ss.getSheetByName('Config');
  
  // Card Pool Spreadsheet
  var StartPoolShtID = shtConfig.getRange(41, 2).getValue();
  var ssStartPool = SpreadsheetApp.openById(StartPoolShtID);
  var shtStartPool = ssStartPool.getSheetByName('Template');
  var shtStartPoolNum;
  var NumSheet = ssStartPool.getNumSheets();
  var SheetsStartPool = ssStartPool.getSheets();
  var SheetName;
  var PlayerFound = 0;
  
  // Players 
  var shtPlayers = ss.getSheetByName('Players'); 
  var NbPlayers = shtPlayers.getRange(2,1).getValue();
  
  var shtPlyrStartPool;
  var shtPlyrName;
  var PlyrRow;
  var StartPoolNumSht;
  var shtName;
  
  // Loops through each player starting from the first
  for (var plyr = 1; plyr <= NbPlayers; plyr++){
  
    // Update the Player Row and Get Player Name from Config File
    PlyrRow = plyr + 2; // 2 is the row where the player list starts
    shtPlyrName = shtPlayers.getRange(PlyrRow, 2).getValue();
    
    // Resets the Player Found flag before searching
    PlayerFound = 0;
    
    // Look if player exists, if yes, skip, if not, create player
    for(var sheet = NumSheet - 1; sheet >= 0; sheet --){
      SheetName = SheetsStartPool[sheet].getSheetName();
      
      if (SheetName == shtPlyrName) PlayerFound = 1;
    }
    
    if (PlayerFound == 0){
      // Get the Template sheet index
      StartPoolNumSht = ssStartPool.getNumSheets();
      // INSERTS TAB BEFORE "Card DB" TAB
      // English Version
      ssStartPool.insertSheet(shtPlyrName, StartPoolNumSht-1, {template: shtStartPool});
      shtPlyrStartPool = ssStartPool.getSheetByName(shtPlyrName).showSheet();
      
      // Opens the new sheet and modify appropriate data (Player Name, Header)
      shtPlyrStartPool.getRange(1,2).setValue(shtPlyrName);
      shtPlyrStartPool.getRange(2,2).setValue('Not Processed');
      shtPlyrStartPool.getRange(3,2).setValue(ssID);
      
      //Hides the 3rd row
      shtPlyrStartPool.hideRows(3);
    }
  }
  
  // Hide Template Sheet
  shtPlyrStartPool = ssStartPool.getSheets()[0];
  shtName = shtPlyrStartPool.getName();
  ssStartPool.setActiveSheet(shtPlyrStartPool);
   if (shtName != 'Template') ssStartPool.getSheetByName('Template').hideSheet();
}

// **********************************************
// function fcnGenPlayerWeekBstr()
//
// This function generates Weekly Booster Lists 
// for all players from the Config File
//
// **********************************************

function fcnGenPlayerWeekBstr(){
    
  // Main Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ssID = ss.getId();
  
  // Config Spreadsheet
  var shtConfig = ss.getSheetByName('Config');
  
  // Card Pool Spreadsheet
  var WeekBstrShtID = shtConfig.getRange(40, 2).getValue();
  var ssWeekBstr = SpreadsheetApp.openById(WeekBstrShtID);
  var shtWeekBstr = ssWeekBstr.getSheetByName('Template');
  var shtWeekBstrNum;
  var NumSheet = ssWeekBstr.getNumSheets();
  var SheetsWeekBstr = ssWeekBstr.getSheets();
  var SheetName;
  var PlayerFound = 0;
  
  // Players 
  var shtPlayers = ss.getSheetByName('Players'); 
  var NbPlayers = shtPlayers.getRange(2,1).getValue();
  
  var shtPlyrWeekBstr;
  var shtPlyrName;
  var PlyrRow;
  var WeekBstrNumSht;
  var shtName;
  
  // Loops through each player starting from the first
  for (var plyr = 1; plyr <= NbPlayers; plyr++){
  
    // Update the Player Row and Get Player Name from Config File
    PlyrRow = plyr + 2; // 2 is the row where the player list starts
    shtPlyrName = shtPlayers.getRange(PlyrRow, 2).getValue();
    
    // Resets the Player Found flag before searching
    PlayerFound = 0;
    
    // Look if player exists, if yes, skip, if not, create player
    for(var sheet = NumSheet - 1; sheet >= 0; sheet --){
      SheetName = SheetsWeekBstr[sheet].getSheetName();
      
      if (SheetName == shtPlyrName) PlayerFound = 1;
    }
    
    if (PlayerFound == 0){
      // Get the Template sheet index
      WeekBstrNumSht = ssWeekBstr.getNumSheets();
      // INSERTS TAB BEFORE "Template" TAB
      // English Version
      ssWeekBstr.insertSheet(shtPlyrName, WeekBstrNumSht-1, {template: shtWeekBstr});
      shtPlyrWeekBstr = ssWeekBstr.getSheetByName(shtPlyrName).showSheet();
      
      // Opens the new sheet and modify appropriate data (Player Name, Header)
      shtPlyrWeekBstr.getRange(1,1).setValue(shtPlyrName);
    }
  }
  
  // Hide Template Sheet
  shtPlyrWeekBstr = ssWeekBstr.getSheets()[0];
  shtName = shtPlyrWeekBstr.getName();
  ssWeekBstr.setActiveSheet(shtPlyrWeekBstr);
   if (shtName != 'Template') ssWeekBstr.getSheetByName('Template').hideSheet();
}

// **********************************************
// function fcnDelPlayerCardDB()
//
// This function deletes all Card DB for all 
// players from the Config File
//
// **********************************************

function fcnDelPlayerCardDB(){

  // Main Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Config Spreadsheet
  var shtConfig = ss.getSheetByName('Config');
  
  // Card DB Spreadsheet
  var CardDBShtID = shtConfig.getRange(31, 2).getValue();
  var ssCardDB = SpreadsheetApp.openById(CardDBShtID);
  var shtTemplate = ssCardDB.getSheetByName('Template');
  var ssNbSheet = ssCardDB.getNumSheets();
  
  // Routine Variables
  var shtCurr;
  var shtCurrName;
  
  // Activates Template Sheet
  ssCardDB.setActiveSheet(shtTemplate);
  
  for (var sht = 0; sht < ssNbSheet - 1; sht ++){
    shtCurr = ssCardDB.getSheets()[0];
    shtCurrName = shtCurr.getName();
    if(shtCurrName != 'Template') ssCardDB.deleteSheet(shtCurr);
    }
}

// **********************************************
// function fcnDelPlayerCardList()
//
// This function deletes all Card Pools for all 
// players from the Config File
//
// **********************************************

function fcnDelPlayerCardList(){

  // Main Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Config Spreadsheet
  var shtConfig = ss.getSheetByName('Config');
  
  // Card Pool Spreadsheet
  var CardListShtIDEn = shtConfig.getRange(32, 2).getValue();
  var CardListShtIDFr = shtConfig.getRange(33, 2).getValue();
  var ssCardListEn = SpreadsheetApp.openById(CardListShtIDEn);
  var ssCardListFr = SpreadsheetApp.openById(CardListShtIDFr);
  var shtTemplateEn = ssCardListEn.getSheetByName('Template');
  var shtTemplateFr = ssCardListFr.getSheetByName('Template');
  var ssNbSheetEn = ssCardListEn.getNumSheets();
  var ssNbSheetFr = ssCardListFr.getNumSheets();  
  
    // Routine Variables
  var shtCurrEn;
  var shtCurrNameEn;
  var shtCurrFr;
  var shtCurrNameFr;
  var NbSheet;
  
  // Show Template sheet
  shtTemplateEn.showSheet();
  shtTemplateFr.showSheet();
  
  // Activates Template Sheet
  ssCardListEn.setActiveSheet(shtTemplateEn);
  ssCardListFr.setActiveSheet(shtTemplateFr);
  
  // Check greater number of sheets
  if (ssNbSheetEn >= ssNbSheetFr) NbSheet = ssNbSheetEn;
  if (ssNbSheetFr >= ssNbSheetEn) NbSheet = ssNbSheetFr;  
  
  for (var sht = 0; sht < NbSheet - 1; sht ++){
    
    // English Version
    shtCurrEn = ssCardListEn.getSheets()[0];
    shtCurrNameEn = shtCurrEn.getName();
    if(shtCurrNameEn != 'Template') ssCardListEn.deleteSheet(shtCurrEn);
    
    // French Version   
    shtCurrFr = ssCardListFr.getSheets()[0];
    shtCurrNameFr = shtCurrFr.getName();
    Logger.log(shtCurrNameFr);
    if(shtCurrNameFr != 'Template') ssCardListFr.deleteSheet(shtCurrFr);
  }
}

// **********************************************
// function fcnDelPlayerStartPool()
//
// This function deletes all Starting Pool for all 
// players from the Config File
//
// **********************************************

function fcnDelPlayerStartPoolMain(){

  // Main Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Config Spreadsheet
  var shtConfig = ss.getSheetByName('Config');
  
  // Card Pool Spreadsheet
  var StartPoolShtID = shtConfig.getRange(41, 2).getValue();
  var ssStartPool = SpreadsheetApp.openById(StartPoolShtID);
  var shtTemplate = ssStartPool.getSheetByName('Template');
  var ssNbSheet = ssStartPool.getNumSheets();
  
  // Routine Variables
  var shtCurr;
  var shtCurrName;
  
  // Show Template sheet
  shtTemplate.showSheet();
  
  // Activates Template Sheet
  ssStartPool.setActiveSheet(shtTemplate);
  
  for (var sht = 0; sht < ssNbSheet - 1; sht ++){
    
    // English Version
    shtCurr = ssStartPool.getSheets()[0];
    shtCurrName = shtCurr.getName();
    if(shtCurrName != 'Template') ssStartPool.deleteSheet(shtCurr);
  }
}

// **********************************************
// function fcnDelPlayerWeekBstr()
//
// This function deletes all Starting Pool for all 
// players from the Config File
//
// **********************************************

function fcnDelPlayerWeekBstr(){

  // Main Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Config Spreadsheet
  var shtConfig = ss.getSheetByName('Config');
  
  // Card Pool Spreadsheet
  var WeekBstrShtID = shtConfig.getRange(40, 2).getValue();
  var ssWeekBstr = SpreadsheetApp.openById(WeekBstrShtID);
  var shtTemplate = ssWeekBstr.getSheetByName('Template');
  var ssNbSheet = ssWeekBstr.getNumSheets();
  
  // Routine Variables
  var shtCurr;
  var shtCurrName;
  
  // Show Template sheet
  shtTemplate.showSheet();
  
  // Activates Template Sheet
  ssWeekBstr.setActiveSheet(shtTemplate);
  
  for (var sht = 0; sht < ssNbSheet - 1; sht ++){
    
    shtCurr = ssWeekBstr.getSheets()[2];
    shtCurrName = shtCurr.getName();
    Logger.log(shtCurrName);
    if(shtCurrName != 'Template') ssWeekBstr.deleteSheet(shtCurr);
  }
}

// **********************************************
// function fcnSetupResponseSht()
//
// This function sets up the new Responses sheets 
// and deletes the old ones
//
// **********************************************

function fcnSetupResponseSht(){

  // Main Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Open Responses Sheets
  var shtOldRespEN = ss.getSheetByName('Responses EN');
  var shtOldRespFR = ss.getSheetByName('Responses FR');
  var shtNewRespEN = ss.getSheetByName('New Responses EN');
  var shtNewRespFR = ss.getSheetByName('New Responses FR');
    
  var OldRespMaxCol = shtOldRespEN.getMaxColumns();
  var NewRespMaxRow = shtNewRespEN.getMaxRows();
  var ColWidth;
  
  // Copy Header from Old to New sheet - Loop to Copy Value and Format from cell to cell, copy formula (or set) in last cell
  for (var col = 1; col <= OldRespMaxCol; col++){
    // Insert Column if it doesn't exist (col >=24)
    if (col >= 24 && col < OldRespMaxCol){
      shtNewRespEN.insertColumnAfter(col);
      shtNewRespFR.insertColumnAfter(col);
    }
    // Set New Response Sheet Values 
    shtOldRespEN.getRange(1,col).copyTo(shtNewRespEN.getRange(1,col));
    shtOldRespFR.getRange(1,col).copyTo(shtNewRespFR.getRange(1,col));
    ColWidth = shtOldRespEN.getColumnWidth(col);
    shtNewRespEN.setColumnWidth(col,ColWidth);
    shtNewRespFR.setColumnWidth(col,ColWidth);
  }
  // Hides Columns 25, 27-30
  shtNewRespEN.hideColumns(25);
  shtNewRespEN.hideColumns(27,4);
  shtNewRespFR.hideColumns(25);
  shtNewRespFR.hideColumns(27,4);
  
  // Deletes all Rows but 1-2
  shtNewRespEN.deleteRows(3, NewRespMaxRow - 2);
  shtNewRespFR.deleteRows(3, NewRespMaxRow - 2);
    
  // Delete Old Sheets
  ss.deleteSheet(shtOldRespEN);
  ss.deleteSheet(shtOldRespFR);
  
  // Rename New Sheets
  shtNewRespEN.setName('Responses EN');
  shtNewRespFR.setName('Responses FR');

}
