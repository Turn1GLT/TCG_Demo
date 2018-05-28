// **********************************************
// function fcnUpdateLinksIDs()
//
// This function updates all sheets Links and IDs  
// in the Config File
//
// **********************************************

function fcnUpdateLinksIDs(){
  
  Logger.log("Routine: fcnUpdateLinksIDs");
  
  // Main Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Config Spreadsheet
  var shtConfig = ss.getSheetByName('Config');
  
  // Copy Log Spreadsheet
  var shtCopyLogID = shtConfig.getRange(9,15).getValue();
  var LinksStatus = shtConfig.getRange(9,16).getValue();
  
  if (shtCopyLogID != '' && LinksStatus =='') {
    var shtCopyLog = SpreadsheetApp.openById(shtCopyLogID).getSheets()[0];
  
    var CopyLogNbFiles = shtCopyLog.getRange(2, 6).getValue();
    var rowStartCopyLog = 5;
    var rowStartConfig = 4;
    var colShtId = 7;
    var colShtUrl = 8;
    
    var CopyLogVal = shtCopyLog.getRange(rowStartCopyLog, 2, CopyLogNbFiles, 3).getValues();
    // [0]= Sheet Name, [1]= Sheet URL, [2]= Sheet ID
    
    var FileName;
    var Link;
    var Formula;
    var rowCfg = 'Not Found';
    
    // Clear Sheet IDs
    shtConfig.getRange(rowStartConfig, colShtId,20,1).clearContent();
    // Clear Sheet URLs
    shtConfig.getRange(rowStartConfig,colShtUrl,20,1).clearContent();
    
    // Loop through all Copied Sheets and get their Link and ID
    for (var row = 0; row < CopyLogNbFiles; row++){
      // Get File Name
      FileName = CopyLogVal[row][0];
      
      switch(FileName){
        case 'Master TCG Event' :
          rowCfg = rowStartConfig + 0; break;
        case 'Master TCG Log' :
          rowCfg = rowStartConfig + 1; break;
        case 'Master TCG Card DB' :
          rowCfg = rowStartConfig + 2; break;
        case 'Master TCG Card Lists EN' :
          rowCfg = rowStartConfig + 3; break;
        case 'Master TCG Card Lists FR' :
          rowCfg = rowStartConfig + 4; break;
        case 'Master TCG Standings EN' :
          rowCfg = rowStartConfig + 5; break;
        case 'Master TCG Standings FR' :
          rowCfg = rowStartConfig + 6; break;	
        case 'Master TCG Player Records' :
          rowCfg = rowStartConfig + 13; break;
        case 'Master TCG Player List & Round Bonus' :
          rowCfg = rowStartConfig + 14; break;
        case 'Master TCG Starting Pool' :
          rowCfg = rowStartConfig + 15; break;        
        default : 
          rowStartConfig = 'Not Found'; break;
      }
      
      // Set the Appropriate Sheet ID Value and URL in the Config File
      if (rowCfg != 'Not Found') {
        shtConfig.getRange(rowCfg, colShtId).setValue(CopyLogVal[row][2]);
        // Opens Spreadsheet by ID to get URL
        Link = SpreadsheetApp.openById(CopyLogVal[row][2]).getUrl();        
        shtConfig.getRange(rowCfg, colShtUrl).setValue(Link);
      }
    }
  }
}

// **********************************************
// function fcnInitializeEvent()
//
// This function clears all data from sheets  
// to start a new Event (League / Tournament)
//
// **********************************************

function fcnInitializeEvent(){
  
  Logger.log("Routine: fcnInitializeEvent");

  // Main Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shtConfig = ss.getSheetByName('Config');
  var cfgEventType = shtConfig.getRange(7,4).getValue();
  
  // Insert ui to confirm
  var ui = SpreadsheetApp.getUi();
  var title = "Clear "+ cfgEventType +" Data Confirmation";
  var msg = "Click OK to clear all "+ cfgEventType +" Data to start a new " + cfgEventType;
  var uiResponse = ui.alert(title, msg, ui.ButtonSet.OK_CANCEL);
    
  // If Confirmed (OK), Initialize all League Data
  if(uiResponse == "OK"){
    //  if(cfgEventType == "League" || cfgEventType == "Tournament"){
    // Config Sheet to get options
    var shtIDs = shtConfig.getRange(4,7,24,1).getValues();
    var cfgColRspSht = shtConfig.getRange(4,15,16,1).getValues();
    var cfgColRndSht = shtConfig.getRange(4,18,16,1).getValues();
    var cfgEvntParam = shtConfig.getRange(4,4,48,1).getValues();
    // Registration Form Construction 
    // Column 1 = Category Name
    // Column 2 = Category Order in Form
    // Column 3 = Column Value in Player/Team Sheet
    var cfgRegFormCnstrVal = shtConfig.getRange(4,23,20,3).getValues();
    
    // Event Parameters
    var evntLocation = cfgEvntParam[0][0];
    var evntNameEN =   cfgEvntParam[7][0];
    var evntNameFR =   cfgEvntParam[8][0];
    var evntCntctGrpNameEN = evntLocation + " " + evntNameEN;
    var evntCntctGrpNameFR = evntLocation + " " + evntNameFR;
    var ContactGroupEN;
    var ContactGroupFR;
    
    // Columns from Config File
    var colRspMatchID        = cfgColRspSht[1][0];
    var colRspMatchIDLastVal = cfgColRspSht[6][0];
    
    // Column Round Sheets
    var colRndMP             = cfgColRndSht[2][0];
    
    var colPlyrName   = cfgRegFormCnstrVal[ 2][2];
    var colPlyrStatus = cfgRegFormCnstrVal[16][2];
    
    // Sheets
    var shtStandings =   ss.getSheetByName('Standings');
    var shtRound       = ss.getSheetByName('Round1');
    var shtMatchRslt   = ss.getSheetByName('Match Results');
    var shtResponses   = ss.getSheetByName('Responses');
    var shtMatchRespEN = ss.getSheetByName('MatchResp EN');
    var shtMatchRespFR = ss.getSheetByName('MatchResp FR');
    var shtPlayers =     ss.getSheetByName('Players');
    var ssStrPlayers = SpreadsheetApp.openById(shtIDs[10][0]); // Store Player List Spreadsheet
    var shtStrPlayers = ssStrPlayers.getSheetByName('Players');// Store Player List Sheet
    
    // Max Rows / Columns
    var MaxRowStdg = shtStandings.getMaxRows();
    var MaxColStdg = shtStandings.getMaxColumns();
    var MaxRowRslt = shtMatchRslt.getMaxRows();
    var MaxColRslt = shtMatchRslt.getMaxColumns();
    var MaxRowRspn = shtResponses.getMaxRows();
    var MaxColRspn = shtResponses.getMaxColumns();
    var MaxRowRspnEN = shtMatchRespEN.getMaxRows();
    var MaxColRspnEN = shtMatchRespEN.getMaxColumns();
    var MaxRowRspnFR = shtMatchRespFR.getMaxRows();
    var MaxColRspnFR = shtMatchRespFR.getMaxColumns();
    var MaxRowRndSht = shtRound.getMaxRows();
    var MaxColRndSht = shtRound.getMaxColumns();
    var MaxRowPlayers = shtPlayers.getMaxRows();
    var MaxColPlayers = shtPlayers.getMaxColumns();
        
    // Clear Data
    // Standings
    shtStandings.getRange(6,2,MaxRowStdg-5,MaxColStdg-1).clearContent();
    // Match Results (does not clear the last column)
    shtMatchRslt.getRange(6,2,MaxRowRslt-5,MaxColRslt-1).clearContent();
    // Responses
    shtResponses.getRange(2,1,MaxRowRspn-1,MaxColRspn).clearContent();
    shtResponses.getRange(1,colRspMatchIDLastVal).setValue(0);
    shtMatchRespEN.getRange(2,1,MaxRowRspnEN-1,MaxColRspnEN).clearContent();
    shtMatchRespFR.getRange(2,1,MaxRowRspnFR-1,MaxColRspnFR).clearContent()
    
    // Round Results
    for (var RoundNum = 1; RoundNum <= 8; RoundNum++){
      shtRound = ss.getSheetByName('Round'+RoundNum);
      shtRound.getRange(5,colRndMP,MaxRowRndSht-4,MaxColRndSht-colRndMP+1).clearContent();
    }
    Logger.log('Event Data Cleared');
    
    // Clear Player List
    // From Player Name to Status
    shtPlayers.getRange(3, 2, MaxRowPlayers-2, colPlyrStatus-colPlyrName).clearContent();
    // From Status to rest of File
    shtPlayers.getRange(3, colPlyrStatus+1, MaxRowPlayers-2, MaxColPlayers-colPlyrStatus).clearContent();
    shtStrPlayers.getRange(3, 2, MaxRowPlayers-2, MaxColPlayers-1).clearContent();
    Logger.log('Player List Cleared');
    
    // Delete Contact Groups
    // Get Contact Group
    ContactGroupEN = ContactsApp.getContactGroup(evntCntctGrpNameEN);
    ContactGroupFR = ContactsApp.getContactGroup(evntCntctGrpNameFR);
    // If Contact Group exists, Delete it
    if(ContactGroupEN != null) ContactsApp.deleteContactGroup(ContactGroupEN);
    if(ContactGroupFR != null) ContactsApp.deleteContactGroup(ContactGroupFR);
    Logger.log('Contact Groups Deleted');
    
    // Update Standings Copies
    fcnCopyStandingsSheets(ss, shtConfig, cfgEvntParam, cfgColRndSht, 0, 1);
    Logger.log('Standings Updated');
    
    // Clear Players DB and Card Pools
    fcnDelPlayerCardDB();
    fcnDelPlayerCardList();
    fcnDelEventRecord();
    Logger.log('Card DB and Card Lists Cleared');
        
    title = cfgEventType +" Data Cleared";
    msg = "All " + cfgEventType +" Data has been cleared. You are now ready to start a new " + cfgEventType;
    uiResponse = ui.alert(title, msg, ui.ButtonSet.OK);    
  }
}

// **********************************************
// function fcnClearMatchResults()
//
// This function clears all Results data but
// does not clear Responses
//
// **********************************************

function fcnClearMatchResults(){
  
  Logger.log("Routine: fcnClearMatchResults");

  // Main Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shtConfig = ss.getSheetByName('Config');
  var cfgEventType = shtConfig.getRange(7,4).getValue();
  
  // Insert ui to confirm
  var ui = SpreadsheetApp.getUi();
  var title = "Reset " + cfgEventType + " Match Results";
  var msg = "Click OK to clear all "+ cfgEventType +" match results";
  var uiResponse = ui.alert(title, msg, ui.ButtonSet.OK_CANCEL);
    
  // If Confirmed (OK), Initialize all League Data
  if(uiResponse == "OK"){
    
    // Config Sheet to get options
    var shtIDs = shtConfig.getRange(4,7,24,1).getValues();
    var cfgColRspSht = shtConfig.getRange(4,15,16,1).getValues();
    var cfgColRndSht = shtConfig.getRange(4,18,16,1).getValues();
    var cfgEvntParam = shtConfig.getRange(4,4,48,1).getValues();
    
    // Columns from Config File
    var colRspMatchID        = cfgColRspSht[1][0];
    var colRspMatchIDLastVal = cfgColRspSht[6][0];
    
    // Column Round Sheets
    var colRndMP             = cfgColRndSht[2][0];
    
    // Sheets
    var shtStandings   = ss.getSheetByName('Standings');
    var shtRound       = ss.getSheetByName('Round1');
    var shtMatchRslt   = ss.getSheetByName('Match Results');
    var shtResponses   = ss.getSheetByName('Responses');
    var shtMatchRespEN = ss.getSheetByName('MatchResp EN');
    var shtMatchRespFR = ss.getSheetByName('MatchResp FR');
    
    // Max Rows / Columns
    var MaxRowStdg = shtStandings.getMaxRows();
    var MaxColStdg = shtStandings.getMaxColumns();
    var MaxRowRslt = shtMatchRslt.getMaxRows();
    var MaxColRslt = shtMatchRslt.getMaxColumns();
    var MaxRowRspn = shtResponses.getMaxRows();
    var MaxColRspn = shtResponses.getMaxColumns();
    var MaxRowRspnEN = shtMatchRespEN.getMaxRows();
    var MaxColRspnEN = shtMatchRespEN.getMaxColumns();
    var MaxRowRspnFR = shtMatchRespFR.getMaxRows();
    var MaxColRspnFR = shtMatchRespFR.getMaxColumns();
    var MaxRowRndSht = shtRound.getMaxRows();
    var MaxColRndSht = shtRound.getMaxColumns();
    
    // Clear Data
    shtStandings.getRange(6,2,MaxRowStdg-5,MaxColStdg-1).clearContent();
    shtMatchRslt.getRange(6,2,MaxRowRslt-5,MaxColRslt-1).clearContent();
    shtResponses.getRange(2,1,MaxRowRspn-1,MaxColRspn).clearContent();
    shtResponses.getRange(1,colRspMatchIDLastVal).setValue(0);
    shtMatchRespEN.getRange(2,colRspMatchID,MaxRowRspnEN-1,MaxColRspnEN-colRspMatchID+1).clearContent();
    shtMatchRespFR.getRange(2,colRspMatchID,MaxRowRspnFR-1,MaxColRspnFR-colRspMatchID+1).clearContent();
    
    // Round Results
    for (var RoundNum = 1; RoundNum <= 8; RoundNum++){
      shtRound = ss.getSheetByName('Round'+RoundNum);
      shtRound.getRange(5,colRndMP,MaxRowRndSht-4,MaxColRndSht-colRndMP+1).clearContent();
    }
    
    // Clear Event Records
    fcnClrEvntRecord();
    
    Logger.log('Match Data Cleared');
    
    // Update Standings Copies
    fcnCopyStandingsSheets(ss, shtConfig, cfgEvntParam, cfgColRndSht, 0, 1);
    Logger.log('Standings Updated');
    
    title = "Match Results Cleared";
    msg = "All Match Results have been cleared. You are now ready to submit Match Reports";
    uiResponse = ui.alert(title, msg, ui.ButtonSet.OK);
  }
}

// **********************************************
// function fcnCrtEvntRecord()
//
// This function generates all Players Records 
// from the Config File
//
// **********************************************

function fcnCrtEvntRecord(){
  
  Logger.log("Routine: fcnCrtEvntRecord");
    
  // Main Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Config Spreadsheet
  var shtConfig = ss.getSheetByName('Config');
  var shtIDs = shtConfig.getRange(4,7,24,1).getValues();
  var cfgEvntParam =  shtConfig.getRange( 4, 4,48,1).getValues();
  var cfgColShtPlyr = shtConfig.getRange( 4,25,30,1).getValues();
  var cfgColShtTeam = shtConfig.getRange(24,25,30,1).getValues();
  
  // Event Log Spreadsheet
  var ssEventRecord = SpreadsheetApp.openById(shtIDs[9][0]);
  var shtTemplate = ssEventRecord.getSheetByName('Template');
  var shtArmyListNum;
  
  // Column Values
  var colShtPlyrName = cfgColShtPlyr[2][0];
  var colShtPlyrLang = cfgColShtPlyr[5][0];
  var colShtTeamName = cfgColShtTeam[7][0];
  var colShtTeamLang = cfgColShtTeam[5][0];
  
  // Event Parameters
  var evntFormat = cfgEvntParam[ 9][0];
  
  // Sheets Values
  var NbSheet = ssEventRecord.getNumSheets();
  var ssSheets = ssEventRecord.getSheets();
  
  // Players 
  var shtPlayers = ss.getSheetByName('Players'); 
  var NbPlayers = shtPlayers.getRange(2,1).getValue();
  // Get Players Names and Languages 
  var PlyrNames = shtPlayers.getRange(2,colShtPlyrName, NbPlayers+1, 1).getValues();
  var PlyrLang =  shtPlayers.getRange(2,colShtPlyrLang, NbPlayers+1, 1).getValues();
  
  // Teams 
  var shtTeams = ss.getSheetByName('Teams'); 
  var NbTeams = shtTeams.getRange(2,1).getValue();
  // Get Teams Names and Languages 
  var TeamNames = shtTeams.getRange(2,colShtTeamName, NbTeams+1, 1).getValues();
  var TeamLang =  shtTeams.getRange(2,colShtTeamLang, NbTeams+1, 1).getValues();
  
  // Routine Variables
  var shtPT;
  var namePT;
  var langPT;
  var nameSheet;
  var GlobalHdr;
  var HstryHdr;
  var LoopMax;
  var PTFound = 0;
  
  // Defines Loop Parameters
  if(evntFormat == "Single") LoopMax = NbPlayers;
  
  if(evntFormat == "Team") LoopMax = NbTeams;
  
  // Loops through each player starting from the Last
  for (var PT = LoopMax; PT > 0; PT--){
    
    // Gets the Player/Team Name and Language
    if(evntFormat == "Single"){
      namePT = PlyrNames[PT][0];
      langPT = PlyrLang[PT][0];
    }
    if(evntFormat == "Team"){
      namePT = TeamNames[PT][0];
      langPT = TeamLang[PT][0];
    }
    
    // Resets the Player/Team Found flag before searching
    PTFound = 0;
    // Look if Player/Team exists, if yes, skip, if not, create Player/Team
    for(var sheet = NbSheet; sheet > 0; sheet --){
      nameSheet = ssSheets[sheet-1].getSheetName();
      if (nameSheet == namePT) PTFound = 1;
    }
          
    // If Player/Team is not found, add a tab
    if(PTFound == 0){
      // Inserts Sheet before Template (Last Sheet in Spreadsheet)
      ssEventRecord.insertSheet(namePT, NbSheet-1, {template: shtTemplate});
      shtPT = ssEventRecord.getSheetByName(namePT);
      shtPT.showSheet();
      
      // Updates the number of sheets
      NbSheet = ssEventRecord.getNumSheets();
      ssSheets = ssEventRecord.getSheets();
      
      // Opens the new sheet and modify appropriate data (Player/Team Name, Header)
      shtPT.getRange(2,1).setValue(namePT);
      
      // Translate Header if Player/Team Language Preference is French
      if(langPT == 'Français'){
        // Set Global Header
        GlobalHdr = shtPT.getRange(3,1,1,6).getValues();
        GlobalHdr[0][0] = 'Joué';           // Played
        GlobalHdr[0][1] = 'Victoires';      // Win
        GlobalHdr[0][2] = 'Défaites';       // Loss
        GlobalHdr[0][3] = 'Nulles';         // Tie
        GlobalHdr[0][4] = '';               // N/A Pts Scored
        GlobalHdr[0][5] = '% Victoire';     // Win%
        // Update Header
        shtPT.getRange(3,1,1,6).setValues(GlobalHdr);
      
        // Set History Header
        HstryHdr = shtPT.getRange(6,1,1,6).getValues();
        HstryHdr[0][0] = 'Événement';      // Event Name
        HstryHdr[0][1] = '';               // Event Name (merged cell)
        HstryHdr[0][2] = 'Ronde';          // Round
        HstryHdr[0][3] = 'Résultat';       // Match Result
        HstryHdr[0][4] = 'Joué contre';    // Played vs
        HstryHdr[0][5] = '';               // Played vs (merged cell)
        // Update Header
        shtPT.getRange(6,1,1,6).setValues(HstryHdr);
      }
    }
  }
  // English Version
  ssEventRecord.setActiveSheet(ssEventRecord.getSheets()[0]);
  ssEventRecord.getSheetByName('Template').hideSheet();
}


// **********************************************
// function fcnDelEventRecord()
//
// This function deletes all Players/Teams Record Sheets
// from the Config File
//
// **********************************************

function fcnDelEventRecord(){
  
  Logger.log("Routine: fcnDelEventRecord");

  // Main Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Config Spreadsheet
  var shtConfig = ss.getSheetByName('Config');
  
  // Sheet IDs
  var shtIDs = shtConfig.getRange(4,7,24,1).getValues();
  
  // Template Sheet
  var ssDel = SpreadsheetApp.openById(shtIDs[9][0]);
  var shtTemplate = ssDel.getSheetByName('Template').showSheet();
  
  // Delete Event Records
  subDelPlayerSheets(shtIDs[9][0]);

}

// **********************************************
// function fcnClrEvntRecord()
//
// This function clears all data in Player Record Sheets
//
// **********************************************

function fcnClrEvntRecord(){

  Logger.log("Routine: fcnClrEvntRecord");

  // Main Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Config Spreadsheet
  var shtConfig = ss.getSheetByName('Config');
  
  // Get Player Log Spreadsheet
  var shtIDs = shtConfig.getRange(4,7,24,1).getValues();
  var ssEvntPlyrRec = SpreadsheetApp.openById(shtIDs[9][0]);
  var evntPlyrRecNbSheets = ssEvntPlyrRec.getNumSheets();
  var evntPlyrSheets = ssEvntPlyrRec.getSheets();
  var evntPlyrRowStart = 7;
  
  // Routine Variables
  var sheet;
  var shtMaxCol;
  var shtMaxRow;
  
  // Loop through all Players Sheets
  for(var sht = 0; sht < evntPlyrRecNbSheets; sht++){
    // Get Sheet
    sheet = evntPlyrSheets[sht];
    shtMaxCol = sheet.getMaxColumns();
    shtMaxRow = sheet.getMaxRows();
    
    // Clear Player Record
    sheet.getRange(4,1,1,shtMaxCol).clearContent();
    
    // Delete all History Rows from Row 8 to Max Row
    if(shtMaxRow > evntPlyrRowStart) sheet.deleteRows(evntPlyrRowStart+1, shtMaxRow-evntPlyrRowStart);
    
    // Clear Player History
    sheet.getRange(evntPlyrRowStart, 1, 1, shtMaxCol).clearContent();
  }
}



// **********************************************
// function fcnCrtPlayerCardDB()
//
// This function generates all Army DB for all 
// players from the Config File
//
// **********************************************

function fcnCrtPlayerCardDB(){
  
  Logger.log("Routine: fcnCrtPlayerCardDB");
  
  // Main Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Config Spreadsheet
  var shtConfig = ss.getSheetByName('Config');
  var shtIDs = shtConfig.getRange(4,7,24,1).getValues();
  
  // Configuration Data
  var cfgColShtPlyr = shtConfig.getRange(4,25,20,1).getValues();
  var cfgDeckBuild =  shtConfig.getRange(4,30,20,1).getValues();
  
  // Legal Sets Data from Config File
  var cfgSetData =    shtConfig.getRange(3,31,9,5).getValues();
  // [x][0] = Set Presence, [x][1] = Set Number, [x][2] = Set Abreviation, [x][3] = Set Name, [x][4] = Set Masterpiece Series
  // x = Set Number (1-8)
  var cfgNbSet = cfgSetData[0][0];

  // Column Values
  var colShtPlyrName = cfgColShtPlyr[2][0];
  var colShtPlyrTeam = cfgColShtPlyr[5][0];
  
  // Deck Building Configuration
  var cfgCardPack = cfgDeckBuild[7][0];
  
  // Card DB Spreadsheet
  var ssCardDB = SpreadsheetApp.openById(shtIDs[2][0]);
  var shtTemplate = ssCardDB.getSheetByName('Template');
  var CardDBHeader = shtTemplate.getRange(4,1,4,48).getValues();
  var NbSheet = ssCardDB.getNumSheets();
  var ssSheets = ssCardDB.getSheets();
  var SheetName;
  var PlayerFound = 0;

  // Players 
  var shtPlayers = ss.getSheetByName('Players'); 
  var MaxColPlayers = shtPlayers.getMaxColumns();
  var NbPlayers = shtPlayers.getRange(2,1).getValue();
  
  // Get Players Data
  var PlayerData = shtPlayers.getRange(2,colShtPlyrName, NbPlayers+1, MaxColPlayers-1).getValues();
     
  var shtPlyr;
  var PlyrName;
  var SetNum;
  var frmlNbPack;
  
  // Gets the Card Set Data from Config File to Populate the Template Header
  for (var col = 0; col < 48; col++){
    SetNum = CardDBHeader[0][col];
    // Only executes if the value in the cell is a number between 1 and 8
    if(SetNum >= 1 && SetNum <= 8){
      CardDBHeader[1][col] = cfgSetData[SetNum][2];
      if (col < 32) CardDBHeader[2][col] = cfgSetData[SetNum][3];
      if (col > 32) CardDBHeader[2][col] = cfgSetData[SetNum][4];
    }
  }
        
  // Set Card Set Names and Codes in the template sheet
  shtTemplate.getRange(4,1,4,48).setValues(CardDBHeader);
  // Set Formula to count number of Packs
  frmlNbPack = "=G3/" + cfgCardPack;
  shtTemplate.getRange(3, 11).setValue(frmlNbPack)
  
  // Loops through each player starting from the last
  for (var plyr = NbPlayers; plyr > 0; plyr--){
    // Gets the Player Name 
    PlyrName = PlayerData[plyr][0];
    // Resets the Player Found flag before searching
    PlayerFound = 0;            
    
    // Look if player exists, if yes, skip, if not, create player
    for(var sheet = NbSheet; sheet > 0; sheet --){
      SheetName = ssSheets[sheet-1].getSheetName();
      if (SheetName == PlyrName) PlayerFound = 1;
    }
    
    // If Player is not found, create a tab with the player's name
    if(PlayerFound == 0){
      // Get the Template sheet index
      NbSheet = ssCardDB.getNumSheets();
      // Inserts Sheet before Template (Last Sheet in Spreadsheet)
      ssCardDB.insertSheet(PlyrName, NbSheet-2, {template: shtTemplate});
      shtPlyr = ssCardDB.getSheetByName(PlyrName);
      
      // Opens the new sheet and modify appropriate data (Player Name, Header)
      shtPlyr.getRange(3,3).setValue(PlyrName);
      //shtPlyr.getRange(4,1,4,48).setValues(CardDBHeader);
      
      // Updates the number of sheets before relooping
      NbSheet = ssCardDB.getNumSheets();
      ssSheets = ssCardDB.getSheets();
    }
  }
  shtPlyr = ssCardDB.getSheets()[0];
  ssCardDB.setActiveSheet(shtPlyr);
}


// **********************************************
// function fcnCrtPlayerCardList()
//
// This function generates all accessible Card Lists
// for all players from the Config File
//
// **********************************************

function fcnCrtPlayerCardList(){
  
  Logger.log("Routine: fcnCrtPlayerCardList");
    
  // Main Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shtTest = ss.getSheetByName('Test');
  
  // Config Spreadsheet
  var shtConfig = ss.getSheetByName('Config');
  var shtIDs = shtConfig.getRange(4,7,24,1).getValues();
  
  // Configuration Data
  var cfgColShtPlyr = shtConfig.getRange(4,25,30,1).getValues();
  var cfgCardBuild = shtConfig.getRange(4,30,20,1).getValues();
  
  // Column Values
  var colShtPlyrName = cfgColShtPlyr[2][0];
  
  // Card DB Spreadsheet
  var ssCardDB = SpreadsheetApp.openById(shtIDs[2][0]); 
  
  // Card Lists Spreadsheet
  var ssCardListEN = SpreadsheetApp.openById(shtIDs[3][0]);
  var ssCardListFR = SpreadsheetApp.openById(shtIDs[4][0]);
  var shtTemplateEN = ssCardListEN.getSheetByName('Template');
  var shtTemplateFR = ssCardListFR.getSheetByName('Template');
  var shtCardListNum;
  
  var NbSheet = ssCardListEN.getNumSheets();
  var ssSheets = ssCardListEN.getSheets();
  var SheetName;
  var PlayerFound = 0;
  
  // Players 
  var shtPlayers = ss.getSheetByName('Players'); 
  var NbPlayers = shtPlayers.getRange(2,1).getValue();
  var PlayerData = shtPlayers.getRange(2,colShtPlyrName, NbPlayers+1, 1).getValues();
    
  var shtPlyrEN;
  var shtPlyrFR;
  var PlyrName;
  
  // Loops through each player starting from the last
  for (var plyr = NbPlayers; plyr > 0; plyr--){
    // Gets the Player Name 
    PlyrName = PlayerData[plyr][0];
    // Resets the Player Found flag before searching
    PlayerFound = 0;
    
    // Look if player exists, if yes, skip, if not, create player
    for(var sheet = NbSheet; sheet > 0; sheet --){
      SheetName = ssSheets[sheet-1].getSheetName();
      if (SheetName == PlyrName) PlayerFound = 1;
    }
          
    // If Player is not found, create a tab with the player's name
    if(PlayerFound == 0){

      // Gets the Player Card DB Number of Cards and Boosters
      var shtCardDBPlyr = ssCardDB.getSheetByName(PlyrName);
      var CardTotal =     shtCardDBPlyr.getRange(3, 7).getValue();
      var BstrTotal =     shtCardDBPlyr.getRange(3,11).getValue();
      
      // Inserts Sheet before Template (Last Sheet in Spreadsheet)
      // English Version
      ssCardListEN.insertSheet(PlyrName, NbSheet-1, {template: shtTemplateEN});
      shtPlyrEN = ssCardListEN.getSheetByName(PlyrName).showSheet();
      
      // Opens the new sheet and modify appropriate data (Player Name, Header)
      shtPlyrEN.getRange(2,1).setValue(PlyrName);
      shtPlyrEN.getRange(3,1).setValue(BstrTotal);
      shtPlyrEN.getRange(4,1).setValue(CardTotal);
      
      // French Version
      ssCardListFR.insertSheet(PlyrName, NbSheet-1, {template: shtTemplateFR});
      shtPlyrFR = ssCardListFR.getSheetByName(PlyrName).showSheet();
      // Opens the new sheet and modify appropriate data (Player Name, Header)
      shtPlyrFR.getRange(2,1).setValue(PlyrName);
      shtPlyrFR.getRange(3,1).setValue(BstrTotal);
      shtPlyrFR.getRange(4,1).setValue(CardTotal);

      // Updates the number of sheets before relooping
      NbSheet = ssCardListEN.getNumSheets();
      ssSheets = ssCardListEN.getSheets();
      
      // Call function to generate clean card list from Player Card DB
      fcnUpdateCardList(shtConfig, PlyrName, shtTest);
      
    }
  }
  // Selects the first Tab and hides the Template Tab
  // English Version
  ssCardListEN.setActiveSheet(ssCardListEN.getSheets()[0]);
  ssCardListEN.getSheetByName('Template').hideSheet();
  
  // French Version
  ssCardListFR.setActiveSheet(ssCardListFR.getSheets()[0]);
  ssCardListFR.getSheetByName('Template').hideSheet();
}


// **********************************************
// function fcnCrtPlayerStartPool()
//
// This function generates Starting Pool for all 
// players from the Config File
//
// **********************************************

function fcnCrtPlayerStartPool(){
      
  Logger.log("Routine: fcnCrtPlayerStartPool");
    
  // Main Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ssID = ss.getId();
  
  // Config Spreadsheet
  var shtConfig = ss.getSheetByName('Config');
  var shtIDs = shtConfig.getRange(4,7,24,1).getValues();
  
  // Configuration Data
  var cfgColShtPlyr = shtConfig.getRange(4,25,20,1).getValues();
  
  // Column Values
  var colShtPlyrName = cfgColShtPlyr[2][0];
  var colShtPlyrTeam = cfgColShtPlyr[5][0];
  
  // Starting Pool Spreadsheet
  var ssStartPool = SpreadsheetApp.openById(shtIDs[22][0]);
  var shtTemplate = ssStartPool.getSheetByName('Template');
  var NbSheet = ssStartPool.getNumSheets();
  var ssSheets = ssStartPool.getSheets();
  var SheetName;
  var PlayerFound = 0;

  // Players 
  var shtPlayers = ss.getSheetByName('Players'); 
  var MaxColPlayers = shtPlayers.getMaxColumns();
  var NbPlayers = shtPlayers.getRange(2,1).getValue();
  
  // Get Players Data
  var PlayerData = shtPlayers.getRange(2,colShtPlyrName, NbPlayers+1, MaxColPlayers-1).getValues();
     
  var shtPlyr;
  var PlyrName;
  var ValidSetList = new Array(8);
  
  // Legal Sets Data from Config File
  var cfgSetData =    shtConfig.getRange(4,34,8,1).getValues();
  // [x][0] = Set Name
  // x = Set Number (1-8)
  var cfgNbSet = cfgSetData[0][0];
  
  // Create Data Validation
  for(var i = 0; i < 8; i++){
    ValidSetList[i] = cfgSetData[7-i][0];
  }
    
  var ruleSet = SpreadsheetApp.newDataValidation().requireValueInList(ValidSetList, true).build();
  shtTemplate.getRange(5,2,1,6).setDataValidation(ruleSet);
    
    
  // Loops through each player starting from the last
  for (var plyr = NbPlayers; plyr > 0; plyr--){
    // Gets the Player Name 
    PlyrName = PlayerData[plyr][0];
    // Resets the Player Found flag before searching
    PlayerFound = 0;
    
    // Look if player exists, if yes, skip, if not, create player
    for(var sheet = NbSheet; sheet > 0; sheet --){
      SheetName = ssSheets[sheet-1].getSheetName();
      if (SheetName == PlyrName) PlayerFound = 1;
    }
    
    // If Player is not found, create a tab with the player's name
    if(PlayerFound == 0){
      // Get the Template sheet index
      NbSheet = ssStartPool.getNumSheets();
      // Inserts Sheet before Template (Last Sheet in Spreadsheet)
      ssStartPool.insertSheet(PlyrName, NbSheet-1, {template: shtTemplate});
      shtPlyr = ssStartPool.getSheetByName(PlyrName);
      
      // Opens the new sheet and modify appropriate data (Player Name, Header)
      shtPlyr.getRange(1,2).setValue(PlyrName);
      shtPlyr.getRange(2,2).setValue('Not Processed');
      shtPlyr.getRange(3,2).setValue(ssID);
      
      //Hides the 3rd row
      shtPlyr.hideRows(3);
      
      // Updates the number of sheets before relooping
      NbSheet = ssStartPool.getNumSheets();
      ssSheets = ssStartPool.getSheets();
    }
  }
  
  // Hide Template Sheet
  ssStartPool.setActiveSheet(ssStartPool.getSheets()[0]);
  ssStartPool.getSheetByName('Template').hideSheet();
}

// **********************************************
// function fcnCrtPlayerEscltnBonus()
//
// This function generates the Escalation Bonus 
// for all players from the Config File
//
// **********************************************

function fcnCrtPlayerEscltnBonus(){
      
  Logger.log("Routine: fcnCrtPlayerEscltnBonus");
    
  // Main Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ssID = ss.getId();
  
  // Config Spreadsheet
  var shtConfig = ss.getSheetByName('Config');
  var shtIDs = shtConfig.getRange(4,7,24,1).getValues();
  
  // Configuration Data
  var cfgColShtPlyr = shtConfig.getRange(4,25,20,1).getValues();

  // Column Values
  var colShtPlyrName = cfgColShtPlyr[2][0];
  var colShtPlyrTeam = cfgColShtPlyr[5][0];
  
  // Card DB Spreadsheet
  var ssEscltnBonus = SpreadsheetApp.openById(shtIDs[21][0]);
  var shtTemplate = ssEscltnBonus.getSheetByName('Template');
  var NbSheet = ssEscltnBonus.getNumSheets();
  var ssSheets = ssEscltnBonus.getSheets();
  var SheetName;
  var PlayerFound = 0;

  // Players 
  var shtPlayers = ss.getSheetByName('Players'); 
  var MaxColPlayers = shtPlayers.getMaxColumns();
  var NbPlayers = shtPlayers.getRange(2,1).getValue();
  
  // Get Players Data
  var PlayerData = shtPlayers.getRange(2,colShtPlyrName, NbPlayers+1, MaxColPlayers-1).getValues();
     
  var shtPlyr;
  var PlyrName;
  
  // Loops through each player starting from the last
  for (var plyr = NbPlayers; plyr > 0; plyr--){
    // Gets the Player Name 
    PlyrName = PlayerData[plyr][0];
    // Resets the Player Found flag before searching
    PlayerFound = 0;
    
    // Look if player exists, if yes, skip, if not, create player
    for(var sheet = NbSheet; sheet > 0; sheet --){
      SheetName = ssSheets[sheet-1].getSheetName();
      if (SheetName == PlyrName) PlayerFound = 1;
    }
    
    if (PlayerFound == 0){
      // Get the Template sheet index
      NbSheet = ssEscltnBonus.getNumSheets();
      // INSERTS TAB BEFORE "Template" TAB
      // English Version
      ssEscltnBonus.insertSheet(shtPlyr, NbSheet-1, {template: shtTemplate});
      shtPlyr = ssEscltnBonus.getSheetByName(shtPlyr).showSheet();
      
      // Opens the new sheet and modify appropriate data (Player Name, Header)
      shtPlyr.getRange(1,1).setValue(PlyrName);
    }
  }
  
  // Hide Template Sheet
  ssEscltnBonus.setActiveSheet(ssEscltnBonus.getSheets()[0]);
  ssEscltnBonus.getSheetByName('Template').hideSheet();
}


// **********************************************
// function fcnDelPlayerCardDB()
//
// This function deletes all Players Card DB Sheets 
// from the Config File
//
// **********************************************

function fcnDelPlayerCardDB(){
  
  Logger.log("Routine: fcnDelPlayerCardDB");

  // Main Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Config Spreadsheet
  var shtConfig = ss.getSheetByName('Config');
  
  // Sheet IDs
  var shtIDs = shtConfig.getRange(4,7,24,1).getValues();
   
  // Template Sheet
  var ssDel = SpreadsheetApp.openById(shtIDs[2][0]);
  var shtTemplate = ssDel.getSheetByName('Template').showSheet();
    
  // Delete Player Card DB
  subDelPlayerSheets(shtIDs[2][0]);
 }


// **********************************************
// function fcnDelPlayerCardList()
//
// This function deletes all Players Card List Sheets
// from the Config File
//
// **********************************************

function fcnDelPlayerCardList(){
  
  Logger.log("Routine: fcnDelPlayerCardList");

  // Main Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Config Spreadsheet
  var shtConfig = ss.getSheetByName('Config');
  
  // Sheet IDs
  var shtIDs = shtConfig.getRange(4,7,24,1).getValues();
  
  // Template Sheet
  var ssDel = SpreadsheetApp.openById(shtIDs[3][0]);
  var shtTemplate = ssDel.getSheetByName('Template').showSheet();
  
  // Template Sheet
  var ssDel = SpreadsheetApp.openById(shtIDs[4][0]);
  var shtTemplate = ssDel.getSheetByName('Template').showSheet();
  
  // Delete Players Card Lists EN
  subDelPlayerSheets(shtIDs[3][0]);
  
  // Delete Players Card Lists FR
  subDelPlayerSheets(shtIDs[4][0]);
}

// **********************************************
// function fcnDelPlayerStartPool()
//
// This function deletes all Players Starting Pool Sheets
// from the Config File
//
// **********************************************

function fcnDelPlayerStartPool(){
  
  Logger.log("Routine: fcnDelPlayerStartPool");

  // Main Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Config Spreadsheet
  var shtConfig = ss.getSheetByName('Config');
  
  // Sheet IDs
  var shtIDs = shtConfig.getRange(4,7,24,1).getValues();
  
  // Template Sheet
  var ssDel = SpreadsheetApp.openById(shtIDs[22][0]);
  var shtTemplate = ssDel.getSheetByName('Template').showSheet();
  
  // Delete Players Starting Pool
  subDelPlayerSheets(shtIDs[22][0]);
}

// **********************************************
// function fcnDelPlayerEscltnBonus()
//
// This function deletes all Players Escalation Bonus Sheets
// from the Config File
//
// **********************************************

function fcnDelPlayerEscltnBonus(){
  
  Logger.log("Routine: fcnDelPlayerEscltnBonus");

  // Main Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Config Spreadsheet
  var shtConfig = ss.getSheetByName('Config');
  
  // Sheet IDs
  var shtIDs = shtConfig.getRange(4,7,24,1).getValues();
  
  // Template Sheet
  var ssDel = SpreadsheetApp.openById(shtIDs[21][0]);
  var shtTemplate = ssDel.getSheetByName('Template').showSheet();
  
  // Delete Players Escalation Bonus Sheets
  subDelPlayerSheets(shtIDs[21][0]);
}
