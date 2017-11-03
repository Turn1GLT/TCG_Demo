// **********************************************
// function fcnRegistrationTCG
//
// This function adds the new player to
// the Player's List and calls other functions
// to create its complete profile
//
// **********************************************

function fcnRegistrationTCG(shtResponse, RowResponse){

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var shtConfig = ss.getSheetByName('Config');
  var shtPlayers = ss.getSheetByName('Players');
  
  var PlayerData = new Array(10);
  PlayerData[0] = 0 ; // Function Status
  PlayerData[1] = ''; // Number of Players
  PlayerData[2] = ''; // New Player Full Name
  PlayerData[3] = ''; // New Player Email
  PlayerData[4] = ''; // New Player Language
  PlayerData[5] = ''; // New Player Phone Number
  PlayerData[6] = ''; // New Player DCI Number
  PlayerData[7] = ''; // New Player Team Name
  PlayerData[8] = ''; // New Player Spare
  PlayerData[9] = ''; // New Player Spare
  
  // Add Player to Player List
  PlayerData = fcnAddPlayerTCG(shtConfig, shtPlayers, shtResponse, RowResponse, PlayerData);
  var NbPlayers  = PlayerData[1];
  var PlayerName = PlayerData[2];
  
  // If Player was succesfully added, Generate Card DB, Generate Card Pool, Generate Startin Pool, Modify Match Report Form and Add Player to Weekly Booster
  if(PlayerData[0] == "New Player") {
    fcnGenPlayerCardDB();
    Logger.log('Card Database Generated'); 
    fcnGenPlayerCardList();
    Logger.log('Card Pool Generated');
    fcnGenPlayerStartPoolMain();
    Logger.log('Starting Pool Generated');   
    fcnGenPlayerWeekBstr();
    Logger.log('Weekly Booster Generated');    
    fcnModifyReportFormTCG(ss, shtConfig, shtPlayers);
    
    // Execute Ranking function in Standing tab
    fcnUpdateStandings(ss, shtConfig);
    
    // Copy all data to Standing League Spreadsheet
    fcnCopyStandingsSheets(ss, shtConfig, 0, 1);
    
    // Send Confirmation to New Player
    fcnSendNewPlayerConf(shtConfig, PlayerData);
    Logger.log('Confirmation Email Sent');
    
    // Send Log for new Registration
    var recipient = 'turn1glt@gmail.com';
    var subject = 'New Player Registration: ' + PlayerName;
    var body = Logger.getLog();
    MailApp.sendEmail(recipient, subject, body);
    
    // Send Confirmation to Location
    // fcnSendNewPlayerConfLocation(shtConfig, PlayerData)
  }
  // If Player is not a new player, send error to turn1glt@gmail.com
  if(PlayerData[0] != "New Player") {
    // Send Log for new Registration
    var recipient = 'turn1glt@gmail.com';
    var subject = 'Registration Error: ' + PlayerName;
    var body = Logger.getLog();
    MailApp.sendEmail(recipient, subject, body);
  }
}




// **********************************************
// function fcnAddPlayerTCG
//
// This function adds the new player to
// the Player's List
//
// **********************************************

function fcnAddPlayerTCG(shtConfig, shtPlayers, shtResponse, RowResponse, PlayerData) {

  // Opens Players List File
  var ssPlayersListID = shtConfig.getRange(40,2).getValue();
  var ssPlayersList = SpreadsheetApp.openById(ssPlayersListID);
  var shtPlayersList = ssPlayersList.getSheetByName('Players');
  
  // Current Player List
  var NbPlayers = shtPlayers.getRange(2,1).getValue();
  var NextPlayerRow = NbPlayers + 3;
  var CurrPlayers = shtPlayers.getRange(2, 2, NbPlayers+1, 1).getValues();
  var Status = "New Player";
  
  // Response Columns from Configuration File
  // [x][0] = Response Columns
  // [x][1] = Players Table Columns in Main Sheet
  var colRegRespValues = shtConfig.getRange(56,7,12,2).getValues();
  
  // Response Columns
  var colRespEmail = colRegRespValues[0][0];
  var colRespName = colRegRespValues[1][0];
  var colRespFirstName = colRegRespValues[2][0];
  var colRespLastName = colRegRespValues[3][0];
  var colRespPhone = colRegRespValues[4][0];
  var colRespLanguage = colRegRespValues[5][0];
  var colRespDCI = colRegRespValues[6][0];
  var colRespTeamName = colRegRespValues[7][0];
  
  // Player Table Columns
  var colTableEmail = colRegRespValues[0][1];
  var colTableName = colRegRespValues[1][1];
  var colTableFirstName = colRegRespValues[2][1];
  var colTableLastName = colRegRespValues[3][1];
  var colTablePhone = colRegRespValues[4][1];
  var colTableLanguage = colRegRespValues[5][1];
  var colTableDCI = colRegRespValues[6][1];
  var colTableTeamName = colRegRespValues[7][1];
  
  // Get All Values from Response Sheet
  var shtRespMaxCol = shtResponse.getMaxColumns();
  var Responses = shtResponse.getRange(RowResponse,1,1,shtRespMaxCol).getValues();
  
  // Email
  var EmailAddress = Responses[0][colRespEmail];
  Logger.log(EmailAddress);
  
  // Player First Name / Last Name (if used)
  if(colRespFirstName != '' && colRespFirstName != 0 && colRespLastName != '' && colRespLastName != 0){
    var FirstName = Responses[0][colRespFirstName];
    var LastName = Responses[0][colRespLastName];
    var PlayerName = FirstName + ' ' + LastName;
    Logger.log(PlayerName);
  }
  
  // Player Full Name (if used)
  if(colRespName != '' && colRespName != 0) {
    var PlayerName = Responses[0][colRespName];
    Logger.log(PlayerName);
  }
  
  // Player Language Preference
  var Language = Responses[0][colRespLanguage];
  Logger.log(Language);  
  
  // Player Phone Number
  if(colRespPhone != '' && colRespPhone != 0) {
    var Phone = Responses[0][colRespPhone];
    Logger.log(Phone);
  }
  
  // Player DCI Number 
  if(colRespDCI != '' && colRespDCI != 0) {
    var DCINum = Responses[0][colRespDCI];
    Logger.log(DCINum);
  }
  
  // Team Name
  if(colRespTeamName != '' && colRespTeamName != 0) {
    var TeamName = Responses[0][colRespTeamName];
    Logger.log(TeamName);
  }
  
  // Check if Player exists in List
  for(var i = 1; i <= NbPlayers; i++){
    if(PlayerName == CurrPlayers[i][0]){
      Status = "Cannot complete registration for " + PlayerName + ", Duplicate Player Found in List";
      Logger.log(DuplicatePlyr)
    }
  }

  // Copy Values to Players Sheet at the Next Empty Spot (Number of Players + 3)
  // Copy Values to Players List for Store Access
  if(Status == "New Player"){
	// Name
    shtPlayers.getRange(NextPlayerRow, colTableName).setValue(PlayerName);
    shtPlayersList.getRange(NextPlayerRow, colTableName).setValue(PlayerName);
    Logger.log('Player Name: %s',PlayerName);
    // Email Address
    shtPlayers.getRange(NextPlayerRow, colTableEmail).setValue(EmailAddress);
    shtPlayersList.getRange(NextPlayerRow, colTableEmail).setValue(EmailAddress);
    Logger.log('Email Address: %s',EmailAddress);
    // Language
    shtPlayers.getRange(NextPlayerRow, colTableLanguage).setValue(Language);
    shtPlayersList.getRange(NextPlayerRow, colTableLanguage).setValue(Language);
    Logger.log('Language: %s',Language);
    // Phone Number
    if(colTablePhone != '' && colTablePhone != 0){
	shtPlayers.getRange(NextPlayerRow, colTablePhone).setValue(Phone);
    shtPlayersList.getRange(NextPlayerRow, colTablePhone).setValue(Phone);
    Logger.log('Phone: %s',Phone);  
    }
	// DCI Number
    if(colTableDCI != '' && colTableDCI != 0){
	shtPlayers.getRange(NextPlayerRow, colTableDCI).setValue(DCINum);
    shtPlayersList.getRange(NextPlayerRow, colTableDCI).setValue(DCINum);
    Logger.log('DCI: %s',DCINum);  Logger.log('-----------------------------');
  }
	// Team Name
    if(colTableTeamName != '' && colTableTeamName != 0){
	shtPlayers.getRange(NextPlayerRow, colTableTeamName).setValue(TeamName);
    shtPlayersList.getRange(NextPlayerRow, colTableTeamName).setValue(TeamName);
    Logger.log('Team Name: %s',TeamName);  Logger.log('-----------------------------');
	}
  }
  PlayerData[0] = Status;
  PlayerData[1] = NbPlayers + 1;
  PlayerData[2] = PlayerName;
  PlayerData[3] = EmailAddress;
  PlayerData[4] = Language;
  PlayerData[5] = Phone;
  PlayerData[6] = DCINum;
  PlayerData[7] = TeamName;
  
  return PlayerData;
}


// **********************************************
// function fcnModifyReportFormTCG
//
// This function modifies the Match Report Form
// to add new added players
//
// **********************************************

function fcnModifyReportFormTCG(ss, shtConfig, shtPlayers) {

  var MatchFormEN = FormApp.openById(shtConfig.getRange(36, 2).getValue());
  var MatchFormItemEN = MatchFormEN.getItems();
  var MatchFormFR = FormApp.openById(shtConfig.getRange(37, 2).getValue());
  var MatchFormItemFR = MatchFormFR.getItems();
  var NbMatchFormItem = MatchFormItemFR.length;
  
  var WeekBstrFormEN = FormApp.openById(shtConfig.getRange(42, 2).getValue());
  var WeekBstrFormItemEN = WeekBstrFormEN.getItems();
  var WeekBstrFormFR = FormApp.openById(shtConfig.getRange(43, 2).getValue());
  var WeekBstrFormItemFR = WeekBstrFormFR.getItems();
  var NbWeekBstrFormItem = WeekBstrFormItemFR.length;

  // Function Variables
  var ItemTitle;
  var ItemPlayerListEN;
  var ItemPlayerListFR;
  var ItemPlayerChoice;
  
  var NbPlayers = shtPlayers.getRange(2, 1).getValue();
  var Players = shtPlayers.getRange(3, 2, NbPlayers, 1).getValues();
  var ListPlayers = [];
  
  // Loops in Match Form to Find Players List
  for(var item = 0; item < NbMatchFormItem; item++){
    ItemTitle = MatchFormItemEN[item].getTitle();
    if(ItemTitle == 'Winning Player' || ItemTitle == 'Losing Player'){
      
      // Get the List Item from the Match Report Form
      ItemPlayerListEN = MatchFormItemEN[item].asListItem();
      ItemPlayerListFR = MatchFormItemFR[item].asListItem();
      
      // Build the Player List from the Players Sheet     
      for (i = 0; i < NbPlayers; i++){
        ListPlayers[i] = Players[i][0];
      }
      // Set the Player List to the Match Report Forms
      ItemPlayerListEN.setChoiceValues(ListPlayers);
      ItemPlayerListFR.setChoiceValues(ListPlayers);
    }
  }
  
  // Loops in Weekly Booster Form to Find Players List
  for(var item = 0; item < NbWeekBstrFormItem; item++){
    ItemTitle = WeekBstrFormItemEN[item].getTitle();
    if(ItemTitle == 'Winning Player' || ItemTitle == 'Losing Player'){
      
      // Get the List Item from the Weekly Booster Report Form
      ItemPlayerListEN = WeekBstrFormItemEN[item].asListItem();
      ItemPlayerListFR = WeekBstrFormItemFR[item].asListItem();
      
      // Build the Player List from the Players Sheet     
      for (i = 0; i < NbPlayers; i++){
        ListPlayers[i] = Players[i][0];
      }
      // Set the Player List to the Weekly Booster Report Forms
      ItemPlayerListEN.setChoiceValues(ListPlayers);
      ItemPlayerListFR.setChoiceValues(ListPlayers);
    }
  }
}
