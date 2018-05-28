// **********************************************
// function fcnRegistrationPlyrTCG
//
// This function adds the new player to
// the Player's List and calls other functions
// to create its complete profile
//
// **********************************************

function fcnRegistrationPlyrTCG(shtResponse, RowResponse){

  Logger.log("Routine: fcnRegistrationPlyrTCG");
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shtConfig = ss.getSheetByName("Config");
  var shtPlayers = ss.getSheetByName("Players");
  var shtPlayersMaxCol = shtPlayers.getMaxColumns();
  var shtPlayersMaxRow = shtPlayers.getMaxRows();
  
  var shtIDs = shtConfig.getRange(4,7,24,1).getValues();
  var cfgEvntParam = shtConfig.getRange(4,4,48,1).getValues();
  var cfgColRspSht = shtConfig.getRange(4,15,16,1).getValues();
  var cfgColRndSht = shtConfig.getRange(4,18,16,1).getValues();
  var cfgExecData  = shtConfig.getRange(4,21,16,1).getValues();
  
  // Player Registration Form Construction 
  // Column 1 = Category Name
  // Column 2 = Category Order in Form
  // Column 3 = Column Value in Player/Team Sheet
  var cfgRegFormCnstrVal = shtConfig.getRange(4,23,20,3).getValues();
  
  // Log Sheet
  var shtLog = SpreadsheetApp.openById(shtIDs[1][0]).getSheetByName("Log");

  // Player Info Store Sheet
  var shtStrPlayers = SpreadsheetApp.openById(shtIDs[10][0]).getSheetByName("Players");
  
  // Execution Parameters
  var exeMemberLink = cfgExecData[7][0];
  
  // Event Parameters
  var evntEscalation =  cfgEvntParam[19][0];
  var evntLogCardList = cfgEvntParam[44][0];
  
  // Match Report Form IDs
  var MatchFormIdEN = shtIDs[11][0];
  var MatchFormIdFR = shtIDs[12][0];
  
  // Create Member 
  var Member = subCreateArray(16,1);
  //  Member[ 0] = Member ID
  //  Member[ 1] = Member Record File ID
  //  Member[ 2] = Member Record File Link
  //  Member[ 3] = Member Full Name
  //  Member[ 4] = Member First Name
  //  Member[ 5] = Member Last Name
  //  Member[ 6] = Member Email
  //  Member[ 7] = Member Language
  //  Member[ 8] = Member Phone Number
  //  Member[ 9] = Member DCI
  //  Member[10] = Member Spare
  //  Member[11] = Member Spare
  //  Member[12] = Member Spare
  //  Member[13] = Member Spare
  //  Member[14] = Member Spare
  //  Member[15] = Member Spare
  
  var memberFullName;
  var memberFileID;
  
  // Log new Registration
  Logger.log( "------- New Player Registration -------");

  // Get All Values from Response Sheet
  var shtRespMaxCol = shtResponse.getMaxColumns();
  var RegRspnVal = shtResponse.getRange(RowResponse,1,1,shtRespMaxCol).getValues();
  
  // Add Player to Player List
  Member = fcnAddPlayerTCG(shtIDs, shtConfig, shtPlayers, RegRspnVal, cfgEvntParam, cfgRegFormCnstrVal, Member);
  memberFullName = Member[3];
  
  // If Player was succesfully added, the Full Name will be created, then execute the following
  if(memberFullName != "") {
    
    // If Link to Membership is Enabled
    if(exeMemberLink == "Enabled"){
      // Search if Player is Member of Turn1 GLT
      Member = fcnSearchMember(Member);
      memberFileID = Member[1];
      
      if(memberFileID != "Member Not Found") Logger.log("Member %s already exists",memberFullName);
      Logger.log("Member File ID: %s",memberFileID);
      // If the Member Record File does not exist, the Player is not a member, create it 
      if(memberFileID == "Member Not Found") {
        Member = fcnCreateMember(Member);
        memberFileID = Member[1];
        if(memberFileID != "Member Not Found") Logger.log("Member %s created",memberFullName);
      }
      // Update Player File ID in Player Sheet
      subUpdatePlayerMember(shtConfig, shtPlayers, Member);
    }
    
    // Create Player Army DB
    if(evntLogCardList == "Enabled"){
      fcnCrtPlayerCardDB();
      Logger.log("Card Database Generated");
      
      // Create Player Card Lists (Player Access)
      fcnCrtPlayerCardList();
      Logger.log("Card List Generated");  
    }
    
    // Create Player Event Record (Player Access)
    fcnCrtEvntRecord();
    Logger.log("Player Record Generated");  
    
    // If Escalation is Enabled, Create Player Escalation Bonus sheet 
    if(evntEscalation == "Enabled"){
      fcnCrtPlayerEscltBonus();
      Logger.log("Round Unit Sheet Generated");   
    }
    // Add Player to Match Report Forms
    if(MatchFormIdEN != "" && MatchFormIdFR != ""){
      fcnModifyReportFormTCG(shtConfig, shtIDs, shtPlayers, cfgEvntParam, evntEscalation);
      Logger.log("Match Report Form Updated");  
    }
    
    // Execute Ranking function in Standing tab
    fcnUpdateStandings();
      Logger.log("Overall Standings Updated");  
    
    // Copy all data to Standing League Spreadsheet
    fcnCopyStandingsSheets(ss, shtConfig, cfgEvntParam, cfgColRndSht, 0, 1);
      Logger.log("Standing Sheets Updated");  
    
    // Send Confirmation to New Player
    //fcnSendNewPlayerConf(shtConfig, PlayerData);
    //Logger.log("Confirmation Email Sent");
    
    // Send Confirmation to Organizer
    // fcnSendNewPlayerConfOrgnzr(shtConfig, PlayerData)
    
    // Copy Main File Players List to Store Players List
    var MainPlayers = shtPlayers.getRange(3,2,shtPlayersMaxRow-2,shtPlayersMaxCol-1).getValues();
    shtStrPlayers.getRange(3,2,shtPlayersMaxRow-2,shtPlayersMaxCol-1).setValues(MainPlayers);
    //rngPlayers.copyTo(rngStrPlayers);
    
  }

  // Post Log to Log Sheet
  subPostLog(shtLog, Logger.getLog());
}


// **********************************************
// function fcnAddPlayerTCG
//
// This function adds the new player to
// the Player's List
//
// **********************************************

function fcnAddPlayerTCG(shtIDs, shtConfig, shtPlayers, RegRspnVal, cfgEvntParam, cfgRegFormCnstrVal, Member) {
  
  // Current Player List
  var NbPlayers = shtPlayers.getRange(2,1).getValue();
  var NextPlayerRow = NbPlayers + 3;
  var CurrPlayers = shtPlayers.getRange(2, 2, NbPlayers+1, 1).getValues();
  var Status = "New Player";
  
  // Event Properties
  var evntFormat =     cfgEvntParam[9][0];
  var evntNbPlyrTeam = cfgEvntParam[10][0];
  
  // Response Columns
  var colRspEmail =        cfgRegFormCnstrVal[ 1][1];
  var colRspFullName =     cfgRegFormCnstrVal[ 2][1];
  var colRspFrstName =     cfgRegFormCnstrVal[ 3][1];
  var colRspLastName =     cfgRegFormCnstrVal[ 4][1];
  var colRspLanguage =     cfgRegFormCnstrVal[ 5][1];
  var colRspPhone =        cfgRegFormCnstrVal[ 6][1];
  var colRspTeamName =     cfgRegFormCnstrVal[ 7][1];
  var colRspDCI =          cfgRegFormCnstrVal[ 8][1];
  var colRspDeckDef =      cfgRegFormCnstrVal[ 9][1];
  var colRspDeckCmdr =     cfgRegFormCnstrVal[10][1];
  var colRspDeckList =     cfgRegFormCnstrVal[11][1];
  
  // Team Table Columns
  var colTblEmail =        cfgRegFormCnstrVal[ 1][2];
  var colTblFullName =     cfgRegFormCnstrVal[ 2][2];
  var colTblFrstName =     cfgRegFormCnstrVal[ 3][2];
  var colTblLastName =     cfgRegFormCnstrVal[ 4][2];
  var colTblLanguage =     cfgRegFormCnstrVal[ 5][2];
  var colTblPhone =        cfgRegFormCnstrVal[ 6][2];
  var colTblTeamName =     cfgRegFormCnstrVal[ 7][2];
  var colTblDCI =          cfgRegFormCnstrVal[ 8][2];
  var colTblDeckDef =      cfgRegFormCnstrVal[ 9][2];
  var colTblDeckCmdr =     cfgRegFormCnstrVal[10][2];
  var colTblDeckList =     cfgRegFormCnstrVal[11][2];
  
  var colTblStatus =          cfgRegFormCnstrVal[16][2];
  var colTblMemberFileID =    cfgRegFormCnstrVal[17][2];
  var colTblEmailContact =    cfgRegFormCnstrVal[18][2];
  var colTblEmailContactGrp = cfgRegFormCnstrVal[19][2];
  
  // Routine Variables
  var PlyrEmail = "";
  var PlyrFullName = "";
  var PlyrFrstName = "";
  var PlyrLastName = "";
  var PlyrLanguage = "";
  var PlyrPhone = "";
  var PlyrDCI = "";
  var PlyrDeckCmdr = "";
  var PlyrDeckList = "";
  var PlyrTeamName = "";
  var PlyrTeamMember1 = "";
  var PlyrTeamMember2 = "";
  var PlyrTeamMember3 = "";
  var PlyrTeamMember4 = "";
  
  var DeckDefOffset = 2;
  
  var ContactInfo = new Array(4); // [0]= First Name, [1]= Last Name, [2]= Email Address, [3]= Language Preference
  
  // Email
  PlyrEmail = RegRspnVal[0][colRspEmail-1];
  
  // Player First and Last Name
  PlyrFrstName = RegRspnVal[0][colRspFrstName-1];
  PlyrLastName = RegRspnVal[0][colRspLastName-1];
  
  // Create Full Name
  PlyrFullName = PlyrFrstName + " " + PlyrLastName;
  
  // Player Language Preference
  PlyrLanguage = RegRspnVal[0][colRspLanguage-1];
  
  // Player Phone Number
  if(colRspPhone != "") PlyrPhone = RegRspnVal[0][colRspPhone-1];

  // Player DCI
  if(colRspDCI != "") PlyrDCI = RegRspnVal[0][colRspDCI-1];
  
  // Team Name
  if(colRspTeamName != "") PlyrTeamName = RegRspnVal[0][colRspTeamName-1];
  
  // Player Deck List
  // Player Deck List Definition
  if(colRspDeckCmdr != "") {
    PlyrDeckCmdr = RegRspnVal[0][colRspDeckCmdr-DeckDefOffset];
    Logger.log("PlyrArmyWarlord: %s",PlyrDeckCmdr);
  }
    
  // Check if Player exists in List
  for(var i = 1; i <= NbPlayers; i++){
    if(PlyrFullName == CurrPlayers[i][0]){
      Status = "Cannot complete registration for " + PlyrFullName + ", Duplicate Player Found in List";
      Logger.log(Status)
    }
  }
  // If New Player
  // Copy Values to Players Sheet at the Next Empty Spot (Number of Players + 3)
  // Copy Values to Players List for Store Access
  if(Status == "New Player"){
    
    // Player Full Name
    shtPlayers.getRange(NextPlayerRow, colTblFullName).setValue(PlyrFullName);
    Logger.log("Player Name: %s",PlyrFullName);
    
    // Email Address
    shtPlayers.getRange(NextPlayerRow, colTblEmail).setValue(PlyrEmail);
    Logger.log("Email Address: %s",PlyrEmail);
    
    // Language
    shtPlayers.getRange(NextPlayerRow, colTblLanguage).setValue(PlyrLanguage);
    Logger.log("Language: %s",PlyrLanguage);
    
    // Phone Number
    if(PlyrPhone != ""){
      shtPlayers.getRange(NextPlayerRow, colTblPhone).setValue(PlyrPhone);
      Logger.log("Phone: %s",PlyrPhone); 
    }
    // DCI Number
    if(PlyrDCI != ""){
      shtPlayers.getRange(NextPlayerRow, colTblDCI).setValue(PlyrDCI);
      Logger.log("Player DCI: %s",PlyrDCI); 
    }
    Logger.log("-----------------------------");
	
    // Team Name
    if(PlyrTeamName != ""){
      shtPlayers.getRange(NextPlayerRow, colTblTeamName).setValue(PlyrTeamName);
      Logger.log("Team Name: %s",PlyrTeamName);  
	}
    
    // Deck Commander
    if(PlyrDeckCmdr != ""){
      shtPlayers.getRange(NextPlayerRow, colTblDeckCmdr).setValue(PlyrDeckCmdr);
      Logger.log("Commander: %s",PlyrDeckCmdr);  
    }
   
    // Deck List
    if(PlyrDeckList != ""){
      // INSERT NEW FUNCTION TO PROCESS DECK LIST INFORMATION
      // fcnProcessDeckList();
      shtPlayers.getRange(NextPlayerRow, colTblDeckList).setValue(PlyrDeckList);
      Logger.log("Deck List: %s",PlyrDeckList);  
    }
    Logger.log("-----------------------------");

    // Set Player Contact Info 
    ContactInfo[0]= PlyrFrstName;
    ContactInfo[1]= PlyrLastName;
    ContactInfo[2]= PlyrEmail;
    ContactInfo[3]= PlyrLanguage;
    
    // Add Player Info to Contact List and Contact Group
    var CntctStatus = subCrtContact(ContactInfo);
    if(CntctStatus == "Contact Created" || CntctStatus == "Contact Updated") {
      // Set Contact Created in Players Sheet
      if(CntctStatus == "Contact Created") shtPlayers.getRange(NextPlayerRow, colTblEmailContact).setValue("Created");
      if(CntctStatus == "Contact Updated") shtPlayers.getRange(NextPlayerRow, colTblEmailContact).setValue("Updated");
      // Add to Contact Group   
      var CntctGrpStatus = subAddToContactGroup(shtConfig, ContactInfo);
      if(CntctGrpStatus == "Contact added to Contact Group") {
        // Set Added in Contact Group in Players Sheet
        shtPlayers.getRange(NextPlayerRow, colTblEmailContactGrp).setValue("Added");
        Logger.log("Contact Added to Contact Group");
      }
      else Logger.log("Contact NOT Added to Contact Group");
    }
    else Logger.log("Contact NOT Created");
   
  }
  
  // Update Member Data
  Member[ 0] = "";           // Member ID
  Member[ 1] = "";           // Member Record File ID
  Member[ 2] = "";           // Member Record File Link
  Member[ 3] = PlyrFullName; // Member Full Name
  Member[ 4] = PlyrFrstName; // Member First Name
  Member[ 5] = PlyrLastName; // Member Last Name
  Member[ 6] = PlyrEmail;    // Member Email
  Member[ 7] = PlyrLanguage; // Member Language
  Member[ 8] = PlyrPhone;    // Member Phone Number
  Member[ 9] = PlyrDCI;      // Member DCI
  Member[10] = "";           // Member Spare
  Member[11] = "";           // Member Spare
  Member[12] = "";           // Member Spare
  Member[13] = "";           // Member Spare
  Member[14] = "";           // Member Spare
  Member[15] = "";           // Member Spare
  
  return Member;
}

// **********************************************
// function fcnRegistrationTeamTCG
//
// This function adds the new player to
// the Player's List and calls other functions
// to create its complete profile
//
// **********************************************

function fcnRegistrationTeamTCG(shtResponse, RowResponse){

  Logger.log("Routine: fcnRegistrationTeamTCG");
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shtConfig = ss.getSheetByName("Config");
  var shtTeams = ss.getSheetByName("Teams");
  var shtPlayers = ss.getSheetByName("Players");
  var shtPlayersMaxCol = shtPlayers.getMaxColumns();
  var shtPlayersMaxRow = shtPlayers.getMaxRows();
  
  var shtIDs = shtConfig.getRange(4,7,24,1).getValues();  
  var cfgEvntParam =    shtConfig.getRange( 4, 4,48,1).getValues();
  var cfgColRspSht =    shtConfig.getRange( 4,15,16,1).getValues();
  var cfgColRndSht =    shtConfig.getRange( 4,18,16,1).getValues();
  var cfgExecData  =    shtConfig.getRange( 4,21,16,1).getValues();
  var cfgColMatchRep =  shtConfig.getRange( 4,28,20,1).getValues();
  var cfgColMatchRslt = shtConfig.getRange(21,15,32,1).getValues();
  
  // Team Registration Form Construction 
  // Column 1 = Category Name
  // Column 2 = Category Order in Form
  // Column 3 = Column Value in Player/Team Sheet
  var cfgRegFormCnstrVal = shtConfig.getRange(24,23,20,3).getValues();
  
  // Log Sheet
  var shtLog = SpreadsheetApp.openById(shtIDs[1][0]).getSheetByName("Log");
  
  // Opens Store Players/Teams List File
  var ssStrPlyrInfo = SpreadsheetApp.openById(shtIDs[10][0]);
  var shtStrPlayers = ssStrPlyrInfo.getSheetByName("Players");
  var shtStrTeams =   ssStrPlyrInfo.getSheetByName("Teams");
  
  // Execution Parameters
  var exeMemberLink = cfgExecData[7][0];
  
  // Event Parameters
  var evntEscalation =  cfgEvntParam[19][0];
  var evntLogCardList = cfgEvntParam[44][0];
  
  // Match Report Form IDs
  var MatchFormIdEN = shtIDs[11][0];
  var MatchFormIdFR = shtIDs[12][0];
  
  // Create Team 
  var Team = subCreateArray(24,1);
  //  Team[ 0] = Team ID
  //  Team[ 1] = Team Record File ID
  //  Team[ 2] = Team Record File Link
  //  Team[ 3] = Team Name
  //  Team[ 4] = Team Member 1
  //  Team[ 5] = Team Member 2
  //  Team[ 6] = Team Member 3
  //  Team[ 7] = Team Member 4
  //  Team[ 8] = Team Member 5
  //  Team[ 9] = Team Member 6
  //  Team[10] = Team Member 7
  //  Team[11] = Team Member 8
  //  Team[12] = Team Contact First Name
  //  Team[13] = Team Contact Last Name 
  //  Team[14] = Team Contact Email 
  //  Team[15] = Team Contact Language 
  //  Team[16] = Team Contact Phone Number 
  //  Team[17] = Team Spare
  //  Team[18] = Team Spare
  //  Team[19] = Team Spare
  //  Team[20] = Team Spare
  //  Team[21] = Team Spare
  //  Team[22] = Team Spare
  //  Team[23] = Team Spare
  
  var teamName;
  var teamID;
  
  // Log new Registration
  Logger.log( "------- New Team Registration -------");

  // Get All Values from Response Sheet
  var shtRespMaxCol = shtResponse.getMaxColumns();
  var RegRspnVal = shtResponse.getRange(RowResponse,1,1,shtRespMaxCol).getValues();
  
  // Add Team to Team List
  Team = fcnAddTeamTCG(shtIDs, shtConfig, shtTeams, RegRspnVal, cfgEvntParam, cfgRegFormCnstrVal, Team);
  teamName = Team[3];
    
  // If Team was succesfully added, the Team Name will be created, then execute the following
  if(teamName != "") {
    
    // Create Team Event Record (Player Access)
    //fcnCrtEvntRecord();
    Logger.log("Team Record Generated");  
    
    // If Escalation is Enabled, Create Player Escalation Bonus sheet 
    if(evntEscalation == "Enabled"){
      fcnCrtPlayerEscltBonus();
      Logger.log("Round Unit Sheet Generated");   
    }
    // Add Team to Match Report Forms
    if(MatchFormIdEN != "" && MatchFormIdFR != ""){
      fcnModifyReportFormTCG(shtConfig, shtIDs, shtTeams, cfgEvntParam, evntEscalation);
      Logger.log("Match Report Form Updated");  
    }
    
    // Execute Ranking function in Standing tab
    fcnUpdateStandings();
      Logger.log("Overall Standings Updated");  
    
    // Copy all data to Standing League Spreadsheet
    fcnCopyStandingsSheets(ss, shtConfig, cfgEvntParam, cfgColRndSht, 0, 1);
      Logger.log("Standing Sheets Updated");  
    
    // Send Confirmation to Team Contact
    //fcnSendNewPlayerConf(shtConfig, PlayerData);
    //Logger.log("Confirmation Email Sent");
    
    // Send Confirmation to Organizer
    // fcnSendNewPlayerConfOrgnzr(shtConfig, PlayerData)
    
    // Copy Main File Players List to Store Players List
    var rngPlayers = shtPlayers.getRange(3,2,shtPlayersMaxRow-2,shtPlayersMaxCol-1);
    var rngStrPlayers = shtStrPlayers.getRange(3,2,shtPlayersMaxRow-2,shtPlayersMaxCol-1);
    rngPlayers.copyTo(rngStrPlayers);
  }

  // Post Log to Log Sheet
  subPostLog(shtLog, Logger.getLog());
}

// **********************************************
// function fcnAddTeamTCG
//
// This function adds the new Team to
// the Team's List
//
// **********************************************

function fcnAddTeamTCG(shtIDs, shtConfig, shtTeams, RegRspnVal, cfgEvntParam, cfgRegFormCnstrVal, Team) {

  // Current Team List
  var NbTeams = shtTeams.getRange(2,1).getValue();
  var NextTeamRow = NbTeams + 3;
  var CurrTeams = shtTeams.getRange(2, 2, NbTeams+1, 1).getValues();
  var Status = "New Team";
  
  // Event Properties
  var evntFormat = cfgEvntParam[9][0];      // Single, Team or Team+Players
  var evntNbPlyrTeam = cfgEvntParam[10][0]; // Nb Player per Team
  
  // Response Columns
  var colRspCntctEmail =    cfgRegFormCnstrVal[ 1][1];
  var colRspCntctFullName = cfgRegFormCnstrVal[ 2][1];
  var colRspCntctFrstName = cfgRegFormCnstrVal[ 3][1];
  var colRspCntctLastName = cfgRegFormCnstrVal[ 4][1];
  var colRspCntctLanguage = cfgRegFormCnstrVal[ 5][1];
  var colRspCntctPhone =    cfgRegFormCnstrVal[ 6][1];
  var colRspTeamName =      cfgRegFormCnstrVal[ 7][1];
  var colRspNameP1 =        cfgRegFormCnstrVal[ 8][1];
  var colRspNameP2 =        cfgRegFormCnstrVal[ 9][1];
  var colRspNameP3 =        cfgRegFormCnstrVal[10][1];
  var colRspNameP4 =        cfgRegFormCnstrVal[11][1];
  var colRspNameP5 =        cfgRegFormCnstrVal[12][1];
  var colRspNameP6 =        cfgRegFormCnstrVal[13][1];
  var colRspNameP7 =        cfgRegFormCnstrVal[14][1];
  var colRspNameP8 =        cfgRegFormCnstrVal[15][1];
  
  // Team Table Columns
  var colTblCntctEmail =    cfgRegFormCnstrVal[ 1][2];
  var colTblCntctFullName = cfgRegFormCnstrVal[ 2][2];
  var colTblCntctFrstName = cfgRegFormCnstrVal[ 3][2];
  var colTblCntctLastName = cfgRegFormCnstrVal[ 4][2];
  var colTblCntctLanguage = cfgRegFormCnstrVal[ 5][2];
  var colTblCntctPhone =    cfgRegFormCnstrVal[ 6][2];
  var colTblTeamName =      cfgRegFormCnstrVal[ 7][2];
  var colTblNameP1 =        cfgRegFormCnstrVal[ 8][2];
  var colTblNameP2 =        cfgRegFormCnstrVal[ 9][2];
  var colTblNameP3 =        cfgRegFormCnstrVal[10][2];
  var colTblNameP4 =        cfgRegFormCnstrVal[11][2];
  var colTblNameP5 =        cfgRegFormCnstrVal[12][2];
  var colTblNameP6 =        cfgRegFormCnstrVal[13][2];
  var colTblNameP7 =        cfgRegFormCnstrVal[14][2];
  var colTblNameP8 =        cfgRegFormCnstrVal[15][2];
  
  var colTblStatus =       cfgRegFormCnstrVal[16][2];
  var colTblMemberFileID = cfgRegFormCnstrVal[17][2];
  var colTblContact =      cfgRegFormCnstrVal[18][2];
  var colTblContactGrp =   cfgRegFormCnstrVal[19][2];
  
  // Routine Variables
  var TeamCntctEmail = "";
  var TeamCntctFullName = "";
  var TeamCntctFrstName = "";
  var TeamCntctLastName = "";
  var TeamCntctLanguage = "";
  var TeamCntctPhone = "";
  var TeamName =  "";
  var TeamPlyr = new Array(evntNbPlyrTeam);
  
  var colRspTeamPlyr;
  var colTblTeamPlyr = new Array(evntNbPlyrTeam);
  
  var ArmyDefOffset = 2;
  
  var CntctInfo = new Array(4); // [0]= First Name, [1]= Last Name, [2]= Email Address, [3]= Language Preference
  
  // Email
  TeamCntctEmail = RegRspnVal[0][colRspCntctEmail-1];
  
  // Team Contact First and Last Name
  TeamCntctFrstName = RegRspnVal[0][colRspCntctFrstName-1];
  TeamCntctLastName = RegRspnVal[0][colRspCntctLastName-1];
  
  // Create Team Contact Full Name
  TeamCntctFullName = TeamCntctFrstName + " " + TeamCntctLastName;
  
  // Team Contact Language Preference
  TeamCntctLanguage = RegRspnVal[0][colRspCntctLanguage-1];
  
  // Team Contact Phone Number
  if(colRspCntctPhone != "") TeamCntctPhone = RegRspnVal[0][colRspCntctPhone-1];
  
  // Team Name
  TeamName = RegRspnVal[0][colRspTeamName-1];
  
  // Team Player 1-8
  for(var x = 0; x < evntNbPlyrTeam; x++){
    // Copies colTeamPx Value (colTeamP1 starts at [8][1] )
    colRspTeamPlyr = cfgRegFormCnstrVal[x+8][1];
    // If Value in Response Column is not null
    if(colRspTeamPlyr != "") {
      // Gets Team Table Column Value
      colTblTeamPlyr[x] = cfgRegFormCnstrVal[x+8][2];
      TeamPlyr[x] = RegRspnVal[0][colRspTeamPlyr-1];
    }
  }
  
  // Check if Team exists in List
  for(var i = 1; i <= NbTeams; i++){
    if(TeamName == CurrTeams[i][0]){
      Status = "Cannot complete registration for " + TeamName + ", Duplicate Team Found in List";
      Logger.log(Status)
    }
  }
  // If New Team
  // Copy Values to Teams Sheet at the Next Empty Spot (Number of Teams + 3)
  // Copy Values to Teams List for Organizer Access
  if(Status == "New Team"){
    
    // Team Contact Full Name
    shtTeams.getRange(NextTeamRow, colTblCntctFullName).setValue(TeamCntctFullName);
    Logger.log("Team Contact Name: %s",TeamCntctFullName);
    
    // Team Contact Email Address
    shtTeams.getRange(NextTeamRow, colTblCntctEmail).setValue(TeamCntctEmail);
    Logger.log("Team Contact Email Address: %s",TeamCntctEmail);
    
    // Team Contact Language
    shtTeams.getRange(NextTeamRow, colTblCntctLanguage).setValue(TeamCntctLanguage);
    Logger.log("Team Contact Language: %s",TeamCntctLanguage); 
    
    // Team Contact Phone Number
    if(TeamCntctPhone != ""){
      shtTeams.getRange(NextTeamRow, colTblCntctPhone).setValue(TeamCntctPhone);
      Logger.log("Team Contact Phone: %s",TeamCntctPhone); 
    }
    Logger.log("-----------------------------");
	
    // Team Name
    shtTeams.getRange(NextTeamRow, colTblTeamName).setValue(TeamName);
    Logger.log("Team Name: %s",TeamName); 

    // Team Player 1-8
    for(x = 0; x < evntNbPlyrTeam; x++){
      if(TeamPlyr[x] != "") {
        shtTeams.getRange(NextTeamRow, colTblTeamPlyr[x]).setValue(TeamPlyr[x]);
        Logger.log("Team Player %s: %s",x+1,TeamPlyr[x]);  
      }
    }
    Logger.log("-----------------------------");
    
    // Set Team Contact Info 
    CntctInfo[0]= TeamCntctFrstName;
    CntctInfo[1]= TeamCntctLastName;
    CntctInfo[2]= TeamCntctEmail;
    CntctInfo[3]= TeamCntctLanguage;
    
    // Add Team Info to Contact List and Contact Group
    var CntctStatus = subCrtContact(CntctInfo);
    if(CntctStatus == "Contact Created" || CntctStatus == "Contact Updated") {
      // Set Contact Created in Teams Sheet
      if(CntctStatus == "Contact Created") shtTeams.getRange(NextTeamRow, colTblContact).setValue("Created");
      if(CntctStatus == "Contact Updated") shtTeams.getRange(NextTeamRow, colTblContact).setValue("Updated");
      // Add to Contact Group   
      var CntctGrpStatus = subAddToContactGroup(shtConfig, CntctInfo);
      if(CntctGrpStatus == "Contact added to Contact Group") {
        // Set Added in Contact Group in Teams Sheet
        shtTeams.getRange(NextTeamRow, colTblContactGrp).setValue("Added");
        Logger.log("Contact Added to Contact Group");
      }
      else Logger.log("Contact NOT Added to Contact Group");
    }
    else Logger.log("Contact NOT Created");
 
  }
  
  // Update Team Data
  Team[ 0] = ""; // Team ID
  Team[ 1] = ""; // Team Record File ID
  Team[ 2] = ""; // Team Record File Link
  Team[ 3] = TeamName; // Team Name
  Team[ 4] = TeamPlyr[0]; // Team Member 1
  Team[ 5] = TeamPlyr[1]; // Team Member 2
  if(evntNbPlyrTeam >= 3) Team[ 6] = TeamPlyr[2]; // Team Member 3
  else Team[ 6] = "";
  if(evntNbPlyrTeam >= 4) Team[ 7] = TeamPlyr[3]; // Team Member 4
  else Team[ 7] = "";
  if(evntNbPlyrTeam >= 5) Team[ 8] = TeamPlyr[4]; // Team Member 5
  else Team[ 8] = "";
  if(evntNbPlyrTeam >= 6) Team[ 9] = TeamPlyr[5]; // Team Member 6
  else Team[ 9] = "";
  if(evntNbPlyrTeam >= 7) Team[10] = TeamPlyr[6]; // Team Member 7
  else Team[10] = "";
  if(evntNbPlyrTeam >= 8) Team[11] = TeamPlyr[7]; // Team Member 8
  else Team[11] = "";
  Team[12] = TeamCntctFrstName; // Team Contact First Name
  Team[13] = TeamCntctLastName; // Team Contact Last Name 
  Team[14] = TeamCntctEmail; // Team Contact Email 
  Team[15] = TeamCntctLanguage; // Team Contact Language 
  Team[16] = TeamCntctPhone; // Team Contact Phone Number 
  Team[17] = "" ;// Team Spare
  Team[18] = "" ;// Team Spare
  Team[19] = "" ;// Team Spare
  Team[20] = "" ;// Team Spare
  Team[21] = "" ;// Team Spare
  Team[22] = "" ;// Team Spare
  Team[23] = "" ;// Team Spare
    
  Logger.log(Team);
  
  return Team;
}


// **********************************************
// function fcnProcessDeckList
//
// This function processes the Deck List Info
//
// **********************************************

function fcnProcessDeckList(shtIDs, shtConfig, shtPlayers, shtResponse, RegRspnVal, Member){
  
  // Get Response Sheet Name
  var RespSheetName = shtResponse.getSheetName();
  
  // Get All Values from Response Sheet
  var shtRespMaxCol = shtResponse.getMaxColumns();
  var RespHeader = shtResponse.getRange(1,1,1,shtRespMaxCol).getValues();
  
  Logger.log(RespHeader);
}

// **********************************************
// function fcnModifyReportFormTCG
//
// This function modifies the Match Report Form
// to add new added players
//
// **********************************************

function fcnModifyReportFormTCG(shtConfig, shtIDs, shtCompetitors, cfgEvntParam, evntEscalation) {

  var MatchFormEN = FormApp.openById(shtIDs[11][0]);
  var MatchFormItemEN = MatchFormEN.getItems();
  var MatchFormFR = FormApp.openById(shtIDs[12][0]);
  var MatchFormItemFR = MatchFormFR.getItems();
  var MatchFormNbItem = MatchFormItemEN.length;
 
  // Function Variables
  var ItemTitle;
  var ItemComptrListEN;
  var ItemComptrListFR;
  
  var CompetitorList = subCrtMatchRepComptrList(shtConfig, shtCompetitors, cfgEvntParam);
  
  // Loops in Match Form to Find Players List
  for(var item = 0; item < MatchFormNbItem; item++){
    ItemTitle = MatchFormItemEN[item].getTitle();
    if(ItemTitle == "Winning Player" || ItemTitle == "Losing Player" || ItemTitle == "Winning Team" || ItemTitle == "Losing Team"){
      
      // Get the List Item from the Match Report Form
      ItemComptrListEN = MatchFormItemEN[item].asListItem();
      ItemComptrListFR = MatchFormItemFR[item].asListItem();
      
      // Set the Player List to the Match Report Forms
      ItemComptrListEN.setChoiceValues(CompetitorList);
      ItemComptrListFR.setChoiceValues(CompetitorList);
    }
  }
  
  if(evntEscalation == "Enabled"){
    
    var EscltBonusFormEN = FormApp.openById(shtIDs[17][0]);
    var EscltBonusFormItemEN = EscltBonusFormEN.getItems();
    var EscltBonusFormFR = FormApp.openById(shtIDs[18][0]);
    var EscltBonusFormItemFR = EscltBonusFormFR.getItems();
    var EscltBonusFormNbItem = EscltBonusFormItemEN.length;
    
    // Loops in Escalation Bonus Form to Find Players List
    for(var item = 0; item < EscltBonusFormNbItem; item++){
      ItemTitle = EscltBonusFormNbItem[item].getTitle();
      if(ItemTitle == "Player"){
        
        // Get the List Item from the Round Booster Report Form
        ItemComptrListEN = EscltBonusFormItemEN[item].asListItem();
        ItemComptrListFR = EscltBonusFormItemFR[item].asListItem();

        // Set the Player List to the Round Booster Report Forms
        ItemComptrListEN.setChoiceValues(CompetitorList);
        ItemComptrListFR.setChoiceValues(CompetitorList);
      }
    }
  }
}