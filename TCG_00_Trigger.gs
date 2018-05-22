// **********************************************
// function onSubmitTCG_()
//
// This function analyzes the form submitted
// and executes the appropriate functions
//
// **********************************************

function onSubmitTCG_DemoMtG(e) {
      
  // Get Row from New Response
  var RowResponse = e.range.getRow();
    
  // Get Sheet from New Response
  var shtResponse = SpreadsheetApp.getActiveSheet();
  var ShtName = shtResponse.getSheetName();
  
  Logger.log("------- New Response Received -------");
  Logger.log("Sheet: %s",ShtName);
  Logger.log("Response Row: %s",RowResponse);
  
  // If Form Submitted is a Match Report, process results
  if(ShtName == "MatchResp EN" || ShtName == "MatchResp FR") {
    Logger.log("Match Report Submission Received");
    fcnProcessMatchTCG();
  }
  
  // If Form Submitted is a Player Subscription, Register Player
  if(ShtName == "RegPlyr EN" || ShtName == "RegPlyr FR"){
    Logger.log("Player Registration Received");
    fcnRegistrationPlyrTCG(shtResponse, RowResponse);
  }

  // If Form Submitted is a Team Subscription, Register Team
  if(ShtName == "RegTeam EN" || ShtName == "RegTeam FR"){
    Logger.log("Team Registration Received");
    fcnRegistrationTeamTCG(shtResponse, RowResponse);
  }
  
  // If Form Submitted is a Round Bonus Unit, Enter Bonus Unit
  if(ShtName == "EscltBonus EN" || ShtName == "EscltBonus FR"){
    Logger.log("Escalation Bonus Submission Received");
    fcnEscltBonusTCG(shtResponse, RowResponse);
  }
} 


// **********************************************
// function OnOpenTCG_()
//
// Function executed everytime the Spreadsheet is
// opened or refreshed
//
// **********************************************

function onOpenTCG_DemoMtG() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shtConfig = ss.getSheetByName("Config");
  var cfgEvntParam = shtConfig.getRange( 4, 4,48,1).getValues();
  
  // Event Parameters
  var evntFormat =      cfgEvntParam[ 9][0];
  var evntEscalation =  cfgEvntParam[19][0];
  var evntLogArmyDef =  cfgEvntParam[46][0];
  var evntLogArmyList = cfgEvntParam[47][0];
    
  var AnalyzeDataMenu  = [];
  AnalyzeDataMenu.push({name: "Analyze New Match Entry", functionName: "fcnProcessMatchTCG"});
  AnalyzeDataMenu.push({name: "Update Standings", functionName: "fcnUpdateStandings"})
  AnalyzeDataMenu.push({name: "Clear Match Results and Entries", functionName:"fcnClearMatchResults"});
  
  var EventMenu = [];
  EventMenu.push({name:"Refresh Menus", functionName:"onOpenTCG_DemoMtG"});
  EventMenu.push({name:"Initialize Event", functionName:"fcnInitializeEvent"});
  EventMenu.push({name:"Update Config ID & Links", functionName:"fcnUpdateLinksIDs"});
  EventMenu.push({name:"Create Match Report Forms", functionName:"fcnCrtMatchReportForm_TCG"});
  EventMenu.push({name:"Setup Match Response Sheets",functionName:"fcnSetupMatchResponseSht"});
  if(evntFormat == "Single" || evntFormat == "Team+Players") EventMenu.push({name:"Create Player Registration Forms", functionName:"fcnCrtRegstnFormPlyr_TCG"});
  if(evntFormat == "Team" || evntFormat == "Team+Players")   EventMenu.push({name:"Create Team Registration Forms", functionName:"fcnCrtRegstnFormTeam_TCG"});
  if(evntEscalation == "Enabled") EventMenu.push({name:"Create Escalation Bonus Forms", functionName:"fcnCrtEscltForm_TCG"});
  EventMenu.push(null);
  // If Army Lists are used
  if(evntLogArmyDef == "Enabled" || evntLogArmyList == "Enabled"){
    EventMenu.push({name:"Create Players Card DBs", functionName:"fcnCrtPlayerCardDB"});
    EventMenu.push({name:"Create Players Card Lists", functionName:"fcnCrtPlayerCardList"});
  }
  EventMenu.push({name:"Create Players/Teams Records", functionName:"fcnCrtEvntRecord"});
  if(evntEscalation == "Enabled") EventMenu.push({name:"Create Players Escalation Bonus Sheets", functionName:"fcnCrtPlayerEscltBonus"});
  EventMenu.push(null);
  // If Army Lists are used
  if(evntLogArmyDef == "Enabled" || evntLogArmyList == "Enabled"){
    EventMenu.push({name:"Delete Players Card DBs", functionName:"fcnDelPlayerCardDB"});
    EventMenu.push({name:"Delete Players Card Lists", functionName:"fcnDelPlayerCardList"});
  }
  EventMenu.push({name:"Delete Players Records", functionName:"fcnDelEventPlayerRecord"});
  if(evntEscalation == "Enabled") EventMenu.push({name:"Delete Players Escalation Bonus Sheets", functionName:"fcnDelPlayerEscltBonus"});
  
  var TestMenu  = [];
  TestMenu.push({name: "Test Email Log", functionName: "fcnTestEmail"});
  
  ss.addMenu("Manage Event", EventMenu);
  ss.addMenu("Process Data", AnalyzeDataMenu);
  //ss.addMenu("Test Menu", TestMenu);  
}



// **********************************************
// function fcnRoundChangeTCG()
//
// When the Round number changes, this function 
// executes different functions 
//
// **********************************************

function onRoundChangeTCG_DemoMtG(){

  // Main Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Open Configuration Spreadsheet
  var shtConfig = ss.getSheetByName('Config');
  
  // Configuration Parameters
  var shtIDs = shtConfig.getRange(4,7,24,1).getValues();
  var cfgUrl = shtConfig.getRange(4,11,20,1).getValues();
  var cfgEvntParam = shtConfig.getRange(4,4,48,1).getValues();
  var cfgColRspSht = shtConfig.getRange(4,15,16,1).getValues();
  var cfgColRndSht = shtConfig.getRange(4,18,16,1).getValues();
  var cfgExecData  = shtConfig.getRange(4,21,16,1).getValues();

  // Get Log Sheet
  var shtLog = SpreadsheetApp.openById(shtIDs[1][0]).getSheetByName('Log');
  
  // Get Document URLs
  var urlStandingsEN = cfgUrl[5][0];
  var urlStandingsFR = cfgUrl[6][0];
  
  // Facebook Page Link
  var urlFacebook = shtConfig.getRange(4,15).getValue(); 
  
  // League / Tournament Parameters
  var evntNameEN = cfgEvntParam[0][0] + ' ' + cfgEvntParam[7][0];
  var evntNameFR = cfgEvntParam[8][0] + ' ' + cfgEvntParam[0][0];
  var evntMinGame = cfgEvntParam[15][0];
  var evntLocationEmail = cfgEvntParam[1][0];
  
  // Email Variables
  // [0][0]= Recipients, [0][1]= Subject, [0][2]= Message, [0][3]= Language 
  var EmailDataEN = subCreateArray(1,4);
  var EmailDataFR = subCreateArray(1,4);
  EmailDataEN[0][3] = 'English';
  EmailDataFR[0][3] = 'Français';
  
  var GenRecipients;
  var AdminEmail = Session.getActiveUser().getEmail();
  
  // Function Values
  var shtCumul = ss.getSheetByName('Cumulative Results');
  var Round = shtCumul.getRange(2,3).getValue();
  var LastRound = Round - 1;
  var RoundShtName = 'Round'+LastRound;
  var shtRound = ss.getSheetByName(RoundShtName);
  var shtPlayers = ss.getSheetByName('Players');
  var NbPlayers = shtPlayers.getRange(2,1).getValue();
  
  // Function Variables
  var PenaltyTable;
  var RoundData;
  var TotalMatch = 0;
  var TotalWins = 0;
  var TotalLoss = 0;
  var TotalMatchStore = 0;
  var MostParam;
  
  // Array to Find Player with Most Matches Played in Store
  // [0][0]= Type (Wins, Loss, Win%, Store, PunPack)
  // [0][1]= Round Award Description
  // [0][2]= 
  // [0][3]= 
  // [x][0]= Player Name, [x][1]= Value
  
  // Generate the Round Report
  fcnGenerateRoundReport();
  
  // RoundPrize Data
  var RoundPrizeData = shtConfig.getRange(4,36,10,4).getValues();
  
  // [x][1]= Prize Category 1 				[x][2]= Prize Category 2				[x][3]= Prize Category 3					
  // [0][1]= Round Prize                    [0][2]= Round Prize                    [0][3]= Round Prize
  // [1][1]= Type				 			[1][2]= Type 							[1][3]= Type				 													
  
  var PlayerMost1 = subCreateArray(NbPlayers+1,4);
  PlayerMost1[0][0]= RoundPrizeData[2][1];

  // Array to Find Player with Most Losses
  var PlayerMost2 = subCreateArray(NbPlayers+1,4);
  PlayerMost2[0][0]= RoundPrizeData[2][2];

  // Array to Find Player with Most Losses
  var PlayerMost3 = subCreateArray(NbPlayers+1,4);
  PlayerMost3[0][0]= RoundPrizeData[2][3];

  // Modify the Round Number in the Match Report Sheet
  fcnModifyRoundMatchReport(ss, shtConfig);
  
  // Verify Round Matches Data Integrity
  RoundData = shtRound.getRange(5,4,NbPlayers,7).getValues(); //[0]= Matches Played [1]= Wins [2]= Losses [3]= Ties [6]= Matches in Store
  var RoundTotals = shtRound.getRange(2, 4, 1, 6).getValues();
  // Get Total Matches Played
  TotalMatch = RoundTotals[0][0];
  
  // Get Total Wins
  TotalWins = RoundTotals[0][1];
  
  // Get Total Losses
  TotalLoss = RoundTotals[0][2];
  
  // Get Amount of matches played at the store this Round.
  TotalMatchStore = RoundTotals[0][5];
  
  // Create RoundStats Array
  var RoundStats = subCreateArray(2,4); 
    
  RoundStats[0][0] = LastRound;
  RoundStats[0][1] = Round;
  
  RoundStats[1][0] = TotalMatch;
  RoundStats[1][1] = TotalMatchStore;  
  RoundStats[1][2] = TotalWins;
  RoundStats[1][3] = TotalLoss;
  
  
  // If All Totals are equal, Round Data is Valid, Send Round Report
  if(TotalMatch == TotalWins &&  TotalMatch == TotalLoss && TotalWins == TotalLoss) {
    
    // Round Awards
    //Player with Most 1
    PlayerMost1 = subPlayerWithMost(PlayerMost1, NbPlayers, shtRound);
    
    // Player with Most 2
    PlayerMost2 = subPlayerWithMost(PlayerMost2, NbPlayers, shtRound);
  
    // Player with Most 3
    PlayerMost3 = subPlayerWithMost(PlayerMost3, NbPlayers, shtRound);

    // Send Round Report Email
    
    // Email Subject
    EmailDataEN[0][1] = evntNameEN +" - Round Report " + LastRound;
    EmailDataFR[0][1] = evntNameFR +" - Rapport du Round " + LastRound;
    
    // Generate Round Report Messages
    // Email Message
    EmailDataEN = fcnGenRoundReportMsg(ss, shtConfig, EmailDataEN, RoundStats, RoundPrizeData, PlayerMost1, PlayerMost2, PlayerMost3);
    
    EmailDataFR = fcnGenRoundReportMsg(ss, shtConfig, EmailDataFR, RoundStats, RoundPrizeData, PlayerMost1, PlayerMost2, PlayerMost3);
    
    
    
    // If there is a minimum games to play per Round, generate the Penalty Losses for players who have played less than the minimum
    if(evntMinGame > 0){
      
      // Players Array to return Penalty Losses
      var PlayerData = new Array(32); // 0= Player Name, 1= Penalty Losses
      for(var plyr = 0; plyr < 32; plyr++){
        PlayerData[plyr] = new Array(2); 
        for (var val = 0; val < 2; val++) PlayerData[plyr][val] = '';
      }
      
      // Analyze if Players have missing matches to apply Loss Penalties
      PlayerData = fcnAnalyzeLossPenalty(ss, Round, PlayerData);
      
      // Logs All Players Record
      for(var row = 0; row<32; row++){
        if (PlayerData[row][0] != '') Logger.log('Player: %s - Missing: %s',PlayerData[row][0], PlayerData[row][1]);
      }
      
      // Populate the Penalty Table for the Round Report
      PenaltyTable = subEmailPlayerPenaltyTable(PlayerData);  
      // Update the Email message to add the Penalty Losses table
      EmailDataEN[0][2] += PenaltyTable;
      EmailDataFR[0][2] += PenaltyTable;
    }
    
    
    
    
    // English Custom Message
    // Add Final Tournament
    EmailDataEN[0][2] += '<br><br><b><font size="3">Final Tournament<font></b>'+
      '<br>The first 8 players with the best Win Ratio <b>AND at least 10 games played</b> will qualify for the final tournament.';
    // Add Standings Link
    EmailDataEN[0][2] += "<br><br>Click here to access the League Standings and Results:<br>" + urlStandingsEN ;
    // Add Facebook Page Link
    EmailDataEN[0][2] += "<br><br>Please join the Community Facebook page to chat with other players and plan matches.<br><br>" + urlFacebook;
    // Turn1 Signature
    EmailDataEN[0][2] += "<br><br>Thank you for using TCG Booster League Manager from Turn 1 Gaming Leagues & Tournaments";
    
    // French Custom Message
    // Add Final Tournament
    EmailDataFR[0][2] += '<br><br><b><font size="3">Tournoi Final<font></b>'+
      '<br>Les 8 premiers joueurs qui ont le meilleur ratio de victoire <b>ET au moins 10 parties jouées</b> vont se qualifier pour le tournoi final.';
    // Add Standings Link
    EmailDataFR[0][2] += "<br><br>Cliquez ici pour accéder aux résutlats et classement de la ligue:<br>" + urlStandingsFR ;
    // Add Facebook Page Link
    EmailDataFR[0][2] += "<br><br>Joignez vous à la page Facebook de la communauté pour discuter avec les autres joueurs et organiser vos parties.<br><br>" + urlFacebook;
    // Turn1 Signature
    EmailDataFR[0][2] += "<br><br>Merci d'utiliser TCG Booster League Manager de Turn 1 Gaming Ligues & Tournois";
    
    // General Recipients
    GenRecipients = AdminEmail + ', ' + evntLocationEmail;
    //Recipients = Session.getActiveUser().getEmail();
    
    // Get English Players Email
    EmailDataEN[0][0] = subGetEmailRecipients(shtPlayers, EmailDataEN[0][3]);
    
    // Get French Players Email
    EmailDataFR[0][0] = subGetEmailRecipients(shtPlayers, EmailDataFR[0][3]);
    
    // Send English Email
    MailApp.sendEmail(GenRecipients, EmailDataEN[0][1],"",{bcc:EmailDataEN[0][0],name:'Turn 1 Gaming League Manager',htmlBody:EmailDataEN[0][2]});
    
    // Send French Email
    MailApp.sendEmail(GenRecipients, EmailDataFR[0][1],"",{bcc:EmailDataFR[0][0],name:'Turn 1 Gaming League Manager',htmlBody:EmailDataFR[0][2]});
    
    // Execute Ranking function in Standing tab
    fcnUpdateStandings();
    
    // Copy all data to League Spreadsheet
    fcnCopyStandingsSheets(ss, shtConfig, cfgEvntParam, cfgColRndSht, LastRound, 0);
  }
  
  // If Round Match Data is not Valid
  else{
    Logger.log('Round Match Data is not Valid');
    Logger.log('Total Match Played: %s',TotalMatch);
    Logger.log('Total Wins: %s',TotalWins);
    Logger.log('Total Losses: %s',TotalLoss);
  
    // Send Log by email
    var recipient = Session.getActiveUser().getEmail();
    var subject = LeagueNameEN + ' - Round ' + LastRound + ' - Round Data is not Valid';
    var body = Logger.getLog();
    MailApp.sendEmail(recipient, subject, body)
  }
  
  // Post Log to Log Sheet
  subPostLog(shtLog, Logger.getLog());
  
}