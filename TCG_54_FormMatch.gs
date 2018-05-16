// **********************************************
// function fcnCrtMatchReportForm_TCG()
//
// This function creates the Match Report Form 
// based on the parameters in the Config File
//
// **********************************************

function fcnCrtMatchReportForm_TCG() {
  
  Logger.log("Routine: fcnCrtMatchReportForm_TCG");
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shtConfig =  ss.getSheetByName('Config');
  var shtPlayers = ss.getSheetByName('Players');
  var shtTeams =   ss.getSheetByName('Teams');
    
  // Configuration Data
  var shtIDs = shtConfig.getRange(4,7,20,1).getValues();
  var cfgEvntParam = shtConfig.getRange(4,4,48,1).getValues();
  var cfgColRspSht = shtConfig.getRange(4,18,16,1).getValues();
  var cfgColRndSht = shtConfig.getRange(4,21,16,1).getValues();
  var cfgExecData  = shtConfig.getRange(4,24,16,1).getValues();
  var cfgArmyBuild = shtConfig.getRange(4,33,20,1).getValues();
  
  // Registration Form Construction 
  // Column 1 = Category Name
  // Column 2 = Category Order in Form
  // Column 3 = Column Value in Player/Team Sheet
  var cfgReportFormCnstrVal = shtConfig.getRange(4,30,20,2).getValues();
  
  // Execution Parameters
  var exeGnrtResp = cfgExecData[3][0];
  
  // Event Properties
  var evntLocation =       cfgEvntParam[0][0];
  var evntName =           cfgEvntParam[7][0];
  var evntFormat =         cfgEvntParam[9][0];
  var evntTeamNbPlyr =     cfgEvntParam[10][0];
  var evntTeamMatch =      cfgEvntParam[11][0];
  var evntLocationBonus =  cfgEvntParam[23][0];
  var evntMatchPtsMin =    0;
  var evntPtsGainedMatch = cfgEvntParam[27][0];
  var evntMatchPtsMax =    cfgEvntParam[28][0];
  var evntTiePossible =    cfgEvntParam[32][0];
  
  var RoundNum = shtConfig.getRange(7,2).getValue();
  var RoundArray = new Array(1); RoundArray[0] = RoundNum;
  
  var NbPlyr = shtConfig.getRange(13,2).getValue();
  var NbTeam = shtConfig.getRange(14,2).getValue();
  
  // Log Sheet
  var shtLog = SpreadsheetApp.openById(shtIDs[1][0]).getSheetByName('Log');

  // Registration ID from the Config File
  var ssID =     shtIDs[0][0];
  var FormIdEN = shtIDs[7][0];
  var FormIdFR = shtIDs[8][0];
 
  // Row Column Values to Write Form IDs and URLs
  var rowFormEN  = 11;
  var rowFormFR  = 12;
  var colFormID  = 7;
  var colFormURL = 11
  
  var ssTexts = SpreadsheetApp.openById('1DkSr5HbGqZ_c38DlHKiBhgcBXw3fr3CK9zDE04187fE');
  var shtTxtReport = ssTexts.getSheetByName('Match Report TCG');
  var ConfirmMsgEN = shtTxtReport.getRange(4,2).getValue();
  var ConfirmMsgFR = shtTxtReport.getRange(4,3).getValue();
  
  var QuestionOrder = 2;
    
  var ssSheets;
  var NbSheets;
  var SheetName;
  var shtResp;
  var shtRespMaxRow;
  var shtRespMaxCol;
  var FirstCellVal;
  
  var formEN;
  var FormIdEN;
  var FormNameEN;
  var FormItemsEN;
  
  var formFR;
  var FormIdFR;
  var FormNameFR;
  var FormItemsFR;
  
  var Players;
  var PlayerList;
  var Player1List;
  var Player2List;
  var Teams;
  var TeamList;
  var Team1List;
  var Team2List;
  var TeamListLength;
  
  var shtResp1;
  var shtResp2;
  var shtRespName1;
  var shtRespName2;
  var IndexResponses = ss.getSheetByName("Responses").getIndex();
  var FormsCreated = 0;
  var FormsDeleted = 0;
    
  // Insert ui to confirm
  var ui = SpreadsheetApp.getUi();
  var title;
  var msg;
  var uiResponse;

  // If Form Exists, Log Error Message
  if(FormIdEN != '' || FormIdFR != ''){
    title = "Match Report Forms Overwrite";
    msg = "The Match Report Forms already exist. Click OK to overwrite.";
    var uiResponse = ui.alert(title, msg, ui.ButtonSet.OK_CANCEL);
    
    if(uiResponse == "OK"){
      // Clear IDs and URLs
      shtConfig.getRange(rowFormEN, colFormID).clearContent();
      shtConfig.getRange(rowFormFR, colFormID).clearContent();
      shtConfig.getRange(rowFormEN, colFormURL).clearContent(); 
      shtConfig.getRange(rowFormFR, colFormURL).clearContent();
      
      // If Responses Sheets exist, Unlink and Delete them
      shtResp1 = ss.getSheets()[IndexResponses];
      shtRespName1 = shtResp1.getName();
      shtResp2 = ss.getSheets()[IndexResponses+1];
      shtRespName2 = shtResp2.getName();
      
      // First Sheet After Responses is MatchResp EN
      if(shtRespName1 == "MatchResp EN"){
        FormApp.openById(FormIdEN).removeDestination();
        ss.deleteSheet(shtResp1);
      }
      
      // Second Sheet After Responses is MatchResp EN
      if(shtRespName2 == "MatchResp EN"){
        FormApp.openById(FormIdEN).removeDestination();
        ss.deleteSheet(shtResp2);
      }
      
      // First Sheet After Responses is MatchResp EN
      if(shtRespName1 == "MatchResp FR"){
        FormApp.openById(FormIdFR).removeDestination();
        ss.deleteSheet(shtResp1);
      }
      
      // Second Sheet After Responses is MatchResp FR
      if(shtRespName2 == "MatchResp FR"){
        FormApp.openById(FormIdFR).removeDestination();
        ss.deleteSheet(shtResp2);
      }
      // Forms Deleted Flag
      FormsDeleted = 1;
    }
  }

  // CREATE VALIDATIONS
  // Points Validation
  var PointsValidationEN = FormApp.createTextValidation()
  .setHelpText("Enter a number between " + evntMatchPtsMin + " and " + evntMatchPtsMax)
  .requireNumberBetween(evntMatchPtsMin, evntMatchPtsMax)
  .build();
  
  var PointsValidationFR = FormApp.createTextValidation()
  .setHelpText("Entrez un nombre entre " + evntMatchPtsMin + " et " + evntMatchPtsMax)
  .requireNumberBetween(evntMatchPtsMin, evntMatchPtsMax)
  .build();
  
  // Create Forms
  if ((FormIdEN == "" && FormIdFR == "") || FormsDeleted == 1){
    
    //---------------------------------------------
    // TITLE SECTION
    // English
    FormNameEN = evntLocation + " " + evntName + " Match Reporter EN";
    formEN = FormApp.create(FormNameEN).setTitle(FormNameEN)
    .setDescription("Please enter the following information to submit your match result");
    // French    
    FormNameFR = evntLocation + " " + evntName + " Match Reporter FR";
    formFR = FormApp.create(FormNameFR).setTitle(FormNameFR)
    .setDescription("SVP, entrez les informations suivantes pour soumettre votre rapport de match");
    
    // Create Player List for Match Report
    if(NbPlyr > 0) PlayerList = subCrtMatchRepPlyrList(shtConfig, shtPlayers, cfgEvntParam);
    
    // Create Team List for Match Report
    if(NbTeam > 0) TeamList = subCrtMatchRepTeamList(shtConfig, shtTeams, cfgEvntParam);

         
    // Loops in Response Columns Values and Create Appropriate Question
    for(var i = 1; i < cfgReportFormCnstrVal.length; i++){
      // Look for Col Equal to Question Order
      if(QuestionOrder == cfgReportFormCnstrVal[i][1]){
        Logger.log("Switch");
        Logger.log("Qstn:%s - Value:%s",QuestionOrder,cfgReportFormCnstrVal[i][1]);
        Logger.log(cfgReportFormCnstrVal[i][0]);
        switch(cfgReportFormCnstrVal[i][0]){
            
            //---------------------------------------------
            // PASSWORD SECTION
          case 'Password':{ 
            // English
            formEN.addTextItem()
            .setTitle("Event Password")
            .setHelpText("Please enter the Event Password to send your match report")
            .setRequired(true);
            
            // French
            formFR.addTextItem()
            .setTitle("Mot de passe de l'événement")
            .setHelpText("SVP, entrez le mot de passe de l'événement pour envoyer votre rapport de match")
            .setRequired(true);
            
            break;
          }
            
            //---------------------------------------------
            // ROUND NUMBER
          case 'Round Number':{ 
            // English
            if(evntFormat == 'Single') formEN.addPageBreakItem().setTitle("Round Number & Players");
            if(evntFormat == 'Team')   formEN.addPageBreakItem().setTitle("Round Number & Teams");
            // Round
            formEN.addListItem()
            .setTitle("Round")
            .setRequired(true)
            .setChoiceValues(RoundArray);
                
            // French
            if(evntFormat == 'Single') formFR.addPageBreakItem().setTitle("Numéro de Semaine & Joueurs");
            if(evntFormat == 'Team')   formFR.addPageBreakItem().setTitle("Numéro de Semaine & Équipes");
            
            // Semaine
            formFR.addListItem()
            .setTitle("Ronde")
            .setRequired(true)
            .setChoiceValues(RoundArray);
            
            break;
          }
            
            //---------------------------------------------
            // PLAYERS
            // Player 1 List
          case 'Player 1':{ 
            // If Points Gained in Match are used
            if(evntPtsGainedMatch == "Enabled"){
              // English
              Player1List = formEN.addListItem()
              .setTitle("Player 1")
              .setHelpText("Select your name")
              .setRequired(true);
              if (NbPlyr > 0) Player1List.setChoiceValues(PlayerList);
              
              // French
              Player1List = formFR.addListItem()
              .setTitle("Joueur 1")
              .setHelpText("Sélectionnez votre nom")
              .setRequired(true);
              if (NbPlyr > 0) Player1List.setChoiceValues(PlayerList);
            }
            // If Points Gained in Match are not used
            if(evntPtsGainedMatch == "Disabled"){
              // English
              Player1List = formEN.addListItem()
              .setTitle("Winning Player")
              .setHelpText("If Game is a Tie, select your name")
              .setRequired(true);
              if (NbPlyr > 0) Player1List.setChoiceValues(PlayerList);
              
              // French
              Player1List = formFR.addListItem()
              .setTitle("Joueur Gagnant")
              .setHelpText("Si la partie est nulle, sélectionnez votre nom")
              .setRequired(true);
              if (NbPlyr > 0) Player1List.setChoiceValues(PlayerList);
            }
            break;
          }
            // Player 2 List
          case 'Player 2':{ 
            // If Points Gained in Match are used
            if(evntPtsGainedMatch == "Enabled"){
              // English
              Player2List = formEN.addListItem()
              .setTitle("Player 2")
              .setHelpText("Select your opponent")
              .setRequired(true);
              if (NbPlyr > 0) Player2List.setChoiceValues(PlayerList); 
              
              // French
              Player2List = formFR.addListItem()
              .setTitle("Joueur 2")
              .setHelpText("Sélectionnez votre adversaire")
              .setRequired(true);
              if (NbPlyr > 0) Player2List.setChoiceValues(PlayerList);
            }
            // If Points Gained in Match are not used
            if(evntPtsGainedMatch == "Disabled"){
              // English
              Player2List = formEN.addListItem()
              .setTitle("Losing Player")
              .setHelpText("If Game is a Tie, select your opponent")
              .setRequired(true);
              if (NbPlyr > 0) Player2List.setChoiceValues(PlayerList); 
              
              // French
              Player2List = formFR.addListItem()
              .setTitle("Joueur Perdant")
              .setHelpText("Si la partie est nulle, sélectionnez votre adversaire")
              .setRequired(true);
              if (NbPlyr > 0) Player2List.setChoiceValues(PlayerList);
            }
            break;
          }
            
            //---------------------------------------------
            // TEAMS
            // Team 1 List
          case 'Team 1':{ 
            // If Points Gained in Match are used
            if(evntPtsGainedMatch == "Enabled"){
              // English
              Team1List = formEN.addListItem()
              .setTitle("Team 1")
              .setHelpText("Select your team")
              .setRequired(true);
              if (NbTeam > 0) Team1List.setChoiceValues(TeamList);
              
              // French
              Team1List = formFR.addListItem()
              .setTitle("Équipe 1")
              .setHelpText("Sélectionnez votre équipe")
              .setRequired(true);
              if (NbTeam > 0) Team1List.setChoiceValues(TeamList);
            }
            // If Points Gained in Match are not used
            if(evntPtsGainedMatch == "Disabled"){
              // English
              Team1List = formEN.addListItem()
              .setTitle("Winning Team")
              .setHelpText("If Game is a Tie, select your team")
              .setRequired(true);
              if (NbTeam > 0) Team1List.setChoiceValues(TeamList);
              
              // French
              Team1List = formFR.addListItem()
              .setTitle("Équipe Gagnante")
              .setHelpText("Si la partie est nulle, sélectionnez votre équipe")
              .setRequired(true);
              if (NbTeam > 0) Team1List.setChoiceValues(TeamList);
            }
            break;
          }
            // Team 2 List
          case 'Team 2':{ 
            // If Points Gained in Match are used
            if(evntPtsGainedMatch == "Enabled"){ 
              // English
              Team2List = formEN.addListItem()
              .setTitle("Team 2")
              .setHelpText("Select the opposing team")
              .setRequired(true);
              if (NbTeam > 0) Team2List.setChoiceValues(TeamList); 
              
              // French
              Team2List = formFR.addListItem()
              .setTitle("Équipe 2")
              .setHelpText("Sélectionnez l'équipe adverse")
              .setRequired(true);
              if (NbTeam > 0) Team2List.setChoiceValues(TeamList);
            } 
            // If Points Gained in Match are not used
            if(evntPtsGainedMatch == "Disabled"){ 
              // English
              Team2List = formEN.addListItem()
              .setTitle("Losing Team")
              .setHelpText("If Game is a Tie, select the opposing team")
              .setRequired(true);
              if (NbTeam > 0) Team2List.setChoiceValues(TeamList); 
              
              // French
              Team2List = formFR.addListItem()
              .setTitle("Équipe Perdante")
              .setHelpText("Si la partie est nulle, sélectionnez l'équipe adverse")
              .setRequired(true);
              if (NbTeam > 0) Team2List.setChoiceValues(TeamList);
            } 
            break;
          }
            //---------------------------------------------
            // WINNING POINTS
          case 'P/T Points 1':{ 
            if(evntPtsGainedMatch == 'Enabled'){
              // English
              formEN.addTextItem()
              .setTitle("Points Scored")
              .setHelpText("Enter the points scored by Player 1 or Team 1")
              .setValidation(PointsValidationEN)
              .setRequired(true);
              
              // French
              formFR.addTextItem()
              .setTitle("Points Marqués")
              .setHelpText("Entrez les points accumulés par le joueur 1 ou l'équipe 1")
              .setValidation(PointsValidationFR)
              .setRequired(true);
            }
            break;
          }
            
            //---------------------------------------------
            // LOSING POINTS
          case 'P/T Points 2':{ 
            if(evntPtsGainedMatch == 'Enabled'){
              // English
              formEN.addTextItem()
              .setTitle("Points Scored")
              .setHelpText("Enter the points scored by Player 2 or Team 2")
              .setValidation(PointsValidationEN)
              .setRequired(true);
              
              // French
              formFR.addTextItem()
              .setTitle("Points Marqués")
              .setHelpText("Entrez les points accumulés par le joueur 2 ou l'équipe 2")
              .setValidation(PointsValidationFR)
              .setRequired(true);
            }
            break;
          }
            //---------------------------------------------
            // GAME TIE
          case 'Game is Tie':{
            if(evntTiePossible == "Enabled"){
              // English
              formEN.addMultipleChoiceItem()
              .setTitle("Game is a Tie?")
              .setHelpText("OPTIONAL")
              .setChoiceValues(["No","Yes"]);
              
              // French
              formFR.addMultipleChoiceItem()
              .setTitle("Partie est Nulle?")
              .setHelpText("OPTIONNEL")
              .setChoiceValues(["Non","Oui"]);
            }
            break;
          }

            //---------------------------------------------
            // LOCATION SECTION
          case 'Location':{ 
            // English
            formEN.addPageBreakItem().setTitle("Location")
            formEN.addMultipleChoiceItem()
            .setTitle("Location Bonus")
            .setHelpText("Did you play at the store?")
            .setRequired(true)
            .setChoiceValues(["Yes","No"]);
            
            // French
            formFR.addPageBreakItem().setTitle("Localisation")
            formFR.addMultipleChoiceItem()
            .setTitle("Bonus de Localisation")
            .setHelpText("Avez-vous joué au magasin?")
            .setRequired(true)
            .setChoiceValues(["Oui","Non"]);
            break;
          }
          default : break;
        }
        // Increment to Next Question
        QuestionOrder++;
        // Reset Loop if new question was added
        i = -1;
      }
      
      //---------------------------------------------
      // CONFIRMATION MESSAGE
      
      // English
      formEN.setConfirmationMessage(ConfirmMsgEN);
      
      // French
      formFR.setConfirmationMessage(ConfirmMsgFR);
      
      // Forms Created Flag
      FormsCreated = 1;
    
    }

    // RESPONSE SHEETS
    // Create Response Sheet in Main File and Rename
    if(exeGnrtResp == "Enabled" && FormsCreated == 1){
      Logger.log("Generating Response Sheets and Form Links");
      IndexResponses = ss.getSheetByName("Responses").getIndex();
      
      // English Form
      formEN.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
      
      // Find and Rename Response Sheet
      ss = SpreadsheetApp.openById(ssID);
      ssSheets = ss.getSheets();
      ssSheets[0].setName('MatchResp EN');
      
      // Move Response Sheet to appropriate spot in file
      shtResp = ss.getSheetByName('MatchResp EN');
      ss.moveActiveSheet(IndexResponses+1);
      shtRespMaxRow = shtResp.getMaxRows();
      shtRespMaxCol = shtResp.getMaxColumns();
      
      // Delete All Empty Rows
      shtResp.deleteRows(3, shtRespMaxRow - 2);
      
      // Delete All Empty Columns
      for(var c = 1;  c <= shtRespMaxCol; c++){
        FirstCellVal = shtResp.getRange(1, c).getValue();
        if(FirstCellVal == '') {
          shtResp.deleteColumns(c,shtRespMaxCol-c+1);
          c = shtRespMaxCol + 1;
        }
      }
      
      // French Form
      formFR.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
      
      // Find and Rename Response Sheet
      ss = SpreadsheetApp.openById(ssID);
      ssSheets = ss.getSheets();
      ssSheets[0].setName('MatchResp FR');
      
      // Move Response Sheet to appropriate spot in file
      shtResp = ss.getSheetByName('MatchResp FR');
      ss.moveActiveSheet(IndexResponses+2);
      shtRespMaxRow = shtResp.getMaxRows();
      shtRespMaxCol = shtResp.getMaxColumns();
      
      // Delete All Empty Rows
      shtResp.deleteRows(3, shtRespMaxRow - 2);
      
      // Delete All Empty Columns
      for(var c = 1;  c <= shtRespMaxCol; c++){
        FirstCellVal = shtResp.getRange(1, c).getValue();
        if(FirstCellVal == '') {
          shtResp.deleteColumns(c,shtRespMaxCol-c+1);
          c = shtRespMaxCol + 1;
        }
      }
      
      // Set Match Report IDs in Config File
      FormIdEN = formEN.getId();
      shtConfig.getRange(rowFormEN, colFormID).setValue(FormIdEN);
      FormIdFR = formFR.getId();
      shtConfig.getRange(rowFormFR, colFormID).setValue(FormIdFR);
      
      // Create Links to add to Config File  
      shtConfig.getRange(rowFormEN, colFormURL).setValue(formEN.getPublishedUrl()); 
      shtConfig.getRange(rowFormFR, colFormURL).setValue(formFR.getPublishedUrl());
      
      Logger.log("Response Sheets and Form Links Generated");
      
    }
  }

  // Post Log to Log Sheet
  subPostLog(shtLog,Logger.getLog());
  
}

// **********************************************
// function fcnSetupResponseSht()
//
// This function sets up the new Responses sheets 
// and deletes the old ones
//
// **********************************************

function fcnSetupMatchResponseSht(){
  
  Logger.log("Routine: fcnSetupResponseSht");

  // Main Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // Configuration Sheet
  var shtConfig = ss.getSheetByName('Config');
  var cfgColRspSht = shtConfig.getRange(4,18,16,1).getValues();
  
  // Open Responses Sheets
  var shtNewRespEN = ss.getSheetByName('MatchResp EN');
  var shtNewRespFR = ss.getSheetByName('MatchResp FR');
  
  var ColWidth;
  
  // Columns Values and Parameters
  var RspnDataInputs =      cfgColRspSht[0][0]; // from Time Stamp to Data Processed
  var colMatchID =          cfgColRspSht[1][0];
  var colPrcsd =            cfgColRspSht[2][0];
  var colDataConflict =     cfgColRspSht[3][0];
  var colStatus =           cfgColRspSht[4][0];
  var colStatusMsg =        cfgColRspSht[5][0];
  var colMatchIDLastVal =   cfgColRspSht[6][0];
  var colNextEmptyRow =     cfgColRspSht[7][0];
  var colNbUnprcsdEntries = cfgColRspSht[8][0];
  
  var LastCol = colNbUnprcsdEntries;
  var value;
  
  // Copy Header from Old to New sheet - Loop to Copy Value and Format from cell to cell, copy formula (or set) in last cell
  for (var col = 1; col <= LastCol; col++){
    // Insert Column if it doesn't exist
    if (col >= colMatchID && col <= LastCol){
      // Insert New Column
      shtNewRespEN.insertColumnAfter(col-1);
      shtNewRespFR.insertColumnAfter(col-1);
    
      // Set New Response Sheet Values 
      switch(col){
        case colMatchID :{
          // Set Value
          value = '=CONCATENATE("Match ID",CHAR(10),"(data copied to Match Results)")';
          shtNewRespEN.getRange(1, col)
          .setValue(value)
          .setHorizontalAlignment("center")
          .setWrap(true);
          shtNewRespFR.getRange(1, col)
          .setValue(value)
          .setHorizontalAlignment("center")
          .setWrap(true); 
          // Set Width
          shtNewRespEN.setColumnWidth(col, 100);
          shtNewRespFR.setColumnWidth(col, 100);
          break;
        }
		case colPrcsd :{
          // Set Value
          value = '=CONCATENATE("Data",CHAR(10),"Processed",CHAR(10),"Status")';
          shtNewRespEN.getRange(1, col)
          .setValue(value)
          .setHorizontalAlignment("center")
          .setWrap(true); 
          shtNewRespFR.getRange(1, col)
          .setValue(value)
          .setHorizontalAlignment("center")
          .setWrap(true); 
          // Set Width
          shtNewRespEN.setColumnWidth(col, 100);
          shtNewRespFR.setColumnWidth(col, 100);
          break;
        }
        case colDataConflict :{
          // Set Value
          value = '=CONCATENATE("Data",CHAR(10),"Conflict")';
          shtNewRespEN.getRange(1, col)
          .setValue(value)
          .setHorizontalAlignment("center")
          .setNote("Data Conflict will be validated when both players have sent their form. It will compare every field to make sure they are equal. If not, the Data Conflict column will get the value of the data number mismatching")
          .setWrap(true); 
          shtNewRespFR.getRange(1, col)
          .setValue(value)
          .setHorizontalAlignment("center")
          .setNote("Data Conflict will be validated when both players have sent their form. It will compare every field to make sure they are equal. If not, the Data Conflict column will get the value of the data number mismatching")
          .setWrap(true); 
          // Set Width
          shtNewRespEN.setColumnWidth(col, 80);
          shtNewRespFR.setColumnWidth(col, 80);
          break;
        }
        case colStatus :{
          // Set Value
          value = '=CONCATENATE("Process",CHAR(10),"Status")';
          shtNewRespEN.getRange(1, col)
          .setValue(value)
          .setHorizontalAlignment("center")
          .setWrap(true); 
          shtNewRespFR.getRange(1, col)
          .setValue(value)
          .setHorizontalAlignment("center")
          .setWrap(true); 
          // Set Width
          shtNewRespEN.setColumnWidth(col, 80);
          shtNewRespFR.setColumnWidth(col, 80);
          break;
        }
        case colStatusMsg :{
          // Set Value
          value = 'Status Message';
          shtNewRespEN.getRange(1, col)
          .setValue(value)
          .setHorizontalAlignment("center")
          .setWrap(true); 
          shtNewRespFR.getRange(1, col)
          .setValue(value)
          .setHorizontalAlignment("center")
          .setWrap(true); 
          // Set Width
          shtNewRespEN.setColumnWidth(col, 250);
          shtNewRespFR.setColumnWidth(col, 250);
          break;
        }
        case colMatchIDLastVal :{
          // Set Value
          value = 0;
          shtNewRespEN.getRange(1, col)
          .setValue(value)
          .setHorizontalAlignment("center")
          .setNote("Last Match ID Generated"); 
          shtNewRespFR.getRange(1, col)
          .setValue(value)
          .setHorizontalAlignment("center")
          .setNote("Last Match ID Generated"); 
          // Set Width
          shtNewRespEN.setColumnWidth(col, 40);
          shtNewRespFR.setColumnWidth(col, 40);
          break;
        }
        case colNextEmptyRow :{
          // Set Value
          value = '=SUM(indirect("R[1]C[0]",FALSE):indirect("R[301]C[0]",FALSE))+2';
          shtNewRespEN.getRange(1, col)
          .setValue(value)
          .setHorizontalAlignment("center")
          .setNote("Next Empty Row"); 
          shtNewRespFR.getRange(1, col)
          .setValue(value)
          .setHorizontalAlignment("center")
          .setNote("Next Empty Row"); 
          // Set Width
          shtNewRespEN.setColumnWidth(col, 40);
          shtNewRespFR.setColumnWidth(col, 40);
          break;
        }
        case colNbUnprcsdEntries :{
          // Set Value
          value = '=SUM(indirect("R[1]C[0]",FALSE):indirect("R[301]C[0]",FALSE))';
          shtNewRespEN.getRange(1, col)
          .setValue(value)
          .setHorizontalAlignment("center")
          .setNote("Number of Unprocessed Entries"); 
          shtNewRespFR.getRange(1, col)
          .setValue(value)
          .setHorizontalAlignment("center")
          .setNote("Number of Unprocessed Entries"); 
          // Set Width
          shtNewRespEN.setColumnWidth(col, 40);
          shtNewRespFR.setColumnWidth(col, 40);
          break;
        }
      }
    }
  }
  
  // Duplicate New Response EN and rename
  var shtResponses = ss.getSheetByName("Responses")
  var IndexResponses = shtResponses.getIndex();
  ss.deleteSheet(shtResponses);
  shtNewRespEN.activate();
  ss.duplicateActiveSheet();
  ss.getSheetByName("Copy of MatchResp EN").setName("Responses").activate();
  ss.moveActiveSheet(IndexResponses);
  
  // Hides Columns 
  shtNewRespEN.hideColumns(colMatchID);
  shtNewRespEN.hideColumns(colDataConflict);
  shtNewRespEN.hideColumns(colStatus);
  shtNewRespEN.hideColumns(colStatusMsg);
  shtNewRespEN.hideColumns(colMatchIDLastVal);
  
  shtNewRespFR.hideColumns(colMatchID);
  shtNewRespFR.hideColumns(colDataConflict);
  shtNewRespFR.hideColumns(colStatus);
  shtNewRespFR.hideColumns(colStatusMsg);
  shtNewRespFR.hideColumns(colMatchIDLastVal);
  
  Logger.log("Match Response Sheet Setup Complete");
  

}