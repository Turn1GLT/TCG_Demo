// **********************************************
// function fcnCrtRegstnFormPlyr_TCG()
//
// This function creates the Player Registration Form 
// based on the parameters in the Config File
//
// **********************************************

function fcnCrtRegstnFormPlyr_TCG() {
  
  Logger.log("Routine: fcnCreateRegForm_TCG");
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shtConfig = ss.getSheetByName('Config');
  var shtPlayers = ss.getSheetByName('Players');
    
  // Configuration Data
  var shtIDs = shtConfig.getRange(4,7,24,1).getValues();
  var cfgEvntParam =    shtConfig.getRange( 4, 4,48,1).getValues();
  var cfgColRspSht =    shtConfig.getRange( 4,15,16,1).getValues();
  var cfgColRndSht =    shtConfig.getRange( 4,18,16,1).getValues();
  var cfgExecData  =    shtConfig.getRange( 4,21,16,1).getValues();
  var cfgColMatchRep =  shtConfig.getRange( 4,28,20,1).getValues();
  var cfgColMatchRslt = shtConfig.getRange(21,15,32,1).getValues();
  var cfgArmyBuild =    shtConfig.getRange( 4,30,20,1).getValues();
  
  // Registration Form Construction 
  // Column 1 = Category Name
  // Column 2 = Category Order in Form
  // Column 3 = Column Value in Player/Team Sheet
  var cfgRegFormCnstrVal = shtConfig.getRange(4,23,20,3).getValues();
  var cfgArmyBuild =       shtConfig.getRange(4,30,16,1).getValues();
  
  // Execution Parameters
  var exeGnrtResp = cfgExecData[3][0];
  
  // Event Properties
  var evntLocation   = cfgEvntParam[0][0];
  var evntGameSystem = cfgEvntParam[5][0];
  var evntName       = cfgEvntParam[7][0];
  var evntFormat     = cfgEvntParam[9][0];
  var evntNbPlyrTeam = cfgEvntParam[10][0];
  
  var armyBuildRatingVal = cfgArmyBuild[0][0];
  var armyBuildStartVal  = cfgArmyBuild[1][0];
    
  // Log Sheet
  var shtLog = SpreadsheetApp.openById(shtIDs[1][0]).getSheetByName('Log');
  
  // Registration ID from the Config File
  var ssID = shtIDs[0][0];
  var FormIdEN = shtIDs[13][0];
  var FormIdFR = shtIDs[14][0];
 
  // Row Column Values to Write Form IDs and URLs
  var rowFormEN  = 17;
  var rowFormFR  = 18;
  var colFormID  = 7;
  var colFormURL = 8;
  
  var ErrorVal = '';
  var QuestionOrder = 2;
  
  // Army Building Options
  if(evntGameSystem == "Magic the Gathering"){
    var shtConfigWH40k = ss.getSheetByName('ConfigMtG');
    
    var StdCardMax = 4;
    var EDHCardMax = 1;
    
  }
  
  // Routine Variables
  var ssSheets;
  var NbSheets;
  var SheetName;
  var shtResp;
  var shtRespMaxRow;
  var shtRespMaxCol;
  var FirstCellVal;
    
  var formEN;
  var FormNameEN;
  var FormItemsEN;
  var urlFormEN;
  
  var formFR;
  var FormNameFR;
  var FormItemsFR;
  var urlFormFR;

  var TestCol = 1;
  
  var shtResp1;
  var shtResp2;
  var shtRespName1;
  var shtRespName2;
  var IndexPlayers = ss.getSheetByName("Players").getIndex();
  var FormsCreated = 0;
  var FormsDeleted = 0;
  
  var ui;
  var title;
  var msg;
  var uiResponse;
  
  // If Event Format is not Single or Team+Players, Pop up Error Message
  if(evntFormat != "Single" && evntFormat != "Team+Players"){
    ui = SpreadsheetApp.getUi();
    title = "Registration Forms Error";
    msg = "The Event does not support Players Registration. Please review Event configuration";
    uiResponse = ui.alert(title, msg, ui.ButtonSet.OK);
  }
  
    // Checks if Event Format is Team or Team+Players
  if(evntFormat == "Single" || evntFormat == "Team+Players"){
    // If Form Exists, Log Error Message
    if(FormIdEN != '' || FormIdFR != ''){
      ErrorVal = 1;
      ui = SpreadsheetApp.getUi();
      title = "Players Registration Forms";
      msg = "The Registration Forms already exist. Click OK to overwrite.";
      uiResponse = ui.alert(title, msg, ui.ButtonSet.OK_CANCEL);
      
      if(uiResponse == "OK"){
        // Clear IDs and URLs
        shtConfig.getRange(rowFormEN, colFormID).clearContent();
        shtConfig.getRange(rowFormFR, colFormID).clearContent();
        shtConfig.getRange(rowFormEN, colFormURL).clearContent(); 
        shtConfig.getRange(rowFormFR, colFormURL).clearContent();
        
        // If Responses Sheets exist, Unlink and Delete them
        shtResp1 = ss.getSheets()[IndexPlayers];
        shtRespName1 = shtResp1.getName();
        shtResp2 = ss.getSheets()[IndexPlayers+1];
        shtRespName2 = shtResp2.getName();
        
        // First Sheet After Players is RegPlyr EN
        if(shtRespName1 == "RegPlyr EN"){
          FormApp.openById(FormIdEN).removeDestination();
          ss.deleteSheet(shtResp1);
        }
        
        // Second Sheet After Players is RegPlyr FR
        if(shtRespName2 == "RegPlyr FR"){
          FormApp.openById(FormIdFR).removeDestination();
          ss.deleteSheet(shtResp2);
        }
        
        // First Sheet After Players is RegPlyr FR
        if(shtRespName1 == "RegPlyr FR"){
          FormApp.openById(FormIdFR).removeDestination();
          ss.deleteSheet(shtResp1);
        }
        
        // Second Sheet After Players is RegPlyr EN
        if(shtRespName2 == "RegPlyr EN"){
          FormApp.openById(FormIdEN).removeDestination();
          ss.deleteSheet(shtResp2);
        }
        
        // Forms Deleted Flag
        FormsDeleted = 1;
      }
    }
    
    // Create Forms
    if ((FormIdEN == "" && FormIdFR == "") || FormsDeleted == 1){
      // Create Forms
      FormNameEN = evntLocation + " " + evntName + " Player Registration EN";
      formEN = FormApp.create(FormNameEN).setTitle(FormNameEN);
      
      FormNameFR = evntLocation + " " + evntName + " Player Registration FR";
      formFR = FormApp.create(FormNameFR).setTitle(FormNameFR);
      
      // Loops in Response Columns Values and Create Appropriate Question
      for(var i = 1; i < cfgRegFormCnstrVal.length; i++){
        // Check for Question Order in Response Column Value in Configuration File
        if(QuestionOrder == cfgRegFormCnstrVal[i][1]){
          
          switch(cfgRegFormCnstrVal[i][0]){
              
              // EMAIL
            case 'Email': {
              // Set Registration Email collection
              formEN.setCollectEmail(true);
              formFR.setCollectEmail(true);
              break;
            }
              // FULL NAME
            case 'Full Name': {
              // ENGLISH
              formEN.addTextItem()
              .setTitle("Name")
              .setHelpText("Please, Remove any space at the beginning or end of the name")
              .setRequired(true);
              
              // FRENCH
              formFR.addTextItem()
              .setTitle("Nom")
              .setHelpText("SVP, enlevez les espaces au début ou à la fin du nom")
              .setRequired(true);
              break;
            }
              // FIRST NAME
            case 'First Name': {
              // ENGLISH
              formEN.addTextItem()
              .setTitle("First Name")
              .setHelpText("Please, Remove any space at the beginning or end of the name")
              .setRequired(true);
              
              // FRENCH
              formFR.addTextItem()
              .setTitle("Prénom")
              .setHelpText("SVP, enlevez les espaces au début ou à la fin du nom")
              .setRequired(true);
              break;
            }
              // LAST NAME
            case 'Last Name': {
              // ENGLISH
              formEN.addTextItem()
              .setTitle("Last Name")
              .setHelpText("Please, Remove any space at the beginning or end of the name")
              .setRequired(true);
              
              // FRENCH
              formFR.addTextItem()
              .setTitle("Nom de Famille")
              .setHelpText("SVP, enlevez les espaces au début ou à la fin du nom")
              .setRequired(true);
              break;
            }
              // LANGUAGE
            case 'Language': {
              // ENGLISH
              formEN.addMultipleChoiceItem()
              .setTitle("Language Preference")
              .setHelpText("Which Language do you prefer to use? The application is available in English and French")
              .setRequired(true)
              .setChoiceValues(["English","Français"]);
              
              // FRENCH
              formFR.addMultipleChoiceItem()
              .setTitle("Préférence de Langue")
              .setHelpText("Quelle langue préférez-vous utiliser? L'application est disponible en anglais et en français.")
              .setRequired(true)
              .setChoiceValues(["English","Français"]);
              break;
            }
              // PHONE NUMBER
            case 'Phone Number': {
              // ENGLISH
              formEN.addTextItem()
              .setTitle("Phone Number")
              .setRequired(true);
              
              // FRENCH
              formFR.addTextItem()
              .setTitle("Numéro de téléphone")
              .setRequired(true);
              break;
            }
              // TEAM NAME
            case 'Team Name': {
              if(evntFormat == 'Team'){
                // ENGLISH
                formEN.addPageBreakItem().setTitle("Team");
                formEN.addTextItem()
                .setTitle("Team Name")
                .setRequired(true);
                
                // FRENCH
                formFR.addPageBreakItem().setTitle("Équipe");
                formFR.addTextItem()
                .setTitle("Nom d'équipe")
                .setRequired(true);
              }
              break;
            }
              // DCI NUMBER
            case 'DCI Number': {
              formEN.addTextItem()
              .setTitle("DCI Number")
              .setRequired(true);
              
              formFR.addTextItem()
              .setTitle("Numéro DCI")
              .setRequired(true);
              break;
            }
              // DECK DEFINITION
            case 'Deck Definition': {
              Logger.log('%s - %s',QuestionOrder,cfgRegFormCnstrVal[i][0]); 
              // ENGLISH
              formEN.addPageBreakItem()
              .setTitle("Deck Definition");
              
              // FRENCH
              formFR.addPageBreakItem()
              .setTitle("Définition de Deck");
              
              break;
            }
              // DECK LIST
            case 'Deck List': {
              Logger.log('%s - %s',QuestionOrder,cfgRegFormCnstrVal[i][0]);  
              // Army List
              // ENGLISH
              formEN.addPageBreakItem()
              .setTitle("Deck List");
              formEN.addTextItem()
              .setTitle("Deck List")
              .setHelpText("Please, enter your Deck List. One Line per card with the following format: N CARD ex. 3 Counterspell")
              .setRequired(true);
              
              // FRENCH
              formFR.addPageBreakItem()
              .setTitle("Deck List");
              formFR.addTextItem()
              .setTitle("Deck List")
              .setHelpText("SVP, entrez la liste de carte qui composent votre Deck. Une ligne par carte, en suivant le format suivant: N CARTE ex. 3 Counterspell")
              .setRequired(true);
              break;
            }
              
              // COMMANDER NAME
            case 'Deck Commander (EDH)' :{
              // ENGLISH
              formEN.addTextItem()
              .setTitle("Deck Commander (EDH)")
              .setHelpText("Enter your Legendary Creature card your Deck's Commander")
              .setRequired(true); 
              
              // FRENCH
              formFR.addTextItem()
              .setTitle("Commandeur de votre Deck")
              .setHelpText("Entrez le nom de la carte qui de votre armée")
              .setRequired(true); 
              break;
            }
          }
          // Increment to Next Question
          QuestionOrder++;
          // Reset Loop if new question was added
          i = -1;
        }
        // Forms Created Flag
        FormsCreated = 1;
      }
      
      // RESPONSE SHEETS
      // Create Response Sheet in Main File and Rename
      if(exeGnrtResp == "Enabled" && FormsCreated == 1){
        Logger.log("Generating Response Sheets and Form Links");
        var IndexPlayers = ss.getSheetByName("Players").getIndex();
        // English Form
        formEN.setDestination(FormApp.DestinationType.SPREADSHEET, ssID);
        
        // Find and Rename Response Sheet
        ss = SpreadsheetApp.openById(ssID);
        ssSheets = ss.getSheets();
        ssSheets[0].setName('RegPlyr EN');
        // Move Response Sheet to appropriate spot in file
        shtResp = ss.getSheetByName('RegPlyr EN');
        ss.moveActiveSheet(IndexPlayers+1);
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
        formFR.setDestination(FormApp.DestinationType.SPREADSHEET, ssID);
        
        // Find and Rename Response Sheet
        ss = SpreadsheetApp.openById(ssID);
        ssSheets = ss.getSheets();
        ssSheets[0].setName('RegPlyr FR');
        
        // Move Response Sheet to appropriate spot in file
        shtResp = ss.getSheetByName('RegPlyr FR');
        ss.moveActiveSheet(IndexPlayers+2);
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
        urlFormEN = formEN.getPublishedUrl();
        shtConfig.getRange(rowFormEN, colFormURL).setValue(urlFormEN); 
        
        urlFormFR = formFR.getPublishedUrl();
        shtConfig.getRange(rowFormFR, colFormURL).setValue(urlFormFR);
        
        Logger.log("Response Sheets and Form Links Generated");
        
        // Format Players Sheet
        // Hide Unused Columns
        // Loop through RespCol and Hide Matching Table Column if RespCol == "" 
        
        
      }
    }
  }
  // Post Log to Log Sheet
  subPostLog(shtLog, Logger.getLog());
}