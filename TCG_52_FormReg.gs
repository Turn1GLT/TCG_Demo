// **********************************************
// function fcnCreateRegForm()
//
// This function creates the Registration Form 
// based on the parameters in the Config File
//
// **********************************************

function fcnCreateRegForm() {
  
  var ss = SpreadsheetApp.getActive();
  var shtConfig = ss.getSheetByName('Config');
  var shtPlayers = ss.getSheetByName('Players');
  var ssID = shtConfig.getRange(30,2).getValue();
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

  var RowFormUrlEN = 23;
  var RowFormUrlFR = 24;
  var RowFormIdEN = 38;
  var RowFormIdFR = 39;
  
  var ErrorVal = '';
  var QuestionOrder = 1;
  
  // Response Columns from Configuration File
  // [x][0] = Response Columns
  var colRegRespValues = shtConfig.getRange(56,6,12,2).getValues();
  
  // Response Columns
  var colRespEmail = colRegRespValues[0][1];
  var colRespName = colRegRespValues[1][1];
  var colRespFirstName = colRegRespValues[2][1];
  var colRespLastName = colRegRespValues[3][1];
  var colRespPhone = colRegRespValues[4][1];
  var colRespLanguage = colRegRespValues[5][1];
  var colRespDCI = colRegRespValues[6][1];
  var colRespTeamName = colRegRespValues[7][1];
  
  // Gets the Subscription ID from the Config File
  FormIdEN = shtConfig.getRange(RowFormIdEN, 2).getValue();
  FormIdFR = shtConfig.getRange(RowFormIdFR, 2).getValue();

  // If Form Exists, Log Error Message
  if(FormIdEN != ''){
    ErrorVal = 1;
    Logger.log('Error! FormIdEN already exists. Unlink Response and Delete Form');
  }
  // If Form Exists, Log Error Message
  if(FormIdEN != ''){
    ErrorVal = 1;
    Logger.log('Error! FormIdFR already exists. Unlink Response and Delete Form');
  }

  if (FormIdEN == '' && FormIdFR == ''){
    // Create Forms
    FormNameEN = shtConfig.getRange(3, 2).getValue() + " Registration EN";
    formEN = FormApp.create(FormNameEN).setTitle(FormNameEN);
    
    FormNameFR = shtConfig.getRange(3, 2).getValue() + " Registration FR";
    formFR = FormApp.create(FormNameFR).setTitle(FormNameFR);
    
    // Loops in Response Columns Values and Create Appropriate Question
    for(var i = 0; i < colRegRespValues.length; i++){
      // Look for Col Equal to Question Order
      if(QuestionOrder == colRegRespValues[i][1]){
        //
        switch(colRegRespValues[i][0]){
          case 'Email': {
            Logger.log('%s - %s',QuestionOrder,colRegRespValues[i][0]);
            // EMAIL
            // Set Registration Email collection
            formEN.setCollectEmail(true);
            formFR.setCollectEmail(true);
            break;
          }
          case 'Name': {
            Logger.log('%s - %s',QuestionOrder,colRegRespValues[i][0]); 
            // FULL NAME   
            formEN.addTextItem()
            .setTitle("Name")
            .setHelpText("Please, Remove any space at the end of the name")
            .setRequired(true);
            
            formFR.addTextItem()
            .setTitle("Nom")
            .setHelpText("SVP, enlevez les espaces à la fin du nom")
            .setRequired(true);
            break;
          }
          case 'First Name': {
            Logger.log('%s - %s',QuestionOrder,colRegRespValues[i][0]); 
            // FIRST NAME  
              formEN.addTextItem()
              .setTitle("First Name")
              .setHelpText("Please, Remove any space at the end of the name")
              .setRequired(true);
              
            formFR.addTextItem()
            .setTitle("Prénom")
            .setHelpText("SVP, enlevez les espaces à la fin du nom")
            .setRequired(true);
            break;
          }
          case 'Last Name': {
            Logger.log('%s - %s',QuestionOrder,colRegRespValues[i][0]); 
            // LAST NAME 
            formEN.addTextItem()
            .setTitle("Last Name")
            .setHelpText("Please, Remove any space at the end of the name")
            .setRequired(true);
            
            formFR.addTextItem()
            .setTitle("Nom de Famille")
            .setHelpText("SVP, enlevez les espaces à la fin du nom")
            .setRequired(true);
            break;
          }
          case 'Phone Number': {
            Logger.log('%s - %s',QuestionOrder,colRegRespValues[i][0]); 
            // PHONE NUMBER    
            formEN.addTextItem()
            .setTitle("Phone Number")
            .setRequired(true);
            
            formFR.addTextItem()
            .setTitle("Numéro de téléphone")
            .setRequired(true);
            break;
          }
          case 'Language': {
            Logger.log('%s - %s',QuestionOrder,colRegRespValues[i][0]); 
            // LANGUAGE
            formEN.addMultipleChoiceItem()
            .setTitle("Language Preference")
            .setHelpText("Which Language do you prefer to use? The application is available in English and French")
            .setRequired(true)
            .setChoiceValues(["English","Français"]);
            
            formFR.addMultipleChoiceItem()
            .setTitle("Préférence de Langue")
            .setHelpText("Quelle langue préférez-vous utiliser? L'application est disponible en anglais et en français.")
            .setRequired(true)
            .setChoiceValues(["English","Français"]);
            break;
          }
          case 'DCI': {
            Logger.log('%s - %s',QuestionOrder,colRegRespValues[i][0]); 
            // DCI NUMBER
            formEN.addTextItem()
            .setTitle("DCI Number")
            .setRequired(true);
            
            formFR.addTextItem()
            .setTitle("Numéro DCI")
            .setRequired(true);
            break;
          }
          case 'Team Name': {
            Logger.log('%s - %s',QuestionOrder,colRegRespValues[i][0]); 
            // TEAM NAME
            formEN.addTextItem()
            .setTitle("Team Name")
            .setRequired(true);
            
            formFR.addTextItem()
            .setTitle("Nom d'équipe")
            .setRequired(true);
            break;
          }
        }
        // Increment to Next Question
        QuestionOrder++;
        // Reset Loop 
        i = -1;
      }
    }

    // RESPONSE SHEETS
    // Create Response Sheet in Main File and Rename
    
    // English Form
    formEN.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
    
    // Find and Rename Response Sheet
    ss = SpreadsheetApp.openById(ssID);
    ssSheets = ss.getSheets();
    ssSheets[0].setName('Registration EN');
    // Move Response Sheet to appropriate spot in file
    shtResp = ss.getSheetByName('Registration EN');
    ss.moveActiveSheet(17);
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
    ssSheets[0].setName('Registration FR');
    
    // Move Response Sheet to appropriate spot in file
    shtResp = ss.getSheetByName('Registration FR');
    ss.moveActiveSheet(18);
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
    shtConfig.getRange(RowFormIdEN, 2).setValue(FormIdEN);
    FormIdFR = formFR.getId();
    shtConfig.getRange(RowFormIdFR, 2).setValue(FormIdFR);
    
    // Create Links to add to Config File  
    urlFormEN = formEN.getPublishedUrl();
    shtConfig.getRange(RowFormUrlEN, 2).setValue(urlFormEN); 
    
    urlFormFR = formFR.getPublishedUrl();
    shtConfig.getRange(RowFormUrlFR, 2).setValue(urlFormFR);
  }
}  