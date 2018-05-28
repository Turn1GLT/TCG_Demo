// **********************************************
// function fcnUpdateCardDB()
//
// This function updates the Player card database  
// with the list of cards sent in arguments
//
// **********************************************

function fcnUpdateCardDB(shtConfig, Player, CardList, PackData, shtTest){
  
  // Player Card DB Spreadsheet
  var shtIDs = shtConfig.getRange(4,7,24,1).getValues();
  var shtCardDB = SpreadsheetApp.openById(shtIDs[2][0]).getSheetByName(Player);
  var CardDBSet = shtCardDB.getRange(4,1,1,32).getValues();
  var MstrSet = shtCardDB.getRange(4,33,1,16).getValues();
  
  var ColCard = 0;
  var SetCardList = new Array(300);
  var CardNameList;
  var NbCardSet = 0;
  var NewCardName;
  var SetNum;
  var CardID;
  var CardQty;
  var CardNum;
  var CardName;
  var CardRarity;
  var CardListSet = CardList[0];
  var CardInfo; 
    
  // Updates the Set Name to return to Main Function
  PackData[0][0] = CardListSet;
  
  // Find Set Column according to Set in Cardlist (CardList[0]) and get all card quantities (first card starts at row 8, row 7 = card 0)
  for (var ColSet = 0; ColSet <= 31; ColSet++){   
    if (CardListSet == CardDBSet[0][ColSet]){
      ColCard = ColSet+1;
      ColSet = 32;
    }
  }

  // Get Card Info (Quantity, Card Number, Card Name, Rarity) // [0][0]= Card in Pack, [0][1]= Card Number, [0][2]= Card Name, [0][3]= Card Rarity
  CardInfo = shtCardDB.getRange(8, ColCard-2,300,4).getValues();
  
  // Build the Card List Comparator 
  for(var i=0; i<300; i++){
    if(CardNameList[i][0] != "") {
      // Concatenate Card Name and Rarity
      CardName   = CardInfo[i][2];
      CardRarity = CardInfo[i][3];
      SetCardList[i] = CardName + " - " + CardRarity;
      NbCardSet++;
    }
    else i = 301;
  }
  
  // Loop through each card in CardList to find the appropriate column to find card (Masterpiece or not)
  for (var CardListNb = 1; CardListNb <= 14; CardListNb++){
    // Get Card Name
    NewCardName = CardList[CardListNb];
    
    // Search for the Card in the DB
    for(var i=0; i<300; i++){
      if(NewCardName == SetCardList[i]) {
        // Update the Card DB
        CardQty    = CardInfo[i][0];
        CardNum    = CardInfo[i][1];
        CardName   = CardInfo[i][2];
        CardRarity = CardInfo[i][3];
        
        // Update Card Quantity in Card DB
        shtCardDB.getRange(i+8, ColCard-2).setValue(CardQty + 1);

        // Store Card Info to return to Main Function
        PackData[CardListNb][0] = CardListNb;  // Card in Pack
        PackData[CardListNb][1] = CardInfo[0][1]; // Card Number
        PackData[CardListNb][2] = CardInfo[0][2]; // Card Name
        PackData[CardListNb][3] = CardInfo[0][3]; // Card Rarity
        // Exit the Loop
        i = 301;
      }
    }
  }
  
  // Debug
  //shtTest.getRange(1,1,16,4).setValues(PackData);
  
  // Call function to generate clean card list from Player Card DB
  fcnUpdateCardList(shtConfig, Player, shtTest);
  
  // Return Value
  return PackData;
}


// **********************************************
// function fcnCardListPerSet(ssCardDB)
//
// This function updates the Player card database  
// with the list of cards sent in arguments
//
// **********************************************

function fcnCardListPerSet(ssCardDB, SetNum){

  var shtCardDB = ssCardDB.getSheetByName('Template');
  var CardDBSet = shtCardDB.getRange(4,1,1,32).getValues();
  var MstrSet = shtCardDB.getRange(4,33,1,16).getValues();
  
  // Routine Variables
  var ColCard = 0;
  var ColCardMstr = 0;
  var CardList = new Array(300);
  var CardNameList;
  var NbCardSet = 0;
  
  // Find Set Column according to Set Number. Card Column is Set Column + 1
  for (var ColSet = 0; ColSet <= 31; ColSet++){   
    if (SetNum == CardDBSet[0][ColSet]){
      ColCard = ColSet + 1;
      ColSet = 32;
    }
  }
  
  // Get Regular Card Names and Rarity
  CardNameList = shtCardDB.getRange(8, ColCard,300,2).getValues();
  // Detect the amount of cards in set
  for(var i=0; i<300; i++){
    if(CardNameList[i][0] != "") {
      // Concatenate Card Name and Rarity
      CardList[i] = CardNameList[i][0] + " - " + CardNameList[i][1];
      NbCardSet++;
    }
    else i = 301;
  }

  // Set the Card List Array Length
  CardList.length = NbCardSet;

  return CardList;
}
