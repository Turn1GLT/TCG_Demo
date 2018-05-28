// **********************************************
// function fcnUpdateCardList()
//
// This function updates the Player card database  
// with the list of cards sent in arguments
//
// **********************************************

function fcnUpdateCardList(shtConfig, Player, shtTest){
  
  // Card List Spreadsheet
  var shtIDs = shtConfig.getRange(4,7,24,1).getValues();
  var shtCardListEn = SpreadsheetApp.openById(shtIDs[3][0]).getSheetByName(Player);
  var CardListEnMaxRows = shtCardListEn.getMaxRows();
  var shtCardListFr = SpreadsheetApp.openById(shtIDs[4][0]).getSheetByName(Player);
  var CardListFrMaxRows = shtCardListFr.getMaxRows();
  var rngCardListEn = shtCardListEn.getRange(6, 1, CardListEnMaxRows-6, 5); // 0 = Card Qty, 1 = Card Number, 2 = Card Name, 3 = Rarity, 4 = Set Name 
  var rngCardListFr = shtCardListFr.getRange(6, 1, CardListFrMaxRows-6, 5); // 0 = Card Qty, 1 = Card Number, 2 = Card Name, 3 = Rarity, 4 = Set Name 
  var CardList; // Where Card Data will be populated
  
  var shtCardDB = SpreadsheetApp.openById(shtIDs[2][0]).getSheetByName(Player);
  var CardDBSetTotal = shtCardDB.getRange(2,1,1,48).getValues(); // Gets Sum of all set Quantity, if > 0, set is present in card pool
  var CardTotal = shtCardDB.getRange(3, 7).getValue();
  var BstrTotal = shtCardDB.getRange(3,11).getValue();
  var SetData;
  var SetName;
  var colSet;
  var CardNb = 0;
    
  // Clear Player Card Pool
  rngCardListEn.clearContent();
  CardList = rngCardListEn.getValues();
  rngCardListFr.clearContent();
  CardList = rngCardListFr.getValues();
    
  // Look for Set with cards present in pool
  for (var col = 0; col <= 48; col++){   
    // if Set Card Quantity > 0, Set has cards in pool, Loop through all cards in Set
    if (CardDBSetTotal[0][col] > 0){
      colSet = col + 1;
      SetName = shtCardDB.getRange(6,colSet+2).getValue();

      // Get all Cards Data from set
      SetData = shtCardDB.getRange(7, colSet, 301, 4).getValues();
      
      // Loop through each card in Set and get Card Data
      for (var CardID = 1; CardID < 300; CardID++){
        if (SetData[CardID][0] > 0) {
          CardList[CardNb][0] = SetData[CardID][0]; // Quantity
          CardList[CardNb][1] = SetData[CardID][2]; // Card Name
          CardList[CardNb][2] = SetData[CardID][3]; // Card Rarity
          CardList[CardNb][3] = SetData[CardID][1]; // Card Number (ID)
          CardList[CardNb][4] = SetName;            // Set Name    
          CardNb++;
        }
      }
    }
  }
  // Updates the Player Card Pool
  rngCardListEn.setValues(CardList);
  shtCardListEn.getRange(3,1).setValue(BstrTotal);
  shtCardListEn.getRange(4,1).setValue(CardTotal);
  rngCardListFr.setValues(CardList);
  shtCardListFr.getRange(3,1).setValue(BstrTotal);
  shtCardListFr.getRange(4,1).setValue(CardTotal);
  
  
  // Return Value
}
