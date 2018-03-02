function arrangeEmail() {
  var sss = SpreadsheetApp.openById('1nMZyk60hkShbQX8p5n-ROWLGgfEXyyL2wi73hhFJ6vY');
  var ss = sss.getSheetByName('data');
  //var range = ss.getRange('A:I'); //assign the range you want to copy
  var range = ss.getDataRange();
  var data = range.getValues();

  var tss = SpreadsheetApp.openById('1nMZyk60hkShbQX8p5n-ROWLGgfEXyyL2wi73hhFJ6vY');
  var ts = tss.getSheetByName('output');
  
  var count = 1;
  var finalData = "";
  for(i = 0; i < data.length; i++){
    if(finalData == ""){
      finalData = data[i][0];
    }
    else{
      finalData = finalData + "," +data[i][0];
    }
    count++;
    
    if(count%25 == 0){
      Logger.log(finalData);
      ts.getRange(count/25,1).setValue(finalData);
      finalData = "";
    } 
  }
  
  if(finalData != ""){
    ts.getRange((count+25)/25,1).setValue(finalData);
  }

}