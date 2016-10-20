function onEdit() {
  
  // 本程式剛寫好必須試執行，才能允許權限通過。
  
  // 取得目前的工作表
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // 取得最後一行的位置
  var row = sheet.getLastRow();
  // 因為第一行是標題，所以真正的報名人數 = 行數-1
  var num = row - 1;
  
  var receiver_name = sheet.getRange(row, 2).getValue();
  var receiver_email = sheet.getRange(row, 4).getValue();
  // 建構信箱如右： 詹嘉隆 <taichis@shsh.ylc.edu.tw>
  var receiver = receiver_name + " <" + receiver_email + ">";
  
  // 寄信
  MailApp.sendEmail(receiver, "感謝您報名 Google Apps Script 研習", receiver_name + "您好：  目前已經有"+num+"個人報名，我們的上課教材在 https://goo.gl/PB2R8C");
  
  // 有寄信的加個註記
  sheet.getRange(row, 5).setValue("已寄信");
    
}
