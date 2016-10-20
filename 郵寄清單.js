// This constant is written in column C for rows for which an email
// has been sent successfully.
var EMAIL_SENT = "EMAIL_SENT";

function sendEmails() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 2;  // 從第幾行開始寄送
  var numRows = sheet.getMaxRows() - 1;   // 總共要寄送最大行數減一（標題不需要寄送）
  // Fetch the range of cells A2:B3
  var dataRange = sheet.getRange(startRow, 1, numRows, 5);    //getRange(row, column, numRows, numColumns)
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  for (var i = 0; i < data.length; ++i) {
    var row = data[i];
    
    var class=row[0];
    var student=row[1];
    var subject=row[4];
    var teacher=row[2];
    var teacher_id=row[3];
    var emailSent = row[5];
    
    var emailAddress=teacher_id + "@shsh.ylc.edu.tw";
    var title="請您參加 8/20 "+ class + student + "的期初IEP會議";
    
    // 內文如果要用 HTML 格式，可以在協作平台中打好，然後直接寄送該網址。
    var site = SitesApp.getPageByUrl('https://sites.google.com/a/shsh.ylc.edu.tw/iep-collection/iep-kai-hui-tong-zhi');
    var body = site.getHtmlContent();
    message = body.replace("{class}", class);
    message = message.replace("{subject}", subject);
    message = message.replace("{teacher}", teacher);
    message = message.replace("{student}", student);
    
    
    if (emailSent != EMAIL_SENT) {  // 避免重複寄信的設定
      
      // 進行寄信動作
      MailApp.sendEmail({
        to: emailAddress,
        subject: title, 
        htmlBody: message} );
      
      sheet.getRange(startRow + i, 6).setValue(EMAIL_SENT);
      // Make sure the cell is updated right away in case the script is interrupted
      SpreadsheetApp.flush();
    }
  }
}
