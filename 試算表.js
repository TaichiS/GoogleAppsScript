/* 可以串接要執行的東西 */
function onOpen() {
  
  // 如果這個功能是在 Google Docs 使用，則應該使用 DocumentApp.getUi()....
  SpreadsheetApp.getUi().createMenu('自定選單')
  .addItem('彈出警告視窗', 'showAlert')  // showAlert
  .addItem('彈出提示視窗', 'showPrompt')  // showPrompt
  .addItem('彈出對話框', 'showDialog')   // showDialog
  .addSeparator()
  .addItem('文件設定成「公開可編輯」', 'setOpen')   // setOpen
  .addItem('文件設定成「不公開」', 'setPrivate')  // setPrivate
  .addToUi();
}

// 彈出警告視窗，可以使用 alert 
function showAlert(){
   SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
     .alert('嗨～你按了彈出警告視窗！');
}

// 也可以跟使用者取值（是或否），這裡順便示範沒有串接的寫法。
function showPrompt(){
  
  var ui = SpreadsheetApp.getUi(); // Same variations.

  var result = ui.alert(
     '請確認',
     '您是否要繼續？',
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (result == ui.Button.YES) {
    // User clicked "Yes".
    ui.alert('您已經按了確認YES！');
  } else {
    // User clicked "No" or X in the title bar.
    ui.alert('您已經按了取消NO！');
  }
  
}

function showDialog(){
  var ui = SpreadsheetApp.getUi(); // Same variations.

  var result = ui.prompt(
      'Let\'s get to know each other!',
      'Please enter your name:',
      ui.ButtonSet.OK_CANCEL);

  // Process the user's response.
  var button = result.getSelectedButton();
  var text = result.getResponseText();
  if (button == ui.Button.OK) {
    // User clicked "OK".
    ui.alert('Your name is ' + text + '.');
  } else if (button == ui.Button.CANCEL) {
    // User clicked "Cancel".
    ui.alert('I didn\'t get your name.');
  } else if (button == ui.Button.CLOSE) {
    // User clicked X in the title bar.
    ui.alert('You closed the dialog.');
  }
}

function setOpen(){
   myid = SpreadsheetApp.getActiveSpreadsheet().getId();
   DriveApp.getFileById(myid).setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.EDIT);
}

function setPrivate(){
   myid = SpreadsheetApp.getActiveSpreadsheet().getId();
   DriveApp.getFileById(myid).setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
}  
