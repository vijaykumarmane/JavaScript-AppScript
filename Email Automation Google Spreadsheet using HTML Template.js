/* 

To get HTML template get sent or received email in Gmail brwoser and
do Inspect element and from it copy the HTML element of Email body. and done

*/

function onOpen() {
  //function runs by default when the google sheet is open  
  
  // add button on user interface 
  var user_interface = SpreadsheetApp.getUi();
  
  user_interface.createMenu("Send Email").addItem("Send to All", 'loopOn').addToUi();
  
}

function loopOn() {

var ss = SpreadsheetApp.getActiveSpreadsheet();  
var sheet = ss.getSheetByName("Email")
var sheet2 = ss.getSheetByName("Sent Data")
var rangeData = sheet2.getDataRange();
//var lastColumn = rangeData.getLastColumn();
var lastRowPaste = rangeData.getLastRow()+1;
// sheet.getRange(2,3).setValue(lastRow-1);
  
//sheet.getRange("A1").activate();
  //sheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate();
//var lastRow = sheet.getCurrentCell().getRowIndex()
  //sheet.getRange(2, 2).setFormula("=UNIQUE(A2:A)")
  //Utilities.sleep(10000)
  //sheet.getRange("B2:B").copyTo(sheet.getRange("B2:B"), {contentsOnly: true});
 //sheet.getRange(2,6).setValue(lastR);
 //sheet.getRange(2,7).setValue(lastRow);
  //sheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.UP).activate();
var mailQuota = (MailApp.getRemainingDailyQuota());
var iterate = Math.floor(mailQuota*0.5) + 1;
sheet.getRange(2,6).setValue(Math.floor(mailQuota));
var emailSent = 0;
sheet.getRange(2,7).setValue(emailSent);
var user = Session.getActiveUser().getEmail()
var date = new Date()

for(var i = 2; i <= iterate; i++) {
  
  if (sheet.getRange(i, 2).getValue() != "" && sheet.getRange(i, 3).getValue() == ""){
    
    var to_email = sheet.getRange(i, 2).getValue().trim();
    
    //To check valid email ID
    if (/^\w+([\.-]?\w+)*@\w+([\.-]?\w+)*(\.\w{2,3})+$/.test(to_email)){
    sendEmail(to_email);
    sheet.getRange(i,3).setValue("Sent");
    sheet.getRange(i, 4).setValue(user);
    sheet.getRange(i, 5).setValue(date);
    emailSent++
    sheet.getRange(2,7).setValue(emailSent);
    }
    else{
      sheet.getRange(i,3).setValue("Invalid Email ID");
      sheet.getRange(i, 4).setValue(user);
      sheet.getRange(i, 5).setValue(date);
    }
  }
  
  else{
    if(sheet.getRange(i, 3).getValue() != ""){
    sheet.getRange(i,3).setValue("Already sent");
     }
    if(sheet.getRange(i,3).getValue() == "Invalid Email ID"){
      sheet.getRange(i,3).setValue("Invalid Email ID");
    }
    else{
      break
    }
  }
  
sheet.getRange(2,6).setValue(MailApp.getRemainingDailyQuota());

}
  i = i - 1
  if(i > 1){
   
  sheet.getRange("B2:E"+i).copyTo(sheet2.getRange("A"+ lastRowPaste), {contentsOnly: true});
  sheet2.getRange("E"+ lastRowPaste).setValue(i +" " + lastRowPaste)
/*
  sheet2.getRange("E"+ lastRowPaste).setValue(i +" " + lastRowPaste)
  sheet.getRange("B2:B").copyTo(sheet.getRange("B2:B"), {contentsOnly: true});
  sheet2.insertRows(lastRowPaste, i);  
    lastRowPaste = lastRowPaste + 15;
    for(var j = 35; j > 0; j--){
      if(sheet2.getRange(lastRowPaste, 1).getValue() == ""){
        sheet2.deleteRow(lastRowPaste);
      }
      lastRowPaste--;
    }
     sheet2.insertRows(lastRowPaste+150, i+20);
*/  
  }
  
  
}


// send email function
function sendEmail(to_email){
  
  // final email body
  var html = templateEmail();
  
  
  // Send Email 
  MailApp.sendEmail({to:to_email,
                     cc:"mohitagrawal.pune@gmail.com",
                     replyTo: "mohit@interlinkcapital.in",
                     name: "Interlink Capital Advisors",
                     subject:"Claim Eligible State and Central government subsidies", htmlBody: html})
  
  // Create Draft
  //GmailApp.createDraft(to_email, "Claim Eligible State and Central government subsidies", html.evaluate().getContent(),
    //                   { cc:"mohitagrawal.pune@gmail.com"
      //                });

}


// email template and user data input to modify give function parameters
// For now is standard email for all
function templateEmail(){
  
  // load template email
  var main = HtmlService.createTemplateFromFile('email_template');
  
  // add user relatd values
  //main.user_data = user_data;
  
  // final email 
  var html = main.evaluate().getContent();
  
  return html;
  //return main;
}
