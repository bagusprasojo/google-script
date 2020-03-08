function GenerateURL_Edit() {
  var form = FormApp.openById('1O3U4n9Iz0bOYJN16dCvmYDFBg3reH1fcpc9BDSv4-9w');
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Daftar');
  var data = sheet.getDataRange().getValues();
  var urlCol = 29;
  var responses = form.getResponses();  
  
  var d = new Date('February 17, 2018 13:00:00 -0500');
  var lastEdit = d.getTime();
  var barisKirimEmail = -1;
  
  
  var timestamps = [], urls = [], resultUrls = [];
  
  for (var i = 0; i < responses.length; i++) {
    timestamps.push(responses[i].getTimestamp().setMilliseconds(0));
    
    if (responses[i].getTimestamp().setMilliseconds(0) > lastEdit) {
      lastEdit = responses[i].getTimestamp().setMilliseconds(0); 
      barisKirimEmail = i; 
    }    
    
    
    urls.push(responses[i].getEditResponseUrl());
  }
  
  for (var j = 1; j < data.length; j++) {
    resultUrls.push([data[j][1]?urls[timestamps.indexOf(data[j][0].setMilliseconds(0))]:'']);
    //ui.alert(resultUrls[j]);
  }
  
  sheet.getRange(2, urlCol,resultUrls.length).setValues(resultUrls);   
  
  if (barisKirimEmail > -1){
    //var emailAddress = responses[barisKirimEmail].getItemResponses()[1].getResponse();
    //var message = responses[barisKirimEmail].getEditResponseUrl();
    //MailApp.sendEmail(emailAddress, 'Test', message);
    
    SendEmail(responses, barisKirimEmail);
    
//    responses[1]
    
  }
  
}

function SendEmail(AResponses, ABaris){
  
  var templ = HtmlService.createTemplateFromFile('candidate-email');
  
  var candidate = 
      {
        name : AResponses[ABaris].getItemResponses()[0].getResponse(),
        email: AResponses[ABaris].getItemResponses()[1].getResponse(),
        urlEdit: AResponses[ABaris].getEditResponseUrl(),
        data : ""
      };
  
  templ.candidate = candidate;
  var message = templ.evaluate().getContent();
  
  
  MailApp.sendEmail({
    to: candidate.email,
    subject: "Terima Kasih " + candidate.name + " Anda Telah Melakukan Pendaftaran Seminar Kami",
    htmlBody: message
  });
  
  var emailQuotaRemaining = MailApp.getRemainingDailyQuota();
  Logger.log("Remaining email quota: " + emailQuotaRemaining);
}

function onOpen() 
{
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Approve Peserta Seminar')
      .addItem('Approve', 'doApprove')
      .addItem('Un Approve', 'doApprove')
      .addToUi();
}