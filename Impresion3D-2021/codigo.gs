function getMessage() {
  var htmlOutput = HtmlService.createHtmlOutputFromFile('Cuerpo');
  
  var message = htmlOutput.getContent()
  var SS = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1IiLGuK6h0mL6zuw9UwW9MmakSk8Kv3Rx7stFYAaDCBA/edit#gid=657634916").getSheetByName("Respuestas de formulario 1")
  message = message.replace("%name", SS.getRange(SS.getLastRow(),3).getValue());
  //message = message.replace("%name", "Test.Subject");
  return message;
}

function sendEmail() {
  var SS = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1IiLGuK6h0mL6zuw9UwW9MmakSk8Kv3Rx7stFYAaDCBA/edit#gid=657634916").getSheetByName("Respuestas de formulario 1");
  Utilities.sleep(1999);
  SpreadsheetApp.flush();
  SS = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1IiLGuK6h0mL6zuw9UwW9MmakSk8Kv3Rx7stFYAaDCBA/edit#gid=657634916").getSheetByName("Respuestas de formulario 1");
  var message = getMessage()
  MailApp.sendEmail(SS.getRange(SS.getLastRow(),2).getValue(), "Inscripción a curso impresión 3D recibida!", message, {htmlBody : message});
}

function lateLooper()
{
  var SS = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1IiLGuK6h0mL6zuw9UwW9MmakSk8Kv3Rx7stFYAaDCBA/edit#gid=657634916").getSheetByName("Respuestas de formulario 1")
  for (let i=2;i<=SS.getLastRow();i++)
  {
    var htmlOutput = HtmlService.createHtmlOutputFromFile('Cuerpo');
    var message = htmlOutput.getContent()
    message = message.replace("%name", SS.getRange(i,3).getValue());
    Logger.log("a");
    Logger.log(SS.getRange(i,2).getValue());
    Logger.log(i);
    Utilities.sleep(500);
    MailApp.sendEmail(SS.getRange(i,2).getValue(), "Inscripción a curso impresión 3D recibida!", message, {htmlBody : message});
  }
}