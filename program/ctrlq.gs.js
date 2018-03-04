/*
 * @name Save Emails and Attachments
 * @version July 18, 2015
 * @author JellyFishTech
 * * тустовый гуглДиск папка= 1C6ljR4bjyKN28v281orNxzjyUPzyqGct
*/

function onOpen(e) {
  var menu = SpreadsheetApp.getUi().createMenu("🚀Save Emails")
  .addItem("Create New Rule", "createRules_");

  if (rulesCount() > 0)
    menu.addItem("Manage Rules", "manageRules_");

  menu.addSeparator()
  // .addItem("Video Tutorial & Support", "showHelpWindow_")
  .addToUi();
    var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('rules');
    var range = ss.getRange("F1:H1");
    ss.hideColumn(range);
}

/*
 * Создание правил
 * @private
 */
function createRules_() {

    var html = HtmlService.createTemplateFromFile("rules")
    .evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setTitle("Create a New Rule")
    .setWidth(580)
    .setHeight(365);

    SpreadsheetApp.getActive().show(html);
}

/*
 * Manage Rules
 * @private
 */
function manageRules_() {

  var html = HtmlService.createTemplateFromFile("manage");
  html.rules = getRulesFromSheet();
  html = html.evaluate()
  .setSandboxMode(HtmlService.SandboxMode.IFRAME)
  .setTitle("Manage Rules")
  .setWidth(400)
  .setHeight(120);

  SpreadsheetApp.getActive().show(html);
}

/*
 * Video Tutorial & Support
 * @private
 */
function showHelpWindow_() {

  var html = HtmlService.createHtmlOutputFromFile("support")
  .setTitle("Video Tutorial & Support")
  .setWidth(500)
  .setHeight(340);

  SpreadsheetApp.getActive().show(html);
}
