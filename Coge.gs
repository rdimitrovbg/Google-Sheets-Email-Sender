const ui = SpreadsheetApp.getUi();
const sent_to = ["@","@","@"];
const sheets = ["application", "employees", "holidays"];
const sheetToPDF = sheets[0];

var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheetApp = ss.getSheetByName(sheets[0]);
var sheetEmployees = ss.getSheetByName(sheets[1]);
var sheetHolydays = ss.getSheetByName(sheets[2]);
/* ========================================================================== */
/* ========================================================================== */
function onOpen(e) {
  lock();
  createSendMenu();
  whoami();
  ui.alert('Use \'📧 Send Application\' button to send the email located at the menu bar.',
    '', ui.ButtonSet.OK);
  unLock();
}
/* ========================================================================== */
function createSendMenu() {
  var menu = ui.createMenu('📧 Send Application');
  menu.addItem('📤 Send', 'sendMailAction')
  menu.addToUi();
}
/* ========================================================================== */
function sendMailAction() {
  if (showConfirmSendEmailAlert()) {
    sendMail();
  }
}
/* ========================================================================== */
function whoami() {
  var email = Session.getActiveUser().getEmail();
  sheetEmployees.getRange('F1').setValue(email);
}
/* ========================================================================== */
function lock() {
  ss.toast(arguments.callee.name, 'TODO :)', 1);
  return;
}
/* ========================================================================== */
function unLock() {
  ss.toast(arguments.callee.name, 'TODO :)', 1);
  return;
}
/* ========================================================================== */
function showConfirmSendEmailAlert() {
  const title = "Please confirm";
  const message = [
    "Hello ".concat(sheetEmployees.getRange('G1').getValue()),
    "Are you sure you want to send this email?",
    "То:", sent_to.join("\r\n")
  ];

  return ui.Button.YES == ui.alert(title, message.join("\r\n"), ui.ButtonSet.YES_NO);
}
/* ========================================================================== */
function hideSheets() {
  ss.getSheets().forEach(
    function (s) {
      if (s.getName() !== sheetToPDF) { s.hideSheet(); }
    });
}
/* ========================================================================== */
function getDataFromCell(arg_1) {
  return sheetApp.getRange(arg_1).getValue()
}
/* ========================================================================== */
function calendarCalculation() {
  const dateCellFrom = "C28:D28";
  const dateCellTo = "F28:G28";
  const dateCellOff = "C26:D26";

  var dateFrom = new Date(getDataFromCell(dateCellFrom).toString());
  var dateTo = new Date(getDataFromCell(dateCellTo).toString());
  var offDays = getDataFromCell(dateCellOff).toString();

  var formatDate = function (date) {
    var d = date.getDate() < 10 ? "0" + date.getDate() : date.getDate();
    var m = (date.getMonth() + 1) < 10 ? "0" + (date.getMonth() + 1) : (date.getMonth() + 1);
    return d + "." + m + "." + date.getFullYear();
  };

  return {
    from: formatDate(dateFrom),
    to: formatDate(dateTo),
    daysOff: offDays
  };
}
/* ========================================================================== */
function sendMail() {
  var dt = calendarCalculation();
  var name_cell = "C14:I14";
  hideSheets();
  var bodyMessage = ["Hello there!",
    `The file is a leave request for ${dt.daysOff} ${dt.daysOff == 1 ? 'day' : 'days'}, in the period from/to - ${dt.from}/${dt.to}.`,
    "Thank you!"];

  var message = {
    to: sent_to.join(";"),
    subject: "application for leave approval",
    body: bodyMessage.join("\r\n"),
    attachments: {
      attachments: [SpreadsheetApp.getActiveSpreadsheet().getAs(MimeType.PDF).setName(Utilities.formatString("application-for-leave-approval-%s-[%s-%s].pdf", sheetEmployees.getRange('G1').getValue(), dt.from, dt.to))],
      name: getDataFromCell(name_cell).toString()
    }
  }
  try {
    GmailApp.sendEmail(message.to, message.subject, message.body, message.attachments);
    ui.alert("Send successfully","The message has been send successfully",ui.ButtonSet.OK);
  } catch (error) {
    ui.alert("Error occurred", "An error occurred while sending an email.\r\n" + error, ui.ButtonSet.OK);
  }
}
/* ========================================================================== */
function debug() { }
