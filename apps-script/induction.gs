/**
 * NBF site induction form - receiver.
 * Saves each submission as a row in a private Google Sheet and stores the filled-in PDF in a private Drive folder.
 * Everything is created in the Google account that deploys this script, and stays private to that account.
 * Run setup() once to authorise it, then deploy as a Web app (Execute as: Me, Who has access: Anyone).
 */
var SHEET_NAME = "NBF Site Induction Responses";
var FOLDER_NAME = "NBF Site Induction Forms (private)";
var MAX_PDF_CHARS = 4000000;

var COLUMNS = [
  ["Submitted","submitted"],["Reference","ref"],["Name","name"],["Date of birth","dob"],["Mobile","mobile"],["Home address","address"],
  ["Company","company"],["Nationality","nationality"],["Induction type","inductionType"],["Job number","jobNumber"],["Start date","startDate"],
  ["Worked for JBD before","workedBefore"],
  ["Next of kin","kinName"],["Relationship","kinRel"],["Next of kin home tel","kinHome"],["Emergency tel","kinEmergency"],
  ["Disabilities","disabilities"],["Asthma","asthma"],["Heart conditions","heart"],["Eye conditions","eye"],["Epilepsy","epilepsy"],
  ["Allergies","allergies"],["Lung conditions","lung"],["Hearing conditions","hearing"],["Back problems","back"],["HAVS","havs"],
  ["Other health","healthOther"],["Health details","healthDetails"],
  ["Taking medication","medication"],["Medication for","medFor"],["Medication kept","medWhere"],
  ["GP contact permission","gpPermission"],["Last GP visit","gpLast"],["GP name","gpName"],["GP tel","gpTel"],
  ["Cannot work at height","atHeight"],["Cannot work below ground","belowGround"],["Cannot work in confined spaces","confined"],
  ["Cannot wear breathing apparatus","breathing"],["Other restrictions","condOther"],
  ["Valid UK licence","licence"],["Driving company vehicles","companyVehicles"],["Driving offence (5 yrs)","convicted"],
  ["CSCS/CPCS reg no","cscsNo"],
  ["SafePass (Solas)","t_cscs"],["SafePass (Solas) expiry","x_cscs"],["CPCS card","t_cpcs"],["CPCS expiry","x_cpcs"],["Supervisor Course","t_sssts"],["Supervisor Course expiry","x_sssts"],
  ["First aid","t_firstAid"],["First aid expiry","x_firstAid"],["Manual handling","t_manual"],["Manual handling expiry","x_manual"],
  ["Abrasive wheels","t_abrasive"],["Abrasive wheels expiry","x_abrasive"],["Confined spaces training","t_confinedT"],["Confined spaces expiry","x_confinedT"],
  ["Cable avoidance","t_cable"],["Cable avoidance expiry","x_cable"],["Hand arm vibration training","t_hav"],["HAV expiry","x_hav"],
  ["New roads & street works","t_roads"],["New roads expiry","x_roads"],["Other training","trainOther"],
  ["Declaration signed date","signDate"],["PDF","pdfUrl"]
];

// Earlier column headings, renamed in place so existing sheets keep their data and line up with the new names.
var RENAMED = { "CSCS card": "SafePass (Solas)", "CSCS expiry": "SafePass (Solas) expiry", "SSSTS": "Supervisor Course", "SSSTS expiry": "Supervisor Course expiry" };

function setup() {
  var s = getSheet_(); getFolder_();
  Logger.log("Sheet: " + s.getParent().getUrl());
  Logger.log("PDF folder: " + getFolder_().getUrl());
}

function doGet() { return out_({ ok: true, service: "NBF site induction form" }); }

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var d = JSON.parse(e.postData.contents);
    if (d.hp) return out_({ ok: true });
    var f = d.fields || {};
    if (!f.name || !d.pdf || String(d.pdf).length > MAX_PDF_CHARS) return out_({ ok: false, error: "Invalid submission." });
    var ref = String(d.ref || "").replace(/[^A-Za-z0-9\-]/g, "").slice(0, 30) || ("IND-" + Date.now());
    var stamp = Utilities.formatDate(new Date(), "Europe/Dublin", "dd-MM-yyyy");
    var fileName = "Induction - " + String(f.name).replace(/[\\\/:*?"<>|]/g, "").slice(0, 60) + " - " + stamp + " - " + ref + ".pdf";
    var blob = Utilities.newBlob(Utilities.base64Decode(d.pdf), "application/pdf", fileName);
    var file = getFolder_().createFile(blob);
    var sheet = getSheet_();
    var values = {};
    COLUMNS.forEach(function (c) { values[c[0]] = f[c[1]] === undefined ? "" : f[c[1]]; });
    values["Submitted"] = Utilities.formatDate(new Date(), "Europe/Dublin", "dd/MM/yyyy HH:mm");
    values["Reference"] = ref;
    values["PDF"] = file.getUrl();
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    sheet.appendRow(headers.map(function (h) { return values[h] === undefined ? "" : String(values[h]); }));
    return out_({ ok: true, ref: ref });
  } catch (err) {
    return out_({ ok: false, error: "Could not save the form. Please try again." });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function getSheet_() {
  var p = PropertiesService.getScriptProperties(), id = p.getProperty("SHEET_ID"), ss;
  if (id) { ss = SpreadsheetApp.openById(id); }
  else {
    ss = SpreadsheetApp.create(SHEET_NAME);
    p.setProperty("SHEET_ID", ss.getId());
    var sh = ss.getSheets()[0];
    var heads = COLUMNS.map(function (c) { return c[0]; });
    sh.getRange(1, 1, 1, heads.length).setValues([heads]).setFontWeight("bold").setBackground("#012060").setFontColor("#ffffff");
    sh.setFrozenRows(1); sh.setFrozenColumns(3);
  }
  var sheet = ss.getSheets()[0];
  var head = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0], changed = false;
  head = head.map(function (h) { if (RENAMED[h]) { changed = true; return RENAMED[h]; } return h; });
  if (changed) sheet.getRange(1, 1, 1, head.length).setValues([head]);
  return sheet;
}

function getFolder_() {
  var p = PropertiesService.getScriptProperties(), id = p.getProperty("FOLDER_ID");
  if (id) return DriveApp.getFolderById(id);
  var f = DriveApp.createFolder(FOLDER_NAME);
  p.setProperty("FOLDER_ID", f.getId());
  return f;
}
