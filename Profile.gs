/**
 * ============================================================
 *  VOICE OF HERs — Member Registry
 *  FILE 2 of 2:  Profile.gs
 * ============================================================
 *
 *  On every form submission this:
 *    1. issues a member ID  (VOH + 2 letters + team code + number)
 *    2. builds a profile document from the organization's template
 *    3. saves it as a Word .docx in the profiles folder
 *    4. writes the ID and links into the responses spreadsheet
 *    5. emails the HR team with the document attached
 *
 *  Repeat submissions are recognised by phone number: the member keeps
 *  their original ID and the earlier document is replaced.
 * ============================================================
 */

// ==================== CONFIGURATION ====================

var VOH_EMAIL_ON_SUBMIT     = true;   // email HR on each submission
var VOH_SHOW_PHOTO_IN_SHEET = true;   // photo preview column, or a link
var VOH_DELETE_GOOGLE_DOC   = true;   // keep only the .docx

// How repeat submissions are recognised: 'phone', 'email', or 'none'
var VOH_DEDUPE_BY = 'phone';

// Used only if Forms.gs is missing or has not defined them
var VOH_FALLBACK_HR_EMAIL = 'hr@example.org';
var VOH_FALLBACK_LOGO_ID  = 'YOUR_LOGO_DRIVE_FILE_ID';

// ==================== DOCUMENT WORDING AND STYLE ====================

var VOH_DOC_TITLE      = 'Team Data (info)';
var VOH_ABOUT_HEADING  = 'About the person';
var VOH_ABOUT_QUESTION = 'About Employee';   // a hint only, matching is flexible

var VOH_PURPLE      = '#4A1F44';
var VOH_LIGHT_PANEL = '#F3EAF8';
var VOH_YELLOW      = '#F2C41A';

// ==================== TEAM CODES ====================
// Keys must match VOH_DEPARTMENTS in Forms.gs exactly.
// ID format: VOH + abbr + (code + sequence), e.g. the first HR member is VOHhr301.

var VOH_DEPARTMENT_CODES = {
  'Legal Team':              { abbr: 'le', code: 100 },
  'Secretary Team':          { abbr: 'se', code: 200 },
  'HR Team':                 { abbr: 'hr', code: 300 },
  'IT Team':                 { abbr: 'it', code: 400 },
  'Partnership Team':        { abbr: 'pa', code: 500 },
  'Sponsorship Team':        { abbr: 'sp', code: 600 },
  'Communication Team':      { abbr: 'co', code: 700 },
  'Accounting Team':         { abbr: 'ac', code: 800 },
  'Event and Training Team': { abbr: 'ev', code: 900 },
  'Logistics Team':          { abbr: 'lo', code: 1000 },
  'Graphics Team':           { abbr: 'gr', code: 2000 },
  'Project Management Team': { abbr: 'pr', code: 3000 },
  'Social Media Team':       { abbr: 'so', code: 4000 }
};

function vohHrEmail() {
  return (typeof VOH_HR_EMAIL !== 'undefined' && VOH_HR_EMAIL)
    ? VOH_HR_EMAIL : VOH_FALLBACK_HR_EMAIL;
}

function vohLogoId() {
  return (typeof VOH_LOGO_FILE_ID !== 'undefined' && VOH_LOGO_FILE_ID)
    ? VOH_LOGO_FILE_ID : VOH_FALLBACK_LOGO_ID;
}

// ==================== LOOKUPS ====================

function vohProps() { return PropertiesService.getScriptProperties(); }

function vohGetForm() {
  var id = vohProps().getProperty('VOH_FORM_ID');
  if (!id) throw new Error('No form yet. Run createAndBrandForm in Forms.gs first.');
  return FormApp.openById(id);
}

function vohGetFolder() {
  var id = vohProps().getProperty('VOH_FOLDER_ID');
  if (!id) throw new Error('No profiles folder yet.');
  return DriveApp.getFolderById(id);
}

function vohGetSheet() {
  var id = vohProps().getProperty('VOH_SHEET_ID');
  if (!id) throw new Error('No responses sheet yet.');
  return SpreadsheetApp.openById(id).getSheets()[0];
}

// ==================== TRIGGER ====================

function setUpProfileGenerator() {
  var form = vohGetForm();
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'onMemberFormSubmit') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('onMemberFormSubmit').forForm(form).onFormSubmit().create();
  Logger.log('Document generation is switched on.');
}

function onMemberFormSubmit(e) {
  vohProcessResponse(e.response, null);
}

// ==================== DUPLICATE CONTROL ====================

/** Run ONCE. Stops response editing and limits to one per Google account. */
function vohLockDownDuplicates() {
  var form = vohGetForm();
  form.setAllowResponseEdits(false);
  form.setLimitOneResponsePerUser(true);
  Logger.log('Response editing is OFF, one response per account is ON.');
}

/** Wipes the phone-to-ID index. Use when clearing test data. */
function vohClearMemberIndex() {
  var props = vohProps();
  var all = props.getProperties();
  var cleared = 0;
  for (var key in all) {
    if (key.indexOf('MEM_') === 0) { props.deleteProperty(key); cleared++; }
  }
  Logger.log('Cleared ' + cleared + ' entries from the member index.');
}

/** Lists the phone-to-ID index. */
function vohShowMemberIndex() {
  var all = vohProps().getProperties();
  var count = 0;
  Logger.log('--- member index ---');
  for (var key in all) {
    if (key.indexOf('MEM_') === 0) {
      Logger.log('  ' + key.substring(6) + '  ->  ' + all[key]);
      count++;
    }
  }
  Logger.log(count + ' members recorded.');
}

function vohIdentityKey(data) {
  if (VOH_DEDUPE_BY === 'none') return '';

  if (VOH_DEDUPE_BY === 'email' && data.email) {
    return 'MEM_E_' + data.email.trim().toLowerCase();
  }

  // keep digits only and use the last 9, so 0912345678 and
  // +251912345678 are recognised as the same person
  var digits = String(data.phone || '').replace(/\D/g, '');
  if (digits.length >= 9) return 'MEM_P_' + digits.slice(-9);

  if (data.email) return 'MEM_E_' + data.email.trim().toLowerCase();
  return '';
}

/** Moves any earlier documents for this member to the bin. */
function vohRemoveOldDocuments(folder, memberId) {
  var prefix = memberId + ' - ';
  var files = folder.getFiles();
  var removed = 0;
  while (files.hasNext()) {
    var f = files.next();
    if (f.getName().indexOf(prefix) === 0) { f.setTrashed(true); removed++; }
  }
  if (removed) Logger.log('Replaced ' + removed + ' earlier document(s) for ' + memberId);
}

// ==================== ANSWER MATCHING ====================

function vohGet(answers, title) {
  return (answers[title] && answers[title].value) ? String(answers[title].value) : '';
}

/**
 * Finds the About answer without relying on the question title:
 * the exact title, then any title containing "about", then the
 * long paragraph question.
 */
function vohFindAbout(formResponse, answers) {
  var exact = vohGet(answers, VOH_ABOUT_QUESTION);
  if (exact) { Logger.log('About: matched the exact title.'); return exact; }

  for (var title in answers) {
    if (title.toLowerCase().indexOf('about') > -1) {
      var byWord = vohGet(answers, title);
      if (byWord) { Logger.log('About: matched on "' + title + '".'); return byWord; }
    }
  }

  var responses = formResponse.getItemResponses();
  for (var i = 0; i < responses.length; i++) {
    var item = responses[i].getItem();
    if (item.getType() === FormApp.ItemType.PARAGRAPH_TEXT) {
      var v = responses[i].getResponse();
      if (v) {
        Logger.log('About: used the paragraph question "' + item.getTitle() + '".');
        return String(v);
      }
    }
  }

  Logger.log('About: nothing found — probably left blank.');
  return '';
}

/**
 * Works out which uploaded file is the photo and which is the ID.
 * The ID is checked FIRST, so a question titled "photo of your National Id"
 * is not mistaken for the passport photo.
 */
function vohFindUploads(formResponse) {
  var out   = { photoId: '', nationalId: '' };
  var found = [];

  var responses = formResponse.getItemResponses();
  for (var i = 0; i < responses.length; i++) {
    var item = responses[i].getItem();
    if (item.getType() !== FormApp.ItemType.FILE_UPLOAD) continue;
    var ids = [].concat(responses[i].getResponse() || []);
    if (!ids.length) continue;
    found.push({ title: item.getTitle(), lower: item.getTitle().toLowerCase(),
                 fileId: ids[0], used: false });
  }

  for (var a = 0; a < found.length; a++) {
    var la = found[a].lower;
    if (!out.nationalId &&
        (la.indexOf('kebele') > -1 || la.indexOf('national') > -1 ||
         la.indexOf('identity') > -1 || la.indexOf('identification') > -1)) {
      out.nationalId = found[a].fileId; found[a].used = true;
    }
  }

  for (var b = 0; b < found.length; b++) {
    if (found[b].used) continue;
    var lb = found[b].lower;
    if (!out.photoId && (lb.indexOf('passport') > -1 || lb.indexOf('photo') > -1)) {
      out.photoId = found[b].fileId; found[b].used = true;
    }
  }

  for (var c = 0; c < found.length; c++) {
    if (found[c].used) continue;
    if (!out.photoId)         { out.photoId    = found[c].fileId; found[c].used = true; }
    else if (!out.nationalId) { out.nationalId = found[c].fileId; found[c].used = true; }
  }

  Logger.log('--- uploads found: ' + found.length + ' ---');
  for (var d = 0; d < found.length; d++) {
    var slot = (found[d].fileId === out.photoId) ? 'PHOTO'
             : (found[d].fileId === out.nationalId) ? 'ID DOC' : 'unused';
    Logger.log('  "' + found[d].title + '" -> ' + slot);
  }
  return out;
}

/** Works out the status wording from the two yes/no questions. */
function vohDeriveStatus(answers) {
  var direct = vohGet(answers, 'Current Status');
  if (direct) return direct;
  var working  = vohGet(answers, 'Are you currently working?')  === 'Yes';
  var studying = vohGet(answers, 'Are you currently studying?') === 'Yes';
  if (working && studying) return 'Both working and studying';
  if (working)  return 'Working';
  if (studying) return 'Student';
  return 'Neither at the moment';
}

// ==================== PROCESS ONE RESPONSE ====================

function vohProcessResponse(formResponse, reuseId) {
  var answers = {};
  var responses = formResponse.getItemResponses();
  for (var i = 0; i < responses.length; i++) {
    var t = responses[i].getItem().getTitle();
    var v = responses[i].getResponse();
    if (!answers[t] || !answers[t].value) answers[t] = { value: v };
  }

  var email = vohGet(answers, 'Email Address');
  if (!email) {
    try { email = formResponse.getRespondentEmail() || ''; } catch (err) { email = ''; }
  }

  var emgName = vohGet(answers, 'Emergency Contact Name');
  if (!emgName) emgName = vohGet(answers, 'Emergency Contact Name and Relation');

  var placeOfWork  = vohGet(answers, 'Place of Work');
  var jobTitle     = vohGet(answers, 'Position / Job Title');
  var placeOfStudy = vohGet(answers, 'Place of Study');
  var fieldOfStudy = vohGet(answers, 'Field of Study');
  if (!placeOfWork && !placeOfStudy) {
    placeOfWork = vohGet(answers, 'Place of Work / Place of Study');
    jobTitle    = vohGet(answers, 'Position / Field of Study');
  }

  var uploads = vohFindUploads(formResponse);

  var data = {
    name:         vohGet(answers, 'Full Name'),
    phone:        vohGet(answers, 'Phone No'),
    email:        email,
    age:          vohGet(answers, 'Age'),
    joined:       vohGet(answers, 'Joined Date'),
    department:   vohGet(answers, 'Department at VOICE OF HERs'),
    isBoard:      vohGet(answers, 'Are you also a Board Member?') === 'Yes',
    status:       vohDeriveStatus(answers),
    placeOfWork:  placeOfWork,
    jobTitle:     jobTitle,
    placeOfStudy: placeOfStudy,
    fieldOfStudy: fieldOfStudy,
    about:        vohFindAbout(formResponse, answers),
    emgPhone:     vohGet(answers, 'Emergency Contact'),
    emgName:      emgName,
    relation:     vohGet(answers, 'Relation'),
    photoId:      uploads.photoId,
    nationalId:   uploads.nationalId
  };

  // ---------- member ID, with duplicate handling ----------
  var isUpdate = false;
  var key = vohIdentityKey(data);

  if (reuseId) {
    data.memberId = reuseId;
    isUpdate = true;
  } else {
    var known = key ? vohProps().getProperty(key) : null;
    if (known) {
      data.memberId = known;
      isUpdate = true;
      Logger.log('REPEAT SUBMISSION — keeping the existing ID ' + known);
    } else {
      data.memberId = vohGenerateMemberId(data.department, data.isBoard);
      if (key) vohProps().setProperty(key, data.memberId);
    }
  }

  var files = vohBuildProfileDocument(data);
  vohWriteToSheet(formResponse.getTimestamp(), data, files);

  if (VOH_EMAIL_ON_SUBMIT && files.docx) {
    try {
      MailApp.sendEmail({
        to: vohHrEmail(),
        subject: 'VOICE OF HERs — ' + (isUpdate ? 'UPDATED' : 'new') +
                 ' member profile: ' + data.name + ' (' + data.memberId + ')',
        body: data.name + (isUpdate
                ? ' has submitted the form again. This replaces the earlier document — ' +
                  'the member keeps the same ID.\n\n'
                : ' has submitted the member information form.\n\n') +
              'Member ID:  ' + data.memberId + '\n' +
              'Department: ' + data.department + (data.isBoard ? ' (Board Member)' : '') + '\n' +
              'Phone:      ' + data.phone + '\n' +
              'Email:      ' + data.email + '\n\n' +
              'The profile document is attached and saved in the profiles folder.',
        attachments: [files.docx]
      });
    } catch (err) {
      Logger.log('Email to HR failed, document still saved: ' + err.message);
    }
  }

  Logger.log('Saved: ' + files.fileUrl);
}

// ==================== ID GENERATION ====================

function vohGenerateMemberId(department, isBoard) {
  var dept = VOH_DEPARTMENT_CODES[department];
  if (!dept) return 'VOHxx000-CHECK';   // department name did not match

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var key = 'SEQ_' + dept.abbr;
    var next = parseInt(vohProps().getProperty(key) || '0', 10) + 1;
    vohProps().setProperty(key, String(next));
    var id = 'VOH' + dept.abbr + (dept.code + next);
    return isBoard ? id + '-B' : id;
  } finally {
    lock.releaseLock();
  }
}

/**
 * If members already have IDs issued by hand, put the last number used
 * for each team here and run once. Example: HR is up to VOHhr304 -> 'hr': 4
 */
function vohSetStartingNumbers() {
  var starts = {
    // 'hr': 4,
    // 'it': 2
  };
  for (var abbr in starts) vohProps().setProperty('SEQ_' + abbr, String(starts[abbr]));
  Logger.log('Starting numbers updated.');
}

// ==================== DOCUMENT BUILDER ====================

/**
 * Inserts an image, or a clickable link when the file is a PDF, HEIC, or
 * anything Google Docs cannot embed, so an upload is never silently lost.
 */
function vohInsertImageOrLink(para, fileId, label, width, fixedHeight) {
  if (!fileId) { para.setText('[no ' + label + ' submitted]'); return; }
  try {
    var file = DriveApp.getFileById(fileId);
    var blob = file.getBlob();
    var type = blob.getContentType() || '';
    var embeddable = (type.indexOf('image/') === 0 &&
                      type.indexOf('heic') < 0 && type.indexOf('heif') < 0);

    if (embeddable) {
      var img = para.appendInlineImage(blob);
      if (fixedHeight) {
        img.setWidth(width).setHeight(fixedHeight);
      } else {
        var ratio = img.getHeight() / img.getWidth();
        img.setWidth(width).setHeight(Math.round(width * ratio));
      }
      return;
    }

    Logger.log(label + ' is ' + type + ' — cannot embed, inserting a link.');
    para.setText('Open the uploaded ' + label + ' (' + file.getName() + ')');
    para.editAsText().setLinkUrl('https://drive.google.com/file/d/' + fileId + '/view');
  } catch (err) {
    Logger.log(label + ' FAILED: ' + err.message);
    para.setText('[' + label + ' could not be read]');
  }
}

function vohBuildProfileDocument(data) {
  var folder  = vohGetFolder();
  var docName = data.memberId + ' - ' + data.name;

  vohRemoveOldDocuments(folder, data.memberId);

  var doc  = DocumentApp.create(docName);
  var body = doc.getBody();
  body.setMarginTop(24).setMarginBottom(24).setMarginLeft(30).setMarginRight(30);

  // ---------- header band: logo, then the name in bold yellow ----------
  var band = body.appendTable([['']]);
  band.setBorderWidth(0);
  var bandCell = band.getCell(0, 0);
  bandCell.setBackgroundColor(VOH_PURPLE);
  bandCell.setPaddingTop(10).setPaddingBottom(8);

  var logoPara = bandCell.getChild(0).asParagraph();
  logoPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  logoPara.setSpacingAfter(2);
  try {
    var logo = logoPara.appendInlineImage(DriveApp.getFileById(vohLogoId()).getBlob());
    var ratio = logo.getHeight() / logo.getWidth();
    logo.setWidth(110).setHeight(Math.round(110 * ratio));
  } catch (err) {
    Logger.log('LOGO FAILED: ' + err.message);
  }

  var orgName = bandCell.appendParagraph('VOICE OF HERs');
  orgName.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  orgName.setSpacingBefore(0).setSpacingAfter(0);
  orgName.editAsText().setBold(true).setFontSize(15).setForegroundColor(VOH_YELLOW);

  var title = body.appendParagraph(VOH_DOC_TITLE);
  title.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  title.editAsText().setBold(true).setFontSize(22).setForegroundColor('#000000');

  vohRule(body);

  // ---------- two columns ----------
  var main = body.appendTable([['', '']]);
  main.setBorderWidth(0);
  main.setColumnWidth(0, 215);

  var left = main.getCell(0, 0);
  left.setBackgroundColor(VOH_LIGHT_PANEL);
  left.setPaddingTop(12).setPaddingBottom(12).setPaddingLeft(10).setPaddingRight(10);

  vohInsertImageOrLink(left.getChild(0).asParagraph(), data.photoId,
                       'passport photo', 105, 125);

  var emergencyPerson = data.emgName + (data.relation ? ' — ' + data.relation : '');

  var rows = [
    ['Name',                        data.name],
    ['Age',                         data.age],
    ['Phone No',                    data.phone],
    ['Email',                       data.email],
    ['Joined date',                 data.joined],
    ['Department at voice of hers', data.department + (data.isBoard ? ' — Board Member' : '')],
    ['Current Status',              data.status],
    ['ID No',                       data.memberId],
    ['Emergency Contact',           data.emgPhone],
    ['Emergency Contact Name and Relation', emergencyPerson]
  ];

  var info = left.appendTable(rows);
  info.setBorderWidth(0);
  info.setColumnWidth(0, 100);
  for (var r = 0; r < rows.length; r++) {
    var labelCell = info.getCell(r, 0);
    labelCell.setPaddingTop(5).setPaddingBottom(5).setPaddingLeft(0).setPaddingRight(4);
    labelCell.editAsText().setBold(true).setFontSize(9.5);
    var valueCell = info.getCell(r, 1);
    valueCell.setPaddingTop(5).setPaddingBottom(5);
    valueCell.editAsText().setBold(false).setFontSize(9);
  }

  // ---------- right column: work and study shown separately ----------
  var right = main.getCell(0, 1);
  right.setPaddingTop(12).setPaddingLeft(16).setPaddingRight(4);

  var blocks = [];
  if (data.placeOfWork || data.jobTitle) {
    blocks.push(['Place of Work', data.placeOfWork]);
    blocks.push(['Position',      data.jobTitle]);
  }
  if (data.placeOfStudy || data.fieldOfStudy) {
    blocks.push(['Place of Study', data.placeOfStudy]);
    blocks.push(['Field of Study', data.fieldOfStudy]);
  }
  if (!blocks.length) blocks.push(['Place of Work / Place of Study', '']);

  var firstHeading = right.getChild(0).asParagraph();
  firstHeading.setText(blocks[0][0]);
  firstHeading.editAsText().setBold(true).setFontSize(11);
  vohBodyText(right, blocks[0][1]);

  for (var b = 1; b < blocks.length; b++) {
    vohHeading(right, blocks[b][0]);
    vohBodyText(right, blocks[b][1]);
  }

  vohRule(right);

  vohHeading(right, VOH_ABOUT_HEADING);
  vohBodyText(right, data.about);

  var caption = right.appendParagraph('National Id / Kebele Id');
  caption.setSpacingBefore(10);
  caption.editAsText().setBold(true).setUnderline(true).setFontSize(10);

  vohInsertImageOrLink(right.appendParagraph(''), data.nationalId,
                       'National Id / Kebele Id', 250, null);

  var first = body.getChild(0);
  if (first.getType() === DocumentApp.ElementType.PARAGRAPH &&
      first.asParagraph().getText() === '' && body.getNumChildren() > 1) {
    body.removeChild(first);
  }

  doc.saveAndClose();

  // ---------- export as Word, then remove the Google Doc ----------
  var docId  = doc.getId();
  var result = { docx: null, fileUrl: doc.getUrl() };

  try {
    var blob = UrlFetchApp.fetch(
      'https://docs.google.com/feeds/download/documents/export/Export?id=' +
        docId + '&exportFormat=docx',
      { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } }
    ).getBlob().setName(docName + '.docx');

    var wordFile = folder.createFile(blob);
    result.docx    = blob;
    result.fileUrl = wordFile.getUrl();

    if (VOH_DELETE_GOOGLE_DOC) {
      DriveApp.getFileById(docId).setTrashed(true);
    } else {
      DriveApp.getFileById(docId).moveTo(folder);
    }
  } catch (err) {
    Logger.log('Word export failed, keeping the Google Doc: ' + err.message);
    try { DriveApp.getFileById(docId).moveTo(folder); } catch (e2) {}
  }

  return result;
}

function vohRule(container) {
  var rule = container.appendTable([['']]);
  rule.setBorderWidth(0);
  var cell = rule.getCell(0, 0);
  cell.setBackgroundColor(VOH_PURPLE);
  cell.setPaddingTop(0).setPaddingBottom(0);
  cell.getChild(0).asParagraph().editAsText().setFontSize(2);
  return rule;
}

function vohHeading(cell, text) {
  var p = cell.appendParagraph(text);
  p.setSpacingBefore(10);
  p.editAsText().setBold(true).setFontSize(11).setUnderline(false);
  return p;
}

function vohBodyText(cell, text) {
  var p = cell.appendParagraph(text || '—');
  p.editAsText().setBold(false).setFontSize(10).setUnderline(false);
  return p;
}

// ==================== SPREADSHEET ====================

/**
 * One clean set of columns at the right of the sheet. Google Forms creates
 * its own columns per question; these merge the values the document uses.
 */
var VOH_EXTRA_COLUMNS = [
  'Member ID', 'Age',
  'Work Place', 'Position', 'Study Place', 'Study Field',
  'Photo', 'ID Document', 'Profile Doc'
];

function vohEnsureColumns() {
  var sheet = vohGetSheet();
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var positions = {};
  for (var i = 0; i < VOH_EXTRA_COLUMNS.length; i++) {
    var name = VOH_EXTRA_COLUMNS[i];
    var at = headers.indexOf(name) + 1;
    if (at === 0) {
      at = sheet.getLastColumn() + 1;
      sheet.getRange(1, at).setValue(name).setFontWeight('bold');
      headers.push(name);
    }
    positions[name] = at;
  }
  return positions;
}

function vohFindRow(timestamp) {
  var sheet = vohGetSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  var stamps = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var target = timestamp.getTime();
  for (var i = stamps.length - 1; i >= 0; i--) {
    var cell = stamps[i][0];
    if (cell instanceof Date && Math.abs(cell.getTime() - target) < 300000) return i + 2;
  }
  return lastRow;
}

function vohReadCell(row, columnName) {
  var sheet = vohGetSheet();
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var col = headers.indexOf(columnName) + 1;
  if (col === 0) return null;
  var v = sheet.getRange(row, col).getValue();
  return v ? String(v) : null;
}

function vohWriteToSheet(timestamp, data, files) {
  try {
    var sheet = vohGetSheet();
    var cols  = vohEnsureColumns();
    var row   = vohFindRow(timestamp);
    if (!row) { Logger.log('Could not find the sheet row.'); return; }

    sheet.getRange(row, cols['Member ID']).setValue(data.memberId);
    sheet.getRange(row, cols['Age']).setValue(data.age);
    sheet.getRange(row, cols['Work Place']).setValue(data.placeOfWork);
    sheet.getRange(row, cols['Position']).setValue(data.jobTitle);
    sheet.getRange(row, cols['Study Place']).setValue(data.placeOfStudy);
    sheet.getRange(row, cols['Study Field']).setValue(data.fieldOfStudy);

    if (data.photoId && VOH_SHOW_PHOTO_IN_SHEET) {
      sheet.getRange(row, cols['Photo']).setFormula(
        '=IMAGE("https://drive.google.com/thumbnail?id=' + data.photoId + '&sz=w120")');
      sheet.setRowHeight(row, 110);
    } else if (data.photoId) {
      sheet.getRange(row, cols['Photo']).setFormula(
        '=HYPERLINK("https://drive.google.com/file/d/' + data.photoId + '/view","Open photo")');
    }

    if (data.nationalId) {
      sheet.getRange(row, cols['ID Document']).setFormula(
        '=HYPERLINK("https://drive.google.com/file/d/' + data.nationalId + '/view","Open ID")');
    }

    if (files.fileUrl) {
      sheet.getRange(row, cols['Profile Doc']).setFormula(
        '=HYPERLINK("' + files.fileUrl + '","Open profile")');
    }

    Logger.log('Sheet row ' + row + ' updated with ' + data.memberId);
  } catch (err) {
    Logger.log('Could not update the sheet: ' + err.message);
  }
}

// ==================== DIAGNOSTICS ====================

function vohCheckLogo() {
  try {
    var file = DriveApp.getFileById(vohLogoId());
    Logger.log('Logo found: ' + file.getName() + '  [' + file.getBlob().getContentType() + ']');
  } catch (err) {
    Logger.log('Logo could NOT be read: ' + err.message);
  }
}

function vohReportFile(label, fileId) {
  if (!fileId) { Logger.log(label + ': NOT FOUND'); return; }
  try {
    var file = DriveApp.getFileById(fileId);
    var type = file.getBlob().getContentType() || 'unknown';
    var ok = (type.indexOf('image/') === 0 &&
              type.indexOf('heic') < 0 && type.indexOf('heif') < 0);
    Logger.log(label + ': ' + file.getName() + '  [' + type + ']  ' +
               (ok ? 'can be embedded' : 'CANNOT be embedded — will become a link'));
  } catch (err) {
    Logger.log(label + ': file could not be opened — ' + err.message);
  }
}

function vohCheckUploads() {
  var all = vohGetForm().getResponses();
  if (!all.length) { Logger.log('No responses yet.'); return; }
  var f = vohFindUploads(all[all.length - 1]);
  vohReportFile('photo',  f.photoId);
  vohReportFile('id doc', f.nationalId);
}

function vohDebugLastResponse() {
  var all = vohGetForm().getResponses();
  if (!all.length) { Logger.log('No responses yet.'); return; }
  var last = all[all.length - 1];
  var items = last.getItemResponses();
  Logger.log('--- last response, ' + items.length + ' answers ---');
  for (var i = 0; i < items.length; i++) {
    var item = items[i].getItem();
    Logger.log('[' + item.getType() + ']  "' + item.getTitle() + '"  =  ' +
               JSON.stringify(items[i].getResponse()));
  }
  vohFindUploads(last);
}

/** Builds a sample document so the layout can be checked without a response. */
function vohTestProfileDoc() {
  var files = vohBuildProfileDocument({
    memberId:     'VOHhr301',
    name:         'Sample Member',
    age:          '27',
    phone:        '0900000000',
    email:        'sample@example.org',
    joined:       '2024-01-15',
    department:   'HR Team',
    isBoard:      false,
    status:       'Both working and studying',
    placeOfWork:  'Sample Organization',
    jobTitle:     'Sample Position',
    placeOfStudy: 'Sample University',
    fieldOfStudy: 'Sample Program',
    about:        'Sample record for checking the layout.',
    emgPhone:     '0900000000',
    emgName:      'Sample Contact',
    relation:     'Sister',
    photoId:      '',
    nationalId:   ''
  });
  Logger.log('Sample document: ' + files.fileUrl);
}

// ==================== CORRECTIONS ====================

/** Rebuilds the last submission's document, keeping the same ID. */
function vohRebuildLastResponse() {
  var all = vohGetForm().getResponses();
  if (!all.length) { Logger.log('No responses yet.'); return; }
  var last = all[all.length - 1];
  var row = vohFindRow(last.getTimestamp());
  var existingId = row ? vohReadCell(row, 'Member ID') : null;
  vohProcessResponse(last, existingId);
  Logger.log('Rebuilt' + (existingId ? ' as ' + existingId : ''));
}

/**
 * Reissues a member's ID after they move to a different team.
 * Set DRY_RUN to true first: it reports without changing anything.
 * Any older row for that member must be deleted from the sheet by hand.
 */
function vohReassignMember() {
  var PHONE             = '0900000000';   // the member's phone number
  var ROLLBACK_OLD_TEAM = true;           // return the unused number to the old team
  var DRY_RUN           = true;           // true = report only

  var digits = String(PHONE).replace(/\D/g, '');
  if (digits.length < 9) { Logger.log('That phone number looks wrong.'); return; }
  var key = 'MEM_P_' + digits.slice(-9);

  var props = vohProps();
  var oldId = props.getProperty(key);
  if (!oldId) { Logger.log('No index entry for that number.'); return; }

  var all = vohGetForm().getResponses();
  var mine = [];
  for (var i = 0; i < all.length; i++) {
    var items = all[i].getItemResponses();
    for (var j = 0; j < items.length; j++) {
      if (items[j].getItem().getTitle() !== 'Phone No') continue;
      var d = String(items[j].getResponse()).replace(/\D/g, '');
      if (d.slice(-9) === digits.slice(-9)) mine.push(all[i]);
    }
  }
  if (!mine.length) { Logger.log('No form response found for that number.'); return; }

  var latest = mine[mine.length - 1];
  var dept = '';
  var itemsL = latest.getItemResponses();
  for (var k = 0; k < itemsL.length; k++) {
    if (itemsL[k].getItem().getTitle() === 'Department at VOICE OF HERs') {
      dept = String(itemsL[k].getResponse());
    }
  }

  Logger.log('Current ID:        ' + oldId);
  Logger.log('Responses found:   ' + mine.length);
  Logger.log('Latest department: ' + dept);
  if (mine.length > 1) Logger.log('NOTE: delete the older row from the sheet by hand.');

  if (DRY_RUN) {
    Logger.log('DRY RUN — nothing changed. Set DRY_RUN to false to apply.');
    return;
  }

  vohRemoveOldDocuments(vohGetFolder(), oldId);

  if (ROLLBACK_OLD_TEAM) {
    var base = oldId.replace('-B', '');
    for (var name in VOH_DEPARTMENT_CODES) {
      var t = VOH_DEPARTMENT_CODES[name];
      var seq = parseInt(props.getProperty('SEQ_' + t.abbr) || '0', 10);
      if (seq > 0 && base === 'VOH' + t.abbr + (t.code + seq)) {
        props.setProperty('SEQ_' + t.abbr, String(seq - 1));
        Logger.log('Rolled ' + name + ' back to ' + (seq - 1) + '.');
      }
    }
  }

  props.deleteProperty(key);
  vohProcessResponse(latest, null);
  Logger.log('Done. The new ID is shown above.');
}
