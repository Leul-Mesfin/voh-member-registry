/**
 * ============================================================
 *  VOICE OF HERs — Member Registry
 *  FILE 1 of 2:  Forms.gs
 * ============================================================
 *
 *  Builds and maintains the member registration form.
 *
 *  FIRST TIME SETUP
 *    1. Fill in the CONFIGURATION block below
 *    2. Run createAndBrandForm
 *    3. Add the two file upload questions by hand on the last page.
 *       Apps Script cannot create them on personal Google accounts.
 *       The titles must contain these keywords:
 *         - "photo" or "passport"          for the passport photo
 *         - "national", "kebele", or "id"  for the identity document
 *    4. Run vohLockDownDuplicates in Profile.gs
 *
 *  BRANCHING
 *    Page 1 ... Are you currently working?
 *        Yes -> Work Information -> continues to Education
 *        No  -> Education
 *    Education ... Are you currently studying?
 *        Yes -> Study Information -> continues to Additional Information
 *        No  -> Additional Information
 *
 *  No setGoToPage is used anywhere. That method controls the page BEFORE
 *  the page break it is set on, which is a common off-by-one branching bug.
 * ============================================================
 */

// ==================== CONFIGURATION ====================
// Replace these two before running anything.

var VOH_LOGO_FILE_ID = 'YOUR_LOGO_DRIVE_FILE_ID';
var VOH_HR_EMAIL     = 'hr@example.org';

var VOH_FOLDER_NAME  = 'VOICE OF HERs — Member Profiles';

// ==================== FORM STRUCTURE ====================

var VOH_PAGE_WORK      = 'Work Information';
var VOH_PAGE_EDUCATION = 'Education';
var VOH_PAGE_STUDY     = 'Study Information';
var VOH_PAGE_FINAL     = 'Additional Information';

var VOH_Q_WORKING  = 'Are you currently working?';
var VOH_Q_STUDYING = 'Are you currently studying?';

var VOH_AGE_MIN = 15;
var VOH_AGE_MAX = 100;

// Must match the keys of VOH_DEPARTMENT_CODES in Profile.gs exactly.
var VOH_DEPARTMENTS = [
  'Legal Team', 'Secretary Team', 'HR Team', 'IT Team',
  'Partnership Team', 'Sponsorship Team', 'Communication Team',
  'Accounting Team', 'Event and Training Team', 'Logistics Team',
  'Graphics Team', 'Project Management Team', 'Social Media Team'
];

var VOH_RELATIONS = [
  'Mother', 'Father', 'Sister', 'Brother', 'Spouse',
  'Son / Daughter', 'Uncle / Aunt', 'Cousin', 'Friend'
];

var VOH_CONFIDENTIALITY_NOTICE =
  'CONFIDENTIALITY NOTICE\n\n' +
  'The information you provide in this form — including your phone number, ' +
  'passport photo, and National ID / Kebele ID — is strictly confidential.\n\n' +
  'It is collected only for VOICE OF HERs internal records. Access is limited ' +
  'to the HR team. Your details will not be shared with other departments, ' +
  'other members, or any person or organization outside VOICE OF HERs without ' +
  'your consent, except where we are required to do so by law.\n\n' +
  'If you have a question about your data, or you want it corrected or removed, ' +
  'contact the HR team at ' + VOH_HR_EMAIL + '.';

// ==================== BUILD THE FORM ====================

function createAndBrandForm() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('VOH_FORM_ID')) {
    Logger.log('A form is already connected. Run vohStartFresh first.');
    return;
  }

  var form = FormApp.create('VOICE OF HERs — Member Information Form');
  form.setDescription(
    'This form collects member information for VOICE OF HERs internal records.\n\n' +
    VOH_CONFIDENTIALITY_NOTICE);
  form.setCollectEmail(true);
  form.setProgressBar(true);
  form.setConfirmationMessage(
    'Thank you. Your information has been recorded and your member ID will be ' +
    'issued by the HR team.\n\n' +
    'Your details are held confidentially and are accessible only to the HR team. ' +
    'To correct or remove your information, contact ' + VOH_HR_EMAIL + '.');

  // ---------- logo ----------
  try {
    var logo = form.addImageItem()
        .setTitle('VOICE OF HERs')
        .setImage(DriveApp.getFileById(VOH_LOGO_FILE_ID).getBlob())
        .setAlignment(FormApp.Alignment.CENTER)
        .setWidth(420);
    form.moveItem(logo.getIndex(), 0);
  } catch (err) {
    Logger.log('Logo could not be added: ' + err.message);
  }

  // ================= PAGE 1 =================
  form.addSectionHeaderItem().setTitle('1. Personal Information');

  form.addTextItem().setTitle('Full Name')
      .setHelpText('As written on your ID.')
      .setRequired(true);

  var phone = form.addTextItem().setTitle('Phone No')
      .setHelpText('Example: 0912345678 or +251912345678').setRequired(true);
  phone.setValidation(FormApp.createTextValidation()
      .setHelpText('Enter a valid phone number')
      .requireTextMatchesPattern('^(\\+251|0)?[79]\\d{8}$').build());

  var email = form.addTextItem().setTitle('Email Address')
      .setHelpText('Your personal or work email address.').setRequired(true);
  email.setValidation(FormApp.createTextValidation()
      .setHelpText('Enter a valid email address').requireTextIsEmail().build());

  var age = form.addTextItem().setTitle('Age')
      .setHelpText('Your age in years, in numbers only.').setRequired(true);
  age.setValidation(FormApp.createTextValidation()
      .setHelpText('Enter a number between ' + VOH_AGE_MIN + ' and ' + VOH_AGE_MAX)
      .requireNumberBetween(VOH_AGE_MIN, VOH_AGE_MAX).build());

  form.addSectionHeaderItem().setTitle('2. Your Role in VOICE OF HERs');

  form.addListItem().setTitle('Department at VOICE OF HERs')
      .setHelpText('Choose the one team you serve in. Board members: choose your team, not "board".')
      .setChoiceValues(VOH_DEPARTMENTS).setRequired(true);

  form.addMultipleChoiceItem().setTitle('Are you also a Board Member?')
      .setHelpText('Board membership is recorded in addition to your team.')
      .setChoiceValues(['No', 'Yes']).setRequired(true);

  form.addDateItem().setTitle('Joined Date')
      .setHelpText('The date you joined. An approximate date is fine.')
      .setRequired(true);

  form.addSectionHeaderItem().setTitle('3. Work and Education');

  // must be the LAST question on page 1
  var workingQ = form.addMultipleChoiceItem()
      .setTitle(VOH_Q_WORKING)
      .setHelpText('Answer No if you are not employed at the moment.')
      .setRequired(true);

  // ================= WORK PAGE =================
  var workPage = form.addPageBreakItem()
      .setTitle(VOH_PAGE_WORK).setHelpText('About where you work.');

  form.addTextItem().setTitle('Place of Work')
      .setHelpText('Name of the company or organization.').setRequired(true);
  form.addTextItem().setTitle('Position / Job Title')
      .setHelpText('Your role, e.g. Accountant, Teacher.').setRequired(true);

  // ================= EDUCATION QUESTION PAGE =================
  var educationPage = form.addPageBreakItem().setTitle(VOH_PAGE_EDUCATION);

  // must be the ONLY question on this page
  var studyingQ = form.addMultipleChoiceItem()
      .setTitle(VOH_Q_STUDYING)
      .setHelpText('Answer No if you are not a student at the moment.')
      .setRequired(true);

  // ================= STUDY PAGE =================
  var studyPage = form.addPageBreakItem()
      .setTitle(VOH_PAGE_STUDY).setHelpText('About where you study.');

  form.addTextItem().setTitle('Place of Study')
      .setHelpText('Name of your university, college, or school.').setRequired(true);
  form.addTextItem().setTitle('Field of Study')
      .setHelpText('Your program, e.g. Computer Science, Law.').setRequired(true);

  // ================= FINAL PAGE =================
  var finalPage = form.addPageBreakItem().setTitle(VOH_PAGE_FINAL);

  form.addParagraphTextItem().setTitle('About Employee')
      .setHelpText('A short paragraph about yourself: your experience, skills, and interests.');

  form.addSectionHeaderItem().setTitle('Emergency Contact');

  var emg = form.addTextItem().setTitle('Emergency Contact')
      .setHelpText('Phone number of the person we should call in an emergency.')
      .setRequired(true);
  emg.setValidation(FormApp.createTextValidation()
      .setHelpText('Enter a valid phone number')
      .requireTextMatchesPattern('^(\\+251|0)?[79]\\d{8}$').build());

  form.addTextItem().setTitle('Emergency Contact Name')
      .setHelpText('Full name of that person.').setRequired(true);

  form.addMultipleChoiceItem().setTitle('Relation')
      .setHelpText('How is this person related to you? Choose Other to type your own.')
      .setChoiceValues(VOH_RELATIONS).showOtherOption(true).setRequired(true);

  form.addSectionHeaderItem().setTitle('Documents')
      .setHelpText('You must be signed in to a Google account to upload files.');

  // ================= BRANCHING =================
  workingQ.setChoices([
    workingQ.createChoice('Yes', workPage),
    workingQ.createChoice('No',  educationPage)
  ]);

  studyingQ.setChoices([
    studyingQ.createChoice('Yes', studyPage),
    studyingQ.createChoice('No',  finalPage)
  ]);

  // ================= SHEET AND FOLDER =================
  var sheet = SpreadsheetApp.create('VOICE OF HERs — Member Data (Responses)');
  form.setDestination(FormApp.DestinationType.SPREADSHEET, sheet.getId());
  var folder = DriveApp.createFolder(VOH_FOLDER_NAME);

  props.setProperties({
    'VOH_FORM_ID':   form.getId(),
    'VOH_SHEET_ID':  sheet.getId(),
    'VOH_FOLDER_ID': folder.getId()
  });

  setUpProfileGenerator();

  Logger.log('====================================================');
  Logger.log('EDIT THE FORM:    ' + form.getEditUrl());
  Logger.log('SHARE THIS LINK:  ' + form.getPublishedUrl());
  Logger.log('RESPONSES SHEET:  ' + sheet.getUrl());
  Logger.log('PROFILES FOLDER:  ' + folder.getUrl());
  Logger.log('====================================================');
  Logger.log('NOW ADD BY HAND on the last page: two File upload questions,');
  Logger.log('one titled with "photo", one titled with "national" or "kebele".');
}

// ==================== MAINTENANCE ====================

/** Adds a new team to the live form's dropdown. */
function vohAddDepartmentToForm() {
  var NEW_DEPARTMENT = 'Social Media Team';   // change to the team you are adding

  var form = FormApp.openById(vohRequireProp('VOH_FORM_ID'));
  var items = form.getItems(FormApp.ItemType.LIST);

  for (var i = 0; i < items.length; i++) {
    var list = items[i].asListItem();
    if (list.getTitle() !== 'Department at VOICE OF HERs') continue;

    var values = [];
    var choices = list.getChoices();
    for (var c = 0; c < choices.length; c++) {
      if (choices[c].getValue() === NEW_DEPARTMENT) {
        Logger.log('Already in the list. Nothing to do.');
        return;
      }
      values.push(choices[c].getValue());
    }

    values.push(NEW_DEPARTMENT);
    list.setChoiceValues(values);
    Logger.log('Added "' + NEW_DEPARTMENT + '". The list now has ' + values.length + ' teams.');
    Logger.log('Remember to add its code to VOH_DEPARTMENT_CODES in Profile.gs.');
    return;
  }
  Logger.log('Could not find the department question.');
}

/** Adds the Age question to a form built before Age existed. */
function vohAddAgeQuestion() {
  var form = FormApp.openById(vohRequireProp('VOH_FORM_ID'));

  var items = form.getItems();
  for (var i = 0; i < items.length; i++) {
    if (items[i].getTitle() === 'Age') {
      Logger.log('Age already exists. Nothing to do.');
      return;
    }
  }

  var age = form.addTextItem()
      .setTitle('Age')
      .setHelpText('Your age in years, in numbers only.')
      .setRequired(true);
  age.setValidation(FormApp.createTextValidation()
      .setHelpText('Enter a number between ' + VOH_AGE_MIN + ' and ' + VOH_AGE_MAX)
      .requireNumberBetween(VOH_AGE_MIN, VOH_AGE_MAX)
      .build());

  var anchor = -1;
  var all = form.getItems();
  for (var j = 0; j < all.length; j++) {
    if (all[j].getTitle() === 'Email Address') anchor = all[j].getIndex();
  }
  if (anchor < 0) {
    for (var k = 0; k < all.length; k++) {
      if (all[k].getTitle() === 'Phone No') anchor = all[k].getIndex();
    }
  }
  if (anchor >= 0) form.moveItem(age.getIndex(), anchor + 1);

  Logger.log('Age question added.');
}

// ==================== FRESH START ====================

/** Disconnects the current form and resets ID numbering to 1. */
function vohStartFresh() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) ScriptApp.deleteTrigger(triggers[i]);
  PropertiesService.getScriptProperties().deleteAllProperties();
  Logger.log('Cleared. Now run createAndBrandForm.');
  Logger.log('Nothing was deleted from Drive.');
}

/** Resets only the ID numbering, keeping the current form. */
function vohResetIdNumbers() {
  var props = PropertiesService.getScriptProperties();
  var all = props.getProperties();
  var cleared = 0;
  for (var key in all) {
    if (key.indexOf('SEQ_') === 0) { props.deleteProperty(key); cleared++; }
  }
  Logger.log('Reset ' + cleared + ' team counters.');
}

// ==================== HELPERS ====================

function vohRequireProp(key) {
  var v = PropertiesService.getScriptProperties().getProperty(key);
  if (!v) throw new Error('Nothing saved for ' + key + '. Run createAndBrandForm first.');
  return v;
}

function showMyLinks() {
  var props = PropertiesService.getScriptProperties();
  var formId = props.getProperty('VOH_FORM_ID');
  if (!formId) { Logger.log('No form connected. Run createAndBrandForm first.'); return; }

  var form = FormApp.openById(formId);
  Logger.log('EDIT THE FORM:    ' + form.getEditUrl());
  Logger.log('SHARE THIS LINK:  ' + form.getPublishedUrl());
  Logger.log('RESPONSES SHEET:  ' + SpreadsheetApp.openById(props.getProperty('VOH_SHEET_ID')).getUrl());
  Logger.log('PROFILES FOLDER:  ' + DriveApp.getFolderById(props.getProperty('VOH_FOLDER_ID')).getUrl());
}

/** Prints every page and question in order. */
function vohListFormQuestions() {
  var items = FormApp.openById(vohRequireProp('VOH_FORM_ID')).getItems();
  Logger.log('--- form contents in order ---');
  for (var i = 0; i < items.length; i++) {
    Logger.log(i + '  [' + items[i].getType() + ']  "' + items[i].getTitle() + '"');
  }
}
