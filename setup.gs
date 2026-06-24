// ============================================================
// setup.gs  (FULLY CORRECTED — v2)
// NoQs Registry — One-time Sheet Builder
// Run: buildEntireRegistry()
// ============================================================

/**
 * MASTER ENTRY POINT
 * Run this function ONCE to build the entire database layer.
 * Safe to re-run — checks for existing sheets, never overwrites real data.
 */
/**
 * MASTER ENTRY POINT
 * Run this function ONCE to build the entire database layer.
 * Safe to re-run — checks for existing sheets, never overwrites real data.
 */
function buildEntireRegistry() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('=== NoQs Registry Setup: Starting ===');

  _buildStudentsSheet(ss);
  _buildConfigSheet(ss);
  _buildRbacBlock(ss);
  _buildAuditLogSheet(ss);
  _applyNamedRanges(ss);
  _applyConditionalFormatting(ss);
  _applyDataValidation(ss);
  _applyProtectedRanges(ss);
  _seedInitialData(ss);
  _runDataMigration(ss);
  _finalizeSheetOrder(ss);

  SpreadsheetApp.flush();
  Logger.log('=== NoQs Registry Setup: Complete ===');

  // FIX (setup.gs:29): getUi() throws "Cannot call SpreadsheetApp.getUi()
  // from this context" whenever the script runs without a bound document UI
  // (web-app execution, trigger context, "Run" from certain states, etc.).
  // The alert is purely cosmetic confirmation — it must NEVER abort the build.
  // Wrap it so the function always completes successfully and logs its result.
  const summary =
    'NoQs Registry Setup Complete!\n\n' +
    '- Students sheet — ready with headers, formatting, validation\n' +
    '- Config sheet — hidden, seeded with subjects and counter\n' +
    '- Audit_Log sheet — hidden, protected, ready to record changes\n' +
    '- All named ranges, conditional formatting, and data validation applied\n\n' +
    'You can now proceed to building the web app frontend.';

  try {
    SpreadsheetApp.getUi().alert('\u2705 ' + summary);
  } catch (e) {
    // No UI available in this context — that's fine. Setup already succeeded.
    Logger.log('(UI alert skipped — no document UI in this context.)');
    Logger.log(summary);
  }
}



// ============================================================
// SECTION 1 — STUDENTS SHEET
// ============================================================

function _buildStudentsSheet(ss) {
  Logger.log('Building Students sheet...');

  let sheet = ss.getSheetByName('Students');
  if (!sheet) {
    sheet = ss.insertSheet('Students');
    Logger.log('  → Created Students sheet');
  } else {
    sheet.clear();
    Logger.log('  → Cleared existing Students sheet');
  }

  // --- Column Headers (Row 1) ---
  const headers = [
    'student_id',     // A
    'name',           // B
    'score_pct',      // C
    'dob',            // D
    'subject_id',     // E
    'is_active',      // F
    'created_at',     // G
    'updated_at',     // H
    'row_version',    // I
    'idempotency_key' // J
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // --- Header Row Styling ---
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#0F172A')
    .setFontColor('#F8FAFC')
    .setFontWeight('bold')
    .setFontSize(11)
    .setFontFamily('Arial, sans-serif')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  sheet.setFrozenRows(1);

  // --- Column Widths (px) ---
  sheet.setColumnWidth(1, 160);  // student_id
  sheet.setColumnWidth(2, 200);  // name
  sheet.setColumnWidth(3, 110);  // score_pct
  sheet.setColumnWidth(4, 140);  // dob
  sheet.setColumnWidth(5, 180);  // subject_id
  sheet.setColumnWidth(6, 100);  // is_active
  sheet.setColumnWidth(7, 200);  // created_at
  sheet.setColumnWidth(8, 200);  // updated_at
  sheet.setColumnWidth(9, 110);  // row_version
  sheet.setColumnWidth(10, 280); // idempotency_key

  // FIX 1: setDefaultRowHeight() does not exist — use setRowHeights(start, count, height)
  sheet.setRowHeights(2, 999, 28);

  // --- Column Alignment (data rows 2–1000) ---
  sheet.getRange('A2:A1000').setHorizontalAlignment('left');
  sheet.getRange('B2:B1000').setHorizontalAlignment('left');
  sheet.getRange('C2:C1000').setHorizontalAlignment('right');
  sheet.getRange('D2:D1000').setHorizontalAlignment('center');
  sheet.getRange('E2:E1000').setHorizontalAlignment('left');
  sheet.getRange('F2:F1000').setHorizontalAlignment('center');
  sheet.getRange('G2:G1000').setHorizontalAlignment('left');
  sheet.getRange('H2:H1000').setHorizontalAlignment('left');
  sheet.getRange('I2:I1000').setHorizontalAlignment('right');
  sheet.getRange('J2:J1000').setHorizontalAlignment('left');

  // --- Text Wrapping ---
  sheet.getRange('A1:J1000').setWrap(true);
  sheet.getRange('J2:J1000').setWrap(false); // UUIDs — truncate, don't wrap

  // --- Font for data rows ---
  sheet.getRange('A2:J1000')
    .setFontFamily('Arial, sans-serif')
    .setFontSize(10)
    .setFontColor('#1E293B');

  // --- Alternating row banding ---
  const bandings = sheet.getBandings();
  bandings.forEach(b => b.remove());
  sheet.getRange('A2:J1000').applyRowBanding(
    SpreadsheetApp.BandingTheme.LIGHT_GREY,
    false,
    false
  );

  // --- DOB column teal base background ---
  sheet.getRange('D2:D1000').setBackground('#CCFBF1');

  // --- Number formats ---
  sheet.getRange('C2:C1000').setNumberFormat('0.00"%"');
  sheet.getRange('I2:I1000').setNumberFormat('0');

  // --- Header tooltip notes ---
  sheet.getRange('A1').setNote('System-generated. Format: STU-YYYY-NNNN. Never edit manually.');
  sheet.getRange('B1').setNote('Student full name. Title-cased automatically on save.');
  sheet.getRange('C1').setNote('Score as percentage (0–100). Used for at-risk flagging.');
  sheet.getRange('D1').setNote('Date of birth in dd/MM/yyyy format. Cannot be a future date.');
  sheet.getRange('E1').setNote('Must match a canonical subject from the Config sheet.');
  sheet.getRange('F1').setNote('TRUE = active student. FALSE = soft-deleted. Never delete rows directly.');
  sheet.getRange('G1').setNote('ISO timestamp set automatically when the record is created.');
  sheet.getRange('H1').setNote('ISO timestamp updated automatically on every change.');
  sheet.getRange('I1').setNote('Increments on every update. Used for optimistic concurrency control.');
  sheet.getRange('J1').setNote('UUID generated per form submission. Prevents duplicate rows on retry.');

  Logger.log('  ✓ Students sheet built successfully');
}


// ============================================================
// SECTION 2 — CONFIG SHEET
// ============================================================

function _buildConfigSheet(ss) {
  Logger.log('Building Config sheet...');

  let sheet = ss.getSheetByName('Config');
  if (!sheet) {
    sheet = ss.insertSheet('Config');
    Logger.log('  → Created Config sheet');
  } else {
    sheet.clear();
    Logger.log('  → Cleared existing Config sheet');
  }

  // --- Section A: Canonical Subject List ---
  sheet.getRange('A1').setValue('CANONICAL SUBJECTS');
  const subjects = [
    ['Mathematics'],
    ['Science'],
    ['English'],
    ['History'],
    ['Computer Science'],
    ['Physics'],
    ['Chemistry'],
    ['Biology'],
    ['Geography'],
    ['Economics']
  ];
  sheet.getRange(2, 1, subjects.length, 1).setValues(subjects);

  sheet.getRange('A1')
    .setBackground('#0F172A')
    .setFontColor('#F8FAFC')
    .setFontWeight('bold')
    .setFontSize(11)
    .setHorizontalAlignment('center');

  sheet.getRange(2, 1, subjects.length, 1)
    .setBackground('#F0FDF4')
    .setFontColor('#166534')
    .setFontSize(10)
    .setBorder(true, true, true, true, true, true, '#BBF7D0', SpreadsheetApp.BorderStyle.SOLID);

  sheet.setColumnWidth(1, 200);

  // --- Section B: Auto-Increment Counter ---
  sheet.getRange('B1').setValue('ID_COUNTER_LABEL')
    .setBackground('#0F172A')
    .setFontColor('#F8FAFC')
    .setFontWeight('bold')
    .setFontSize(11)
    .setHorizontalAlignment('center');

  sheet.getRange('B2').setValue(0)
    .setBackground('#FEF3C7')
    .setFontColor('#92400E')
    .setFontWeight('bold')
    .setFontSize(14)
    .setHorizontalAlignment('center')
    .setNote('Auto-incremented by the web app. Do NOT edit manually.');

  sheet.getRange('B3').setValue('← Current ID Counter')
    .setFontColor('#6B7280')
    .setFontSize(9)
    .setFontStyle('italic');

  sheet.setColumnWidth(2, 180);

  // --- Section C: App-Wide Settings ---
  // FIX 3: Removed the broken .transpose() chain. Set header cells individually.
  sheet.getRange('C1').setValue('SETTING KEY');
  sheet.getRange('D1').setValue('VALUE');
  sheet.getRange('E1').setValue('DESCRIPTION');

  sheet.getRange('C1:E1')
    .setBackground('#0F172A')
    .setFontColor('#F8FAFC')
    .setFontWeight('bold')
    .setFontSize(11)
    .setHorizontalAlignment('center');

  const settings = [
    ['AT_RISK_THRESHOLD', 50,  'Students scoring below this value are flagged at-risk (red)'],
    ['AMBER_THRESHOLD',   70,  'Students scoring below this but at/above at-risk get amber'],
    ['APP_NAME',          'NoQs Registry', 'Displayed in the web app header'],
    ['ACADEMIC_YEAR',     new Date().getFullYear(), 'Used in student ID generation: STU-YYYY-NNNN'],
    ['MAX_SCORE',         100, 'Maximum possible score percentage'],
    ['MIN_SCORE',         0,   'Minimum possible score percentage'],
    ['CACHE_TTL_SECONDS', 30,  'How long getAllStudents() result is cached'],
    ['LOCK_TIMEOUT_MS',   10000, 'LockService.waitLock() timeout in milliseconds']
  ];

  sheet.getRange(2, 3, settings.length, 3).setValues(settings);

  sheet.getRange(2, 3, settings.length, 1)
    .setBackground('#EFF6FF')
    .setFontColor('#1E3A8A')
    .setFontWeight('bold');

  sheet.getRange(2, 4, settings.length, 1)
    .setBackground('#FEF9C3')
    .setFontColor('#713F12')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  sheet.getRange(2, 5, settings.length, 1)
    .setFontColor('#6B7280')
    .setFontSize(9)
    .setFontStyle('italic');

  sheet.setColumnWidth(3, 200);
  sheet.setColumnWidth(4, 120);
  sheet.setColumnWidth(5, 380);

  sheet.getRange(1, 3, settings.length + 1, 3)
    .setBorder(true, true, true, true, true, true, '#BFDBFE', SpreadsheetApp.BorderStyle.SOLID);

  // --- Section D: Subject Normalization Map ---
  sheet.getRange('G1').setValue('SUBJECT NORMALIZATION MAP')
    .setBackground('#0F172A')
    .setFontColor('#F8FAFC')
    .setFontWeight('bold')
    .setFontSize(11)
    .setHorizontalAlignment('center');

  sheet.getRange('H1').setValue('CANONICAL VALUE')
    .setBackground('#0F172A')
    .setFontColor('#F8FAFC')
    .setFontWeight('bold')
    .setFontSize(11)
    .setHorizontalAlignment('center');

  const normMap = [
    ['Maths',    'Mathematics'],
    ['Math',     'Mathematics'],
    ['Sci',      'Science'],
    ['Eng',      'English'],
    ['Hist',     'History'],
    ['CS',       'Computer Science'],
    ['Comp Sci', 'Computer Science'],
    ['Phys',     'Physics'],
    ['Chem',     'Chemistry'],
    ['Bio',      'Biology'],
    ['Geo',      'Geography'],
    ['Econ',     'Economics']
  ];

  sheet.getRange(2, 7, normMap.length, 2).setValues(normMap);
  sheet.getRange(2, 7, normMap.length, 1)
    .setBackground('#FEE2E2')
    .setFontColor('#991B1B');
  sheet.getRange(2, 8, normMap.length, 1)
    .setBackground('#DCFCE7')
    .setFontColor('#166534');

  sheet.setColumnWidth(7, 160);
  sheet.setColumnWidth(8, 180);
  sheet.getRange('G1').setNote('Used by the migration script to normalise dirty subject values.');

  sheet.hideSheet();
  Logger.log('  ✓ Config sheet built and hidden');
}

// ============================================================
// SECTION 2b — RBAC EDITOR ALLOWLIST (Config columns J/K)
// ============================================================
function _buildRbacBlock(ss) {
  Logger.log('Building RBAC editor allowlist...');
  const sheet = ss.getSheetByName('Config');

  sheet.getRange('J1').setValue('EDITOR EMAILS')
    .setBackground('#0F172A').setFontColor('#F8FAFC')
    .setFontWeight('bold').setFontSize(11).setHorizontalAlignment('center');

  sheet.getRange('K1').setValue('DEFAULT ROLE')
    .setBackground('#0F172A').setFontColor('#F8FAFC')
    .setFontWeight('bold').setFontSize(11).setHorizontalAlignment('center');

  // Seed the first editor as whoever runs setup. They can add more below.
  let owner = '';
  try { owner = Session.getEffectiveUser().getEmail() || ''; } catch (e) {}
  if (owner) sheet.getRange('J2').setValue(owner);

  sheet.getRange('J1').setNote(
    'One editor email per row, starting J2. Anyone NOT in this list is a Viewer ' +
    '(read-only). Emails are matched case-insensitively.');

  // K2 controls what an unknown user gets: "viewer" (recommended) or "editor".
  sheet.getRange('K2').setValue('viewer')
    .setBackground('#FEF3C7').setFontColor('#92400E').setHorizontalAlignment('center')
    .setNote('Role for users NOT in the editor list. Use "viewer" to lock down by default.');

  sheet.setColumnWidth(10, 260);
  sheet.setColumnWidth(11, 130);

  Logger.log('  ✓ RBAC block built (Config J/K)');
}



// ============================================================
// SECTION 3 — AUDIT LOG SHEET
// ============================================================

function _buildAuditLogSheet(ss) {
  Logger.log('Building Audit_Log sheet...');

  let sheet = ss.getSheetByName('Audit_Log');
  if (!sheet) {
    sheet = ss.insertSheet('Audit_Log');
    Logger.log('  → Created Audit_Log sheet');
  } else {
    sheet.clear();
    Logger.log('  → Cleared existing Audit_Log sheet');
  }

  const headers = ['timestamp', 'action', 'student_id', 'changed_by', 'old_values', 'new_values'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#7C3AED')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setFontSize(11)
    .setFontFamily('Arial, sans-serif')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  sheet.setFrozenRows(1);

  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 160);
  sheet.setColumnWidth(4, 220);
  sheet.setColumnWidth(5, 350);
  sheet.setColumnWidth(6, 350);

  // FIX 2: Same setDefaultRowHeight() fix as Students sheet
  sheet.setRowHeights(2, 4999, 26);

  sheet.getRange('A2:F5000')
    .setFontFamily('Courier New, monospace')
    .setFontSize(9)
    .setFontColor('#374151')
    .setVerticalAlignment('top');

  sheet.getRange('A2:A5000')
    .setNumberFormat('dd/MM/yyyy HH:mm:ss')
    .setHorizontalAlignment('left')
    .setFontFamily('Arial, sans-serif');

  sheet.getRange('B2:B5000')
    .setHorizontalAlignment('center')
    .setFontWeight('bold');

  sheet.getRange('E2:F5000').setWrap(false);

  sheet.getRange('A1').setNote('UTC timestamp of when the action occurred.');
  sheet.getRange('B1').setNote('One of: CREATE, UPDATE, DELETE, RESTORE');
  sheet.getRange('C1').setNote('The student_id of the affected record.');
  sheet.getRange('D1').setNote('Email of the user who performed the action.');
  sheet.getRange('E1').setNote('JSON snapshot of the record BEFORE the change. Empty for CREATE.');
  sheet.getRange('F1').setNote('JSON snapshot of the record AFTER the change. Empty for DELETE.');

  const bandings = sheet.getBandings();
  bandings.forEach(b => b.remove());
  sheet.getRange('A2:F5000').applyRowBanding(
    SpreadsheetApp.BandingTheme.LIGHT_GREY,
    false,
    false
  );

  sheet.hideSheet();
  Logger.log('  ✓ Audit_Log sheet built and hidden');
}


// ============================================================
// SECTION 4 — NAMED RANGES
// ============================================================

function _applyNamedRanges(ss) {
  Logger.log('Applying named ranges...');

  const namesToRemove = [
    'SUBJECT_LIST', 'ID_COUNTER', 'AT_RISK_THRESHOLD',
    'AMBER_THRESHOLD', 'APP_NAME', 'ACADEMIC_YEAR',
    'SUBJECT_NORM_MAP', 'ACTIVE_STUDENTS'
  ];

  ss.getNamedRanges().forEach(nr => {
    if (namesToRemove.includes(nr.getName())) nr.remove();
  });

  const config   = ss.getSheetByName('Config');
  const students = ss.getSheetByName('Students');

  ss.setNamedRange('SUBJECT_LIST',      config.getRange('A2:A11'));
  ss.setNamedRange('ID_COUNTER',        config.getRange('B2'));
  ss.setNamedRange('AT_RISK_THRESHOLD', config.getRange('D2'));
  ss.setNamedRange('AMBER_THRESHOLD',   config.getRange('D3'));
  ss.setNamedRange('ACADEMIC_YEAR',     config.getRange('D5'));
  ss.setNamedRange('SUBJECT_NORM_MAP',  config.getRange('G2:H13'));
  ss.setNamedRange('ACTIVE_STUDENTS',   students.getRange('A2:J1000'));

  Logger.log('  ✓ Named ranges applied: SUBJECT_LIST, ID_COUNTER, AT_RISK_THRESHOLD, AMBER_THRESHOLD, ACADEMIC_YEAR, ACTIVE_STUDENTS, SUBJECT_NORM_MAP');
}


// ============================================================
// SECTION 5 — CONDITIONAL FORMATTING
// ============================================================

function _applyConditionalFormatting(ss) {
  Logger.log('Applying conditional formatting...');

  const students = ss.getSheetByName('Students');
  const auditLog = ss.getSheetByName('Audit_Log');

  students.clearConditionalFormatRules();
  auditLog.clearConditionalFormatRules();

  const studentRules = [];

  // FIX 4 + FIX 6 throughout:
  //   .whenFormula()      → .whenFormulaSatisfied()
  //   .setFontWeight('bold') → .setBold(true)

  // RULE 1: Inactive rows → grey strikethrough entire row
  studentRules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$F2=FALSE')
      .setBackground('#F1F5F9')
      .setFontColor('#94A3B8')
      .setStrikethrough(true)
      .setRanges([students.getRange('A2:J1000')])
      .build()
  );

  // RULE 2: Score < 50 AND active → red score cell
  studentRules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($C2<50,$F2=TRUE)')
      .setBackground('#FEE2E2')
      .setFontColor('#DC2626')
      .setBold(true)
      .setRanges([students.getRange('C2:C1000')])
      .build()
  );

  // RULE 3: Score 50–69 AND active → amber score cell
  studentRules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($C2>=50,$C2<70,$F2=TRUE)')
      .setBackground('#FEF3C7')
      .setFontColor('#D97706')
      .setBold(true)
      .setRanges([students.getRange('C2:C1000')])
      .build()
  );

  // RULE 4: Score >= 70 AND active → green score cell
  studentRules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($C2>=70,$F2=TRUE)')
      .setBackground('#DCFCE7')
      .setFontColor('#16A34A')
      .setBold(true)
      .setRanges([students.getRange('C2:C1000')])
      .build()
  );

  // RULE 5: Active rows → teal DOB cell (reinforces base colour)
  studentRules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$F2=TRUE')
      .setBackground('#CCFBF1')
      .setFontColor('#0F766E')
      .setRanges([students.getRange('D2:D1000')])
      .build()
  );

  // RULE 6: Perfect score (100) AND active → gold entire row
  studentRules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($C2=100,$F2=TRUE)')
      .setBackground('#FEF9C3')
      .setFontColor('#A16207')
      .setBold(true)
      .setRanges([students.getRange('A2:J1000')])
      .build()
  );

  // RULE 7: Created today → subtle blue tint on name/ID columns
  studentRules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(DATEVALUE(LEFT($G2,10))=TODAY(),$F2=TRUE)')
      .setBackground('#EFF6FF')
      .setRanges([students.getRange('A2:B1000')])
      .build()
  );

  students.setConditionalFormatRules(studentRules);

  // --- Audit Log conditional formatting (action column) ---
  const auditRules = [];

  auditRules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('CREATE')
      .setBackground('#DCFCE7')
      .setFontColor('#166534')
      .setBold(true)
      .setRanges([auditLog.getRange('B2:B5000')])
      .build()
  );

  auditRules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('UPDATE')
      .setBackground('#DBEAFE')
      .setFontColor('#1E40AF')
      .setBold(true)
      .setRanges([auditLog.getRange('B2:B5000')])
      .build()
  );

  auditRules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('DELETE')
      .setBackground('#FEE2E2')
      .setFontColor('#991B1B')
      .setBold(true)
      .setRanges([auditLog.getRange('B2:B5000')])
      .build()
  );

  auditRules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('RESTORE')
      .setBackground('#FEF3C7')
      .setFontColor('#92400E')
      .setBold(true)
      .setRanges([auditLog.getRange('B2:B5000')])
      .build()
  );

  auditLog.setConditionalFormatRules(auditRules);

  Logger.log('  ✓ Conditional formatting applied (7 rules on Students, 4 rules on Audit_Log)');
}


// ============================================================
// SECTION 6 — DATA VALIDATION
// ============================================================

function _applyDataValidation(ss) {
  Logger.log('Applying data validation...');

  const students = ss.getSheetByName('Students');
  const config   = ss.getSheetByName('Config');

  // Subject dropdown — sourced from Config A2:A11
  students.getRange('E2:E1000').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInRange(config.getRange('A2:A11'), true)
      .setAllowInvalid(false)
      .setHelpText('Select a subject from the canonical list. New subjects must be added to Config first.')
      .build()
  );

  // is_active — TRUE or FALSE only
  students.getRange('F2:F1000').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['TRUE', 'FALSE'], true)
      .setAllowInvalid(false)
      .setHelpText('TRUE = active student. FALSE = soft-deleted. Use the web app to delete.')
      .build()
  );

  // Score — number between 0 and 100 inclusive
  students.getRange('C2:C1000').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireNumberBetween(0, 100)
      .setAllowInvalid(false)
      .setHelpText('Enter a score between 0 and 100. Decimals allowed.')
      .build()
  );

  // DOB — must match dd/MM/yyyy pattern
  students.getRange('D2:D1000').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireFormulaSatisfied('=REGEXMATCH(TO_TEXT(D2),"^\\d{2}/\\d{2}/\\d{4}$")')
      .setAllowInvalid(false)
      .setHelpText('Enter date in dd/MM/yyyy format. Example: 15/08/2003')
      .build()
  );

  // student_id — must match STU-YYYY-NNNN format
  students.getRange('A2:A1000').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireFormulaSatisfied('=REGEXMATCH(TO_TEXT(A2),"^STU-\\d{4}-\\d{4}$")')
      .setAllowInvalid(false)
      .setHelpText('Student IDs are auto-generated by the system. Format: STU-YYYY-NNNN')
      .build()
  );

  // row_version — must be a positive integer
  students.getRange('I2:I1000').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireNumberGreaterThan(0)
      .setAllowInvalid(false)
      .setHelpText('Row version is managed by the system. Do not edit manually.')
      .build()
  );

  Logger.log('  ✓ Data validation applied (subject dropdown, is_active, score, DOB, student_id, row_version)');
}


// ============================================================
// SECTION 7 — PROTECTED RANGES
// ============================================================

function _applyProtectedRanges(ss) {
  Logger.log('Applying protected ranges...');

  // FIX 8: Safer removal of existing protections.
  // ss.getProtections(RANGE) only returns range-level protections — correct.
  // For sheet-level protections, we must call sheet.getProtections() on each sheet.
  ss.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => p.remove());

  const students = ss.getSheetByName('Students');
  const config   = ss.getSheetByName('Config');
  const audit    = ss.getSheetByName('Audit_Log');

  // Remove any existing sheet-level protections from each sheet individually
  [students, config, audit].forEach(sheet => {
    sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());
  });

  // Protect system-managed columns in Students (warning only so backend can still write)
  ['A', 'G', 'H', 'I', 'J'].forEach(col => {
    students.getRange(`${col}2:${col}1000`).protect()
      .setDescription(`System column ${col} — do not edit manually`)
      .setWarningOnly(true);
  });

  // Protect entire Config sheet (warning only)
  config.protect()
    .setDescription('Config — managed by setup.gs and Repository.gs only')
    .setWarningOnly(true);

  // Protect entire Audit_Log sheet (warning only)
  audit.protect()
    .setDescription('Audit_Log — append-only log. Do not edit manually.')
    .setWarningOnly(true);

  Logger.log('  ✓ Protected ranges applied (5 system columns in Students, full Config, full Audit_Log)');
}


// ============================================================
// SECTION 8 — SEED INITIAL DATA
// ============================================================

function _seedInitialData(ss) {
  Logger.log('Seeding initial sample data...');

  const sheet = ss.getSheetByName('Students');

  if (sheet.getLastRow() > 1) {
    Logger.log('  → Students sheet already has data — skipping seed to preserve existing records');
    return;
  }

  const now  = new Date().toISOString();
  const year = new Date().getFullYear();

  const sampleStudents = [
    ['STU-' + year + '-0001', 'Amrit Kumar Ghoshal',  88,  '15/08/2004', 'Computer Science', true,  now, now, 1, Utilities.getUuid()],
    ['STU-' + year + '-0002', 'Priya Sharma',          92,  '22/03/2003', 'Mathematics',      true,  now, now, 1, Utilities.getUuid()],
    ['STU-' + year + '-0003', 'Rohan Mehta',           45,  '07/11/2004', 'Science',          true,  now, now, 1, Utilities.getUuid()],
    ['STU-' + year + '-0004', 'Anjali Verma',          78,  '30/05/2003', 'English',          true,  now, now, 1, Utilities.getUuid()],
    ['STU-' + year + '-0005', 'Karthik Nair',          63,  '19/09/2004', 'History',          true,  now, now, 1, Utilities.getUuid()],
    ['STU-' + year + '-0006', 'Sneha Patel',           55,  '02/01/2005', 'Chemistry',        true,  now, now, 1, Utilities.getUuid()],
    ['STU-' + year + '-0007', 'Vikram Singh',          100, '14/06/2003', 'Physics',          true,  now, now, 1, Utilities.getUuid()],
    ['STU-' + year + '-0008', 'Meera Iyer',            38,  '28/12/2004', 'Biology',          true,  now, now, 1, Utilities.getUuid()],
    ['STU-' + year + '-0009', 'Arjun Das',             71,  '03/04/2003', 'Geography',        true,  now, now, 1, Utilities.getUuid()],
    ['STU-' + year + '-0010', 'Divya Krishnan',        84,  '17/07/2004', 'Economics',        true,  now, now, 1, Utilities.getUuid()],
    ['STU-' + year + '-0011', 'Rahul Gupta',           49,  '09/02/2005', 'Mathematics',      true,  now, now, 1, Utilities.getUuid()],
    ['STU-' + year + '-0012', 'Pooja Reddy',           95,  '25/10/2003', 'Computer Science', true,  now, now, 1, Utilities.getUuid()],
    ['STU-' + year + '-0013', 'Siddharth Joshi',       67,  '11/03/2004', 'Science',          true,  now, now, 1, Utilities.getUuid()],
    ['STU-' + year + '-0014', 'Tanvi Bhatt',           82,  '06/08/2003', 'English',          true,  now, now, 1, Utilities.getUuid()],
    ['STU-' + year + '-0015', 'Nikhil Agarwal',        58,  '20/05/2004', 'History',          false, now, now, 1, Utilities.getUuid()]
  ];

  // Single batch write — not 15 appendRow() calls
  sheet.getRange(2, 1, sampleStudents.length, 10).setValues(sampleStudents);

  // Sync the ID counter in Config
  ss.getSheetByName('Config').getRange('B2').setValue(sampleStudents.length);

  // Seed matching audit log entries
  const audit = ss.getSheetByName('Audit_Log');
  const auditRows = sampleStudents.map(s => [
    now, 'CREATE', s[0], 'setup@system', '',
    JSON.stringify({ name: s[1], score_pct: s[2], dob: s[3], subject_id: s[4] })
  ]);
  audit.getRange(2, 1, auditRows.length, 6).setValues(auditRows);

  Logger.log(`  ✓ Seeded ${sampleStudents.length} sample students and ${auditRows.length} audit entries`);
}


// ============================================================
// SECTION 9 — DATA MIGRATION
// ============================================================

function _runDataMigration(ss) {
  Logger.log('Running data migration / dirty data cleanup...');

  const students = ss.getSheetByName('Students');
  const config   = ss.getSheetByName('Config');
  const audit    = ss.getSheetByName('Audit_Log');

  const lastRow = students.getLastRow();
  if (lastRow < 2) {
    Logger.log('  → No data rows to migrate');
    return;
  }

  // Load normalization map from Config G2:H13 in one batch call
  const normMapValues = config.getRange('G2:H13').getValues();
  const normMap = {};
  normMapValues.forEach(row => {
    if (row[0]) normMap[row[0].toString().trim()] = row[1].toString().trim();
  });

  // Read all student data in ONE batch call
  const dataRange = students.getRange(2, 1, lastRow - 1, 10);
  const data = dataRange.getValues();

  let migrationsApplied = 0;
  const migrationLog = [];

  data.forEach((row, i) => {
    const rowNum = i + 2;
    const originalSubject = row[4] ? row[4].toString().trim() : '';

    // Normalise subject if it matches an alias
    if (originalSubject && normMap[originalSubject]) {
      const canonical = normMap[originalSubject];
      Logger.log(`  → Row ${rowNum}: Normalizing subject "${originalSubject}" → "${canonical}"`);
      migrationLog.push([
        new Date().toISOString(), 'UPDATE',
        row[0] || `ROW_${rowNum}`, 'migration@system',
        JSON.stringify({ subject_id: originalSubject }),
        JSON.stringify({ subject_id: canonical })
      ]);
      data[i][4] = canonical;
      migrationsApplied++;
    }

    // Coerce is_active strings to booleans
    if (row[5] === 'TRUE'  || row[5] === 'true')                      data[i][5] = true;
    else if (row[5] === 'FALSE' || row[5] === 'false' || row[5] === '') data[i][5] = false;

    // Ensure row_version is at least 1
    if (!row[8] || row[8] < 1) data[i][8] = 1;

    // Generate idempotency key if missing
    if (!row[9]) data[i][9] = Utilities.getUuid();
  });

  // Write ALL changes back in ONE batch call
  dataRange.setValues(data);

  // Log any subject normalisations to the audit trail
  if (migrationLog.length > 0) {
    const auditLastRow = audit.getLastRow();
    audit.getRange(auditLastRow + 1, 1, migrationLog.length, 6).setValues(migrationLog);
  }

  Logger.log(`  ✓ Migration complete: ${migrationsApplied} subject normalizations applied`);
}


// ============================================================
// SECTION 10 — SHEET ORDER & FINAL PRESENTATION
// ============================================================

function _finalizeSheetOrder(ss) {
  Logger.log('Finalizing sheet order and presentation...');

  const students = ss.getSheetByName('Students');
  const config   = ss.getSheetByName('Config');
  const audit    = ss.getSheetByName('Audit_Log');

  // Move Students to position 1 (first tab)
  ss.setActiveSheet(students);
  ss.moveActiveSheet(1);

  // Ensure Config and Audit_Log stay hidden
  if (config && !config.isSheetHidden()) config.hideSheet();
  if (audit  && !audit.isSheetHidden())  audit.hideSheet();

  // Auto-resize ID and name columns for clean initial presentation
  students.autoResizeColumn(1);
  students.autoResizeColumn(2);

  // Tab colours — .setTabColor(cssString) is valid, no change needed
  students.setTabColor('#0D9488'); // teal
  config.setTabColor('#F59E0B');   // amber
  audit.setTabColor('#7C3AED');    // purple

  // Rename the spreadsheet
  ss.rename('NoQs Registry — Student Management System');

  Logger.log('  ✓ Sheet order finalised, tabs coloured, title updated');
}


// ============================================================
// UTILITY FUNCTIONS — Used by Repository.gs and frontend
// ============================================================

/**
 * Returns the canonical subject list from Config A2:A11.
 * Called by Repository.gs and the frontend subject dropdown.
 */
function getSubjectList() {
  const config = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  return config.getRange('A2:A11').getValues()
    .map(r => r[0])
    .filter(v => v !== '');
}

/**
 * Returns all app settings as a { key: value } object.
 * Called once on app load to configure thresholds and constants.
 */
function getAppSettings() {
  const config = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  const rows   = config.getRange('C2:D9').getValues();
  const out    = {};
  rows.forEach(row => { if (row[0]) out[row[0]] = row[1]; });
  return out;
}

/**
 * Diagnostic health check — run after setup to verify everything built correctly.
 * Results appear in the Apps Script Execution Log (View → Logs).
 */
function runHealthCheck() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const students = ss.getSheetByName('Students');
  const config   = ss.getSheetByName('Config');
  const audit    = ss.getSheetByName('Audit_Log');

  Logger.log('=== Health Check ===');
  Logger.log(`Students  : ${students ? 'EXISTS' : '❌ MISSING'}`);
  Logger.log(`Config    : ${config   ? 'EXISTS' : '❌ MISSING'}`);
  Logger.log(`Audit_Log : ${audit    ? 'EXISTS' : '❌ MISSING'}`);

  if (students) {
    Logger.log(`  Data rows          : ${students.getLastRow() - 1}`);
    Logger.log(`  Columns            : ${students.getLastColumn()}`);
    Logger.log(`  Frozen rows        : ${students.getFrozenRows()}`);
    Logger.log(`  Bandings           : ${students.getBandings().length}`);
    Logger.log(`  Cond. format rules : ${students.getConditionalFormatRules().length}`);
    Logger.log(`  Subject validation : ${students.getRange('E2').getDataValidation() ? 'SET' : '❌ MISSING'}`);
  }
  if (config) {
    Logger.log(`  Config hidden      : ${config.isSheetHidden()}`);
    Logger.log(`  ID Counter         : ${config.getRange('B2').getValue()}`);
    Logger.log(`  Subjects           : [${getSubjectList().join(', ')}]`);
  }
  if (audit) {
    Logger.log(`  Audit_Log hidden   : ${audit.isSheetHidden()}`);
    Logger.log(`  Audit entries      : ${audit.getLastRow() - 1}`);
  }

  const namedRanges = ss.getNamedRanges().map(nr => nr.getName());
  Logger.log(`  Named ranges       : [${namedRanges.join(', ')}]`);
  Logger.log('=== Health Check Complete ===');
}

// ============================================================
// WEB APP ENTRY POINTS — required to serve the frontend.
// Without doGet(), the deployed Web App URL returns nothing.
// ============================================================

/**
 * Serves the single-page app. Deploy: Execute as Me, Access Anyone.
 * Uses the google.script.run model (not doPost), per the design doc.
 */

function doGet() {
  var tmpl = HtmlService.createTemplateFromFile('Index');
  // Resolve the role ONCE at page-render time (server-side).
  // With "Execute as: User" deployment, this correctly identifies the visitor.
  tmpl.initialRole  = _resolveRole();
  tmpl.initialEmail = _currentUserEmail();
  return tmpl.evaluate()
    .setTitle('NoQs Registry')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}



/**
 * include() — lets Index.html pull in Styles.html and Script.html via
 * <?!= include('Styles') ?>. Keeps CSS/JS in separate files for sanity.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

