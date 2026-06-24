// ============================================================
// Repository.gs  — DATA ACCESS LAYER
// ------------------------------------------------------------
// THE ONLY FILE PERMITTED TO CALL SpreadsheetApp.
// Every other layer (Service.gs, frontend) goes through these
// functions. Swap the database -> change only this file.
//
// Responsibilities:
//   • Read/write the Students sheet (batch operations only)
//   • LockService.getScriptLock() around EVERY mutation
//   • CacheService caching for repo_getAllStudents()
//   • Atomic student-ID generation (counter in Config)
//   • Append-only Audit_Log writes
//
// IMPORTANT: Repository functions THROW on hard failure.
// Service.gs is responsible for catching and converting those
// throws into the { success, error, code } client contract.
// Repository never builds the client response shape itself —
// it returns raw data objects or throws.
// ============================================================


// ------------------------------------------------------------
// CONSTANTS
// ------------------------------------------------------------
const SHEET_STUDENTS = 'Students';
const SHEET_CONFIG   = 'Config';
const SHEET_AUDIT    = 'Audit_Log';

// Column order MUST match the Students sheet headers exactly (A..J).
// Centralized here so a schema change is a one-line edit.
const STUDENT_COLS = [
  'student_id',      // 0  A
  'name',            // 1  B
  'score_pct',       // 2  C
  'dob',             // 3  D
  'subject_id',      // 4  E
  'is_active',       // 5  F
  'created_at',      // 6  G
  'updated_at',      // 7  H
  'row_version',     // 8  I
  'idempotency_key'  // 9  J
];

const CACHE_KEY_STUDENTS = 'ALL_STUDENTS_V1';
const CACHE_TTL_SECONDS  = 30;     // matches the design doc
const LOCK_TIMEOUT_MS    = 10000;  // wait up to 10s for the script lock
const FIRST_DATA_ROW     = 2;      // row 1 is headers


// ============================================================
// PRIVATE SHEET HELPERS  (all SpreadsheetApp access funnels here)
// ============================================================

function _ss() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function _studentsSheet() {
  const sh = _ss().getSheetByName(SHEET_STUDENTS);
  if (!sh) throw new Error('Students sheet not found. Run buildEntireRegistry() first.');
  return sh;
}

function _auditSheet() {
  const sh = _ss().getSheetByName(SHEET_AUDIT);
  if (!sh) throw new Error('Audit_Log sheet not found.');
  return sh;
}

function _configSheet() {
  const sh = _ss().getSheetByName(SHEET_CONFIG);
  if (!sh) throw new Error('Config sheet not found.');
  return sh;
}

/**
 * Convert a raw sheet row (array) into a typed student OBJECT.
 * This is the BACKEND DATA BOUNDARY — the frontend receives exactly
 * this shape (keys == STUDENT_COLS). Booleans/numbers are coerced here
 * so the client never has to second-guess types.
 */
function _rowToObject(row) {
  const obj = {};
  STUDENT_COLS.forEach((key, i) => { obj[key] = row[i]; });
  // Type coercion at the boundary
  obj.score_pct   = Number(obj.score_pct);
  obj.row_version = Number(obj.row_version);
  obj.is_active   = (obj.is_active === true || obj.is_active === 'TRUE' || obj.is_active === 'true');
  return obj;
}

/**
 * Convert a student OBJECT back into a sheet ROW (array) in column order.
 */
function _objectToRow(obj) {
  return STUDENT_COLS.map(key => obj[key]);
}


// ============================================================
// READ — repo_getAllStudents()  (CACHED)
// ============================================================

/**
 * Returns ALL student objects (active AND inactive — the frontend filters).
 *
 * CACHING: results are held in CacheService for CACHE_TTL_SECONDS. Repeated
 * page loads / filter operations within that window do NOT re-hit the sheet.
 * Any mutation (create/update/delete/restore) calls _invalidateCache() so the
 * next read is fresh.
 *
 * BATCH READ: the entire data range is pulled in ONE getValues() call —
 * orders of magnitude faster than per-cell reads.
 *
 * @returns {Object[]} array of student objects (THROWS on sheet failure)
 */
function repo_getAllStudents() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(CACHE_KEY_STUDENTS);
  if (cached) {
    // Cache stores a JSON string; parse back to objects.
    return JSON.parse(cached);
  }

  const sh = _studentsSheet();
  const lastRow = sh.getLastRow();

  // No data rows yet -> empty array (still cache the empty result briefly)
  let students = [];
  if (lastRow >= FIRST_DATA_ROW) {
    const values = sh.getRange(FIRST_DATA_ROW, 1, lastRow - 1, STUDENT_COLS.length).getValues();
    students = values
      .filter(r => r[0] !== '' && r[0] != null) // skip blank rows
      .map(_rowToObject);
  }

  // CacheService values must be strings and < 100KB. JSON is fine here.
  // For very large datasets this could exceed the limit; chunking would be
  // the next step, but is out of scope for the current dataset size.
  cache.put(CACHE_KEY_STUDENTS, JSON.stringify(students), CACHE_TTL_SECONDS);
  return students;
}

/** Invalidate the students cache. Called after every mutation. */
function _invalidateCache() {
  CacheService.getScriptCache().remove(CACHE_KEY_STUDENTS);
}


// ============================================================
// LOOKUP HELPERS — find a row by id / idempotency key
// ------------------------------------------------------------
// These read the SHEET DIRECTLY (not the cache) because mutations must
// operate on authoritative, real-time data — never on a possibly-stale cache.
// They return the 1-based SHEET ROW NUMBER plus the parsed object, so callers
// can write back to the exact row.
// ============================================================

/**
 * @returns {{rowNum:number, student:Object}|null}
 */
function _findRowByStudentId(sh, studentId) {
  const lastRow = sh.getLastRow();
  if (lastRow < FIRST_DATA_ROW) return null;
  const values = sh.getRange(FIRST_DATA_ROW, 1, lastRow - 1, STUDENT_COLS.length).getValues();
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === studentId) {
      return { rowNum: FIRST_DATA_ROW + i, student: _rowToObject(values[i]) };
    }
  }
  return null;
}

/**
 * Used for the IDEMPOTENCY check (TODO 4). Returns the existing student
 * object if this idempotency_key was already processed, else null.
 * @returns {Object|null}
 */
function _findByIdempotencyKey(sh, key) {
  if (!key) return null;
  const lastRow = sh.getLastRow();
  if (lastRow < FIRST_DATA_ROW) return null;
  // idempotency_key is column J (index 9, 1-based col 10)
  const keys = sh.getRange(FIRST_DATA_ROW, 10, lastRow - 1, 1).getValues();
  for (let i = 0; i < keys.length; i++) {
    if (keys[i][0] === key) {
      const row = sh.getRange(FIRST_DATA_ROW + i, 1, 1, STUDENT_COLS.length).getValues()[0];
      return _rowToObject(row);
    }
  }
  return null;
}


// ============================================================
// ID GENERATION — atomic counter in Config (TODO 1)
// ------------------------------------------------------------
// Counter lives in Config!B2 (named range ID_COUNTER). This is called ONLY
// from inside a write lock (repo_createStudent), so the read-increment-write
// sequence is safe from races. Format: STU-YYYY-NNNN.
// ============================================================

function _nextStudentId() {
  const config = _configSheet();
  const counterCell = config.getRange('B2'); // named range ID_COUNTER
  const current = Number(counterCell.getValue()) || 0;
  const next = current + 1;
  counterCell.setValue(next);

  // Academic year from settings (Config!D5). Fall back to current year.
  let year = config.getRange('D5').getValue();
  if (!year) year = new Date().getFullYear();

  const padded = String(next).padStart(4, '0');
  return `STU-${year}-${padded}`;
}


// ============================================================
// AUDIT LOG — append-only (one row per mutation)
// ============================================================

/**
 * @param {string} action  CREATE | UPDATE | DELETE | RESTORE
 * @param {string} studentId
 * @param {Object|null} oldValues  snapshot before change (null for CREATE)
 * @param {Object|null} newValues  snapshot after change  (null for DELETE)
 */
function _writeAudit(action, studentId, oldValues, newValues) {
  const audit = _auditSheet();
  let changedBy = 'unknown';
  try { changedBy = Session.getActiveUser().getEmail() || 'anonymous'; } catch (e) {}

  audit.appendRow([
    new Date().toISOString(),
    action,
    studentId,
    changedBy,
    oldValues ? JSON.stringify(oldValues) : '',
    newValues ? JSON.stringify(newValues) : ''
  ]);
}


// ============================================================
// CREATE — repo_createStudent()  [WRITE-LOCKED + IDEMPOTENT]
// ============================================================

/**
 * Inserts a fully-formed student record. The SERVICE layer is responsible
 * for validation and for filling derived fields; the repository owns:
 *   • the write lock
 *   • the idempotency short-circuit
 *   • atomic ID + timestamp + row_version assignment
 *   • the audit row
 *
 * @param {Object} data  validated payload:
 *        { name, score_pct, dob, subject_id, idempotency_key }
 * @returns {Object} the authoritative stored student object (THROWS on failure)
 */
function repo_createStudent(data) {
  const lock = LockService.getScriptLock();
  // waitLock THROWS if it can't acquire within the timeout — Service.gs catches it.
  lock.waitLock(LOCK_TIMEOUT_MS);
  try {
    const sh = _studentsSheet();

    // IDEMPOTENCY CHECK (inside the lock so two retries can't both insert):
    // if this submission was already processed, return the existing row.
    const existing = _findByIdempotencyKey(sh, data.idempotency_key);
    if (existing) {
      return existing; // no duplicate inserted — frontend treats this as success
    }

    const nowIso = new Date().toISOString();
    const student = {
      student_id:      _nextStudentId(),       // atomic, locked
      name:            data.name,
      score_pct:       data.score_pct,
      dob:             data.dob,
      subject_id:      data.subject_id,
      is_active:       true,
      created_at:      nowIso,
      updated_at:      nowIso,
      row_version:     1,
      idempotency_key: data.idempotency_key
    };

    // Append in ONE write (the next empty row).
    const targetRow = sh.getLastRow() + 1;
    sh.getRange(targetRow, 1, 1, STUDENT_COLS.length).setValues([_objectToRow(student)]);

    // Re-apply the teal DOB background on the new cell (TODO 6) so direct
    // sheet viewing stays consistent even outside conditional formatting.
    sh.getRange(targetRow, 4).setBackground('#CCFBF1');

    _writeAudit('CREATE', student.student_id, null, _publicFields(student));
    _invalidateCache();
    return student;
  } finally {
    lock.releaseLock(); // ALWAYS release, even on throw
  }
}


// ============================================================
// UPDATE — repo_updateStudent()  [WRITE-LOCKED + VERSION CHECK]
// ============================================================

/**
 * Applies changes to an existing record with OPTIMISTIC CONCURRENCY CONTROL.
 *
 * The version check happens INSIDE the lock against the live sheet value, so
 * it can't be defeated by a race. On mismatch we DO NOT write; we return a
 * sentinel so Service.gs can surface VERSION_CONFLICT to the client.
 *
 * @param {string} studentId
 * @param {Object} changes          { name?, score_pct?, dob?, subject_id? }
 * @param {number} expectedVersion  row_version the client last saw
 * @returns {Object} one of:
 *   { conflict: true,  current: <serverRecord> }   // version mismatch
 *   { conflict: false, student: <updatedRecord> }  // success
 *   (THROWS if the student doesn't exist)
 */
function repo_updateStudent(studentId, changes, expectedVersion) {
  const lock = LockService.getScriptLock();
  lock.waitLock(LOCK_TIMEOUT_MS);
  try {
    const sh = _studentsSheet();
    const found = _findRowByStudentId(sh, studentId);
    if (!found) throw new Error('Student not found: ' + studentId);

    const current = found.student;

    // OPTIMISTIC CONCURRENCY: compare versions on the LIVE row.
    if (Number(expectedVersion) !== Number(current.row_version)) {
      // No write performed. Return the current server record for the diff view.
      return { conflict: true, current: current };
    }

    // Merge allowed editable fields only — never let the client overwrite
    // system fields (id, created_at, row_version, idempotency_key).
    const updated = {
      ...current,
      name:        changes.name        !== undefined ? changes.name        : current.name,
      score_pct:   changes.score_pct   !== undefined ? changes.score_pct   : current.score_pct,
      dob:         changes.dob         !== undefined ? changes.dob         : current.dob,
      subject_id:  changes.subject_id  !== undefined ? changes.subject_id  : current.subject_id,
      updated_at:  new Date().toISOString(),
      row_version: Number(current.row_version) + 1  // bump version on success
    };

    sh.getRange(found.rowNum, 1, 1, STUDENT_COLS.length).setValues([_objectToRow(updated)]);
    sh.getRange(found.rowNum, 4).setBackground('#CCFBF1'); // keep DOB teal

    _writeAudit('UPDATE', studentId, _publicFields(current), _publicFields(updated));
    _invalidateCache();
    return { conflict: false, student: updated };
  } finally {
    lock.releaseLock();
  }
}


// ============================================================
// SOFT DELETE / RESTORE  [WRITE-LOCKED]
// ============================================================

/**
 * Soft delete: flips is_active to FALSE. Data is NEVER physically removed.
 * Bumps row_version and updated_at like any other mutation.
 * @returns {Object} the updated student (THROWS if not found)
 */
function repo_softDeleteStudent(studentId) {
  return _setActiveFlag(studentId, false, 'DELETE');
}

/**
 * Restore: flips is_active back to TRUE.
 * @returns {Object} the updated student (THROWS if not found)
 */
/**
 * True Restore: Reconstructs the row state from the last DELETE action in Audit_Log.
 * @returns {Object} the updated student (THROWS if not found or no audit history)
 */
// ============================================================
// BULK CREATE — repo_bulkCreateStudents()  [WRITE-LOCKED]  (NEW)
// ------------------------------------------------------------
// ONE lock, ONE batched setValues, ONE cache invalidation for the whole batch.
// Idempotency is enforced against the live sheet + within the batch itself.
// ============================================================
function repo_bulkCreateStudents(rows) {
  const lock = LockService.getScriptLock();
  lock.waitLock(LOCK_TIMEOUT_MS);
  try {
    const sh = _studentsSheet();
    const created = [];
    const newRowArrays = [];
    const auditRows = [];
    const seenKeys = {}; // dedupe within this batch

    // Snapshot existing idempotency keys ONCE (single column read).
    const existingKeys = _allIdempotencyKeys(sh);

    let changedBy = 'unknown';
    try { changedBy = Session.getActiveUser().getEmail() || 'anonymous'; } catch (e) {}

    rows.forEach(data => {
      const key = data.idempotency_key;
      if (key && (existingKeys[key] || seenKeys[key])) {
        return; // duplicate -> skip silently (idempotent)
      }
      if (key) seenKeys[key] = true;

      const nowIso = new Date().toISOString();
      const student = {
        student_id:      _nextStudentId(),
        name:            data.name,
        score_pct:       data.score_pct,
        dob:             data.dob,
        subject_id:      data.subject_id,
        is_active:       true,
        created_at:      nowIso,
        updated_at:      nowIso,
        row_version:     1,
        idempotency_key: key
      };
      created.push(student);
      newRowArrays.push(_objectToRow(student));
      auditRows.push([
        nowIso, 'CREATE', student.student_id, changedBy, '',
        JSON.stringify(_publicFields(student))
      ]);
    });

    if (newRowArrays.length > 0) {
      const startRow = sh.getLastRow() + 1;
      sh.getRange(startRow, 1, newRowArrays.length, STUDENT_COLS.length)
        .setValues(newRowArrays);
      // Keep the DOB column teal for the inserted block (single range call).
      sh.getRange(startRow, 4, newRowArrays.length, 1).setBackground('#CCFBF1');

      const audit = _auditSheet();
      audit.getRange(audit.getLastRow() + 1, 1, auditRows.length, 6).setValues(auditRows);

      _invalidateCache();
    }

    return created;
  } finally {
    lock.releaseLock();
  }
}

/** Returns a { key: true } map of all idempotency keys currently in the sheet. */
function _allIdempotencyKeys(sh) {
  const map = {};
  const lastRow = sh.getLastRow();
  if (lastRow < FIRST_DATA_ROW) return map;
  const keys = sh.getRange(FIRST_DATA_ROW, 10, lastRow - 1, 1).getValues();
  keys.forEach(r => { if (r[0]) map[r[0]] = true; });
  return map;
}
// ============================================================
// RESTORE — repo_restoreStudent()  [WRITE-LOCKED]  (OPTIMIZED)
// ------------------------------------------------------------
// Reconstructs the row from the most recent DELETE entry in Audit_Log.
// Reads only the columns it needs (action, student_id, changed_by, old_values)
// and scans newest-first, instead of pulling the whole sheet with getDataRange.
// ============================================================
function repo_restoreStudent(studentId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(LOCK_TIMEOUT_MS);
  try {
    const sh = _studentsSheet();
    const auditSh = _auditSheet();

    const found = _findRowByStudentId(sh, studentId);
    if (!found) throw new Error('Student not found: ' + studentId);

    // Read only audit columns B..E (action, student_id, changed_by, old_values).
    const lastAuditRow = auditSh.getLastRow();
    let lastKnownState = null;
    if (lastAuditRow >= 2) {
      const auditVals = auditSh.getRange(2, 2, lastAuditRow - 1, 4).getValues();
      // auditVals[i] = [action, student_id, changed_by, old_values]
      for (let i = auditVals.length - 1; i >= 0; i--) {
        if (auditVals[i][0] === 'DELETE' && auditVals[i][1] === studentId) {
          try { lastKnownState = JSON.parse(auditVals[i][3]); } catch (e) {}
          break;
        }
      }
    }
    if (!lastKnownState) {
      throw new Error('No historical DELETE record found to restore from.');
    }

    const current = found.student;
    const updated = {
      ...current,
      name:        lastKnownState.name        != null ? lastKnownState.name        : current.name,
      score_pct:   lastKnownState.score_pct   != null ? lastKnownState.score_pct   : current.score_pct,
      dob:         lastKnownState.dob         != null ? lastKnownState.dob         : current.dob,
      subject_id:  lastKnownState.subject_id  != null ? lastKnownState.subject_id  : current.subject_id,
      is_active:   true,
      updated_at:  new Date().toISOString(),
      row_version: Number(current.row_version) + 1
    };

    sh.getRange(found.rowNum, 1, 1, STUDENT_COLS.length).setValues([_objectToRow(updated)]);
    sh.getRange(found.rowNum, 4).setBackground('#CCFBF1');

    _writeAudit('RESTORE', studentId, _publicFields(current), _publicFields(updated));
    _invalidateCache();
    return updated;
  } finally {
    lock.releaseLock();
  }
}



/** Shared implementation for soft-delete and restore. */
function _setActiveFlag(studentId, activeValue, action) {
  const lock = LockService.getScriptLock();
  lock.waitLock(LOCK_TIMEOUT_MS);
  try {
    const sh = _studentsSheet();
    const found = _findRowByStudentId(sh, studentId);
    if (!found) throw new Error('Student not found: ' + studentId);

    const current = found.student;
    const updated = {
      ...current,
      is_active:   activeValue,
      updated_at:  new Date().toISOString(),
      row_version: Number(current.row_version) + 1
    };

    sh.getRange(found.rowNum, 1, 1, STUDENT_COLS.length).setValues([_objectToRow(updated)]);

    _writeAudit(action, studentId, _publicFields(current), _publicFields(updated));
    _invalidateCache();
    return updated;
  } finally {
    lock.releaseLock();
  }
}


// ============================================================
// AUDIT SNAPSHOT HELPER
// ------------------------------------------------------------
// Strips internal-only noise from audit JSON; keeps the meaningful fields.
// ============================================================
function _publicFields(s) {
  return {
    name:        s.name,
    score_pct:   s.score_pct,
    dob:         s.dob,
    subject_id:  s.subject_id,
    is_active:   s.is_active,
    row_version: s.row_version
  };
}
