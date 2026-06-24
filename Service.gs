// ============================================================
// Service.gs  — BUSINESS LOGIC + CLIENT API SURFACE
// ------------------------------------------------------------
// Every function exposed to google.script.run:
//   • Validates input BEFORE touching the repository
//   • Wraps ALL repository calls in try/catch
//   • NEVER throws a raw Error to the client
//   • ALWAYS returns the plain serializable contract:
//        { success, data?, error?, field?, code? }
//
// This file does NOT call SpreadsheetApp directly — only Repository.gs does.
// ============================================================


// ------------------------------------------------------------
// RESPONSE BUILDERS — single source of truth for the contract shape
// ------------------------------------------------------------
function _ok(data)         { return { success: true,  data: data }; }
function _fail(msg, field) { return { success: false, error: msg, field: field || null }; }
function _conflict(current) {
  return { success: false, code: 'VERSION_CONFLICT',
           error: 'Record was modified by someone else.', data: current };
}


// ============================================================
// READ — exposed to client
// ============================================================
function service_getAllStudents() {
  try {
    return _ok(repo_getAllStudents());
  } catch (err) {
    return _fail('Failed to load students: ' + err.message);
  }
}

// ============================================================
// COMBINED INITIAL LOAD — service_getInitialData()
// Called once on page load instead of three separate round-trips.
// ============================================================
function service_getInitialData() {
  try {
    return {
      success: true,
      data: {
        students: repo_getAllStudents(),
        subjects: getSubjectList(),
        settings: getAppSettings(),
        role:     _resolveRole(),          // ← NEW
        email:    _currentUserEmail()      // ← NEW (optional, for a header badge)
      }
    };
  } catch (err) {
    return { success: false, error: 'Failed to load initial data: ' + err.message };
  }
}


// ============================================================
// VALIDATION — shared, pure (no SpreadsheetApp)
// ============================================================
function _validateStudentPayload(p) {
  if (!p || typeof p.name !== 'string' || !p.name.trim()) {
    return { message: 'Name is required.', field: 'name' };
  }

  const score = Number(p.score_pct);
  if (p.score_pct === '' || p.score_pct === null || isNaN(score)) {
    return { message: 'Score must be a number.', field: 'score_pct' };
  }
  if (score < 0 || score > 100) {
    return { message: 'Score must be between 0 and 100.', field: 'score_pct' };
  }

  const dobErr = _validateDob(p.dob);
  if (dobErr) return { message: dobErr, field: 'dob' };

  const subjects = _safeGetSubjects();
  if (!p.subject_id || subjects.indexOf(p.subject_id) === -1) {
    return { message: 'Subject must be one of the canonical values.', field: 'subject_id' };
  }

  return null;
}

function _validateDob(dob) {
  if (!dob || typeof dob !== 'string') return 'Date of birth is required.';
  const m = dob.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (!m) return 'Date must be in DD/MM/YYYY format.';
  const dd = Number(m[1]), mm = Number(m[2]), yyyy = Number(m[3]);
  const d = new Date(yyyy, mm - 1, dd);
  if (d.getFullYear() !== yyyy || d.getMonth() !== mm - 1 || d.getDate() !== dd) {
    return 'Not a real calendar date.';
  }
  if (d > new Date()) return 'Date of birth cannot be in the future.';
  return null;
}

function _normalizeName(name) {
  return name.trim().replace(/\s+/g, ' ')
    .replace(/\w\S*/g, w => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase());
}

function _safeGetSubjects() {
  try { return getSubjectList() || []; } catch (e) { return []; }
}


// ============================================================
// CREATE — service_createStudent()  [exposed to client]   (FIXED)
// ------------------------------------------------------------
// FIX: the old version had unreachable dead code after the first return
//      that referenced an undefined `result` variable. Removed entirely.
// ============================================================
function service_createStudent(payload) {
  try {
     _requireEditor(); 
    const err = _validateStudentPayload(payload);
    if (err) return _fail(err.message, err.field);

    const idempotencyKey = payload.idempotency_key || Utilities.getUuid();

    const clean = {
      name:            _normalizeName(payload.name),
      score_pct:       Number(payload.score_pct),
      dob:             payload.dob,
      subject_id:      payload.subject_id,
      idempotency_key: idempotencyKey
    };

    const stored = repo_createStudent(clean);
    _checkAndSendRiskAlert(stored);   // best-effort; never throws
    return _ok(stored);
  } catch (err) {
    if (err.__rbac) return { success: false, code: 'FORBIDDEN', error: err.message };  // ← NEW
    return _fail('Could not create student: ' + err.message);
  }
}


// ============================================================
// BULK CREATE — service_bulkCreateStudents()  [exposed to client]  (NEW)
// ------------------------------------------------------------
// Replaces the frontend's per-row await loop. ONE call, ONE lock acquisition
// inside the repo, ONE cache invalidation. Each row is validated independently
// so one bad row doesn't sink the batch — per-row errors are reported back.
// ============================================================
function service_bulkCreateStudents(rows) {
  try {
     _requireEditor(); 
    if (!Array.isArray(rows) || rows.length === 0) {
      return _fail('No rows provided for import.');
    }
    if (rows.length > 200) {
      return _fail('Import limited to 200 rows per batch.');
    }

    const valid = [];
    const errors = [];

    rows.forEach((row, i) => {
      const err = _validateStudentPayload(row);
      if (err) {
        errors.push({ index: i, name: row && row.name, message: err.message, field: err.field });
        return;
      }
      valid.push({
        name:            _normalizeName(row.name),
        score_pct:       Number(row.score_pct),
        dob:             row.dob,
        subject_id:      row.subject_id,
        idempotency_key: row.idempotency_key || Utilities.getUuid()
      });
    });

    const created = valid.length ? repo_bulkCreateStudents(valid) : [];

    _sendBatchRiskAlerts(created);

    return _ok({ created: created, createdCount: created.length, errors: errors });
  } catch (err) {
    if (err.__rbac) return { success: false, code: 'FORBIDDEN', error: err.message };  // ← NEW
    return _fail('Bulk import failed: ' + err.message);
  }
}


// ============================================================
// UPDATE — service_updateStudent()  [exposed to client]
// ============================================================
function service_updateStudent(req) {
  try {
    _requireEditor(); 
    if (!req || !req.student_id) return _fail('Missing student_id.');
    if (req.expected_version === undefined || req.expected_version === null) {
      return _fail('Missing expected_version (required for concurrency control).');
    }

    const err = _validateStudentPayload(req.changes);
    if (err) return _fail(err.message, err.field);

    const clean = {
      name:       _normalizeName(req.changes.name),
      score_pct:  Number(req.changes.score_pct),
      dob:        req.changes.dob,
      subject_id: req.changes.subject_id
    };

    const result = repo_updateStudent(req.student_id, clean, req.expected_version);

    if (result.conflict) {
      return _conflict(result.current);
    }
    _checkAndSendRiskAlert(result.student);   // alert on updates too
    return _ok(result.student);
  } catch (err) {
    if (err.__rbac) return { success: false, code: 'FORBIDDEN', error: err.message };  // ← NEW
    return _fail('Could not update student: ' + err.message);
  }
}


// ============================================================
// SOFT DELETE — service_deleteStudent()  [exposed to client]
// ============================================================
function service_deleteStudent(req) {
  try {
    _requireEditor();                    // ← NEW

    if (!req || !req.student_id) return _fail('Missing student_id.');
    return _ok(repo_softDeleteStudent(req.student_id));
  } catch (err) {
    if (err.__rbac) return { success: false, code: 'FORBIDDEN', error: err.message };  // ← NEW
    return _fail('Could not delete student: ' + err.message);
  }
}


// ============================================================
// RESTORE — service_restoreStudent()  [exposed to client]
// ============================================================
function service_restoreStudent(req) {
  try {
    _requireEditor();
    if (!req || !req.student_id) return _fail('Missing student_id.');
    return _ok(repo_restoreStudent(req.student_id));
  } catch (err) {
    if (err.__rbac) return { success: false, code: 'FORBIDDEN', error: err.message };
    return _fail('Could not restore student: ' + err.message);
  }
}


// ============================================================
// AI BULK IMPORT — ai_parseStudents()  [exposed to client]
// ------------------------------------------------------------
// ONLY proposes structured rows; never writes. Every proposed row must still
// pass through service_createStudent / service_bulkCreateStudents.
// ============================================================
// ============================================================
// ===  REPLACE ai_parseStudents() IN Service.gs WITH THIS  ===
// ============================================================
 
/**
 * AI BULK IMPORT — ai_parseStudents()  [exposed to client]
 *
 * Strategy:
 *   1. Try Gemini (with internal retries + model fallback).
 *   2. If Gemini fails for any reason, silently fall through to
 *      the local regex parser (_parseStudentsLocally).
 *   3. Only fail if BOTH paths are unavailable or produce nothing.
 *
 * ONLY proposes rows — never writes to the sheet.
 */
function ai_parseStudents(rawText) {
  try {
    _requireEditor();
    if (!rawText || !rawText.trim()) return _fail('No text provided to parse.');
 
    // --- PATH 1: Gemini ---
    if (typeof ai_parseWithGemini === 'function') {
      try {
        var proposed = ai_parseWithGemini(rawText);
        if (Array.isArray(proposed) && proposed.length > 0) {
          return _ok(proposed);
        }
      } catch (aiErr) {
        // Gemini failed (network, 503, quota, etc.) — log and fall through.
        Logger.log('Gemini unavailable: ' + aiErr.message + '. Falling back to local parser.');
      }
    }
 
    // --- PATH 2: Local regex parser (offline fallback) ---
    if (typeof _parseStudentsLocally === 'function') {
      var localProposed = _parseStudentsLocally(rawText);
      if (Array.isArray(localProposed) && localProposed.length > 0) {
        Logger.log('Local parser used as fallback; produced ' + localProposed.length + ' row(s).');
        return _ok(localProposed);
      }
    }
 
    return _fail('Could not parse any student rows from the provided text. ' +
                 'Check that each line contains a name, score, subject, and date of birth.');
 
  } catch (err) {
    if (err.__rbac) return { success: false, code: 'FORBIDDEN', error: err.message };
    return _fail('AI parsing failed: ' + err.message);
  }
}
 


// ============================================================
// AT-RISK EMAIL ALERTS  (FIXED — throttled, single Config read, escaped)
// ============================================================
function _checkAndSendRiskAlert(student) {
  try {
    const riskThreshold = _getRiskThreshold();
    if (Number(student.score_pct) >= riskThreshold) return;
    _sendRiskEmail(student, riskThreshold);
  } catch (e) {
    Logger.log('Email alert failed: ' + e.message);
  }
}

/** For bulk imports: hard-cap the number of emails to protect the daily quota. */
function _sendBatchRiskAlerts(students) {
  try {
    if (!students || students.length === 0) return;
    const riskThreshold = _getRiskThreshold();
    const atRisk = students.filter(s => Number(s.score_pct) < riskThreshold);
    if (atRisk.length === 0) return;

    const MAX_EMAILS = 5;
    atRisk.slice(0, MAX_EMAILS).forEach(s => {
      try { _sendRiskEmail(s, riskThreshold); } catch (e) { Logger.log(e.message); }
    });
    if (atRisk.length > MAX_EMAILS) {
      Logger.log('Suppressed ' + (atRisk.length - MAX_EMAILS) +
        ' additional risk alerts (batch cap).');
    }
  } catch (e) {
    Logger.log('Batch alert failed: ' + e.message);
  }
}

function _getRiskThreshold() {
  try {
    const settings = getAppSettings();
    return Number(settings.AT_RISK_THRESHOLD || 50);
  } catch (e) {
    return 50;
  }
}

function _sendRiskEmail(student, riskThreshold) {
  let email = '';
  try { email = Session.getActiveUser().getEmail(); } catch (e) {}
  if (!email) return;

  // Escape interpolated values so a crafted name can't inject HTML.
  const esc = s => String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

  const subject = 'Action Required: Student At Risk (' + student.name + ')';
  const body =
    '<h3>At-Risk Student Alert</h3>' +
    '<p>The following student is below the ' + riskThreshold + '% threshold:</p>' +
    '<ul>' +
    '<li><b>Name:</b> ' + esc(student.name) + '</li>' +
    '<li><b>ID:</b> ' + esc(student.student_id) + '</li>' +
    '<li><b>Score:</b> <span style="color:red;">' + esc(student.score_pct) + '%</span></li>' +
    '<li><b>Subject:</b> ' + esc(student.subject_id) + '</li>' +
    '</ul>' +
    '<p>Please review their record in the NoQs Registry.</p>';

  MailApp.sendEmail({ to: email, subject: subject, htmlBody: body });
}
