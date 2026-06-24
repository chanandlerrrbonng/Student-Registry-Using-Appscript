// ============================================================
// ===  PASTE THIS INTO AI.gs  (replace the whole file)  ===
// ============================================================
 
// Primary model — stable, production-ready. Google guarantees
// this name keeps working across minor releases.
const GEMINI_MODEL          = 'gemini-2.0-flash';
// Model-level fallback in case the primary endpoint is degraded.
const GEMINI_MODEL_FALLBACK = 'gemini-1.5-flash';
 
// HTTP codes that are transient and worth retrying.
const GEMINI_RETRY_CODES = [429, 500, 503, 504];
const GEMINI_MAX_RETRIES = 3;
 
 
/**
 * Call Gemini to parse rawText into an array of proposed student rows.
 * Retries on transient errors, then falls back to the secondary model,
 * then throws so the caller (ai_parseStudents) can use the local parser.
 *
 * BOUNDARY: this file ONLY proposes rows. Nothing here writes to the sheet.
 */
function ai_parseWithGemini(rawText) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    throw new Error('GEMINI_API_KEY missing. Add it to Script Properties.');
  }
 
  // Try primary model first; on failure try the fallback model once.
  const modelsToTry = [GEMINI_MODEL, GEMINI_MODEL_FALLBACK];
 
  for (var modelIdx = 0; modelIdx < modelsToTry.length; modelIdx++) {
    var model = modelsToTry[modelIdx];
 
    try {
      return _callGeminiWithRetry(apiKey, model, rawText);
    } catch (err) {
      if (modelIdx < modelsToTry.length - 1) {
        // Log and try the next model.
        Logger.log('Model ' + model + ' failed (' + err.message + '); trying fallback model.');
        continue;
      }
      // All models exhausted — re-throw so the caller can use local parser.
      throw err;
    }
  }
}
 
 
/**
 * Makes the actual HTTP request with up to GEMINI_MAX_RETRIES retries
 * using exponential back-off for transient failures.
 */
function _callGeminiWithRetry(apiKey, model, rawText) {
  // Pull canonical subjects so the prompt + schema stay in sync with Config.
  var subjects = [];
  try { subjects = getSubjectList() || []; } catch (e) { subjects = []; }
  var subjectList = subjects.length
    ? subjects.join(', ')
    : 'Mathematics, Science, English, History, Computer Science, Physics, Chemistry, Biology, Geography, Economics';
 
  var prompt =
    'You are a data-parsing assistant. Convert the unstructured text below into ' +
    'student records.\n\n' +
    'Rules for each field:\n' +
    '- name: full name in Title Case.\n' +
    '- score_pct: numeric 0-100, digits only (strip any "%" sign).\n' +
    '- dob: date of birth. Output ONLY in DD/MM/YYYY format. Convert ANY input ' +
    'format (e.g. "2003-11-22", "30.08.2005", "5 dec 2003", "9th july 2004", ' +
    '"25th June 2006", "30-04-2004", "26/05/2006").\n' +
    '- subject_id: map shorthand to one of these EXACT canonical values: ' +
    subjectList + '. ' +
    'Examples: "Maths"/"Math" -> "Mathematics", "bio" -> "Biology", ' +
    '"CS" -> "Computer Science", "geo" -> "Geography", "Phys" -> "Physics", ' +
    '"Sci" -> "Science", "Eng" -> "English".\n\n' +
    'Text to parse:\n"""' + rawText + '"""';
 
  var payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: {
      temperature: 0.1,
      responseMimeType: 'application/json',
      responseSchema: {
        type: 'ARRAY',
        items: {
          type: 'OBJECT',
          properties: {
            name:       { type: 'STRING' },
            score_pct:  { type: 'NUMBER' },
            dob:        { type: 'STRING' },
            subject_id: { type: 'STRING' }
          },
          required: ['name', 'score_pct', 'dob', 'subject_id']
        }
      }
    }
  };
 
  var options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true   // so we can read error bodies and decide whether to retry
  };
 
  var endpoint =
    'https://generativelanguage.googleapis.com/v1beta/models/' +
    model + ':generateContent?key=' + apiKey;
 
  var lastCode = 0;
  var lastText = '';
 
  for (var attempt = 0; attempt <= GEMINI_MAX_RETRIES; attempt++) {
    // Exponential back-off: 0 s, 1 s, 2 s, 4 s (skipped on attempt 0).
    if (attempt > 0) {
      Utilities.sleep(Math.pow(2, attempt - 1) * 1000);
      Logger.log('Gemini retry ' + attempt + '/' + GEMINI_MAX_RETRIES +
                 ' for model ' + model);
    }
 
    var response = UrlFetchApp.fetch(endpoint, options);
    lastCode = response.getResponseCode();
    lastText = response.getContentText();
 
    if (lastCode === 200) {
      // Success — parse and return.
      return _extractGeminiRows(lastText);
    }
 
    var isRetryable = GEMINI_RETRY_CODES.indexOf(lastCode) !== -1;
    if (!isRetryable || attempt === GEMINI_MAX_RETRIES) {
      // Non-retryable error, or retries exhausted.
      Logger.log('Gemini API Error (model=' + model + ') HTTP ' + lastCode + ': ' + lastText);
      throw new Error('AI provider returned HTTP ' + lastCode +
        ' (model: ' + model + '). Will attempt local parsing fallback.');
    }
 
    Logger.log('Gemini transient HTTP ' + lastCode + ' on attempt ' + attempt +
               ' — will retry.');
  }
 
  // Should be unreachable, but satisfy the code path.
  throw new Error('Gemini: all retries exhausted (HTTP ' + lastCode + ').');
}
 
 
/**
 * Extract the rows array from a successful (HTTP 200) Gemini response body.
 */
function _extractGeminiRows(responseText) {
  var json;
  try {
    json = JSON.parse(responseText);
  } catch (e) {
    throw new Error('AI provider returned a non-JSON envelope.');
  }
 
  var aiText;
  try {
    aiText = json.candidates[0].content.parts[0].text;
  } catch (e) {
    Logger.log('Unexpected Gemini envelope: ' + responseText);
    throw new Error('AI provider returned no usable content.');
  }
 
  var parsed = _safeParseJsonArray(aiText);
  if (!Array.isArray(parsed)) {
    throw new Error('AI output was not a JSON array of rows.');
  }
 
  // Normalize every row server-side so we never depend on the model being perfect.
  return parsed.map(_normalizeProposedRow);
}
 
 
// ============================================================
// LOCAL FALLBACK PARSER
// ------------------------------------------------------------
// Pure regex — no network call. Used when Gemini is unavailable.
// Handles the same variety of inputs the prompt describes.
//
// Format assumption: each LINE is one student. Field ORDER within
// the line can vary, so we detect fields by their shape, not position.
//   • score  : a standalone integer 0–100
//   • date   : dd.mm.yyyy | dd/mm/yyyy | dd-mm-yyyy | yyyy-mm-dd |
//              "25th June 2006" | "June 25 2006" | "25 June 2006"
//   • subject: matched against the canonical list + alias table
//   • name   : whatever text remains after the above are stripped
// ============================================================
 
function _parseStudentsLocally(rawText) {
  var lines = rawText.split('\n').map(function(l) { return l.trim(); }).filter(Boolean);
 
  var MONTHS = {
    jan:1, feb:2, mar:3, apr:4, may:5,  jun:6,
    jul:7, aug:8, sep:9, oct:10,nov:11, dec:12
  };
 
  // All aliases, longest first so "Computer Science" beats "Science".
  var SUBJECT_ALIASES = {
    'computer science': 'Computer Science',
    'comp sci':         'Computer Science',
    'mathematics':      'Mathematics',
    'maths':            'Mathematics',
    'math':             'Mathematics',
    'science':          'Science',
    'english':          'English',
    'history':          'History',
    'physics':          'Physics',
    'chemistry':        'Chemistry',
    'biology':          'Biology',
    'geography':        'Geography',
    'economics':        'Economics',
    'cs':               'Computer Science',
    'sci':              'Science',
    'eng':              'English',
    'hist':             'History',
    'phys':             'Physics',
    'chem':             'Chemistry',
    'bio':              'Biology',
    'geo':              'Geography',
    'econ':             'Economics'
  };
 
  // Sort alias keys longest-first to prefer the most specific match.
  var aliasKeys = Object.keys(SUBJECT_ALIASES).sort(function(a, b) {
    return b.length - a.length;
  });
 
  function pad2(v) { return String(v).padStart(2, '0'); }
 
  return lines.map(function(line) {
    var remaining = line;
    var dob = '';
    var score_pct = '';
    var subject_id = '';
    var m;
 
    // ------------------------------------------------------------------
    // STEP 1 — Extract date
    // Priority order: numeric patterns first (most unambiguous), then
    // text-month patterns.
    // ------------------------------------------------------------------
 
    // Pattern A: dd.mm.yyyy  /  dd/mm/yyyy  /  dd-mm-yyyy
    m = remaining.match(/\b(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{4})\b/);
    if (m) {
      dob = pad2(m[1]) + '/' + pad2(m[2]) + '/' + m[3];
      remaining = remaining.replace(m[0], ' ');
    }
 
    // Pattern B: yyyy-mm-dd  /  yyyy.mm.dd  /  yyyy/mm/dd
    if (!dob) {
      m = remaining.match(/\b(\d{4})[.\/-](\d{1,2})[.\/-](\d{1,2})\b/);
      if (m) {
        dob = pad2(m[3]) + '/' + pad2(m[2]) + '/' + m[1];
        remaining = remaining.replace(m[0], ' ');
      }
    }
 
    // Pattern C: "25th June 2006"  /  "25 June 2006"  /  "25 jun 2006"
    if (!dob) {
      m = remaining.match(/\b(\d{1,2})(?:st|nd|rd|th)?\s+([a-zA-Z]{3,9})\s+(\d{4})\b/i);
      if (m) {
        var mon = MONTHS[m[2].toLowerCase().slice(0, 3)];
        if (mon) {
          dob = pad2(m[1]) + '/' + pad2(String(mon)) + '/' + m[3];
          remaining = remaining.replace(m[0], ' ');
        }
      }
    }
 
    // Pattern D: "June 25 2006"  /  "June 25, 2006"
    if (!dob) {
      m = remaining.match(/\b([a-zA-Z]{3,9})\s+(\d{1,2})(?:st|nd|rd|th)?,?\s+(\d{4})\b/i);
      if (m) {
        var mon = MONTHS[m[1].toLowerCase().slice(0, 3)];
        if (mon) {
          dob = pad2(m[2]) + '/' + pad2(String(mon)) + '/' + m[3];
          remaining = remaining.replace(m[0], ' ');
        }
      }
    }
 
    // Clean up leftover separators from the date removal.
    remaining = remaining.replace(/\s{2,}/g, ' ').trim();
 
    // ------------------------------------------------------------------
    // STEP 2 — Extract score  (standalone integer 0–100)
    // We look for a word-boundary-delimited number that is NOT immediately
    // adjacent to a "/" or "." on both sides (which would indicate it's
    // still part of a date that the regex above missed).
    // ------------------------------------------------------------------
    m = remaining.match(/(?<![\/.])\b(100|[0-9]{1,2})\b(?![\/.])/);
    if (m) {
      score_pct = parseFloat(m[1]);
      // Replace only the first occurrence to avoid eating a name digit.
      remaining = remaining.replace(m[0], ' ');
    }
 
    remaining = remaining.replace(/\s{2,}/g, ' ').trim();
 
    // ------------------------------------------------------------------
    // STEP 3 — Extract subject  (longest match wins)
    // ------------------------------------------------------------------
    for (var i = 0; i < aliasKeys.length; i++) {
      var key    = aliasKeys[i];
      // Build a word-boundary regex that handles spaces within the alias
      // (e.g., "Computer Science" → \bComputer\s+Science\b).
      var regexSrc = '\\b' + key.replace(/ +/g, '\\s+') + '\\b';
      var subjectRe = new RegExp(regexSrc, 'i');
      if (subjectRe.test(remaining)) {
        subject_id = SUBJECT_ALIASES[key];
        remaining  = remaining.replace(subjectRe, ' ');
        break;
      }
    }
 
    remaining = remaining.replace(/\s{2,}/g, ' ').trim();
 
    // ------------------------------------------------------------------
    // STEP 4 — Remaining text is the name
    // ------------------------------------------------------------------
    var name = remaining.replace(/\s+/g, ' ').trim();
    name = _aiTitleCase(name);
 
    return _normalizeProposedRow({
      name:       name,
      score_pct:  score_pct,
      dob:        dob,
      subject_id: subject_id
    });
  });
}
 
 
// ============================================================
// JSON EXTRACTION HELPERS
// ============================================================
 
/**
 * Parse a JSON array even if the model wrapped it in ```json fences or added
 * stray prose. Returns the parsed value or throws.
 */
function _safeParseJsonArray(s) {
  var t = String(s || '').trim();
  // Strip markdown code fences if present.
  t = t.replace(/^```(?:json)?\s*/i, '').replace(/\s*```$/i, '').trim();
  try {
    return JSON.parse(t);
  } catch (e) {
    // Last resort: extract the first [...] block.
    var start = t.indexOf('[');
    var end   = t.lastIndexOf(']');
    if (start !== -1 && end !== -1 && end > start) {
      return JSON.parse(t.slice(start, end + 1));
    }
    throw new Error('Could not parse AI output into JSON.');
  }
}
 
 
// ============================================================
// ROW NORMALIZATION  (shared by Gemini path AND local parser)
// ============================================================
 
/**
 * Clean a single proposed row so the preview + downstream validation behave.
 * Best-effort only; invalid rows are caught later by _validateStudentPayload.
 */
function _normalizeProposedRow(row) {
  row = row || {};
  return {
    name:       _aiTitleCase(String(row.name || '').trim()),
    score_pct:  _aiCoerceScore(row.score_pct),
    dob:        _aiNormalizeDobString(String(row.dob || '').trim()),
    subject_id: _aiNormalizeSubject(String(row.subject_id || '').trim())
  };
}
 
function _aiTitleCase(s) {
  return s.replace(/\s+/g, ' ')
          .replace(/\w\S*/g, function(w) {
            return w.charAt(0).toUpperCase() + w.slice(1).toLowerCase();
          });
}
 
function _aiCoerceScore(v) {
  var n = parseFloat(String(v).replace(/[^\d.]/g, ''));
  return isNaN(n) ? '' : n;
}
 
/**
 * Convert many human date formats into dd/MM/yyyy.
 * Handles: dd/mm/yyyy, dd-mm-yyyy, dd.mm.yyyy (NEW), yyyy-mm-dd,
 *          "5 dec 2003", "9th july 2004", "july 9 2004", "25th June 2006".
 */
function _aiNormalizeDobString(s) {
  if (!s) return '';
 
  // dd/MM/yyyy, dd-mm-yyyy, or dd.mm.yyyy  ← dot-separator added here
  var m = s.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{4})$/);
  if (m) return _aiPad(m[1]) + '/' + _aiPad(m[2]) + '/' + m[3];
 
  // ISO yyyy-mm-dd (also accepts yyyy.mm.dd and yyyy/mm/dd)
  m = s.match(/^(\d{4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})$/);
  if (m) return _aiPad(m[3]) + '/' + _aiPad(m[2]) + '/' + m[1];
 
  var months = {
    jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6,
    jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12
  };
 
  // "5 dec 2003", "9th july 2004", "5 December, 2003", "25th June 2006"
  m = s.toLowerCase().match(/(\d{1,2})(?:st|nd|rd|th)?\s+([a-z]+)\.?,?\s+(\d{4})/);
  if (m) {
    var mon = months[m[2].slice(0, 3)];
    if (mon) return _aiPad(m[1]) + '/' + _aiPad(String(mon)) + '/' + m[3];
  }
 
  // "july 9 2004" style (month first)
  m = s.toLowerCase().match(/([a-z]+)\.?\s+(\d{1,2})(?:st|nd|rd|th)?,?\s+(\d{4})/);
  if (m) {
    var mon2 = months[m[1].slice(0, 3)];
    if (mon2) return _aiPad(m[2]) + '/' + _aiPad(String(mon2)) + '/' + m[3];
  }
 
  return s; // leave as-is; validation will reject if unusable
}
 
function _aiPad(v) { return String(v).padStart(2, '0'); }
 
/**
 * Map shorthand subjects to canonical values: exact match first, then aliases.
 */
function _aiNormalizeSubject(s) {
  if (!s) return '';
  var canonical = _aiSafeGetSubjects();
 
  // exact match (case-insensitive)
  var exact = canonical.find(function(c) { return c.toLowerCase() === s.toLowerCase(); });
  if (exact) return exact;
 
  var aliases = {
    maths: 'Mathematics', math: 'Mathematics', sci: 'Science', eng: 'English',
    hist: 'History', cs: 'Computer Science', 'comp sci': 'Computer Science',
    phys: 'Physics', chem: 'Chemistry', bio: 'Biology', geo: 'Geography',
    econ: 'Economics'
  };
  var alias = aliases[s.toLowerCase()];
  if (alias && canonical.indexOf(alias) !== -1) return alias;
 
  return s; // unknown; validation rejects it
}
 
function _aiSafeGetSubjects() {
  try { return getSubjectList() || []; } catch (e) { return []; }
}
 
 
