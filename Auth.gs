// ============================================================
// Auth.gs — Role-Based Access Control (RBAC)
// Google authenticates the user; we authorize them here.
// ============================================================

const ROLE_EDITOR = 'editor';
const ROLE_VIEWER = 'viewer';


/** Returns the verified email of the current user, or '' if unavailable. */
function _currentUserEmail() {
  try {
    var email = Session.getActiveUser().getEmail();
    return (email || '').trim().toLowerCase();
  } catch (e) {
    return '';
  }
}


/** Reads editor emails (Config J2:J) + default role (Config K2). Cached briefly. */
// AFTER — everything is inside the try/catch:
function _loadRbacConfig() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('RBAC_CONFIG_V1');
  if (cached) return JSON.parse(cached);

  let editors = [];
  let defaultRole = ROLE_VIEWER;
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config'); // ← moved inside
    const last = sh.getLastRow();
    if (last >= 2) {
      editors = sh.getRange(2, 10, last - 1, 1).getValues()
        .map(r => String(r[0] || '').trim().toLowerCase())
        .filter(Boolean);
    }
    const dr = String(sh.getRange('K2').getValue() || '').trim().toLowerCase();
    if (dr === ROLE_EDITOR || dr === ROLE_VIEWER) defaultRole = dr;
  } catch (e) { /* fall through to safe defaults — empty editors, viewer role */ }

  const cfg = { editors: editors, defaultRole: defaultRole };
  cache.put('RBAC_CONFIG_V1', JSON.stringify(cfg), 60);
  return cfg;
}

/** Resolve the current user's role: 'editor' or 'viewer'. */
function _resolveRole() {
  const email = _currentUserEmail();

  // FAIL CLOSED: no verified identity => always viewer, never editor.
  if (!email) return ROLE_VIEWER;

  const cfg = _loadRbacConfig();
  if (cfg.editors.indexOf(email) !== -1) return ROLE_EDITOR;

  // Known-but-not-allowlisted users get the default, but cap it at viewer
  // unless the default is explicitly editor for a verified user.
  return cfg.defaultRole === ROLE_EDITOR ? ROLE_EDITOR : ROLE_VIEWER;
}


/** Throw if the current user is not an editor. Used to guard every mutation. */
function _requireEditor() {
  if (_resolveRole() !== ROLE_EDITOR) {
    const e = new Error('You have view-only access. Editing is restricted to authorized users.');
    e.__rbac = true;
    throw e;
  }
}

/** Clear the RBAC cache after you change the editor list. */
function rbac_clearCache() {
  CacheService.getScriptCache().remove('RBAC_CONFIG_V1');
}
