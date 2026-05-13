/**
 * KAL ACADEMY - Global constants shared across all .gs files.
 * Declared here once to avoid "Identifier already declared" errors
 * when multiple script files are present in the same project.
 */

// ── Column indices (1-based) ──────────────────────────────────────────────────
const COL = Object.freeze({
  ROW_NUM:    1,   // A
  DESC:       2,   // B
  FILENAME:   3,   // C
  FILETYPE:   4,   // D
  VERSION:    5,   // E
  FOLDER:     6,   // F
  LINK:       7,   // G
  FOR_WHO:    8,   // H
  KAL_CHECK:  9,   // I
  DEST_DRIVE: 10,  // J
  TEMPLATE:   11,  // K
  ABSTRACT:   12,  // L
  OWNER:      13   // M — email address for notifications (optional)
});

const LAST_COL   = COL.ABSTRACT; // rightmost formatted/data column (L)
const DATA_START = 2;            // row 1 = header; data rows begin here

// ── Sheet names ───────────────────────────────────────────────────────────────
const SHEET = Object.freeze({
  REGISTRY: 'Registry',          // main file list  (was "Docs")
  CODES:    'CODES & Descriptions', // validation data (was "Levels")
  SETTINGS: 'Settings',
  TEMP:     'Temp'
});

// ── Row highlight colours (priority model) ────────────────────────────────────
const COLOR = Object.freeze({
  RED:    '#f4cccc',  // 🔴 Priority 1 – structural / naming error
  YELLOW: '#fff2cc',  // 🟡 Priority 2 – valid name, file missing in Drive
  GREEN:  '#d9ead3',  // 🟢 Priority 3 – finalized (vFINAL)
  ERROR:  '#fff2cc',  // amber – script error during batch (same as yellow)
  NONE:   null        // no highlight
});

// ── Header / separator colours ───────────────────────────────────────────────
const HEADER_BLUE  = '#111184';  // deep navy — matches sidebar --blue
const SEPARATOR_RED = '#cc0000'; // red separator rows between drive-code groups

// Drive-code prefix order for rebuildRegistryFromDrive()
// Edit this array to change the grouping order. Unknown codes follow, sorted.
const REBUILD_PREFIX_ORDER = ['OP', 'KAL', 'LP', 'PC'];

// ── KAL logo (Base64-encoded PNG) ────────────────────────────────────────────
// Paste the full base64 string between the quotes.
// Used by showUserGuide() to embed the logo in the sidebar.
const KAL_LOGO_BASE64 = '';
