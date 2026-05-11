/**
 * KAL ACADEMY - Global constants shared across all .gs files.
 * Declared here once to avoid "Identifier already declared" errors
 * when multiple script files are present in the same project.
 */

// ── Column indices (1-based) ──────────────────────────────────────────────────
// Keep in sync with the Docs sheet layout.
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
  ABSTRACT:   12   // L
});

const LAST_COL   = COL.ABSTRACT; // rightmost data column index
const DATA_START = 2;            // row 1 = header; data rows begin here

// ── KAL logo (Base64-encoded PNG) ────────────────────────────────────────────
// Paste the full base64 string between the quotes below.
// Used by showUserGuide() to render the logo in the sidebar.
const KAL_LOGO_BASE64 = '';
