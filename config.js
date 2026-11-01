/* ============================================================
   Two Chimps customer lookup - SETTINGS
   This is the only file you should need to change.
   ============================================================ */
window.TCC_CONFIG = {
  // From the Entra app registration ("Application (client) ID") - step 2 of the setup guide
  clientId: "PASTE-CLIENT-ID-HERE",
  // Your Microsoft 365 tenant ("Directory (tenant) ID") - same page
  tenantId: "PASTE-TENANT-ID-HERE",

  // Sharing link to the spreadsheet (Excel > Share > Copy link).
  // TESTING: the CLAUDE TEST copy.  GOING LIVE: swap in the live file's link.
  workbookShareUrl: "PASTE-SHARE-LINK-TO-CLAUDE-TEST-COPY-HERE",
  sheetName: "Sheet1",

  firstDataRow: 3,          // customers start on row 3 (rows 1-2 are headers)
  lastFormulaRow: 13048,    // formulas are filled down to here
  lastColumn: "CD",

  // Not real email addresses - match these customers on name only
  placeholderEmails: ["noths"],

  // Tick marks: "a" in Webdings looks like a tick
  tickValue: "a",
  tickFont: "Webdings",

  // Fields shown in the panel. type: text | longtext | date | select | tick
  fields: [
    { col: "A",  label: "Name",            type: "text",  key: "name" },
    { col: "B",  label: "Email",           type: "text",  key: "email" },
    { col: "N",  label: "Phone",           type: "text",  key: "phone" },
    { col: "U",  label: "Frequency",       type: "select",
      options: ["Weekly", "Fortnightly", "Monthly", "Once", "Sample", "Cancelled", "Gift Subs", "NOTHS Subs"] },
    { col: "Q",  label: "Hand delivery day", type: "select", options: ["Wed", "Thur"] },
    { col: "R",  label: "1st order posted", type: "date" },
    { col: "S",  label: "Last order date", type: "date" },
    { col: "T",  label: "Discount code",   type: "text" },
    { col: "M",  label: "Flavour wheel leaflet sent", type: "tick" },
    { col: "O",  label: "Course leaflet sent", type: "tick" },
    { col: "P",  label: "Notes",           type: "longtext" }
  ],

  // "Due now" shows phases due in the last this-many days; older unticked ones are tucked away
  dueWindowDays: 60,

  // Phases: due = formula column (date, or "yes"), done = tick column,
  // canSend = optional formula column; "no" there means not due yet.
  phases: [
    { due: "AC", done: "AD" }, { due: "AE", done: "AF" }, { due: "AG", done: "AH" },
    { due: "AI", done: "AJ" }, { due: "AK", done: "AL" }, { due: "AM", done: "AO", canSend: "AN" },
    { due: "AP", done: "AQ" }, { due: "AR", done: "AS" }, { due: "AT", done: "AV", canSend: "AU" },
    { due: "AW", done: "AY", canSend: "AX" }, { due: "AZ", done: "BA" },
    { due: "BB", done: "BD", canSend: "BC" }, { due: "BG", done: "BH" }, { due: "BI", done: "BJ" },
    { due: "BK", done: "BL" }, { due: "BM", done: "BN" }, { due: "BO", done: "BP" },
    { due: "BQ", done: "BR" }, { due: "BS", done: "BT" }, { due: "BU", done: "BV" },
    { due: "BW", done: "BX" }, { due: "BY", done: "BZ" }, { due: "CA", done: "CB" },
    { due: "CC", done: "CD" }
  ]
};
