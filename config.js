/* ============================================================
   Two Chimps customer lookup - SETTINGS
   This is the only file you should need to change.
   Columns are found by their HEADER NAME, not their letter, so you can
   add, move or insert columns in the spreadsheet without changing this.
   ============================================================ */
window.TCC_CONFIG = {
  // From the Entra app registration
  clientId: "32360b19-8075-4d21-ba39-034d56514803",
  tenantId: "7513cbbc-9f58-498a-be21-b7c12713405d",

  // The spreadsheet the panel reads and edits.
  // TESTING: "on boarding updated CLAUDE TEST.xlsx" (Two Chimps Coffee site > Documents)
  // GOING LIVE: Claude will give you the live file's two IDs to swap in here.
  driveId: "b!66b-DgKPbUaA7ma0wCDb7qWYb636lyxAsGXW7dSgiwgYQHIYNrRLR4W8vzjHycI7",
  itemId: "014EVWT2SXYZTWG3COCNG2JLBBFB2OYHUW",
  // Backup: the sharing link, only used if the two IDs above are blank
  workbookShareUrl: "https://twochimpscoffee.sharepoint.com/:x:/r/sites/TwoChimpsCoffee/Shared%20Documents/on%20boarding%20updated%20CLAUDE%20TEST.xlsx?d=w6367c6574e6c4d13a4ac212874ec1e96&csf=1&web=1&e=bijphm",
  sheetName: "Sheet1",

  headerRows: 2,            // rows 1-2 are headers, customers start on row 3

  // Not real email addresses - match these customers on name only
  placeholderEmails: ["noths"],

  // Tick marks: "a" in Webdings looks like a tick
  tickValue: "a",
  tickFont: "Webdings",

  // The main fields, found by header text (row 1 or row 2), shown in this order.
  // type: text | longtext | date | select | tick
  coreFields: [
    { key: "name",      header: "Name",            label: "Name",            type: "text" },
    { key: "email",     header: "Email address",   label: "Email",           type: "text" },
    { key: "phone",     header: "Phone",           label: "Phone",           type: "text" },
    { key: "frequency", header: "Frequency",       label: "Frequency",       type: "select",
      options: ["Weekly", "Fortnightly", "Monthly", "Once", "Sample", "Cancelled", "Gift Subs", "NOTHS Subs"] },
    { key: "delivery",  header: "Hand Delivery Day", label: "Hand delivery day", type: "select", options: ["Wed", "Thur"] },
    { key: "posted",    header: "Posted date",     label: "1st order posted", type: "date" },
    { key: "lastOrder", header: "Last Order Date", label: "Last order date", type: "date" },
    { key: "discount",  header: "Discount Code",   label: "Discount code",   type: "text" },
    { key: "flavour",   header: "Flavour Wheel Leaflet Sent?", label: "Flavour wheel leaflet sent", type: "tick" },
    { key: "course",    header: "Course Leaflet Sent", label: "Course leaflet sent", type: "tick" },
    { key: "notes",     header: "Notes",           label: "Notes",           type: "longtext" }
  ],

  // Every OTHER column with a header is shown automatically under "More details",
  // as a tick box, date or text (worked out from the column). To force a type,
  // add the header here, e.g.  "xmas brochure": "tick"
  fieldTypes: {},

  // Columns to leave out of the panel (by header text)
  hideHeaders: [],

  // Helper columns that hold the dropdown lists - never shown or edited
  ignoreHeaders: ["Weekly", "Fortnightly", "Delivery Day"],

  // Which emails the panel looks up when it's pinned open. Anything else just shows "Not an order email".
  // An email counts as an order if it's from one of these senders, or its subject matches one of these patterns.
  orderEmailSenders: ["noreply@twochimpscoffee.com"],
  orderEmailSubjects: ["\\[Two Chimps Coffee\\].*order"],

  // "Due now" shows phases due in the last this-many days; older unticked ones are tucked away.
  // Phases are found automatically: a formula column (or two, e.g. "Can we send?")
  // followed by the column where it's ticked off.
  dueWindowDays: 60
};
