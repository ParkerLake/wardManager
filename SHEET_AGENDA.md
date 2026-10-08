# Sheet-based Ward Council agenda items

Off by default. In `src/config.js` (local file, not in git): `WC_AGENDA_SOURCE: "app"` = today's behavior,
`"sheet"` = discussion topics are read from / written to the **Agenda** tab. Rollback = set back to `"app"` and `npm run deploy`.
Old topic rows in `WardCouncilMeeting` are never deleted.

## 1. Before you start
In the Ward Council spreadsheet: **File > Version history > Name current version** ("before Agenda tab").

## 2. Create the tab
Add a tab named exactly `Agenda`. Each meeting is a 6-row block; the newest block is always last and nothing goes below it.
Rows below are for a block starting at row 1 (copy/paste the whole block downward for each new week). No helper columns.

| Cell | Contents |
|---|---|
| A1 | the date as a real date (type `10/11/2026`), then Format > Number > Custom date: `dddd mmmm d` so it reads "Sunday October 11" |
| A2 / B2 | `Hymn` / formula below with `opening_song` |
| A3 / B3 | `Opening Prayer` / formula with `opening_prayer` |
| A4 / B4 | `Thought/Handbook` / formula with `spirit_thought` |
| A5 / B5 | `Closing Prayer` / formula with `closing_prayer` |
| A6 | `Agenda Items` |
| A7... | one topic per row. Type `Done` in column B to mark one complete. |

The app reads the date from what column A displays ("Sunday October 11"; it works out the year from the weekday).
Typed text works for the app too, but the assignment formulas need a real date.

Assignment formula (B2; change the item key for each row; `$A1` is that block's date cell, relative so it follows when you copy the block):

```
=IFERROR(INDEX(FILTER(WardCouncilMeeting!$C$2:$C, ARRAYFORMULA(TEXT(WardCouncilMeeting!$A$2:$A,"yyyy-mm-dd"))=TEXT($A$1,"yyyy-mm-dd"), WardCouncilMeeting!$B$2:$B="opening_song"),1),"")
```
Use `$A1` in place of `$A$1` in all four formulas before you copy the block down.

## 3. New week
Copy the last block, paste it directly below, change the date. Delete finished items, keep the rest (that is the carry-over).

## 4. One-time import of existing topics (optional)
Apps Script (Extensions > Apps Script), run `importTopics` once with the latest block already created and its date a real date:

```js
function importTopics() {
  const ss = SpreadsheetApp.getActive();
  const ag = ss.getSheetByName('Agenda');
  const a = ag.getRange(1, 1, ag.getLastRow(), 1).getValues().map(r => r[0]);
  let ai = -1; a.forEach((v, i) => { if (String(v).trim().toLowerCase() === 'agenda items') ai = i; });
  if (ai < 5) throw new Error('No block found');
  const d = a[ai - 5];                       // date cell is 5 rows above "Agenda Items"
  if (!(d instanceof Date)) throw new Error('Make the block date a real date first');
  const date = Utilities.formatDate(d, ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  const rows = ss.getSheetByName('WardCouncilMeeting').getDataRange().getValues().slice(1)
    .filter(r => String(r[0]) === date && String(r[1]).indexOf('topic_') === 0)
    .sort((x, y) => (Number(x[7]) || 0) - (Number(y[7]) || 0));
  if (!rows.length) return;
  ag.getRange(ai + 2, 1, rows.length, 2).setValues(rows.map(r => [r[4], String(r[3]).toLowerCase() === 'true' ? 'Done' : '']));
}
```

## 5. Show only the current week (optional)
Same Apps Script project:

```js
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Agenda')
    .addItem('Show current week only', 'showCurrentWeek').addItem('Show all weeks', 'showAll').addToUi();
}
function showCurrentWeek() {
  const sh = SpreadsheetApp.getActive().getSheetByName('Agenda');
  const a = sh.getRange(1, 1, sh.getLastRow(), 1).getValues().map(r => String(r[0]).trim().toLowerCase());
  const ai = a.lastIndexOf('agenda items');          // 0-based; header is 5 rows above
  sh.showRows(1, sh.getMaxRows());
  if (ai > 5) sh.hideRows(1, ai - 5);                // hides everything above the last block's date row
}
function showAll() { const sh = SpreadsheetApp.getActive().getSheetByName('Agenda'); sh.showRows(1, sh.getMaxRows()); }
```

## 6. Turn it on
1. Set `WC_AGENDA_SOURCE: "sheet"` in `src/config.js`, run `npm run deploy`.
2. Check the Ward Council tab shows the topics from the sheet; add one in the app and confirm it appears in the sheet.

## Rules
- Nothing below the last block's items. Keep each block's shape (date row, 4 assignment rows, `Agenda Items` row, then items) and don't type over the assignment formulas.
- Typing an item that looks like a date (e.g. `Oct 25`) is fine.
- The app can edit only the **latest** week's items; older weeks are edited in the sheet.
- Assignments are read-only in the sheet; change them in the app.
- If an app save fails with an Agenda error, reload the page.
