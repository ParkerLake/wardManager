# Sheet-based Ward Council agenda items

Off by default. In `src/config.js` (local file, not in git): `WC_AGENDA_SOURCE: "app"` = today's behavior,
`"sheet"` = discussion topics are read from / written to the **Agenda** tab. Rollback = set back to `"app"` and `npm run deploy`.
Old topic rows in `WardCouncilMeeting` are never deleted.

## 1. Before you start
In the Ward Council spreadsheet: **File > Version history > Name current version** ("before Agenda tab").

## 2. Create the tab
Add a tab named exactly `Agenda`. Each meeting is a 6-row block; the newest block is always last and nothing goes below it.
Rows below are for a block starting at row 1 (copy/paste the whole block downward for each new week).

| Cell | Contents |
|---|---|
| A1 | the date as a real date, format Format > Number > Custom: `dddd mmmm d` |
| C1 | `=TEXT(A1,"yyyy-mm-dd")`  (the app finds blocks by this cell; **hide column C**) |
| A2 / B2 | `Hymn` / formula below with `opening_song` |
| A3 / B3 | `Opening Prayer` / formula with `opening_prayer` |
| A4 / B4 | `Thought/Handbook` / formula with `spirit_thought` |
| A5 / B5 | `Closing Prayer` / formula with `closing_prayer` |
| A6 | `Agenda Items` |
| A7... | one topic per row. Type `Done` in column B to mark one complete. |

Assignment formula (B2; change the item key for each row; `$C1` must point at that block's date row):

```
=IFERROR(INDEX(FILTER(WardCouncilMeeting!$C$2:$C, ARRAYFORMULA(TEXT(WardCouncilMeeting!$A$2:$A,"yyyy-mm-dd"))=$C$1, WardCouncilMeeting!$B$2:$B="opening_song"),1),"")
```
When you copy a block down, change `$C$1` to a relative row (`$C1`) in all four formulas first so they follow the block.

## 3. New week
Copy the last block, paste it directly below, change the date. Delete finished items, keep the rest (that is the carry-over).

## 4. One-time import of existing topics (optional)
Apps Script (Extensions > Apps Script), run `importTopics` once with the latest block already created:

```js
function importTopics() {
  const ss = SpreadsheetApp.getActive();
  const ag = ss.getSheetByName('Agenda');
  const col = ag.getRange('C1:C' + ag.getLastRow()).getValues().map(r => String(r[0]));
  let hdr = -1; col.forEach((v, i) => { if (/^\d{4}-\d{2}-\d{2}$/.test(v)) hdr = i; });
  if (hdr < 0) throw new Error('No block found');
  const date = col[hdr];
  const a = ag.getRange(hdr + 1, 1, ag.getLastRow() - hdr, 1).getValues().map(r => String(r[0]).toLowerCase());
  const items = hdr + 1 + a.indexOf('agenda items') + 1; // 0-based row index of first item
  const rows = ss.getSheetByName('WardCouncilMeeting').getDataRange().getValues().slice(1)
    .filter(r => String(r[0]) === date && String(r[1]).indexOf('topic_') === 0)
    .sort((x, y) => (Number(x[7]) || 0) - (Number(y[7]) || 0));
  if (!rows.length) return;
  ag.getRange(items + 1, 1, rows.length, 2).setValues(rows.map(r => [r[4], String(r[3]).toLowerCase() === 'true' ? 'Done' : '']));
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
  const col = sh.getRange('C1:C' + sh.getLastRow()).getValues().map(r => String(r[0]));
  let hdr = 0; col.forEach((v, i) => { if (/^\d{4}-\d{2}-\d{2}$/.test(v)) hdr = i + 1; });
  sh.showRows(1, sh.getMaxRows());
  if (hdr > 1) sh.hideRows(1, hdr - 1);
}
function showAll() { const sh = SpreadsheetApp.getActive().getSheetByName('Agenda'); sh.showRows(1, sh.getMaxRows()); }
```

## 6. Turn it on
1. Set `WC_AGENDA_SOURCE: "sheet"` in `src/config.js`, run `npm run deploy`.
2. Check the Ward Council tab shows the topics from the sheet; add one in the app and confirm it appears in the sheet.

## Rules
- Nothing below the last block's items. Don't edit date/C cells or assignment formulas by hand except when copying a block.
- The app can edit only the **latest** week's items; older weeks are edited in the sheet.
- Assignments are read-only in the sheet; change them in the app.
- If an app save fails with an Agenda error, reload the page.
