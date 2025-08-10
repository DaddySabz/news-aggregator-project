/**
 * Daily maintenance:
 * 1. Sort Feed!A:I by publish_date (column A) newest‑first
 * 2. Move rows older than 24 h to Archive sheet
 */
function sortAndArchiveFeed() {
  const ss     = SpreadsheetApp.openById('10Ebl6jdFkXtDSQUMM1wO9GhQx5QZb8pE3TkRHgO6hdQ');   // ← replace
  const feed   = ss.getSheetByName('Feed');
  const archive= ss.getSheetByName('Archive');
  if (!feed || !archive) return;

  const lastRow = feed.getLastRow();
  if (lastRow < 3) return;                       // nothing to do

  // --- 1) Sort newest‑first (keeps header row)
  feed.getRange(2, 1, lastRow - 1, 9)
      .sort({ column: 1, ascending: false });

  // --- 2) Archive rows older than 24 h
  const cutoff = Date.now() - 24 * 60 * 60 * 1000;
  const data   = feed.getRange(2, 1, lastRow - 1, 9).getValues();

  const toArchive = [];
  const keep      = [];

  for (const row of data) {
    const ts = Date.parse(row[0]);               // publish_date in col A
    (isNaN(ts) || ts < cutoff) ? toArchive.push(row) : keep.push(row);
  }

  if (toArchive.length) {
    // Add header to Archive if empty
    if (archive.getLastRow() === 0) {
      const header = feed.getRange(1, 1, 1, 9).getValues();
      archive.appendRow(header[0]);
    }
    archive.getRange(archive.getLastRow()+1,1,toArchive.length,9)
           .setValues(toArchive);
  }

  // Rewrite Feed (header + rows to keep)
  feed.clearContents();
  const header = ['Publish date','Scrape date','Headline','Link','Source','Authors','Source Icon','Thumbnail','Snippet'];
  feed.appendRow(header);
  if (keep.length) feed.getRange(2,1,keep.length,9).setValues(keep);
}
