//*******COLUMN Setting****//
const COL_LINK = 9;   // I
const COL_COUNT = 10; // J
const START_ROW = 2;  // if header is in row 1

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Slides Tools')
    .addItem('Count Slides (I→J)', 'countSlidesInColumn')
    .addToUi();
}

function countSlidesInColumn() {
  const sh = SpreadsheetApp.getActiveSheet();
  const lastRow = sh.getLastRow();
  if (lastRow < START_ROW) return;

  const linkRange = sh.getRange(START_ROW, COL_LINK, lastRow - START_ROW + 1, 1);
  const linkVals = linkRange.getValues();
  const richVals = linkRange.getRichTextValues();

  const out = [];
  for (let r = 0; r < linkVals.length; r++) {
    let url = null;

    // 1) If it has RichText
    try {
      url = richVals[r][0] && richVals[r][0].getLinkUrl();
    } catch (e) { /* نادیده بگیر */ }

    // 2) if URL is in cell as text
    if (!url) {
      const t = (linkVals[r][0] || '').toString().trim();
      if (t.startsWith('http')) url = t;
    }

    // 3) Not Found
    if (!url) {
      out.push(['']);
      continue;
    }

    // 4) extract presentation id 
    const id = extractDriveId(url);
    if (!id) {
      out.push(['-']); // link not valid
      continue;
    }

    // 5) counts slides
    let count = '';
    try {
      // first tun needs access
      const pres = SlidesApp.openById(id);
      count = pres.getSlides().length;
    } catch (err) {
      count = '⚠️';
    }

    out.push([count]);
    Utilities.sleep(50); // a bit delay for limitations
  }

  sh.getRange(START_ROW, COL_COUNT, out.length, 1).setValues(out);
}

// EXTRACTING DRIVE ID
function extractDriveId(url) {
  if (!url) return null;
  const patterns = [
    /\/presentation\/u\/\d+\/d\/([a-zA-Z0-9_-]{10,})\b/,  // حالت u/0/d/
    /\/presentation\/d\/([a-zA-Z0-9_-]{10,})\b/,          // حالت ساده d/
    /\/file\/d\/([a-zA-Z0-9_-]{10,})\b/,
    /[?&]id=([a-zA-Z0-9_-]{10,})\b/,
    /https:\/\/drive\.google\.com\/open\?id=([a-zA-Z0-9_-]{10,})\b/
  ];
  for (const re of patterns) {
    const m = url.match(re);
    if (m && m[1]) return m[1];
  }
  return null;
}
