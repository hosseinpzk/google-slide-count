# Google-Slide-Count
Let's say you have a Google Excel sheet, and in column I, you have a link to Google Slides. You want to write the number of slides in front of that. This Google App Script helps you well :)

Google Apps Script to count the number of slides in Google Slides and show it in your Google Sheet.

## Setup
1. Open your Google Sheet → **Extensions → Apps Script**.
2. Paste the code from this repo.
3. Configure:
   - `COL_LINK` → the column where your slide URLs are (default: **I = 9**).
   - `COL_COUNT` → the column where the number of slides will be written (default: **J = 10**)
4. Save & refresh the sheet.

## Usage
- A new menu **Slides Tools** appears.  
- Click **Slides Tools → Count Slides (I→J)**.  
- Script fills column J with slide counts.

## Notes
- Works only with Google Slides links (`/presentation/d/...` or `/u/0/d/...`).  
- If no access → `⚠️ No access`.  
- If not a Slides file → `Not Slides`.

<img width="749" height="113" alt="image" src="https://github.com/user-attachments/assets/2c9472db-ca60-4f52-b025-59e7543af866" />
