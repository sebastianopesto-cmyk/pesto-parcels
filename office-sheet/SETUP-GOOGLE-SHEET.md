# Office sheet setup (about 5 minutes, free)

1. Go to sheets.google.com and create a new blank sheet. Name it "Pesto Parcels".
2. Menu **Extensions → Apps Script**. Delete what's there and paste the whole of `Code.gs`. Click the save icon.
3. Left side, gear icon **Project Settings → Script properties → Add script property**:
   name `TEAM_CODE`, value: a short code you choose (for example a word + 4 digits).
   Tell colleagues this code in person or by phone, never in a public chat.
4. Top right **Deploy → New deployment**. Gear next to "Select type" → **Web app**.
   - Execute as: **Me**
   - Who has access: **Anyone**
   Click **Deploy**, allow the permissions Google asks for (it's your own script).
5. Copy the **Web app URL** (ends in `/exec`) and send it to Claude in the thread
   (this link is fine to share; the team code is the lock).
6. On each phone: open the app, tap the grey status line at the bottom, fill in team code and your name. Done.

Every scan becomes a new row in the "Parcels" tab of the sheet. The app's Office tab reads the same rows.

## Update (after a new version of Code.gs)

The web address stays the same, so phones don't need anything new.

1. Open the sheet, **Extensions → Apps Script**. Select all the code in `Code.gs`, delete it, paste the new `Code.gs`, click save.
2. **Deploy → Manage deployments** → pencil icon (Edit) → Version: **New version** → **Deploy**.

Do not use "New deployment" for an update: that makes a new web address.
