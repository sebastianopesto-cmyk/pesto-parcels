# PIMPI on Google Sheets: setup (about 5 minutes, free)

Use a **new** Google Sheet: the scanner's sheet keeps working as it is.

1. Go to sheets.google.com, create a blank sheet, name it "PIMPI".
2. **Extensions → Apps Script**. Delete what's there, paste all of `pimpi/Code.gs`, click save.
3. Gear icon **Project Settings → Script properties → Add script property**: `TEAM_CODE` = a short code you choose.
   Tell colleagues in person, never in a chat.
4. **Deploy → New deployment** → gear → **Web app**. Execute as **Me**, who has access **Anyone**. Deploy.
   Google asks for permission twice (sheet and Drive, because label photos are saved in a Drive folder "PIMPI foto"). Allow.
5. Send the **Web app URL** (ends in `/exec`) in the thread. The team code stays with you.

Update later: **Deploy → Manage deployments → pencil → Version: New version → Deploy**. The link stays the same.
