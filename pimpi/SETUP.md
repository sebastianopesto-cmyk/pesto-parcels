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

## Offices (one login per office)

`TEAM_CODE` is the head-office code: it sees every office. To add an office, log in with it, open **Tariffe → Uffici e accessi**, type the office name and city, press **Aggiungi ufficio**. The app creates an 8-character access code; send it to that office by voice or WhatsApp.

- Each office sees and changes only its own parcels, and can have its own rates. Yachts are shared by all offices.
- Parcels saved before the first office existed belong to the first office you add (add Genova first).
- **Nuovo codice** replaces an office's code (the old one stops working). **Disattiva** blocks that office.
- The harbour map on the Live page is drawn for Genoa, so it shows only for offices whose city is Genova.

## Paid add-ons (optional, keys go only in Script properties)

**AI label reading (Claude).** console.anthropic.com → sign up → Billing: add prepaid credit (start with $5) → API keys → Create key. Add script property `ANTHROPIC_API_KEY` = the key. Default model is `claude-opus-5-5` (most accurate, about 1–3 cents per label). For about 20x cheaper set `AI_MODEL` = `claude-haiku-5-5`. If the key is missing or a read fails, the app uses the free reader as before.

**Live courier status (17TRACK).** 17track.net → API → sign up (100 free tracking numbers to test). Copy the API key, add script property `TRACK17_KEY`. Then **Triggers (clock icon) → Add trigger → `refreshTracking` → Time-driven → Hour timer → Every hour**. The Live page also has an **Aggiorna stato** button.

After adding keys: **Deploy → Manage deployments → New version** (keys are read at run time, but new code needs a new version).
