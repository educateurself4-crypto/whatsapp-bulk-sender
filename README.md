# WhatsApp Bulk Sender: Python desktop app + n8n + Google Sheets

Uses the **official WhatsApp Cloud API**. Marketing messages must use **approved templates** and go only to contacts who **opted in**.

```
Desktop app (Python/Tkinter)  --HTTPS-->  n8n webhooks  -->  WhatsApp Cloud API
   Excel, template, progress                  |
                                              +--> Google Sheets (Messages, Events)
Meta status webhooks (delivered/read/failed) -+
```
The Meta token lives only in n8n. The desktop app only knows your n8n URL and an API key.

## What is in this folder
| File | Purpose |
|---|---|
| `desktop-app/app.py` | The desktop UI |
| `desktop-app/core.py` | Excel reading, validation, n8n client, campaign runner |
| `desktop-app/test_core.py` | Tests for the logic (11 tests) |
| `desktop-app/build_exe.bat` | Builds `WABulkSender.exe` |
| `n8n/01_WA_Send_Campaign.json` | Sends templates, logs to the `Messages` sheet |
| `n8n/02_WA_Status_Webhook.json` | Receives Meta delivery/read/failed events and STOP replies |
| `n8n/03_WA_App_API.json` | Endpoints the app calls: list templates, campaign status, opt-outs |

## Step 1: Meta / WhatsApp setup
1. Create a Meta Business Portfolio and complete **Business Verification**.
2. developers.facebook.com -> create app (type **Business**) -> add **WhatsApp**.
3. Add your sending phone number (it cannot be active on the normal WhatsApp app). Until then you can use Meta's test number, which only messages up to 5 verified recipients.
4. Note the **Phone Number ID** and **WhatsApp Business Account ID (WABA ID)**.
5. Business Settings -> **System Users** -> create one -> assign your app + WABA -> generate a **permanent token** with `whatsapp_business_messaging` and `whatsapp_business_management`.
6. Add a payment method in WhatsApp Manager.
7. Create a **Marketing template** (example: `Hi {{1}}, our {{2}} sale is live! Get {{3}} off till {{4}}.`) with a "Stop" quick-reply button. Wait for **Approved**.

This app supports templates with text body variables `{{1}}, {{2}}...`, static buttons and a text footer. Templates with image/video/document headers, variable URL buttons, coupon codes, carousels or named variables show as "not supported".

## Step 2: Google Sheet
Create one spreadsheet with two tabs. Row 1 must contain **exactly** these headers:

- Tab `Messages`: `campaign_id | phone | wa_message_id | send_status | error | sent_at`
- Tab `Events`: `campaign_id | wa_message_id | phone | status | error_code | error_text | timestamp`

**Important:** select the whole `phone` and `wa_message_id` columns -> Format -> Number -> **Plain text**. Otherwise Sheets turns long numbers into `9.19E+11` and matching breaks.

Copy the spreadsheet ID from its URL (`docs.google.com/spreadsheets/d/<THIS PART>/edit`).

## Step 3: n8n
Use n8n Cloud, or self-host with a real HTTPS domain (Meta will not call plain http).

**A. Fill the placeholders.** Open the three JSON files in a text editor and find/replace:

| Placeholder | Value | Files |
|---|---|---|
| `YOUR_SPREADSHEET_ID` | your sheet ID | 01, 02, 03 |
| `YOUR_PHONE_NUMBER_ID` | Phone Number ID | 01 |
| `YOUR_WABA_ID` | WABA ID | 03 |
| `YOUR_VERIFY_TOKEN` | any random string you invent | 02 |

The Graph API version is set to `v23.0` in the HTTP nodes. Older versions stay supported for a long time, but check Meta's changelog and update it if you like.

**B. Create 3 credentials in n8n**
1. *Header Auth* named **WA n8n Webhook Key**: Name `x-api-key`, Value = a long random secret (this goes in the desktop app too).
2. *Header Auth* named **WhatsApp Cloud Token**: Name `Authorization`, Value `Bearer <your permanent token>`.
3. *Google Sheets OAuth2* named **Google Sheets account**.

**C. Import** each JSON: Workflows -> Import from file. Nodes with a red warning need a credential selected (Webhook nodes -> WA n8n Webhook Key; HTTP nodes -> WhatsApp Cloud Token; Sheets nodes -> Google account). In each Google Sheets node, confirm the spreadsheet and tab show correctly (re-select them if a field looks empty).

**D. Activate all 3 workflows.** Use the **production** URLs (`/webhook/...`), not `/webhook-test/...`.

**E. Connect Meta -> n8n.** In the Meta app -> WhatsApp -> Configuration -> Webhook:
- Callback URL: `https://YOUR-N8N-DOMAIN/webhook/wa-status`
- Verify token: the same string you used for `YOUR_VERIFY_TOKEN`
- Click Verify, then **subscribe to the `messages` field**.

## Step 4: Desktop app
```
cd desktop-app
pip install -r requirements.txt
python app.py
```
(Tkinter ships with the standard Python installer for Windows.) Then:
1. Enter the n8n base URL (e.g. `https://n8n.yourdomain.com`), the API key, and click **Save + Load templates**.
2. Pick a template. The preview shows the text and how many variables it needs.
3. Browse to your Excel/CSV. Pick the phone column and map each `{{n}}` to a column.
4. **Validate contacts.** It removes invalid numbers, duplicates, empty variables and opted-out numbers.
5. **Send 1 test message** to your own number first.
6. Tick the opt-in confirmation, set "Max messages this run", and **Start sending**.

To build an .exe: run `build_exe.bat`.

## How sending works
The app sends 50 contacts at a time to n8n and waits until n8n has logged results for all 50 before sending the next chunk. n8n sends one message every 1.5 seconds (adjustable in the "Send via WhatsApp API" node -> Options -> Batching). Delivery and read status arrive later through Meta's webhook.

## Limits and things to know
- **Daily limit:** new numbers start at about 250 unique recipients per 24 hours and increase with quality and verification. Check WhatsApp Manager and set "Max messages this run" accordingly.
- **Per-user marketing cap:** Meta limits marketing templates per user across all businesses (error 131049). These show up as failures and cannot be bypassed.
- **Cost:** you pay per delivered marketing message, by recipient country. Check Meta's current rate card.
- **Google Sheets quotas:** about 60 reads and 60 writes per minute per user. This design keeps usage low (one write per chunk, plus status events), which suits a few thousand messages per day. For much larger volumes, replace the Sheets nodes with Postgres or Supabase.
- **"Accepted" is not "delivered."** Check the delivered/failed counts after a few minutes.
- **Stop button** stops new chunks only. The chunk already handed to n8n (up to 50 messages) still finishes.
- **If a chunk times out, do not resend** before checking the `Messages` sheet, or people will get duplicates.
- The status webhook (POST) is public. Meta's signature check (`X-Hub-Signature-256`) is not implemented. Add it before heavy production use.
- Phone checking is format-only. It cannot tell whether a number is on WhatsApp (those come back as failures).
- The API key is saved in `~/.wa_bulk_sender/settings.json` in plain text.

## What was tested
- Python logic (Excel/CSV reading, phone validation, template parsing, chunking, cancel, timeout) passes 11 automated tests against a mock n8n server: `python -m unittest test_core -v`.
- The JavaScript in all n8n Code nodes was run against sample Meta payloads (success, error responses, delivered/failed statuses, STOP replies, empty sheets).
- **Not tested:** the Tkinter window itself and a live n8n/Meta/Google connection. Those need your accounts, so use the test-message step first.
