# EmailVerify Pro - Streamlit Guide

## Files To Create

Create these files:

```text
streamlit_app.py
verify_emails.py
requirements.txt
```

Run:

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

Open:

```text
http://localhost:8501
```

## App Pages

| Page | Use |
| --- | --- |
| Setup | Connect Google Sheets or upload Excel |
| Dashboard | See counts and verify emails 1 by 1 |
| Instructions | New-user guide and direct Google links |

## Columns

Default columns:

| Column | Meaning |
| --- | --- |
| C | Email input |
| W | Valid / Invalid output |
| X | Score / reason output |

Change these in Setup if your file uses different columns.

## Direct Google Links

- Create project: https://console.cloud.google.com/projectcreate
- Google Cloud Console: https://console.cloud.google.com
- Enable Google Sheets API: https://console.cloud.google.com/apis/library/sheets.googleapis.com
- Enable Google Drive API: https://console.cloud.google.com/apis/library/drive.googleapis.com
- Create service account: https://console.cloud.google.com/iam-admin/serviceaccounts/create
- Service accounts page: https://console.cloud.google.com/iam-admin/serviceaccounts
- Google Sheets: https://sheets.google.com

## Google Sheet Setup

1. Go to **Create project**.
2. Create or select a Google Cloud project.
3. Open **Enable Google Sheets API** and click **Enable**.
4. Open **Enable Google Drive API** and click **Enable**.
5. Open **Create service account**.
6. Name it something simple, like:

```text
email-verify
```

7. Finish service account creation.
8. Open the service account.
9. Go to **Keys**.
10. Click **Add key**.
11. Click **Create new key**.
12. Choose **JSON**.
13. Download the JSON.

Google only lets you download this private JSON once. Keep it private.

## Share Your Sheet With The JSON Email

Open the downloaded JSON and find:

```json
"client_email": "email-verify@your-project.iam.gserviceaccount.com"
```

Copy only the email value.

Then:

1. Open your Google Sheet.
2. Click **Share**.
3. Paste the `client_email`.
4. Give **Editor** permission.
5. Click **Send**.

## What To Paste In Setup

For Google Sheets:

```text
Source: Google Sheets
Spreadsheet URL: paste your Google Sheet URL
Worksheet / tab name: your sheet's tab name
Google service account JSON: paste full downloaded JSON
Email column: C
Result column: W
Score / reason column: X
Start row: 2
End row: 10
Delay after each email: 8 to 15
```

You can tick:

```text
Save this JSON on this computer
```

That saves it locally in:

```text
.email_verify_app/google_credentials.json
```

Use **Forget saved JSON** if you want to remove it.

## Dashboard

The dashboard shows:

| Metric | Meaning |
| --- | --- |
| Rows selected | Total rows in your selected range |
| Emails waiting | Rows with email and no result yet |
| Unique emails | Emails to check after duplicate removal |
| Already done | Rows already filled in result column |
| Blank emails | Rows with no email |
| Happened now | Rows verified in this app session |

Verification runs **1 by 1**.

Batch mode is shown as:

```text
Coming soon
```

## Excel Option

1. Open Setup.
2. Choose **Excel**.
3. Upload `.xlsx`.
4. Leave worksheet blank to use the first sheet, or type the exact sheet name.
5. Keep columns as `C`, `W`, `X`.
6. Click **Save and open dashboard**.
7. Verify 1 by 1.
8. Download the verified Excel file.

## Safe Speed

Very gentle:

```text
Delay after each email: 15 to 30 seconds
```

Normal safe:

```text
Delay after each email: 8 to 15 seconds
```

## Fix Proxy Error

If you see:

```text
HTTPSConnectionPool(host='oauth2.googleapis.com', port=443)
ProxyError
127.0.0.1:9
```

Run in PowerShell:

```powershell
Remove-Item Env:HTTP_PROXY -ErrorAction SilentlyContinue
Remove-Item Env:HTTPS_PROXY -ErrorAction SilentlyContinue
Remove-Item Env:ALL_PROXY -ErrorAction SilentlyContinue
streamlit run streamlit_app.py
```
