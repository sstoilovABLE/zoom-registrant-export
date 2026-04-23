# Export Unique Join Links for Zoom Meeting Registrants

A PowerShell script that exports all registered participants from a Zoom meeting with their **unique personal join links** to a timestamped CSV file. Uses the Zoom Server-to-Server OAuth API with automatic pagination to handle any number of registrants.

> 🤖 **Disclaimer:** All code in this repository was vibe coded with AI and tested by a human (me 🧔‍♂️). It works, but treat it accordingly — review before using in production environments, and contribute improvements if you spot them.

***

## Who Is This For?

Anyone who:

- Organizes Zoom meetings with registration enabled
- Needs to extract unique join links for all registrants in bulk (e.g. to send via email/mail merge, import into a CRM, or log for compliance)
- Wants a no-fuss, no-install solution on Windows — no Python, no third-party tools required

***

## Prerequisites

- Windows with PowerShell - the script works with the modern PowerShell 7, as well as with the built-in Windows PowerShell
- A Zoom account with **Server-to-Server OAuth** app credentials (Pro plan or higher)
- The meeting must have **registration enabled**

### Getting Your Zoom API Credentials

1. Go to [marketplace.zoom.us](https://marketplace.zoom.us) → **Develop → Build App**
2. Select **Server-to-Server OAuth**
3. Add the scope: `meeting:read:admin`
4. Note down your **Account ID**, **Client ID**, and **Client Secret**
> ⚠️ Never share these credentials! Anyone who has them can use them to access your Zoom account.

***

## How to Use

### 1. Download the script

Clone this repo or download `GetZoomLinks.ps1` directly. 

### 2. Allow PowerShell scripts to run (one-time)

Open **PowerShell as Administrator**. You can do this multiple ways: 
* Pressing **Win+X**, selecting **Terminal (Admin)**, and then in the Command Prompt enter the command `pwsh` (PowerShell 7) or `powershell` (Windows PowerShell)
* Opening the Start menu, searching for "PowerShell", and then selecting "Run as Administrator" underneath the option that appears (PowerShell 7 or Windows PowerShell)

Then run:

```powershell
Set-ExecutionPolicy RemoteSigned
```

Type `Y` to confirm. You only need to do this once.

### 3. Run the script

```powershell
.\GetZoomLinks.ps1
```

### 4. First run — enter your credentials

On first run, the script will prompt you for your Zoom API credentials:

```
=== First Run Setup ===
Please enter your Zoom credentials:

Account ID: xxxxxxxxxxxxxxxx
Client ID: xxxxxxxxxxxxxxxx
Client Secret: ****************
```

Credentials are saved locally to `zoom-secrets.json` in the same folder. The script will also offer to create or update a `.gitignore` to ensure your secrets file and exported CSVs are never accidentally committed to version control.

### 5. Enter your Meeting ID

```
Enter Meeting ID: 123456789
```

### 6. Done

The script fetches all pages of registrants automatically and saves the results to a timestamped CSV in the same directory. The output filename is in the format `yyyymmdd_hhmmss_meetingid_registrants.csv`. 

```
Done! 450 registrants saved to 20260424_014500_123456789_registrants.csv
```

### Output Format

The CSV contains the following columns:

| Column | Description |
|---|---|
| `email` | Registrant's email address |
| `first_name` | First name |
| `last_name` | Last name |
| `join_url` | Unique personal join link for this registrant |

***

## How Credentials Are Stored

On first run, your credentials are saved to `zoom-secrets.json` in the script directory:

```json
{
  "account_id": "...",
  "client_id": "...",
  "client_secret": "..."
}
```

On subsequent runs, the script loads credentials from this file automatically — no need to re-enter them.

**The script will never commit this file** — it adds `zoom-secrets.json` to `.gitignore` automatically on first run.

> 🔒 Keep `zoom-secrets.json` out of any shared or synced folders. Treat your Client Secret like a password.

***

## Pagination

The Zoom API returns a maximum of 300 registrants per request. This script handles pagination automatically using `next_page_token` — it will loop through all pages until every registrant is retrieved, regardless of total count.

***

## License

MIT — free to use, modify, and redistribute. See [LICENSE](LICENSE).

***

## Reuse, Fork & Contribute

Found this useful? Feel free to fork it, adapt it for your own needs, or extend it (webinar support, additional CSV fields, email sending, n8n integration — the possibilities are there).

Pull requests are welcome. Whether it's a bug fix, a new feature, or just tidying up the vibe-coded corners — contributions of all sizes are appreciated.

If this saved you time, a ⭐ on the repo is always welcome.