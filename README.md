# Daily Google Drive Excel Mailer

This repository now includes an automated Python workflow that:

1. Reads the latest Excel file from a Google Drive folder
2. Processes the Excel content (sheet/row/column summary)
3. Sends the output by email through Gmail
4. Runs daily using GitHub Actions

## Files

- `/home/runner/work/MadhabNr/MadhabNr/scripts/daily_drive_mailer.py` – main automation script
- `/home/runner/work/MadhabNr/MadhabNr/.github/workflows/daily-drive-mail.yml` – scheduled job
- `/home/runner/work/MadhabNr/MadhabNr/requirements.txt` – dependencies

## Google setup

1. Create a Google Cloud project.
2. Enable **Google Drive API**.
3. Create a **Service Account** and download `credential.json`.
4. Share the target Drive folder with the service-account email.
5. In Gmail, enable 2FA and generate an **App Password** for SMTP sending.

## Required GitHub Secrets

Set these repository secrets:

- `GOOGLE_SERVICE_ACCOUNT_JSON` (full JSON content from `credential.json`)
- `DRIVE_FOLDER_ID`
- `RECIPIENTS` (comma-separated emails)
- `SMTP_EMAIL` (your Gmail address)
- `SMTP_APP_PASSWORD` (Gmail app password)

Optional:

- `MAIL_SUBJECT`

## Schedule configuration

Workflow file: `/home/runner/work/MadhabNr/MadhabNr/.github/workflows/daily-drive-mail.yml`

- Current cron: `30 3 * * *` (UTC)
- GitHub Actions cron always uses UTC. Update it to your required local time.
- Manual run is available using **workflow_dispatch**.

## Local run

```bash
pip install -r requirements.txt
export GOOGLE_SERVICE_ACCOUNT_FILE=credential.json
export DRIVE_FOLDER_ID="your_drive_folder_id"
export RECIPIENTS="a@example.com,b@example.com"
export SMTP_EMAIL="your_gmail@gmail.com"
export SMTP_APP_PASSWORD="your_16_char_app_password"
python scripts/daily_drive_mailer.py
```
