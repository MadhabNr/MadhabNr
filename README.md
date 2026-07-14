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
2. Enable **Google Drive API** and **Gmail API**.
3. Create OAuth client credentials.
4. Obtain a refresh token for the Gmail/Drive account you want to use.
5. Share the Drive folder with that account if required.

## Required GitHub Secrets

Set these repository secrets:

- `GOOGLE_CLIENT_ID`
- `GOOGLE_CLIENT_SECRET`
- `GOOGLE_REFRESH_TOKEN`
- `DRIVE_FOLDER_ID`
- `RECIPIENTS` (comma-separated emails)

Optional:

- `GOOGLE_USER_EMAIL` (default: `me`)
- `MAIL_SUBJECT`

## Schedule configuration

Workflow file: `/home/runner/work/MadhabNr/MadhabNr/.github/workflows/daily-drive-mail.yml`

- Current cron: `30 3 * * *` (UTC)
- GitHub Actions cron always uses UTC. Update it to your required local time.
- Manual run is available using **workflow_dispatch**.

## Local run

```bash
pip install -r requirements.txt
python scripts/daily_drive_mailer.py
```
