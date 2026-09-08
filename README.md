# Outlook Bulk Email Sender

Sends an HTML email to every address in an Excel column using Outlook
on this computer. Rotates through several sender accounts, writes the
status of each send back into the sheet, and autosaves as it goes.

## Setup

```
pip install -r requirements.txt
```

## Run

Outlook must be open and signed in, then:

```
python main.py
```

## What to edit before running

`SENDER_IDENTITIES` and `EMAIL_SUBJECT` at the top of `main.py`.
Fill in your own Outlook accounts before running.

## Never commit

Patient or client data (`.xlsx`, `.pdf`, `.csv`), `service_account.json`,
and chromedriver binaries. All are covered by `.gitignore`.
