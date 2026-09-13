# Outlook Bulk Email Sender

Sends personalised HTML email to an Excel mailing list through desktop Outlook,
rotating sender identities and writing delivery status back to the sheet.

![Python](https://img.shields.io/badge/Python-3.9%2B-3776AB?logo=python&logoColor=white) ![Outlook](https://img.shields.io/badge/Outlook-0078D4?logo=microsoftoutlook&logoColor=white) ![License](https://img.shields.io/badge/License-MIT-2ea44f)

## What it does

Drives the Outlook client already installed and signed in on your machine, so mail
goes out from real mailboxes rather than a third-party relay.

- **Sender rotation** - switches account after a configurable number of successful
  sends, instead of pushing an entire list through one mailbox
- **Status write-back** - each row records its own outcome
- **Autosave** - the sheet is saved as it goes, so an interrupted run is never
  ambiguous about where it stopped
- **Pre-flight check** - verifies every configured sender actually exists in your
  Outlook profile and fails immediately if one does not, rather than halfway through

## Requirements

- Python 3.9+, Windows
- Outlook desktop, open and signed in
- `pandas`, `pywin32`, `openpyxl`

## Install

```bash
pip install -r requirements.txt
```

## Configure

At the top of `main.py`:

| Setting | What it is |
|---|---|
| `SENDER_IDENTITIES` | The Outlook accounts to send from, with display names |
| `EMAIL_SUBJECT` | Subject line |
| `BATCH_SIZE_PER_SENDER` | Successful sends before switching account (default 30) |

## Run

```bash
python main.py
```

## Notes on data

This repository contains **no client or patient data**. `.gitignore` already excludes
`.xlsx`, `.pdf` and `.csv` files, `service_account.json`, and chromedriver binaries -
keep it that way if you fork this.

## License

MIT © Muhammad Sharaz Khalid - see [LICENSE](LICENSE).
