# Naukri Automation – Candidate Data Collection

A Python-based automation project designed to assist with candidate search, candidate data collection, and resume downloading through Naukri Resdex, using browser automation.

## Features

- Automates supported browser interactions for candidate search.
- Helps collect candidate information from search results.
- Supports resume downloading through permitted user-account actions.
- Uses Python and Playwright for browser automation.
- Supports parsing downloaded PDF and DOCX resumes.
- Uses Pandas to organize extracted data for Excel export.
- Keeps candidate data and browser sessions locally rather than committing them to GitHub.

## Tech Stack

- **Python** – Core automation logic
- **Playwright** – Browser automation
- **Pandas** – Data processing and Excel export
- **PyMuPDF** – PDF text extraction
- **python-docx** – DOCX resume parsing

## Project Structure

```text
naukri-automation/
├── main.py
├── recorded.py
├── .gitignore
├── Naukri_Automation_User_Guide.pdf
├── local_data/       # Local data; not committed
├── session/          # Local session files; not committed
├── venv/             # Python virtual environment; not committed
└── chrome-profile/   # Local browser profile; not committed
```

The actual files and directories may differ depending on your local setup.

## Prerequisites

- Python 3.10 or later (recommended; verify compatibility with your installed dependencies)
- Google Chrome or another supported browser
- An authorized Naukri Resdex account
- Git, if you want to clone the repository

## Installation

### 1. Clone the repository

```bash
git clone https://github.com/rajjaiswal3858/Naukri_automation_CANDIDATE_DATA_COLLECTION.git
cd Naukri_automation_CANDIDATE_DATA_COLLECTION
```

### 2. Create a virtual environment

**Windows PowerShell:**

```powershell
python -m venv venv
.\venv\Scripts\Activate.ps1
```

### 3. Install dependencies

Install the packages required by the project:

```powershell
python -m pip install playwright pandas pymupdf python-docx openpyxl
```

If the project has a `requirements.txt` file, prefer installing the pinned dependencies from it:

```powershell
pip install -r requirements.txt
```

### 4. Run the automation

Check the available command-line options in `main.py`, then run the script using the appropriate arguments. For example, if supported by your version:

```powershell
python main.py --limit 1
```

The exact arguments depend on the current implementation.

## Data and Privacy

- Keep browser profiles, authentication cookies, session files, and credentials private.
- Do not commit `.env` files, real candidate records, or downloaded resumes to a public repository.
- Store candidate data locally and restrict access to authorized users.
- Collect and process candidate information only when authorized and in accordance with applicable privacy requirements and Naukri's terms.
- Review `.gitignore` before every upload to ensure sensitive files are excluded.

## Troubleshooting

**Browser session not found**

Follow the login and session setup procedure implemented in the project before running automation.

**No candidates found**

Check the search criteria, account permissions, current page structure, and whether the search page loaded successfully.

**Resume parsing issues**

Confirm that downloaded files are valid PDF or DOCX documents and that the appropriate parsing libraries are installed.

**Git push or authentication issues**

Confirm that Git is configured and that your GitHub account has write access to the repository.

## Repository

GitHub: https://github.com/rajjaiswal3858/Naukri_automation_CANDIDATE_DATA_COLLECTION

## Disclaimer

This project is intended for authorized recruitment workflows and data processing. Users are responsible for ensuring compliance with applicable laws, platform terms, account permissions, and candidate privacy requirements.

## License

No license has been specified yet. Unless a license is added to this repository, reuse and redistribution are subject to applicable copyright law.
