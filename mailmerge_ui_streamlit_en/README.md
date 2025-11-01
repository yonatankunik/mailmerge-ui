# Mail-Merge Letters – Streamlit UI (English, Light Theme)

A friendly Streamlit UI to generate personalized letters (DOCX) from an Excel guest list and per-group templates.
Bright, inviting design with a light theme.

## Run locally
```bash
python -m venv .venv
# Windows:
.venv\Scripts\activate
# macOS / Linux:
source .venv/bin/activate

pip install -r requirements.txt
streamlit run app.py
```

## Expected Excel columns
`FullName`, `Address`, `Institution`, `Group` (values: כחול/ירוק/צהוב for Blue/Green/Yellow).
Templates support placeholders: `{{FullName}}`, `{{Address}}`, `{{Institution}}`.
