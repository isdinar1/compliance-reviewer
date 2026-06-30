# Course Compliance Reviewer

Upload a course script and any regulation/guidelines document — get back a Word doc with every conflict (yellow) and gap (cyan) highlighted.

## What it does

- **Input 1**: Course document (upload .docx, .pdf, .txt — or paste text) + optional Course SKU
- **Input 2**: Regulation/guidelines documents (upload one or more files)
- **Input 3**: URLs to scan for rules or regulations (paste links, one per line)
- **Output**: Downloadable Word doc with yellow highlights for conflicts and cyan highlights for gaps, plus a full Comment Legend

## Setup

### 1. Get a free Gemini API key
Go to [aistudio.google.com](https://aistudio.google.com), sign in with a Google account, and create an API key. No credit card required.

### 2. Install dependencies
```bash
pip install -r requirements.txt
```

### 3. Add your API key
Create `.streamlit/secrets.toml`:
```toml
GEMINI_API_KEY = "your-key-here"
```

### 4. Run locally
```bash
streamlit run app.py
```

## Deploy to Streamlit Community Cloud (free)

1. Push this repo to GitHub
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Connect your GitHub repo
4. Add `GEMINI_API_KEY` in the **Secrets** section of the app settings
5. Deploy — anyone with the link can use it, no setup needed

## Output format

- 🟡 **Yellow highlight** = ISSUE: course text directly conflicts with the regulation
- 🔵 **Cyan highlight** = GAP: regulation requires content that is absent from the course
- `[n]` markers link inline highlights to a full Comment Legend on the last page
