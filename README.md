# Inscribe 2.0 — Hebrew OCR (Online Test Version)

Streamlit app using Google Gemini AI for Hebrew OCR (printed + handwriting).

## Deploy for free on Streamlit Cloud

1. Create a free account at [streamlit.io](https://streamlit.io)
2. Connect your GitHub account
3. Push this folder to a GitHub repo
4. In Streamlit Cloud → "New app" → select your repo → `app.py`
5. Deploy — you get a public link to share with anyone

## Get your free Gemini API key

1. Go to [aistudio.google.com](https://aistudio.google.com)
2. Sign in with Google
3. Click "Get API key" → "Create API key"
4. Paste it in the sidebar when you open the app

## Files
- `app.py` — main Streamlit app
- `requirements.txt` — Python dependencies
- `.streamlit/config.toml` — UI theme config


## Public-link setup (recommended)

In Streamlit Community Cloud, open **App settings → Secrets** and add:

```toml
GEMINI_API_KEY = "your-key-here"
GEMINI_MODEL = "gemini-2.5-flash"
```

Never commit the API key to GitHub. With a secret configured, visitors can use the app without seeing or entering your key. Free quotas and model availability can change; set strict provider quotas before sharing broadly.
