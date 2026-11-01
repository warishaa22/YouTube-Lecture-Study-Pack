# 🎓 YouTube Lecture → Study Pack

A Streamlit app that turns a YouTube lecture into a complete study pack:

- 📝 **Summary**: a short 5–7 sentence overview
- 📚 **Notes**: structured notes with headings, plus key terms (download as Markdown)
- 🖼️ **Slides**: an auto-generated slide deck (download as `.pptx`)
- ❓ **Quiz**: interactive multiple-choice questions with instant feedback and a score

> Add a screenshot or GIF here: `![Demo](docs/demo.png)`

## How it works

1. Extracts the video ID from the link
2. Downloads the transcript with `youtube-transcript-api`
3. Long lectures are split into chunks and condensed first
4. One Gemini call returns the whole study pack as JSON
5. The output is validated and shown in tabs; slides are built with `python-pptx`

## Languages

Lectures in Hindi (or other languages) work. Choose the **Output language** in the sidebar:
English, Hindi (Devanagari) or Hinglish (Hindi in English letters).

## Tech stack

Python · Streamlit · Google Gemini (Flash) · youtube-transcript-api · python-pptx

## Run locally

```bash
git clone https://github.com/<your-username>/yt-study-pack.git
cd yt-study-pack

python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

Add your API key (get one free at https://aistudio.google.com/app/apikey):

```bash
cp .streamlit/secrets.toml.example .streamlit/secrets.toml
# then edit .streamlit/secrets.toml and paste your key
```

Start the app:

```bash
streamlit run app.py
```

## Push to GitHub

```bash
git init
git add .
git commit -m "Initial commit: YouTube Lecture to Study Pack"
git branch -M main
git remote add origin https://github.com/<your-username>/yt-study-pack.git
git push -u origin main
```

Before pushing, check that `git status` does **not** list `.streamlit/secrets.toml`.

## Deploy on Streamlit Cloud

1. Go to https://share.streamlit.io and sign in with GitHub
2. Click **Create app** → choose your repo, branch `main`, main file `app.py`
3. Open **Advanced settings → Secrets** and paste:
   ```toml
   GEMINI_API_KEY = "your-real-key"
   ```
4. Click **Deploy**

## Troubleshooting

| Problem | Fix |
|---|---|
| "YouTube blocked the transcript request" | YouTube blocks most cloud-server IPs. Use the **Paste / upload transcript** tab (YouTube → …more → Show transcript → copy), run the app locally, or add proxy credentials (see below). |
| "Transcripts are disabled" / "No captions" | The video has no captions. Pick a lecture that shows a CC option. |
| "No API key found" | Add `GEMINI_API_KEY` to `.streamlit/secrets.toml` (local) or the Secrets panel (cloud). |
| Model not found error | Set `GEMINI_MODEL` in secrets to a Flash model your key supports. |
| Rate-limit / quota errors | The free tier has limits. Wait a minute, or try a shorter lecture or fewer quiz questions. |
| Quiz resets when I change a setting | Settings only apply when you press **Generate** again. |

## Optional: proxy for the link method on Streamlit Cloud

Add one of these to your secrets so transcript requests go through a proxy:

```toml
# Option A: Webshare (use a "Residential" proxy plan)
WEBSHARE_USERNAME = "your-username"
WEBSHARE_PASSWORD = "your-password"

# Option B: any other HTTP proxy
PROXY_URL = "http://user:pass@host:port"
```

Proxies usually cost money, so the paste/upload tab is the free alternative.

## Ideas for next steps

- Timestamps in notes that link back to the video moment
- Export notes as PDF
- Flashcard mode built from the key terms
- Save past study packs in SQLite

## License

MIT
