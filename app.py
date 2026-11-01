"""
YouTube Lecture → Study Pack
----------------------------
Paste a YouTube lecture link and get:
  1. A short summary
  2. Structured study notes + key terms
  3. A slide deck (shown in the app and downloadable as .pptx)
  4. An interactive multiple-choice quiz

Flow: link -> transcript -> (chunk + condense if long) -> ONE JSON call to Gemini
      -> validate -> show in tabs.
"""

import io
import json
import os
import re
import time

import google.generativeai as genai
import streamlit as st
from pptx import Presentation
from pptx.util import Pt
from youtube_transcript_api import (
    NoTranscriptFound,
    TranscriptsDisabled,
    VideoUnavailable,
    YouTubeTranscriptApi,
)

# ----------------------------------------------------------------------------
# Settings
# ----------------------------------------------------------------------------
DEFAULT_MODEL = "gemini-flash-latest"  # override with GEMINI_MODEL in secrets
SINGLE_CALL_LIMIT = 25_000  # transcripts shorter than this go straight to the final call
CHUNK_SIZE = 12_000  # characters per chunk for long transcripts
MAX_CHUNKS = 20  # safety limit so a huge video doesn't burn your API quota
PREFERRED_LANGUAGES = ["en", "en-US", "en-GB"]


class AppError(Exception):
    """An error whose message is safe and friendly to show to the user."""


# ----------------------------------------------------------------------------
# Secrets / API helpers
# ----------------------------------------------------------------------------
def get_secret(name: str, default: str | None = None) -> str | None:
    """Read from Streamlit secrets first, then environment variables."""
    try:
        return st.secrets[name]
    except Exception:
        return os.environ.get(name, default)


def call_gemini(prompt: str, json_mode: bool = False) -> str:
    """Send a prompt to Gemini (with a few retries) and return the text."""
    api_key = get_secret("GEMINI_API_KEY")
    if not api_key:
        raise AppError(
            "No API key found. Add GEMINI_API_KEY to .streamlit/secrets.toml "
            "(local) or to the Secrets panel on Streamlit Cloud."
        )

    genai.configure(api_key=api_key)
    model = genai.GenerativeModel(get_secret("GEMINI_MODEL", DEFAULT_MODEL))

    config = {"temperature": 0.4}
    if json_mode:
        config["response_mime_type"] = "application/json"

    last_error = None
    for attempt in range(3):
        try:
            response = model.generate_content(prompt, generation_config=config)
            return response.text
        except Exception as error:  # network issues, rate limits, blocked output...
            last_error = error
            time.sleep(2 * (attempt + 1))
    raise AppError(f"The AI request failed after 3 tries: {last_error}")


# ----------------------------------------------------------------------------
# YouTube helpers
# ----------------------------------------------------------------------------
def get_video_id(url: str) -> str | None:
    """Extract the 11-character video ID from the common YouTube URL formats."""
    pattern = r"(?:v=|youtu\.be/|embed/|shorts/|live/)([A-Za-z0-9_-]{11})"
    match = re.search(pattern, url.strip())
    return match.group(1) if match else None


def make_transcript_api() -> YouTubeTranscriptApi:
    """
    Build the transcript client. YouTube blocks most cloud-server IPs, so if you
    add proxy credentials to your secrets the request goes through a proxy.
    """
    webshare_user = get_secret("WEBSHARE_USERNAME")
    webshare_pass = get_secret("WEBSHARE_PASSWORD")
    proxy_url = get_secret("PROXY_URL")

    if webshare_user and webshare_pass:
        from youtube_transcript_api.proxies import WebshareProxyConfig

        return YouTubeTranscriptApi(
            proxy_config=WebshareProxyConfig(
                proxy_username=webshare_user, proxy_password=webshare_pass
            )
        )
    if proxy_url:
        from youtube_transcript_api.proxies import GenericProxyConfig

        return YouTubeTranscriptApi(
            proxy_config=GenericProxyConfig(http_url=proxy_url, https_url=proxy_url)
        )
    return YouTubeTranscriptApi()


def clean_transcript_text(raw: str) -> str:
    """
    Clean a transcript the user pasted or uploaded. Handles plain text copied from
    YouTube's "Show transcript" panel as well as .srt and .vtt subtitle files.
    """
    text = raw.replace("WEBVTT", " ")
    text = re.sub(r"^\s*\d+\s*$", " ", text, flags=re.MULTILINE)  # SRT line counters
    text = re.sub(  # timestamps such as 0:05, 01:02:03,500 or 00:01 --> 00:04
        r"\d{1,2}:\d{2}(?::\d{2})?(?:[.,]\d{1,3})?"
        r"(?:\s*-->\s*\d{1,2}:\d{2}(?::\d{2})?(?:[.,]\d{1,3})?)?",
        " ",
        text,
    )
    text = re.sub(r"<[^>]+>", " ", text)  # tags inside .vtt files
    text = re.sub(r"\[[^\]]*\]", " ", text)  # [Music], [Applause]
    text = re.sub(r"\s+", " ", text).strip()
    if len(text) < 200:
        raise AppError("That transcript is too short (under 200 characters) to build a study pack.")
    return text


@st.cache_data(show_spinner=False)
def get_transcript(video_id: str) -> str:
    """Download the transcript text. Cached so the same video isn't fetched twice."""
    api = make_transcript_api()
    try:
        try:
            fetched = api.fetch(video_id, languages=PREFERRED_LANGUAGES)
        except NoTranscriptFound:
            # No English transcript: fall back to whatever language exists.
            available = list(api.list(video_id))
            if not available:
                raise
            fetched = available[0].fetch()
    except TranscriptsDisabled:
        raise AppError("The owner of this video has disabled transcripts.")
    except NoTranscriptFound:
        raise AppError("No captions/transcript are available for this video.")
    except VideoUnavailable:
        raise AppError("This video is unavailable (private, deleted or region-locked).")
    except Exception as error:
        message = str(error).lower()
        if "block" in message or "ip" in message:
            raise AppError(
                "YouTube blocked the transcript request (this is common on cloud servers). "
                "Use the 'Paste / upload transcript' tab instead, or run the app locally."
            )
        raise AppError(f"Could not fetch the transcript: {error}")

    text = " ".join(snippet.text for snippet in fetched)
    text = re.sub(r"\[[^\]]*\]", " ", text)  # drop tags like [Music] or [Applause]
    text = re.sub(r"\s+", " ", text).strip()
    if len(text) < 200:
        raise AppError("The transcript is too short to build a study pack from.")
    return text


# ----------------------------------------------------------------------------
# AI pipeline
# ----------------------------------------------------------------------------
def split_into_chunks(text: str, size: int = CHUNK_SIZE) -> list[str]:
    """Split a long transcript into word-aligned chunks of roughly `size` characters."""
    chunks, current, length = [], [], 0
    for word in text.split():
        current.append(word)
        length += len(word) + 1
        if length >= size:
            chunks.append(" ".join(current))
            current, length = [], 0
    if current:
        chunks.append(" ".join(current))
    return chunks


@st.cache_data(show_spinner=False)
def condense_chunk(chunk: str, position: int, total: int) -> str:
    """Turn one raw transcript chunk into compact, faithful notes (plain text)."""
    prompt = (
        f"This is part {position} of {total} of a lecture transcript. "
        "Rewrite it as compact, faithful study notes: keep every key concept, "
        "definition, formula, example and conclusion; remove filler and repetition. "
        "Plain text bullet points only.\n\nTRANSCRIPT PART:\n" + chunk
    )
    return call_gemini(prompt)


def build_final_prompt(material: str, n_slides: int, n_questions: int, difficulty: str) -> str:
    """The main prompt: asks for ONE JSON object containing the whole study pack."""
    return (
        "You are an expert teaching assistant. Using ONLY the lecture material below, "
        "return ONE valid JSON object (no markdown, no commentary) with exactly this shape:\n"
        "{\n"
        '  "title": "short lecture title",\n'
        '  "summary": "5-7 sentence summary",\n'
        '  "notes": [{"heading": "topic", "points": ["bullet", "bullet"]}],\n'
        '  "key_terms": [{"term": "...", "meaning": "..."}],\n'
        '  "slides": [{"title": "...", "bullets": ["...", "..."]}],\n'
        '  "quiz": [{"question": "...", "options": ["...", "...", "...", "..."], '
        '"answer": "must be exactly one of the options", "explanation": "..."}]\n'
        "}\n\n"
        "Rules:\n"
        f"- Exactly {n_slides} slides, each with a title and at most 4 short bullets "
        "(max 12 words per bullet). Slides should follow the lecture's order.\n"
        f"- Exactly {n_questions} multiple-choice questions, each with 4 options, "
        f"difficulty level: {difficulty}. Spread correct answers across positions.\n"
        "- 4-8 note sections with clear headings; 5-12 key terms.\n"
        "- Do not invent facts that are not in the material.\n\n"
        "LECTURE MATERIAL:\n" + material
    )


def parse_json(text: str) -> dict:
    """Parse the model's reply, tolerating ```json fences or extra text around it."""
    cleaned = re.sub(r"^```(?:json)?|```$", "", text.strip(), flags=re.MULTILINE).strip()
    try:
        return json.loads(cleaned)
    except json.JSONDecodeError:
        start, end = cleaned.find("{"), cleaned.rfind("}")
        if start != -1 and end > start:
            try:
                return json.loads(cleaned[start : end + 1])
            except json.JSONDecodeError:
                pass
    raise AppError("The AI returned an unreadable response. Please try again.")


def normalize_pack(raw: dict) -> dict:
    """Validate and clean the AI output so the UI never crashes on odd data."""
    if not isinstance(raw, dict):
        raise AppError("The AI response had an unexpected format. Please try again.")

    pack = {
        "title": str(raw.get("title") or "Lecture Study Pack"),
        "summary": str(raw.get("summary") or "No summary was generated."),
        "notes": [],
        "key_terms": [],
        "slides": [],
        "quiz": [],
    }

    for section in raw.get("notes") or []:
        if isinstance(section, dict) and section.get("heading"):
            points = [str(p) for p in (section.get("points") or [])]
            pack["notes"].append({"heading": str(section["heading"]), "points": points})

    for item in raw.get("key_terms") or []:
        if isinstance(item, dict) and item.get("term"):
            pack["key_terms"].append(
                {"term": str(item["term"]), "meaning": str(item.get("meaning", ""))}
            )

    for slide in raw.get("slides") or []:
        if isinstance(slide, dict) and slide.get("title"):
            bullets = [str(b) for b in (slide.get("bullets") or [])][:4]
            pack["slides"].append({"title": str(slide["title"]), "bullets": bullets})

    for q in raw.get("quiz") or []:
        if not isinstance(q, dict):
            continue
        options = [str(o) for o in (q.get("options") or [])]
        if not q.get("question") or len(options) < 2:
            continue
        answer = str(q.get("answer", ""))
        if answer not in options:
            # The model sometimes answers with a letter ("B") or a different casing.
            letters = "ABCD"
            if answer.strip().upper()[:1] in letters and len(answer.strip()) <= 2:
                index = letters.index(answer.strip().upper()[:1])
                answer = options[index] if index < len(options) else options[0]
            else:
                lowered = {o.lower(): o for o in options}
                answer = lowered.get(answer.strip().lower(), options[0])
        pack["quiz"].append(
            {
                "question": str(q["question"]),
                "options": options,
                "answer": answer,
                "explanation": str(q.get("explanation", "")),
            }
        )

    if not pack["slides"] and not pack["notes"]:
        raise AppError("The AI response was missing the main content. Please try again.")
    return pack


@st.cache_data(show_spinner=False)
def generate_pack(material: str, n_slides: int, n_questions: int, difficulty: str) -> dict:
    """The single JSON call that produces the whole study pack (cached)."""
    prompt = build_final_prompt(material, n_slides, n_questions, difficulty)
    try:
        return normalize_pack(parse_json(call_gemini(prompt, json_mode=True)))
    except AppError as first_error:
        # One extra attempt, since malformed JSON is usually a one-off.
        if "unreadable" in str(first_error) or "unexpected format" in str(first_error):
            return normalize_pack(parse_json(call_gemini(prompt, json_mode=True)))
        raise


def build_study_pack(get_text, n_slides: int, n_questions: int, difficulty: str) -> dict:
    """
    Run the whole pipeline with live progress messages in the UI.
    `get_text` is a function that returns the transcript text (from YouTube or pasted).
    """
    with st.status("Building your study pack...", expanded=True) as status:
        st.write("📥 Getting transcript...")
        transcript = get_text()
        st.write(f"✅ Transcript ready ({len(transcript):,} characters)")

        if len(transcript) <= SINGLE_CALL_LIMIT:
            material = transcript
        else:
            chunks = split_into_chunks(transcript)
            if len(chunks) > MAX_CHUNKS:
                raise AppError(
                    f"This lecture is very long ({len(chunks)} parts). "
                    f"The limit is {MAX_CHUNKS} parts. Try a shorter video."
                )
            st.write(f"✂️ Long lecture: condensing {len(chunks)} parts...")
            progress = st.progress(0.0)
            condensed = []
            for i, chunk in enumerate(chunks, start=1):
                condensed.append(condense_chunk(chunk, i, len(chunks)))
                progress.progress(i / len(chunks))
            material = "\n\n".join(condensed)

        st.write("🧠 Generating summary, notes, slides and quiz...")
        pack = generate_pack(material, n_slides, n_questions, difficulty)
        status.update(label="Study pack ready!", state="complete", expanded=False)
    return pack


# ----------------------------------------------------------------------------
# Exports
# ----------------------------------------------------------------------------
def build_pptx(pack: dict) -> bytes:
    """Create a .pptx file (title slide, content slides, key terms) in memory."""
    prs = Presentation()

    title_slide = prs.slides.add_slide(prs.slide_layouts[0])
    title_slide.shapes.title.text = pack["title"]
    title_slide.placeholders[1].text = "Study pack generated from a YouTube lecture"

    def add_bullet_slide(title: str, bullets: list[str]) -> None:
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = title
        slide.shapes.title.text_frame.paragraphs[0].font.size = Pt(34)
        frame = slide.placeholders[1].text_frame
        frame.word_wrap = True
        for i, bullet in enumerate(bullets):
            paragraph = frame.paragraphs[0] if i == 0 else frame.add_paragraph()
            paragraph.text = bullet
            paragraph.font.size = Pt(24)

    for slide_data in pack["slides"]:
        add_bullet_slide(slide_data["title"], slide_data["bullets"])

    if pack["key_terms"]:
        terms = [f"{t['term']}: {t['meaning']}" for t in pack["key_terms"][:6]]
        add_bullet_slide("Key Terms", terms)

    buffer = io.BytesIO()
    prs.save(buffer)
    return buffer.getvalue()


def build_markdown(pack: dict) -> str:
    """Notes, summary and key terms as a Markdown file students can keep."""
    lines = [f"# {pack['title']}", "", "## Summary", pack["summary"], "", "## Notes"]
    for section in pack["notes"]:
        lines += ["", f"### {section['heading']}"]
        lines += [f"- {point}" for point in section["points"]]
    if pack["key_terms"]:
        lines += ["", "## Key Terms"]
        lines += [f"- **{t['term']}**: {t['meaning']}" for t in pack["key_terms"]]
    return "\n".join(lines) + "\n"


# ----------------------------------------------------------------------------
# UI
# ----------------------------------------------------------------------------
def render_quiz(quiz: list[dict]) -> None:
    """Interactive quiz with instant feedback and a running score."""
    if not quiz:
        st.info("No quiz questions were generated.")
        return

    round_id = st.session_state.get("quiz_round", 0)
    score = 0
    answered = 0

    for i, q in enumerate(quiz):
        choice = st.radio(
            f"**Q{i + 1}. {q['question']}**",
            q["options"],
            index=None,
            key=f"quiz_{round_id}_{i}",
        )
        if choice is not None:
            answered += 1
            if choice == q["answer"]:
                score += 1
                st.success(f"Correct! {q['explanation']}")
            else:
                st.error(f"Not quite. Correct answer: **{q['answer']}**. {q['explanation']}")
        st.divider()

    st.metric("Score", f"{score} / {len(quiz)}", help=f"{answered} answered so far")
    if answered == len(quiz) and score == len(quiz):
        st.balloons()
    if st.button("🔄 Retake quiz"):
        st.session_state["quiz_round"] = round_id + 1
        st.rerun()


def render_results(pack: dict) -> None:
    st.header(pack["title"])
    tab_summary, tab_notes, tab_slides, tab_quiz = st.tabs(
        ["📝 Summary", "📚 Notes", "🖼️ Slides", "❓ Quiz"]
    )

    with tab_summary:
        st.write(pack["summary"])

    with tab_notes:
        for section in pack["notes"]:
            st.subheader(section["heading"])
            for point in section["points"]:
                st.markdown(f"- {point}")
        if pack["key_terms"]:
            st.subheader("🔑 Key terms")
            for term in pack["key_terms"]:
                st.markdown(f"**{term['term']}**: {term['meaning']}")
        st.download_button(
            "⬇️ Download notes (.md)",
            build_markdown(pack),
            file_name="lecture_notes.md",
            mime="text/markdown",
        )

    with tab_slides:
        for number, slide in enumerate(pack["slides"], start=1):
            with st.container(border=True):
                st.markdown(f"**{number}. {slide['title']}**")
                for bullet in slide["bullets"]:
                    st.markdown(f"- {bullet}")
        st.download_button(
            "⬇️ Download slides (.pptx)",
            build_pptx(pack),
            file_name="lecture_slides.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )

    with tab_quiz:
        render_quiz(pack["quiz"])


def run_pipeline(get_text, n_slides: int, n_questions: int, difficulty: str) -> None:
    """Run the pipeline, save the result in session_state, and show friendly errors."""
    try:
        st.session_state["pack"] = build_study_pack(get_text, n_slides, n_questions, difficulty)
        st.session_state["quiz_round"] = 0
    except AppError as error:
        st.error(str(error))
        if "blocked" in str(error).lower():
            st.info(
                "💡 Tip: open YouTube → click **…more** under the video → **Show transcript** → "
                "copy the text, then paste it in the **Paste / upload transcript** tab."
            )
    except Exception as error:  # last-resort safety net
        st.error(f"Something unexpected went wrong: {error}")


def main() -> None:
    st.set_page_config(page_title="Lecture → Study Pack", page_icon="🎓", layout="wide")
    st.title("🎓 YouTube Lecture → Study Pack")
    st.caption("Paste a lecture link (or its transcript) and get a summary, notes, slides and a quiz.")

    with st.sidebar:
        st.header("⚙️ Settings")
        n_slides = st.slider("Number of slides", 6, 15, 10)
        n_questions = st.slider("Quiz questions", 3, 10, 5)
        difficulty = st.selectbox("Quiz difficulty", ["easy", "medium", "hard"], index=1)
        st.markdown("---")
        st.caption("Works best with lectures that have captions (manual or auto-generated).")

    tab_link, tab_paste = st.tabs(["🔗 YouTube link", "📋 Paste / upload transcript"])

    with tab_link:
        url = st.text_input(
            "YouTube lecture link", placeholder="https://www.youtube.com/watch?v=..."
        )
        if st.button("✨ Generate study pack", type="primary", key="generate_link"):
            video_id = get_video_id(url) if url else None
            if not video_id:
                st.error("That doesn't look like a valid YouTube link. Please check it and try again.")
            else:
                run_pipeline(lambda: get_transcript(video_id), n_slides, n_questions, difficulty)

    with tab_paste:
        st.markdown(
            "**Use this if the link method is blocked.** On YouTube, click **…more** under the "
            "video → **Show transcript**, copy everything, and paste it below. "
            "You can also upload a `.txt`, `.srt` or `.vtt` file."
        )
        pasted = st.text_area("Transcript text", height=220)
        uploaded = st.file_uploader("...or upload a transcript file", type=["txt", "srt", "vtt"])
        if st.button("✨ Generate study pack", type="primary", key="generate_paste"):
            raw = uploaded.getvalue().decode("utf-8", errors="ignore") if uploaded else pasted
            if not raw.strip():
                st.error("Please paste a transcript or upload a file first.")
            else:
                run_pipeline(lambda: clean_transcript_text(raw), n_slides, n_questions, difficulty)

    pack = st.session_state.get("pack")
    if pack:
        render_results(pack)


if __name__ == "__main__":
    main()
