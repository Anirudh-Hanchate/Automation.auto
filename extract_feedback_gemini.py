"""
extract_feedback_gemini.py
---------------------------
Uses Google's Gemini API (via Google GenAI SDK) to read scanned, hand-filled
feedback form PDFs and transcribes the circled/ticked scores and open-ended
suggestions into structured JSON matching the fill_excel.py schema.

Setup:
    pip install -r requirements.txt
    export GEMINI_API_KEY=AIza... (or create a .env file with GEMINI_API_KEY=...)

Usage:
    python extract_feedback_gemini.py CRE_Feedback.pdf --pages-per-participant 2 --out extracted_data.json
"""
import argparse
from concurrent.futures import ThreadPoolExecutor, as_completed
from io import BytesIO
import json
import os
import re
import sys
import time

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

try:
    import config
except ImportError:
    config = None

try:
    from pypdf import PdfReader, PdfWriter
except ImportError:
    sys.exit("Error: pypdf not found. Run: pip install pypdf")

try:
    from google import genai
    from google.genai import types
except ImportError:
    sys.exit("Error: google-genai not found. Run: pip install google-genai")

DEFAULT_MODEL = getattr(config, "DEFAULT_GEMINI_MODEL", "gemini-3.6-flash") if config else "gemini-3.6-flash"
FALLBACK_MODEL = getattr(config, "FALLBACK_GEMINI_MODEL", "gemini-3.5-flash-lite") if config else "gemini-3.5-flash-lite"
DEFAULT_MAX_WORKERS = 5


SYSTEM_PROMPT = """You are an expert AI evaluator transcribing scanned handwritten training feedback forms.
Each form has multiple day blocks. Each day has 1 to 3 topic/session blocks.
Each topic block has exactly 5 criteria rows:
1. Course Content (or Course Contenet)
2. Structure & Flow
3. Time Management
4. Faculty Delivery
5. Faculty Participants Interaction

Each row is a printed 10-9-8-7-6-5-4-3-2-1 scale where the participant circled, ticked, or marked ONE number.
Read the hand-drawn mark accurately (circles, checkmarks, ticks, underlines, or cross marks) and extract the circled integer (1-10).

Return ONLY valid JSON matching this exact schema:
{
  "staff_no": "<string, participant staff no printed on the form>",
  "name": "<string, participant name>",
  "sbu_csg": "<string or null>",
  "department": "<string or null>",
  "days": [
    {
      "day": <integer day number, 1-based>,
      "topics": [
        {
          "topic_index": <integer, 1-based position of topic within this day>,
          "topic_name": "<topic title as printed>",
          "criteria": {
            "Course Content": <integer 1-10>,
            "Structure & Flow": <integer 1-10>,
            "Time Management": <integer 1-10>,
            "Faculty Delivery": <integer 1-10>,
            "Faculty Participants Interaction": <integer 1-10>
          }
        }
      ]
    }
  ],
  "overall": {
    "<overall criteria label>": <integer 1-10>
  },
  "recommend_continue": "Yes" | "No" | null,
  "suggestion_1": "<answer to 'other training programs to add', or empty string>",
  "suggestion_2": "<answer to 'any suggestion for future program', or empty string>"
}

Rules:
- Capture all topics and criteria rows present on the scanned pages.
- If a row is left blank, use null.
- Extract numbers exactly as circled (1-10).
- Preserve exact staff_no as written (including letters/numbers).
"""


def encode_pdf_chunk(reader: PdfReader, page_indices) -> bytes:
    writer = PdfWriter()
    for i in page_indices:
        writer.add_page(reader.pages[i])
    buf = BytesIO()
    writer.write(buf)
    return buf.getvalue()


def extract_json(text: str):
    """Extract and parse JSON from response text even if wrapped in markdown."""
    text = text.strip()
    text = re.sub(r"^```(?:json)?\s*", "", text, flags=re.IGNORECASE)
    text = re.sub(r"\s*```$", "", text)
    start_idx = text.find("{")
    end_idx = text.rfind("}")
    if start_idx != -1 and end_idx != -1:
        text = text[start_idx:end_idx + 1]
    return json.loads(text)


def call_gemini(client, pdf_bytes: bytes, chunk_label: str, model_name: str = DEFAULT_MODEL, max_retries=5):
    current_model = model_name or DEFAULT_MODEL
    if current_model in ("gemini-2.5-flash", "gemini-2.0-flash", "gemini-1.5-flash"):
        current_model = DEFAULT_MODEL

    for attempt in range(1, max_retries + 1):
        try:
            resp = client.models.generate_content(
                model=current_model,
                contents=[
                    types.Part.from_bytes(data=pdf_bytes, mime_type="application/pdf"),
                    "Transcribe all handwritten circled ratings and suggestions from this participant feedback form into JSON.",
                ],
                config=types.GenerateContentConfig(
                    system_instruction=SYSTEM_PROMPT,
                    temperature=0,
                    response_mime_type="application/json",
                ),
            )
            return extract_json(resp.text)
        except Exception as e:
            err_str = str(e)
            is_rate_limit = "429" in err_str or "RESOURCE_EXHAUSTED" in err_str
            is_unavailable = "503" in err_str or "UNAVAILABLE" in err_str

            retry_match = re.search(r"retry in (\d+(?:\.\d+)?)s", err_str, re.IGNORECASE)
            delay_sec = float(retry_match.group(1)) if retry_match else (10 * attempt if is_rate_limit else 3 * attempt)
            wait = max(delay_sec + 1, 4)

            if (is_rate_limit or is_unavailable) and attempt >= 2 and current_model != FALLBACK_MODEL:
                print(f"  [{chunk_label}] Switching to fallback model '{FALLBACK_MODEL}'...", file=sys.stderr)
                current_model = FALLBACK_MODEL

            print(f"  [{chunk_label}] attempt {attempt} failed ({e}) -- retrying in {int(wait)}s", file=sys.stderr)
            if attempt == max_retries:
                raise
            time.sleep(wait)


def extract_all_parallel(client, reader: PdfReader, pages_per_participant=2, model_name=DEFAULT_MODEL, max_workers=DEFAULT_MAX_WORKERS, progress_callback=None, max_participants=None):
    """
    Extracts all participant feedback forms from a PDF in parallel using concurrent workers.
    Returns ordered list of extracted participant records.
    """
    n_pages = len(reader.pages)
    step = pages_per_participant
    chunk_starts = list(range(0, n_pages, step))
    if max_participants:
        chunk_starts = chunk_starts[:max_participants]

    total_chunks = len(chunk_starts)
    records = [None] * total_chunks

    # Pre-encode all chunks in memory
    encoded_chunks = []
    for idx, start in enumerate(chunk_starts):
        idxs = list(range(start, min(start + step, n_pages)))
        label = f"Participant {idx+1}/{total_chunks} (pages {idxs[0]+1}-{idxs[-1]+1})"
        chunk_pdf = encode_pdf_chunk(reader, idxs)
        encoded_chunks.append((idx, chunk_pdf, label))

    def _worker(item):
        chunk_idx, pdf_data, lbl = item
        rec = call_gemini(client, pdf_data, lbl, model_name)
        return chunk_idx, rec, lbl

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = {executor.submit(_worker, chunk_item): chunk_item for chunk_item in encoded_chunks}
        for future in as_completed(futures):
            item = futures[future]
            try:
                chunk_idx, rec, lbl = future.result()
                records[chunk_idx] = rec
                if progress_callback:
                    progress_callback(chunk_idx + 1, total_chunks, rec, None)
            except Exception as exc:
                print(f"  [ERROR] Failed {item[2]}: {exc}", file=sys.stderr)
                if progress_callback:
                    progress_callback(item[0] + 1, total_chunks, None, str(exc))

    # Filter out any failed None records
    valid_records = [r for r in records if r is not None]
    return valid_records


def main():
    ap = argparse.ArgumentParser(description="Extract handwritten feedback form data from PDF using Parallel Gemini Vision.")
    ap.add_argument("pdf_path", help="Path to input PDF file")
    ap.add_argument("--pages-per-participant", type=int, default=2, help="Number of pages per participant form (default: 2)")
    ap.add_argument("--out", default="extracted_data.json", help="Output JSON path (default: extracted_data.json)")
    ap.add_argument("--model", default=DEFAULT_MODEL, help=f"Gemini model identifier (default: {DEFAULT_MODEL})")
    ap.add_argument("--api-key", default=None, help="Gemini API Key (default: GEMINI_API_KEY env var)")
    ap.add_argument("--workers", type=int, default=DEFAULT_MAX_WORKERS, help=f"Concurrent parallel workers (default: {DEFAULT_MAX_WORKERS})")
    ap.add_argument("--max-participants", type=int, default=None, help="Limit number of participants to process")
    args = ap.parse_args()

    if not os.path.exists(args.pdf_path):
        sys.exit(f"Error: PDF file not found: {args.pdf_path}")

    api_key = args.api_key or (config.GEMINI_API_KEY if config else None) or os.environ.get("GEMINI_API_KEY")
    if not api_key:
        sys.exit("Error: GEMINI_API_KEY is not configured. Set it in config.py, .env, or pass --api-key.")

    client = genai.Client(api_key=api_key, vertexai=False)
    reader = PdfReader(args.pdf_path)
    n = len(reader.pages)
    step = args.pages_per_participant

    print(f"[START] Starting High-Speed Parallel Vision Extraction...")
    print(f"  PDF Pages: {n} ({n // step} participant forms) | Parallel Workers: {args.workers} | Model: {args.model}")
    start_t = time.time()

    def on_progress(curr, total, rec, err):
        if rec:
            name = rec.get("name") or "Staff"
            staff_no = rec.get("staff_no") or "?"
            print(f"  [OK] [{curr:02d}/{total:02d}] Extracted: {name:<25} (Staff No: {staff_no})")
        else:
            print(f"  [FAIL] [{curr:02d}/{total:02d}] Failed: {err}")

    records = extract_all_parallel(
        client=client,
        reader=reader,
        pages_per_participant=step,
        model_name=args.model,
        max_workers=args.workers,
        progress_callback=on_progress,
        max_participants=args.max_participants,
    )

    elapsed = time.time() - start_t
    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(records, f, indent=2, ensure_ascii=False)

    print(f"\n[DONE] Completed in {elapsed:.1f}s! Wrote {len(records)} participant records -> {args.out}")


if __name__ == "__main__":
    main()

