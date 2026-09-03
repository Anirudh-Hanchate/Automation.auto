"""
extract_feedback.py
--------------------
Reads the scanned feedback-form PDF (one participant = N consecutive pages,
default 2) and uses Anthropic's Claude API vision capability to read handwritten
circles, ticks, and checkmarks on each form, outputting structured JSON.

Output: extracted_data.json, a list of participant records in the schema
fill_excel.py expects.

Setup:
    pip install -r requirements.txt
    export ANTHROPIC_API_KEY=sk-ant-... (or create a .env file with ANTHROPIC_API_KEY=...)

Usage:
    python extract_feedback.py CRE_Feedback.pdf --pages-per-participant 2 --out extracted_data.json
"""
import argparse
import base64
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
    import anthropic
except ImportError:
    sys.exit("Error: anthropic not found. Run: pip install anthropic")

DEFAULT_MODEL = getattr(config, "DEFAULT_CLAUDE_MODEL", "claude-3-7-sonnet-latest") if config else "claude-3-7-sonnet-latest"


SYSTEM_PROMPT = """You are transcribing a scanned, hand-filled training feedback form.
Each topic block has several criteria rows. Each row is a printed 10-9-8-7-6-5-4-3-2-1
scale and the participant circled, ticked, or underlined exactly ONE number per row --
that circled/ticked number is the answer, not the leftmost or rightmost number.
Read the actual mark carefully; marks are sometimes messy or offset from the number.

Return ONLY valid JSON (no markdown code blocks, no preamble, no commentary) matching exactly this schema:

{
  "staff_no": "<string, as printed on the form, keep leading zeros / letters exactly>",
  "name": "<participant name>",
  "sbu_csg": "<string or null>",
  "department": "<string or null>",
  "days": [
    {
      "day": <integer day number, 1-based>,
      "topics": [
        {
          "topic_index": <integer, 1-based position of this topic block within THIS day>,
          "topic_name": "<topic name as printed>",
          "criteria": {"<criteria name as printed, e.g. Course Content>": <integer 1-10>, ...}
        }
      ]
    }
  ],
  "overall": {"<overall feedback row label>": <integer 1-10>, ...},
  "recommend_continue": "Yes" | "No" | null,
  "suggestion_1": "<answer to 'other quality training programs to add', or empty string>",
  "suggestion_2": "<answer to 'any suggestion for future program', or empty string>"
}

Rules:
- Only include days/topics/criteria that actually appear on the pages given.
- If a field is blank/illegible, use null (for numbers) or "" (for text) -- never guess.
- Preserve criteria names exactly as printed on the form (including any typos like "Contenet").
- staff_no must be a plain string exactly as written, including any letters/dashes.
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


def call_claude(client, pdf_bytes: bytes, chunk_label: str, model_name: str, max_retries=3):
    b64 = base64.standard_b64encode(pdf_bytes).decode("utf-8")
    for attempt in range(1, max_retries + 1):
        try:
            resp = client.messages.create(
                model=model_name,
                max_tokens=4000,
                temperature=0,
                system=SYSTEM_PROMPT,
                messages=[{
                    "role": "user",
                    "content": [
                        {"type": "document", "source": {"type": "base64", "media_type": "application/pdf", "data": b64}},
                        {"type": "text", "text": "Transcribe this participant's feedback form into the JSON schema described in the system prompt."},
                    ],
                }],
            )
            text = "".join(b.text for b in resp.content if b.type == "text")
            return extract_json(text)
        except Exception as e:
            print(f"  [{chunk_label}] attempt {attempt} failed: {e}", file=sys.stderr)
            if attempt == max_retries:
                raise
            time.sleep(2 * attempt)


def main():
    ap = argparse.ArgumentParser(description="Extract handwritten feedback form data from PDF using Claude Vision.")
    ap.add_argument("pdf_path", help="Path to input PDF file")
    ap.add_argument("--pages-per-participant", type=int, default=2, help="Number of pages per participant form (default: 2)")
    ap.add_argument("--out", default="extracted_data.json", help="Output JSON path (default: extracted_data.json)")
    ap.add_argument("--model", default=DEFAULT_MODEL, help=f"Claude model identifier (default: {DEFAULT_MODEL})")
    ap.add_argument("--api-key", default=None, help="Anthropic API Key (default: ANTHROPIC_API_KEY env var)")
    ap.add_argument("--max-participants", type=int, default=None, help="Limit number of participants to process")
    args = ap.parse_args()

    if not os.path.exists(args.pdf_path):
        sys.exit(f"Error: PDF file not found: {args.pdf_path}")

    api_key = args.api_key or (config.ANTHROPIC_API_KEY if config else None) or os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        sys.exit("Error: ANTHROPIC_API_KEY is not configured. Set it in config.py, .env, or pass --api-key.")


    client = anthropic.Anthropic(api_key=api_key)
    reader = PdfReader(args.pdf_path)
    n = len(reader.pages)
    step = args.pages_per_participant
    if n % step != 0:
        print(f"WARNING: {n} pages is not a multiple of {step}; last chunk will have fewer pages.")

    records = []
    chunk_starts = list(range(0, n, step))
    if args.max_participants:
        chunk_starts = chunk_starts[:args.max_participants]

    total_chunks = len(chunk_starts)
    print(f"Processing {total_chunks} participant forms ({n} total pages in PDF, {step} pages/participant)...")

    for idx, start in enumerate(chunk_starts, 1):
        idxs = list(range(start, min(start + step, n)))
        label = f"Participant {idx}/{total_chunks} (pages {idxs[0]+1}-{idxs[-1]+1})"
        print(f"Extracting {label} ...")
        chunk_pdf = encode_pdf_chunk(reader, idxs)
        try:
            record = call_claude(client, chunk_pdf, label, args.model)
            staff_no = record.get('staff_no')
            name = record.get('name')
            print(f"  -> Successfully extracted: {name} (Staff No: {staff_no})")
            records.append(record)
        except Exception as e:
            print(f"  FAILED {label}: {e} -- skipping, check manually", file=sys.stderr)

    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(records, f, indent=2, ensure_ascii=False)
    print(f"\nWrote {len(records)} participant records -> {args.out}")


if __name__ == "__main__":
    main()

