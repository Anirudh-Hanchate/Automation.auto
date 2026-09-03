# Feedback Form to Excel Automation

Turns scanned PDFs of hand-filled BAE Academy feedback forms into fully populated `Participant's Feedback` Excel workbooks, eliminating manual data entry.

---

## Architecture & How It Works

Feedback forms contain **hand-drawn circles, ticks, checkmarks, or underlines** next to pre-printed 1-10 rating scales. Standard OCR engines (like Tesseract) cannot reliably distinguish which number was circled by hand. 

This pipeline uses **multimodal vision AI** (Google Gemini Vision or Anthropic Claude Vision) to interpret the handwriting and marks like a human evaluator, and then uses **deterministic Excel mapping** (`openpyxl`) to place every score and suggestion into its exact cell.

```
                  ┌─────────────────────────────────┐
                  │    CRE_Feedback.pdf (scanned)   │
                  └────────────────┬────────────────┘
                                   │
              ┌────────────────────┴────────────────────┐
              │                                         │
    [extract_feedback_gemini.py]              [extract_feedback.py]
         (Google Gemini)                        (Anthropic Claude)
              │                                         │
              └────────────────────┬────────────────────┘
                                   │
                                   ▼
                       ┌───────────────────────┐
                       │  extracted_data.json  │
                       └───────────┬───────────┘
                                   │
                           [fill_excel.py]
                           [sheet_map.py]
                                   │
                                   ▼
                       ┌───────────────────────┐
                       │      filled.xlsx      │
                       └───────────────────────┘
```

---

## Features

- **End-to-End Orchestrator (`run_pipeline.py`)**: Run extraction + validation + Excel filling with a single command.
- **Offline Filling Supported**: Use pre-extracted JSON (`example_extracted_data.json`) to fill workbooks instantly without API calls.
- **Dynamic Header Mapping (`sheet_map.py`)**: Automatically detects Staff No columns, Day/Topic blocks, Criteria rows, and Overall Feedback rows. Fails safely and loudly instead of misaligning cells if sheet layout shifts.
- **Multimodal AI Vision Support**: Supports both **Google Gemini** (`gemini-2.5-flash`, `gemini-3.7-flash`) and **Anthropic Claude** (`claude-3-7-sonnet-latest`).
- **Safe Execution & Validation**: Includes `--dry-run`, `--overwrite` safeguards, and detailed execution summaries.
- **Automatic Formula Sheet Population**: Populates `Participant's Feedback` which automatically flows into `Overall Feedback` and `Suggestions` sheets via Excel formulas.

---

## Quick Start

### 1. Installation

Install required dependencies:

```bash
pip install -r requirements.txt
```

### 2. Launch the Web UI (Recommended)

Start the interactive web application:

```bash
python app.py
```
Open **http://127.0.0.1:8000** in your browser to:
- Drag-and-drop scanned PDF feedback forms from your computer.
- Click **Start Automation & Fill Excel**.
- Download the generated `.xlsx` workbook directly with one click.

---

### 3. Configure API Keys (Optional if using Pre-Extracted Data)

Create a `.env` file or export your API key:

**For Google Gemini:**
```bash
export GEMINI_API_KEY="your-gemini-api-key"
```

**For Anthropic Claude:**
```bash
export ANTHROPIC_API_KEY="your-anthropic-api-key"
```


---

## Usage Examples

### Option A: Complete Pipeline in One Command (`run_pipeline.py`)

#### 1. Fill Excel from pre-extracted JSON (Instant / Offline):
```bash
python run_pipeline.py --json example_extracted_data.json --template blank_template.xlsx --out filled.xlsx
```

#### 2. Extract from PDF and fill Excel using Gemini:
```bash
python run_pipeline.py --pdf CRE_Feedback.pdf --template blank_template.xlsx --out filled.xlsx --engine gemini
```

#### 3. Extract from PDF and fill Excel using Claude:
```bash
python run_pipeline.py --pdf CRE_Feedback.pdf --template blank_template.xlsx --out filled.xlsx --engine claude
```

#### 4. Dry-run validation (checks all column & row mappings without saving):
```bash
python run_pipeline.py --json example_extracted_data.json --template blank_template.xlsx --dry-run
```

---

### Option B: Step-by-Step Execution

#### Step 1: Extract from PDF to JSON

Using **Gemini**:
```bash
python extract_feedback_gemini.py CRE_Feedback.pdf --pages-per-participant 2 --out extracted_data.json
```

Using **Claude**:
```bash
python extract_feedback.py CRE_Feedback.pdf --pages-per-participant 2 --out extracted_data.json
```

#### Step 2: Fill the Excel Workbook

```bash
python fill_excel.py extracted_data.json blank_template.xlsx "Mod-1 Participant's Feedback" --out filled.xlsx
```

To overwrite existing cell values:
```bash
python fill_excel.py extracted_data.json blank_template.xlsx "Mod-1 Participant's Feedback" --out filled.xlsx --overwrite
```

---

## Project Structure

| File | Description |
|---|---|
| `run_pipeline.py` | Unified CLI orchestrator for extraction, validation, and filling |
| `extract_feedback_gemini.py` | PDF $\rightarrow$ JSON extractor using Google Gemini Vision |
| `extract_feedback.py` | PDF $\rightarrow$ JSON extractor using Anthropic Claude Vision |
| `fill_excel.py` | JSON $\rightarrow$ Excel workbook writer |
| `sheet_map.py` | Dynamic sheet structure parser (maps headers, staff columns, and criteria) |
| `blank_template.xlsx` | Excel workbook template containing feedback sheets and formulas |
| `example_extracted_data.json` | Sample extracted JSON records for all 19 participants |
| `CRE_Feedback.pdf` | Scanned handwritten feedback forms (38 pages / 19 participants) |
| `filled.xlsx` | Output generated Excel workbook populated with extracted feedback |
| `requirements.txt` | Python package dependencies |

---

## JSON Data Schema

Each participant record in `extracted_data.json` follows this schema:

```json
{
  "staff_no": "218399",
  "name": "Pavan C Pawar",
  "sbu_csg": "QM/BAE",
  "department": "Quality",
  "days": [
    {
      "day": 1,
      "topics": [
        {
          "topic_index": 1,
          "topic_name": "Reliability Fundamentals",
          "criteria": {
            "Course Content": 10,
            "Structure & Flow": 8,
            "Time Management": 8,
            "Faculty Delivery": 8,
            "Faculty Participants Interaction": 9
          }
        }
      ]
    }
  ],
  "overall": {
    "Delivery Notes (Course Material)": 9,
    "Communication & Overall Coordination": 9,
    "Classroom Facilities": 9,
    "Relevance to your current job": 10
  },
  "recommend_continue": "Yes",
  "suggestion_1": "None",
  "suggestion_2": "Some topics were rushed because of time shortage"
}
```

---

## Verification & Testing

All 19 participants (950 score cells + 20 suggestions) have been round-trip verified against the template. When `filled.xlsx` is opened in Microsoft Excel or LibreOffice, dependent sheets (`Mod-1 Overall Feedback` and `Mod-1Suggestions`) automatically compute their summaries through built-in formulas.

