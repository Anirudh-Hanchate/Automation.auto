"""
fill_excel.py
-------------
Takes the JSON produced by extract_feedback.py / extract_feedback_gemini.py
(one record per participant) and writes it into the correct cells of an existing
"Participant's Feedback" sheet, using the sheet's own headers to figure out
which row/column each answer belongs in (see sheet_map.py). It never invents
new rows/columns for scores, so it will not silently corrupt a sheet whose
layout has changed -- it raises a clear error instead.

Usage:
    python fill_excel.py extracted_data.json blank_template.xlsx "Mod-1 Participant's Feedback" --out filled.xlsx

Record schema expected per participant:
{
  "staff_no": "218399",
  "name": "Pavan C Pawar",
  "days": [
    {"day": 1, "topics": [
        {"topic_index": 1, "criteria": {"Course Content": 10, "Structure & Flow": 8, ...}},
        {"topic_index": 2, "criteria": {...}}
    ]},
    ...
  ],
  "overall": {"Delivery Notes (Course Material)": 9, ...},
  "suggestion_1": "None",
  "suggestion_2": "Some topics were rushed..."
}
"""
import argparse
import json
import os
import sys

import openpyxl
from sheet_map import build_sheet_map, find_criteria_row, find_overall_row, _norm


def fill_participant(ws, smap, record, overwrite=False, auto_add_column=True):
    staff_no = str(record["staff_no"]).strip()
    col = None

    if staff_no in smap.staff_col:
        col = smap.staff_col[staff_no]
    else:
        # Check if float/int representation matches
        matching_key = None
        for k in smap.staff_col:
            if k == staff_no or k.lstrip("0") == staff_no.lstrip("0"):
                matching_key = k
                break
        if matching_key:
            col = smap.staff_col[matching_key]
        elif auto_add_column:
            # Dynamically insert a new column for this staff member
            from openpyxl.utils import get_column_letter
            insert_at = smap.last_data_col + 1
            ws.insert_cols(idx=insert_at, amount=1)
            ws.cell(row=smap.staff_row, column=insert_at, value=staff_no)
            col = get_column_letter(insert_at)
            smap.staff_col[staff_no] = col
            smap.last_data_col = insert_at
        else:
            raise KeyError(
                f"Staff No {staff_no} has no column in this sheet. "
                f"Available columns: {list(smap.staff_col.keys())}. "
                f"Add them to row {smap.staff_row} first, then re-run."
            )

    written_scores = 0
    written_overall = 0
    written_suggestions = 0


    # 1. Fill Day / Topic / Criteria scores
    for day in record.get("days", []):
        day_num = day.get("day")
        for topic in day.get("topics", []):
            topic_idx = topic.get("topic_index", 1)
            for criteria_name, score in topic.get("criteria", {}).items():
                if score is None or score == "":
                    continue
                try:
                    score_val = int(score) if str(score).isdigit() else score
                except Exception:
                    score_val = score

                row = None
                try:
                    row = find_criteria_row(smap, day_num, topic_idx, criteria_name)
                except ValueError:
                    # Fallback: search across all topics of that day
                    day_rows = [cr for cr in smap.criteria_rows if _norm(cr.day) == f"day{day_num}"]
                    norm_crit = _norm(criteria_name)
                    for cr in day_rows:
                        if _norm(cr.criteria) == norm_crit or norm_crit in _norm(cr.criteria) or _norm(cr.criteria) in norm_crit:
                            row = cr.row
                            break

                if row is None:
                    continue

                cell = ws[f"{col}{row}"]
                if cell.value is not None and not overwrite:
                    raise ValueError(
                        f"Cell {col}{row} already has a value ({cell.value!r}); "
                        f"pass --overwrite to replace it."
                    )
                cell.value = score_val
                written_scores += 1

    # 2. Fill Overall Feedback scores (if sheet has an overall section)
    if smap.overall_rows and record.get("overall"):
        for crit_name, score in record["overall"].items():
            if score is None or score == "":
                continue
            try:
                score_val = int(score) if str(score).isdigit() else score
                row = find_overall_row(smap, crit_name)
                cell = ws[f"{col}{row}"]
                if cell.value is not None and not overwrite:
                    raise ValueError(
                        f"Cell {col}{row} (overall) already has a value ({cell.value!r}); "
                        f"pass --overwrite to replace it."
                    )
                cell.value = score_val
                written_overall += 1
            except ValueError:
                pass

    # 3. Fill Suggestions
    if smap.suggestion1_row and record.get("suggestion_1") is not None:
        s1 = str(record["suggestion_1"]).strip()
        if s1:
            ws[f"{col}{smap.suggestion1_row}"] = s1
            written_suggestions += 1

    if smap.suggestion2_row and record.get("suggestion_2") is not None:
        s2 = str(record["suggestion_2"]).strip()
        if s2:
            ws[f"{col}{smap.suggestion2_row}"] = s2
            written_suggestions += 1

    return {
        "scores": written_scores,
        "overall": written_overall,
        "suggestions": written_suggestions,
        "total": written_scores + written_overall + written_suggestions,
    }


def main():
    ap = argparse.ArgumentParser(description="Fill Excel feedback sheet with extracted participant JSON records.")
    ap.add_argument("json_path", help="Path to extracted_data.json")
    ap.add_argument("xlsx_path", help="Path to workbook xlsx")
    ap.add_argument("sheet_name", nargs="?", default="Mod-1 Participant's Feedback", help="Target sheet name")
    ap.add_argument("--out", default=None, help="Output path (default: overwrite in place or specify path)")
    ap.add_argument("--overwrite", action="store_true", help="Overwrite non-empty cells")
    ap.add_argument("--dry-run", action="store_true", help="Validate without saving changes to disk")
    args = ap.parse_args()

    if not os.path.exists(args.json_path):
        sys.exit(f"Error: JSON file not found: {args.json_path}")
    if not os.path.exists(args.xlsx_path):
        sys.exit(f"Error: Excel workbook not found: {args.xlsx_path}")

    with open(args.json_path, encoding="utf-8") as f:
        records = json.load(f)

    wb = openpyxl.load_workbook(args.xlsx_path)
    if args.sheet_name not in wb.sheetnames:
        sys.exit(f"Error: Sheet '{args.sheet_name}' not found. Available sheets: {wb.sheetnames}")

    ws = wb[args.sheet_name]
    smap = build_sheet_map(ws)

    print(f"Loaded sheet '{args.sheet_name}' with {len(smap.staff_col)} participant columns.")
    print(f"Criteria rows mapped: {len(smap.criteria_rows)}, Overall rows: {len(smap.overall_rows)}")
    print(f"Suggestion 1 row: {smap.suggestion1_row}, Suggestion 2 row: {smap.suggestion2_row}\n")

    total_scores = 0
    total_overall = 0
    total_suggestions = 0

    for rec in records:
        name_label = rec.get("name") or rec.get("staff_no")
        res = fill_participant(ws, smap, rec, overwrite=args.overwrite)
        total_scores += res["scores"]
        total_overall += res["overall"]
        total_suggestions += res["suggestions"]
        print(f"  Participant {name_label} ({rec.get('staff_no')}): {res['scores']} scores, {res['overall']} overall, {res['suggestions']} suggestions")

    print(f"\nSummary:")
    print(f"  Participants processed: {len(records)}")
    print(f"  Score cells filled     : {total_scores}")
    print(f"  Overall cells filled   : {total_overall}")
    print(f"  Suggestions filled     : {total_suggestions}")
    print(f"  Total cells written    : {total_scores + total_overall + total_suggestions}")

    if not args.dry_run:
        out = args.out or args.xlsx_path
        wb.save(out)
        print(f"  Saved workbook -> {out}")
    else:
        print("  Dry-run complete. No file was saved.")


if __name__ == "__main__":
    main()

