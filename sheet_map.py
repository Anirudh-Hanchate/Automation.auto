"""
sheet_map.py
------------
Reads the *structure* of a "Participant's Feedback" sheet directly from the
workbook itself (day blocks, topic blocks, criteria rows, overall feedback,
staff-no columns), so the filler script never hardcodes row/column numbers.
That means the same code keeps working even if BAE Academy changes the topic
order, adds a day, or adds/removes a criterion, as long as they keep filling
the same style of sheet.

Layout this expects:
    Row 4/5 (merged)      : "STAFF NO" header, staff numbers start in col D
    Col A (merged blocks)  : "DAY 1\n(06-07-2026)" etc. or "OVERALL FEEDBACK"
    Col B (merged blocks)  : "Topic name\n(Faculty)"
    Col C                  : criteria name, one per row
    Col D..V (or wherever) : one column per participant, holds the 1-10 score
    Row with "Suggestion 1": Q2 answer ("other programs you want us to add")
    Row with "Suggestion 2": Q3 answer ("any suggestion for future program")
"""
import re
from dataclasses import dataclass, field
from openpyxl.utils import get_column_letter


@dataclass
class CriteriaCell:
    day: str
    day_index: int
    topic_index: int
    topic: str
    criteria: str
    row: int


@dataclass
class OverallCell:
    criteria: str
    row: int


@dataclass
class SheetMap:
    staff_col: dict          # {staff_no (str): column_letter}
    staff_row: int
    first_data_col: int
    last_data_col: int
    criteria_rows: list       # list[CriteriaCell], in sheet order
    overall_rows: list        # list[OverallCell], in sheet order
    suggestion1_row: int | None
    suggestion2_row: int | None


def _clean(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v)).strip()
    return str(v).strip()


def _norm(s: str) -> str:
    """Normalize string for fuzzy-matching headers/criteria."""
    if not s:
        return ""
    return (
        s.lower()
        .replace("contenet", "content")
        .replace("impresion", "impression")
        .replace("&", "and")
        .replace("-", "")
        .replace("_", "")
        .replace("'", "")
        .replace('"', "")
        .replace(" ", "")
        .replace("\n", "")
        .replace("\r", "")
        .strip()
    )


def build_sheet_map(ws) -> SheetMap:
    # 1. Find the header row containing "STAFF NO" in column A-E area
    header_row = None
    for r in range(1, 10):
        for c in range(1, 6):
            val = _clean(ws.cell(row=r, column=c).value).upper()
            if "STAFF NO" in val:
                header_row = r
                break
        if header_row:
            break
    if header_row is None:
        raise ValueError("Could not locate the 'STAFF NO' header row in this sheet")

    # The actual staff-number values may sit on this same row (to the right of the
    # label) or on the row directly below it (if "STAFF NO" is a merged label row).
    # Pick whichever of the two has more numeric-looking entries in columns D..J.
    def numeric_count(row_idx):
        n = 0
        for c in range(4, 15):
            v = ws.cell(row=row_idx, column=c).value
            if v is None:
                continue
            s = _clean(v)
            if s.replace("-", "").replace("N", "").replace("K", "").isdigit() or isinstance(v, (int, float)):
                n += 1
        return n

    values_row = header_row if numeric_count(header_row) >= numeric_count(header_row + 1) else header_row + 1

    # 2. Walk columns D.. until we hit an empty run of 3+ columns
    staff_col = {}
    empty_run = 0
    c = 4
    last_data_col = 3
    while empty_run < 3 and c < 200:
        v = ws.cell(row=values_row, column=c).value
        if v is None or _clean(v) == "":
            empty_run += 1
        else:
            empty_run = 0
            staff_col[_clean(v)] = get_column_letter(c)
            last_data_col = c
        c += 1

    # 3. Walk down col A/B/C from just below the header to build day/topic/criteria rows
    criteria_rows = []
    overall_rows = []
    current_day = None
    current_topic = None
    current_day_index = 0
    current_topic_index = 0
    is_overall = False
    suggestion1_row = suggestion2_row = None
    r = values_row + 1
    max_row = ws.max_row

    while r <= max_row:
        a = _clean(ws.cell(row=r, column=1).value)
        b = _clean(ws.cell(row=r, column=2).value)
        cc = _clean(ws.cell(row=r, column=3).value)

        if a:
            if "OVERALL" in a.upper():
                is_overall = True
                current_day = "OVERALL"
                current_day_index = 999
                current_topic = "OVERALL"
                current_topic_index = 1
            else:
                is_overall = False
                current_day = a.split("\n")[0].strip()
                day_match = re.search(r"\d+", current_day)
                if day_match:
                    current_day_index = int(day_match.group())
                else:
                    current_day_index += 1
                current_topic_index = 0

        if b:
            b_upper = b.upper()
            if "SUGGESTION 1" in b_upper or "SUGGESTION1" in b_upper:
                suggestion1_row = r
            elif "SUGGESTION 2" in b_upper or "SUGGESTION2" in b_upper:
                suggestion2_row = r
            elif not is_overall:
                current_topic_index += 1
                current_topic = b.split("\n")[0].strip()

        if cc:
            cc_upper = cc.upper()
            # Ignore rating scale / legend rows if present in Col C at the bottom
            is_scale_label = any(cc_upper.startswith(kw) for kw in ["EXCELLENT", "VERY GOOD", "GOOD", "SATISFACTORY", "UNSATISFACTORY"])
            if not is_scale_label:
                if is_overall or (current_day and "OVERALL" in current_day.upper()):
                    overall_rows.append(OverallCell(cc, r))
                elif current_day and not (b and "SUGGESTION" in b.upper()):
                    criteria_rows.append(CriteriaCell(
                        day=current_day,
                        day_index=current_day_index,
                        topic_index=current_topic_index or 1,
                        topic=current_topic or "",
                        criteria=cc,
                        row=r
                    ))

        r += 1

    return SheetMap(
        staff_col=staff_col,
        staff_row=values_row,
        first_data_col=4,
        last_data_col=last_data_col,
        criteria_rows=criteria_rows,
        overall_rows=overall_rows,
        suggestion1_row=suggestion1_row,
        suggestion2_row=suggestion2_row,
    )


def find_criteria_row(smap: SheetMap, day_index: int, topic_index: int, criteria_name: str) -> int:
    """
    day_index: 1-based day number (1 = DAY 1, 6 = DAY-6)
    topic_index: 1-based position of the topic block WITHIN that day
    criteria_name: e.g. 'Course Content' (matched case/space-insensitively, ignoring typos)
    """
    day_rows = [cr for cr in smap.criteria_rows if cr.day_index == day_index or _norm(cr.day) == f"day{day_index}"]
    if not day_rows:
        raise ValueError(f"Day {day_index} not found in sheet")

    # Filter by topic_index within that day
    topic_rows = [cr for cr in day_rows if cr.topic_index == topic_index]
    if not topic_rows:
        topic_rows = day_rows

    norm_crit = _norm(criteria_name)

    # 1. Exact normalized match
    for cr in topic_rows:
        if _norm(cr.criteria) == norm_crit:
            return cr.row

    # 2. Substring match
    for cr in topic_rows:
        nc = _norm(cr.criteria)
        if norm_crit in nc or nc in norm_crit:
            return cr.row

    # 3. Keyword / Token overlap match
    crit_words = set(re.findall(r"\w+", criteria_name.lower().replace("contenet", "content").replace("impresion", "impression")))
    best_row = None
    best_overlap = 0
    for cr in topic_rows:
        row_words = set(re.findall(r"\w+", cr.criteria.lower().replace("contenet", "content").replace("impresion", "impression")))
        overlap = len(crit_words & row_words)
        if overlap > best_overlap and overlap >= 2:
            best_overlap = overlap
            best_row = cr.row

    if best_row is not None:
        return best_row

    raise ValueError(f"Criteria '{criteria_name}' not found for Day {day_index} / topic #{topic_index}")


def find_overall_row(smap: SheetMap, criteria_name: str) -> int:
    """
    Finds row index for an overall feedback criterion (e.g. 'Classroom Facilities', 'Delivery Notes').
    """
    norm_crit = _norm(criteria_name)
    for oc in smap.overall_rows:
        if _norm(oc.criteria) == norm_crit:
            return oc.row
    for oc in smap.overall_rows:
        if norm_crit in _norm(oc.criteria) or _norm(oc.criteria) in norm_crit:
            return oc.row
    raise ValueError(f"Overall criteria '{criteria_name}' not found in sheet")

