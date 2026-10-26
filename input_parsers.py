"""Read the five department datasets from Excel (.xlsx/.xls), CSV, PDF or manually typed text.

Every function returns a pandas DataFrame with the standard column names used by the rest of the
application, so departments.create_department() does not care where the data came from.
"""
import io
import re

import pandas as pd

SUBJECT_COLS = ["Subject Code", "Subject Name", "Semester", "Credits", "Schema", "Department", "Year"]
CO_COLS = ["Subject Code", "CO ID", "Course Outcome"]
OUTCOME_COLS = {
    "po": ("PO ID", "Programme Outcomes", "PO"),
    "peo": ("PEO ID", "Programme Educational Objectives", "PEO"),
    "pso": ("PSO ID", "Programme Specific Outcomes", "PSO"),
}
KINDS = ("subjects", "co", "po", "peo", "pso")

ALLOWED_EXT = (".xlsx", ".xls", ".csv", ".pdf")
CODE_RE = r"[A-Z]{1,6}\s?\d{1,3}[A-Z]{0,3}\d{0,2}[A-Z]?"      # BCS301, 21CS51 handled below, BCS456B ...
CODE_RE2 = r"\d{2}[A-Z]{2,4}\d{1,3}[A-Z]?"                      # 21CS51, 22MAT11
ANY_CODE = rf"(?:{CODE_RE2}|{CODE_RE})"


class ParseError(Exception):
    pass


# ------------------------------------------------------------------ helpers
def _clean(text):
    return re.sub(r"\s+", " ", str(text or "")).strip()


def _norm_id(prefix, number):
    return f"{prefix}{int(number)}"


def _pdf_text(stream):
    from pypdf import PdfReader
    try:
        reader = PdfReader(stream)
        pages = [(p.extract_text() or "") for p in reader.pages]
    except Exception as exc:
        raise ParseError(f"could not read the PDF ({exc}).")
    text = "\n".join(pages)
    if not text.strip():
        raise ParseError("the PDF has no selectable text (it looks like a scan). Use a text-based PDF, an Excel file, or type the data manually.")
    return text


def _find(df, names):
    wanted = [re.sub(r"[^a-z0-9]", "", n.lower()) for n in names]
    for c in df.columns:
        if re.sub(r"[^a-z0-9]", "", str(c).lower()) in wanted:
            return c
    return None


def normalize_frame(kind, df, default_codes=None):
    """Give any spreadsheet the standard column names. Falls back to column POSITION when the
    headers are different (or missing), and numbers outcomes automatically when IDs are absent."""
    df = df.copy().fillna("")
    df.columns = [str(c).strip() for c in df.columns]
    df = df.loc[:, [not str(c).startswith("Unnamed") or df[c].astype(str).str.strip().ne("").any() for c in df.columns]]
    ncols = len(df.columns)
    if ncols == 0 or df.empty:
        return df
    default_codes = [c for c in (default_codes or []) if c]
    if kind == "subjects":
        c_code = _find(df, ["Subject Code", "Course Code", "Code"])
        c_name = _find(df, ["Subject Name", "Course Name", "Name", "Subject", "Course Title", "Title"])
        if c_code and c_name:
            return df
        if ncols == 1:
            names = df.iloc[:, 0].astype(str)
            return pd.DataFrame({"Subject Code": [f"SUB{i}" for i in range(1, len(names) + 1)], "Subject Name": names})
        return df.rename(columns={df.columns[0]: "Subject Code", df.columns[1]: "Subject Name"})
    if kind == "co":
        c_code = _find(df, ["Subject Code", "Course Code", "Code"])
        c_id = _find(df, ["CO ID", "Course Outcome ID", "CO", "CO No", "CO Number"])
        c_text = _find(df, ["Course Outcome", "Course Outcomes", "CO Statement", "Description", "Statement", "Outcome"])
        if c_code and c_id and c_text:
            return df
        cols = list(df.columns)
        if ncols >= 3:
            out = pd.DataFrame({"Subject Code": df[cols[0]], "CO ID": df[cols[1]], "Course Outcome": df[cols[2]]})
        elif ncols == 2:
            code = default_codes[0] if len(default_codes) == 1 else ""
            out = pd.DataFrame({"Subject Code": code, "CO ID": df[cols[0]], "Course Outcome": df[cols[1]]})
        else:
            code = default_codes[0] if len(default_codes) == 1 else ""
            out = pd.DataFrame({"Subject Code": code, "CO ID": [f"CO{i}" for i in range(1, len(df) + 1)], "Course Outcome": df[cols[0]]})
        out["CO ID"] = [v if re.search(r"\d", str(v)) and re.match(r"(?i)\s*co", str(v)) else (f"CO{re.sub(r'[^0-9]', '', str(v)) or i}")
                        for i, v in enumerate(out["CO ID"], 1)]
        return out
    id_col, text_col, prefix = OUTCOME_COLS[kind]
    c_id = _find(df, [id_col, prefix, f"{prefix} No", f"{prefix} Number"])
    c_text = _find(df, [text_col, f"{prefix} Description", f"{prefix} Statement", "Description", "Statement", "Outcome", "Objective"])
    if c_id and c_text:
        return df
    cols = list(df.columns)
    if ncols >= 2:
        ids, texts = df[cols[0]], df[cols[1]]
    else:
        ids, texts = pd.Series(range(1, len(df) + 1)), df[cols[0]]
    ids = [f"{prefix}{re.sub(r'[^0-9]', '', str(v)) or i}" for i, v in enumerate(ids, 1)]
    return pd.DataFrame({id_col: ids, text_col: texts.astype(str).tolist()})


def _read_excel_like(stream, filename):
    name = (filename or "").lower()
    if name.endswith(".csv"):
        return pd.read_csv(stream).fillna("")
    engine = "xlrd" if name.endswith(".xls") else "openpyxl"
    try:
        return pd.read_excel(stream, sheet_name=0, engine=engine).fillna("")
    except ImportError:
        raise ParseError("reading .xls needs the 'xlrd' package: python -m pip install xlrd")


# ------------------------------------------------------------- text parsers
def _split_inline(text, prefix):
    """Put each 'CO1:' / 'PO2.' style marker on its own line, so a whole list typed or pasted on
    ONE line ("CO1: ... CO2: ... CO3: ...") is read item by item."""
    return re.sub(rf"(?<![A-Za-z0-9])({prefix}\s*[-_]?\s*\d+\s*[:.\-–—)])", r"\n\1", text.replace("\r", ""), flags=re.I)


def parse_outcomes_text(text, kind):
    """PO / PEO / PSO text.  Accepts 'PO1: text', 'PO 1 - text', 'PO1 | text', 'PO1<TAB>text'
    and statements that continue over several lines."""
    id_col, text_col, prefix = OUTCOME_COLS[kind]
    text = _split_inline(text, prefix)
    pattern = re.compile(rf"^\s*{prefix}\s*[-_ ]?\s*(\d+)\s*(?:[:.\-–—)|\t]|\s)\s*(.*)$", re.I)
    items, current = [], None
    for raw in text.replace("\r", "").split("\n"):
        line = raw.strip()
        if not line:
            continue
        m = pattern.match(line)
        if m:
            current = [_norm_id(prefix, m.group(1)), m.group(2).strip()]
            items.append(current)
        elif current is not None:
            if re.fullmatch(r"\d+", line):      # stray page number
                continue
            current[1] = (current[1] + " " + line).strip()
    items = [(i, _clean(t)) for i, t in items if _clean(t)]
    if not items:
        # No IDs typed at all: every non-empty line is the next outcome (IDs are numbered automatically)
        header = re.compile(r"^(programme|program)\b.{0,40}(outcomes?|objectives?)\s*:?$", re.I)
        lines = [_clean(l) for l in text.replace("\r", "").split("\n")]
        lines = [l for l in lines if l and not header.match(l) and not re.fullmatch(r"\d+", l)]
        items = [(f"{prefix}{n}", l) for n, l in enumerate(lines, 1)]
    if not items:
        raise ParseError(f"nothing typed for {prefix}. Write one {prefix} statement per line.")
    return pd.DataFrame(items, columns=[id_col, text_col])


def parse_subjects_text(text, department=""):
    """One subject per line, any style:  'CODE | Name | Sem | Credits | Schema', tab separated,
    'CODE - Name', 'CODE Name', or just a name (a code SUB1, SUB2 ... is then generated)."""
    rows, auto = [], 0
    for raw in text.replace("\r", "").split("\n"):
        line = raw.strip()
        if not line:
            continue
        if "|" in line or "\t" in line:
            parts = [p.strip() for p in re.split(r"\||\t", line)]
            if len(parts) == 1 or not parts[1]:
                parts = [parts[0], parts[0]]
        else:
            m = re.match(rf"^({ANY_CODE})\s*[-–—:,]?\s+(.+)$", line, re.I) or re.match(rf"^({ANY_CODE})\s*[-–—:,]\s*(.+)$", line, re.I)
            if m:
                code, rest = m.group(1), m.group(2)
                sem = cred = ""
                tail = re.search(r"\s+(\d(?:st|nd|rd|th)?)\s+(\d{1,2})$", rest, re.I)
                if tail:
                    sem, cred, rest = tail.group(1), tail.group(2), rest[:tail.start()]
                parts = [code, rest, sem, cred]
            else:
                auto += 1
                parts = [f"SUB{auto}", line]
        if re.match(r"^(sub(ject)?|course)\s*code$", parts[0], re.I):   # header row
            continue
        parts = (parts + [""] * 6)[:6]
        if not parts[0]:
            continue
        rows.append([parts[0].upper().replace(" ", ""), parts[1] or parts[0], parts[2], parts[3], parts[4], department, parts[5]])
    if not rows:
        raise ParseError("nothing typed for Subjects. Write one subject per line (code and name).")
    return pd.DataFrame(rows, columns=SUBJECT_COLS)


def parse_co_text(text, default_codes=None):
    """Course outcomes.  Either one per line  'BCS301 | CO1 | statement'  or grouped:

        Subject Code: BCS301          (or just a line starting with the code)
        CO1: statement ...
        CO2: statement ...
    Statements may continue over several lines."""
    co_re = re.compile(r"^\s*CO\s*[-_ ]?(\d+)\s*(?:[:.\-–—)|\t]|\s)\s*(.*)$", re.I)
    full_re = re.compile(rf"^\s*({ANY_CODE})\s*(?:[|\t,:\-–]|\s)\s*CO\s*[-_ ]?(\d+)\s*(?:[|\t:.\-–—)]|\s)\s*(.*)$", re.I)
    head_re = re.compile(rf"(?:sub(?:ject)?|course)?\s*code\s*[:\-]?\s*({ANY_CODE})\b|^\s*({ANY_CODE})\b", re.I)
    text = _split_inline(text, "CO")
    default_codes = [c for c in (default_codes or []) if c]
    # CO lines typed with no subject code belong to the only subject (if there is just one)
    rows, code, current = [], (default_codes[0] if len(default_codes) == 1 else ""), None
    first_line = ""
    for raw in text.replace("\r", "").split("\n"):
        line = raw.strip()
        if not line:
            continue
        first_line = first_line or line
        m = full_re.match(line)
        if m:
            code = m.group(1).upper().replace(" ", "")
            current = [code, _norm_id("CO", m.group(2)), m.group(3).strip()]
            rows.append(current)
            continue
        m = co_re.match(line)
        if m and code:
            current = [code, _norm_id("CO", m.group(1)), m.group(2).strip()]
            rows.append(current)
            continue
        m = head_re.search(line)
        if m and not co_re.match(line) and len(line) < 120 and not line.lower().startswith(("co", "course outcome")):
            code = (m.group(1) or m.group(2)).upper().replace(" ", "")
            current = None
            continue
        if current is not None and not re.fullmatch(r"\d+", line):
            current[2] = (current[2] + " " + line).strip()
    rows = [[c, i, _clean(t)] for c, i, t in rows if _clean(t)]
    if not rows and len(default_codes) == 1:
        # No 'CO1' markers: every non-empty line is the next course outcome of the only subject
        lines = [_clean(l) for l in text.replace("\r", "").split("\n")]
        lines = [l for l in lines if l and not re.fullmatch(r"\d+", l)]
        rows = [[default_codes[0], f"CO{n}", l] for n, l in enumerate(lines, 1)]
    if not rows:
        hint = ""
        if re.search(r"^\s*CO\s*[-_ ]?\d+", text, re.I | re.M) and not default_codes:
            hint = " Your CO lines have no subject code - add a line 'Subject Code: BCS301' above them, or start each line with the code."
        raise ParseError(
            "no course outcomes found (first line read: '" + first_line[:60] + "')." + hint +
            " Accepted: 'BCS301 | CO1 | statement' on each line, or 'Subject Code: BCS301' followed by 'CO1: statement' lines.")
    return pd.DataFrame(rows, columns=CO_COLS)


def parse_text(kind, text, department="", default_codes=None):
    if kind == "subjects":
        return parse_subjects_text(text, department)
    if kind == "co":
        return parse_co_text(text, default_codes)
    return parse_outcomes_text(text, kind)


# ----------------------------------------------------------------- public API
def load_dataset(kind, file_storage=None, typed_text="", department="", default_codes=None):
    """Return a DataFrame for one dataset from an uploaded file OR typed text.
    Raises ParseError with a user-friendly message."""
    label = kind.upper() if kind != "subjects" else "Subjects"
    if typed_text and typed_text.strip():
        try:
            return parse_text(kind, typed_text, department, default_codes)
        except ParseError as exc:
            raise ParseError(f"{label} (typed): {exc}")
    if not file_storage or not file_storage.filename:
        raise ParseError(f"{label}: upload a file or type the data.")
    fname = file_storage.filename
    if not fname.lower().endswith(ALLOWED_EXT):
        raise ParseError(f"{label}: '{fname}' is not supported. Use .xlsx, .xls, .csv or .pdf.")
    data = file_storage.read()
    try:
        if fname.lower().endswith(".pdf"):
            return parse_text(kind, _pdf_text(io.BytesIO(data)), department, default_codes)
        return normalize_frame(kind, _read_excel_like(io.BytesIO(data), fname), default_codes)
    except ParseError as exc:
        raise ParseError(f"{label} ({fname}): {exc}")
    except Exception as exc:
        raise ParseError(f"{label} ({fname}): could not read the file ({exc}).")
