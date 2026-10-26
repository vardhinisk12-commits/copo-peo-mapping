import os

# Keep the heavy NLP libraries quiet and predictable on Windows (set before they are imported).
os.environ.setdefault("TOKENIZERS_PARALLELISM", "false")
os.environ.setdefault("HF_HUB_DISABLE_TELEMETRY", "1")
os.environ.setdefault("HF_HUB_DISABLE_SYMLINKS_WARNING", "1")

from flask import Flask, render_template, request, send_file, redirect, session, make_response
import pandas as pd
from pypdf import PdfReader, PdfWriter
# ==============================
# REPORTLAB IMPORTS
# ==============================
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.units import inch

from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    PageBreak,
    Image,
    KeepTogether
)
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.graphics.shapes import Drawing, String
from reportlab.graphics.charts.barcharts import VerticalBarChart
from docx.shared import Pt

# ==============================
# OPENPYXL IMPORTS
# ==============================
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.drawing.image import Image as ExcelImage


# ==============================
# PYTHON-DOCX IMPORTS
# ==============================
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT

from flask_sqlalchemy import SQLAlchemy
from werkzeug.security import generate_password_hash, check_password_hash

from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle
)
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape

from docx import Document

from blooms import detect_bloom


# ============================================================
# FLASK CONFIGURATION
# ============================================================

app = Flask(__name__)

app.secret_key = os.environ.get("FLASK_SECRET_KEY", "dev-only-change-this-secret-key")

_db_url = os.environ.get("DATABASE_URL", "sqlite:///users.db")
if _db_url.startswith("postgres://"):          # Render gives postgres://, SQLAlchemy needs postgresql://
    _db_url = _db_url.replace("postgres://", "postgresql://", 1)
if _db_url.startswith("postgresql"):
    # Make sure a PostgreSQL driver is installed; otherwise fall back to SQLite instead of crashing.
    _driver_ok = False
    for _mod in (("psycopg",) if "+psycopg" in _db_url and "+psycopg2" not in _db_url else ("psycopg2",)):
        try:
            __import__(_mod)
            _driver_ok = True
        except Exception:
            pass
    if not _driver_ok:
        for _mod, _prefix in (("psycopg", "postgresql+psycopg://"), ("psycopg2", "postgresql+psycopg2://")):
            try:
                __import__(_mod)
                _db_url = _prefix + _db_url.split("://", 1)[1]
                _driver_ok = True
                break
            except Exception:
                continue
    if not _driver_ok:
        print("WARNING: no PostgreSQL driver installed - using SQLite (accounts will not persist).")
        _db_url = "sqlite:///users.db"
app.config["SQLALCHEMY_DATABASE_URI"] = _db_url
app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024  # uploads

db = SQLAlchemy(app)


# ============================================================
# PATHS
# ============================================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

import departments as dept_store
from flask import has_request_context


class _DeptPath(os.PathLike):
    """Path inside the *currently selected department's* dataset folder.

    Every place in this app that uses DATASET_DIR / CO_FILE / ... therefore automatically reads the
    data of the department chosen after login (default: the built-in CSE data in Dataset/).
    """

    def __init__(self, filename=None):
        self.filename = filename

    def _resolve(self):
        slug = session.get("department") if has_request_context() else None
        folder = dept_store.dataset_dir_for(BASE_DIR, slug)
        return str(folder if not self.filename else folder / self.filename)

    def __fspath__(self):
        return self._resolve()

    def __str__(self):
        return self._resolve()

    __repr__ = __str__


DATASET_DIR = _DeptPath()
SUBJECT_FILE = _DeptPath("Subject_Details.xlsx")
CO_FILE = _DeptPath("CO_Dataset-1.xlsx")
CO_PO_FILE = _DeptPath("co_po_mapping.csv")
CO_PEO_FILE = _DeptPath("co_peo_mapping.csv")
CO_PO_JUSTIFICATION_FILE = _DeptPath("co_po_justification.csv")
CO_PEO_JUSTIFICATION_FILE = _DeptPath("co_peo_justification.csv")

COLLEGE_PDF_TEMPLATE = os.path.join(
    BASE_DIR,
    "report_templates",
    "College_Template.pdf"
)


# ============================================================
# USER MODEL
# ============================================================

class User(db.Model):

    id = db.Column(
        db.Integer,
        primary_key=True
    )

    username = db.Column(
        db.String(100),
        unique=True,
        nullable=False
    )

    password = db.Column(
        db.String(255),
        nullable=False
    )


with app.app_context():
    db.create_all()


# ============================================================
# BASIC HELPERS
# ============================================================

def safe_string(value):

    if value is None:
        return ""

    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass

    return str(value).strip()


def safe_int(value):

    try:

        if value is None:
            return 0

        if pd.isna(value):
            return 0

        return int(float(value))

    except Exception:

        return 0


def normalize_code(value):

    return safe_string(value).upper().strip()


def clean_column_name(value):

    return safe_string(value).replace("\n", " ").strip()


# ============================================================
# FIND COLUMN
# ============================================================

def find_column(df, possible_columns):

    if df is None or df.empty:
        return None

    normalized = {}

    for column in df.columns:

        key = (
            safe_string(column)
            .lower()
            .replace("_", " ")
            .replace("-", " ")
            .strip()
        )

        normalized[key] = column

    for possible in possible_columns:

        key = (
            safe_string(possible)
            .lower()
            .replace("_", " ")
            .replace("-", " ")
            .strip()
        )

        if key in normalized:

            return normalized[key]

    return None


# ============================================================
# MAPPING STRENGTH
# ============================================================

def get_mapping_strength(value):

    value = safe_int(value)

    if value >= 3:
        return "High"

    if value == 2:
        return "Moderate"

    if value == 1:
        return "Low"

    return "None"


# ============================================================
# MAPPING DESCRIPTION
# ============================================================

def get_mapping_description(value):

    value = safe_int(value)

    if value >= 3:

        return (
            "Strong relationship between the Course Outcome "
            "and the target outcome."
        )

    if value == 2:

        return (
            "Moderate relationship between the Course Outcome "
            "and the target outcome."
        )

    if value == 1:

        return (
            "Low relationship between the Course Outcome "
            "and the target outcome."
        )

    return "No significant relationship identified."


# ============================================================
# BLOOM REASON
# ============================================================

def generate_bloom_reason(
    co_statement,
    bloom_level
):

    statement = safe_string(
        co_statement
    )

    level = safe_string(
        bloom_level
    )

    if not statement:

        return (
            f"The Bloom level is identified as {level} "
            "based on the available outcome information."
        )

    text = statement.lower()

    keyword_map = {

        "K1": [
            "define",
            "identify",
            "list",
            "state",
            "name",
            "recall"
        ],

        "K2": [
            "describe",
            "explain",
            "discuss",
            "summarize",
            "understand"
        ],

        "K3": [
            "apply",
            "implement",
            "use",
            "calculate",
            "demonstrate"
        ],

        "K4": [
            "analyze",
            "analyse",
            "compare",
            "differentiate",
            "examine"
        ],

        "K5": [
            "evaluate",
            "justify",
            "assess",
            "critique"
        ],

        "K6": [
            "design",
            "create",
            "develop",
            "construct",
            "formulate"
        ]
    }

    matched = []

    for key, keywords in keyword_map.items():

        for keyword in keywords:

            if keyword in text:

                matched.append(keyword)

    if matched:

        return (
            f"The Course Outcome is classified as {level} "
            f"because it contains action word(s) such as "
            f"{', '.join(matched[:3])}, which indicate "
            "the corresponding cognitive skill."
        )

    return (
        f"The Course Outcome is classified as {level} "
        "based on NLP-based Bloom's Taxonomy analysis "
        "of the outcome statement."
    )


# ============================================================
# SUBJECT DETAILS
# ============================================================

def get_subject_details(subject_code):

    if not os.path.exists(SUBJECT_FILE):

        raise FileNotFoundError(
            f"Subject file not found:\n{SUBJECT_FILE}"
        )

    df = pd.read_excel(
        SUBJECT_FILE
    )

    df.columns = [
        clean_column_name(column)
        for column in df.columns
    ]

    subject_code_column = find_column(
        df,
        [
            "Subject Code",
            "SubjectCode",
            "Code"
        ]
    )

    subject_name_column = find_column(
        df,
        [
            "Subject Name",
            "SubjectName",
            "Name"
        ]
    )

    if not subject_code_column:

        raise KeyError(
            "Subject Code column not found in Subject_Details.xlsx"
        )

    if not subject_name_column:

        raise KeyError(
            "Subject Name column not found in Subject_Details.xlsx"
        )

    target = normalize_code(
        subject_code
    )

    values = (
        df[subject_code_column]
        .astype(str)
        .str.strip()
        .str.upper()
    )

    match = df[
        values == target
    ]

    if match.empty:

        raise ValueError(
            f"Subject code '{subject_code}' was not found."
        )

    result = match.iloc[0].to_dict()

    # Standard names for the template
    result["Subject Code"] = safe_string(
        result.get(
            subject_code_column,
            subject_code
        )
    )

    result["Subject Name"] = safe_string(
        result.get(
            subject_name_column,
            ""
        )
    )

    # Scheme / Schema: read whatever the uploaded file calls it
    # (Schema, Scheme, Scheme Year, Syllabus Scheme ...) and keep the
    # value from the file instead of a fixed 2022.
    scheme_column = find_column(
        df,
        [
            "Scheme",
            "Schema",
            "Scheme Year",
            "Syllabus Scheme",
            "Syllabus",
            "Regulation",
            "Batch Scheme"
        ]
    )

    scheme_value = ""

    if scheme_column:

        scheme_value = safe_string(
            result.get(
                scheme_column,
                ""
            )
        )

        # 2022.0 (Excel number) -> 2022
        if scheme_value.endswith(".0") and scheme_value[:-2].isdigit():
            scheme_value = scheme_value[:-2]

    result["Scheme"] = scheme_value

    return result


# ============================================================
# LOAD COURSE OUTCOMES
# ============================================================

def load_course_outcomes(subject_code):

    if not os.path.exists(CO_FILE):

        raise FileNotFoundError(
            f"CO file not found:\n{CO_FILE}"
        )

    df = pd.read_excel(
        CO_FILE
    )

    df.columns = [
        clean_column_name(column)
        for column in df.columns
    ]

    if df.empty:

        return pd.DataFrame(
            columns=[
                "Subject Code",
                "CO ID",
                "Course Outcome",
                "Bloom Level",
                "Bloom Reason"
            ]
        )

    # --------------------------------------------------------
    # SUBJECT COLUMN
    # --------------------------------------------------------

    subject_column = find_column(
        df,
        [
            "Subject Code",
            "subject_code",
            "SubjectCode",
            "Code"
        ]
    )

    target = normalize_code(
        subject_code
    )

    if subject_column:

        values = (
            df[subject_column]
            .astype(str)
            .str.strip()
            .str.upper()
        )

        df = df[
            values == target
        ].copy()

    else:

        # Try Unique Subject style columns
        possible = []

        for column in df.columns:

            text = safe_string(
                column
            ).upper()

            if (
                "SUBJECT" in text
                and "CODE" in text
            ):

                possible.append(column)

        if possible:

            column = possible[0]

            values = (
                df[column]
                .astype(str)
                .str.strip()
                .str.upper()
            )

            df = df[
                values == target
            ].copy()

        else:

            return pd.DataFrame(
                columns=[
                    "CO ID",
                    "Course Outcome",
                    "Bloom Level",
                    "Bloom Reason"
                ]
            )

    if df.empty:

        return pd.DataFrame(
            columns=[
                "CO ID",
                "Course Outcome",
                "Bloom Level",
                "Bloom Reason"
            ]
        )

    # --------------------------------------------------------
    # CO ID
    # --------------------------------------------------------

    co_id_column = find_column(
        df,
        [
            "CO ID",
            "CO",
            "CO Code",
            "Course Outcome ID",
            "Course Outcome Code"
        ]
    )

    if co_id_column:

        df["CO ID"] = (
            df[co_id_column]
            .apply(normalize_code)
        )

    else:

        df["CO ID"] = [
            f"CO{i + 1}"
            for i in range(len(df))
        ]

    # --------------------------------------------------------
    # CO STATEMENT
    # --------------------------------------------------------

    statement_column = find_column(
        df,
        [
            "Course Outcome",
            "CO Statement",
            "Course Outcome Statement",
            "Statement",
            "Outcome"
        ]
    )

    if statement_column:

        df["Course Outcome"] = (
            df[statement_column]
            .apply(safe_string)
        )

    else:

        df["Course Outcome"] = ""

    # --------------------------------------------------------
    # REMOVE INVALID ROWS
    # --------------------------------------------------------

    df = df[
        df["CO ID"].astype(str).str.strip() != ""
    ].copy()

    # --------------------------------------------------------
    # BLOOM
    # --------------------------------------------------------

    bloom_levels = []

    for statement in df["Course Outcome"]:

        statement = safe_string(
            statement
        )

        try:

            level = detect_bloom(
                statement
            )

            level = safe_string(
                level
            )

            if not level:

                level = "Unknown"

        except Exception as error:

            print(
                "Bloom detection error:",
                error
            )

            level = "Unknown"

        bloom_levels.append(
            level
        )

    df["Bloom Level"] = bloom_levels

    # --------------------------------------------------------
    # BLOOM REASON
    # --------------------------------------------------------

    df["Bloom Reason"] = [

        generate_bloom_reason(
            statement,
            level
        )

        for statement, level
        in zip(
            df["Course Outcome"],
            df["Bloom Level"]
        )
    ]

    return df.reset_index(
        drop=True
    )


# ============================================================
# EXTRACT CO ID
# ============================================================

def extract_co_id(value):

    value = normalize_code(
        value
    )

    if not value:

        return ""

    # Example:
    # BCS301-CO1 -> CO1
    # BCS301_CO1 -> CO1

    if "-" in value:

        parts = value.split("-")

        last = parts[-1].strip()

        if last.startswith("CO"):

            return last

    if "_" in value:

        parts = value.split("_")

        last = parts[-1].strip()

        if last.startswith("CO"):

            return last

    # Search for CO1, CO2, etc.
    upper = value.upper()

    position = upper.find("CO")

    if position >= 0:

        remaining = upper[position:]

        if remaining.startswith("CO"):

            return remaining

    return value


# ============================================================
# LOAD CO-PO
# ============================================================

def load_co_po(subject_code):

    if not os.path.exists(CO_PO_FILE):

        print(
            "CO-PO file not found:",
            CO_PO_FILE
        )

        return pd.DataFrame()

    df = pd.read_csv(
        CO_PO_FILE
    )

    df.columns = [
        clean_column_name(column)
        for column in df.columns
    ]

    if df.empty:

        return df

    target = normalize_code(
        subject_code
    )

    # --------------------------------------------------------
    # FILTER USING UNIQUE_CO
    # --------------------------------------------------------

    unique_column = find_column(
        df,
        [
            "Unique_CO",
            "Unique CO",
            "UniqueCO"
        ]
    )

    if unique_column:

        values = (
            df[unique_column]
            .astype(str)
            .str.strip()
            .str.upper()
        )

        filtered = df[
            values.str.startswith(
                target + "-"
            )
        ].copy()

        if filtered.empty:

            filtered = df[
                values.str.startswith(
                    target + "_"
                )
            ].copy()

        if filtered.empty:

            filtered = df[
                values.str.startswith(
                    target
                )
            ].copy()

        df = filtered

    else:

        # ----------------------------------------------------
        # FILTER USING CO ID
        # ----------------------------------------------------

        co_column = find_column(
            df,
            [
                "CO ID",
                "CO",
                "CO Code"
            ]
        )

        if co_column:

            values = (
                df[co_column]
                .astype(str)
                .str.strip()
                .str.upper()
            )

            filtered = df[
                values.str.startswith(
                    target + "-"
                )
            ].copy()

            if filtered.empty:

                filtered = df[
                    values.str.startswith(
                        target
                    )
                ].copy()

            df = filtered

        else:

            # ------------------------------------------------
            # FILTER USING SUBJECT CODE
            # ------------------------------------------------

            subject_column = find_column(
                df,
                [
                    "Subject Code",
                    "SubjectCode",
                    "Subject"
                ]
            )

            if subject_column:

                values = (
                    df[subject_column]
                    .astype(str)
                    .str.strip()
                    .str.upper()
                )

                df = df[
                    values == target
                ].copy()

    return df.reset_index(
        drop=True
    )


# ============================================================
# LOAD CO-PEO
# ============================================================

def load_co_peo(subject_code):

    if not os.path.exists(CO_PEO_FILE):

        print(
            "CO-PEO file not found:",
            CO_PEO_FILE
        )

        return pd.DataFrame()

    df = pd.read_csv(
        CO_PEO_FILE
    )

    df.columns = [
        clean_column_name(column)
        for column in df.columns
    ]

    if df.empty:

        return df

    target = normalize_code(
        subject_code
    )

    # --------------------------------------------------------
    # UNIQUE CO
    # --------------------------------------------------------

    unique_column = find_column(
        df,
        [
            "Unique_CO",
            "Unique CO",
            "UniqueCO"
        ]
    )

    if unique_column:

        values = (
            df[unique_column]
            .astype(str)
            .str.strip()
            .str.upper()
        )

        filtered = df[
            values.str.startswith(
                target + "-"
            )
        ].copy()

        if filtered.empty:

            filtered = df[
                values.str.startswith(
                    target + "_"
                )
            ].copy()

        if filtered.empty:

            filtered = df[
                values.str.startswith(
                    target
                )
            ].copy()

        df = filtered

    else:

        co_column = find_column(
            df,
            [
                "CO ID",
                "CO",
                "CO Code"
            ]
        )

        if co_column:

            values = (
                df[co_column]
                .astype(str)
                .str.strip()
                .str.upper()
            )

            filtered = df[
                values.str.startswith(
                    target + "-"
                )
            ].copy()

            if filtered.empty:

                filtered = df[
                    values.str.startswith(
                        target
                    )
                ].copy()

            df = filtered

        else:

            subject_column = find_column(
                df,
                [
                    "Subject Code",
                    "SubjectCode",
                    "Subject"
                ]
            )

            if subject_column:

                values = (
                    df[subject_column]
                    .astype(str)
                    .str.strip()
                    .str.upper()
                )

                df = df[
                    values == target
                ].copy()

    return df.reset_index(
        drop=True
    )


# ============================================================
# FIND MAPPING COLUMNS
# ============================================================

def get_mapping_columns(df):

    if df is None or df.empty:

        return []

    columns = []

    for column in df.columns:

        text = (
            safe_string(column)
            .upper()
            .replace(" ", "")
            .replace("_", "")
            .replace("-", "")
        )

        # PO1, PO2, PO12
        if text.startswith("PO"):

            remaining = text[2:]

            if remaining.isdigit():

                columns.append(
                    column
                )

                continue

        # PEO1, PEO2, PEO4
        if text.startswith("PEO"):

            remaining = text[3:]

            if remaining.isdigit():

                columns.append(
                    column
                )

    # --------------------------------------------------------
    # If normal detection did not work,
    # check names containing PO / PEO.
    # --------------------------------------------------------

    if not columns:

        for column in df.columns:

            text = (
                safe_string(column)
                .upper()
                .replace(" ", "")
                .replace("_", "")
                .replace("-", "")
            )

            if text.startswith("PO"):

                columns.append(
                    column
                )

            elif text.startswith("PEO"):

                columns.append(
                    column
                )

    return columns


# ============================================================
# SORT PO / PEO COLUMNS
# ============================================================

def sort_mapping_columns(columns):

    def number(value):

        text = (
            safe_string(value)
            .upper()
            .replace(" ", "")
            .replace("_", "")
            .replace("-", "")
        )

        digits = ""

        for char in text:

            if char.isdigit():

                digits += char

        if digits:

            return int(digits)

        return 999

    return sorted(
        columns,
        key=number
    )


# ============================================================
# GET CO STATEMENT
# ============================================================

def get_co_statement(
    co_df,
    co_id
):

    if co_df is None or co_df.empty:

        return ""

    target = normalize_code(
        extract_co_id(co_id)
    )

    for _, row in co_df.iterrows():

        current = normalize_code(
            row.get(
                "CO ID",
                ""
            )
        )

        current_short = normalize_code(
            extract_co_id(current)
        )

        if (
            current == target
            or
            current_short == target
        ):

            return safe_string(
                row.get(
                    "Course Outcome",
                    ""
                )
            )

    return ""


# ============================================================
# AVERAGE DATA
# ============================================================

def numeric_average_data(df):

    if df is None or df.empty:

        return [], []

    columns = sort_mapping_columns(
        get_mapping_columns(df)
    )

    labels = []
    values = []

    for column in columns:

        numeric = pd.to_numeric(
            df[column],
            errors="coerce"
        )

        numeric = numeric.dropna()

        if numeric.empty:

            average = 0.0

        else:

            average = float(
                numeric.mean()
            )

        labels.append(
            safe_string(column)
        )

        values.append(
            round(
                average,
                2
            )
        )

    return labels, values


# ============================================================
# BUILD CO ROWS
# ============================================================

def build_co_rows(co_df):

    rows = []

    if co_df is None or co_df.empty:

        return rows

    for _, row in co_df.iterrows():

        co_id = safe_string(
            row.get(
                "CO ID",
                ""
            )
        )

        statement = safe_string(
            row.get(
                "Course Outcome",
                ""
            )
        )

        bloom = safe_string(
            row.get(
                "Bloom Level",
                "Unknown"
            )
        )

        if not bloom:

            bloom = "Unknown"

        reason = safe_string(
            row.get(
                "Bloom Reason",
                ""
            )
        )

        if not reason:

            reason = generate_bloom_reason(
                statement,
                bloom
            )

        rows.append({

            "CO ID":
                co_id,

            "CO Statement":
                statement,

            "Course Outcome":
                statement,

            "Bloom Level":
                bloom,

            "Bloom Reason":
                reason
        })

    return rows


# ============================================================
# GET MAPPING CO ID
# ============================================================

def get_mapping_co_id(
    mapping_row
):

    value = ""

    for column in [
        "Unique_CO",
        "Unique CO",
        "UniqueCO",
        "CO ID",
        "CO",
        "CO Code"
    ]:

        if column in mapping_row.index:

            value = safe_string(
                mapping_row.get(
                    column,
                    ""
                )
            )

            if value:

                break

    return extract_co_id(
        value
    )


# ============================================================
# BUILD CO-PO MATRIX
# ============================================================

def build_co_po_matrix(
    co_df,
    co_po_df
):

    matrix = []

    if (
        co_po_df is None
        or co_po_df.empty
    ):

        return matrix

    columns = sort_mapping_columns(
        get_mapping_columns(
            co_po_df
        )
    )

    if not columns:

        return matrix

    # IMPORTANT:
    # No header row is added here.
    # subject_report.html expects:
    # row[0] = CO ID
    # row[1:] = PO values

    for _, mapping_row in co_po_df.iterrows():

        co_id = get_mapping_co_id(
            mapping_row
        )

        if not co_id:

            continue

        row_data = [
            co_id
        ]

        for column in columns:

            value = safe_int(
                mapping_row.get(
                    column,
                    0
                )
            )

            row_data.append(
                value
            )

        matrix.append(
            row_data
        )

    return matrix


# ============================================================
# BUILD CO-PEO MATRIX
# ============================================================

def build_co_peo_matrix(
    co_df,
    co_peo_df
):

    matrix = []

    if (
        co_peo_df is None
        or co_peo_df.empty
    ):

        return matrix

    columns = sort_mapping_columns(
        get_mapping_columns(
            co_peo_df
        )
    )

    if not columns:

        return matrix

    for _, mapping_row in co_peo_df.iterrows():

        co_id = get_mapping_co_id(
            mapping_row
        )

        if not co_id:

            continue

        row_data = [
            co_id
        ]

        for column in columns:

            value = safe_int(
                mapping_row.get(
                    column,
                    0
                )
            )

            row_data.append(
                value
            )

        matrix.append(
            row_data
        )

    return matrix


# ============================================================
# LOAD JUSTIFICATION
# ============================================================

def load_justification(
    file_path,
    subject_code
):

    if not os.path.exists(
        file_path
    ):

        return pd.DataFrame()

    try:

        df = pd.read_csv(
            file_path
        )

    except Exception as error:

        print(
            "Justification read error:",
            error
        )

        return pd.DataFrame()

    if df.empty:

        return df

    df.columns = [
        clean_column_name(column)
        for column in df.columns
    ]

    target = normalize_code(
        subject_code
    )

    # --------------------------------------------------------
    # CO ID FILTER
    # --------------------------------------------------------

    co_column = find_column(
        df,
        [
            "CO ID",
            "CO",
            "Unique_CO",
            "Unique CO"
        ]
    )

    if co_column:

        values = (
            df[co_column]
            .astype(str)
            .str.strip()
            .str.upper()
        )

        filtered = df[
            values.str.startswith(
                target + "-"
            )
        ].copy()

        if filtered.empty:

            filtered = df[
                values.str.startswith(
                    target
                )
            ].copy()

        # If subject is not encoded in CO,
        # check subject column.
        if filtered.empty:

            subject_column = find_column(
                df,
                [
                    "Subject Code",
                    "SubjectCode",
                    "Subject"
                ]
            )

            if subject_column:

                subject_values = (
                    df[subject_column]
                    .astype(str)
                    .str.strip()
                    .str.upper()
                )

                filtered = df[
                    subject_values == target
                ].copy()

        df = filtered

    return df.reset_index(
        drop=True
    )


# ============================================================
# FIND EXISTING JUSTIFICATION
# ============================================================

def is_generic_justification(text):
    """Return True for old placeholder justifications that should be regenerated."""
    text = safe_string(text).strip().lower()
    generic_phrases = (
        "high semantic similarity between",
        "medium semantic similarity between",
        "low semantic similarity between",
        "no significant relationship",
    )
    return any(text.startswith(p) for p in generic_phrases)


def find_existing_justification(
    justification_df,
    co_id,
    target_id,
    target_type
):

    if (
        justification_df is None
        or justification_df.empty
    ):

        return ""

    co_column = find_column(
        justification_df,
        [
            "CO ID",
            "CO",
            "Unique_CO"
        ]
    )

    if not co_column:

        return ""

    if target_type == "PO":

        target_column = find_column(
            justification_df,
            [
                "PO ID",
                "PO",
                "PO Code"
            ]
        )

    else:

        target_column = find_column(
            justification_df,
            [
                "PEO ID",
                "PEO",
                "PEO Code"
            ]
        )

    if not target_column:

        return ""

    justification_column = find_column(
        justification_df,
        [
            "Justification",
            "Reason",
            "Explanation"
        ]
    )

    if not justification_column:

        return ""

    target_co = normalize_code(
        extract_co_id(co_id)
    )

    target_target = normalize_code(
        target_id
    )

    for _, row in justification_df.iterrows():

        row_co = normalize_code(
            extract_co_id(
                row.get(
                    co_column,
                    ""
                )
            )
        )

        row_target = normalize_code(
            row.get(
                target_column,
                ""
            )
        )

        if (
            row_co == target_co
            and
            row_target == target_target
        ):

            result = safe_string(
                row.get(
                    justification_column,
                    ""
                )
            )

            if result:

                return result

    return ""


# ============================================================
# DETAILED JUSTIFICATION
# ============================================================

def generate_detailed_justification(
    co_statement,
    target_id,
    target_type,
    mapping_value
):

    co_statement = safe_string(
        co_statement
    )

    target_id = safe_string(
        target_id
    )

    target_type = safe_string(
        target_type
    )

    mapping_value = safe_int(
        mapping_value
    )

    strength = get_mapping_strength(
        mapping_value
    )

    description = get_mapping_description(
        mapping_value
    )

    if not co_statement:

        co_statement = (
            "the given Course Outcome"
        )

    if mapping_value == 0:

        return (
            f"The Course Outcome '{co_statement}' "
            f"has no significant mapping with "
            f"{target_type} {target_id}. "
            f"The NLP-based semantic similarity analysis "
            f"resulted in a mapping value of 0/3, indicating "
            f"no meaningful relationship."
        )

    return (
        f"The Course Outcome '{co_statement}' "
        f"shows a {strength.lower()} relationship with "
        f"{target_type} {target_id}. "
        f"The mapping value is {mapping_value}/3. "
        f"{description} "
        f"The relationship was identified through "
        f"NLP-based semantic similarity analysis between "
        f"the Course Outcome and the corresponding "
        f"{target_type} description."
    )


# ============================================================
# BUILD CO-PO REASON ROWS
# ============================================================

def build_po_reason_rows(
    co_df,
    co_po_df,
    po_justification
):

    rows = []

    if (
        co_po_df is None
        or co_po_df.empty
    ):

        return rows

    columns = sort_mapping_columns(
        get_mapping_columns(
            co_po_df
        )
    )

    for _, mapping_row in co_po_df.iterrows():

        co_id = get_mapping_co_id(
            mapping_row
        )

        if not co_id:

            continue

        co_statement = get_co_statement(
            co_df,
            co_id
        )

        for po_id in columns:

            mapping_value = safe_int(
                mapping_row.get(
                    po_id,
                    0
                )
            )

            existing = (
                find_existing_justification(
                    po_justification,
                    co_id,
                    po_id,
                    "PO"
                )
            )

            if existing and not is_generic_justification(existing):

                reason = existing

            else:

                reason = (
                    generate_detailed_justification(
                        co_statement,
                        po_id,
                        "PO",
                        mapping_value
                    )
                )

            rows.append({

                "CO ID":
                    co_id,

                "PO ID":
                    safe_string(
                        po_id
                    ),

                "Mapping Value":
                    mapping_value,

                "Strength":
                    get_mapping_strength(
                        mapping_value
                    ),

                "CO Statement":
                    co_statement,

                "Reason":
                    reason
            })

    return rows


# ============================================================
# BUILD CO-PEO REASON ROWS
# ============================================================

def build_peo_reason_rows(
    co_df,
    co_peo_df,
    peo_justification
):

    rows = []

    if (
        co_peo_df is None
        or co_peo_df.empty
    ):

        return rows

    columns = sort_mapping_columns(
        get_mapping_columns(
            co_peo_df
        )
    )

    for _, mapping_row in co_peo_df.iterrows():

        co_id = get_mapping_co_id(
            mapping_row
        )

        if not co_id:

            continue

        co_statement = get_co_statement(
            co_df,
            co_id
        )

        for peo_id in columns:

            mapping_value = safe_int(
                mapping_row.get(
                    peo_id,
                    0
                )
            )

            existing = (
                find_existing_justification(
                    peo_justification,
                    co_id,
                    peo_id,
                    "PEO"
                )
            )

            if existing and not is_generic_justification(existing):

                reason = existing

            else:

                reason = (
                    generate_detailed_justification(
                        co_statement,
                        peo_id,
                        "PEO",
                        mapping_value
                    )
                )

            rows.append({

                "CO ID":
                    co_id,

                "PEO ID":
                    safe_string(
                        peo_id
                    ),

                "Mapping Value":
                    mapping_value,

                "Strength":
                    get_mapping_strength(
                        mapping_value
                    ),

                "CO Statement":
                    co_statement,

                "Reason":
                    reason
            })

    return rows


# ============================================================
# BUILD AI REASON ROWS
# ============================================================

def build_ai_reason_rows(
    po_reason_rows,
    peo_reason_rows
):

    rows = []

    for item in po_reason_rows:

        rows.append({

            "CO ID":
                safe_string(
                    item.get(
                        "CO ID",
                        ""
                    )
                ),

            "Target ID":
                safe_string(
                    item.get(
                        "PO ID",
                        ""
                    )
                ),

            "Target Type":
                "PO",

            "Mapping Value":
                safe_int(
                    item.get(
                        "Mapping Value",
                        0
                    )
                ),

            "Strength":
                safe_string(
                    item.get(
                        "Strength",
                        ""
                    )
                ),

            "CO Statement":
                safe_string(
                    item.get(
                        "CO Statement",
                        ""
                    )
                ),

            "Justification":
                safe_string(
                    item.get(
                        "Reason",
                        ""
                    )
                )
        })

    for item in peo_reason_rows:

        rows.append({

            "CO ID":
                safe_string(
                    item.get(
                        "CO ID",
                        ""
                    )
                ),

            "Target ID":
                safe_string(
                    item.get(
                        "PEO ID",
                        ""
                    )
                ),

            "Target Type":
                "PEO",

            "Mapping Value":
                safe_int(
                    item.get(
                        "Mapping Value",
                        0
                    )
                ),

            "Strength":
                safe_string(
                    item.get(
                        "Strength",
                        ""
                    )
                ),

            "CO Statement":
                safe_string(
                    item.get(
                        "CO Statement",
                        ""
                    )
                ),

            "Justification":
                safe_string(
                    item.get(
                        "Reason",
                        ""
                    )
                )
        })

    return rows


# ============================================================
# BUILD BLOOM REASON ROWS
# ============================================================

def build_bloom_reason_rows(
    co_df
):

    rows = []

    if co_df is None or co_df.empty:

        return rows

    for _, row in co_df.iterrows():

        co_id = safe_string(
            row.get(
                "CO ID",
                ""
            )
        )

        statement = safe_string(
            row.get(
                "Course Outcome",
                ""
            )
        )

        bloom = safe_string(
            row.get(
                "Bloom Level",
                "Unknown"
            )
        )

        reason = safe_string(
            row.get(
                "Bloom Reason",
                ""
            )
        )

        if not reason:

            reason = generate_bloom_reason(
                statement,
                bloom
            )

        rows.append({

            "CO ID":
                co_id,

            "Bloom Level":
                bloom,

            "Explanation":
                reason
        })

    return rows


# ============================================================
# CREATE HTML TABLE
# ============================================================

def dataframe_to_html(
    df
):

    if df is None or df.empty:

        return ""

    try:

        return df.to_html(
            index=False,
            classes="data-table",
            escape=True
        )

    except Exception:

        return ""


# ============================================================
# SUBJECT REPORT
# ============================================================


def add_word_section_heading(document, text, *args, **kwargs):
    """Safe Word heading that does not require Heading 1 style."""
    paragraph = document.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

    run = paragraph.add_run(str(text))
    run.bold = True
    run.font.size = Pt(14)

    return paragraph

@app.route(
    "/subject_report",
    methods=["GET", "POST"]
)
def subject_report():

    if "user" not in session:

        return redirect("/")

    # ========================================================
    # GET SUBJECT CODE
    # ========================================================

    subject_code = safe_string(
        request.values.get(
            "subject",
            ""
        )
    )

    if not subject_code:

        subject_code = safe_string(
            request.values.get(
                "subject_code",
                ""
            )
        )

    if not subject_code:

        subject_code = safe_string(
            session.get(
                "subject_code",
                ""
            )
        )

    if not subject_code:

        return redirect(
            "/dashboard"
        )

    subject_code = normalize_code(
        subject_code
    )

    print("\n")
    print("=" * 70)
    print("LOADING SUBJECT REPORT")
    print("SUBJECT:", subject_code)
    print("=" * 70)

    # ========================================================
    # SUBJECT DETAILS
    # ========================================================

    try:

        details = get_subject_details(
            subject_code
        )

    except Exception as error:

        print(
            "SUBJECT DETAILS ERROR:",
            error
        )

        return f"""
        <html>
        <body style="
            font-family:Arial;
            text-align:center;
            padding:100px;
        ">

        <h2>Unable to load subject</h2>

        <p>{safe_string(error)}</p>

        <a href="/dashboard">
            Back to Dashboard
        </a>

        </body>
        </html>
        """

    session["subject_code"] = subject_code

    # ========================================================
    # SUBJECT INFORMATION
    # ========================================================

    subject_name = safe_string(
        details.get(
            "Subject Name",
            ""
        )
    )

    semester = safe_string(
        details.get(
            "Semester",
            ""
        )
    )

    credits = safe_string(
        details.get(
            "Credits",
            ""
        )
    )

    department = safe_string(
        details.get(
            "Department",
            "Computer Science and Engineering"
        )
    )

    scheme = safe_string(
        details.get(
            "Scheme",
            ""
        )
    )

    # Remove VTU if already present
    if scheme.upper().startswith("VTU"):

        scheme = scheme[3:].strip()

    if not scheme:

        scheme = "N/A"

    scheme_display = scheme

    year = safe_string(details.get("Year", ""))
    if year.endswith(".0") and year[:-2].isdigit():
        year = year[:-2]

    # ========================================================
    # LOAD CO
    # ========================================================

    try:

        co = load_course_outcomes(
            subject_code
        )

    except Exception as error:

        print(
            "CO ERROR:",
            error
        )

        co = pd.DataFrame()

    # ========================================================
    # LOAD CO-PO
    # ========================================================

    try:

        co_po = load_co_po(
            subject_code
        )

    except Exception as error:

        print(
            "CO-PO ERROR:",
            error
        )

        co_po = pd.DataFrame()

    # ========================================================
    # LOAD CO-PEO
    # ========================================================

    try:

        co_peo = load_co_peo(
            subject_code
        )

    except Exception as error:

        print(
            "CO-PEO ERROR:",
            error
        )

        co_peo = pd.DataFrame()

    # ========================================================
    # JUSTIFICATION
    # ========================================================

    try:

        po_justification = (
            load_justification(
                CO_PO_JUSTIFICATION_FILE,
                subject_code
            )
        )

    except Exception as error:

        print(
            "PO JUSTIFICATION ERROR:",
            error
        )

        po_justification = pd.DataFrame()

    try:

        peo_justification = (
            load_justification(
                CO_PEO_JUSTIFICATION_FILE,
                subject_code
            )
        )

    except Exception as error:

        print(
            "PEO JUSTIFICATION ERROR:",
            error
        )

        peo_justification = pd.DataFrame()

    # ========================================================
    # BUILD CO ROWS
    # ========================================================

    co_rows = build_co_rows(
        co
    )

    co_count = len(
        co_rows
    )

    # ========================================================
    # BUILD CO-PO MATRIX
    # ========================================================

    co_po_matrix = (
        build_co_po_matrix(
            co,
            co_po
        )
    )

    # ========================================================
    # BUILD CO-PEO MATRIX
    # ========================================================

    co_peo_matrix = (
        build_co_peo_matrix(
            co,
            co_peo
        )
    )

    # ========================================================
    # LOAD MANUALLY MAINTAINED PSO DATA + CO-PSO MAPPING
    # ========================================================
    co_pso_matrix = []
    pso_labels = []
    pso_records = []
    pso_reason_rows = []
    pso_graph_labels = []
    pso_graph_values = []
    co_pso_subject = pd.DataFrame()
    try:
        from pso_mapping import generate_pso_mapping
        pso_result = generate_pso_mapping(BASE_DIR, DATASET_DIR)
        pso_all = pso_result["mapping"].fillna("")
        pso_file = os.path.join(DATASET_DIR, "PSO_Dataset.xlsx")
        pso_df = pd.read_excel(pso_file).fillna("")
        pso_df.columns = [clean_column_name(c) for c in pso_df.columns]
        pso_id_col = find_column(pso_df, ["PSO ID", "PSO"])
        pso_desc_col = find_column(pso_df, ["Programme Specific Outcomes", "Program Specific Outcomes", "PSO Description", "Description"])
        if pso_id_col and pso_desc_col:
            pso_records = pso_df[[pso_id_col, pso_desc_col]].to_dict(orient="records")
            pso_labels = [str(v).strip() for v in pso_df[pso_id_col].tolist() if safe_string(v)]
        if "Subject Code" in pso_all.columns:
            co_pso_subject = pso_all[pso_all["Subject Code"].astype(str).str.upper().str.strip() == subject_code].copy()
        for _, map_row in co_pso_subject.iterrows():
            row_values = [safe_string(map_row.get("CO ID", ""))]
            row_values.extend(safe_int(map_row.get(pid, 0)) for pid in pso_labels)
            if row_values[0]:
                co_pso_matrix.append(row_values)
        for pid in pso_labels:
            values = [safe_int(v) for v in co_pso_subject[pid].tolist()] if pid in co_pso_subject.columns else []
            pso_graph_labels.append(pid)
            pso_graph_values.append(round(sum(values) / len(values), 2) if values else 0)
        just_all = pso_result["justifications"].fillna("")
        if "Subject Code" in just_all.columns:
            just_all = just_all[just_all["Subject Code"].astype(str).str.upper().str.strip() == subject_code]
        co_statement_lookup = {}
        if not co.empty:
            coid_col = find_column(co, ["CO ID", "Course Outcome ID", "Outcome ID"])
            cotext_col = find_column(co, ["Course Outcome", "CO Statement", "Outcome Description", "Description"])
            if coid_col and cotext_col:
                co_statement_lookup = {str(r[coid_col]).strip(): str(r[cotext_col]).strip() for _, r in co.iterrows()}
        for _, jr in just_all.iterrows():
            value = safe_int(jr.get("Mapping Level", 0))
            pso_reason_rows.append({
                "CO ID": safe_string(jr.get("CO ID", "")),
                "Target ID": safe_string(jr.get("PSO ID", "")),
                "Target Type": "PSO",
                "Mapping Value": value,
                "Strength": safe_string(jr.get("Strength", get_mapping_strength(value))),
                "CO Statement": co_statement_lookup.get(safe_string(jr.get("CO ID", "")), ""),
                "Justification": safe_string(jr.get("Justification", ""))
            })
    except Exception as error:
        print("CO-PSO DATA ERROR:", error)

    # ========================================================
    # CO-PO CHART
    # ========================================================

    co_po_labels = []
    co_po_values = []

    if (
        co_po is not None
        and not co_po.empty
    ):

        try:

            (
                co_po_labels,
                co_po_values
            ) = numeric_average_data(
                co_po
            )

        except Exception as error:

            print(
                "CO-PO CHART ERROR:",
                error
            )

    # ========================================================
    # CO-PEO CHART
    # ========================================================

    co_peo_labels = []
    co_peo_values = []

    if (
        co_peo is not None
        and not co_peo.empty
    ):

        try:

            (
                co_peo_labels,
                co_peo_values
            ) = numeric_average_data(
                co_peo
            )

        except Exception as error:

            print(
                "CO-PEO CHART ERROR:",
                error
            )

    # ========================================================
    # BLOOM CHART
    # ========================================================

    bloom_labels = []
    bloom_values = []

    if co_rows:

        counts = {}

        for row in co_rows:

            level = safe_string(
                row.get(
                    "Bloom Level",
                    "Unknown"
                )
            )

            if not level:

                level = "Unknown"

            counts[level] = (
                counts.get(level, 0) + 1
            )

        # K1-K6 order
        preferred_order = [
            "K1",
            "K2",
            "K3",
            "K4",
            "K5",
            "K6",
            "Unknown"
        ]

        for level in preferred_order:

            if level in counts:

                bloom_labels.append(
                    level
                )

                bloom_values.append(
                    counts[level]
                )

        # Any unexpected levels
        for level, value in counts.items():

            if level not in bloom_labels:

                bloom_labels.append(
                    level
                )

                bloom_values.append(
                    value
                )

    # ========================================================
    # BLOOM REASONS
    # ========================================================

    bloom_reason_rows = (
        build_bloom_reason_rows(
            co
        )
    )

    # ========================================================
    # CO-PO REASONS
    # ========================================================

    po_reason_rows = (
        build_po_reason_rows(
            co,
            co_po,
            po_justification
        )
    )

    # ========================================================
    # CO-PEO REASONS
    # ========================================================

    peo_reason_rows = (
        build_peo_reason_rows(
            co,
            co_peo,
            peo_justification
        )
    )

    # ========================================================
    # AI JUSTIFICATION
    # ========================================================

    ai_reason_rows = (
        build_ai_reason_rows(
            po_reason_rows,
            peo_reason_rows
        )
    )
    # Include PSO justifications in the same searchable report panel.
    ai_reason_rows.extend(pso_reason_rows)

    ai_justifications = ai_reason_rows

    # ========================================================
    # COUNTS
    # ========================================================

    po_count = len(
        co_po_labels
    )

    if po_count == 0:

        po_count = 11

    peo_count = len(
        co_peo_labels
    )

    if peo_count == 0:

        peo_count = 4

    # ========================================================
    # OLD TABLE VARIABLES
    # ========================================================

    co_table = dataframe_to_html(
        co
    )

    co_po_table = dataframe_to_html(
        co_po
    )

    co_peo_table = dataframe_to_html(
        co_peo
    )

    if po_reason_rows:

        justification_table = (
            pd.DataFrame(
                po_reason_rows
            ).to_html(
                index=False,
                classes="data-table",
                escape=True
            )
        )

    else:

        justification_table = ""

    if peo_reason_rows:

        co_peo_justification_table = (
            pd.DataFrame(
                peo_reason_rows
            ).to_html(
                index=False,
                classes="data-table",
                escape=True
            )
        )

    else:

        co_peo_justification_table = ""

    if ai_reason_rows:

        ai_justification_table = (
            pd.DataFrame(
                ai_reason_rows
            ).to_html(
                index=False,
                classes="data-table",
                escape=True
            )
        )

    else:

        ai_justification_table = ""

    # ========================================================
    # DEBUG
    # ========================================================

    print()
    print("SUBJECT CODE:", subject_code)
    print("SUBJECT NAME:", subject_name)
    print("CO COUNT:", co_count)
    print("CO ROWS:", len(co_rows))
    print("CO-PO DATASET ROWS:", len(co_po))
    print("CO-PEO DATASET ROWS:", len(co_peo))
    print("CO-PO MATRIX ROWS:", len(co_po_matrix))
    print("CO-PEO MATRIX ROWS:", len(co_peo_matrix))
    print("CO-PO LABELS:", co_po_labels)
    print("CO-PO VALUES:", co_po_values)
    print("CO-PEO LABELS:", co_peo_labels)
    print("CO-PEO VALUES:", co_peo_values)
    print("BLOOM LABELS:", bloom_labels)
    print("BLOOM VALUES:", bloom_values)
    print("PO REASON ROWS:", len(po_reason_rows))
    print("PEO REASON ROWS:", len(peo_reason_rows))
    print("AI ROWS:", len(ai_reason_rows))
    print("=" * 70)
    print()

    # ========================================================
    # RENDER
    # ========================================================

    return render_template(

        "subject_report.html",

        # ----------------------------------------------------
        # SUBJECT
        # ----------------------------------------------------

        subject_code=subject_code,

        subject=subject_code,

        subject_name=subject_name,

        semester=semester,

        credits=credits,

        year=year,

        scheme=scheme_display,

        department=department,

        # ----------------------------------------------------
        # COUNTS
        # ----------------------------------------------------

        co_count=co_count,

        po_count=po_count,

        peo_count=peo_count,

        # ----------------------------------------------------
        # CO
        # ----------------------------------------------------

        co_rows=co_rows,

        # ----------------------------------------------------
        # MATRICES
        # ----------------------------------------------------

        co_po_matrix=co_po_matrix,

        co_peo_matrix=co_peo_matrix,
        co_pso_matrix=co_pso_matrix,
        pso_labels=pso_labels,
        pso_records=pso_records,
        pso_graph_labels=pso_graph_labels,
        pso_graph_values=pso_graph_values,

        # ----------------------------------------------------
        # BLOOM
        # ----------------------------------------------------

        bloom_labels=bloom_labels,

        bloom_values=bloom_values,

        bloom_reason_rows=bloom_reason_rows,

        # ----------------------------------------------------
        # CO-PO CHART
        # ----------------------------------------------------

        co_po_labels=co_po_labels,

        co_po_values=co_po_values,

        graph_labels=co_po_labels,

        graph_values=co_po_values,

        # ----------------------------------------------------
        # CO-PEO CHART
        # ----------------------------------------------------

        co_peo_labels=co_peo_labels,

        co_peo_values=co_peo_values,

        peo_graph_labels=co_peo_labels,

        peo_graph_values=co_peo_values,

        # ----------------------------------------------------
        # JUSTIFICATION
        # ----------------------------------------------------

        po_reason_rows=po_reason_rows,

        peo_reason_rows=peo_reason_rows,

        ai_reason_rows=ai_reason_rows,

        ai_justifications=ai_justifications,

        # ----------------------------------------------------
        # OLD VARIABLES
        # ----------------------------------------------------

        justification_table=justification_table,

        co_peo_justification_table=(
            co_peo_justification_table
        ),

        ai_justification_table=(
            ai_justification_table
        ),

        co_table=co_table,

        co_po_table=co_po_table,

        co_peo_table=co_peo_table
    )


# ============================================================
# LOGIN
# ============================================================


def add_word_section_heading(document, text, *args, **kwargs):
    """Safe Word heading that does not require Heading 1 style."""
    paragraph = document.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

    run = paragraph.add_run(str(text))
    run.bold = True
    run.font.size = Pt(14)

    return paragraph

@app.route(
    "/",
    methods=["GET", "POST"]
)
def login():

    if request.method == "POST":

        username = safe_string(
            request.form.get(
                "username",
                ""
            )
        )

        password = request.form.get(
            "password",
            ""
        )

        user = User.query.filter_by(
            username=username
        ).first()

        if (
            user
            and check_password_hash(
                user.password,
                password
            )
        ):

            session["user"] = username
            session.pop("department", None)

            return redirect(
                "/departments"
            )

        return render_template(
            "login.html",
            error="Invalid username or password."
        )

    return render_template(
        "login.html"
    )


# ============================================================
# REGISTER
# ============================================================


def add_word_section_heading(document, text, *args, **kwargs):
    """Safe Word heading that does not require Heading 1 style."""
    paragraph = document.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

    run = paragraph.add_run(str(text))
    run.bold = True
    run.font.size = Pt(14)

    return paragraph

@app.route("/healthz")
def healthz():
    return "ok", 200


@app.route(
    "/register",
    methods=["GET", "POST"]
)
def register():

    if request.method == "POST":

        username = safe_string(
            request.form.get(
                "username",
                ""
            )
        )

        password = request.form.get(
            "password",
            ""
        )

        if not username or not password:

            return render_template(
                "register.html",
                error=(
                    "Please enter username "
                    "and password."
                )
            )

        existing = User.query.filter_by(
            username=username
        ).first()

        if existing:

            return render_template(
                "register.html",
                error=(
                    "Username already exists."
                )
            )

        user = User(
            username=username,
            password=generate_password_hash(
                password
            )
        )

        db.session.add(
            user
        )

        db.session.commit()

        return redirect("/")

    # Always serve a blank, non-cached form so a new username and
    # password are typed every time (no autofill / no back-button refill).
    response = make_response(
        render_template(
            "register.html"
        )
    )
    response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    response.headers["Pragma"] = "no-cache"
    return response


# ============================================================
# DASHBOARD
# ============================================================


def add_word_section_heading(document, text, *args, **kwargs):
    """Safe Word heading that does not require Heading 1 style."""
    paragraph = document.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

    run = paragraph.add_run(str(text))
    run.bold = True
    run.font.size = Pt(14)

    return paragraph

@app.route(
    "/dashboard"
)
def dashboard():

    if "user" not in session:

        return redirect("/")

    if "department" not in session:

        return redirect("/departments")

    subjects = []

    graph_labels = []
    graph_values = []

    peo_graph_labels = []
    peo_graph_values = []

    bloom_labels = []
    bloom_values = []

    co_po_labels = []
    co_po_values = []

    co_peo_labels = []
    co_peo_values = []

    # PSO dashboard data: only the PSO list and CO–PSO strength chart are shown.
    pso_records = []
    pso_graph_labels = []
    pso_graph_values = []

    subject_count = 0
    co_count = 0

    po_count = 11
    peo_count = 4

    total_po_mapping = 0
    total_peo_mapping = 0

    ai_accuracy = "SBERT"

    # ========================================================
    # SUBJECTS
    # ========================================================

    try:

        if os.path.exists(
            SUBJECT_FILE
        ):

            subject_df = pd.read_excel(
                SUBJECT_FILE
            )

            subject_df.columns = [
                clean_column_name(column)
                for column in subject_df.columns
            ]

            code_column = find_column(
                subject_df,
                [
                    "Subject Code",
                    "SubjectCode",
                    "Code"
                ]
            )

            name_column = find_column(
                subject_df,
                [
                    "Subject Name",
                    "SubjectName",
                    "Name"
                ]
            )

            if (
                code_column
                and
                name_column
            ):

                subject_df = (
                    subject_df[
                        [
                            code_column,
                            name_column
                        ]
                    ]
                    .fillna("")
                )

                subjects = []

                for _, row in subject_df.iterrows():

                    code = safe_string(
                        row.get(
                            code_column,
                            ""
                        )
                    )

                    name = safe_string(
                        row.get(
                            name_column,
                            ""
                        )
                    )

                    if code:

                        subjects.append(
                            [
                                code,
                                name
                            ]
                        )

                subject_count = len(
                    subjects
                )

    except Exception as error:

        print(
            "Dashboard subject error:",
            error
        )

    # ========================================================
    # CO COUNT
    # ========================================================

    try:

        if os.path.exists(
            CO_FILE
        ):

            all_co = pd.read_excel(
                CO_FILE
            )

            co_count = len(
                all_co
            )

    except Exception as error:

        print(
            "Dashboard CO error:",
            error
        )

    # ========================================================
    # CO-PO
    # ========================================================

    try:

        if os.path.exists(
            CO_PO_FILE
        ):

            df = pd.read_csv(
                CO_PO_FILE
            )

            df.columns = [
                clean_column_name(column)
                for column in df.columns
            ]

            (
                co_po_labels,
                co_po_values
            ) = numeric_average_data(
                df
            )

            graph_labels = (
                co_po_labels
            )

            graph_values = (
                co_po_values
            )

            for column in get_mapping_columns(
                df
            ):

                numeric = pd.to_numeric(
                    df[column],
                    errors="coerce"
                )

                total_po_mapping += int(
                    (numeric > 0).sum()
                )

    except Exception as error:

        print(
            "Dashboard CO-PO error:",
            error
        )

    # ========================================================
    # CO-PEO
    # ========================================================

    try:

        if os.path.exists(
            CO_PEO_FILE
        ):

            df = pd.read_csv(
                CO_PEO_FILE
            )

            df.columns = [
                clean_column_name(column)
                for column in df.columns
            ]

            (
                co_peo_labels,
                co_peo_values
            ) = numeric_average_data(
                df
            )

            peo_graph_labels = (
                co_peo_labels
            )

            peo_graph_values = (
                co_peo_values
            )

            for column in get_mapping_columns(
                df
            ):

                numeric = pd.to_numeric(
                    df[column],
                    errors="coerce"
                )

                total_peo_mapping += int(
                    (numeric > 0).sum()
                )

    except Exception as error:

        print(
            "Dashboard CO-PEO error:",
            error
        )

    # ========================================================
    # PSO LIST + CO–PSO MAPPING STRENGTH CHART
    # ========================================================

    try:
        pso_dataset_path = os.path.join(DATASET_DIR, "PSO_Dataset.xlsx")
        if os.path.exists(pso_dataset_path):
            pso_df = pd.read_excel(pso_dataset_path).fillna("")
            pso_df.columns = [clean_column_name(column) for column in pso_df.columns]
            pso_id_column = find_column(pso_df, ["PSO ID", "PSO"])
            pso_desc_column = find_column(
                pso_df,
                ["Programme Specific Outcomes", "Program Specific Outcomes", "PSO Description", "Description"]
            )
            if pso_id_column and pso_desc_column:
                pso_records = [
                    {"id": safe_string(row[pso_id_column]), "description": safe_string(row[pso_desc_column])}
                    for _, row in pso_df.iterrows()
                    if safe_string(row[pso_id_column]) and safe_string(row[pso_desc_column])
                ]

        # Regenerate from the current CO and manually maintained PSO Excel datasets.
        from pso_mapping import generate_pso_mapping
        pso_result = generate_pso_mapping(BASE_DIR, DATASET_DIR)
        pso_mapping_df = pso_result["mapping"].fillna("")
        for pso in pso_records:
            pso_id = pso["id"]
            if pso_id in pso_mapping_df.columns:
                numeric = pd.to_numeric(pso_mapping_df[pso_id], errors="coerce").fillna(0)
                pso_graph_labels.append(pso_id)
                pso_graph_values.append(round(float(numeric.mean()), 2) if len(numeric) else 0)
    except Exception as error:
        print("Dashboard CO-PSO error:", error)

    # ========================================================
    # BLOOM
    # ========================================================

    try:

        if os.path.exists(
            CO_FILE
        ):

            df = pd.read_excel(
                CO_FILE
            )

            df.columns = [
                clean_column_name(column)
                for column in df.columns
            ]

            statement_column = find_column(
                df,
                [
                    "Course Outcome",
                    "CO Statement",
                    "Course Outcome Statement",
                    "Statement",
                    "Outcome"
                ]
            )

            if statement_column:

                levels = []

                for statement in df[
                    statement_column
                ]:

                    try:

                        level = detect_bloom(
                            safe_string(
                                statement
                            )
                        )

                    except Exception:

                        level = "Unknown"

                    if not level:

                        level = "Unknown"

                    levels.append(
                        safe_string(level)
                    )

                counts = pd.Series(
                    levels
                ).value_counts()

                bloom_labels = (
                    counts.index.tolist()
                )

                bloom_values = [
                    int(value)
                    for value
                    in counts.values.tolist()
                ]

    except Exception as error:

        print(
            "Dashboard Bloom error:",
            error
        )

    # ========================================================
    # PO / PEO LISTS + DEPARTMENT NAME (from the selected department's datasets)
    # ========================================================

    def read_outcome_list(filename, id_names, text_names):
        items = []
        path = os.path.join(DATASET_DIR, filename)
        try:
            if os.path.exists(path):
                odf = pd.read_excel(path).fillna("")
                odf.columns = [clean_column_name(c) for c in odf.columns]
                id_col = find_column(odf, id_names)
                text_col = find_column(odf, text_names)
                if id_col and text_col:
                    for _, orow in odf.iterrows():
                        oid = safe_string(orow[id_col]).replace(" ", "")
                        text = safe_string(orow[text_col])
                        if not oid or not text:
                            continue
                        title, body = "", text
                        head_part, sep, rest = text.partition(":")
                        if sep and len(head_part) <= 60:
                            title, body = head_part.strip(), rest.strip()
                        items.append({"id": oid, "title": title, "description": body})
        except Exception as error:
            print("Dashboard outcome list error:", error)
        return items

    po_records = read_outcome_list(
        "PO_Dataset.xlsx", ["PO ID", "PO"],
        ["Programme Outcomes", "Program Outcomes", "PO Description", "PO Statement", "Description"])
    peo_records = read_outcome_list(
        "PEO_Dataset.xlsx", ["PEO ID", "PEO"],
        ["Programme Educational Objectives", "Program Educational Objectives",
         "PEO Description", "PEO Statement", "Description"])
    if po_records:
        po_count = len(po_records)
    if peo_records:
        peo_count = len(peo_records)

    department_name = dept_store.department_name(BASE_DIR, session.get("department"))

    # ========================================================
    # REPORT COUNT
    # ========================================================

    report_count = subject_count

    # ========================================================
    # RENDER DASHBOARD
    # ========================================================

    return render_template(

        "dashboard.html",

        subjects=subjects,

        subject_count=subject_count,

        co_count=co_count,

        po_count=po_count,

        peo_count=peo_count,

        report_count=report_count,

        total_po_mapping=total_po_mapping,

        total_peo_mapping=total_peo_mapping,

        ai_accuracy=ai_accuracy,

        graph_labels=graph_labels,

        graph_values=graph_values,

        peo_graph_labels=peo_graph_labels,

        peo_graph_values=peo_graph_values,

        bloom_labels=bloom_labels,

        bloom_values=bloom_values,

        co_po_labels=co_po_labels,

        co_po_values=co_po_values,

        co_peo_labels=co_peo_labels,

        co_peo_values=co_peo_values,

        pso_records=pso_records,

        po_records=po_records,

        peo_records=peo_records,

        department_name=department_name,

        pso_graph_labels=pso_graph_labels,

        pso_graph_values=pso_graph_values
    )


# ============================================================
# LOGOUT
# ============================================================


def add_word_section_heading(document, text, *args, **kwargs):
    """Safe Word heading that does not require Heading 1 style."""
    paragraph = document.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

    run = paragraph.add_run(str(text))
    run.bold = True
    run.font.size = Pt(14)

    return paragraph

@app.route(
    "/logout"
)
def logout():

    session.clear()

    return redirect("/")


# ============================================================
# GET DOWNLOAD SUBJECT
# ============================================================

def get_download_subject():

    subject_code = safe_string(
        request.args.get(
            "subject",
            ""
        )
    )

    if not subject_code:

        subject_code = safe_string(
            request.args.get(
                "subject_code",
                ""
            )
        )

    if not subject_code:

        subject_code = safe_string(
            session.get(
                "subject_code",
                ""
            )
        )

    if not subject_code:

        return ""

    return normalize_code(
        subject_code
    )


# ============================================================
# REPORT HELPER FUNCTIONS
# PDF + WORD + EXCEL
# ============================================================

def get_report_logo():

    """
    Returns the college logo path if available.
    """

    logo_candidates = [

        os.path.join(
            BASE_DIR,
            "static",
            "images",
            "ATME.jpeg"
        ),

        os.path.join(
            BASE_DIR,
            "static",
            "images",
            "ATME.jpg"
        ),

        os.path.join(
            BASE_DIR,
            "static",
            "images",
            "ATME.png"
        )
    ]

    for path in logo_candidates:

        if os.path.exists(path):

            return path

    return None


def get_report_college_name(details):

    """
    Gets college name from subject details.
    """

    possible_columns = [

        "College",
        "College Name",
        "Institute",
        "Institute Name",
        "CollegeName"
    ]

    for column in possible_columns:

        if column in details:

            value = safe_string(
                details.get(column)
            )

            if value:

                return value

    return "ATME College of Engineering"


def get_report_department(details):

    """
    Gets department name from subject details.
    """

    possible_columns = [

        "Department",
        "Department Name",
        "Branch",
        "Branch Name"
    ]

    for column in possible_columns:

        if column in details:

            value = safe_string(
                details.get(column)
            )

            if value:

                return value

    return "Computer Science and Engineering"


def get_report_subject_details(details):

    """
    Creates a clean subject-details list
    for PDF, Word and Excel.
    """

    return [

        [
            "College",
            get_report_college_name(details)
        ],

        [
            "Department",
            get_report_department(details)
        ],

        [
            "Subject Code",
            safe_string(
                details.get(
                    "Subject Code",
                    ""
                )
            )
        ],

        [
            "Subject Name",
            safe_string(
                details.get(
                    "Subject Name",
                    ""
                )
            )
        ],

        [
            "Semester",
            safe_string(
                details.get(
                    "Semester",
                    ""
                )
            )
        ],

        [
            "Credits",
            safe_string(
                details.get(
                    "Credits",
                    ""
                )
            )
        ],

        [
            "Scheme",
            safe_string(
                details.get(
                    "Scheme",
                    ""
                )
            ) or "N/A"
        ]
    ]


def get_report_po_details(
    co_po
):

    """
    Creates PO detail table from the PO columns
    available in the CO-PO mapping file.

    If a separate PO dataset is available,
    its descriptions are used.
    """

    columns = sort_mapping_columns(
        get_mapping_columns(
            co_po
        )
    )

    rows = []

    # --------------------------------------------------------
    # Try separate PO dataset
    # --------------------------------------------------------

    po_dataset_candidates = [

        os.path.join(
            DATASET_DIR,
            "PO_Dataset.xlsx"
        ),

        os.path.join(
            DATASET_DIR,
            "PO_Dataset-1.xlsx"
        ),

        os.path.join(
            DATASET_DIR,
            "PO.xlsx"
        )
    ]

    po_df = pd.DataFrame()

    for path in po_dataset_candidates:

        if os.path.exists(path):

            try:

                po_df = pd.read_excel(
                    path
                )

                break

            except Exception:

                po_df = pd.DataFrame()

    po_id_column = None
    po_statement_column = None

    if not po_df.empty:

        po_df.columns = [
            clean_column_name(column)
            for column in po_df.columns
        ]

        po_id_column = find_column(
            po_df,
            [
                "PO ID",
                "PO",
                "PO Code",
                "PO Number",
                "Program Outcome ID"
            ]
        )

        po_statement_column = find_column(
            po_df,
            [
                "PO Statement",
                "PO Description",
                "Program Outcome",
                "Programme Outcomes",
                "Program Outcomes",
                "Description",
                "Statement"
            ]
        )

    # --------------------------------------------------------
    # Build rows
    # --------------------------------------------------------

    for po_id in columns:

        description = ""

        if (
            not po_df.empty
            and po_id_column
            and po_statement_column
        ):

            for _, row in po_df.iterrows():

                current_id = safe_string(
                    row.get(
                        po_id_column,
                        ""
                    )
                ).upper().replace(
                    " ",
                    ""
                )

                target_id = safe_string(
                    po_id
                ).upper().replace(
                    " ",
                    ""
                )

                if current_id == target_id:

                    description = safe_string(
                        row.get(
                            po_statement_column,
                            ""
                        )
                    )

                    break

        if not description:

            description = (
                f"Program Outcome {po_id}"
            )

        rows.append({

            "PO ID":
                safe_string(po_id),

            "PO Statement":
                description
        })

    return rows


def get_report_peo_details(
    co_peo
):

    """
    Creates PEO detail table from PEO columns.
    """

    columns = sort_mapping_columns(
        get_mapping_columns(
            co_peo
        )
    )

    rows = []

    # --------------------------------------------------------
    # Try separate PEO dataset
    # --------------------------------------------------------

    peo_dataset_candidates = [

        os.path.join(
            DATASET_DIR,
            "PEO_Dataset.xlsx"
        ),

        os.path.join(
            DATASET_DIR,
            "PEO_Dataset-1.xlsx"
        ),

        os.path.join(
            DATASET_DIR,
            "PEO.xlsx"
        )
    ]

    peo_df = pd.DataFrame()

    for path in peo_dataset_candidates:

        if os.path.exists(path):

            try:

                peo_df = pd.read_excel(
                    path
                )

                break

            except Exception:

                peo_df = pd.DataFrame()

    peo_id_column = None
    peo_statement_column = None

    if not peo_df.empty:

        peo_df.columns = [
            clean_column_name(column)
            for column in peo_df.columns
        ]

        peo_id_column = find_column(
            peo_df,
            [
                "PEO ID",
                "PEO",
                "PEO Code",
                "PEO Number",
                "Program Educational Objective ID"
            ]
        )

        peo_statement_column = find_column(
            peo_df,
            [
                "PEO Statement",
                "PEO Description",
                "Program Educational Objective",
                "Programme Educational Objectives",
                "Program Educational Objectives",
                "Description",
                "Statement"
            ]
        )

    # --------------------------------------------------------
    # Build rows
    # --------------------------------------------------------

    for peo_id in columns:

        description = ""

        if (
            not peo_df.empty
            and peo_id_column
            and peo_statement_column
        ):

            for _, row in peo_df.iterrows():

                current_id = safe_string(
                    row.get(
                        peo_id_column,
                        ""
                    )
                ).upper().replace(
                    " ",
                    ""
                )

                target_id = safe_string(
                    peo_id
                ).upper().replace(
                    " ",
                    ""
                )

                if current_id == target_id:

                    description = safe_string(
                        row.get(
                            peo_statement_column,
                            ""
                        )
                    )

                    break

        if not description:

            description = (
                f"Program Educational Objective {peo_id}"
            )

        rows.append({

            "PEO ID":
                safe_string(peo_id),

            "PEO Statement":
                description
        })

    return rows


def report_mapping_scale():

    return [

        {
            "Value": 3,
            "Level": "High",
            "Description":
                "Strong relationship"
        },

        {
            "Value": 2,
            "Level": "Moderate",
            "Description":
                "Moderate relationship"
        },

        {
            "Value": 1,
            "Level": "Low",
            "Description":
                "Low relationship"
        },

        {
            "Value": 0,
            "Level": "None",
            "Description":
                "No significant relationship"
        }
    ]


def report_co_po_matrix(
    co,
    co_po
):

    rows = []

    if (
        co_po is None
        or co_po.empty
    ):

        return rows

    columns = sort_mapping_columns(
        get_mapping_columns(
            co_po
        )
    )

    for _, mapping_row in co_po.iterrows():

        co_id = get_mapping_co_id(
            mapping_row
        )

        if not co_id:

            continue

        row = {

            "CO ID":
                co_id
        }

        for column in columns:

            row[
                safe_string(column)
            ] = safe_int(
                mapping_row.get(
                    column,
                    0
                )
            )

        rows.append(row)

    return rows


def report_co_peo_matrix(
    co,
    co_peo
):

    rows = []

    if (
        co_peo is None
        or co_peo.empty
    ):

        return rows

    columns = sort_mapping_columns(
        get_mapping_columns(
            co_peo
        )
    )

    for _, mapping_row in co_peo.iterrows():

        co_id = get_mapping_co_id(
            mapping_row
        )

        if not co_id:

            continue

        row = {

            "CO ID":
                co_id
        }

        for column in columns:

            row[
                safe_string(column)
            ] = safe_int(
                mapping_row.get(
                    column,
                    0
                )
            )

        rows.append(row)

    return rows


def report_pdf_paragraph(
    text,
    style
):

    """
    Safely converts text into a ReportLab Paragraph.
    """

    text = safe_string(text)

    text = (
        text
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )

    return Paragraph(
        text,
        style
    )


def report_pdf_table(
    data,
    widths=None,
    font_size=7
):

    """
    Creates a professional ReportLab table.
    """

    if not data:

        return None

    styles = getSampleStyleSheet()

    # Use independent styles.  Previously both variables referenced the
    # same BodyText style, so setting the header text to white also made
    # every table data cell white/invisible.
    cell_style = ParagraphStyle(
        "ReportTableCell",
        parent=styles["BodyText"],
        fontName="Helvetica",
        fontSize=font_size,
        leading=font_size + 2,
        textColor=colors.black
    )

    header_style = ParagraphStyle(
        "ReportTableHeader",
        parent=styles["BodyText"],
        fontName="Helvetica-Bold",
        fontSize=font_size,
        leading=font_size + 2,
        textColor=colors.white
    )

    formatted_data = []

    for row_index, row in enumerate(data):

        formatted_row = []

        for value in row:

            if row_index == 0:

                formatted_row.append(
                    report_pdf_paragraph(
                        value,
                        header_style
                    )
                )

            else:

                formatted_row.append(
                    report_pdf_paragraph(
                        value,
                        cell_style
                    )
                )

        formatted_data.append(
            formatted_row
        )

    # The college PDF template is portrait (612 x 792 pt).
    # Automatically scale wide tables so no content is clipped.
    fitted_widths = widths
    if widths:
        available_width = 612 - 25 - 25
        total_width = sum(widths)
        if total_width > available_width:
            scale = available_width / total_width
            fitted_widths = [w * scale for w in widths]

    table = Table(
        formatted_data,
        colWidths=fitted_widths,
        repeatRows=1,
        hAlign="LEFT"
    )

    table.setStyle(
        TableStyle([

            (
                "BACKGROUND",
                (0, 0),
                (-1, 0),
                colors.HexColor("#1A6B69")
            ),

            (
                "TEXTCOLOR",
                (0, 0),
                (-1, 0),
                colors.white
            ),

            (
                "FONTNAME",
                (0, 0),
                (-1, 0),
                "Helvetica-Bold"
            ),

            (
                "GRID",
                (0, 0),
                (-1, -1),
                0.4,
                colors.grey
            ),

            (
                "VALIGN",
                (0, 0),
                (-1, -1),
                "TOP"
            ),

            (
                "LEFTPADDING",
                (0, 0),
                (-1, -1),
                5
            ),

            (
                "RIGHTPADDING",
                (0, 0),
                (-1, -1),
                5
            ),

            (
                "TOPPADDING",
                (0, 0),
                (-1, -1),
                5
            ),

            (
                "BOTTOMPADDING",
                (0, 0),
                (-1, -1),
                5
            )
        ])
    )

    return table


def add_word_table(
    document,
    headers,
    rows,
    widths=None
):

    """
    Adds a clean, fixed-width table to Word: sensible column widths, shaded repeating
    header row, rows that never split across pages, centred ID / value columns.
    """

    if not headers:

        return

    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import RGBColor

    section = document.sections[-1]

    usable = (
        section.page_width
        - section.left_margin
        - section.right_margin
    ) / 914400.0  # inches

    ncols = len(headers)
    names = [safe_string(h).strip().lower() for h in headers]

    import re as _re
    is_matrix = (
        ncols >= 3
        and names[0] == "co id"
        and all(_re.fullmatch(r"(po|peo|pso)\s*\d+", n) for n in names[1:])
    ) or ncols > 6

    # ---------------- column widths (inches) ----------------
    if widths is None:

        narrow = {
            "co id": 0.8, "po id": 0.8, "peo id": 0.8, "pso id": 0.8,
            "value": 0.65, "level": 0.95, "strength": 0.95,
            "bloom level": 1.1, "mapping value": 0.8, "field": 1.8,
        }

        if is_matrix:
            # mapping matrix: ID column + equal narrow value columns (capped so few columns stay compact)
            first = 0.9
            rest = min((usable - first) / (ncols - 1), 1.1)
            widths = [first] + [rest] * (ncols - 1)

        else:
            fixed = {i: narrow[n] for i, n in enumerate(names) if n in narrow}
            weights = {}
            for i, n in enumerate(names):
                if i in fixed:
                    continue
                if "justification" in n or "reason" in n:
                    weights[i] = 3.0
                elif any(k in n for k in ("statement", "outcome", "description", "details")):
                    weights[i] = 2.0
                else:
                    weights[i] = 1.0
            remaining = max(usable - sum(fixed.values()), 1.0)
            total_w = sum(weights.values()) or 1.0
            widths = [
                fixed[i] if i in fixed else remaining * weights[i] / total_w
                for i in range(ncols)
            ]

    if not is_matrix:
        scale = usable / sum(widths)
        widths = [w * scale for w in widths]

    centre_cols = {
        i for i, n in enumerate(names)
        if n in ("co id", "po id", "peo id", "pso id", "value", "level", "strength",
                 "bloom level", "mapping value") or is_matrix
    }

    table = document.add_table(
        rows=1,
        cols=ncols
    )

    table.style = "Table Grid"
    table.autofit = False
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    tbl_pr = table._tbl.tblPr

    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tbl_pr.append(layout)

    margins = OxmlElement("w:tblCellMar")
    for side, val in (("top", 40), ("left", 80), ("bottom", 40), ("right", 80)):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:w"), str(val))
        el.set(qn("w:type"), "dxa")
        margins.append(el)
    tbl_pr.append(margins)

    def shade(cell, fill):
        tc_pr = cell._tc.get_or_add_tcPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), fill)
        tc_pr.append(shd)

    def write(cell, text, size, bold=False, color=None, centre=False):
        cell.text = ""
        paragraph = cell.paragraphs[0]
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(0)
        if centre:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = paragraph.add_run(safe_string(text))
        run.font.size = Pt(size)
        run.bold = bold
        if color:
            run.font.color.rgb = RGBColor(*color)

    def keep_row_together(row, header=False):
        tr_pr = row._tr.get_or_add_trPr()
        tr_pr.append(OxmlElement("w:cantSplit"))
        if header:
            tr_pr.append(OxmlElement("w:tblHeader"))

    # Header row (repeats on every page)
    for index, header in enumerate(headers):

        cell = table.rows[0].cells[index]
        write(cell, header, 9, bold=True, color=(255, 255, 255), centre=index in centre_cols)
        shade(cell, "1F4E78")

    keep_row_together(table.rows[0], header=True)

    # Data rows
    for r_index, row in enumerate(rows):

        new_row = table.add_row()
        keep_row_together(new_row)
        cells = new_row.cells

        for index in range(ncols):

            value = row[index] if index < len(row) else ""
            write(cells[index], value, 8.5, centre=index in centre_cols)

            if r_index % 2 == 1:
                shade(cells[index], "F2F6FA")

    # Fixed widths on the grid and on every cell
    for index, width in enumerate(widths):

        table.columns[index].width = Inches(width)

        for row in table.rows:
            row.cells[index].width = Inches(width)

    # Short tables stay on one page (every row except the last keeps with the next)
    if len(table.rows) <= 12:
        for row in table.rows[:-1]:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    paragraph.paragraph_format.keep_with_next = True

    spacer = document.add_paragraph()
    spacer.paragraph_format.space_after = Pt(4)

    return table


def style_excel_sheet(
    worksheet
):

    """
    Applies professional formatting to Excel.
    """

    thin = Side(
        style="thin",
        color="B7B7B7"
    )

    border = Border(
        left=thin,
        right=thin,
        top=thin,
        bottom=thin
    )

    header_fill = PatternFill(
        "solid",
        fgColor="1F4E78"
    )

    header_font = Font(
        bold=True,
        color="FFFFFF"
    )

    for row in worksheet.iter_rows():

        for cell in row:

            cell.border = border

            cell.alignment = Alignment(
                vertical="top",
                wrap_text=True
            )

    if worksheet.max_row >= 1:

        for cell in worksheet[1]:

            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True
            )

    # Column widths
    for column_cells in worksheet.columns:

        max_length = 0

        column_letter = (
            column_cells[0]
            .column_letter
        )

        for cell in column_cells:

            try:

                length = len(
                    str(
                        cell.value
                        if cell.value is not None
                        else ""
                    )
                )

                if length > max_length:

                    max_length = length

            except Exception:

                pass

        worksheet.column_dimensions[
            column_letter
        ].width = min(
            max(
                max_length + 3,
                12
            ),
            45
        )

    worksheet.freeze_panes = "A2"

    worksheet.auto_filter.ref = (
        worksheet.dimensions
    )


# ============================================================
# PDF GRAPH HELPERS
# ============================================================

def report_pdf_bar_chart(labels, values, title, y_label="Average Mapping Value"):
    """Create a compact, printable bar chart for the PDF report."""
    if not labels or not values:
        return None

    clean_labels = [safe_string(x) for x in labels]
    clean_values = []
    for value in values:
        try:
            clean_values.append(float(value))
        except Exception:
            clean_values.append(0.0)

    width = 540
    height = 230
    drawing = Drawing(width, height)

    chart = VerticalBarChart()
    chart.x = 45
    chart.y = 45
    chart.width = 475
    chart.height = 145
    chart.data = [clean_values]
    chart.categoryAxis.categoryNames = clean_labels
    chart.categoryAxis.labels.fontName = "Helvetica"
    chart.categoryAxis.labels.fontSize = 7
    chart.categoryAxis.labels.angle = 45 if len(clean_labels) > 7 else 0
    chart.categoryAxis.labels.dy = -12 if len(clean_labels) > 7 else -2
    chart.valueAxis.valueMin = 0
    chart.valueAxis.valueMax = max(3.0, max(clean_values) + 0.5)
    chart.valueAxis.valueStep = 1
    chart.valueAxis.labels.fontName = "Helvetica"
    chart.valueAxis.labels.fontSize = 7
    chart.barWidth = max(8, min(24, 430 / max(1, len(clean_labels))))
    chart.groupSpacing = 8
    chart.barSpacing = 2
    chart.valueAxis.visibleGrid = True
    chart.strokeColor = colors.grey
    chart.fillColor = colors.white
    chart.bars[0].fillColor = colors.HexColor("#238987")
    chart.bars[0].strokeColor = colors.HexColor("#1A6B69")

    drawing.add(chart)
    drawing.add(String(width / 2, 208, title, textAnchor="middle",
                       fontName="Helvetica-Bold", fontSize=11))
    drawing.add(String(18, 118, y_label, textAnchor="middle",
                       fontName="Helvetica", fontSize=7, angle=90))
    drawing.add(String(width / 2, 10, "Mapping Categories", textAnchor="middle",
                       fontName="Helvetica", fontSize=7))
    return drawing


# ============================================================
# DOWNLOAD PDF
# ============================================================


def add_word_section_heading(document, text, *args, **kwargs):
    """Safe Word heading that does not require Heading 1 style."""
    paragraph = document.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

    run = paragraph.add_run(str(text))
    run.bold = True
    run.font.size = Pt(14)

    return paragraph


def load_subject_pso_report_data(subject_code):
    """Return PSO definitions, selected-subject CO–PSO mappings and justifications."""
    from pso_mapping import generate_pso_mapping
    result = generate_pso_mapping(BASE_DIR, DATASET_DIR)
    dataset_path = os.path.join(DATASET_DIR, "PSO_Dataset.xlsx")
    pso_dataset = pd.read_excel(dataset_path).fillna("") if os.path.exists(dataset_path) else pd.DataFrame()
    mapping = result.get("mapping", pd.DataFrame()).fillna("")
    justifications = result.get("justifications", pd.DataFrame()).fillna("")
    code = str(subject_code).upper().strip()
    if "Subject Code" in mapping.columns:
        mapping = mapping[mapping["Subject Code"].astype(str).str.upper().str.strip() == code].copy()
    else:
        mapping = pd.DataFrame()
    if "Subject Code" in justifications.columns:
        justifications = justifications[justifications["Subject Code"].astype(str).str.upper().str.strip() == code].copy()
    else:
        justifications = pd.DataFrame()
    return pso_dataset, mapping, justifications

@app.route(
    "/download_pdf"
)
def download_pdf():

    if "user" not in session:

        return redirect("/")

    subject_code = (
        get_download_subject()
    )

    if not subject_code:

        return redirect(
            "/dashboard"
        )

    try:

        # ----------------------------------------------------
        # LOAD DATA
        # ----------------------------------------------------

        details = get_subject_details(
            subject_code
        )

        co = load_course_outcomes(
            subject_code
        )

        co_po = load_co_po(
            subject_code
        )

        co_peo = load_co_peo(
            subject_code
        )

        po_justification = (
            load_justification(
                CO_PO_JUSTIFICATION_FILE,
                subject_code
            )
        )

        peo_justification = (
            load_justification(
                CO_PEO_JUSTIFICATION_FILE,
                subject_code
            )
        )

    except Exception as error:

        print(
            "PDF REPORT ERROR:",
            error
        )

        return (
            safe_string(error),
            500
        )

    # --------------------------------------------------------
    # REPORT DATA
    # --------------------------------------------------------

    co_rows = build_co_rows(
        co
    )

    po_rows = get_report_po_details(
        co_po
    )

    peo_rows = get_report_peo_details(
        co_peo
    )

    po_reason_rows = (
        build_po_reason_rows(
            co,
            co_po,
            po_justification
        )
    )

    peo_reason_rows = (
        build_peo_reason_rows(
            co,
            co_peo,
            peo_justification
        )
    )

    bloom_rows = (
        build_bloom_reason_rows(
            co
        )
    )

    try:
        pso_dataset_pdf, pso_mapping_pdf, pso_just_pdf = load_subject_pso_report_data(subject_code)
    except Exception as error:
        print("PDF PSO REPORT DATA ERROR:", error)
        pso_dataset_pdf, pso_mapping_pdf, pso_just_pdf = pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    # --------------------------------------------------------
    # FILE
    # --------------------------------------------------------

    filename = os.path.join(
        BASE_DIR,
        f"{subject_code}_Subject_Report.pdf"
    )

    # --------------------------------------------------------
    # PDF
    # --------------------------------------------------------

    # Match the official college template page size (US Letter portrait).
    # The template contains the official header, watermark and footer.
    document = SimpleDocTemplate(
        filename,
        pagesize=(612, 792),
        rightMargin=25,
        leftMargin=25,
        topMargin=125,
        bottomMargin=70
    )

    styles = getSampleStyleSheet()

    title_style = styles["Title"]

    title_style.fontName = (
        "Helvetica-Bold"
    )

    title_style.fontSize = 18
    title_style.leading = 22
    title_style.alignment = 1

    heading_style = styles["Heading2"]

    heading_style.fontName = (
        "Helvetica-Bold"
    )

    heading_style.textColor = (
        colors.HexColor("#1A6B69")
    )

    heading_style.spaceBefore = 10
    heading_style.spaceAfter = 8

    normal_style = styles["BodyText"]

    normal_style.fontSize = 9
    normal_style.leading = 12

    story = []

    # ========================================================
    # REPORT TITLE
    # Official college branding is supplied by College_Template.pdf
    # and merged onto every generated page after ReportLab finishes.
    # ========================================================

    story.append(
        Spacer(1, 8)
    )

    story.append(
        Paragraph(
            "AI-Based Automatic CO-PO-PEO Mapping System",
            title_style
        )
    )

    story.append(
        Spacer(1, 15)
    )

    # ========================================================
    # SUBJECT DETAILS
    # ========================================================

    story.append(
        Paragraph(
            "1. Subject Details",
            heading_style
        )
    )

    subject_details = (
        get_report_subject_details(
            details
        )
    )

    subject_data = [
        [
            "Field",
            "Details"
        ]
    ]

    for item in subject_details:

        subject_data.append([
            item[0],
            item[1]
        ])

    subject_table = report_pdf_table(
        subject_data,
        widths=[
            2.0 * inch,
            7.0 * inch
        ],
        font_size=8
    )

    if subject_table:

        story.append(
            subject_table
        )

    # ========================================================
    # COURSE OUTCOMES
    # ========================================================

    story.append(
        Spacer(1, 12)
    )

    story.append(
        Paragraph(
            "2. Course Outcomes and Bloom's Taxonomy",
            heading_style
        )
    )

    co_data = [

        [
            "CO ID",
            "Course Outcome",
            "Bloom Level",
            "Bloom Reason"
        ]
    ]

    for row in co_rows:

        co_data.append([

            row.get(
                "CO ID",
                ""
            ),

            row.get(
                "CO Statement",
                ""
            ),

            row.get(
                "Bloom Level",
                ""
            ),

            row.get(
                "Bloom Reason",
                ""
            )
        ])

    if len(co_data) > 1:

        table = report_pdf_table(

            co_data,

            widths=[
                0.65 * inch,
                3.1 * inch,
                0.85 * inch,
                5.0 * inch
            ],

            font_size=7
        )

        if table:

            story.append(
                table
            )

    else:

        story.append(
            Paragraph(
                "No Course Outcomes available.",
                normal_style
            )
        )

    # ========================================================
    # PO DETAILS
    # ========================================================

    story.append(
        Spacer(1, 15)
    )

    story.append(
        Paragraph(
            "3. Program Outcomes (PO)",
            heading_style
        )
    )

    po_data = [
        [
            "PO ID",
            "PO Statement"
        ]
    ]

    for row in po_rows:

        po_data.append([
            row.get(
                "PO ID",
                ""
            ),

            row.get(
                "PO Statement",
                ""
            )
        ])

    if len(po_data) > 1:

        table = report_pdf_table(

            po_data,

            widths=[
                1.0 * inch,
                8.0 * inch
            ],

            font_size=8
        )

        if table:

            story.append(
                table
            )

    # ========================================================
    # PEO DETAILS
    # ========================================================

    story.append(
        Spacer(1, 15)
    )

    story.append(
        Paragraph(
            "4. Program Educational Objectives (PEO)",
            heading_style
        )
    )

    peo_data = [
        [
            "PEO ID",
            "PEO Statement"
        ]
    ]

    for row in peo_rows:

        peo_data.append([
            row.get(
                "PEO ID",
                ""
            ),

            row.get(
                "PEO Statement",
                ""
            )
        ])

    if len(peo_data) > 1:

        table = report_pdf_table(

            peo_data,

            widths=[
                1.0 * inch,
                8.0 * inch
            ],

            font_size=8
        )

        if table:

            story.append(
                table
            )

    # ========================================================
    # PSO DETAILS (statements from the department's PSO dataset)
    # ========================================================

    story.append(Spacer(1, 15))
    story.append(Paragraph("5. Program Specific Outcomes (PSO)", heading_style))
    _pso_id_col = find_column(pso_dataset_pdf, ["PSO ID", "PSO"]) if not pso_dataset_pdf.empty else None
    _pso_text_col = find_column(
        pso_dataset_pdf,
        ["Programme Specific Outcomes", "Program Specific Outcomes", "PSO Description", "PSO Statement", "Description"]
    ) if not pso_dataset_pdf.empty else None
    pso_data = [["PSO ID", "PSO Statement"]]
    if _pso_id_col and _pso_text_col:
        for _, _row in pso_dataset_pdf.iterrows():
            if safe_string(_row.get(_pso_id_col, "")):
                pso_data.append([safe_string(_row.get(_pso_id_col, "")), safe_string(_row.get(_pso_text_col, ""))])
    if len(pso_data) > 1:
        _t = report_pdf_table(pso_data, widths=[1.0 * inch, 8.0 * inch], font_size=8)
        if _t:
            story.append(_t)
    else:
        story.append(Paragraph("PSO statements are not available.", normal_style))

    # ========================================================
    # CO-PO MAPPING
    # ========================================================

    story.append(
        PageBreak()
    )

    story.append(
        Paragraph(
            "6. CO-PO Mapping Matrix",
            heading_style
        )
    )

    po_columns = sort_mapping_columns(
        get_mapping_columns(
            co_po
        )
    )

    matrix_data = [

        [
            "CO ID"
        ]
        +
        [
            safe_string(
                column
            )
            for column in po_columns
        ]
    ]

    for row in report_co_po_matrix(
        co,
        co_po
    ):

        matrix_data.append([

            safe_string(
                row.get(
                    "CO ID",
                    ""
                )
            )

        ]
        +
        [
            safe_string(
                row.get(
                    column,
                    0
                )
            )

            for column in po_columns
        ])

    if len(matrix_data) > 1:

        matrix_widths = [
            0.65 * inch
        ]

        if po_columns:

            remaining_width = (
                9.0 / len(po_columns)
            )

            matrix_widths += [
                remaining_width * inch
            ] * len(po_columns)

        table = report_pdf_table(

            matrix_data,

            widths=matrix_widths,

            font_size=7
        )

        if table:

            story.append(
                table
            )

    # ========================================================
    # CO-PEO MAPPING
    # ========================================================

    story.append(
        Spacer(1, 15)
    )

    story.append(
        Paragraph(
            "7. CO-PEO Mapping Matrix",
            heading_style
        )
    )

    peo_columns = sort_mapping_columns(
        get_mapping_columns(
            co_peo
        )
    )

    matrix_data = [

        [
            "CO ID"
        ]
        +
        [
            safe_string(
                column
            )
            for column in peo_columns
        ]
    ]

    for row in report_co_peo_matrix(
        co,
        co_peo
    ):

        matrix_data.append([

            safe_string(
                row.get(
                    "CO ID",
                    ""
                )
            )

        ]
        +
        [
            safe_string(
                row.get(
                    column,
                    0
                )
            )

            for column in peo_columns
        ])

    if len(matrix_data) > 1:

        matrix_widths = [
            0.65 * inch
        ]

        if peo_columns:

            remaining_width = (
                9.0 / len(peo_columns)
            )

            matrix_widths += [
                remaining_width * inch
            ] * len(peo_columns)

        table = report_pdf_table(

            matrix_data,

            widths=matrix_widths,

            font_size=7
        )

        if table:

            story.append(
                table
            )

    # ========================================================
    # CO-PSO MAPPING MATRIX (subject report only)
    # ========================================================

    story.append(Spacer(1, 15))
    story.append(Paragraph("8. CO-PSO Mapping Matrix", heading_style))
    pso_id_col = find_column(pso_dataset_pdf, ["PSO ID", "PSO"]) if not pso_dataset_pdf.empty else None
    pso_columns = [safe_string(v) for v in pso_dataset_pdf[pso_id_col].tolist() if safe_string(v)] if pso_id_col else []
    pso_matrix_data = [["CO ID"] + pso_columns]
    for _, pso_row in pso_mapping_pdf.iterrows():
        pso_matrix_data.append([safe_string(pso_row.get("CO ID", ""))] + [safe_string(pso_row.get(pid, 0)) for pid in pso_columns])
    if len(pso_matrix_data) > 1:
        pso_widths = [0.65 * inch] + [(8.8 / max(len(pso_columns), 1)) * inch] * len(pso_columns)
        pso_table = report_pdf_table(pso_matrix_data, widths=pso_widths, font_size=7)
        if pso_table:
            story.append(pso_table)

    # ========================================================
    # GRAPHICAL ANALYSIS
    # ========================================================

    story.append(PageBreak())
    story.append(
        Paragraph(
            "9. Graphical Analysis",
            heading_style
        )
    )
    story.append(
        Paragraph(
            "The following charts summarize the generated CO-PO, CO-PEO, CO-PSO and Bloom's Taxonomy results for the selected subject.",
            normal_style
        )
    )
    story.append(Spacer(1, 8))

    co_po_chart = report_pdf_bar_chart(
        *numeric_average_data(co_po),
        "CO-PO Mapping Analysis"
    )
    if co_po_chart:
        story.append(co_po_chart)
        story.append(Spacer(1, 10))

    co_peo_chart = report_pdf_bar_chart(
        *numeric_average_data(co_peo),
        "CO-PEO Mapping Analysis"
    )
    if co_peo_chart:
        story.append(co_peo_chart)
        story.append(Spacer(1, 10))

    # CO-PSO chart (average mapping value per PSO for the selected subject)
    _pso_chart_labels, _pso_chart_values = [], []
    for _pid in pso_columns:
        if _pid in pso_mapping_pdf.columns:
            _vals = [safe_int(v) for v in pso_mapping_pdf[_pid].tolist()]
            _pso_chart_labels.append(_pid)
            _pso_chart_values.append(round(sum(_vals) / len(_vals), 2) if _vals else 0)
    co_pso_chart = report_pdf_bar_chart(
        _pso_chart_labels,
        _pso_chart_values,
        "CO-PSO Mapping Analysis"
    )
    if co_pso_chart:
        story.append(co_pso_chart)
        story.append(Spacer(1, 10))

    if bloom_rows:
        bloom_counts = {}
        for row in bloom_rows:
            level = safe_string(row.get("Bloom Level", "Unknown")) or "Unknown"
            bloom_counts[level] = bloom_counts.get(level, 0) + 1
        bloom_chart = report_pdf_bar_chart(
            list(bloom_counts.keys()),
            list(bloom_counts.values()),
            "Bloom's Taxonomy Distribution",
            "Number of COs"
        )
        if bloom_chart:
            story.append(bloom_chart)

    # ========================================================
    # CO-PO JUSTIFICATION
    # ========================================================

    story.append(
        PageBreak()
    )

    story.append(
        Paragraph(
            "10. CO-PO Mapping Justification",
            heading_style
        )
    )

    justification_data = [

        [
            "CO ID",
            "PO ID",
            "Value",
            "Level",
            "Justification"
        ]
    ]

    for row in po_reason_rows:

        justification_data.append([

            row.get(
                "CO ID",
                ""
            ),

            row.get(
                "PO ID",
                ""
            ),

            safe_string(
                row.get(
                    "Mapping Value",
                    0
                )
            ),

            row.get(
                "Strength",
                ""
            ),

            row.get(
                "Reason",
                ""
            )
        ])

    if len(justification_data) > 1:

        table = report_pdf_table(

            justification_data,

            widths=[
                0.65 * inch,
                0.65 * inch,
                0.55 * inch,
                0.85 * inch,
                6.3 * inch
            ],

            font_size=6.5
        )

        if table:

            story.append(
                table
            )

    # ========================================================
    # CO-PEO JUSTIFICATION
    # ========================================================

    story.append(
        Spacer(1, 15)
    )

    story.append(
        Paragraph(
            "11. CO-PEO Mapping Justification",
            heading_style
        )
    )

    justification_data = [

        [
            "CO ID",
            "PEO ID",
            "Value",
            "Level",
            "Justification"
        ]
    ]

    for row in peo_reason_rows:

        justification_data.append([

            row.get(
                "CO ID",
                ""
            ),

            row.get(
                "PEO ID",
                ""
            ),

            safe_string(
                row.get(
                    "Mapping Value",
                    0
                )
            ),

            row.get(
                "Strength",
                ""
            ),

            row.get(
                "Reason",
                ""
            )
        ])

    if len(justification_data) > 1:

        table = report_pdf_table(

            justification_data,

            widths=[
                0.65 * inch,
                0.7 * inch,
                0.55 * inch,
                0.85 * inch,
                6.25 * inch
            ],

            font_size=6.5
        )

        if table:

            story.append(
                table
            )

    # ========================================================
    # CO-PSO JUSTIFICATION
    # ========================================================

    story.append(Spacer(1, 15))
    story.append(Paragraph("12. CO-PSO Mapping Justification", heading_style))
    pso_just_data = [["CO ID", "PSO ID", "Value", "Level", "Justification"]]
    for _, row in pso_just_pdf.iterrows():
        pso_just_data.append([safe_string(row.get("CO ID", "")), safe_string(row.get("PSO ID", "")), safe_string(row.get("Mapping Level", 0)), safe_string(row.get("Strength", "")), safe_string(row.get("Justification", ""))])
    if len(pso_just_data) > 1:
        pso_just_table = report_pdf_table(pso_just_data, widths=[0.65*inch, 0.7*inch, 0.55*inch, 0.85*inch, 6.25*inch], font_size=6.5)
        if pso_just_table:
            story.append(pso_just_table)

    # ========================================================
    # MAPPING SCALE
    # ========================================================

    story.append(
        Spacer(1, 15)
    )

    story.append(
        Paragraph(
            "13. Mapping Scale",
            heading_style
        )
    )

    scale_data = [

        [
            "Value",
            "Level",
            "Description"
        ]
    ]

    for item in report_mapping_scale():

        scale_data.append([

            str(
                item["Value"]
            ),

            item["Level"],

            item["Description"]
        ])

    table = report_pdf_table(

        scale_data,

        widths=[
            1.0 * inch,
            2.0 * inch,
            6.0 * inch
        ],

        font_size=8
    )

    if table:

        story.append(
            table
        )

    # ========================================================
    # FACULTY VALIDATION
    # ========================================================

    story.append(
        Spacer(1, 35)
    )

    signature_data = [

        [
            "Faculty Name: ______________________________",
            "Faculty Signature: ______________________________"
        ],

        [
            "Date: ______________________________",
            "HOD Signature: ______________________________"
        ]
    ]

    signature_table = Table(
        signature_data,
        colWidths=[
            3.65 * inch,
            3.65 * inch
        ]
    )

    signature_table.setStyle(
        TableStyle([

            (
                "VALIGN",
                (0, 0),
                (-1, -1),
                "MIDDLE"
            ),

            (
                "TOPPADDING",
                (0, 0),
                (-1, -1),
                12
            ),

            (
                "BOTTOMPADDING",
                (0, 0),
                (-1, -1),
                12
            )
        ])
    )

    story.append(
        signature_table
    )

    # ========================================================
    # BUILD PDF AND APPLY OFFICIAL COLLEGE TEMPLATE
    # ========================================================

    document.build(story)

    if not os.path.exists(COLLEGE_PDF_TEMPLATE):
        raise FileNotFoundError(
            f"College PDF template not found: {COLLEGE_PDF_TEMPLATE}"
        )

    # Put the college template behind every report page.  The generated
    # report remains the foreground layer, so the template's white body
    # never covers the actual CO/PO/PEO data.
    content_reader = PdfReader(filename)
    template_reader = PdfReader(COLLEGE_PDF_TEMPLATE)
    template_page = template_reader.pages[0]
    writer = PdfWriter()

    # IMPORTANT: Do not deepcopy() the template PageObject here.  pypdf can
    # retain shared indirect-object references when a PDF page is deep-copied,
    # which can cause every output page to display the first report page.
    # Create a fresh blank page for every report page and merge the template
    # first, then the report content on top.
    from pypdf._page import PageObject

    for content_page in content_reader.pages:
        background_page = PageObject.create_blank_page(
            width=float(template_page.mediabox.width),
            height=float(template_page.mediabox.height)
        )
        background_page.merge_page(template_page)
        background_page.merge_page(content_page)
        writer.add_page(background_page)

    temp_filename = filename + ".tmp"
    with open(temp_filename, "wb") as output_file:
        writer.write(output_file)

    os.replace(temp_filename, filename)

    return send_file(

        filename,

        as_attachment=True,

        download_name=(
            f"{subject_code}_Subject_Report.pdf"
        ),

        mimetype="application/pdf"
    )

# ============================================================
# DOWNLOAD WORD
# ============================================================

def add_word_section_heading(document, text, level=1):
    """
    Safe Word heading.
    Does not depend on the 'Heading 1' style
    from the Word template.
    """

    paragraph = document.add_paragraph()

    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

    run = paragraph.add_run(str(text))

    run.bold = True

    if level == 1:
        run.font.size = Pt(14)
    elif level == 2:
        run.font.size = Pt(12)
    else:
        run.font.size = Pt(11)

    paragraph.paragraph_format.space_before = Pt(8)
    paragraph.paragraph_format.space_after = Pt(5)
    paragraph.paragraph_format.keep_with_next = True

    return paragraph


# ============================================================
# DOWNLOAD WORD
# ============================================================

@app.route("/download_word")
def download_word():

    if "user" not in session:
        return redirect("/")

    subject_code = get_download_subject()

    if not subject_code:
        return redirect("/dashboard")

    # ========================================================
    # LOAD REPORT DATA
    # ========================================================

    try:

        details = get_subject_details(
            subject_code
        )

        co = load_course_outcomes(
            subject_code
        )

        co_po = load_co_po(
            subject_code
        )

        co_peo = load_co_peo(
            subject_code
        )

        po_justification = load_justification(
            CO_PO_JUSTIFICATION_FILE,
            subject_code
        )

        peo_justification = load_justification(
            CO_PEO_JUSTIFICATION_FILE,
            subject_code
        )

    except Exception as error:

        print(
            "WORD REPORT ERROR:",
            error
        )

        return (
            safe_string(error),
            500
        )

    # ========================================================
    # REPORT DATA
    # ========================================================

    try:

        co_rows = build_co_rows(
            co
        )

        po_rows = get_report_po_details(
            co_po
        )

        peo_rows = get_report_peo_details(
            co_peo
        )

        po_reason_rows = build_po_reason_rows(
            co,
            co_po,
            po_justification
        )

        peo_reason_rows = build_peo_reason_rows(
            co,
            co_peo,
            peo_justification
        )
        pso_dataset_word, pso_mapping_word, pso_just_word = load_subject_pso_report_data(subject_code)

    except Exception as error:

        print(
            "WORD REPORT DATA ERROR:",
            error
        )

        return (
            safe_string(error),
            500
        )

    # ========================================================
    # OUTPUT FILE
    # ========================================================

    filename = os.path.join(
        BASE_DIR,
        f"{subject_code}_Subject_Report.docx"
    )

    # ========================================================
    # WORD TEMPLATE
    # ========================================================

    template_path = os.path.join(
        BASE_DIR,
        "report_templates",
        "College_Template.docx"
    )

    try:

        if os.path.exists(template_path):

            document = Document(
                template_path
            )

        else:

            print(
                "WORD TEMPLATE NOT FOUND:",
                template_path
            )

            document = Document()

    except Exception as error:

        print(
            "WORD TEMPLATE ERROR:",
            error
        )

        return (
            safe_string(error),
            500
        )

    # ========================================================
    # PAGE SETTINGS
    # ========================================================

    section = document.sections[0]

    section.orientation = WD_ORIENT.LANDSCAPE

    section.page_width = Inches(11.69)
    section.page_height = Inches(8.27)

    # top margin must clear the college letterhead stored in the template header
    section.top_margin = Inches(1.65)
    section.header_distance = Inches(0.2)
    section.bottom_margin = Inches(0.7)

    section.left_margin = Inches(0.6)
    section.right_margin = Inches(0.6)

    # The college logo and name come from the letterhead in College_Template.docx,
    # so they are not added again here (that made them overprint).

    # ========================================================
    # DEPARTMENT
    # ========================================================

    paragraph = document.add_paragraph()

    paragraph.alignment = (
        WD_ALIGN_PARAGRAPH.CENTER
    )

    run = paragraph.add_run(
        str(
            get_report_department(
                details
            )
        )
    )

    run.bold = True
    run.font.size = Pt(12)

    # ========================================================
    # REPORT TITLE
    # ========================================================

    paragraph = document.add_paragraph()

    paragraph.alignment = (
        WD_ALIGN_PARAGRAPH.CENTER
    )

    run = paragraph.add_run(
        "AI-Based Automatic CO-PO-PEO Mapping Report"
    )

    run.bold = True
    run.font.size = Pt(16)

    # ========================================================
    # 1. SUBJECT DETAILS
    # ========================================================

    add_word_section_heading(
        document,
        "1. Subject Details",
        level=1
    )

    subject_details = get_report_subject_details(
        details
    )

    add_word_table(
        document,
        [
            "Field",
            "Details"
        ],
        subject_details
    )

    # ========================================================
    # 2. COURSE OUTCOMES
    # ========================================================

    add_word_section_heading(
        document,
        "2. Course Outcomes and Bloom's Taxonomy",
        level=1
    )

    co_word_rows = []

    for row in co_rows:

        co_word_rows.append(
            [
                row.get(
                    "CO ID",
                    ""
                ),

                row.get(
                    "CO Statement",
                    ""
                ),

                row.get(
                    "Bloom Level",
                    ""
                ),

                row.get(
                    "Bloom Reason",
                    ""
                )
            ]
        )

    add_word_table(
        document,
        [
            "CO ID",
            "Course Outcome",
            "Bloom Level",
            "Bloom Reason"
        ],
        co_word_rows
    )

    # ========================================================
    # 3. PROGRAM OUTCOMES
    # ========================================================

    add_word_section_heading(
        document,
        "3. Program Outcomes (PO)",
        level=1
    )

    po_word_rows = []

    for row in po_rows:

        po_word_rows.append(
            [
                row.get(
                    "PO ID",
                    ""
                ),

                row.get(
                    "PO Statement",
                    ""
                )
            ]
        )

    add_word_table(
        document,
        [
            "PO ID",
            "PO Statement"
        ],
        po_word_rows
    )

    # ========================================================
    # 4. PROGRAM EDUCATIONAL OBJECTIVES
    # ========================================================

    add_word_section_heading(
        document,
        "4. Program Educational Objectives (PEO)",
        level=1
    )

    peo_word_rows = []

    for row in peo_rows:

        peo_word_rows.append(
            [
                row.get(
                    "PEO ID",
                    ""
                ),

                row.get(
                    "PEO Statement",
                    ""
                )
            ]
        )

    add_word_table(
        document,
        [
            "PEO ID",
            "PEO Statement"
        ],
        peo_word_rows
    )

    # ========================================================
    # 5. PROGRAM SPECIFIC OUTCOMES
    # ========================================================

    add_word_section_heading(document, "5. Program Specific Outcomes (PSO)", level=1)
    _w_id = find_column(pso_dataset_word, ["PSO ID", "PSO"]) if not pso_dataset_word.empty else None
    _w_text = find_column(
        pso_dataset_word,
        ["Programme Specific Outcomes", "Program Specific Outcomes", "PSO Description", "PSO Statement", "Description"]
    ) if not pso_dataset_word.empty else None
    _w_rows = []
    if _w_id and _w_text:
        for _, _r in pso_dataset_word.iterrows():
            if safe_string(_r.get(_w_id, "")):
                _w_rows.append([safe_string(_r.get(_w_id, "")), safe_string(_r.get(_w_text, ""))])
    add_word_table(document, ["PSO ID", "PSO Statement"], _w_rows)

    # ========================================================
    # 5. CO-PO MAPPING MATRIX
    # ========================================================

    add_word_section_heading(
        document,
        "6. CO-PO Mapping Matrix",
        level=1
    )

    po_columns = sort_mapping_columns(
        get_mapping_columns(
            co_po
        )
    )

    co_po_rows = []

    for row in report_co_po_matrix(
        co,
        co_po
    ):

        co_po_rows.append(
            [
                row.get(
                    "CO ID",
                    ""
                )
            ]
            +
            [
                row.get(
                    column,
                    0
                )
                for column in po_columns
            ]
        )

    add_word_table(
        document,
        [
            "CO ID"
        ]
        +
        po_columns,
        co_po_rows
    )

    # ========================================================
    # 6. CO-PEO MAPPING MATRIX
    # ========================================================

    add_word_section_heading(
        document,
        "7. CO-PEO Mapping Matrix",
        level=1
    )

    peo_columns = sort_mapping_columns(
        get_mapping_columns(
            co_peo
        )
    )

    co_peo_rows = []

    for row in report_co_peo_matrix(
        co,
        co_peo
    ):

        co_peo_rows.append(
            [
                row.get(
                    "CO ID",
                    ""
                )
            ]
            +
            [
                row.get(
                    column,
                    0
                )
                for column in peo_columns
            ]
        )

    add_word_table(
        document,
        [
            "CO ID"
        ]
        +
        peo_columns,
        co_peo_rows
    )

    # ========================================================
    # 7. CO-PSO MAPPING MATRIX
    # ========================================================

    add_word_section_heading(document, "8. CO-PSO Mapping Matrix", level=1)
    pso_id_col_word = find_column(pso_dataset_word, ["PSO ID", "PSO"]) if not pso_dataset_word.empty else None
    pso_columns_word = [safe_string(v) for v in pso_dataset_word[pso_id_col_word].tolist() if safe_string(v)] if pso_id_col_word else []
    pso_word_matrix_rows = []
    for _, row in pso_mapping_word.iterrows():
        pso_word_matrix_rows.append([safe_string(row.get("CO ID", ""))] + [safe_string(row.get(pid, 0)) for pid in pso_columns_word])
    add_word_table(document, ["CO ID"] + pso_columns_word, pso_word_matrix_rows)

    # ========================================================
    # 8. CO-PO JUSTIFICATION
    # ========================================================

    add_word_section_heading(
        document,
        "9. CO-PO Mapping Justification",
        level=1
    )

    po_word_reason_rows = []

    for row in po_reason_rows:

        po_word_reason_rows.append(
            [
                row.get(
                    "CO ID",
                    ""
                ),

                row.get(
                    "PO ID",
                    ""
                ),

                row.get(
                    "Mapping Value",
                    0
                ),

                row.get(
                    "Strength",
                    ""
                ),

                row.get(
                    "Reason",
                    ""
                )
            ]
        )

    add_word_table(
        document,
        [
            "CO ID",
            "PO ID",
            "Value",
            "Level",
            "Justification"
        ],
        po_word_reason_rows
    )

    # ========================================================
    # 8. CO-PEO JUSTIFICATION
    # ========================================================

    add_word_section_heading(
        document,
        "10. CO-PEO Mapping Justification",
        level=1
    )

    peo_word_reason_rows = []

    for row in peo_reason_rows:

        peo_word_reason_rows.append(
            [
                row.get(
                    "CO ID",
                    ""
                ),

                row.get(
                    "PEO ID",
                    ""
                ),

                row.get(
                    "Mapping Value",
                    0
                ),

                row.get(
                    "Strength",
                    ""
                ),

                row.get(
                    "Reason",
                    ""
                )
            ]
        )

    add_word_table(
        document,
        [
            "CO ID",
            "PEO ID",
            "Value",
            "Level",
            "Justification"
        ],
        peo_word_reason_rows
    )

    # ========================================================
    # 10. CO-PSO JUSTIFICATION
    # ========================================================

    add_word_section_heading(document, "11. CO-PSO Mapping Justification", level=1)
    pso_word_reason_rows = []
    for _, row in pso_just_word.iterrows():
        pso_word_reason_rows.append([safe_string(row.get("CO ID", "")), safe_string(row.get("PSO ID", "")), safe_string(row.get("Mapping Level", 0)), safe_string(row.get("Strength", "")), safe_string(row.get("Justification", ""))])
    add_word_table(document, ["CO ID", "PSO ID", "Value", "Level", "Justification"], pso_word_reason_rows)

    # ========================================================
    # 11. MAPPING SCALE
    # ========================================================

    add_word_section_heading(
        document,
        "12. Mapping Scale",
        level=1
    )

    scale_word_rows = []

    for item in report_mapping_scale():

        scale_word_rows.append(
            [
                item.get(
                    "Value",
                    ""
                ),

                item.get(
                    "Level",
                    ""
                ),

                item.get(
                    "Description",
                    ""
                )
            ]
        )

    add_word_table(
        document,
        [
            "Value",
            "Level",
            "Description"
        ],
        scale_word_rows
    )

    # ========================================================
    # 10. FACULTY VALIDATION
    # ========================================================

    add_word_section_heading(
        document,
        "13. Faculty Validation",
        level=1
    )

    document.add_paragraph(
        "Faculty Name: __________________________________________"
    )

    document.add_paragraph(
        "Faculty Signature: ______________________________________"
    )

    document.add_paragraph(
        "Date: ____________________"
    )

    document.add_paragraph(
        "HOD Signature: __________________________________________"
    )

    # ========================================================
    # SAVE WORD DOCUMENT
    # ========================================================

    try:

        document.save(
            filename
        )

    except Exception as error:

        print(
            "WORD SAVE ERROR:",
            error
        )

        return (
            safe_string(error),
            500
        )

    # ========================================================
    # DOWNLOAD
    # ========================================================

    return send_file(
        filename,
        as_attachment=True,
        download_name=(
            f"{subject_code}_Subject_Report.docx"
        ),
        mimetype=(
            "application/vnd.openxmlformats-officedocument."
            "wordprocessingml.document"
        )
    )

# ============================================================
# DOWNLOAD EXCEL
# ============================================================


def add_word_section_heading(document, text, *args, **kwargs):
    """Safe Word heading that does not require Heading 1 style."""
    paragraph = document.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

    run = paragraph.add_run(str(text))
    run.bold = True
    run.font.size = Pt(14)

    paragraph.paragraph_format.space_before = Pt(8)
    paragraph.paragraph_format.space_after = Pt(5)
    paragraph.paragraph_format.keep_with_next = True

    return paragraph

@app.route(
    "/download_excel"
)
def download_excel():

    if "user" not in session:

        return redirect("/")

    subject_code = (
        get_download_subject()
    )

    if not subject_code:

        return redirect(
            "/dashboard"
        )

    try:

        # ----------------------------------------------------
        # LOAD DATA
        # ----------------------------------------------------

        details = get_subject_details(
            subject_code
        )

        co = load_course_outcomes(
            subject_code
        )

        co_po = load_co_po(
            subject_code
        )

        co_peo = load_co_peo(
            subject_code
        )

        po_justification = (
            load_justification(
                CO_PO_JUSTIFICATION_FILE,
                subject_code
            )
        )

        peo_justification = (
            load_justification(
                CO_PEO_JUSTIFICATION_FILE,
                subject_code
            )
        )

        # Include the manually maintained PSO dataset and subject-specific
        # mapping/justification in the downloadable subject report.
        from pso_mapping import generate_pso_mapping
        pso_result = generate_pso_mapping(BASE_DIR, DATASET_DIR)
        pso_dataset_excel = pd.read_excel(os.path.join(DATASET_DIR, "PSO_Dataset.xlsx")).fillna("")
        pso_mapping_all = pso_result["mapping"].fillna("")
        pso_just_all = pso_result["justifications"].fillna("")
        if "Subject Code" in pso_mapping_all.columns:
            pso_mapping_excel = pso_mapping_all[pso_mapping_all["Subject Code"].astype(str).str.upper().str.strip() == subject_code].copy()
        else:
            pso_mapping_excel = pd.DataFrame()
        if "Subject Code" in pso_just_all.columns:
            pso_just_excel = pso_just_all[pso_just_all["Subject Code"].astype(str).str.upper().str.strip() == subject_code].copy()
        else:
            pso_just_excel = pd.DataFrame()

    except Exception as error:

        print(
            "EXCEL REPORT ERROR:",
            error
        )

        return (
            safe_string(error),
            500
        )

    # --------------------------------------------------------
    # REPORT DATA
    # --------------------------------------------------------

    co_rows = build_co_rows(
        co
    )

    po_rows = get_report_po_details(
        co_po
    )

    peo_rows = get_report_peo_details(
        co_peo
    )

    po_reason_rows = (
        build_po_reason_rows(
            co,
            co_po,
            po_justification
        )
    )

    peo_reason_rows = (
        build_peo_reason_rows(
            co,
            co_peo,
            peo_justification
        )
    )

    bloom_rows = (
        build_bloom_reason_rows(
            co
        )
    )

    # --------------------------------------------------------
    # FILE
    # --------------------------------------------------------

    filename = os.path.join(
        BASE_DIR,
        f"{subject_code}_Subject_Report.xlsx"
    )

    # --------------------------------------------------------
    # EXCEL
    # --------------------------------------------------------

    with pd.ExcelWriter(

        filename,

        engine="openpyxl"
    ) as writer:

        # ====================================================
        # 1. SUBJECT DETAILS
        # ====================================================

        subject_data = pd.DataFrame(

            get_report_subject_details(
                details
            ),

            columns=[
                "Field",
                "Details"
            ]
        )

        subject_data.to_excel(

            writer,

            sheet_name="Subject Details",

            index=False
        )

        # ====================================================
        # 2. COURSE OUTCOMES
        # ====================================================

        if co_rows:

            co_excel = pd.DataFrame(
                co_rows
            )

        else:

            co_excel = pd.DataFrame(

                columns=[
                    "CO ID",
                    "CO Statement",
                    "Bloom Level",
                    "Bloom Reason"
                ]
            )

        co_excel.to_excel(

            writer,

            sheet_name="Course Outcomes",

            index=False
        )

        # ====================================================
        # 3. PO DETAILS
        # ====================================================

        po_excel = pd.DataFrame(
            po_rows,

            columns=[
                "PO ID",
                "PO Statement"
            ]
        )

        po_excel.to_excel(

            writer,

            sheet_name="PO Details",

            index=False
        )

        # ====================================================
        # 4. PEO DETAILS
        # ====================================================

        peo_excel = pd.DataFrame(
            peo_rows,

            columns=[
                "PEO ID",
                "PEO Statement"
            ]
        )

        peo_excel.to_excel(

            writer,

            sheet_name="PEO Details",

            index=False
        )

        # ====================================================
        # 5. PSO DETAILS (manually maintained in Excel)
        # ====================================================
        pso_dataset_excel.to_excel(writer, sheet_name="PSO Details", index=False)

        # ====================================================
        # 6. CO-PSO MAPPING
        # ====================================================
        pso_mapping_excel.to_excel(writer, sheet_name="CO-PSO Mapping", index=False)
        pso_just_excel.to_excel(writer, sheet_name="CO-PSO Justification", index=False)

        # ====================================================
        # 7. CO-PO MAPPING
        # ====================================================

        if co_po is not None:

            co_po.to_excel(

                writer,

                sheet_name="CO-PO Mapping",

                index=False
            )

        else:

            pd.DataFrame().to_excel(

                writer,

                sheet_name="CO-PO Mapping",

                index=False
            )

        # ====================================================
        # 6. CO-PEO MAPPING
        # ====================================================

        if co_peo is not None:

            co_peo.to_excel(

                writer,

                sheet_name="CO-PEO Mapping",

                index=False
            )

        else:

            pd.DataFrame().to_excel(

                writer,

                sheet_name="CO-PEO Mapping",

                index=False
            )

        # ====================================================
        # 7. CO-PO JUSTIFICATION
        # ====================================================

        po_reason_excel = pd.DataFrame(

            po_reason_rows,

            columns=[
                "CO ID",
                "PO ID",
                "Mapping Value",
                "Strength",
                "CO Statement",
                "Reason"
            ]
        )

        po_reason_excel.to_excel(

            writer,

            sheet_name="CO-PO Justification",

            index=False
        )

        # ====================================================
        # 8. CO-PEO JUSTIFICATION
        # ====================================================

        peo_reason_excel = pd.DataFrame(

            peo_reason_rows,

            columns=[
                "CO ID",
                "PEO ID",
                "Mapping Value",
                "Strength",
                "CO Statement",
                "Reason"
            ]
        )

        peo_reason_excel.to_excel(

            writer,

            sheet_name="CO-PEO Justification",

            index=False
        )

        # ====================================================
        # 9. BLOOM ANALYSIS
        # ====================================================

        bloom_excel = pd.DataFrame(

            bloom_rows,

            columns=[
                "CO ID",
                "Bloom Level",
                "Explanation"
            ]
        )

        bloom_excel.to_excel(

            writer,

            sheet_name="Bloom Analysis",

            index=False
        )

        # ====================================================
        # 10. MAPPING SCALE
        # ====================================================

        scale_excel = pd.DataFrame(

            report_mapping_scale(),

            columns=[
                "Value",
                "Level",
                "Description"
            ]
        )

        scale_excel.to_excel(

            writer,

            sheet_name="Mapping Scale",

            index=False
        )

    # ========================================================
    # EXCEL FORMATTING
    # ========================================================

    try:

        workbook = load_workbook(
            filename
        )

        # ----------------------------------------------------
        # STYLE ALL SHEETS
        # ----------------------------------------------------

        for worksheet in workbook.worksheets:

            style_excel_sheet(
                worksheet
            )

        # ----------------------------------------------------
        # ADD LOGO TO SUBJECT DETAILS
        # ----------------------------------------------------

        logo = get_report_logo()

        if logo:

            try:

                excel_logo = ExcelImage(
                    logo
                )

                excel_logo.width = 100
                excel_logo.height = 100

                cover = workbook[
                    "Subject Details"
                ]

                cover.add_image(
                    excel_logo,
                    "D2"
                )

            except Exception as error:

                print(
                    "EXCEL LOGO ERROR:",
                    error
                )

        # ----------------------------------------------------
        # SAVE FORMATTED FILE
        # ----------------------------------------------------

        workbook.save(
            filename
        )

    except Exception as error:

        print(
            "EXCEL FORMATTING ERROR:",
            error
        )

    return send_file(

        filename,

        as_attachment=True,

        download_name=(
            f"{subject_code}_Subject_Report.xlsx"
        ),

        mimetype=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        )
    )

# ============================================================
# DEBUG ROUTES
# ============================================================

print()
print("=" * 70)
print("RUNNING THIS APP.PY")
print("APP FILE:", os.path.abspath(__file__))
print()
print("DATASET DIRECTORY:", DATASET_DIR)
print("SUBJECT FILE:", SUBJECT_FILE)
print("CO FILE:", CO_FILE)
print("CO-PO FILE:", CO_PO_FILE)
print("CO-PEO FILE:", CO_PEO_FILE)
print()
print("ROUTES:")

for rule in app.url_map.iter_rules():

    print(
        rule
    )

print("=" * 70)
print()


# ============================================================
# RUN APPLICATION
# ============================================================



# ============================================================
# CO–PSO MAPPING (PSO dataset is maintained manually in Excel)
# ============================================================

@app.route("/pso-mapping", methods=["GET", "POST"])
def pso_mapping_page():
    if "user" not in session:
        return redirect("/")
    from pso_mapping import generate_pso_mapping
    try:
        result = generate_pso_mapping(BASE_DIR, DATASET_DIR)
        mapping = result["mapping"].fillna("").to_dict(orient="records")
        justifications = result["justifications"].fillna("").to_dict(orient="records")
        pso_df = pd.read_excel(os.path.join(DATASET_DIR, "PSO_Dataset.xlsx")).fillna("")
        pso_records = pso_df.to_dict(orient="records")
        pso_ids = [str(c) for c in mapping[0].keys() if str(c).startswith("PSO")] if mapping else []
        graph_labels = [str(x) for x in pso_df.iloc[:, 0].tolist()]
        graph_values = [int((result["mapping"][pid] > 0).sum()) for pid in graph_labels if pid in result["mapping"].columns]
        return render_template(
            "pso_mapping.html",
            mapping=mapping,
            justifications=justifications,
            pso_records=pso_records,
            pso_ids=pso_ids,
            graph_labels=graph_labels,
            graph_values=graph_values,
            model_name=result["model_name"],
            report_ready=os.path.exists(result["report_path"]),
            error=None
        )
    except Exception as exc:
        return render_template(
            "pso_mapping.html", mapping=[], justifications=[], pso_records=[],
            pso_ids=[], graph_labels=[], graph_values=[], model_name="",
            report_ready=False, error=str(exc)
        )

@app.route("/download-pso-report")
def download_pso_report():
    if "user" not in session:
        return redirect("/")
    report_path = os.path.join(DATASET_DIR, "CO_PSO_Mapping_Report.xlsx")
    if not os.path.exists(report_path):
        from pso_mapping import generate_pso_mapping
        generate_pso_mapping(BASE_DIR, DATASET_DIR)
    return send_file(report_path, as_attachment=True, download_name="CO_PSO_Mapping_Report.xlsx")


# ============================================================
# DEPARTMENTS (each department keeps its own CO / PO / PEO / PSO data)
# ============================================================

@app.route("/departments", methods=["GET", "POST"])
def departments_page():
    if "user" not in session:
        return redirect("/")
    if request.method == "POST":
        chosen = request.form.get("department", "")
        if any(d["slug"] == chosen for d in dept_store.list_departments(BASE_DIR)):
            session["department"] = chosen
            session.pop("subject_code", None)
            return redirect("/dashboard")
    return render_template(
        "departments.html",
        departments=dept_store.list_departments(BASE_DIR),
        current=session.get("department"),
        error=request.args.get("error"),
        success=request.args.get("success")
    )


@app.route("/departments/upload", methods=["POST"])
def departments_upload():
    if "user" not in session:
        return redirect("/")
    from urllib.parse import quote
    import input_parsers
    dept_name = request.form.get("name", "").strip()
    uploads, problems = {}, []
    for key in dept_store.FILES:
        # Typed text wins if present; otherwise the uploaded file (.xlsx/.xls/.csv/.pdf) is read
        typed = request.form.get(f"text_{key}", "")
        try:
            uploads[key] = input_parsers.load_dataset(
                key,
                file_storage=request.files.get(key),
                typed_text=typed,
                department=dept_name,
                default_codes=(
                    [str(c).strip().upper() for c in uploads["subjects"].iloc[:, 0]]
                    if "subjects" in uploads else []
                ),
            )
        except input_parsers.ParseError as exc:
            problems.append(str(exc))
    if problems:
        return redirect("/departments?error=" + quote(" || ".join(problems)))
    slug, errors = dept_store.create_department(
        BASE_DIR,
        dept_name,
        uploads,
        replace=bool(request.form.get("replace"))
    )
    if errors:
        return redirect("/departments?error=" + quote(" || ".join(errors)))
    session["department"] = slug
    session.pop("subject_code", None)
    return redirect("/dashboard")


@app.route("/favicon.ico")
def favicon():
    return redirect("/static/images/ATME.jpeg")


# ============================================================
# RUN APPLICATION (Windows-safe development server)
# ============================================================

def _port_is_free(host, port):
    """True only if nothing at all is listening on host:port.

    Werkzeug's server sets SO_REUSEADDR.  On Windows that flag lets a second
    process bind a port that an older (or crashed/orphaned) Flask process still
    owns, and the OS then hands incoming browser connections to either process
    unpredictably: some requests get the page, some get "Not Found" (old
    version) and some get ERR_CONNECTION_RESET (dead process).  curl can hit the
    healthy one while the browser hits the stale one.  So we refuse to start on
    a port that is already taken and pick a free one instead.
    """
    import socket

    probe = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    probe.settimeout(0.5)
    try:
        if probe.connect_ex((host, port)) == 0:
            return False
    finally:
        probe.close()

    test = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        if hasattr(socket, "SO_EXCLUSIVEADDRUSE"):      # Windows
            test.setsockopt(socket.SOL_SOCKET, socket.SO_EXCLUSIVEADDRUSE, 1)
        test.bind((host, port))
        return True
    except OSError:
        return False
    finally:
        test.close()


def _choose_port(host, preferred):
    for port in range(preferred, preferred + 20):
        if _port_is_free(host, port):
            return port
    raise RuntimeError(
        f"No free port found between {preferred} and {preferred + 19}. "
        "Close other Python/Flask windows and try again."
    )


if __name__ == "__main__":

    HOST = os.environ.get("HOST", "127.0.0.1")
    PREFERRED_PORT = int(os.environ.get("PORT", "5000"))
    PORT = _choose_port(HOST, PREFERRED_PORT)

    # The Werkzeug debugger is optional (set FLASK_DEBUG=1).  The auto-reloader
    # is always OFF: with sentence-transformers installed it watches thousands
    # of files in site-packages (transformers, torch ...) and restarts the
    # server over and over, dropping browser connections.
    DEBUG = os.environ.get("FLASK_DEBUG", "0") == "1"

    print()
    print("=" * 70)
    if PORT != PREFERRED_PORT:
        print(f"Port {PREFERRED_PORT} is already in use by another program/old Flask process.")
        print(f"Using port {PORT} instead.")
    # Warm up the AI model in the background so the first upload/generate is not slow.
    try:
        import pso_mapping
        pso_mapping.preload_model()
        print("Loading the AI model in the background (first upload will be faster)...")
    except Exception as _exc:
        print("Model preload skipped:", _exc)
    print(f"Open this address in Chrome or Edge:  http://{HOST}:{PORT}/")
    print("(use http:// not https://)")
    print("=" * 70)
    print()

    app.run(
        host=HOST,
        port=PORT,
        debug=DEBUG,
        use_reloader=False,
        threaded=True
    )
