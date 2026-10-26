"""Department-wise dataset storage, validation and CO-PO / CO-PEO / CO-PSO mapping generation.

The original CSE data stays in  Dataset/.  Every department uploaded through the website is stored in
Departments/<slug>/  using the same five Excel files and the same generated CSV files as Dataset/.
"""
import json
import os
import re
import shutil
import tempfile
from pathlib import Path

import pandas as pd

from pso_mapping import _col, _similarities, generate_pso_mapping

DEFAULT_SLUG = "CSE"
DEFAULT_NAME = "Computer Science and Engineering"

# Standard file names used by the whole application.
FILES = {
    "subjects": "Subject_Details.xlsx",
    "co": "CO_Dataset-1.xlsx",
    "po": "PO_Dataset.xlsx",
    "peo": "PEO_Dataset.xlsx",
    "pso": "PSO_Dataset.xlsx",
}


def departments_root(base_dir):
    # Set DEPARTMENTS_DIR (e.g. /var/data/Departments) to keep uploaded departments on a persistent disk.
    custom = os.environ.get("DEPARTMENTS_DIR")
    return Path(custom) if custom else Path(base_dir) / "Departments"


def make_slug(name):
    slug = re.sub(r"[^A-Za-z0-9]+", "_", str(name)).strip("_")[:40]
    return slug


def dataset_dir_for(base_dir, slug):
    """Dataset folder for a department slug (falls back to the default CSE folder)."""
    base = Path(base_dir)
    if slug and slug != DEFAULT_SLUG and re.fullmatch(r"[A-Za-z0-9_]+", slug):
        path = departments_root(base) / slug
        if (path / FILES["subjects"]).exists():
            return path
    return base / "Dataset"


def _meta_name(path, fallback):
    meta = Path(path) / "meta.json"
    try:
        return json.loads(meta.read_text(encoding="utf-8")).get("name", fallback)
    except Exception:
        return fallback


def list_departments(base_dir):
    base = Path(base_dir)
    result = []
    default_dir = base / "Dataset"
    if (default_dir / FILES["subjects"]).exists():
        name = DEFAULT_NAME
        try:
            df = pd.read_excel(default_dir / FILES["subjects"])
            column = _col(df, ["Department"])
            if column and df[column].dropna().astype(str).str.strip().any():
                name = str(df[column].dropna().astype(str).str.strip().iloc[0])
        except Exception:
            pass
        result.append({"slug": DEFAULT_SLUG, "name": name, "default": True})
    root = departments_root(base)
    if root.exists():
        for folder in sorted(root.iterdir()):
            if folder.is_dir() and (folder / FILES["subjects"]).exists():
                result.append({"slug": folder.name, "name": _meta_name(folder, folder.name), "default": False})
    return result


def department_name(base_dir, slug):
    for item in list_departments(base_dir):
        if item["slug"] == slug:
            return item["name"]
    return DEFAULT_NAME


# ------------------------------------------------------------------ validation
def _read(path_or_file):
    return pd.read_excel(path_or_file, sheet_name=0).fillna("")


def validate_uploads(frames):
    """frames: dict key -> DataFrame.  Returns (errors, info)."""
    errors, info = [], {}

    sub = frames["subjects"]
    c_code, c_name = _col(sub, ["Subject Code", "Course Code"]), _col(sub, ["Subject Name", "Course Name"])
    if not c_code or not c_name:
        errors.append("Subjects file needs the columns: Subject Code, Subject Name (Semester, Credits, Schema, Department are optional).")
    elif not sub[c_code].astype(str).str.strip().ne("").any():
        errors.append("Subjects file has no subject codes.")

    co = frames["co"]
    k_code, k_id, k_text = (_col(co, ["Subject Code", "Course Code"]), _col(co, ["CO ID", "Course Outcome ID"]),
                            _col(co, ["Course Outcome", "CO Statement", "Description"]))
    if not (k_code and k_id and k_text):
        errors.append("CO file needs the columns: Subject Code, CO ID, Course Outcome.")
    elif not co[k_text].astype(str).str.strip().ne("").any():
        errors.append("CO file has no course outcome statements.")

    for key, label, prefix, text_names in (
        ("po", "PO", "PO", ["Programme Outcomes", "Program Outcomes", "PO Description", "PO Statement", "Description"]),
        ("peo", "PEO", "PEO", ["Programme Educational Objectives", "Program Educational Objectives", "PEO Description", "PEO Statement", "Description"]),
        ("pso", "PSO", "PSO", ["Programme Specific Outcomes", "Program Specific Outcomes", "PSO Description", "Description"]),
    ):
        df = frames[key]
        c_id, c_text = _col(df, [f"{label} ID", label]), _col(df, text_names)
        if not c_id or not c_text:
            errors.append(f"{label} file needs two columns: '{label} ID' and the {label} description.")
            continue
        df = df[df[c_text].astype(str).str.strip().ne("")]
        if df.empty:
            errors.append(f"{label} file has no {label} statements.")
            continue
        bad = [str(v) for v in df[c_id] if not re.fullmatch(prefix + r"\s*\d+", str(v).strip(), re.I)]
        if bad:
            errors.append(f"{label} IDs must look like {prefix}1, {prefix}2 ... (found: {', '.join(bad[:3])}).")
    return errors, info


def _write_frame(df, path):
    df.to_excel(path, index=False)


# ---------------------------------------------------------------- generation
def _top3(scores):
    out = []
    for row in scores:
        order = sorted(range(len(row)), key=lambda j: float(row[j]), reverse=True)
        levels = [0] * len(row)
        for rank, j in enumerate(order[:3]):
            levels[j] = 3 - rank
        out.append(levels)
    return out


def _generic_text(strength, left, right):
    return {3: f"High semantic similarity between {left} and {right}",
            2: f"Medium semantic similarity between {left} and {right}",
            1: f"Low semantic similarity between {left} and {right}"}.get(strength, "No significant relationship")


def build_mapping_files(dataset_dir):
    """Create co_po_mapping.csv / co_peo_mapping.csv (+ justification CSVs) from the five Excel files."""
    d = Path(dataset_dir)
    co = _read(d / FILES["co"])
    co.columns = [str(c).strip() for c in co.columns]
    c_code, c_id, c_text = (_col(co, ["Subject Code", "Course Code"]), _col(co, ["CO ID", "Course Outcome ID"]),
                            _col(co, ["Course Outcome", "CO Statement", "Description"]))
    co = co[co[c_text].astype(str).str.strip().ne("")].copy()
    unique = (co[c_code].astype(str).str.strip().str.upper() + "_" + co[c_id].astype(str).str.strip().str.upper()).tolist()
    texts = co[c_text].astype(str).tolist()

    model_name = ""
    for key, label, names, stem in (
        ("po", "PO", ["Programme Outcomes", "Program Outcomes", "PO Description", "PO Statement", "Description"], "co_po"),
        ("peo", "PEO", ["Programme Educational Objectives", "Program Educational Objectives", "PEO Description", "PEO Statement", "Description"], "co_peo"),
    ):
        df = _read(d / FILES[key])
        df.columns = [str(c).strip() for c in df.columns]
        c_tid, c_ttext = _col(df, [f"{label} ID", label]), _col(df, names)
        df = df[df[c_ttext].astype(str).str.strip().ne("")]
        ids = [str(v).strip() for v in df[c_tid]]
        scores, model_name = _similarities(texts, df[c_ttext].astype(str).tolist())
        levels = _top3(scores)
        mapping = pd.DataFrame(levels, columns=ids)
        mapping.insert(0, "Unique_CO", unique)
        mapping.to_csv(d / f"{stem}_mapping.csv", index=False, encoding="utf-8")
        sim = pd.DataFrame(scores, columns=ids)
        sim.insert(0, "Unique_CO", unique)
        sim.to_csv(d / f"{stem}_similarity.csv", index=False, encoding="utf-8")
        rows = [[u, tid, lv, _generic_text(lv, u, tid)]
                for u, row in zip(unique, levels) for tid, lv in zip(ids, row)]
        pd.DataFrame(rows, columns=["CO ID", f"{label} ID", "Mapping Value", "Justification"]).to_csv(
            d / f"{stem}_justification.csv", index=False, encoding="utf-8")
    return model_name


def create_department(base_dir, name, uploads, replace=False):
    """uploads: dict key -> file-like object.  Returns (slug, errors)."""
    name = str(name).strip()
    slug = make_slug(name)
    if not slug:
        return None, ["Please enter a department name (letters and numbers)."]
    if slug.upper() == DEFAULT_SLUG:
        return None, ["The name 'CSE' is reserved for the built-in Computer Science dataset. Use a different name."]
    root = departments_root(base_dir)
    target = root / slug
    if target.exists() and not replace:
        return None, [f"A department named '{name}' already exists. Tick 'Replace existing data' to overwrite it."]

    try:
        # values may be ready-made DataFrames (PDF / typed text / .xls) or Excel file objects
        frames = {key: (uploads[key].fillna("") if isinstance(uploads[key], pd.DataFrame) else _read(uploads[key]))
                  for key in FILES}
    except Exception as exc:
        return None, [f"Could not read an uploaded file: {exc}"]
    for df in frames.values():
        df.columns = [str(c).strip() for c in df.columns]
    errors, _ = validate_uploads(frames)
    if errors:
        return None, errors

    root.mkdir(parents=True, exist_ok=True)
    tmp = Path(tempfile.mkdtemp(prefix="_new_", dir=root))
    try:
        for key, filename in FILES.items():
            _write_frame(frames[key], tmp / filename)
        sub = frames["subjects"]
        if not _col(sub, ["Department"]):
            sub["Department"] = name
            _write_frame(sub, tmp / FILES["subjects"])
        (tmp / "meta.json").write_text(json.dumps({"name": name}, ensure_ascii=False), encoding="utf-8")
        build_mapping_files(tmp)
        generate_pso_mapping(base_dir, tmp)
        if target.exists():
            shutil.rmtree(target)
        os.replace(tmp, target)
    except Exception as exc:
        shutil.rmtree(tmp, ignore_errors=True)
        return None, [f"Could not generate the mappings: {exc}"]
    return slug, []
