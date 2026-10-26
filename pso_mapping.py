"""CO–PSO mapping generation and Excel report utilities."""
from pathlib import Path
import re
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.chart import BarChart, Reference

PSO_DEFAULTS = [
    ("PSO1", "Ability to apply skills in the field of algorithms, database design, web design, cloud computing and data analytics."),
    ("PSO2", "Apply knowledge in the field of computer networks for building network and internet based applications."),
]

def _col(df, candidates):
    norm = {re.sub(r"[^a-z0-9]", "", str(c).lower()): c for c in df.columns}
    for candidate in candidates:
        key = re.sub(r"[^a-z0-9]", "", candidate.lower())
        if key in norm:
            return norm[key]
    for candidate in candidates:
        key = re.sub(r"[^a-z0-9]", "", candidate.lower())
        for nk, original in norm.items():
            if key and (key in nk or nk in key):
                return original
    return None

import os
import threading

# If the model is not cached locally and the PC is offline, fail fast instead of hanging for minutes.
os.environ.setdefault("HF_HUB_ETAG_TIMEOUT", "8")
os.environ.setdefault("HF_HUB_DOWNLOAD_TIMEOUT", "30")

_LOCK = threading.RLock()
_MODEL = None
_MODEL_FAILED = False
_MODEL_LOCK = threading.Lock()
_RESULT_CACHE = {"key": None, "result": None}


def _get_model():
    """Load Sentence-BERT once per process (it used to be reloaded on every request)."""
    global _MODEL, _MODEL_FAILED
    with _MODEL_LOCK:
        return _load_model_locked()


def _load_model_locked():
    global _MODEL, _MODEL_FAILED
    if _MODEL is not None or _MODEL_FAILED:
        return _MODEL
    try:
        from sentence_transformers import SentenceTransformer
        try:
            # Use the local Hugging Face cache first so an offline PC does not wait on network timeouts.
            _MODEL = SentenceTransformer("all-MiniLM-L6-v2", local_files_only=True)
        except Exception:
            _MODEL = SentenceTransformer("all-MiniLM-L6-v2")
    except Exception as exc:
        print("Sentence-BERT unavailable, using TF-IDF fallback:", exc)
        _MODEL_FAILED = True
        _MODEL = None
    return _MODEL


_EMB_CACHE = {}


def _encode(model, texts):
    """Encode only texts not seen before (the CO statements are reused for PO, PEO and PSO)."""
    import numpy as np
    texts = [str(t) for t in texts]
    missing = list(dict.fromkeys(t for t in texts if t not in _EMB_CACHE))
    if missing:
        vectors = model.encode(missing, convert_to_numpy=True, show_progress_bar=False, batch_size=64)
        for t, v in zip(missing, vectors):
            _EMB_CACHE[t] = v
        if len(_EMB_CACHE) > 20000:          # keep memory bounded
            for k in list(_EMB_CACHE)[:5000]:
                _EMB_CACHE.pop(k, None)
    return np.vstack([_EMB_CACHE[t] for t in texts])


def preload_model():
    """Load Sentence-BERT in the background when the server starts, so the first upload is fast."""
    threading.Thread(target=_get_model, daemon=True).start()


def _similarities(co_texts, pso_texts):
    # Prefer Sentence-BERT, with a lightweight TF-IDF fallback when model files/network are unavailable.
    try:
        from sklearn.metrics.pairwise import cosine_similarity
        model = _get_model()
        if model is None:
            raise RuntimeError("Sentence-BERT model not available")
        a = _encode(model, co_texts)
        b = _encode(model, pso_texts)
        return cosine_similarity(a, b), "Sentence-BERT (all-MiniLM-L6-v2)"
    except Exception:
        from sklearn.feature_extraction.text import TfidfVectorizer
        from sklearn.metrics.pairwise import cosine_similarity
        corpus = list(co_texts) + list(pso_texts)
        vec = TfidfVectorizer(stop_words="english", ngram_range=(1,2))
        matrix = vec.fit_transform(corpus)
        return cosine_similarity(matrix[:len(co_texts)], matrix[len(co_texts):]), "TF-IDF cosine similarity (fallback)"

def generate_pso_mapping(base_dir, dataset_dir=None):
    """Return the CO-PSO mapping.  The result is cached until the CO or PSO dataset file changes,
    so page loads no longer re-encode every CO and rewrite the output files each time."""
    base = Path(base_dir)
    dataset = Path(dataset_dir) if dataset_dir is not None else base / "Dataset"
    co_file = dataset / "CO_Dataset-1.xlsx"
    pso_file = dataset / "PSO_Dataset.xlsx"
    key = tuple((str(f), f.stat().st_mtime_ns, f.stat().st_size) for f in (co_file, pso_file))
    with _LOCK:
        cached = _RESULT_CACHE
        if cached["key"] == key and cached["result"] is not None:
            res = cached["result"]
            return {"mapping": res["mapping"].copy(), "justifications": res["justifications"].copy(),
                    "report_path": res["report_path"], "model_name": res["model_name"]}
        result = _generate_pso_mapping(base_dir, dataset)
        _RESULT_CACHE["key"] = key
        _RESULT_CACHE["result"] = result
        return {"mapping": result["mapping"].copy(), "justifications": result["justifications"].copy(),
                "report_path": result["report_path"], "model_name": result["model_name"]}


def _generate_pso_mapping(base_dir, dataset_dir):
    base = Path(base_dir)
    dataset = Path(dataset_dir)
    co_file = dataset / "CO_Dataset-1.xlsx"
    pso_file = dataset / "PSO_Dataset.xlsx"
    co = pd.read_excel(co_file, sheet_name=0).fillna("")
    pso = pd.read_excel(pso_file, sheet_name=0).fillna("")
    co.columns = [str(c).strip() for c in co.columns]
    pso.columns = [str(c).strip() for c in pso.columns]
    co_id = _col(co, ["CO ID", "Course Outcome ID", "Outcome ID"])
    co_text = _col(co, ["Course Outcome", "CO Statement", "Outcome Description", "Description"])
    subject = _col(co, ["Subject Code", "Course Code"])
    pso_id = _col(pso, ["PSO ID", "PSO"])
    pso_text = _col(pso, ["Programme Specific Outcomes", "Program Specific Outcomes", "PSO Description", "Description"])
    if not co_id or not co_text:
        raise ValueError("CO_Dataset-1.xlsx must contain CO ID and Course Outcome columns.")
    if not pso_id or not pso_text:
        raise ValueError("PSO_Dataset.xlsx must contain PSO ID and Programme Specific Outcomes columns.")
    co = co[co[co_text].astype(str).str.strip().ne("")].copy()
    pso = pso[pso[pso_text].astype(str).str.strip().ne("")].copy()
    if co.empty or pso.empty:
        raise ValueError("Both the CO dataset and PSO dataset must contain outcome statements.")
    co_statements = co[co_text].astype(str).tolist()
    pso_statements = pso[pso_text].astype(str).tolist()
    scores, model_name = _similarities(co_statements, pso_statements)
    ids = pso[pso_id].astype(str).tolist()
    map_rows, just_rows = [], []
    for i, (_, row) in enumerate(co.iterrows()):
        cid = str(row[co_id]).strip()
        scode = str(row[subject]).strip() if subject else ""
        # Relative ranking per CO: strongest is 3, next 2, next 1; low scores remain 0.
        ranking = sorted(range(len(ids)), key=lambda j: float(scores[i][j]), reverse=True)
        strengths = {}
        for rank, j in enumerate(ranking[:min(3, len(ranking))]):
            strengths[j] = 3-rank
        mrow = {"Subject Code": scode, "CO ID": cid, "Course Outcome": str(row[co_text])}
        for j, pid in enumerate(ids):
            score = float(scores[i][j])
            level = strengths.get(j, 0)
            mrow[pid] = level
            strength_name = {3:"High",2:"Moderate",1:"Low",0:"None"}[level]
            if level == 3:
                reason = f"{pid} is the strongest semantic match for {cid}; the CO statement is most closely related to the stated PSO (similarity {score:.3f})."
            elif level == 2:
                reason = f"{pid} has a secondary semantic relationship with {cid}, so it receives a moderate mapping level (similarity {score:.3f})."
            elif level == 1:
                reason = f"{pid} has a weaker relative relationship with {cid} than the higher-ranked PSOs and receives a low mapping level (similarity {score:.3f})."
            else:
                reason = f"{pid} is outside the top three relative matches for {cid}; it is marked 0 (similarity {score:.3f}). Review manually if local accreditation rules require otherwise."
            just_rows.append({"Subject Code":scode,"CO ID":cid,"PSO ID":pid,"Mapping Level":level,"Strength":strength_name,"Similarity Score":round(score,4),"Justification":reason})
        map_rows.append(mrow)
    mapping = pd.DataFrame(map_rows)
    justifications = pd.DataFrame(just_rows)
    similarity = pd.DataFrame(scores, columns=ids)
    similarity.insert(0,"CO ID",co[co_id].astype(str).values)
    try:
        mapping.to_csv(dataset/"co_pso_mapping.csv", index=False)
        justifications.to_csv(dataset/"co_pso_justification.csv", index=False)
        similarity.to_csv(dataset/"co_pso_similarity.csv", index=False)
        report = dataset/"CO_PSO_Mapping_Report.xlsx"
        with pd.ExcelWriter(report, engine="openpyxl") as writer:
            mapping.to_excel(writer,index=False,sheet_name="CO-PSO Mapping")
            justifications.to_excel(writer,index=False,sheet_name="Justifications")
            pso.to_excel(writer,index=False,sheet_name="PSO Dataset")
            similarity.to_excel(writer,index=False,sheet_name="Similarity Scores")
            summary = pd.DataFrame([{"PSO ID":pid,"Mapped CO count":int((mapping[pid]>0).sum()),"High (3)":int((mapping[pid]==3).sum()),"Moderate (2)":int((mapping[pid]==2).sum()),"Low (1)":int((mapping[pid]==1).sum()),"None (0)":int((mapping[pid]==0).sum())} for pid in ids])
            summary.to_excel(writer,index=False,sheet_name="Graph Summary")
        book = load_workbook(report)
        for ws in book.worksheets:
            ws.freeze_panes="A2"
            ws.auto_filter.ref=ws.dimensions
            for cell in ws[1]:
                cell.font=Font(bold=True,color="FFFFFF")
                cell.fill=PatternFill("solid",fgColor="244062")
                cell.alignment=Alignment(wrap_text=True,vertical="top")
            for col in ws.columns:
                letter=col[0].column_letter
                max_len=max((len(str(c.value or "")) for c in col[:40]),default=10)
                ws.column_dimensions[letter].width=min(max(max_len+2,12),55)
            for row in ws.iter_rows(min_row=2):
                for cell in row: cell.alignment=Alignment(wrap_text=True,vertical="top")
        ws=book["Graph Summary"]
        chart=BarChart()
        chart.type="col"
        chart.style=10
        chart.title="CO–PSO Mapping Strength Summary"
        chart.y_axis.title="Number of COs"
        chart.x_axis.title="PSO"
        chart.height=8
        chart.width=16
        data=Reference(ws,min_col=3,max_col=5,min_row=1,max_row=ws.max_row)
        cats=Reference(ws,min_col=1,min_row=2,max_row=ws.max_row)
        chart.add_data(data,titles_from_data=True)
        chart.set_categories(cats)
        ws.add_chart(chart,"G2")
        book.save(report)
    except OSError as exc:
        # e.g. CO_PSO_Mapping_Report.xlsx is open in Excel on Windows.  The in-memory results are still valid.
        print("Could not write CO-PSO output files (close them if open in Excel):", exc)
        report = dataset/"CO_PSO_Mapping_Report.xlsx"
    return {"mapping":mapping,"justifications":justifications,"report_path":str(report),"model_name":model_name}
