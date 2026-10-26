# AI-Based Automatic CO–PO–PEO Mapping System Using NLP

A Flask-based academic project that automates **CO–PO–PEO mapping** using NLP and Sentence-BERT semantic similarity.

## Main Features

- Subject-wise CO, PO and PEO mapping
- Sentence-BERT (`all-MiniLM-L6-v2`) semantic similarity
- Cosine-similarity based mapping
- Mapping scale: **3 = High, 2 = Moderate, 1 = Low, 0 = None**
- Bloom's Taxonomy detection for COs
- AI-style mapping justifications
- Interactive dashboard and subject report
- PDF, Word and Excel report downloads
- Dataset upload and subject selection
- Responsive UI with ATME branding

## Project Structure

```text
AI_CO_PO_PEO_MAPPING/
├── app.py
├── blooms.py
├── preprocess.py
├── semantic_similarity.py
├── mapping.py
├── co_peo_mapping.py
├── justification.py
├── co_peo_justification.py
├── requirements.txt
├── RUN_INSTRUCTIONS.txt
├── Dataset/
│   ├── CO_Dataset-1.xlsx
│   ├── PO_Dataset.xlsx
│   ├── PEO_Dataset.xlsx
│   ├── Subject_Details.xlsx
│   └── processed CSV files used by the application
├── static/
│   ├── css/style.css
│   └── images/ATME.jpeg
└── templates/
```

## Important

The **PDF/Word/Excel generators used by the web application are integrated directly into `app.py`**. Separate old generator scripts were removed to avoid duplicate and conflicting implementations.

## Run

```text
python -m pip install -r requirements.txt
python app.py
```

Then open:

`http://127.0.0.1:5000`

On first run, the application creates a fresh SQLite user database automatically.

## AI Pipeline

Input datasets → preprocessing → Sentence-BERT embeddings → cosine similarity → top-3 mapping → justification → dashboard/report.

## Note

The dashboard reports the AI model as **SBERT** rather than displaying an unsupported fixed "accuracy %" value. A numerical accuracy should only be shown when a validated ground-truth evaluation dataset is available.


## CO–PSO Mapping (manual Excel dataset)

- Maintain programme-specific outcomes directly in `Dataset/PSO_Dataset.xlsx` (no website upload required).
- Required columns: `PSO ID` and `Programme Specific Outcomes`.
- The two requested default outcomes, PSO1 and PSO2, are pre-filled. A copy is also included as the `PSO Dataset` sheet in `Dataset/CO_Dataset-1.xlsx`.
- After editing the workbook, sign in and open **CO–PSO Mapping** from the dashboard. The app creates the mapping table, per-pair justifications, graph and `Dataset/CO_PSO_Mapping_Report.xlsx`.
- The report includes `CO-PSO Mapping`, `Justifications`, `PSO Dataset`, `Similarity Scores` and `Graph Summary` sheets, plus an embedded bar chart. Download it from the mapping page.
- Sentence-BERT is preferred; if its model cannot load, the mapping utility falls back to TF-IDF cosine similarity. The ranking-derived 3/2/1/0 values are AI suggestions and should be reviewed by faculty.


## PSO placement update
The dashboard shows only the PSO definitions list and the CO–PSO mapping-strength chart. The detailed CO–PSO matrix and justifications are shown in the selected subject report. Maintain PSO definitions in `Dataset/PSO_Dataset.xlsx`. See `PSO_REPORT_ONLY_STRUCTURE.txt`.
