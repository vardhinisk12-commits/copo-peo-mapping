PSO DASHBOARD + SUBJECT REPORT UPDATE

DASHBOARD
- Shows the manually maintained PSO list from Dataset/PSO_Dataset.xlsx.
- Shows the CO-PSO Mapping Strength chart (average mapping strength per PSO).
- Does NOT show the CO-PSO matrix or CO-PSO justifications.

SUBJECT REPORT
- Contains the subject-specific CO-PSO mapping matrix.
- Contains PSO definitions and CO-PSO mapping-strength chart.
- Contains AI/NLP justifications, including CO-PSO justifications, with a PSO filter.
- The detailed CO-PO and CO-PEO sections remain in the subject report as before.

PSO DATA
Maintain PSO definitions manually in Dataset/PSO_Dataset.xlsx. Keep columns:
  PSO ID | Programme Specific Outcomes
The default PSO1 and PSO2 definitions are provided. Do not upload PSOs through the website.

RUN
1. Extract this ZIP into a new folder.
2. Open a terminal in the folder containing app.py.
3. Install dependencies: python -m pip install -r requirements.txt
4. Run: python app.py
5. Open http://127.0.0.1:5000 and sign in.
6. The app creates a fresh SQLite users database on first run.
7. Select a subject and click View Subject Report for detailed matrices and justifications.

MAPPING NOTE
The 3/2/1/0 values are semantic similarity ranking suggestions and should be reviewed by faculty. Sentence-BERT is preferred; TF-IDF is used as a fallback if the model cannot load.
