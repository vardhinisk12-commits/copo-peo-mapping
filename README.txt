AI-Based Automatic CO-PO-PEO Mapping System  (ATME College of Engineering)
=========================================================================

RUN ON WINDOWS (Python 3.11)
1. Extract the ZIP. Open PowerShell inside the folder that contains app.py.
2. (Recommended) create a virtual environment:
       py -3.11 -m venv venv
       venv\Scripts\activate
3. Install dependencies:
       python -m pip install -r requirements.txt
4. Start the application:
       python app.py
5. Read the line printed in the window:
       Open this address in Chrome or Edge:  http://127.0.0.1:5000/
   Open exactly that address (http://, NOT https://). If port 5000 was busy, the
   program prints and uses the next free port (5001, 5002 ...) - use the printed one.
6. Register an account and log in. You are asked to choose a DEPARTMENT:
   - "Computer Science and Engineering" is the built-in dataset (folder Dataset\).
   - To add another department, fill in section 2 of that page and upload five Excel (.xlsx) files:
         Subjects : Subject Code, Subject Name  (optional: Semester, Credits, Schema, Department)
         CO       : Subject Code, CO ID, Course Outcome
         PO       : PO ID (PO1, PO2 ...), description
         PEO      : PEO ID (PEO1, PEO2 ...), description
         PSO      : PSO ID (PSO1, PSO2 ...), description
     The CO-PO, CO-PEO and CO-PSO mappings are generated automatically (can take a minute).
     Each department is stored separately in Departments\<name>\ and never changes the others.
     Use "Department ... (change / upload)" in the dashboard top bar to switch department later.
7. Pick a subject and click "View Subject Report". Stop the server with Ctrl+C.

NOTES
- The Flask auto-reloader is OFF on purpose (it restarted endlessly because it watched the
  "transformers" package). After editing a .py file, press Ctrl+C and run python app.py again.
  Optional debugger page:  set FLASK_DEBUG=1   then   python app.py
- Charts (Chart.js) and icons (Font Awesome) load from a CDN, so the browser needs internet access.
- Sentence-BERT (all-MiniLM-L6-v2) is downloaded from Hugging Face on first use (~90 MB, needs internet
  once). If it cannot be loaded, the app falls back to TF-IDF similarity and still works.
- Close Dataset\CO_PSO_Mapping_Report.xlsx in Excel before using the app (Windows locks open files).
- If the browser still shows "Not Found" or ERR_CONNECTION_RESET: close ALL old Python/PowerShell windows
  (or run:  taskkill /F /IM python.exe ), start again and use the address that is printed.
- PSO definitions are edited in Dataset\PSO_Dataset.xlsx (columns: PSO ID | Programme Specific Outcomes).
- Old backups and sample generated reports from the original folder are kept in _archive\.
