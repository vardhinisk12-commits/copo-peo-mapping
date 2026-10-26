# Deploy on Render

## Steps
1. Create a GitHub repository and upload everything inside this folder (app.py must be at the repo root).
2. On https://render.com -> New -> **Blueprint** (uses render.yaml) or New -> **Web Service**.
3. If you choose Web Service manually:
   - Runtime: Python 3
   - Build command: `pip install -r requirements.txt`
   - Start command: `gunicorn app:app --workers 1 --threads 4 --timeout 300 --bind 0.0.0.0:$PORT`
   - Environment variable `FLASK_SECRET_KEY` = any long random text
4. Deploy, then open the Render URL and click **Register** to create the first account.

## Free plan (default)
Uses `requirements.txt` (no PyTorch). CO-PO-PEO-PSO mapping runs with TF-IDF cosine similarity because
Sentence-BERT needs more than the 512 MB RAM of the free plan. Reports show the method actually used.

## Sentence-BERT (paid plan, 2 GB RAM or more)
Set build command to `pip install -r requirements-sbert.txt` and choose a Starter/Standard plan.
The model (all-MiniLM-L6-v2) downloads on first use, so the first mapping is slower.

## Data is not permanent on the free plan
Render's free disk is wiped on every deploy/restart. This affects:
- user accounts (SQLite `users.db`)  -> fix: create a Render PostgreSQL DB and set `DATABASE_URL`
- departments uploaded through the website (`Departments/` folder) and generated files
  -> fix: attach a persistent disk (paid) mounted at `/var/data` and set `DEPARTMENTS_DIR=/var/data/Departments`
The built-in CSE data in `Dataset/` is part of the code, so it is always there.

## Other notes
- The free service sleeps after ~15 min idle; the first request afterwards takes about a minute.
- No users are included in this package; register a new one after deploying.
- `--workers 1` is intentional (in-memory caches); do not increase it without testing.

## Other platforms (Fly.io-style / buildpack hosts)
A `Dockerfile` is included (Python 3.12-slim, gunicorn, listens on `$PORT`, default 8080).
There is deliberately **no `.python-version` file**: some buildpack hosts cannot download an exact
patch version such as 3.12.3 and fail with a 404. If your host builds with buildpacks instead of the
Dockerfile, set the Python version it supports (e.g. `3.12`) in its own settings.
