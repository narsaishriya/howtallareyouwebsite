# How Tall Are You

A small Flask app for collecting height measurements and viewing participant rankings.

## Run locally

```powershell
cd web_app
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install flask openpyxl
Copy-Item .env.example .env
python app.py
```

Set `FLASK_SECRET_KEY` and `OWNER_PASSWORD` in the environment before running outside local development. Participant records and uploaded photos are intentionally kept local and are not committed to Git.