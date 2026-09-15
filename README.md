# ICC Course Operational Complexity Calculator

A Flask version of the ICC calculator with live scoring, per-component quantity multipliers,
and downloadable PDF and Excel reports. Quantities default to 0. Reports show quantity,
unit points, and calculated points for each scored component.

## Run locally

```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\\Scripts\\activate
pip install -r requirements.txt
flask --app app run
```

Open `http://127.0.0.1:5000`.

## Production

```bash
gunicorn app:app
```

The app is stateless: no course data is retained on the server. PDF and Excel files are generated in memory when downloaded.
