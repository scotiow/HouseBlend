# HouseBlend Dash App

This repository now includes a Dash UI that lets you:

- Upload a Hansard Excel file (`.xlsx`)
- Run HouseBlend schedule optimisation
- Download the updated Hansard file
- Download the generated mailmerge workbook

## Privacy statement

- Any data uploaded is at the user's own risk.
- Uploaded files are processed for the active request only.
- The app does not intentionally persist uploaded data server-side after processing completes.
- If you deploy behind reverse proxies, observability tooling, or platform logs, review their retention settings separately.

## Run locally

```bash
uv sync
uv run python dash_app.py
```

Then open `http://127.0.0.1:8050`.

## Deployment recommendation

For this specific app, **Cloud Run** is the best default choice:

- Better long-term scalability than PythonAnywhere
- Simpler containerized deployment than App Engine + more predictable runtime
- Handles occasional heavier optimisation jobs more reliably (request CPU/memory per container)

Use **PythonAnywhere** if you want the fastest setup with minimal DevOps and expect low traffic.

Use **App Engine** if your org is already standardized on App Engine and you need tighter integration with that platform.

## PythonAnywhere quick notes

- Use the WSGI entrypoint in [wsgi.py](/Users/Engs1825/Documents/GitHub/HouseBlend/wsgi.py)
- Point the WSGI file to `from dash_app import server as application`
- Install dependencies from `pyproject.toml`

## Cloud Run quick notes

- Use the included [Dockerfile](/Users/Engs1825/Documents/GitHub/HouseBlend/Dockerfile) (`gunicorn dash_app:server`)
- Build and deploy to Cloud Run
- Configure memory/CPU based on optimisation size and CVXPY solver usage
