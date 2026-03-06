FROM python:3.12-slim

WORKDIR /app

COPY pyproject.toml README.md /app/
COPY HouseBlend /app/HouseBlend
COPY dash_app.py wsgi.py /app/

RUN pip install --no-cache-dir .

COPY img /app/img

ENV PORT=8080

CMD ["sh", "-c", "gunicorn dash_app:server --bind 0.0.0.0:${PORT} --workers 1 --threads 4 --timeout 300"]
