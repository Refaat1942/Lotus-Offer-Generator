FROM python:3.11-slim

WORKDIR /app

RUN apt-get update && apt-get install -y --no-install-recommends \
    gcc && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

RUN mkdir -p data static/uploads

ENV PORT=17800
ENV ADMIN_USER=admin
ENV ADMIN_PASS=admin

EXPOSE 17800

CMD ["gunicorn", "--bind", "0.0.0.0:17800", "--workers", "2", "--timeout", "120", "app:app"]
