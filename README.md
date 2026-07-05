# Lotus Offer Generator - Web

Lotus Pro Claims Mix & Match Engine — web version.

## Login

- **Username:** admin
- **Password:** admin

## Run locally

```bash
pip install -r requirements.txt
python app.py
```

Open http://localhost:17800

## Deploy with Docker

```bash
docker build -t lotus-offer .
docker run -d -p 17800:17800 lotus-offer
```

## Deploy on VPS

```bash
bash deploy/deploy.sh
```

## Original desktop app

`Offer.py` — original CustomTkinter desktop version (v1.4).
