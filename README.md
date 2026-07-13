# VARDAN.ai

Full-stack hospital operations dashboard: doctor directory, contact import, reusable media, AI knowledge, campaigns, broadcasts, and delivery logging.

## Run locally

```powershell
npm install
Copy-Item .env.example .env
npm start
```

Open [http://localhost:3000](http://localhost:3000).

## Integrations

Set new credentials in `.env` (never commit the file):

- `GROQ_API_KEY` activates grounded patient AI answers using the uploaded knowledge base and doctor information.
- `WHATSAPP_TOKEN` and `WHATSAPP_PHONE_NUMBER_ID` activate real WhatsApp Cloud API broadcast delivery.
- `GOOGLE_SHEETS_WEBHOOK_URL` receives newly imported contacts for Google Sheets mirroring.

The server persists records in a local SQLite database under `data/` and uploaded files under `uploads/`. Those folders are intentionally excluded from Git.
