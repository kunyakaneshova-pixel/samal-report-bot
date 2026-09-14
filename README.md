# Samal Report Bot

Telegram bot for SamalCakes iiko Excel reports.

Render:
- Build command: `pip install -r requirements.txt`
- Start command: `gunicorn -b 0.0.0.0:$PORT bot:app`
- Environment variable: `BOT_TOKEN`
- After deploy open `/setup` once to register Telegram webhook.
