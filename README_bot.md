# AOW Document Converter Bot

Telegram bot for processing AOW GROUP documents from your phone.

## Setup

### 1. Create Telegram bot
- Open [@BotFather](https://t.me/BotFather)
- Send `/newbot`, follow instructions
- Copy the **token**

### 2. Get your Telegram user ID
- Open [@userinfobot](https://t.me/userinfobot)
- Copy your **ID** (number)

### 3. Deploy to Railway
1. Push this repo to GitHub
2. Go to [railway.app](https://railway.app) → New Project → Deploy from GitHub
3. Add environment variables:
   - `BOT_TOKEN` = your bot token from BotFather
   - `ALLOWED_USER_ID` = your Telegram user ID

### 4. Usage
1. Open your bot in Telegram
2. Send `/start`
3. Send Pro Forma (DOCX) and/or Spec (XLSX) files
4. Press **Continue**
5. Answer the parameter questions
6. Receive processed files

## Files
- `bot.py` — Telegram bot
- `processing.py` — document processing logic
- `requirements.txt` — dependencies
- `railway.toml` — Railway config
