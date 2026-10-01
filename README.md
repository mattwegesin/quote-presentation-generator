# Proposal Presentation Generator Web App

Production-ready Flask web service that accepts a base PowerPoint presentation (`.pptx`) and an Excel quote workbook (`.xlsx`), dynamically compiling financial tables, investment metric summary cards, and legal caveats into a downloadable presentation deck (`Generated_Proposal.pptx`).

---

## Architecture Overview

- **In-Memory Streaming:** All uploaded `.pptx` and `.xlsx` files are read into `io.BytesIO()` binary streams. The generated deck is compiled directly in RAM and streamed back to the client without ever writing temporary files to local storage.
- **Frontend:** Responsive, modern interface with drag-and-drop file upload zones, client-side validation, and progress indicators.
- **Backend:** Flask web server with secure multipart processing, input validation, and Gunicorn WSGI production worker configuration.
- **Engine:** `python-pptx` and `pandas` modular automation layer.

---

## Local Development & Testing

1. **Navigate to the web project directory:**
   ```bash
   cd quote_deck_web
   ```

2. **(Optional) Create and activate a virtual environment:**
   ```bash
   python3 -m venv venv
   source venv/bin/activate
   ```

3. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

4. **Run the local development server:**
   ```bash
   python3 app.py
   ```
   Open `http://localhost:5000` in your web browser.

---

## Deployment to Render via GitHub

### Step 1: Initialize Git Repository & Commit

Run the following commands inside `quote_deck_web`:

```bash
cd /Users/mattwegesin/Documents/Gemini_Start/quote_deck_web
git init
git add .
git commit -m "feat: Initial commit for Quote Presentation Generator web service"
```

### Step 2: Push to GitHub

Create a new repository on [GitHub](https://github.com/new) (e.g., `quote-presentation-generator`), then push your code:

```bash
git branch -M main
git remote add origin https://github.com/<YOUR_GITHUB_USERNAME>/quote-presentation-generator.git
git push -u origin main
```

### Step 3: Link to Render

1. Log into your [Render Dashboard](https://dashboard.render.com/).
2. Click **New +** and select **Web Service** (or **Blueprint** if using `render.yaml`).
3. Select your newly created GitHub repository `quote-presentation-generator`.
4. Configure the service settings:
   - **Name:** `quote-presentation-generator`
   - **Language / Runtime:** `Python`
   - **Branch:** `main`
   - **Region:** Choose your preferred region (e.g., `Oregon (US West)` or `Ohio (US East)`)
   - **Build Command:** `pip install -r requirements.txt`
   - **Start Command:** `gunicorn app:app --bind 0.0.0.0:$PORT --workers 2 --timeout 120`
   - **Instance Type:** `Free` or `Starter`
5. Under **Environment Variables**, add:
   - `PYTHON_VERSION`: `3.11.9`
   - `FLASK_SECRET_KEY`: (Click "Generate" or provide a secure random string)
6. Click **Deploy Web Service**.
7. Once deployed, Render will provide a live HTTPS URL (e.g., `https://quote-presentation-generator.onrender.com`).
