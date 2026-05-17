import io
import os
from dotenv import load_dotenv
import json
import time
import smtplib
import threading
from flask import Flask, render_template_string, jsonify
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from datetime import datetime, timedelta

import requests
from bs4 import BeautifulSoup
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib.colors import Color
import xml.sax.saxutils as saxutils

import schedule
load_dotenv()


# ═══════════════════════════════════════════════════
#  CONFIG  ← edit these values
# ═══════════════════════════════════════════════════
EMAIL_SENDER       = os.getenv("Email")      # Gmail address
EMAIL_APP_PASSWORD =  os.getenv("psswd") 
print(EMAIL_SENDER,EMAIL_APP_PASSWORD)      # 16-char App Password
EMAIL_RECIPIENT    = "hrudayodgaar@gmail.com"       # where to send the report
EMAIL_SUBJECT      = "Inshorts News Report"

HEALTH_CHECK_URL   = "https://trial271225.onrender.com/"       # site to ping (keep-alive)
HEALTH_CHECK_EVERY = 10                             # minutes between pings

CATEGORIES  = ['national', 'business', 'sports', 'world', 'technology', 'startup']
MAX_PAGES   = 10

SEND_INTERVAL_DAYS = 3                             # email every N days
# ═══════════════════════════════════════════════════

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "en-US,en;q=0.9",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
}


# ───────────────────────────────────────────────────
#  SCRAPING
# ───────────────────────────────────────────────────

def fetch_page(url: str, retries: int = 3) -> str | None:
    for attempt in range(1, retries + 1):
        try:
            resp = requests.get(url, headers=HEADERS, timeout=15)
            resp.raise_for_status()
            return resp.text
        except requests.RequestException as e:
            print(f"  [Attempt {attempt}/{retries}] {e}")
            time.sleep(2 * attempt)
    return None


def parse_articles(html: str) -> list[dict]:
    soup      = BeautifulSoup(html, "html.parser")
    headlines = soup.find_all("span", itemprop="headline")
    bodies    = soup.find_all("div",  itemprop="articleBody")
    dates     = soup.find_all(class_="date")
    return [
        {"headline": h.get_text(strip=True),
         "body":     b.get_text(strip=True),
         "date":     d.get_text(strip=True)}
        for h, b, d in zip(headlines, bodies, dates)
    ]


def fetch_category(category: str, max_pages: int) -> list[dict]:
    base_url     = f"https://www.inshorts.com/en/read/{category}"
    seen         = set()
    all_articles = []
    offset       = None

    for page in range(max_pages):
        url  = base_url if offset is None else f"{base_url}?news_offset={offset}"
        html = fetch_page(url)
        if not html:
            break
        articles = parse_articles(html)
        if not articles:
            break
        for art in articles:
            if art["headline"] not in seen:
                seen.add(art["headline"])
                all_articles.append(art)

        soup  = BeautifulSoup(html, "html.parser")
        cards = soup.find_all("div", class_="news-card")
        if cards:
            last_id = cards[-1].get("data-news-id") or cards[-1].get("id")
            if last_id and last_id != offset:
                offset = last_id
            else:
                break
        else:
            break
        time.sleep(1)

    return all_articles


# ───────────────────────────────────────────────────
#  REPORT BUILDER
# ───────────────────────────────────────────────────

def build_report(log_fn=print):
    """Build the .pdf report purely in memory; returns (BytesIO, article_count)."""
    log_fn("Building report...")
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=letter,
        rightMargin=0.5*inch,
        leftMargin=0.5*inch,
        topMargin=0.4*inch,
        bottomMargin=0.4*inch
    )
    
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('Title', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=16, spaceAfter=6)
    date_style = ParagraphStyle('Date', parent=styles['Normal'], fontName='Helvetica', fontSize=9, textColor=Color(120/255, 120/255, 120/255), spaceAfter=12)
    cat_style = ParagraphStyle('Category', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=13, spaceAfter=8, spaceBefore=12)
    headline_style = ParagraphStyle('Headline', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=10, spaceAfter=4)
    art_date_style = ParagraphStyle('ArtDate', parent=styles['Normal'], fontName='Helvetica', fontSize=8, textColor=Color(100/255, 100/255, 100/255), spaceAfter=4)
    body_style = ParagraphStyle('Body', parent=styles['Normal'], fontName='Helvetica', fontSize=9, spaceAfter=12)
    sep_style = ParagraphStyle('Separator', parent=styles['Normal'], fontName='Helvetica', fontSize=8, textColor=Color(150/255, 150/255, 150/255), spaceAfter=12)
    
    story = []
    story.append(Paragraph('Inshorts News Report', title_style))
    story.append(Paragraph(f"Generated: {datetime.now().strftime('%d %B %Y, %I:%M %p')}", date_style))
    
    total = 0
    for cat in CATEGORIES:
        log_fn(f"  Scraping: {cat.upper()}")
        articles = fetch_category(cat, MAX_PAGES)
        story.append(Paragraph(cat.capitalize(), cat_style))
        if not articles:
            story.append(Paragraph("  (No articles found)", body_style))
            continue
            
        for art in articles:
            headline = saxutils.escape(art["headline"])
            body = saxutils.escape(art["body"])
            date_str = saxutils.escape(art["date"])
            
            story.append(Paragraph(headline, headline_style))
            story.append(Paragraph(f"Date: {date_str}", art_date_style))
            story.append(Paragraph(body, body_style))
            story.append(Paragraph("-" * 50, sep_style))
            total += 1
        log_fn(f"    -> {len(articles)} articles")

    doc.build(story)
    buf.seek(0)
    log_fn(f"Report built in memory ({total} articles) — no file saved.")
    return buf, total


# ───────────────────────────────────────────────────
#  EMAIL
# ───────────────────────────────────────────────────

def send_email(doc_buf: io.BytesIO, log_fn=print) -> bool:
    """Attach the in-memory .pdf buffer and send; no file is written to disk."""
    try:
        log_fn(f"Sending email to {EMAIL_RECIPIENT}...")
        filename = f"Inshorts_Report_{datetime.now().strftime('%Y%m%d')}.pdf"

        msg = MIMEMultipart()
        msg["From"]    = EMAIL_SENDER
        msg["To"]      = EMAIL_RECIPIENT
        msg["Subject"] = f"{EMAIL_SUBJECT} - {datetime.now().strftime('%d %b %Y')}"

        body = (
            f"Hi,\n\nPlease find attached the latest Inshorts News Report "
            f"generated on {datetime.now().strftime('%d %B %Y at %I:%M %p')}.\n\n"
            f"This report is sent automatically every {SEND_INTERVAL_DAYS} days.\n\n"
            f"Regards,\nInshorts Reporter Bot"
        )
        msg.attach(MIMEText(body, "plain"))

        part = MIMEBase("application", "pdf")
        part.set_payload(doc_buf.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename={filename}")
        msg.attach(part)

        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(EMAIL_SENDER, EMAIL_APP_PASSWORD)
            server.sendmail(EMAIL_SENDER, EMAIL_RECIPIENT, msg.as_string())

        log_fn("Email sent successfully.")
        return True
    except Exception as e:
        log_fn(f"Email error: {e}")
        return False


# ───────────────────────────────────────────────────
#  HEALTH CHECKER
# ───────────────────────────────────────────────────

def health_check(log_fn=print):
    try:
        resp = requests.get(HEALTH_CHECK_URL, timeout=15)
        status = f"✅ {HEALTH_CHECK_URL} → {resp.status_code}"
    except Exception as e:
        status = f"❌ {HEALTH_CHECK_URL} → {e}"
    log_fn(f"[Health] {datetime.now().strftime('%H:%M:%S')}  {status}")
    return status


# ───────────────────────────────────────────────────
#  COMBINED JOB  (build + send + update state)
# ───────────────────────────────────────────────────

def run_job(state: dict, log_fn=print, refresh_ui_fn=None):
    try:
        doc_buf, total = build_report(log_fn)
        ok = send_email(doc_buf, log_fn)

        now_str = datetime.now().strftime("%d %b %Y  %I:%M %p")
        record  = {
            "sent_at": now_str,
            "status":  "✅ Sent" if ok else "❌ Failed",
            "articles": total,
        }
        state["email_history"].insert(0, record)
        state["email_history"] = state["email_history"][:3]   # keep last 3

        next_ts = (datetime.now() + timedelta(days=SEND_INTERVAL_DAYS)).timestamp()
        state["next_send_ts"] = next_ts
        log_fn(f"Next send scheduled: {datetime.fromtimestamp(next_ts).strftime('%d %b %Y %I:%M %p')}")
    except Exception as e:
        log_fn(f"Job error: {e}")
    finally:
        if refresh_ui_fn:
            refresh_ui_fn()


# ───────────────────────────────────────────────────
#  FLASK APP & BACKGROUND JOBS
# ───────────────────────────────────────────────────

app = Flask(__name__)

# Global state and logs
app_state = {}
app_logs = []
health_status = "Waiting for first ping..."

def add_log(msg: str):
    ts = datetime.now().strftime("%H:%M:%S")
    app_logs.append(f"[{ts}] {msg}")
    if len(app_logs) > 100:
        app_logs.pop(0)

def bg_schedule_loop():
    def check_schedule():
        next_ts = app_state.get("next_send_ts")
        if next_ts and datetime.now().timestamp() >= next_ts:
            do_run()

    next_ts = app_state.get("next_send_ts")
    if next_ts is None:
        do_run()
    else:
        schedule.every(1).minutes.do(check_schedule)
        while True:
            schedule.run_pending()
            time.sleep(30)

def bg_health_loop():
    def do_health():
        global health_status
        health_status = health_check(add_log)
    
    schedule.every(HEALTH_CHECK_EVERY).minutes.do(do_health)
    do_health()
    while True:
        schedule.run_pending()
        time.sleep(10)

def do_run():
    add_log("Starting scheduled job...")
    threading.Thread(
        target=run_job,
        args=(app_state, add_log, None),
        daemon=True,
    ).start()

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Inshorts Reporter</title>
    <script src="https://unpkg.com/feather-icons"></script>
    <style>
        :root {
            --bg-color: #f9fafb;
            --surface-color: #ffffff;
            --text-primary: #111827;
            --text-secondary: #6b7280;
            --border-color: #e5e7eb;
            --hover-color: #f3f4f6;
            --button-bg: #111827;
            --button-text: #ffffff;
            --button-hover: #374151;
        }

        * { box-sizing: border-box; }
        body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
            background-color: var(--bg-color);
            color: var(--text-primary);
            margin: 0;
            padding: 20px 10px;
            display: flex;
            justify-content: center;
        }
        .container {
            width: 100%;
            max-width: 640px;
            display: flex;
            flex-direction: column;
            gap: 20px;
        }
        .header {
            text-align: center;
            padding: 20px 0;
        }
        .header h1 {
            margin: 0 0 8px 0;
            font-size: 28px;
            font-weight: 800;
            letter-spacing: -0.5px;
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 10px;
        }
        .header p {
            margin: 0;
            color: var(--text-secondary);
            font-size: 15px;
        }
        .card {
            background-color: var(--surface-color);
            border: 1px solid var(--border-color);
            border-radius: 12px;
            padding: 24px;
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05), 0 2px 4px -1px rgba(0, 0, 0, 0.03);
            transition: transform 0.2s ease, box-shadow 0.2s ease;
        }
        .card:hover {
            box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.05), 0 4px 6px -2px rgba(0, 0, 0, 0.03);
        }
        .card-header {
            display: flex;
            align-items: center;
            gap: 10px;
            margin-bottom: 16px;
            border-bottom: 1px solid var(--border-color);
            padding-bottom: 12px;
        }
        .card-header h2 {
            margin: 0;
            font-size: 16px;
            font-weight: 600;
        }
        .card-header svg {
            color: var(--text-secondary);
        }
        
        /* Countdown styling */
        .countdown-container {
            display: flex;
            flex-direction: column;
            align-items: flex-start;
        }
        #countdown {
            font-size: 32px;
            font-weight: 700;
            font-variant-numeric: tabular-nums;
            letter-spacing: -0.5px;
            margin-bottom: 4px;
        }
        #next_date {
            font-size: 14px;
            color: var(--text-secondary);
        }

        /* History styling */
        .history-list {
            display: flex;
            flex-direction: column;
            gap: 12px;
        }
        .history-item {
            display: flex;
            align-items: center;
            justify-content: space-between;
            padding: 12px;
            background-color: var(--bg-color);
            border-radius: 8px;
            border: 1px solid var(--border-color);
            font-size: 14px;
        }
        .history-item-left {
            display: flex;
            align-items: center;
            gap: 12px;
        }
        .history-status {
            font-weight: 600;
            display: flex;
            align-items: center;
            gap: 6px;
        }
        .status-sent { color: #10b981; }
        .status-failed { color: #ef4444; }

        /* Health */
        .health-status {
            font-size: 15px;
            display: flex;
            align-items: center;
            gap: 8px;
            padding: 12px;
            background-color: var(--bg-color);
            border-radius: 8px;
            border: 1px solid var(--border-color);
        }

        /* Logs */
        .log-box {
            font-family: SFMono-Regular, Consolas, 'Liberation Mono', Menlo, monospace;
            font-size: 13px;
            height: 200px;
            overflow-y: auto;
            background-color: #111827;
            color: #e5e7eb;
            padding: 16px;
            border-radius: 8px;
            line-height: 1.5;
        }
        .log-box::-webkit-scrollbar { width: 8px; }
        .log-box::-webkit-scrollbar-track { background: #1f2937; border-radius: 4px; }
        .log-box::-webkit-scrollbar-thumb { background: #4b5563; border-radius: 4px; }

        /* Actions */
        .actions {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 16px;
        }
        button {
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 8px;
            background-color: var(--button-bg);
            color: var(--button-text);
            border: none;
            padding: 14px 20px;
            border-radius: 8px;
            font-size: 15px;
            font-weight: 600;
            cursor: pointer;
            transition: background-color 0.2s ease, transform 0.1s ease;
        }
        button:hover { background-color: var(--button-hover); }
        button:active { transform: scale(0.98); }
        button.secondary {
            background-color: var(--surface-color);
            color: var(--text-primary);
            border: 1px solid var(--border-color);
        }
        button.secondary:hover {
            background-color: var(--hover-color);
        }

        @media (max-width: 480px) {
            .actions { grid-template-columns: 1fr; }
            .history-item { flex-direction: column; align-items: flex-start; gap: 8px; }
            #countdown { font-size: 28px; }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1><i data-feather="file-text"></i> Inshorts Reporter</h1>
            <p>Automated news digest via email</p>
        </div>
        
        <div class="card">
            <div class="card-header">
                <i data-feather="clock"></i>
                <h2>Next Email Countdown</h2>
            </div>
            <div class="countdown-container">
                <div id="countdown">Loading...</div>
                <div id="next_date"></div>
            </div>
        </div>

        <div class="card">
            <div class="card-header">
                <i data-feather="mail"></i>
                <h2>Last 3 Emails Sent</h2>
            </div>
            <div class="history-list" id="history">Loading...</div>
        </div>

        <div class="card">
            <div class="card-header">
                <i data-feather="activity"></i>
                <h2>Site Health</h2>
            </div>
            <div class="health-status" id="health">Loading...</div>
        </div>

        <div class="card">
            <div class="card-header">
                <i data-feather="terminal"></i>
                <h2>Activity Log</h2>
            </div>
            <div class="log-box" id="logs">Loading...</div>
        </div>

        <div class="actions">
            <button onclick="runNow()"><i data-feather="play"></i> Run Now</button>
            <button onclick="pingNow()" class="secondary"><i data-feather="zap"></i> Ping Now</button>
        </div>
    </div>

    <script>
        feather.replace();

        function updateUI() {
            fetch('/api/status')
                .then(r => r.json())
                .then(data => {
                    // Update countdown
                    if (data.next_ts === null) {
                        document.getElementById('countdown').innerText = "Send pending...";
                        document.getElementById('next_date').innerText = "(Run once to schedule)";
                    } else {
                        let now = new Date().getTime() / 1000;
                        let delta = data.next_ts - now;
                        if (delta <= 0) {
                            document.getElementById('countdown').innerText = "Sending now...";
                            document.getElementById('next_date').innerText = "";
                        } else {
                            let d = Math.floor(delta / 86400);
                            let h = Math.floor((delta % 86400) / 3600);
                            let m = Math.floor((delta % 3600) / 60);
                            let s = Math.floor(delta % 60);
                            
                            let pad = num => String(num).padStart(2, '0');
                            document.getElementById('countdown').innerText = `${d}d ${pad(h)}h ${pad(m)}m ${pad(s)}s`;
                            
                            let nextDate = new Date(data.next_ts * 1000).toLocaleString();
                            document.getElementById('next_date').innerText = `Scheduled for: ${nextDate}`;
                        }
                    }

                    // Update health
                    let healthEl = document.getElementById('health');
                    let healthText = data.health;
                    // Replace emojis with feather icons if any exist in status text
                    healthText = healthText.replace('✅', '<i data-feather="check-circle" style="color: #10b981; width: 18px; height: 18px;"></i>');
                    healthText = healthText.replace('❌', '<i data-feather="x-circle" style="color: #ef4444; width: 18px; height: 18px;"></i>');
                    healthEl.innerHTML = healthText;

                    // Update history
                    let histHtml = "";
                    if (data.history.length === 0) {
                        histHtml = "<div class='history-item'>No emails sent yet.</div>";
                    } else {
                        data.history.forEach((h, i) => {
                            let isSent = h.status.includes('Sent');
                            let statusClass = isSent ? 'status-sent' : 'status-failed';
                            let iconName = isSent ? 'check' : 'x';
                            
                            // strip emojis from status if they exist
                            let cleanStatus = h.status.replace('✅ ', '').replace('❌ ', '');

                            histHtml += `<div class="history-item">
                                <div class="history-item-left">
                                    <div class="history-status ${statusClass}">
                                        <i data-feather="${iconName}" style="width: 16px; height: 16px;"></i>
                                        ${cleanStatus}
                                    </div>
                                    <div style="color: var(--text-secondary);">${h.sent_at}</div>
                                </div>
                                <div style="font-weight: 500;">${h.articles} articles</div>
                            </div>`;
                        });
                    }
                    document.getElementById('history').innerHTML = histHtml;

                    // Update logs
                    let logBox = document.getElementById('logs');
                    let wasAtBottom = logBox.scrollHeight - logBox.clientHeight <= logBox.scrollTop + 1;
                    logBox.innerHTML = data.logs.join("<br>");
                    if (wasAtBottom) {
                        logBox.scrollTop = logBox.scrollHeight;
                    }
                    
                    // Re-render new icons
                    feather.replace();
                });
        }

        function runNow() {
            if (confirm("Build report and send email now?")) {
                fetch('/api/run', {method: 'POST'});
            }
        }

        function pingNow() {
            fetch('/api/ping', {method: 'POST'})
                .then(() => updateUI());
        }

        setInterval(updateUI, 1000);
        updateUI();
    </script>
</body>
</html>
"""

@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route('/api/status')
def status():
    return jsonify({
        "next_ts": app_state.get("next_send_ts"),
        "history": app_state.get("email_history", []),
        "health": health_status,
        "logs": app_logs
    })

@app.route('/api/run', methods=['POST'])
def api_run():
    do_run()
    return jsonify({"status": "started"})

@app.route('/api/ping', methods=['POST'])
def api_ping():
    global health_status
    health_status = health_check(add_log)
    return jsonify({"status": "pinged", "health": health_status})

if __name__ == "__main__":
    app_state = {"email_history": [], "next_send_ts": None}
    threading.Thread(target=bg_schedule_loop, daemon=True).start()
    threading.Thread(target=bg_health_loop, daemon=True).start()
    app.run(host='0.0.0.0', port=8080)