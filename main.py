import io
import time
import os
from dotenv import load_dotenv
import smtplib
import threading
import tkinter as tk
from tkinter import ttk, messagebox
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from datetime import datetime, timedelta

import requests
from bs4 import BeautifulSoup
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

import schedule
load_dotenv()
# ═══════════════════════════════════════════════════
#  CONFIG  ← edit these values
# ═══════════════════════════════════════════════════
EMAIL_SENDER       = os.getenv("Email")       # Gmail address
EMAIL_APP_PASSWORD = os.getenv("psswd")         # 16-char App Password
EMAIL_RECIPIENT    = "recipient@example.com"       # where to send the report
EMAIL_SUBJECT      = "Inshorts News Report"

HEALTH_CHECK_URL   = "https://your-site.com"       # site to ping (keep-alive)
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
#  STATE  (pure in-memory — no files, no disk writes)
# ───────────────────────────────────────────────────

# Single shared dict; lives only for the duration of this process.
# email_history : list of dicts  (newest first, max 3 entries)
# next_send_ts  : float | None   (Unix timestamp of next scheduled send)
APP_STATE: dict = {
    "email_history": [],
    "next_send_ts":  None,
}


# ───────────────────────────────────────────────────
#  DOCUMENT HELPERS
# ───────────────────────────────────────────────────

def set_margins(doc):
    for section in doc.sections:
        section.top_margin    = Inches(0.4)
        section.bottom_margin = Inches(0.4)
        section.left_margin   = Inches(0.5)
        section.right_margin  = Inches(0.5)


def tight_paragraph(doc, text, bold=False, font_size=10, color=None):
    para = doc.add_paragraph()
    run  = para.add_run(text)
    run.bold = bold
    run.font.size = Pt(font_size)
    if color:
        run.font.color.rgb = RGBColor(*color)
    pPr     = para._p.get_or_add_pPr()
    spacing = OxmlElement('w:spacing')
    spacing.set(qn('w:before'),   '20')
    spacing.set(qn('w:after'),    '20')
    spacing.set(qn('w:line'),    '240')
    spacing.set(qn('w:lineRule'), 'auto')
    pPr.append(spacing)
    return para


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
    """Build the .docx report purely in memory; returns (BytesIO, article_count)."""
    log_fn("Building report...")
    doc = Document()
    set_margins(doc)
    tight_paragraph(doc, 'Inshorts News Report', bold=True, font_size=16)
    tight_paragraph(
        doc,
        f"Generated: {datetime.now().strftime('%d %B %Y, %I:%M %p')}",
        font_size=9, color=(120, 120, 120),
    )
    total = 0
    for cat in CATEGORIES:
        log_fn(f"  Scraping: {cat.upper()}")
        articles = fetch_category(cat, MAX_PAGES)
        tight_paragraph(doc, cat.capitalize(), bold=True, font_size=13)
        if not articles:
            tight_paragraph(doc, "  (No articles found)", font_size=9)
            continue
        for art in articles:
            tight_paragraph(doc, art["headline"], bold=True, font_size=10)
            tight_paragraph(doc, f"Date: {art['date']}", font_size=8, color=(100, 100, 100))
            tight_paragraph(doc, art["body"], font_size=9)
            tight_paragraph(doc, "-" * 30, font_size=8)
            total += 1
        log_fn(f"    -> {len(articles)} articles")

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    log_fn(f"Report built in memory ({total} articles) — no file saved.")
    return buf, total


# ───────────────────────────────────────────────────
#  EMAIL
# ───────────────────────────────────────────────────

def send_email(doc_buf: io.BytesIO, log_fn=print) -> bool:
    """Attach the in-memory .docx buffer and send; no file is written to disk."""
    try:
        log_fn(f"Sending email to {EMAIL_RECIPIENT}...")
        filename = f"Inshorts_Report_{datetime.now().strftime('%Y%m%d')}.docx"

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

        part = MIMEBase("application", "octet-stream")
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
#  TKINTER UI
# ───────────────────────────────────────────────────

class ReporterApp(tk.Tk):
    DARK_BG   = "#1e1e2e"
    CARD_BG   = "#2a2a3e"
    ACCENT    = "#7c6af7"
    TEXT      = "#cdd6f4"
    MUTED     = "#6c7086"
    GREEN     = "#a6e3a1"
    RED       = "#f38ba8"
    AMBER     = "#fab387"

    def __init__(self, state: dict):
        super().__init__()
        self.state     = state
        self.title("Inshorts Reporter")
        self.geometry("680x640")
        self.resizable(False, False)
        self.configure(bg=self.DARK_BG)

        self._build_ui()
        self._start_background_jobs()
        self._tick()          # start 1-second countdown loop

    # ── UI construction ──────────────────────────────

    def _build_ui(self):
        # ── Header ──
        hdr = tk.Frame(self, bg=self.ACCENT, pady=10)
        hdr.pack(fill="x")
        tk.Label(hdr, text="📰  Inshorts Reporter", font=("Helvetica", 16, "bold"),
                 bg=self.ACCENT, fg="white").pack()
        tk.Label(hdr, text="Automated news digest via email",
                 font=("Helvetica", 9), bg=self.ACCENT, fg="#e0d8ff").pack()

        body = tk.Frame(self, bg=self.DARK_BG, padx=20, pady=15)
        body.pack(fill="both", expand=True)

        # ── Countdown card ──
        cd_card = self._card(body, "⏱  Next Email Countdown")
        self.countdown_var = tk.StringVar(value="Calculating…")
        tk.Label(cd_card, textvariable=self.countdown_var,
                 font=("Courier", 22, "bold"),
                 bg=self.CARD_BG, fg=self.ACCENT).pack(pady=(0, 4))
        self.next_date_var = tk.StringVar(value="")
        tk.Label(cd_card, textvariable=self.next_date_var,
                 font=("Helvetica", 9), bg=self.CARD_BG, fg=self.MUTED).pack()

        # ── Email history ──
        hist_card = self._card(body, "📬  Last 3 Emails Sent")
        self.hist_frame = tk.Frame(hist_card, bg=self.CARD_BG)
        self.hist_frame.pack(fill="x")
        self._refresh_history()

        # ── Health checker ──
        hc_card = self._card(body, "🌐  Site Health")
        self.health_var = tk.StringVar(value="Waiting for first ping…")
        tk.Label(hc_card, textvariable=self.health_var,
                 font=("Helvetica", 10), bg=self.CARD_BG,
                 fg=self.TEXT, wraplength=580, justify="left").pack(anchor="w")

        # ── Log box ──
        log_card = self._card(body, "📋  Activity Log")
        self.log_box = tk.Text(log_card, height=7, bg="#12121e", fg=self.TEXT,
                               font=("Courier", 9), relief="flat",
                               state="disabled", wrap="word")
        self.log_box.pack(fill="x")
        sb = ttk.Scrollbar(log_card, command=self.log_box.yview)
        self.log_box.configure(yscrollcommand=sb.set)

        # ── Buttons ──
        btn_row = tk.Frame(self, bg=self.DARK_BG, pady=10)
        btn_row.pack()
        self._btn(btn_row, "▶  Run Now", self.ACCENT, self._manual_run).pack(side="left", padx=8)
        self._btn(btn_row, "🌐  Ping Now", self.MUTED, self._manual_ping).pack(side="left", padx=8)

    def _card(self, parent, title: str) -> tk.Frame:
        outer = tk.Frame(parent, bg=self.DARK_BG, pady=6)
        outer.pack(fill="x")
        tk.Label(outer, text=title, font=("Helvetica", 10, "bold"),
                 bg=self.DARK_BG, fg=self.MUTED).pack(anchor="w")
        inner = tk.Frame(outer, bg=self.CARD_BG, padx=12, pady=10,
                         highlightbackground=self.ACCENT,
                         highlightthickness=1)
        inner.pack(fill="x")
        return inner

    def _btn(self, parent, text, color, cmd):
        return tk.Button(parent, text=text, command=cmd,
                         bg=color, fg="white", font=("Helvetica", 10, "bold"),
                         relief="flat", padx=16, pady=6, cursor="hand2",
                         activebackground=self.ACCENT, activeforeground="white")

    # ── Refresh helpers ──────────────────────────────

    def _refresh_history(self):
        for w in self.hist_frame.winfo_children():
            w.destroy()
        history = self.state.get("email_history", [])
        if not history:
            tk.Label(self.hist_frame, text="No emails sent yet.",
                     bg=self.CARD_BG, fg=self.MUTED,
                     font=("Helvetica", 9)).pack(anchor="w")
            return
        for i, rec in enumerate(history[:3]):
            color = self.GREEN if "Sent" in rec["status"] else self.RED
            row = tk.Frame(self.hist_frame, bg=self.CARD_BG)
            row.pack(fill="x", pady=2)
            tk.Label(row, text=f"{i+1}.", width=2,
                     bg=self.CARD_BG, fg=self.MUTED,
                     font=("Courier", 9)).pack(side="left")
            tk.Label(row, text=rec["status"], width=10,
                     bg=self.CARD_BG, fg=color,
                     font=("Courier", 9, "bold")).pack(side="left")
            tk.Label(row, text=rec["sent_at"],
                     bg=self.CARD_BG, fg=self.TEXT,
                     font=("Helvetica", 9)).pack(side="left", padx=8)
            articles = rec.get("articles", "?")
            tk.Label(row, text=f"({articles} articles)",
                     bg=self.CARD_BG, fg=self.MUTED,
                     font=("Helvetica", 9)).pack(side="left")

    def _refresh_all(self):
        self._refresh_history()

    # ── Countdown tick (every second) ─────────────────

    def _tick(self):
        next_ts = self.state.get("next_send_ts")
        if next_ts is None:
            self.countdown_var.set("Send pending…")
            self.next_date_var.set("(run once to schedule)")
        else:
            delta = datetime.fromtimestamp(next_ts) - datetime.now()
            if delta.total_seconds() <= 0:
                self.countdown_var.set("Sending now…")
                self.next_date_var.set("")
            else:
                total_s  = int(delta.total_seconds())
                days     = total_s // 86400
                hours    = (total_s % 86400) // 3600
                minutes  = (total_s % 3600) // 60
                seconds  = total_s % 60
                self.countdown_var.set(f"{days:02d}d  {hours:02d}h  {minutes:02d}m  {seconds:02d}s")
                self.next_date_var.set(
                    f"Scheduled: {datetime.fromtimestamp(next_ts).strftime('%d %b %Y  %I:%M %p')}"
                )
        self.after(1000, self._tick)

    # ── Logging ──────────────────────────────────────

    def log(self, msg: str):
        ts = datetime.now().strftime("%H:%M:%S")
        self.log_box.configure(state="normal")
        self.log_box.insert("end", f"[{ts}] {msg}\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    # ── Background jobs ──────────────────────────────

    def _start_background_jobs(self):
        # Schedule email job every N days (checks every minute)
        def _schedule_loop():
            next_ts = self.state.get("next_send_ts")
            if next_ts is None:
                # First run: send immediately then schedule
                self._do_run()
            else:
                schedule.every(1).minutes.do(self._check_schedule)
                while True:
                    schedule.run_pending()
                    time.sleep(30)

        # Schedule health check
        def _health_loop():
            schedule.every(HEALTH_CHECK_EVERY).minutes.do(self._do_health)
            self._do_health()   # run once immediately
            while True:
                schedule.run_pending()
                time.sleep(10)

        threading.Thread(target=_schedule_loop, daemon=True).start()
        threading.Thread(target=_health_loop,   daemon=True).start()

    def _check_schedule(self):
        next_ts = self.state.get("next_send_ts")
        if next_ts and datetime.now().timestamp() >= next_ts:
            self._do_run()

    def _do_run(self):
        self.log("Starting scheduled job…")
        threading.Thread(
            target=run_job,
            args=(self.state, self.log, self._refresh_all),
            daemon=True,
        ).start()

    def _do_health(self):
        def _ping():
            status = health_check(self.log)
            color  = self.GREEN if "✅" in status else self.RED
            self.health_var.set(status)
            # update label color
            for widget in self.winfo_children():
                pass   # colour already embedded in status string
        threading.Thread(target=_ping, daemon=True).start()

    # ── Manual buttons ───────────────────────────────

    def _manual_run(self):
        if messagebox.askyesno("Confirm", "Build report and send email now?"):
            self._do_run()

    def _manual_ping(self):
        self._do_health()


# ───────────────────────────────────────────────────
#  ENTRY POINT
# ───────────────────────────────────────────────────

if __name__ == "__main__":
    app = ReporterApp(APP_STATE)
    app.mainloop()