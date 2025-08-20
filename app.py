import io
import os
from datetime import datetime
from typing import Dict, List, Optional, Tuple

import pandas as pd
from flask import Flask, request, render_template_string, redirect, url_for, flash, send_file

app = Flask(__name__)
app.secret_key = "dev-secret"
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024  # 100MB
ALLOWED_EXT = {".csv", ".xlsx", ".xls"}

# -------------------- UI --------------------
# Updated CSS with day/night mode, animations, and better styling
BASE_CSS = """
:root {
  --bg-dark: #0b1020;
  --bg-light: #e0eafc;
  --card-dark: #0e162b;
  --card-light: #ffffff;
  --muted-dark: #9ca3af;
  --muted-light: #526379;
  --text-dark: #e5e7eb;
  --text-light: #1f2937;
  --brand-dark: #7c3aed;
  --brand-2-dark: #22d3ee;
  --field-dark: #0b1222;
  --field-light: #f3f4f6;
  --border-dark: #24304a;
  --border-light: #d1d5db;
  --ok: #16a34a;
  --warn: #f59e0b;
  --danger: #ef4444;
  --shadow: 0 12px 40px rgba(0, 0, 0, .35);

  /* Set initial theme */
  --bg: var(--bg-dark);
  --card: var(--card-dark);
  --muted: var(--muted-dark);
  --text: var(--text-dark);
  --brand: var(--brand-dark);
  --brand-2: var(--brand-2-dark);
  --field: var(--field-dark);
  --border: var(--border-dark);

  transition: all 0.3s ease;
}

body.light-mode {
  --bg: var(--bg-light);
  --card: var(--card-light);
  --muted: var(--muted-light);
  --text: var(--text-light);
  --brand: #4f46e5;
  --brand-2: #3b82f6;
  --field: var(--field-light);
  --border: var(--border-light);
}

/* Background animation */
.animated-bg {
  position: fixed;
  top: 0;
  left: 0;
  width: 100%;
  height: 100%;
  overflow: hidden;
  z-index: -1;
}

.animated-bg div {
  position: absolute;
  background: var(--brand-2);
  border-radius: 50%;
  animation: glow 15s infinite;
  opacity: 0.2;
}
.animated-bg div:nth-child(1) { top: 20%; left: 10%; width: 150px; height: 150px; animation-delay: 0s; }
.animated-bg div:nth-child(2) { top: 60%; left: 80%; width: 200px; height: 200px; animation-delay: 5s; }
.animated-bg div:nth-child(3) { top: 80%; left: 30%; width: 100px; height: 100px; animation-delay: 10s; }
.animated-bg div:nth-child(4) { top: 40%; left: 45%; width: 120px; height: 120px; animation-delay: 2s; }
.animated-bg div:nth-child(5) { top: 70%; left: 50%; width: 180px; height: 180px; animation-delay: 8s; }
.animated-bg div:nth-child(6) { top: 10%; left: 60%; width: 90px; height: 90px; animation-delay: 12s; }
.animated-bg div:nth-child(7) { top: 90%; left: 10%; width: 160px; height: 160px; animation-delay: 4s; }

@keyframes glow {
  0% { transform: scale(1) translate(0, 0); opacity: 0.2; }
  50% { transform: scale(1.2) translate(50px, -50px); opacity: 0.3; }
  100% { transform: scale(1) translate(0, 0); opacity: 0.2; }
}

* { box-sizing: border-box; }
body {
  margin: 0;
  background: var(--bg);
  color: var(--text);
  font-family: Inter, ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Arial;
  transition: background 0.3s ease, color 0.3s ease;
}
.wrap { max-width: 1100px; margin: 40px auto; padding: 0 16px; }
.header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  margin-bottom: 16px;
}
.h1 {
  font-size: 28px;
  font-weight: 800;
  letter-spacing: .2px;
  background: linear-gradient(90deg, var(--brand), var(--brand-2));
  -webkit-background-clip: text;
  background-clip: text;
  color: transparent;
  animation: pulse 4s infinite ease-in-out;
}
.toggle-btn {
  background: var(--card);
  border: 1px solid var(--border);
  color: var(--text);
  padding: 8px 12px;
  border-radius: 999px;
  cursor: pointer;
  font-size: 14px;
  display: flex;
  align-items: center;
  gap: 8px;
  transition: all 0.3s ease;
}
.toggle-btn:hover {
  background: color-mix(in srgb, var(--card) 90%, var(--border));
  transform: scale(1.05);
}
.card {
  background: var(--card);
  border: 1px solid var(--border);
  border-radius: 16px;
  padding: 18px;
  box-shadow: var(--shadow);
  transition: background 0.3s ease, border-color 0.3s ease, box-shadow 0.3s ease;
}
.grid { display: grid; gap: 16px; }
.two { grid-template-columns: 1fr 1fr; }
label { font-weight: 600; margin-bottom: 8px; display: block; }

input[type=text], select {
  width: 100%;
  padding: 12px 14px;
  border-radius: 12px;
  border: 1px solid var(--border);
  background: var(--field);
  color: var(--text);
  transition: all 0.3s ease;
}
input[type=text]:focus, select:focus {
  outline: none;
  border-color: var(--brand);
  box-shadow: 0 0 0 2px color-mix(in srgb, var(--brand) 50%, transparent);
}

.file-input {
  display: flex;
  align-items: center;
  gap: 12px;
  cursor: pointer;
  transition: all 0.3s ease;
}
.file-btn {
  background: var(--brand);
  color: white;
  padding: 12px 16px;
  border-radius: 12px;
  font-weight: 700;
  cursor: pointer;
  white-space: nowrap;
  transition: all 0.3s ease;
  border: none;
  box-shadow: 0 4px 12px rgba(0,0,0,0.2);
}
.file-btn:hover {
  transform: translateY(-2px);
  box-shadow: 0 6px 16px rgba(0,0,0,0.3);
}
.file-name {
  color: var(--muted);
  font-size: 14px;
  overflow: hidden;
  white-space: nowrap;
  text-overflow: ellipsis;
  flex-grow: 1;
}

.hint { color: var(--muted); font-size: 12px; margin-top: 6px; }
.row { display: grid; gap: 12px; grid-template-columns: 1fr 1fr; }
.btns { display: flex; gap: 12px; margin-top: 16px; flex-wrap: wrap; }
.btn {
  cursor: pointer;
  border: 0;
  padding: 12px 16px;
  border-radius: 12px;
  font-weight: 700;
  background: var(--brand);
  color: white;
  transition: all 0.3s ease;
  box-shadow: 0 4px 12px rgba(0,0,0,0.2);
  min-width: 120px; /* Consistent button width */
}
.btn:hover {
  transform: translateY(-2px);
  box-shadow: 0 6px 16px rgba(0,0,0,0.3);
}
.btn.secondary { background: #1f2a44; color: var(--text); }
.btn.secondary:hover { background: #2c3e5e; }
.btn.ghost { background: transparent; border: 1px dashed var(--border); color: var(--text); }
.btn.ghost:hover { background: color-mix(in srgb, var(--bg) 90%, var(--border)); }

.flash { background: #1f2a44; border: 1px solid var(--warn); color: #fde68a; padding: 10px 12px; border-radius: 12px; margin: 12px 0; transition: all 0.3s ease; }
.kv { color: var(--muted); font-size: 13px; }
.dl { text-decoration: none; }
.table-wrap { overflow: auto; max-height: 65vh; border: 1px solid var(--border); border-radius: 12px; transition: all 0.3s ease; }
table { width: 100%; border-collapse: collapse; font-size: 13px; }
th, td { border-bottom: 1px solid var(--border); padding: 8px 10px; text-align: left; transition: all 0.3s ease; }
th { position: sticky; top: 0; background: var(--field); z-index: 1; }
th:last-child, td:last-child { width: 100%; }
td input.remark {
  width: 100%;
  min-width: 300px;
  padding: 8px 10px;
  background: var(--field);
  border: 1px solid var(--border);
  border-radius: 8px;
  color: var(--text);
  transition: all 0.3s ease;
}
td input.remark:focus { outline: none; border-color: var(--brand); box-shadow: 0 0 0 2px color-mix(in srgb, var(--brand) 50%, transparent); }

.badge {
  display: inline-block;
  padding: 4px 8px;
  border-radius: 999px;
  font-size: 12px;
  border: 1px solid var(--border);
  background: var(--field);
  transition: all 0.3s ease;
}

/* Spinner */
.spinner {
  display: none;
  border: 4px solid rgba(255, 255, 255, 0.3);
  border-top: 4px solid white;
  border-radius: 50%;
  width: 24px;
  height: 24px;
  animation: spin 1s linear infinite;
}
.btn.loading .spinner { display: block; }
.btn.loading {
  pointer-events: none;
  background: #3f51b5;
  color: rgba(255, 255, 255, 0.7);
}

@keyframes spin {
  0% { transform: rotate(0deg); }
  100% { transform: rotate(360deg); }
}
@keyframes pulse {
  0%, 100% { transform: scale(1); }
  50% { transform: scale(1.03); }
}
"""

INDEX_PAGE = """
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <title>Summary Enricher</title>
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <style>{{ css }}</style>
</head>
<body>
  <div class="animated-bg">
    <div></div>
    <div></div>
    <div></div>
    <div></div>
    <div></div>
    <div></div>
    <div></div>
  </div>

  <div class="wrap">
    <div class="header">
      <div class="h1">Summary Enricher</div>
      <button id="theme-toggle" class="toggle-btn" aria-label="Toggle light/dark theme">
        <svg id="moon-icon" xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="display:none;"><path d="M12 3a6 6 0 0 0 9 9 9 9 0 1 1-9-9Z"></path></svg>
        <svg id="sun-icon" xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="4"></circle><path d="M12 2v2"></path><path d="M12 20v2"></path><path d="m4.93 4.93 1.41 1.41"></path><path d="m17.66 17.66 1.41 1.41"></path><path d="M2 12h2"></path><path d="M20 12h2"></path><path d="m4.93 19.07 1.41-1.41"></path><path d="m17.66 6.34-1.41 1.41"></path></svg>
      </button>
    </div>

    {% with messages = get_flashed_messages() %}
      {% if messages %}
        {% for m in messages %}
          <div class="flash">{{ m }}</div>
        {% endfor %}
      {% endif %}
    {% endwith %}

    <form class="grid" method="post" enctype="multipart/form-data" action="{{ url_for('process_step1') }}">
      <div class="card">
        <div class="row">
          <div>
            <label>Usersetting file (.csv / .xlsx)</label>
            <div class="file-input">
              <label class="file-btn" for="usersetting-file">Choose File</label>
              <span class="file-name" id="usersetting-name">No file chosen</span>
              <input required type="file" name="usersetting" id="usersetting-file" hidden>
            </div>
            <div class="hint">Header auto-detected; we keep: User Alias, User ID, Max Loss, Telegram.</div>
          </div>
          <div>
            <label>Summary file (.xlsx recommended; multi-sheet supported)</label>
            <div class="file-input">
              <label class="file-btn" for="summary-file">Choose File</label>
              <span class="file-name" id="summary-name">No file chosen</span>
              <input required type="file" name="summary" id="summary-file" hidden>
            </div>
            <div class="hint">All sheets preserved; only the first sheet is enriched.</div>
          </div>
        </div>
      </div>

      <div class="card">
        <div class="row">
          <div>
            <label>ALGO</label>
            <input required type="text" name="ALGO" placeholder="e.g., 7">
          </div>
          <div>
            <label>OPERATOR</label>
            <input required type="text" name="OPERATOR" placeholder="e.g., GAURAVK">
          </div>
        </div>
        <div class="row">
          <div>
            <label>EXPIRY</label>
            <select name="EXPIRY" required>
              <option value="">-- Please Drop Down :) --</option>
              <option>NIFTY 1DTE</option>
              <option>NIFTY 0DTE</option>
              <option>SENSEX 1DTE</option>
              <option>SENSEX 0DTE</option>
              <option>BANKNIFTY 1DTE</option>
              <option>BANKNIFTY 0DTE</option>
            </select>
          </div>
          <div>
            <label>REMARK (optional)</label>
            <input type="text" name="REMARK" placeholder="Will be edited per-row next step">
          </div>
        </div>
        <div class="hint">SERVER will be auto-detected from the first word in your file name.</div>
      </div>

      <div class="btns">
        <button class="btn" type="submit" id="submit-btn">
          <span id="btn-text">Run</span>
          <div class="spinner"></div>
        </button>
        <button class="btn ghost" type="reset">Reset</button>
      </div>
    </form>
  </div>
  <script>
    document.getElementById('usersetting-file').addEventListener('change', function(e) {
      const fileName = e.target.files[0] ? e.target.files[0].name : 'No file chosen';
      document.getElementById('usersetting-name').textContent = fileName;
    });
    document.getElementById('summary-file').addEventListener('change', function(e) {
      const fileName = e.target.files[0] ? e.target.files[0].name : 'No file chosen';
      document.getElementById('summary-name').textContent = fileName;
    });

    const form = document.querySelector('form');
    const submitBtn = document.getElementById('submit-btn');

    form.addEventListener('submit', function() {
      submitBtn.classList.add('loading');
      submitBtn.querySelector('#btn-text').textContent = '';
    });
    
    // Theme toggle logic
    const toggleBtn = document.getElementById('theme-toggle');
    const moonIcon = document.getElementById('moon-icon');
    const sunIcon = document.getElementById('sun-icon');
    const body = document.body;

    toggleBtn.addEventListener('click', () => {
        body.classList.toggle('light-mode');
        const isLightMode = body.classList.contains('light-mode');
        localStorage.setItem('theme', isLightMode ? 'light' : 'dark');
        moonIcon.style.display = isLightMode ? 'block' : 'none';
        sunIcon.style.display = isLightMode ? 'none' : 'block';
    });

    // Apply saved theme on load
    const savedTheme = localStorage.getItem('theme');
    if (savedTheme === 'light') {
        body.classList.add('light-mode');
        moonIcon.style.display = 'block';
        sunIcon.style.display = 'none';
    } else {
        moonIcon.style.display = 'none';
        sunIcon.style.display = 'block';
    }
  </script>
</body>
</html>
"""

PREVIEW_PAGE = """
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <title>Preview & Edit Remarks</title>
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <style>{{ css }}</style>
</head>
<body class="light-mode">
  <div class="animated-bg">
    <div></div>
    <div></div>
    <div></div>
    <div></div>
    <div></div>
    <div></div>
    <div></div>
  </div>
  <div class="wrap">
    <div class="header">
      <div class="h1">Preview (Sheet 1)</div>
      <button id="theme-toggle" class="toggle-btn" aria-label="Toggle light/dark theme">
        <svg id="moon-icon" xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 3a6 6 0 0 0 9 9 9 9 0 1 1-9-9Z"></path></svg>
        <svg id="sun-icon" xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="display:none;"><circle cx="12" cy="12" r="4"></circle><path d="M12 2v2"></path><path d="M12 20v2"></path><path d="m4.93 4.93 1.41 1.41"></path><path d="m17.66 17.66 1.41 1.41"></path><path d="M2 12h2"></path><path d="M20 12h2"></path><path d="m4.93 19.07 1.41-1.41"></path><path d="m17.66 6.34-1.41 1.41"></path></svg>
      </button>
    </div>
    <div class="kv">
      Detected SERVER: <span class="badge">{{ server }}</span> •
      ALGO: <span class="badge">{{ algo }}</span> •
      OPERATOR: <span class="badge">{{ operator }}</span> •
      EXPIRY: <span class="badge">{{ expiry }}</span>
    </div>

    <div class="card" style="margin-bottom:16px; margin-top: 16px;">
      <div class="hint">Only the <b>REMARK</b> column is editable below. Scroll to review. When ready, click <b>Submit</b> to generate the final workbook.</div>
    </div>

    <form method="post" action="{{ url_for('process_step2') }}">
      <input type="hidden" name="key" value="{{ key }}">
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              {% for col in columns %}
                <th>{{ col }}</th>
              {% endfor %}
            </tr>
          </thead>
          <tbody>
            {% for item in rows %}
              <tr>
                {% for col in columns %}
                  {% if col == 'REMARK' %}
                    <td>
                      <input class="remark" type="text" name="remark_{{ item.idx }}" value="{{ item.row.get('REMARK','') }}">
                    </td>
                  {% else %}
                    <td>{{ item.row.get(col,'') }}</td>
                  {% endif %}
                {% endfor %}
              </tr>
            {% endfor %}
          </tbody>
        </table>
      </div>

      <div class="btns">
        <a class="btn secondary dl" href="{{ us_download }}">Download Cleaned Usersetting</a>
        <button class="btn" type="submit">Submit (Build Final Workbook)</button>
      </div>
    </form>
  </div>
  <script>
    // Theme toggle logic
    const toggleBtn = document.getElementById('theme-toggle');
    const moonIcon = document.getElementById('moon-icon');
    const sunIcon = document.getElementById('sun-icon');
    const body = document.body;

    const savedTheme = localStorage.getItem('theme');
    if (savedTheme === 'light') {
        body.classList.add('light-mode');
        moonIcon.style.display = 'block';
        sunIcon.style.display = 'none';
    } else {
        body.classList.remove('light-mode');
        moonIcon.style.display = 'none';
        sunIcon.style.display = 'block';
    }
    
    toggleBtn.addEventListener('click', () => {
        body.classList.toggle('light-mode');
        const isLightMode = body.classList.contains('light-mode');
        localStorage.setItem('theme', isLightMode ? 'light' : 'dark');
        moonIcon.style.display = isLightMode ? 'block' : 'none';
        sunIcon.style.display = isLightMode ? 'none' : 'block';
    });
  </script>
</body>
</html>
"""


FINAL_PAGE = """
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <title>Final Workbook Ready</title>
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <style>{{ css }}</style>
</head>
<body class="light-mode">
  <div class="animated-bg">
    <div></div>
    <div></div>
    <div></div>
    <div></div>
    <div></div>
    <div></div>
    <div></div>
  </div>
  <div class="wrap">
    <div class="header">
      <div class="h1">Done ✓</div>
      <button id="theme-toggle" class="toggle-btn" aria-label="Toggle light/dark theme">
        <svg id="moon-icon" xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 3a6 6 0 0 0 9 9 9 9 0 1 1-9-9Z"></path></svg>
        <svg id="sun-icon" xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="display:none;"><circle cx="12" cy="12" r="4"></circle><path d="M12 2v2"></path><path d="M12 20v2"></path><path d="m4.93 4.93 1.41 1.41"></path><path d="m17.66 17.66 1.41 1.41"></path><path d="M2 12h2"></path><path d="M20 12h2"></path><path d="m4.93 19.07 1.41-1.41"></path><path d="m17.66 6.34-1.41 1.41"></path></svg>
      </button>
    </div>

    <div class="card">
      <div class="kv" style="margin-bottom: 12px;">Your final Summary (all sheets) is ready.</div>
      <div class="btns">
        <a class="btn dl" href="{{ sm_download }}">Download Enriched Summary</a>
        <a class="btn secondary dl" href="{{ us_download }}">Download Cleaned Usersetting</a>
        <a class="btn ghost dl" href="{{ url_for('index') }}">Start Over</a>
      </div>
    </div>
  </div>
  <script>
    // Theme toggle logic
    const toggleBtn = document.getElementById('theme-toggle');
    const moonIcon = document.getElementById('moon-icon');
    const sunIcon = document.getElementById('sun-icon');
    const body = document.body;

    const savedTheme = localStorage.getItem('theme');
    if (savedTheme === 'light') {
        body.classList.add('light-mode');
        moonIcon.style.display = 'block';
        sunIcon.style.display = 'none';
    } else {
        body.classList.remove('light-mode');
        moonIcon.style.display = 'none';
        sunIcon.style.display = 'block';
    }
    
    toggleBtn.addEventListener('click', () => {
        body.classList.toggle('light-mode');
        const isLightMode = body.classList.contains('light-mode');
        localStorage.setItem('theme', isLightMode ? 'light' : 'dark');
        moonIcon.style.display = isLightMode ? 'block' : 'none';
        sunIcon.style.display = isLightMode ? 'none' : 'block';
    });
  </script>
</body>
</html>
"""

# -------------------- Helpers --------------------
CANONICAL_US = ["User Alias", "User ID", "Max Loss", "Telegram"]

def _norm(s: str) -> str:
    if s is None: return ""
    s = str(s)
    return "".join(ch for ch in s.strip().lower() if ch.isalnum())

SYNONYMS_US: Dict[str, str] = {
    _norm("User Alias"): "User Alias",
    _norm("User ID"): "User ID",
    _norm("Max Loss"): "Max Loss",
    _norm("Telegram"): "Telegram",
    _norm("Telegram ID"): "Telegram",
    _norm("Telegram IDs"): "Telegram",
    _norm("Telegram ID(s)"): "Telegram",
}

def _ext_ok(filename: str) -> bool:
    return os.path.splitext(filename or "")[1].lower() in ALLOWED_EXT

# Put this below _append_constants or with other helpers
DESIRED_ORDER = [
    "SNO","Enabled","UserID","Alias","LoggedIn","SqOff Done","Broker","Qty Multiplier",
    "MTM (All)","ALLOCATION","MAX_LOSS","Available Margin","Total Orders","Total Lots",
    "SERVER","ALGO","REMARK","OPERATOR","EXPIRY"
]

def _reorder_summary_columns(df: pd.DataFrame) -> pd.DataFrame:
    cols = list(df.columns)
    desired = [c for c in DESIRED_ORDER if c in cols]
    rest = [c for c in cols if c not in desired]
    # Keep only existing columns; append any extras at the end
    return df[desired + rest]


def _read_raw(file_storage) -> pd.DataFrame:
    """Read raw usersetting file (csv/xlsx) with fixed header row at index 6 (7th row)."""
    name = (file_storage.filename or "").lower()
    if name.endswith(".csv"):
        # Always skip first 6 junk rows, then use row 7 as header
        return pd.read_csv(
            file_storage,
            header=6,           # row index 6 = 7th row
            dtype=str,
            keep_default_na=False,
            low_memory=False
        )
    return pd.read_excel(file_storage, header=6, dtype=str)


def _select_usersetting_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Select and normalize the canonical usersetting columns."""
    norm_to_orig = {_norm(c): c for c in df.columns}
    resolved = {}
    for canonical in CANONICAL_US:
        candidates = [k for k, v in SYNONYMS_US.items() if v == canonical]
        found = None
        for cand in candidates:
            if cand in norm_to_orig:
                found = norm_to_orig[cand]
                break
        if not found:
            raise ValueError(f"Usersetting missing column: {canonical}")
        resolved[canonical] = found
    cleaned = df[[resolved[c] for c in CANONICAL_US]].copy()
    cleaned.columns = CANONICAL_US
    return cleaned


def _find_header_row(df: pd.DataFrame, scan_rows: int = 30) -> Optional[int]:
    targets = {_norm("Enabled"), _norm("User Alias"), _norm("User ID"), _norm("Max Loss")}
    for r in range(min(scan_rows, len(df))):
        row_vals = [_norm(x) for x in df.iloc[r].tolist()]
        if not any(row_vals): continue
        if sum(1 for t in targets if t in set(row_vals)) >= 3:
            return r
    return None

def _coerce_header(df: pd.DataFrame, header_row: int) -> pd.DataFrame:
    headers = df.iloc[header_row].tolist()
    fixed, last = [], ""
    for h in headers:
        h = str(h).strip()
        if h: last = h
        fixed.append(h or last or "")
    out = df.iloc[header_row + 1:].copy()
    out.columns = fixed
    return out

def _select_usersetting_columns(df: pd.DataFrame) -> pd.DataFrame:
    norm_to_orig = {_norm(c): c for c in df.columns}
    resolved = {}
    for canonical in CANONICAL_US:
        candidates = [k for k in SYNONYMS_US if SYNONYMS_US[k] == canonical]
        found = None
        for cand in candidates:
            if cand in norm_to_orig:
                found = norm_to_orig[cand]
                break
        if not found:
            raise ValueError(f"Usersetting missing column: {canonical}")
        resolved[canonical] = found
    cleaned = df[[resolved[c] for c in CANONICAL_US]].copy()
    cleaned.columns = CANONICAL_US
    return cleaned

def _build_lookup(clean_us: pd.DataFrame) -> Dict[str, Tuple[str, str]]:
    return {
        _norm(row["User ID"]): (row["Telegram"], row["Max Loss"])
        for _, row in clean_us.iterrows() if row.get("User ID")
    }

def _insert_allocation_maxloss(df: pd.DataFrame, lookup: Dict[str, Tuple[str, str]], user_id_colname: str) -> pd.DataFrame:
    out = df.copy()
    insert_at = min(9, len(out.columns))  # before column #10 (1-indexed)
    def fetch(uid):
        tel, mls = lookup.get(_norm(uid), ("", ""))
        return pd.Series({"ALLOCATION": tel, "MAX_LOSS": mls})
    new_cols = out[user_id_colname].apply(fetch)
    out.insert(insert_at, "ALLOCATION", new_cols["ALLOCATION"])
    out.insert(insert_at + 1, "MAX_LOSS", new_cols["MAX_LOSS"])
    return out

def _append_constants(df: pd.DataFrame, consts: Dict[str, str]) -> pd.DataFrame:
    out = df.copy()
    for k in ["SERVER","ALGO","OPERATOR","EXPIRY"]:
        out[k] = consts.get(k, "")
    # Add blank REMARK column (user edits in step 2)
    if "REMARK" not in out.columns:
        out["REMARK"] = ""
    return out

def _read_all_sheets(file_storage) -> Dict[str, pd.DataFrame]:
    name = (file_storage.filename or "").lower()
    if name.endswith(".csv"):
        return {"Sheet1": pd.read_csv(file_storage, low_memory=False)}
    xl = pd.ExcelFile(file_storage)
    return {sheet: xl.parse(sheet_name=sheet) for sheet in xl.sheet_names}

def _server_from_filename(name: str) -> str:
    base = os.path.splitext(name or "")[0].strip()
    token = base.replace("_"," ").replace("-"," ").split()
    return token[0] if token else ""

# -------------------- State --------------------
STORE: Dict[str, Dict] = {}   # temp memory across steps
DOWNLOADS: Dict[str, io.BytesIO] = {}

# -------------------- Routes --------------------
@app.route("/", methods=["GET"])
def index():
    return render_template_string(INDEX_PAGE, css=BASE_CSS)

@app.route("/process", methods=["POST"])
def process_step1():
    try:
        us = request.files["usersetting"]
        sm = request.files["summary"]
        if not _ext_ok(us.filename) or not _ext_ok(sm.filename):
            flash("Unsupported file type."); return redirect(url_for("index"))

        # constants (no REMARK here, it's per-row in next step)
        consts = {
            "ALGO": request.form.get("ALGO","").strip(),
            "OPERATOR": request.form.get("OPERATOR","").strip(),
            "EXPIRY": request.form.get("EXPIRY","").strip(),
        }
        # auto SERVER
        consts["SERVER"] = _server_from_filename(us.filename) or _server_from_filename(sm.filename)

        # --- Clean Usersetting (and make download) ---
        raw_us = _read_raw(us)
        us_clean = _select_usersetting_columns(raw_us)

        us_buf = io.BytesIO()
        with pd.ExcelWriter(us_buf, engine="openpyxl") as xw:
            us_clean.to_excel(xw, index=False, sheet_name="Usersetting")
        us_buf.seek(0)
        us_key = "US_" + datetime.now().strftime("%H%M%S%f")
        DOWNLOADS[us_key] = us_buf

        # --- Prepare Summary (only sheet 1 enriched now; others stored to write later) ---
        sheets = _read_all_sheets(sm)
        names = list(sheets.keys())
        first_name = names[0]
        first_df = sheets[first_name]

        # build lookup & enrich 1st sheet (ALLOCATION / MAX_LOSS + constants + REMARK blank)
        lookup = _build_lookup(us_clean)
        uid_col = "UserID" if "UserID" in first_df.columns else ("User ID" if "User ID" in first_df.columns else None)
        enriched_first = first_df.copy()
        if uid_col:
            enriched_first = _insert_allocation_maxloss(enriched_first, lookup, uid_col)
        else:
            insert_at = min(9, len(enriched_first.columns))
            enriched_first.insert(insert_at, "ALLOCATION", "")
            enriched_first.insert(insert_at + 1, "MAX_LOSS", "")
        enriched_first = _append_constants(enriched_first, consts)
        enriched_first = _reorder_summary_columns(enriched_first)


        # Cap preview rows (rendering huge tables in HTML is slow)
        preview_limit = 1000
        preview_first = enriched_first.head(preview_limit)

        # Save to STORE for step 2
        key = "JOB_" + datetime.now().strftime("%H%M%S%f")
        STORE[key] = {
            "consts": consts,
            "usersetting_df": us_clean,
            "summary_sheets": sheets,        # originals (dict name -> df)
            "first_sheet_name": first_name,
            "enriched_first_full": enriched_first,  # full df inc. blank REMARK
            "us_download_key": us_key,
        }

        # Build preview rows WITH indices for Jinja
        columns = list(preview_first.columns)
        records = preview_first.fillna("").astype(str).to_dict(orient="records")
        rows = [{"idx": i, "row": rec} for i, rec in enumerate(records)]


        return render_template_string(
            PREVIEW_PAGE,
            css=BASE_CSS,
            key=key,
            server=consts["SERVER"],
            algo=consts["ALGO"],
            operator=consts["OPERATOR"],
            expiry=consts["EXPIRY"],
            columns=columns,
            rows=rows,
            us_download=url_for("download", key=us_key),
        )
    except Exception as e:
        flash(f"Error: {e}")
        return redirect(url_for("index"))

@app.route("/finalize", methods=["POST"])
def process_step2():
    try:
        key = request.form.get("key","")
        job = STORE.get(key)
        if not job:
            flash("Session expired. Please run again.")
            return redirect(url_for("index"))

        consts = job["consts"]
        us_key = job["us_download_key"]
        sheets = job["summary_sheets"]
        first_name = job["first_sheet_name"]
        enriched_first = job["enriched_first_full"].copy()

        # Update REMARK column from form fields (remark_<rowIndex>_<any>)
        # We’ll map row index to remark text for visible range; if the sheet is larger than preview,
        # missing remarks remain as initially blank.
        # Update REMARK column from form fields (remark_<rowIndex>)
        remarks = {}
        for form_key, value in request.form.items():
            if form_key.startswith("remark_"):
                try:
                    idx = int(form_key.split("_")[1])
                    remarks[idx] = value.strip()
                except:
                    pass


        # Ensure REMARK column exists
        if "REMARK" not in enriched_first.columns:
            enriched_first["REMARK"] = ""

        # Apply remarks to matching rows (by integer position)
        # Only rows visible in preview were posted; others left as-is.
        for idx, text in remarks.items():
            if 0 <= idx < len(enriched_first):
                enriched_first.iat[idx, enriched_first.columns.get_loc("REMARK")] = text

        # Write all sheets: first = enriched_first w/ REMARKs, others unchanged
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine="openpyxl") as xw:
            # Ensure final column order (SNO..ALGO, REMARK, OPERATOR, EXPIRY)
            enriched_first = _reorder_summary_columns(enriched_first)

            # keep original sheet order
            for name in sheets.keys():
                if name == first_name:
                    enriched_first.to_excel(xw, index=False, sheet_name=name[:31])
                else:
                    sheets[name].to_excel(xw, index=False, sheet_name=name[:31])
        out.seek(0)
        sm_key = "SM_" + datetime.now().strftime("%H%M%S%f")
        DOWNLOADS[sm_key] = out

        # cleanup job
        STORE.pop(key, None)

        return render_template_string(
            FINAL_PAGE,
            css=BASE_CSS,
            sm_download=url_for("download", key=sm_key),
            us_download=url_for("download", key=us_key),
        )
    except Exception as e:
        flash(f"Error: {e}")
        return redirect(url_for("index"))

@app.route("/download/<key>")
def download(key):
    buf = DOWNLOADS.get(key)
    if not buf:
        flash("Download expired. Please re-upload.")
        return redirect(url_for("index"))
    return send_file(
        buf,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=f"Output_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
    )

if __name__ == "__main__":
    app.add_url_rule("/process", view_func=process_step1, methods=["POST"])   # alias clarity
    app.add_url_rule("/finalize", view_func=process_step2, methods=["POST"])
    app.run(debug=True, port=5000)
