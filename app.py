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
BASE_CSS = """
:root {
  --bg:#0b1020; --card:#0e162b; --muted:#9aa4b2; --text:#e7ebf3;
  --brand:#7c3aed; --brand-2:#22d3ee; --brand-3:#f59e0b;
  --field:#0b1222; --border:#22314d;
  --ok:#22c55e; --warn:#f59e0b; --danger:#ef4444; --shadow:0 15px 45px rgba(0,0,0,.35);
  --accent:#60a5fa;
}

:root[data-theme="light"] {
  --bg:#f7f8fb; --card:#ffffff; --muted:#6b7280; --text:#0f172a;
  --brand:#6d28d9; --brand-2:#06b6d4; --brand-3:#ea580c;
  --field:#f3f4f6; --border:#e5e7eb;
  --ok:#16a34a; --warn:#b45309; --danger:#dc2626; --shadow:0 10px 30px rgba(2,6,23,.08);
  --accent:#2563eb;
}

/* Respect system preference on first load; we’ll persist with localStorage */
@media (prefers-color-scheme: light) {
  :root:not([data-theme="dark"]) {
    --bg:#f7f8fb; --card:#ffffff; --muted:#6b7280; --text:#0f172a;
    --brand:#6d28d9; --brand-2:#06b6d4; --brand-3:#ea580c;
    --field:#f3f4f6; --border:#e5e7eb; --shadow:0 10px 30px rgba(2,6,23,.08);
    --accent:#2563eb;
  }
}

* { box-sizing: border-box; }
html, body { height:100%; }
body { margin:0; background:var(--bg); color:var(--text); font-family: Inter, ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Arial; }

.wrap { max-width: 1200px; margin: 40px auto; padding: 0 18px; }

.header {
  display:flex; align-items:center; justify-content:space-between; margin-bottom:20px;
  animation: fadeIn .5s ease both;
}

.h1 {
  font-size: 30px; font-weight:900; letter-spacing:.25px;
  background: conic-gradient(from 90deg, var(--brand), var(--brand-2), var(--brand-3), var(--brand));
  background-size: 200% 200%;
  -webkit-background-clip:text; background-clip:text; color:transparent;
  animation: shine 6s linear infinite;
}

.kv { color:var(--muted); font-size:13px; }

.card {
  background: linear-gradient(180deg, color-mix(in srgb, var(--card), transparent 0%), color-mix(in srgb, var(--card), transparent 6%));
  border: 1px solid var(--border);
  border-radius: 16px; padding: 20px; box-shadow: var(--shadow);
  animation: floatIn .4s ease both;
}

.grid { display:grid; gap:16px; }
.two { grid-template-columns: 1fr 1fr; }
@media (max-width: 900px) { .two { grid-template-columns: 1fr; } }

label { font-weight:700; margin-bottom:8px; display:block; }

input[type=file], input[type=text], select {
  width:100%; padding:12px 14px; border-radius:12px; border:1px solid var(--border);
  background:var(--field); color:var(--text); outline: none;
  transition: box-shadow .2s ease, transform .05s ease, border-color .2s ease, background .2s ease;
}

input[type=file]::file-selector-button {
  border:0; padding:10px 12px; border-radius:10px; margin-right:10px; cursor:pointer;
  background: linear-gradient(90deg, var(--brand), var(--brand-2));
  color:white; font-weight:700;
}

input[type=text]:focus, select:focus {
  box-shadow: 0 0 0 4px color-mix(in srgb, var(--accent), transparent 80%);
  border-color: color-mix(in srgb, var(--accent), transparent 10%);
  background: color-mix(in srgb, var(--field), white 4%);
}

.hint { color:var(--muted); font-size:12px; margin-top:6px; }

.row { display:grid; gap:12px; grid-template-columns: 1fr 1fr; }
@media (max-width: 900px) { .row { grid-template-columns: 1fr; } }

.btns { display:flex; gap:12px; margin-top:18px; flex-wrap:wrap; }

.btn {
  cursor:pointer; border:0; padding:12px 16px; border-radius:12px; font-weight:800;
  background: linear-gradient(90deg, var(--brand), var(--brand-2));
  color:white; letter-spacing:.2px; transition: transform .05s ease, box-shadow .2s ease, opacity .2s ease;
  box-shadow: 0 10px 20px rgba(0,0,0,.15);
}
.btn:hover { transform: translateY(-1px); box-shadow: 0 14px 28px rgba(0,0,0,.18); }
.btn:active { transform: translateY(0); }

.btn.secondary {
  background: linear-gradient(90deg, color-mix(in srgb, var(--accent), white 8%), color-mix(in srgb, var(--brand), white 18%));
  color:white;
}
.btn.ghost {
  background:transparent; border:1px dashed var(--border); color:var(--text);
}

.flash {
  background: color-mix(in srgb, var(--warn), var(--card) 75%);
  border:1px solid color-mix(in srgb, var(--warn), black 10%);
  color:#111; padding:10px 12px; border-radius:12px; margin:12px 0;
}

.badge { display:inline-block; padding:4px 10px; border-radius:999px; font-size:12px; border:1px solid var(--border); background:var(--field); }

.table-wrap { overflow:auto; max-height: 70vh; border:1px solid var(--border); border-radius:14px; }

table { width:100%; border-collapse: collapse; font-size: 13px; }
th, td { border-bottom:1px solid color-mix(in srgb, var(--border), transparent 30%); padding:10px 12px; text-align:left; }
th {
  position: sticky; top:0; background:linear-gradient(180deg, color-mix(in srgb, var(--field), white 0%), color-mix(in srgb, var(--field), transparent 40%));
  z-index:1; font-weight:800; font-size:12px; color: var(--muted);
  backdrop-filter: blur(3px);
}
tbody tr { transition: background .15s ease; }
tbody tr:hover { background: color-mix(in srgb, var(--field), white 6%); }

td input.remark {
  width: 100%;
  min-width: 340px;
  padding: 10px 12px; border-radius:10px; border:1px solid var(--border);
  background:var(--field); color:var(--text); outline:none; font-size:13px;
}
td input.remark:focus {
  box-shadow: 0 0 0 4px color-mix(in srgb, var(--brand-2), transparent 80%);
  border-color: color-mix(in srgb, var(--brand-2), transparent 20%);
}

.header-actions { display:flex; align-items:center; gap:10px; }

/* Pretty theme toggle switch */
.toggle {
  display:inline-flex; align-items:center; gap:8px; user-select:none; cursor:pointer;
  padding:6px 10px; border-radius:999px; border:1px solid var(--border); background:var(--field);
  font-size:12px; color:var(--muted);
}
.toggle .dot {
  width:22px; height:22px; border-radius:999px; background:linear-gradient(90deg, var(--brand), var(--brand-2));
  box-shadow: inset 0 0 0 2px rgba(255,255,255,.25);
}

/* Animations */
@keyframes shine {
  0%   { background-position: 0% 50%; }
  100% { background-position: 200% 50%; }
}
@keyframes fadeIn {
  from { opacity:0; transform: translateY(6px); }
  to   { opacity:1; transform: translateY(0); }
}
@keyframes floatIn {
  from { opacity:0; transform: translateY(8px) scale(.98); }
  to   { opacity:1; transform: translateY(0) scale(1); }
}
"""

INDEX_PAGE = """
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <title>Usersetting Cleaner + Summary Enricher</title>
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <style>{{ css }}</style>
  <script>
    // Theme bootstrap (no FOUC)
    (function() {
      try {
        var t = localStorage.getItem('theme');
        if (t) document.documentElement.setAttribute('data-theme', t);
      } catch(e) {}
    })();
  </script>
</head>
<body>
  <div class="wrap">
    <div class="header">
      <div class="h1">Summary Enricher</div>
      <div class="header-actions">
        <div class="kv">SERVER auto-detected • Only sheet #1 is enriched</div>
        <button type="button" class="toggle" id="themeToggle" title="Toggle theme">
          <span class="dot"></span><span id="themeLabel">Theme</span>
        </button>
      </div>
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
            <input required type="file" name="usersetting" accept=".csv,.xlsx,.xls">
            <div class="hint">Header auto-detected; we keep: User Alias, User ID, Max Loss, Telegram.</div>
          </div>
          <div>
            <label>Summary file (.xlsx recommended; multi-sheet supported)</label>
            <input required type="file" name="summary" accept=".csv,.xlsx,.xls">
            <div class="hint">All sheets preserved; only the first sheet is enriched.</div>
          </div>
        </div>
      </div>

      <div class="card">
        <div class="row">
          <div>
            <label>ALGO</label>
            <input required type="text" name="ALGO" placeholder="e.g., A8">
          </div>
          <div>
            <label>OPERATOR</label>
            <input required type="text" name="OPERATOR" placeholder="e.g., Jay">
          </div>
        </div>
        <div class="row">
          <div>
            <label>EXPIRY</label>
            <select name="EXPIRY" required>
              <option value="">-- Select Expiry --</option>
              <option>NIFTY 1DTE</option>
              <option>NIFTY 0DTE</option>
              <option>SENSEX 1DTE</option>
              <option>SENSEX 0DTE</option>
              <option>BANKNIFTY 1DTE</option>
              <option>BANKNIFTY 0DTE</option>
            </select>
            <div class="hint">SERVER auto-detected from the first word in your file name (e.g., VS11).</div>
          </div>
          <div>
            <label>REMARK (optional — will be edited next step)</label>
            <input type="text" name="REMARK" placeholder="Leave empty here; you'll edit per-row next" disabled>
            <div class="hint">Per-row remarks are editable in the preview screen.</div>
          </div>
        </div>
      </div>

      <div class="btns">
        <button class="btn" type="submit">Run</button>
        <button class="btn ghost" type="reset">Reset</button>
      </div>
    </form>
  </div>

  <script>
    (function(){
      const root = document.documentElement;
      const btn = document.getElementById('themeToggle');
      const label = document.getElementById('themeLabel');

      function currentTheme(){
        return root.getAttribute('data-theme') || 'dark';
      }
      function setTheme(t){
        root.setAttribute('data-theme', t);
        try { localStorage.setItem('theme', t); } catch(e){}
        label.textContent = t === 'dark' ? 'Dark' : 'Light';
      }
      // Initialize label
      setTheme(currentTheme());

      btn.addEventListener('click', function(){
        const next = currentTheme() === 'dark' ? 'light' : 'dark';
        setTheme(next);
      });
    })();
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
  <script>
    (function(){ try { var t = localStorage.getItem('theme'); if(t) document.documentElement.setAttribute('data-theme', t);} catch(e){} })();
  </script>
</head>
<body>
  <div class="wrap">
    <div class="header">
      <div class="h1">Preview (Sheet 1) — Edit Remark Only</div>
      <div class="header-actions">
        <div class="kv">
          Detected SERVER: <span class="badge">{{ server }}</span> •
          ALGO: <span class="badge">{{ algo }}</span> •
          OPERATOR: <span class="badge">{{ operator }}</span> •
          EXPIRY: <span class="badge">{{ expiry }}</span>
        </div>
        <button type="button" class="toggle" id="themeToggle" title="Toggle theme">
          <span class="dot"></span><span id="themeLabel">Theme</span>
        </button>
      </div>
    </div>

    <div class="card" style="margin-bottom:16px;">
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
                      <input class="remark" type="text" name="remark_{{ item.idx }}" value="{{ item.row.get('REMARK','') }}" placeholder="Add a note for this row...">
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
    (function(){
      const root = document.documentElement;
      const btn = document.getElementById('themeToggle');
      const label = document.getElementById('themeLabel');

      function currentTheme(){ return root.getAttribute('data-theme') || 'dark'; }
      function setTheme(t){
        root.setAttribute('data-theme', t);
        try { localStorage.setItem('theme', t); } catch(e){}
        label.textContent = t === 'dark' ? 'Dark' : 'Light';
      }
      setTheme(currentTheme());
      btn.addEventListener('click', function(){
        const next = currentTheme() === 'dark' ? 'light' : 'dark';
        setTheme(next);
      });
    })();
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
  <script>
    (function(){ try { var t = localStorage.getItem('theme'); if(t) document.documentElement.setAttribute('data-theme', t);} catch(e){} })();
  </script>
</head>
<body>
  <div class="wrap">
    <div class="header">
      <div class="h1">Done ✓</div>
      <div class="header-actions">
        <div class="kv">Your final Summary (all sheets) is ready.</div>
        <button type="button" class="toggle" id="themeToggle" title="Toggle theme">
          <span class="dot"></span><span id="themeLabel">Theme</span>
        </button>
      </div>
    </div>

    <div class="card">
      <div class="btns">
        <a class="btn dl" href="{{ sm_download }}">Download Enriched Summary</a>
        <a class="btn secondary dl" href="{{ us_download }}">Download Cleaned Usersetting</a>
        <a class="btn ghost dl" href="{{ url_for('index') }}">Start Over</a>
      </div>
    </div>
  </div>

  <script>
    (function(){
      const root = document.documentElement;
      const btn = document.getElementById('themeToggle');
      const label = document.getElementById('themeLabel');

      function currentTheme(){ return root.getAttribute('data-theme') || 'dark'; }
      function setTheme(t){
        root.setAttribute('data-theme', t);
        try { localStorage.setItem('theme', t); } catch(e){}
        label.textContent = t === 'dark' ? 'Dark' : 'Light';
      }
      setTheme(currentTheme());
      btn.addEventListener('click', function(){
        const next = currentTheme() === 'dark' ? 'light' : 'dark';
        setTheme(next);
      });
    })();
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
    return df[desired + rest]

def _read_raw(file_storage) -> pd.DataFrame:
    """Read raw usersetting file (csv/xlsx) with fixed header row at index 6 (7th row)."""
    name = (file_storage.filename or "").lower()
    if name.endswith(".csv"):
        return pd.read_csv(
            file_storage,
            header=6,
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

        consts = {
            "ALGO": request.form.get("ALGO","").strip(),
            "OPERATOR": request.form.get("OPERATOR","").strip(),
            "EXPIRY": request.form.get("EXPIRY","").strip(),
        }
        consts["SERVER"] = _server_from_filename(us.filename) or _server_from_filename(sm.filename)

        raw_us = _read_raw(us)
        us_clean = _select_usersetting_columns(raw_us)

        us_buf = io.BytesIO()
        with pd.ExcelWriter(us_buf, engine="openpyxl") as xw:
            us_clean.to_excel(xw, index=False, sheet_name="Usersetting")
        us_buf.seek(0)
        us_key = "US_" + datetime.now().strftime("%H%M%S%f")
        DOWNLOADS[us_key] = us_buf

        sheets = _read_all_sheets(sm)
        names = list(sheets.keys())
        first_name = names[0]
        first_df = sheets[first_name]

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

        preview_limit = 1000
        preview_first = enriched_first.head(preview_limit)

        key = "JOB_" + datetime.now().strftime("%H%M%S%f")
        STORE[key] = {
            "consts": consts,
            "usersetting_df": us_clean,
            "summary_sheets": sheets,
            "first_sheet_name": first_name,
            "enriched_first_full": enriched_first,
            "us_download_key": us_key,
        }

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

        # collect remarks
        remarks = {}
        for form_key, value in request.form.items():
            if form_key.startswith("remark_"):
                try:
                    idx = int(form_key.split("_")[1])
                    remarks[idx] = value.strip()
                except:
                    pass

        if "REMARK" not in enriched_first.columns:
            enriched_first["REMARK"] = ""

        for idx, text in remarks.items():
            if 0 <= idx < len(enriched_first):
                enriched_first.iat[idx, enriched_first.columns.get_loc("REMARK")] = text

        out = io.BytesIO()
        with pd.ExcelWriter(out, engine="openpyxl") as xw:
            enriched_first = _reorder_summary_columns(enriched_first)
            for name in sheets.keys():
                if name == first_name:
                    enriched_first.to_excel(xw, index=False, sheet_name=name[:31])
                else:
                    sheets[name].to_excel(xw, index=False, sheet_name=name[:31])
        out.seek(0)
        sm_key = "SM_" + datetime.now().strftime("%H%M%S%f")
        DOWNLOADS[sm_key] = out

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
