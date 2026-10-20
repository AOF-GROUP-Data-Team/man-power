# ============================================================
# Branch Manpower Report — Shawarma Classic (567430) + Garatis (671657)
# GitHub Actions version. Latest submission per branch -> cleaned manpower
# -> xlsx -> emailed as attachment.
# Secrets: ZENPUT_API_TOKEN, GMAIL_APP_PASSWORD
# ============================================================
import os, sys, smtplib, requests, re, ast
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.utils import formataddr
from email import encoders
import pytz
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

API_TOKEN     = os.environ.get("ZENPUT_API_TOKEN", "")
GMAIL_USER    = "aof.group.auto@gmail.com"
GMAIL_PASS    = os.environ.get("GMAIL_APP_PASSWORD", "")
SENDER_NAME   = "Business Intelligence"
TO_EMAILS     = ["o.salahaddin@aofgroup.com"]
CC_EMAILS     = ["a.alsalem@aofgroup.com"]
SEND_EMAIL    = os.environ.get("SEND_EMAIL", "true").lower() == "true"
BASE          = "https://www.zenput.com/api/v3/submissions/"
HEADERS       = {"X-API-TOKEN": API_TOKEN}
TZ            = pytz.timezone("Asia/Riyadh")
PAGE          = 100
MAX_PAGES     = 200     # safety cap: 200 x 100 = 20,000 submissions per template
LINK_BASE     = "https://ops-execution.crunchtime.com/reports/#form_id/"

TEMPLATES = {567430: "Shawarma Classic", 671657: "Garatis"}

# Expected branches: brand -> {code: label}
CLASSIC = {
    "B01": "NURUH", "B02": "KHRUH", "B03": "GHRUH", "B04": "NSRUH", "B05": "RWRUH",
    "B06": "DARUH", "B07": "LBRUH", "B08": "SWRUH", "B09": "AZRUH", "B10": "SHRUH",
    "B11": "NRRUH", "B12": "TWRUH", "B13": "AQRUH", "B14": "RBRUH", "B15": "NDRUH",
    "B16": "BDRUH", "B17": "QRRUH", "B18": "TKRUH", "B19": "MURUH", "B21": "KRRUH",
    "B22": "OBJED", "B23": "SLAHS", "B24": "SFJED", "B25": "RWAHS", "B26": "HAJED",
    "B27": "SARUH", "B28": "MAJED", "B30": "QARUH", "B31": "ANRUH", "B32": "FYJED",
    "B33": "HIRJED", "B34": "URRUH", "B35": "IRRUH", "B36": "PSJED", "B37": "SHWMAK",
    "B38": "UHDMM", "B39": "HSRUH", "B40": "MZDMM",
}
GARATIS = {
    "QB01": "السويدي", "QB02": "النرجس", "QB03": "اليرموك", "QB04": "العقيق",
    "QB05": "الخافجي", "QB06": "الخليج", "QB07": "الندوة", "QB08": "المنصورة", "QB09": "النزهة",
}
GARATIS_LOC_IDS = {
    "2232755": "QB01", "2235805": "QB02", "2254072": "QB03", "2256386": "QB04",
    "2268360": "QB05", "2270650": "QB06", "2274188": "QB07", "2276794": "QB08",
}

PLACEHOLDER_NAMES = {"no", "n", "na", "none", "nil", "null", "x", "xx", "nothing", "empty",
                     "vacant", "no name", "no employee", "not available", "not availble",
                     "not avilable", "no one", "nobody"}

# ---------------- helpers ----------------
def parse_dt(s):
    if not s: return None
    try: return datetime.fromisoformat(s.replace("Z", "+00:00")).astimezone(TZ)
    except Exception: return None

def text_of(v):
    if v is None: return ""
    if isinstance(v, list): return ", ".join(map(str, v))
    if isinstance(v, str) and v.strip().startswith("["):
        try:
            lst = ast.literal_eval(v)
            if isinstance(lst, list): return ", ".join(map(str, lst))
        except Exception: pass
    return str(v).strip()

def name_compact(s):
    return re.sub(r"[^a-z\u0600-\u06ff]", "", (s or "").lower())

def is_valid_name(s):
    n = re.sub(r"\s+", " ", re.sub(r"[^a-z\u0600-\u06ff ]", "", (s or "").lower())).strip()
    if len(n.replace(" ", "")) < 2 or n in PLACEHOLDER_NAMES: return False
    if n.startswith("not avail") or n.startswith("no cashier"): return False
    return True

def id_key(s):
    return re.sub(r"\D", "", s or "").lstrip("0") or None

def is_yes(v):
    return str(v).strip().lower() in ("true", "yes", "1")

def resolve_branch(brand, loc):
    name = re.sub(r"\s+", " ", (loc.get("name") or loc.get("display_name") or "")).strip()
    if brand == "Garatis":
        code = GARATIS_LOC_IDS.get(str(loc.get("id") or ""))
        if not code:
            m = re.search(r"\bQB\s?(\d{2})\b", name, re.I) or re.search(r"Garatis+\s*B(\d{2})", name, re.I)
            code = f"QB{m.group(1)}" if m else None
        return code if code in GARATIS else None
    m = re.search(r"\bB(\d{2})\b", name)
    code = f"B{m.group(1)}" if m else None
    if code in CLASSIC and name.upper().startswith(CLASSIC[code]):
        return code
    return None

# ---------------- fetch ----------------
def fetch(template_id):
    out, start = [], 0
    for _ in range(MAX_PAGES):
        r = requests.get(BASE, headers=HEADERS, params={
            "form_template_id": template_id, "limit": PAGE, "start": start}, timeout=60)
        print(f"[{template_id}] GET start={start} -> HTTP {r.status_code}")
        if r.status_code != 200:
            print(r.text[:500]); break
        data = r.json().get("data", [])
        if not data: break
        out.extend(data)
        if len(data) < PAGE: break
        start += PAGE
    return out

# ---------------- manpower count ----------------
def count_manpower(s):
    emps, cur = [], None
    cashier_flag, cashier_name = None, ""
    for a in s.get("answers", []):
        tl = (a.get("title") or "").strip().lower()
        v = a.get("value")
        if tl.startswith("write ") and tl.endswith("employee"):
            cur = {"name": text_of(v), "id": ""}
            emps.append(cur)
        elif tl == "id number" and cur is not None and not cur["id"]:
            cur["id"] = text_of(v)
        elif tl.startswith("this number of employees with cashier"):
            cashier_flag = is_yes(v)
        elif tl == "write name of cashier":
            cashier_name = text_of(v)
    seen_ids, seen_names, valid = set(), set(), 0
    for e in emps:
        if not is_valid_name(e["name"]): continue
        ik, nk = id_key(e["id"]), name_compact(e["name"])
        if (ik and ik in seen_ids) or (not ik and nk in seen_names): continue
        valid += 1
        if ik: seen_ids.add(ik)
        seen_names.add(nk)
    if cashier_flag is False and is_valid_name(cashier_name) and name_compact(cashier_name) not in seen_names:
        valid += 1
    return valid, len(emps)

# ---------------- run ----------------
if not API_TOKEN:
    sys.exit("ZENPUT_API_TOKEN is not set")
rows = []
for tid, brand in TEMPLATES.items():
    subs = fetch(tid)
    print(f"[{tid}] {brand}: {len(subs)} submissions")
    latest = {}
    for s in subs:
        sm = s.get("smetadata") or {}
        code = resolve_branch(brand, sm.get("location") or {})
        dt = parse_dt(sm.get("date_submitted"))
        if not code or not dt:
            continue
        if code not in latest or dt > latest[code][0]:
            latest[code] = (dt, s)
    expected = CLASSIC if brand == "Shawarma Classic" else GARATIS
    structure_warned = False
    for code in sorted(expected, key=lambda c: int(re.sub(r"\D", "", c))):
        label = f"{code} {expected[code]}"
        if code not in latest:
            rows.append([brand, label, None, "No submission", None])
            continue
        dt, s = latest[code]
        total, slots = count_manpower(s)
        if slots == 0 and not structure_warned:
            structure_warned = True
            print(f"\n!! [{tid}] no employee fields recognised in {label} — answer titles:")
            for a in s.get("answers", []):
                print(f"   {a.get('title')} / {a.get('field_type')} / {text_of(a.get('value'))[:60]}")
        rows.append([brand, label, total, dt.strftime("%Y-%m-%d %H:%M"),
                     LINK_BASE + str(s.get("legacy_submission_id") or "")])

# ---------------- xlsx ----------------
wb = Workbook()
ws = wb.active
ws.title = "Manpower"
ws.append(["Brand", "Branch", "Total Manpower", "Last Submitted", "Submission Link"])
thin = Border(*(Side(style="thin", color="BFBFBF"),) * 4)
for r in rows:
    ws.append(r[:4] + ["Open" if r[4] else ""])
    row = ws[ws.max_row]
    if r[4]:
        row[4].hyperlink, row[4].font = r[4], Font(color="0563C1", underline="single")
    else:
        for c in row: c.fill = PatternFill("solid", fgColor="D9D9D9")
ws.append([])
for brand in TEMPLATES.values():
    ws.append([brand, "TOTAL", sum(r[2] or 0 for r in rows if r[0] == brand), "", ""])
    for c in ws[ws.max_row]: c.font = Font(bold=True)
ws.append(["ALL", "TOTAL", sum(r[2] or 0 for r in rows), "", ""])
for c in ws[ws.max_row]: c.font = Font(bold=True)

for c in ws[1]:
    c.fill, c.font = PatternFill("solid", fgColor="1F3864"), Font(bold=True, color="FFFFFF")
for row in ws.iter_rows():
    for c in row:
        c.alignment = Alignment(horizontal="center", vertical="center")
        if c.value is not None: c.border = thin
for col, w in zip("ABCDE", [18, 20, 15, 18, 16]):
    ws.column_dimensions[col].width = w
ws.freeze_panes = "A2"

now = datetime.now(TZ)
fname = f"manpower_report_{now.strftime('%Y-%m-%d')}.xlsx"
wb.save(fname)
print(f"\nSaved {fname}")
for brand in TEMPLATES.values():
    br = [r for r in rows if r[0] == brand]
    print(f"{brand}: {sum(1 for r in br if r[4])}/{len(br)} branches submitted, "
          f"total manpower {sum(r[2] or 0 for r in br)}")


# ---------------- email ----------------
def build_html():
    lines = []
    for brand in TEMPLATES.values():
        br = [r for r in rows if r[0] == brand]
        missing = [r[1] for r in br if not r[4]]
        lines.append(
            f"<tr><td style='padding:6px 12px;border:1px solid #ccc'>{brand}</td>"
            f"<td style='padding:6px 12px;border:1px solid #ccc;text-align:center'>{sum(1 for r in br if r[4])} / {len(br)}</td>"
            f"<td style='padding:6px 12px;border:1px solid #ccc;text-align:center'><b>{sum(r[2] or 0 for r in br)}</b></td>"
            f"<td style='padding:6px 12px;border:1px solid #ccc'>{', '.join(missing) if missing else '-'}</td></tr>")
    return f"""<html><body style="font-family:Arial,sans-serif;font-size:14px;color:#222">
<p>Dear Team,</p>
<p>Please find attached the branch manpower report generated on {now.strftime('%Y-%m-%d %H:%M')} (KSA),
based on the latest submission of each branch.</p>
<table style="border-collapse:collapse">
<tr style="background:#1F3864;color:#fff">
<th style="padding:6px 12px">Brand</th><th style="padding:6px 12px">Branches Submitted</th>
<th style="padding:6px 12px">Total Manpower</th><th style="padding:6px 12px">No Submission</th></tr>
{''.join(lines)}
<tr><td style='padding:6px 12px;border:1px solid #ccc'><b>All</b></td><td style='border:1px solid #ccc'></td>
<td style='padding:6px 12px;border:1px solid #ccc;text-align:center'><b>{sum(r[2] or 0 for r in rows)}</b></td>
<td style='border:1px solid #ccc'></td></tr>
</table>
<p>Best regards,<br>Business Intelligence</p>
</body></html>"""

def send_email(path):
    if not GMAIL_PASS:
        sys.exit("GMAIL_APP_PASSWORD is not set")
    msg = MIMEMultipart()
    msg["From"] = formataddr((SENDER_NAME, GMAIL_USER))
    msg["To"] = ", ".join(TO_EMAILS)
    msg["Cc"] = ", ".join(CC_EMAILS)
    msg["Subject"] = f"Branch Manpower Report - {now.strftime('%Y-%m-%d')}"
    msg.attach(MIMEText(build_html(), "html", "utf-8"))
    with open(path, "rb") as f:
        part = MIMEBase("application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        part.set_payload(f.read())
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f'attachment; filename="{os.path.basename(path)}"')
    msg.attach(part)
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as s:
        s.login(GMAIL_USER, GMAIL_PASS)
        s.sendmail(GMAIL_USER, TO_EMAILS + CC_EMAILS, msg.as_string())
    print(f"Email sent to {TO_EMAILS} cc {CC_EMAILS}")

if SEND_EMAIL:
    send_email(fname)
else:
    print("SEND_EMAIL=false -> email skipped")
