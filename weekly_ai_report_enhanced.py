"""
Weekly AI Company Analysis — Gemini + Google Search  (v3 — correct SDK)
Uses: pip install google-genai pandas openpyxl
"""

import os, sys, re, time
from datetime import datetime
from google import genai
from google.genai import types

# ============================================================================
# CONFIGURATION
# ============================================================================

GEMINI_API_KEY  = os.environ.get("GEMINI_API_KEY", "")
SEND_EMAIL      = os.environ.get("SEND_EMAIL", "false").lower() == "true"
SMTP_SERVER     = os.environ.get("SMTP_SERVER", "smtp.gmail.com")
SMTP_PORT       = int(os.environ.get("SMTP_PORT", "587"))
EMAIL_FROM      = os.environ.get("EMAIL_FROM", "")
EMAIL_PASSWORD  = os.environ.get("EMAIL_PASSWORD", "")
EMAIL_TO        = os.environ.get("EMAIL_TO", "")

GEMINI_MODEL       = "gemini-2.0-flash"   # Fast, current, supports Google Search
COMPANIES_PER_WEEK = 10                     # Increase to 10 once results look good
REQUEST_DELAY      = 10                     # Seconds between companies

# ============================================================================
# STARTUP CHECKS
# ============================================================================

def check_environment():
    print("🔎 Checking environment...")
    errors = []

    if not GEMINI_API_KEY:
        errors.append("❌ GEMINI_API_KEY missing from GitHub Secrets")
    else:
        print(f"   ✅ GEMINI_API_KEY found ({GEMINI_API_KEY[:8]}...)")

    if SEND_EMAIL:
        for var, val in [("EMAIL_FROM", EMAIL_FROM),
                         ("EMAIL_PASSWORD", EMAIL_PASSWORD),
                         ("EMAIL_TO", EMAIL_TO)]:
            if not val:
                errors.append(f"❌ {var} missing from GitHub Secrets")
            else:
                print(f"   ✅ {var} found")

    try:
        from google import genai as _g
        print(f"   ✅ google-genai installed")
    except ImportError:
        errors.append("❌ google-genai not installed — workflow pip install must include: google-genai")

    try:
        import pandas, openpyxl
        print(f"   ✅ pandas + openpyxl installed")
    except ImportError as e:
        errors.append(f"❌ Missing library: {e}")

    if errors:
        print("\n" + "="*70)
        print("STARTUP FAILED — fix these before running:")
        for e in errors:
            print(f"  {e}")
        print("="*70)
        sys.exit(1)

    print("   ✅ All checks passed\n")

# ============================================================================
# DATA STRUCTURE
# ============================================================================

class CompanyAnalysis:
    def __init__(self, name):
        self.name = name
        self.situations = {
            1: {"name": "Resource Constraints",    "score": 0, "points": [], "sources": []},
            2: {"name": "Supply Chain Disruption", "score": 0, "points": [], "sources": []},
            3: {"name": "Margin Pressure",         "score": 0, "points": [], "sources": []},
            4: {"name": "Significant Growth",      "score": 0, "points": [], "sources": []},
        }

# ============================================================================
# GEMINI API  (new google.genai SDK)
# ============================================================================

def get_gemini_response(prompt_text):
    """
    Call Gemini API with Google Search grounding.
    Uses the NEW google.genai SDK (pip install google-genai).
    """
    client = genai.Client(api_key=GEMINI_API_KEY)

    config = types.GenerateContentConfig(
        tools=[types.Tool(google_search=types.GoogleSearch())],
        temperature=0.3,
        max_output_tokens=8192,
    )

    print(f"   📡 Sending to Gemini ({GEMINI_MODEL}) with Google Search...")
    response = client.models.generate_content(
        model=GEMINI_MODEL,
        contents=prompt_text,
        config=config,
    )

    if not response.candidates:
        raise Exception("Gemini returned no candidates — safety block or empty response")

    candidate = response.candidates[0]
    finish = str(candidate.finish_reason)
    if "SAFETY" in finish or "RECITATION" in finish:
        raise Exception(f"Gemini blocked response: finish_reason={finish}")

    content = response.text
    if not content or len(content.strip()) < 50:
        raise Exception(f"Gemini returned near-empty response ({len(content)} chars): '{content[:100]}'")

    # Log which searches Gemini ran
    try:
        meta = candidate.grounding_metadata
        if meta:
            if hasattr(meta, 'web_search_queries') and meta.web_search_queries:
                print(f"   🔍 Searches run: {list(meta.web_search_queries)}")
            if hasattr(meta, 'grounding_chunks') and meta.grounding_chunks:
                print(f"   🔗 Sources used: {len(meta.grounding_chunks)}")
    except Exception:
        pass

    return content

# ============================================================================
# PARSER
# ============================================================================

def parse_response(response_text):
    companies = []

    print(f"\n   Response length : {len(response_text)} chars")
    print(f"   Preview         : {repr(response_text[:300])}")

    if "---COMPANY START---" not in response_text:
        print("⚠️  No ---COMPANY START--- delimiter found.")
        print("   Gemini did not follow the output format.")
        print("   Check the raw .txt artifact to see what it returned.")
        return []

    blocks = re.split(r'---COMPANY START---', response_text)
    print(f"   Found {len(blocks)-1} company block(s)")

    for idx, block in enumerate(blocks[1:], 1):
        if '---COMPANY END---' in block:
            block = block.split('---COMPANY END---')[0]

        m = re.search(r'Company:\s*(.+?)(?:\n|$)', block, re.IGNORECASE)
        if not m:
            print(f"   ⚠️  Block {idx}: no 'Company:' line")
            continue

        company = CompanyAnalysis(m.group(1).strip())
        print(f"   ✅ Parsing: {company.name}")
        found = 0

        for sit_num in range(1, 5):
            pat = (
                rf'SITUATION\s+{sit_num}[:\s].*?\n'
                rf'Score:\s*(\d+).*?\n'
                rf'(?:Key\s+(?:Signals|Points)[^\n]*\n)(.*?)'
                rf'(?:Evidence\s+Links|Sources)[^\n]*\n(.*?)'
                rf'(?=\n\s*SITUATION\s+{sit_num+1}|\n---COMPANY END---|$)'
            )
            match = re.search(pat, block, re.DOTALL | re.IGNORECASE)

            if match:
                score       = int(match.group(1))
                points_raw  = match.group(2)
                sources_raw = match.group(3)

                points = []
                for line in points_raw.split('\n'):
                    line = re.sub(r'^[\s\-\•\*]+', '', line).strip()
                    if line and len(line) > 10:
                        points.append(line)

                sources = []
                for line in sources_raw.split('\n'):
                    url = re.search(r'https?://\S+', line)
                    if url:
                        sources.append(url.group(0).rstrip('.,)'))

                company.situations[sit_num]["score"]   = score
                company.situations[sit_num]["points"]  = points[:3]
                company.situations[sit_num]["sources"] = sources[:3]
                found += 1
            else:
                print(f"      ⚠️  Situation {sit_num}: regex did not match")

        print(f"      → {found}/4 situations parsed")
        companies.append(company)

    print(f"\n   ✅ Parsed {len(companies)} companies total")
    return companies

# ============================================================================
# EXCEL
# ============================================================================

def create_excel_report(companies, week_num, output_path):
    import pandas as pd
    from openpyxl import load_workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

    data = []
    for co in companies:
        for sit_num in range(1, 5):
            s = co.situations[sit_num]
            data.append({
                'Company':     co.name,
                'Situation':   s['name'],
                'Score':       s['score'],
                'Key Point 1': s['points'][0] if len(s['points']) > 0 else '',
                'Key Point 2': s['points'][1] if len(s['points']) > 1 else '',
                'Key Point 3': s['points'][2] if len(s['points']) > 2 else '',
                'Sources':     ' | '.join(s['sources']),
            })

    import pandas as pd
    pd.DataFrame(data).to_excel(output_path, index=False, sheet_name='Company Analysis')

    wb = load_workbook(output_path)
    ws = wb.active

    hdr   = PatternFill("solid", fgColor="366092")
    green = PatternFill("solid", fgColor="C6EFCE")
    yel   = PatternFill("solid", fgColor="FFEB9C")
    red   = PatternFill("solid", fgColor="FFC7CE")
    bdr   = Border(left=Side(style='thin'), right=Side(style='thin'),
                   top=Side(style='thin'),  bottom=Side(style='thin'))

    for cell in ws[1]:
        cell.fill      = hdr
        cell.font      = Font(name='Arial', size=11, bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border    = bdr

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.font      = Font(name='Arial', size=10)
            cell.alignment = Alignment(vertical='top', wrap_text=True)
            cell.border    = bdr
            if cell.column == 3 and isinstance(cell.value, int):
                cell.fill = green if cell.value <= 2 else (yel if cell.value == 3 else red)
            if cell.column == 7 and cell.value:
                u = re.search(r'https?://\S+', str(cell.value))
                if u:
                    cell.hyperlink = u.group(0)
                    cell.font = Font(name='Arial', size=10, color="0563C1", underline="single")

    for col, w in zip('ABCDEFG', [22, 26, 8, 42, 42, 42, 55]):
        ws.column_dimensions[col].width = w
    for row in ws.iter_rows():
        ws.row_dimensions[row[0].row].height = 30 if row[0].row == 1 else 65

    ws.freeze_panes = 'A2'
    wb.save(output_path)
    print(f"   ✅ Excel: {output_path}  ({len(data)} rows)")

# ============================================================================
# HTML EMAIL
# ============================================================================

def generate_html_email(companies, week_num):
    rows = ""
    for co in companies:
        for sit_num in range(1, 5):
            s = co.situations[sit_num]
            sc   = s['score']
            col  = "#27ae60" if sc <= 2 else ("#f39c12" if sc == 3 else "#e74c3c")
            bg   = "#d5f4e6" if sc <= 2 else ("#fff3cd" if sc == 3 else "#f8d7da")
            pts  = "".join(f"<li style='margin-bottom:4px'>{p}</li>" for p in s['points'])
            srcs = " | ".join(
                f'<a href="{u}" target="_blank" style="color:#3498db">Source {i}</a>'
                for i, u in enumerate(s['sources'][:3], 1) if u.startswith('http')
            )
            rows += f"""
            <tr>
              <td style="padding:10px;border:1px solid #ddd;font-weight:bold">{co.name}</td>
              <td style="padding:10px;border:1px solid #ddd;color:#7f8c8d;font-size:.9em">{s['name']}</td>
              <td style="padding:10px;border:1px solid #ddd;text-align:center">
                <span style="background:{bg};color:{col};padding:4px 10px;border-radius:4px;font-weight:bold">{sc}</span>
              </td>
              <td style="padding:10px;border:1px solid #ddd"><ul style="margin:0;padding-left:18px">{pts}</ul></td>
              <td style="padding:10px;border:1px solid #ddd;font-size:.85em">{srcs}</td>
            </tr>"""

    return f"""<!DOCTYPE html>
<html><head><meta charset="utf-8"></head>
<body style="font-family:Arial,sans-serif;background:#f5f5f5;padding:20px">
<div style="max-width:1200px;margin:0 auto;background:white;padding:30px;border-radius:8px">
  <h1 style="color:#2c3e50;border-bottom:3px solid #3498db;padding-bottom:10px">📊 Weekly Company Analysis</h1>
  <div style="background:#ecf0f1;padding:15px;border-radius:5px;margin-bottom:20px">
    <strong>Week {week_num}</strong> | {datetime.now().strftime('%B %d, %Y at %H:%M UTC')} | Companies: {len(companies)}
  </div>
  <table style="width:100%;border-collapse:collapse">
    <thead><tr style="background:#34495e;color:white">
      <th style="padding:12px;width:15%;text-align:left">Company</th>
      <th style="padding:12px;width:18%;text-align:left">Situation</th>
      <th style="padding:12px;width:7%;text-align:left">Score</th>
      <th style="padding:12px;width:45%;text-align:left">Key Evidence</th>
      <th style="padding:12px;width:15%;text-align:left">Sources</th>
    </tr></thead>
    <tbody>{rows}</tbody>
  </table>
  <div style="margin-top:30px;text-align:center;color:#7f8c8d;font-size:.9em">
    Auto-generated | 1–2 Low Risk · 3 Moderate · 4–5 High Risk
  </div>
</div></body></html>"""

# ============================================================================
# EMAIL SENDER
# ============================================================================

def send_html_email(subject, html, excel_path=None):
    if not SEND_EMAIL:
        print("📧 SEND_EMAIL not true — skipping email")
        return
    import smtplib
    from email.mime.text import MIMEText
    from email.mime.multipart import MIMEMultipart
    from email.mime.application import MIMEApplication

    msg = MIMEMultipart()
    msg['From'], msg['To'], msg['Subject'] = EMAIL_FROM, EMAIL_TO, subject
    msg.attach(MIMEText(html, 'html'))
    if excel_path and os.path.exists(excel_path):
        with open(excel_path, 'rb') as f:
            att = MIMEApplication(f.read(), _subtype="xlsx")
            att.add_header('Content-Disposition', 'attachment', filename=os.path.basename(excel_path))
            msg.attach(att)
    srv = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
    srv.starttls()
    srv.login(EMAIL_FROM, EMAIL_PASSWORD)
    srv.sendmail(EMAIL_FROM, EMAIL_TO, msg.as_string())
    srv.quit()
    print(f"   ✅ Email sent → {EMAIL_TO}")

# ============================================================================
# HELPERS
# ============================================================================

def get_companies_for_week(companies, week_num, per_week):
    total_weeks = max(1, (len(companies) + per_week - 1) // per_week)
    week_num    = ((week_num - 1) % total_weeks) + 1
    start       = (week_num - 1) * per_week
    return companies[start:start + per_week], week_num

def calculate_current_week(start_str, per_week, total):
    weeks       = (datetime.now() - datetime.strptime(start_str, "%Y-%m-%d")).days // 7
    total_weeks = max(1, (total + per_week - 1) // per_week)
    return (weeks % total_weeks) + 1

def save_raw(content, week_num):
    ts    = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    fname = f"report_week{week_num}_{ts}.txt"
    with open(fname, "w", encoding="utf-8") as f:
        f.write(content)
    print(f"   💾 Raw saved: {fname}  ({len(content)} chars)")
    return fname

# ============================================================================
# MAIN
# ============================================================================

def main():
    print("=" * 70)
    print("🤖  WEEKLY COMPANY ANALYSIS — GEMINI + GOOGLE SEARCH  (v3)")
    print("=" * 70)
    print(f"Started : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"Model   : {GEMINI_MODEL}")
    print(f"Per week: {COMPANIES_PER_WEEK}\n")

    check_environment()

    # Load files
    print("📂 Loading files...")
    for fname in ["prompt_updated.txt", "companies.txt", "definitions.txt"]:
        if not os.path.exists(fname):
            print(f"❌ Missing required file: {fname}")
            sys.exit(1)

    role_objective = open("prompt_updated.txt",  encoding="utf-8").read().strip()
    all_companies  = [l.strip() for l in open("companies.txt", encoding="utf-8") if l.strip()]
    scoring_defs   = open("definitions.txt", encoding="utf-8").read().strip()
    print(f"   ✅ Loaded {len(all_companies)} companies")

    # Week
    week_num = int(os.environ.get("WEEK", "0"))
    if week_num == 0:
        week_num = calculate_current_week(
            os.environ.get("START_DATE", "2026-02-17"),
            COMPANIES_PER_WEEK, len(all_companies)
        )
    companies_this_week, week_num = get_companies_for_week(all_companies, week_num, COMPANIES_PER_WEEK)

    print(f"\n📋 Week {week_num} — {len(companies_this_week)} companies:")
    for i, c in enumerate(companies_this_week, 1):
        print(f"   {i}. {c}")
    print()

    # Analyse
    raw_result, successful, failed = "", 0, []

    for i, company in enumerate(companies_this_week, 1):
        print(f"[{i}/{len(companies_this_week)}] 🔎 {company}")
        prompt = (
            f"{role_objective}\n\n"
            f"Definitions of Situations and Scoring:\n{scoring_defs}\n\n"
            f"Company to Analyze:\n{company}\n"
        )
        try:
            resp = get_gemini_response(prompt)
            raw_result += resp + "\n\n"
            successful += 1
            print(f"   ✅ {len(resp)} characters received")
            if i < len(companies_this_week):
                time.sleep(REQUEST_DELAY)
        except Exception as e:
            print(f"   ❌ FAILED: {e}")
            failed.append(company)

    print(f"\n✅ Successful: {successful}/{len(companies_this_week)}")
    if failed:
        print(f"❌ Failed    : {failed}")

    save_raw(raw_result, week_num)

    if successful == 0:
        print("\n❌ CRITICAL: 0 companies analysed. Check errors above.")
        sys.exit(1)

    # Parse
    print("\n🔄 Parsing...")
    companies_data = parse_response(raw_result)
    if not companies_data:
        print("❌ Parsing returned 0 companies.")
        print("   Download the raw .txt artifact and share it — I will fix the parser.")

    # Outputs
    ts    = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    xlsx  = f"analysis_week{week_num}_{ts}.xlsx"

    print(f"\n📊 Generating Excel...")
    create_excel_report(companies_data, week_num, xlsx)

    print("📧 Generating HTML...")
    html = generate_html_email(companies_data, week_num)

    if SEND_EMAIL:
        send_html_email(
            f"[AUTO-REPORT] 📊 Weekly Company Analysis - Week {week_num}",
            html, xlsx
        )
    else:
        preview = f"email_preview_week{week_num}_{ts}.html"
        open(preview, "w", encoding="utf-8").write(html)
        print(f"   💾 HTML preview: {preview}")

    print("\n" + "="*70)
    print("✅ DONE")
    print("="*70)
    print(f"💰 Estimated cost: ~${successful * 0.007:.3f}  ({successful} companies)")
    return 0

if __name__ == "__main__":
    sys.exit(main())
