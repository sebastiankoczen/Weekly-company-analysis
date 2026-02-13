"""
Weekly AI Company Analysis Report - Enhanced Version
Runs every Monday to analyze companies and send formatted results via email
"""

import requests
import os
import sys
import re
from datetime import datetime

# ============================================================================
# CONFIGURATION
# ============================================================================

PERPLEXITY_API_KEY = os.environ.get("PERPLEXITY_API_KEY", "")
SEND_EMAIL = os.environ.get("SEND_EMAIL", "false").lower() == "true"
SMTP_SERVER = os.environ.get("SMTP_SERVER", "smtp.gmail.com")
SMTP_PORT = int(os.environ.get("SMTP_PORT", "587"))
EMAIL_FROM = os.environ.get("EMAIL_FROM", "")
EMAIL_PASSWORD = os.environ.get("EMAIL_PASSWORD", "")
EMAIL_TO = os.environ.get("EMAIL_TO", "")

API_URL = "https://api.perplexity.ai/chat/completions"
PERPLEXITY_MODEL = "sonar-pro"  # MUST BE PRO FOR DEEP RESEARCH
COMPANIES_PER_WEEK = 3  # Processing 5 companies for highest quality

# ============================================================================
# DATA STRUCTURES
# ============================================================================

class CompanyAnalysis:
    """Holds analysis data for one company"""
    def __init__(self, name):
        self.name = name
        self.situations = {
            1: {"name": "Resource Constraints", "score": 0, "points": [], "sources": []},
            2: {"name": "Supply Chain Disruption", "score": 0, "points": [], "sources": []},
            3: {"name": "Margin Pressure", "score": 0, "points": [], "sources": []},
            4: {"name": "Significant Growth", "score": 0, "points": [], "sources": []}
        }

# ============================================================================
# PARSING & URL EXTRACTION
# ============================================================================

def parse_perplexity_response(response_text):
    """Parse the structured response from Perplexity"""
    companies = []
    
    if "---COMPANY START---" not in response_text:
        print("⚠️  WARNING: No '---COMPANY START---' delimiter found!")
        return []
    
    company_blocks = re.split(r'---COMPANY START---', response_text)
    
    for idx, block in enumerate(company_blocks[1:], 1):
        if '---COMPANY END---' in block:
            block = block.split('---COMPANY END---')[0]
        
        company_match = re.search(r'Company:\s*(.+?)(?:\n|$)', block, re.IGNORECASE)
        if not company_match:
            continue
            
        company_name = company_match.group(1).strip()
        company = CompanyAnalysis(company_name)
        
        situations_found = 0
        for sit_num in range(1, 5):
            patterns = [
                rf'SITUATION {sit_num}:.*?\nScore:\s*(\d+)\s*\nKey Points:\s*\n(.*?)\nSources:\s*(.+?)(?=\n\nSITUATION|\n---COMPANY END---|$)',
                rf'SITUATION {sit_num}[:\s]+.*?\nScore[:\s]+(\d+)\s*\n.*?Key Points[:\s]+\n(.*?)\nSources[:\s]+(.+?)(?=\n\nSITUATION|\n---COMPANY END---|$)'
            ]
            
            match = None
            for pattern in patterns:
                match = re.search(pattern, block, re.DOTALL | re.IGNORECASE)
                if match:
                    break
            
            if match:
                score = int(match.group(1))
                points_text = match.group(2).strip()
                sources_text = match.group(3).strip()
                
                points = []
                for line in points_text.split('\n'):
                    line = line.strip()
                    if line.startswith('-') or line.startswith('•') or line.startswith('*'):
                        point = re.sub(r'^[-•*]\s*', '', line)
                        if point:
                            points.append(point)
                
                # Extract the real URLs
                sources = [s.strip().strip('[]') for s in re.split(r'[|\n]', sources_text) if s.strip() and s.strip().startswith('http')]
                
                company.situations[sit_num]["score"] = score
                company.situations[sit_num]["points"] = points[:3]
                company.situations[sit_num]["sources"] = sources[:2]
                
                situations_found += 1
        
        if situations_found > 0:
            companies.append(company)
            
    return companies

# ============================================================================
# EXCEL GENERATION
# ============================================================================

def create_excel_report(companies, week_num, output_path):
    """Create an Excel file with formatted analysis"""
    import pandas as pd
    from openpyxl import load_workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    
    data = []
    if not companies:
        data.append({'Company': 'No data parsed', 'Situation': 'Error', 'Score': 0, 'Key Point 1': '', 'Key Point 2': '', 'Key Point 3': '', 'Sources': ''})
    else:
        for company in companies:
            for sit_num in range(1, 5):
                situation = company.situations[sit_num]
                row = {
                    'Company': company.name,
                    'Situation': situation['name'],
                    'Score': situation['score'],
                    'Key Point 1': situation['points'][0] if len(situation['points']) > 0 else '',
                    'Key Point 2': situation['points'][1] if len(situation['points']) > 1 else '',
                    'Key Point 3': situation['points'][2] if len(situation['points']) > 2 else '',
                    'Sources': ' | '.join(situation['sources']) if situation['sources'] else ''
                }
                data.append(row)
    
    df = pd.DataFrame(data)
    df.to_excel(output_path, index=False, sheet_name='Company Analysis')
    
    wb = load_workbook(output_path)
    ws = wb.active
    
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    header_font = Font(name='Arial', size=11, bold=True, color="FFFFFF")
    
    score_fill_high = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    score_fill_med = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
    score_fill_low = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    
    cell_font = Font(name='Arial', size=10)
    cell_alignment = Alignment(vertical='top', wrap_text=True)
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = thin_border
    
    for row_idx in range(2, ws.max_row + 1):
        for col_idx in range(1, ws.max_column + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.font = cell_font
            cell.alignment = cell_alignment
            cell.border = thin_border
            
            if col_idx == 3:  # Score
                score = cell.value
                if isinstance(score, int):
                    if score in [1, 2]: cell.fill = score_fill_high
                    elif score == 3: cell.fill = score_fill_med
                    elif score in [4, 5]: cell.fill = score_fill_low
            
            if col_idx == 7:  # Sources
                sources_text = cell.value
                if sources_text:
                    urls = [s.strip().strip('[]') for s in sources_text.split('|')]
                    for url in urls:
                        if url.startswith('http'):
                            cell.hyperlink = url
                            cell.font = Font(name='Arial', size=10, color="0563C1", underline="single")
                            break
    
    ws.column_dimensions['A'].width = 20
    ws.column_dimensions['B'].width = 25
    ws.column_dimensions['C'].width = 8
    ws.column_dimensions['D'].width = 40
    ws.column_dimensions['E'].width = 40
    ws.column_dimensions['F'].width = 40
    ws.column_dimensions['G'].width = 50
    
    ws.row_dimensions[1].height = 30
    for row_idx in range(2, ws.max_row + 1):
        ws.row_dimensions[row_idx].height = 60
    
    ws.freeze_panes = 'A2'
    wb.save(output_path)
    return True

# ============================================================================
# HTML EMAIL GENERATION
# ============================================================================

def generate_html_email(companies, week_num):
    """Generate HTML email"""
    if not companies:
        return "<html><body><h1>No data</h1></body></html>"
        
    html = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <style>
            body {{ font-family: Arial, sans-serif; background-color: #f5f5f5; padding: 20px; }}
            .container {{ max-width: 1200px; margin: 0 auto; background-color: white; padding: 30px; border-radius: 8px; }}
            h1 {{ color: #2c3e50; border-bottom: 3px solid #3498db; padding-bottom: 10px; }}
            table {{ width: 100%; border-collapse: collapse; margin-top: 20px; }}
            th {{ background-color: #34495e; color: white; padding: 12px; text-align: left; }}
            td {{ padding: 12px; border-bottom: 1px solid #ddd; vertical-align: top; }}
            .score-1, .score-2 {{ background-color: #d5f4e6; color: #27ae60; padding: 4px 8px; border-radius: 4px; font-weight: bold; }}
            .score-3 {{ background-color: #fff3cd; color: #f39c12; padding: 4px 8px; border-radius: 4px; font-weight: bold; }}
            .score-4, .score-5 {{ background-color: #f8d7da; color: #e74c3c; padding: 4px 8px; border-radius: 4px; font-weight: bold; }}
            .key-points {{ margin: 0; padding-left: 20px; }}
        </style>
    </head>
    <body>
        <div class="container">
            <h1>📊 Weekly Company Analysis Report</h1>
            <table>
                <thead>
                    <tr>
                        <th style="width: 15%;">Company</th>
                        <th style="width: 15%;">Situation</th>
                        <th style="width: 8%;">Score</th>
                        <th style="width: 47%;">Key Evidence Points</th>
                        <th style="width: 15%;">Sources</th>
                    </tr>
                </thead>
                <tbody>
    """
    for company in companies:
        for sit_num in range(1, 5):
            situation = company.situations[sit_num]
            score = situation['score']
            
            points_html = "<ul class='key-points'>"
            for point in situation['points']:
                points_html += f"<li>{point}</li>"
            points_html += "</ul>"
            
            sources_html = ""
            if situation['sources']:
                sources_links = []
                for idx, source in enumerate(situation['sources'][:2], 1):
                    if source.startswith('http'):
                        sources_links.append(f'<a href="{source}" target="_blank">Source {idx}</a>')
                sources_html = ' | '.join(sources_links)
            
            html += f"""
                    <tr>
                        <td><strong>{company.name}</strong></td>
                        <td>{situation['name']}</td>
                        <td><span class="score-{score}">{score}</span></td>
                        <td>{points_html}</td>
                        <td>{sources_html}</td>
                    </tr>
            """
    html += "</tbody></table></div></body></html>"
    return html

# ============================================================================
# API FUNCTIONS (WITH PERFECT URL MAPPING)
# ============================================================================

def get_perplexity_response(prompt_text):
    """Send prompt to Perplexity API with strict citation replacement"""
    if not PERPLEXITY_API_KEY:
        raise ValueError("PERPLEXITY_API_KEY not found")
    
    headers = {
        "Authorization": f"Bearer {PERPLEXITY_API_KEY}",
        "Content-Type": "application/json"
    }
    
    body = {
        "model": PERPLEXITY_MODEL,
        "messages": [{"role": "user", "content": prompt_text}],
        "max_tokens": 8000,
        "stream": False,
        "temperature": 0.2,
        "top_p": 0.9
    }
    
    print(f"Sending request to Perplexity API (model: {PERPLEXITY_MODEL})...")
    response = requests.post(API_URL, json=body, headers=headers, timeout=600)
    
    if response.status_code != 200:
        raise Exception(f"Perplexity API Error: {response.text}")
    
    result = response.json()
    content = result["choices"][0]["message"]["content"]
    
    # Grab the real deep links from Perplexity's hidden list
    citations = result.get("citations", [])
    
    # Swap the brackets backwards (so [10] doesn't get messed up by [1])
    for i in range(len(citations)-1, -1, -1):
        url = citations[i]
        citation_marker = f"[{i+1}]" 
        content = content.replace(citation_marker, url)
        
    return content

def send_html_email(subject, html_content, excel_path=None):
    """Send email with HTML and Excel attachment"""
    if not SEND_EMAIL:
        return
    import smtplib
    from email.mime.text import MIMEText
    from email.mime.multipart import MIMEMultipart
    from email.mime.application import MIMEApplication
    
    msg = MIMEMultipart()
    msg['From'] = EMAIL_FROM
    msg['To'] = EMAIL_TO
    msg['Subject'] = subject
    msg.attach(MIMEText(html_content, 'html'))
    
    if excel_path and os.path.exists(excel_path):
        with open(excel_path, 'rb') as f:
            excel_attachment = MIMEApplication(f.read(), _subtype="xlsx")
            excel_attachment.add_header('Content-Disposition', 'attachment', filename=os.path.basename(excel_path))
            msg.attach(excel_attachment)
    
    server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
    server.starttls()
    server.login(EMAIL_FROM, EMAIL_PASSWORD)
    server.sendmail(EMAIL_FROM, EMAIL_TO, msg.as_string())
    server.quit()
    print(f"✅ Email sent to {EMAIL_TO}")

def get_companies_for_week(companies, week_num, per_week=10):
    idx_start = (week_num - 1) * per_week
    return companies[idx_start:idx_start + per_week]

def calculate_current_week(start_date, companies_per_week, total_companies):
    start = datetime.strptime(start_date, "%Y-%m-%d")
    days_diff = (datetime.now() - start).days
    weeks_passed = days_diff // 7
    total_weeks = (total_companies + companies_per_week - 1) // companies_per_week
    return (weeks_passed % total_weeks) + 1

def save_to_file(content, week_num):
    filename = f"report_week{week_num}_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.txt"
    with open(filename, "w", encoding="utf-8") as f:
        f.write(content)

# ============================================================================
# MAIN EXECUTION (ITERATIVE LOOP)
# ============================================================================

def main():
    try:
        with open("prompt_updated.txt", "r", encoding="utf-8") as f: role_objective = f.read().strip()
        with open("companies.txt", "r", encoding="utf-8") as f: all_companies = [line.strip() for line in f if line.strip()]
        with open("definitions.txt", "r", encoding="utf-8") as f: scoring_defs = f.read().strip()
        
        week_num = int(os.environ.get("WEEK", "0"))
        if week_num == 0:
            week_num = calculate_current_week(os.environ.get("START_DATE", "2025-02-17"), COMPANIES_PER_WEEK, len(all_companies))
            
        companies_this_week = get_companies_for_week(all_companies, week_num, COMPANIES_PER_WEEK)
        
        print(f"\n🔍 Requesting deep analysis from Perplexity AI (One-by-One)...")
        raw_result = ""
        
        # This loop forces the AI to search ONE company at a time so URLs NEVER mix!
        for i, company_name in enumerate(companies_this_week, 1):
            print(f"  [{i}/{len(companies_this_week)}] Researching {company_name}...")
            
            single_prompt = (
                f"{role_objective}\n\n"
                f"Definitions of Situations and Scoring:\n{scoring_defs}\n\n"
                f"Company to Analyze:\n- {company_name}\n"
            )
            
            try:
                company_response = get_perplexity_response(single_prompt)
                raw_result += company_response + "\n\n"
            except Exception as e:
                print(f"  ⚠️ Error researching {company_name}: {str(e)}")
                continue
                
        save_to_file(raw_result, week_num)
        companies_data = parse_perplexity_response(raw_result)
        
        excel_filename = f"analysis_week{week_num}_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
        create_excel_report(companies_data, week_num, excel_filename)
        html_content = generate_html_email(companies_data, week_num)
        
        if SEND_EMAIL:
            send_html_email(f"[AUTO-REPORT] 📊 Weekly Company Analysis - Week {week_num}", html_content, excel_filename)
        else:
            with open(f"email_preview_week{week_num}_{datetime.now().strftime('%H-%M-%S')}.html", "w", encoding="utf-8") as f: f.write(html_content)
            
        return 0
    except Exception as e:
        import traceback
        traceback.print_exc()
        return 1

if __name__ == "__main__":
    sys.exit(main())
