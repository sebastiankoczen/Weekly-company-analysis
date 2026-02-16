"""
Weekly AI Company Analysis Report - Optimized for Quality
Processes 10 companies one-by-one, produces 1 consolidated email + Excel
"""

import requests
import os
import sys
import re
import time
from datetime import datetime

# ============================================================================
# CONFIGURATION - OPTIMIZED FOR $5/MONTH BUDGET
# ============================================================================

PERPLEXITY_API_KEY = os.environ.get("PERPLEXITY_API_KEY", "")
SEND_EMAIL = os.environ.get("SEND_EMAIL", "false").lower() == "true"
SMTP_SERVER = os.environ.get("SMTP_SERVER", "smtp.gmail.com")
SMTP_PORT = int(os.environ.get("SMTP_PORT", "587"))
EMAIL_FROM = os.environ.get("EMAIL_FROM", "")
EMAIL_PASSWORD = os.environ.get("EMAIL_PASSWORD", "")
EMAIL_TO = os.environ.get("EMAIL_TO", "")

API_URL = "https://api.perplexity.ai/chat/completions"
PERPLEXITY_MODEL = "sonar-pro"  # Best quality for web search
COMPANIES_PER_WEEK = 10  # Process 10 companies per week
TOKENS_PER_COMPANY = 6000  # Generous allowance for thorough research
REQUEST_DELAY = 2  # Seconds between API calls to avoid rate limits

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
# PARSING FUNCTIONS
# ============================================================================

def parse_perplexity_response(response_text):
    """Parse structured response from Perplexity into CompanyAnalysis objects"""
    companies = []
    
    if "---COMPANY START---" not in response_text:
        print("⚠️  WARNING: No company delimiters found in response")
        return []
    
    company_blocks = re.split(r'---COMPANY START---', response_text)
    
    for idx, block in enumerate(company_blocks[1:], 1):
        if '---COMPANY END---' in block:
            block = block.split('---COMPANY END---')[0]
        
        # Extract company name
        company_match = re.search(r'Company:\s*(.+?)(?:\n|$)', block, re.IGNORECASE)
        if not company_match:
            print(f"  ⚠️  Could not find company name in block {idx}")
            continue
        
        company_name = company_match.group(1).strip()
        company = CompanyAnalysis(company_name)
        print(f"  ✅ Parsing: {company_name}")
        
        # Parse each of 4 situations
        for sit_num in range(1, 5):
            pattern = rf'SITUATION {sit_num}:.*?\nScore:\s*(\d+)\s*\nKey Points:\s*\n(.*?)\nSources:\s*(.+?)(?=\n\nSITUATION|\n---COMPANY END---|$)'
            match = re.search(pattern, block, re.DOTALL | re.IGNORECASE)
            
            if match:
                score = int(match.group(1))
                points_text = match.group(2).strip()
                sources_text = match.group(3).strip()
                
                # Extract bullet points
                points = []
                for line in points_text.split('\n'):
                    line = line.strip()
                    if line.startswith('-') or line.startswith('•') or line.startswith('*'):
                        point = re.sub(r'^[-•*]\s*', '', line)
                        if point:
                            points.append(point)
                
                # Extract URLs (looking for http/https links)
                sources = []
                for s in re.split(r'[|\n]', sources_text):
                    s = s.strip()
                    if s.startswith('http'):
                        sources.append(s)
                
                company.situations[sit_num]["score"] = score
                company.situations[sit_num]["points"] = points[:3]
                company.situations[sit_num]["sources"] = sources[:2]
        
        companies.append(company)
    
    print(f"\n✅ Successfully parsed {len(companies)} companies")
    return companies

# ============================================================================
# EXCEL GENERATION
# ============================================================================

def create_excel_report(companies, week_num, output_path):
    """Create formatted Excel report with color-coded scores and clickable links"""
    try:
        import pandas as pd
        from openpyxl import load_workbook
        from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
        
        # Prepare data
        data = []
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
        
        # Create DataFrame and save
        df = pd.DataFrame(data)
        df.to_excel(output_path, index=False, sheet_name='Company Analysis')
        
        # Load for formatting
        wb = load_workbook(output_path)
        ws = wb.active
        
        # Styling
        header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        header_font = Font(name='Arial', size=11, bold=True, color="FFFFFF")
        
        score_fill_high = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        score_fill_med = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
        score_fill_low = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        
        cell_font = Font(name='Arial', size=10)
        cell_alignment = Alignment(vertical='top', wrap_text=True)
        thin_border = Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'), bottom=Side(style='thin')
        )
        
        # Format header
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = thin_border
        
        # Format data cells
        for row_idx in range(2, ws.max_row + 1):
            for col_idx in range(1, ws.max_column + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.font = cell_font
                cell.alignment = cell_alignment
                cell.border = thin_border
                
                # Color-code scores
                if col_idx == 3:  # Score column
                    score = cell.value
                    if isinstance(score, int):
                        if score in [1, 2]:
                            cell.fill = score_fill_high
                        elif score == 3:
                            cell.fill = score_fill_med
                        elif score in [4, 5]:
                            cell.fill = score_fill_low
                
                # Make first source clickable
                if col_idx == 7:  # Sources column
                    sources_text = cell.value
                    if sources_text and '|' in sources_text:
                        urls = [s.strip() for s in sources_text.split('|')]
                        if urls and urls[0].startswith('http'):
                            cell.hyperlink = urls[0]
                            cell.font = Font(name='Arial', size=10, color="0563C1", underline="single")
        
        # Set column widths
        ws.column_dimensions['A'].width = 20
        ws.column_dimensions['B'].width = 25
        ws.column_dimensions['C'].width = 8
        ws.column_dimensions['D'].width = 40
        ws.column_dimensions['E'].width = 40
        ws.column_dimensions['F'].width = 40
        ws.column_dimensions['G'].width = 50
        
        # Set row heights
        ws.row_dimensions[1].height = 30
        for row_idx in range(2, ws.max_row + 1):
            ws.row_dimensions[row_idx].height = 60
        
        ws.freeze_panes = 'A2'
        
        wb.save(output_path)
        print(f"✅ Excel created: {output_path}")
        return True
        
    except Exception as e:
        print(f"❌ Excel creation failed: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

# ============================================================================
# HTML EMAIL GENERATION
# ============================================================================

def generate_html_email(companies, week_num):
    """Generate professional HTML email with formatted table"""
    
    html = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <style>
            body {{
                font-family: Arial, sans-serif;
                margin: 0;
                padding: 20px;
                background-color: #f5f5f5;
            }}
            .container {{
                max-width: 1200px;
                margin: 0 auto;
                background-color: white;
                padding: 30px;
                border-radius: 8px;
                box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            }}
            h1 {{
                color: #2c3e50;
                border-bottom: 3px solid #3498db;
                padding-bottom: 10px;
            }}
            .week-info {{
                background-color: #ecf0f1;
                padding: 15px;
                border-radius: 5px;
                margin-bottom: 20px;
            }}
            table {{
                width: 100%;
                border-collapse: collapse;
                margin-top: 20px;
            }}
            th {{
                background-color: #34495e;
                color: white;
                padding: 12px;
                text-align: left;
                font-weight: bold;
            }}
            td {{
                padding: 12px;
                border-bottom: 1px solid #ddd;
                vertical-align: top;
            }}
            tr:hover {{
                background-color: #f9f9f9;
            }}
            .company-name {{
                font-weight: bold;
                color: #2c3e50;
            }}
            .situation-name {{
                color: #7f8c8d;
                font-size: 0.9em;
            }}
            .score-1, .score-2 {{
                background-color: #d5f4e6;
                color: #27ae60;
                padding: 4px 8px;
                border-radius: 4px;
                font-weight: bold;
            }}
            .score-3 {{
                background-color: #fff3cd;
                color: #f39c12;
                padding: 4px 8px;
                border-radius: 4px;
                font-weight: bold;
            }}
            .score-4, .score-5 {{
                background-color: #f8d7da;
                color: #e74c3c;
                padding: 4px 8px;
                border-radius: 4px;
                font-weight: bold;
            }}
            .key-points {{
                margin: 0;
                padding-left: 20px;
            }}
            .key-points li {{
                margin-bottom: 5px;
                font-size: 0.9em;
            }}
            .sources {{
                font-size: 0.85em;
                color: #3498db;
            }}
            .sources a {{
                color: #3498db;
                text-decoration: none;
            }}
            .sources a:hover {{
                text-decoration: underline;
            }}
            .footer {{
                margin-top: 30px;
                padding-top: 20px;
                border-top: 2px solid #ecf0f1;
                text-align: center;
                color: #7f8c8d;
                font-size: 0.9em;
            }}
        </style>
    </head>
    <body>
        <div class="container">
            <h1>📊 Weekly Company Analysis Report</h1>
            
            <div class="week-info">
                <strong>Week {week_num}</strong> | Generated on {datetime.now().strftime('%B %d, %Y at %H:%M UTC')}
                <br>Companies Analyzed: {len(companies)}
            </div>
            
            <table>
                <thead>
                    <tr>
                        <th style="width: 15%;">Company</th>
                        <th style="width: 18%;">Situation</th>
                        <th style="width: 8%;">Score</th>
                        <th style="width: 44%;">Key Evidence Points</th>
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
                    else:
                        sources_links.append(source)
                sources_html = ' | '.join(sources_links)
            
            html += f"""
                    <tr>
                        <td class="company-name">{company.name}</td>
                        <td class="situation-name">{situation['name']}</td>
                        <td><span class="score-{score}">{score}</span></td>
                        <td>{points_html}</td>
                        <td class="sources">{sources_html}</td>
                    </tr>
            """
    
    html += """
                </tbody>
            </table>
            
            <div class="footer">
                <p>This report was automatically generated by the Weekly Company Analysis System</p>
                <p>Scoring: 1-2 (Low Risk/Healthy) | 3 (Moderate) | 4-5 (High Risk/Significant Issue)</p>
            </div>
        </div>
    </body>
    </html>
    """
    
    return html

# ============================================================================
# API FUNCTIONS
# ============================================================================

def get_perplexity_response(prompt_text):
    """Send prompt to Perplexity API and get response"""
    if not PERPLEXITY_API_KEY:
        raise ValueError("PERPLEXITY_API_KEY not found in environment variables")
    
    headers = {
        "Authorization": f"Bearer {PERPLEXITY_API_KEY}",
        "Content-Type": "application/json"
    }
    
    body = {
        "model": PERPLEXITY_MODEL,
        "messages": [{"role": "user", "content": prompt_text}],
        "max_tokens": TOKENS_PER_COMPANY,  # 6000 tokens per company
        "stream": False,
        "temperature": 0.2,
        "top_p": 0.9
    }
    
    try:
        response = requests.post(API_URL, json=body, headers=headers, timeout=180)
        
        if response.status_code != 200:
            raise Exception(f"Perplexity API Error {response.status_code}: {response.text}")
        
        result = response.json()
        return result["choices"][0]["message"]["content"]
    
    except requests.exceptions.Timeout:
        raise Exception("Request timed out after 180 seconds")
    except Exception as e:
        raise Exception(f"API request failed: {str(e)}")


def send_html_email(subject, html_content, excel_path=None):
    """Send HTML email with Excel attachment"""
    if not SEND_EMAIL:
        print("📧 Email sending disabled (SEND_EMAIL not true)")
        return
    
    try:
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
                excel_attachment.add_header('Content-Disposition', 'attachment', 
                                          filename=os.path.basename(excel_path))
                msg.attach(excel_attachment)
        
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(EMAIL_FROM, EMAIL_PASSWORD)
        server.sendmail(EMAIL_FROM, EMAIL_TO, msg.as_string())
        server.quit()
        
        print(f"✅ Email sent to {EMAIL_TO}")
        if excel_path:
            print(f"   📎 Attached: {os.path.basename(excel_path)}")
    
    except Exception as e:
        print(f"❌ Email failed: {str(e)}")


def get_companies_for_week(companies, week_num, per_week=10):
    """Get companies for specific week"""
    idx_start = (week_num - 1) * per_week
    idx_end = idx_start + per_week
    
    if idx_start >= len(companies):
        raise ValueError(f"Week {week_num} exceeds available companies")
    
    return companies[idx_start:idx_end]


def calculate_current_week(start_date, companies_per_week, total_companies):
    """Calculate current week based on date"""
    from datetime import datetime
    
    start = datetime.strptime(start_date, "%Y-%m-%d")
    today = datetime.now()
    
    days_diff = (today - start).days
    weeks_passed = days_diff // 7
    
    total_weeks = (total_companies + companies_per_week - 1) // companies_per_week
    current_week = (weeks_passed % total_weeks) + 1
    
    return current_week


def save_to_file(content, week_num):
    """Save raw report to file"""
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    filename = f"report_week{week_num}_{timestamp}.txt"
    
    try:
        with open(filename, "w", encoding="utf-8") as f:
            f.write(content)
        print(f"✅ Raw report saved: {filename}")
        return filename
    except Exception as e:
        print(f"❌ Save failed: {str(e)}")
        return None

# ============================================================================
# MAIN EXECUTION - ONE-BY-ONE PROCESSING
# ============================================================================

def main():
    """Main function - processes 10 companies one-by-one, sends 1 consolidated email"""
    print("=" * 70)
    print("🤖 WEEKLY COMPANY ANALYSIS - OPTIMIZED FOR QUALITY")
    print("=" * 70)
    print(f"Started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"Model: {PERPLEXITY_MODEL}")
    print(f"Tokens per company: {TOKENS_PER_COMPANY}")
    print(f"Companies per week: {COMPANIES_PER_WEEK}")
    print()
    
    try:
        # Load files
        print("📂 Loading configuration files...")
        
        required_files = ["prompt_updated.txt", "companies.txt", "definitions.txt"]
        for file in required_files:
            if not os.path.exists(file):
                raise FileNotFoundError(f"Missing file: {file}")
        
        with open("prompt_updated.txt", "r", encoding="utf-8") as f:
            role_objective = f.read().strip()
        
        with open("companies.txt", "r", encoding="utf-8") as f:
            all_companies = [line.strip() for line in f if line.strip()]
        
        with open("definitions.txt", "r", encoding="utf-8") as f:
            scoring_defs = f.read().strip()
        
        print(f"✅ Loaded {len(all_companies)} companies")
        
        # Determine week
        week_num = int(os.environ.get("WEEK", "0"))
        
        if week_num == 0:
            start_date = os.environ.get("START_DATE", "2025-02-17")
            week_num = calculate_current_week(start_date, COMPANIES_PER_WEEK, len(all_companies))
            print(f"📅 Auto-calculated week: {week_num}")
        else:
            print(f"📅 Manual week: {week_num}")
        
        # Get companies for this week
        companies_this_week = get_companies_for_week(all_companies, week_num, COMPANIES_PER_WEEK)
        
        print(f"\n📋 Week {week_num} - Analyzing {len(companies_this_week)} companies:")
        for i, company in enumerate(companies_this_week, 1):
            print(f"   {i}. {company}")
        print()
        
        # PROCESS COMPANIES ONE-BY-ONE
        print("🔍 Starting one-by-one analysis (this will take a few minutes)...")
        print()
        
        raw_result = ""
        successful_count = 0
        
        for i, company_name in enumerate(companies_this_week, 1):
            print(f"[{i}/{len(companies_this_week)}] 🔎 Researching: {company_name}")
            
            # Compose prompt for single company
            single_prompt = (
                f"{role_objective}\n\n"
                f"Definitions of Situations and Scoring:\n{scoring_defs}\n\n"
                f"Company to Analyze:\n{company_name}\n"
            )
            
            try:
                # API call for this company
                company_response = get_perplexity_response(single_prompt)
                raw_result += company_response + "\n\n"
                successful_count += 1
                print(f"    ✅ Success ({len(company_response)} characters)")
                
                # Delay between requests to avoid rate limits
                if i < len(companies_this_week):
                    time.sleep(REQUEST_DELAY)
                
            except Exception as e:
                print(f"    ❌ Error: {str(e)}")
                continue
        
        print()
        print(f"✅ Successfully analyzed {successful_count}/{len(companies_this_week)} companies")
        print()
        
        # Save raw output
        save_to_file(raw_result, week_num)
        
        # Parse all companies
        print("🔄 Parsing results...")
        companies_data = parse_perplexity_response(raw_result)
        
        if not companies_data:
            print("⚠️  WARNING: No companies were parsed successfully")
            print("Check the raw .txt file for response format issues")
        
        # Generate Excel
        print("\n📊 Generating Excel report...")
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        excel_filename = f"analysis_week{week_num}_{timestamp}.xlsx"
        create_excel_report(companies_data, week_num, excel_filename)
        
        # Generate HTML
        print("📧 Generating HTML email...")
        html_content = generate_html_email(companies_data, week_num)
        
        # Send email or save preview
        if SEND_EMAIL:
            subject = f"📊 Weekly Company Analysis - Week {week_num} - {datetime.now().strftime('%Y-%m-%d')}"
            send_html_email(subject, html_content, excel_filename)
        else:
            html_filename = f"email_preview_week{week_num}_{timestamp}.html"
            with open(html_filename, "w", encoding="utf-8") as f:
                f.write(html_content)
            print(f"✅ Email preview saved: {html_filename}")
        
        print("\n" + "=" * 70)
        print("✅ PROCESS COMPLETED SUCCESSFULLY")
        print("=" * 70)
        print(f"\n📁 Generated files:")
        print(f"   • Raw response: report_week{week_num}_*.txt")
        print(f"   • Excel: {excel_filename}")
        if not SEND_EMAIL:
            print(f"   • HTML preview: email_preview_week{week_num}_*.html")
        print()
        
        # Cost estimate
        total_tokens = successful_count * TOKENS_PER_COMPANY
        estimated_cost = (total_tokens / 1_000_000) * 1.5  # $1.50 per 1M tokens
        print(f"💰 Estimated cost this run: ${estimated_cost:.4f}")
        print(f"   ({total_tokens:,} tokens at ~$1.50/1M)")
        print()
        
        return 0
    
    except Exception as e:
        print(f"\n❌ Error: {str(e)}")
        import traceback
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    sys.exit(main())
