"""
Weekly AI Company Analysis Report - Enhanced Version with Better Parsing
Runs every Monday to analyze 10 companies and send formatted results via email
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
PERPLEXITY_MODEL = "sonar-pro"
COMPANIES_PER_WEEK = 5

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
# IMPROVED PARSING FUNCTIONS
# ============================================================================

def parse_perplexity_response(response_text):
    """
    Parse the structured response from Perplexity with improved error handling
    """
    companies = []
    
    print("\n🔍 DEBUG: Starting to parse response...")
    print(f"Response length: {len(response_text)} characters")
    
    # Check if response contains expected delimiters
    if "---COMPANY START---" not in response_text:
        print("⚠️  WARNING: No '---COMPANY START---' delimiter found!")
        print("This means Perplexity didn't follow the formatting instructions.")
        print("\nAttempting flexible parsing...")
        
        # Try flexible parsing by company names
        return parse_flexible(response_text)
    
    # Split by company blocks
    company_blocks = re.split(r'---COMPANY START---', response_text)
    print(f"Found {len(company_blocks)-1} company blocks")
    
    for idx, block in enumerate(company_blocks[1:], 1):
        print(f"\n📋 Processing company block {idx}...")
        
        if '---COMPANY END---' not in block:
            print(f"  ⚠️  WARNING: No '---COMPANY END---' delimiter in block {idx}")
            # Try to process anyway
            pass
        else:
            block = block.split('---COMPANY END---')[0]
        
        # Extract company name
        company_match = re.search(r'Company:\s*(.+?)(?:\n|$)', block, re.IGNORECASE)
        if not company_match:
            print(f"  ❌ Could not find company name in block {idx}")
            continue
        
        company_name = company_match.group(1).strip()
        print(f"  ✅ Company: {company_name}")
        company = CompanyAnalysis(company_name)
        
        # Parse each situation with more flexible regex
        situations_found = 0
        for sit_num in range(1, 5):
            # More flexible pattern matching
            patterns = [
                # Pattern 1: Exact format
                rf'SITUATION {sit_num}:.*?\nScore:\s*(\d+)\s*\nKey Points:\s*\n(.*?)\nSources:\s*(.+?)(?=\n\nSITUATION|\n---COMPANY END---|$)',
                # Pattern 2: Flexible whitespace
                rf'SITUATION {sit_num}[:\s]+.*?\nScore[:\s]+(\d+)\s*\n.*?Key Points[:\s]+\n(.*?)\nSources[:\s]+(.+?)(?=\n\nSITUATION|\n---COMPANY END---|$)',
                # Pattern 3: Case insensitive
                rf'situation {sit_num}[:\s]+.*?\nscore[:\s]+(\d+)\s*\n.*?key points[:\s]+\n(.*?)\nsources[:\s]+(.+?)(?=\n\nsituation|\n---company end---|$)'
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
                
                # Extract bullet points
                points = []
                for line in points_text.split('\n'):
                    line = line.strip()
                    if line.startswith('-') or line.startswith('•') or line.startswith('*'):
                        point = re.sub(r'^[-•*]\s*', '', line)
                        if point:
                            points.append(point)
                
                # Extract sources
                sources = [s.strip() for s in re.split(r'[|\n]', sources_text) if s.strip() and s.strip().startswith('http')]
                
                company.situations[sit_num]["score"] = score
                company.situations[sit_num]["points"] = points[:3]  # Max 3
                company.situations[sit_num]["sources"] = sources[:2]  # Max 2
                
                situations_found += 1
                print(f"  ✅ Situation {sit_num}: Score {score}, {len(points)} points, {len(sources)} sources")
            else:
                print(f"  ⚠️  Situation {sit_num}: Not found or couldn't parse")
        
        if situations_found > 0:
            companies.append(company)
            print(f"  ✅ Successfully parsed {situations_found}/4 situations for {company_name}")
        else:
            print(f"  ❌ No situations parsed for {company_name} - skipping")
    
    print(f"\n✅ Total companies successfully parsed: {len(companies)}")
    return companies


def parse_flexible(response_text):
    """
    Fallback parser when strict format isn't followed
    Attempts to extract any structured data present
    """
    print("\n🔄 Using flexible parsing mode...")
    companies = []
    
    # Try to identify company sections by common patterns
    # Look for company names followed by situation analysis
    
    print("⚠️  Flexible parsing not fully implemented - raw response will be saved")
    print("Please check the raw .txt file and adjust the prompt if needed")
    
    return companies


# ============================================================================
# EXCEL GENERATION
# ============================================================================

def create_excel_report(companies, week_num, output_path):
    """Create an Excel file with formatted analysis"""
    if not companies:
        print("⚠️  No companies to export - creating empty template")
    
    try:
        import pandas as pd
        from openpyxl import load_workbook
        from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
        
        # Prepare data
        data = []
        
        if not companies:
            # Create empty row as template
            data.append({
                'Company': 'No data parsed',
                'Situation': 'Check raw .txt file for actual response',
                'Score': 0,
                'Key Point 1': 'Parsing may have failed',
                'Key Point 2': 'Check prompt format',
                'Key Point 3': 'Verify Perplexity response structure',
                'Sources': 'N/A'
            })
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
        
        # Create DataFrame and save
        df = pd.DataFrame(data)
        df.to_excel(output_path, index=False, sheet_name='Company Analysis')
        
        # Load workbook for formatting
        wb = load_workbook(output_path)
        ws = wb.active
        
        # Define styles
        header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        header_font = Font(name='Arial', size=11, bold=True, color="FFFFFF")
        
        score_fill_high = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        score_fill_med = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
        score_fill_low = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        
        cell_font = Font(name='Arial', size=10)
        cell_alignment = Alignment(vertical='top', wrap_text=True)
        
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
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
        print(f"✅ Excel file created: {output_path}")
        return True
        
    except Exception as e:
        print(f"❌ Failed to create Excel: {str(e)}")
        import traceback
        traceback.print_exc()
        return False


# ============================================================================
# HTML EMAIL GENERATION
# ============================================================================

def generate_html_email(companies, week_num):
    """Generate HTML email"""
    
    if not companies:
        return f"""
        <!DOCTYPE html>
        <html>
        <body style="font-family: Arial; padding: 20px;">
            <h1>⚠️ Weekly Company Analysis - Week {week_num}</h1>
            <p>No companies were successfully parsed from the Perplexity response.</p>
            <p><strong>Troubleshooting steps:</strong></p>
            <ol>
                <li>Check the raw .txt file in GitHub Artifacts</li>
                <li>Verify Perplexity is returning data in the expected format</li>
                <li>Review the prompt_updated.txt file</li>
                <li>Check if the model "sonar" is correct for your API plan</li>
            </ol>
            <p>Generated on {datetime.now().strftime('%B %d, %Y at %H:%M UTC')}</p>
        </body>
        </html>
        """
    
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
    """Send prompt to Perplexity API"""
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
    
    try:
        print(f"Sending request to Perplexity API (model: {PERPLEXITY_MODEL})...")
        response = requests.post(API_URL, json=body, headers=headers, timeout=180)
        
        if response.status_code != 200:
            print(f"API Error {response.status_code}: {response.text}")
            raise Exception(f"Perplexity API Error: {response.text}")
        
        result = response.json()
        return result["choices"][0]["message"]["content"]
    
    except Exception as e:
        raise Exception(f"API request failed: {str(e)}")


def send_html_email(subject, html_content, excel_path=None):
    """Send email with HTML and Excel attachment"""
    if not SEND_EMAIL:
        print("Email sending is disabled.")
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
    
    except Exception as e:
        print(f"❌ Failed to send email: {str(e)}")


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
    """Save raw report"""
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    filename = f"report_week{week_num}_{timestamp}.txt"
    
    try:
        with open(filename, "w", encoding="utf-8") as f:
            f.write(content)
        print(f"✅ Raw report saved to: {filename}")
        return filename
    except Exception as e:
        print(f"❌ Failed to save file: {str(e)}")
        return None


# ============================================================================
# MAIN EXECUTION
# ============================================================================

def main():
    """Main function"""
    print("=" * 70)
    print("🤖 WEEKLY AI COMPANY ANALYSIS REPORT (DEBUG MODE)")
    print("=" * 70)
    print(f"Started at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print()
    
    try:
        # Load files
        print("📂 Loading files...")
        
        required_files = ["prompt_updated.txt", "companies.txt", "definitions.txt"]
        for file in required_files:
            if not os.path.exists(file):
                raise FileNotFoundError(f"Required file not found: {file}")
        
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
        companies_text = "\n".join(f"- {name}" for name in companies_this_week)
        
        print(f"\n📋 Analyzing {len(companies_this_week)} companies for Week {week_num}:")
        for i, company in enumerate(companies_this_week, 1):
            print(f"   {i}. {company}")
        print()
        
        # Compose prompt
        prompt = (
            f"{role_objective}\n\n"
            f"Definitions of Situations and Scoring:\n{scoring_defs}\n\n"
            f"Companies to Analyze (Week {week_num}):\n{companies_text}\n"
        )
        
        # Get analysis
        print("🔍 Requesting analysis from Perplexity AI...")
        print("(This may take 2-3 minutes for 10 companies)")
        raw_result = get_perplexity_response(prompt)
        
        print(f"\n✅ Received response ({len(raw_result)} characters)")
        
        # Save raw output
        save_to_file(raw_result, week_num)
        
        # Parse response
        print("\n" + "=" * 70)
        companies_data = parse_perplexity_response(raw_result)
        print("=" * 70)
        
        if not companies_data:
            print("\n⚠️  WARNING: No companies were successfully parsed!")
            print("This usually means Perplexity didn't follow the formatting instructions.")
            print("Check the raw .txt file to see what was actually returned.")
        
        # Generate outputs anyway (will show warning if empty)
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        excel_filename = f"analysis_week{week_num}_{timestamp}.xlsx"
        
        print(f"\n📊 Generating Excel report...")
        create_excel_report(companies_data, week_num, excel_filename)
        
        print(f"\n📧 Generating HTML email...")
        html_content = generate_html_email(companies_data, week_num)
        
        if SEND_EMAIL:
            subject = f"[AUTO-REPORT] 📊 Weekly Company Analysis - Week {week_num} - {datetime.now().strftime('%Y-%m-%d')}"
            send_html_email(subject, html_content, excel_filename)
        else:
            html_filename = f"email_preview_week{week_num}_{timestamp}.html"
            with open(html_filename, "w", encoding="utf-8") as f:
                f.write(html_content)
            print(f"✅ Email preview saved to: {html_filename}")
        
        print("\n" + "=" * 70)
        if companies_data:
            print(f"✅ SUCCESS! Processed {len(companies_data)} companies")
        else:
            print("⚠️  COMPLETED WITH WARNINGS - Check raw .txt file")
        print("=" * 70)
        
        print(f"\n📁 Generated files:")
        print(f"   • Raw response: report_week{week_num}_*.txt")
        print(f"   • Excel: {excel_filename}")
        if not SEND_EMAIL:
            print(f"   • HTML: email_preview_week{week_num}_*.html")
        print()
        
        return 0
    
    except Exception as e:
        print(f"\n❌ Error: {str(e)}")
        import traceback
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    sys.exit(main())
