import http.server
import json
import os
import re
import urllib.request
import pandas as pd

PORT = 8080

def get_sheet_id(url):
    match = re.search(r'/d/([a-zA-Z0-9-_]+)', url)
    return match.group(1) if match else None

def download_google_sheet(sheet_url, destination_path):
    sheet_id = get_sheet_id(sheet_url)
    if not sheet_id:
        raise ValueError("Invalid Google Sheets URL format. Could not extract Spreadsheet ID.")
    
    export_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"
    print(f"Downloading Google Sheet from: {export_url}")
    
    # Configure user agent to bypass any standard bot-detection checks
    req = urllib.request.Request(
        export_url, 
        headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}
    )
    
    with urllib.request.urlopen(req) as response:
        with open(destination_path, 'wb') as out_file:
            out_file.write(response.read())
    print("Download completed successfully.")

class MordorTrackerHandler(http.server.SimpleHTTPRequestHandler):
    def do_GET(self):
        if self.path == '/api/data':
            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.send_header('Cache-Control', 'no-store, no-cache, must-revalidate')
            self.end_headers()
            
            try:
                xlsx_path = 'Step Tracker.xlsx'
                
                # Check for config.json to load Google Sheet URL
                if os.path.exists('config.json'):
                    try:
                        with open('config.json', 'r') as f:
                            config = json.load(f)
                            sheet_url = config.get('google_sheet_url', '').strip()
                            if sheet_url:
                                download_google_sheet(sheet_url, xlsx_path)
                    except Exception as e:
                        print(f"Warning: Could not sync with Google Sheets ({e}). Falling back to local cache.")
                
                if not os.path.exists(xlsx_path):
                    self.wfile.write(json.dumps({"error": f"Step Tracker file not found. Place Step Tracker.xlsx in {os.path.abspath(os.getcwd())} or set a Google Sheet URL in config.json."}).encode())
                    return
                
                xl = pd.ExcelFile(xlsx_path)
                member_sheets = [name for name in xl.sheet_names if name not in ['raw_data', 'leaderboard', 'team_totals', 'route_pct']]
                if not member_sheets:
                    self.wfile.write(json.dumps({"error": "No member data sheet found in Excel."}).encode())
                    return
                
                df = xl.parse(member_sheets[0])
                df = df[df['Name'].notna()]
                
                # Curated bright, beautiful colors for the 12 members' map tokens
                colors = [
                    "#3b82f6", # Blue
                    "#10b981", # Emerald
                    "#ec4899", # Pink
                    "#8b5cf6", # Violet
                    "#f43f5e", # Rose
                    "#06b6d4", # Cyan
                    "#14b8a6", # Teal
                    "#84cc16", # Lime
                    "#eab308", # Yellow
                    "#a855f7", # Purple
                    "#f97316", # Orange
                    "#06b6d4"  # Cyan
                ]
                
                members_list = []
                for idx, row in df.iterrows():
                    name = str(row['Name']).strip()
                    if not name or name == 'nan':
                        continue
                    
                    steps = int(row['Total Steps']) if 'Total Steps' in df.columns and not pd.isna(row['Total Steps']) else 0
                    miles = float(row['Total Distance (mi)']) if 'Total Distance (mi)' in df.columns and not pd.isna(row['Total Distance (mi)']) else steps / 2112
                    
                    initial = name[0] if name else "?"
                    color = colors[idx % len(colors)]
                    
                    members_list.append({
                        "name": name,
                        "steps": steps,
                        "miles": miles,
                        "initial": initial,
                        "color": color
                    })
                
                # Add Fellowship Combined Total Member
                if members_list:
                    total_steps = sum(m['steps'] for m in members_list)
                    total_miles = sum(m['miles'] for m in members_list)
                    members_list.append({
                        "name": "Fellowship (Total)",
                        "steps": total_steps,
                        "miles": total_miles,
                        "initial": "F",
                        "color": "#ffd700" # Radiant gold for the combined team marker
                    })
                
                # Sort members by miles descending (leaderboard)
                # Keep Fellowship (Total) first or sort normally. Sorting normally is fun as they race together!
                members_list.sort(key=lambda x: x['miles'], reverse=True)
                
                response = {
                    "members": members_list
                }
                self.wfile.write(json.dumps(response).encode())
            except Exception as e:
                self.wfile.write(json.dumps({"error": str(e)}).encode())
        else:
            super().do_GET()

if __name__ == '__main__':
    # Change CWD to the directory of this script to avoid path issues
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    server_address = ('', PORT)
    httpd = http.server.HTTPServer(server_address, MordorTrackerHandler)
    print(f"Serving Walk to Mordor tracker on http://localhost:{PORT}")
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        pass
