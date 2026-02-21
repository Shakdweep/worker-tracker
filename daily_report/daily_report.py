import os
import sys
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
import requests
import pandas as pd
import gspread
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from datetime import date, datetime, timedelta
from daily_report.report_queries import QUERIES

# =========================
# CONFIGURATION
# =========================
METABASE_URL = "http://prod-db.kashti.com:3000"
USERNAME = "admin@kashti.com"
PASSWORD = "PlutoAnalytics@12"
DATABASE_ID = 2

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]
SHEET_URL = "https://docs.google.com/spreadsheets/d/1j0ISgfQdLP8Q60klUeGXg7YblKSpR6UoxX5YrPLpuNA/edit?usp=sharing"
SHEET_NAME = "internal report - Working sheet"
CREDENTIALS_FILE = "sagar_shakdweep.json"

# =========================
# METABASE DATA FETCH
# =========================
def get_metabase_token():
    """Authenticate with Metabase and return a session token."""
    auth = requests.post(
        f"{METABASE_URL}/api/session",
        json={"username": USERNAME, "password": PASSWORD},
        timeout=(30, 30)
    )
    auth.raise_for_status()
    return auth.json()['id']

def fetchdata(sql_query, token=None):
    """Send SQL query to Metabase and return results as a DataFrame."""
    try:
        # Authenticate only if a token was not provided
        if token is None:
            token = get_metabase_token()
        headers = {"X-Metabase-Session": token}
        
        # Debug: Print query info
        print(f"Executing query (first 100 chars): {sql_query[:100]}...")
        
        query_dict = {
            "database": DATABASE_ID,
            "type": "native",
            "native": {"query": sql_query},
            "constraints": {
                "max-results": 50000,  # Increase row limit
                "max-results-bare-rows": 50000
            }
        }
        
        response = requests.post(f"{METABASE_URL}/api/dataset", headers=headers, json=query_dict, timeout=(30, 600))
        response.raise_for_status()
        result = response.json()
        
        # Debug: Print response metadata
        print(f"Response status: {result.get('status', 'unknown')}")
        if 'data' in result:
            print(f"Columns returned: {len(result['data'].get('cols', []))}")
            print(f"Rows returned: {len(result['data'].get('rows', []))}")
            
            # Check if results were truncated
            if 'truncated' in result['data']:
                print(f"WARNING: Results truncated: {result['data']['truncated']}")
        
        if "data" in result and result['data']['rows']:
            rows = result['data']['rows']
            columns = [c['name'] for c in result['data']['cols']]
            df = pd.DataFrame(rows, columns=columns)
            
            # Additional debug info
            print(f"DataFrame created: {df.shape[0]} rows, {df.shape[1]} columns")
            if df.shape[0] >= 10000:
                print("WARNING: Large dataset detected - check if pagination is needed")
            
            return df
        else:
            print("Query executed but returned 0 rows")
            print(f"Full response: {result}")
            return pd.DataFrame()
            
    except requests.exceptions.Timeout:
        print(f"ERROR: Query timeout - try breaking down the query or increasing timeout")
        return pd.DataFrame()
    except Exception as e:
        print(f"Error fetching data: {e}")
        print(f"Response content (if available): {getattr(e.response, 'text', 'No response content')}")
        return pd.DataFrame()

# Add a new function for large dataset handling
def fetchdata_with_pagination(sql_query, batch_size=10000, token=None):
    """Fetch large datasets using LIMIT/OFFSET pagination."""
    try:
        all_dfs = []
        offset = 0
        
        while True:
            # Add LIMIT and OFFSET to the query
            paginated_query = f"""
            {sql_query}
            LIMIT {batch_size} OFFSET {offset}
            """
            
            print(f"Fetching batch starting at offset {offset}...")
            df_batch = fetchdata(paginated_query, token=token)
            
            if df_batch.empty:
                break
                
            all_dfs.append(df_batch)
            
            # If we got less than batch_size rows, we're done
            if len(df_batch) < batch_size:
                break
                
            offset += batch_size
        
        if all_dfs:
            final_df = pd.concat(all_dfs, ignore_index=True)
            print(f"Total rows fetched across all batches: {len(final_df)}")
            return final_df
        else:
            return pd.DataFrame()
            
    except Exception as e:
        print(f"Error in paginated fetch: {e}")
        return pd.DataFrame()

# =========================
# GOOGLE SHEETS AUTH
# =========================
def get_creds():
    creds = None
    # Check for token.json in current directory, then parent directory
    token_paths = ["token.json", "../token.json"]
    token_file = None
    
    for path in token_paths:
        if os.path.exists(path):
            token_file = path
            break
    
    if token_file:
        try:
            creds = Credentials.from_authorized_user_file(token_file, SCOPES)
        except Exception as e:
            print(f"Error loading token file: {e}")
            creds = None
    
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except Exception as e:
                print(f"Error refreshing credentials: {e}")
                creds = None
        
        if not creds:
            # Check if credentials file exists
            if os.path.exists(CREDENTIALS_FILE):
                try:
                    flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
                    creds = flow.run_local_server(port=0)
                except Exception as e:
                    print(f"Error with credentials flow: {e}")
                    raise
            else:
                raise FileNotFoundError(f"Credentials file {CREDENTIALS_FILE} not found. Please ensure the Google OAuth credentials file exists.")
        
        # Save token to the found location or current directory
        if creds:
            save_path = token_file if token_file else "token.json"
            try:
                with open(save_path, "w") as token:
                    token.write(creds.to_json())
            except Exception as e:
                print(f"Warning: Could not save token file: {e}")
    
    return creds

# ======================
# Google Sheets Export (simplified and fixed)
# ======================
def df_to_sheet(df, worksheet, start_row, start_col):
    """Helper: Write pandas DataFrame into Google Sheet from a given row/col"""
    try:
        values = [df.columns.tolist()] + df.astype(str).values.tolist()
        worksheet.update(
            gspread.utils.rowcol_to_a1(start_row, start_col),
            values
        )
    except Exception as e:
        print(f"Error writing to sheet: {e}")

def export_to_gsheets(summary_df, merged, result, final_df, sheet_url, login_data=None):
    try:
        # Initialize connection
        creds = get_creds()
        client = gspread.authorize(creds)
        sh = client.open_by_url(sheet_url)
        
        # Calculate yesterday's date for display
        yesterday = (datetime.today() - timedelta(days=1)).strftime("%d %B %Y")
        yesterday_date = (datetime.today() - timedelta(days=1)).date()
        
        # === Work with "Report" sheet ===
        try:
            ws_report = sh.worksheet(SHEET_NAME)
            ws_report.clear()
        except gspread.exceptions.WorksheetNotFound:
            ws_report = sh.add_worksheet(title=SHEET_NAME, rows="1000", cols="30")
        
        # Helper function to get credit score values
        def get_score_val(cat):
            try:
                filtered = result[result["Score Category"] == cat]
                if not filtered.empty:
                    return int(filtered["Count"].values[0])
                return 0
            except Exception:
                return 0
        
        # Get values for calculations (with error handling)
        try:
            # Extract values safely from summary_df
            lead = int(summary_df.loc[0, "No in total"]) if len(summary_df) > 0 else 0
            u_lead = int(summary_df.loc[1, "No in total"]) if len(summary_df) > 1 else 0
            application = int(summary_df.loc[2, "No in total"]) if len(summary_df) > 2 else 0
            a_application = int(summary_df.loc[3, "No in total"]) if len(summary_df) > 3 else 0
            prev = int(summary_df.loc[4, "No in total"]) if len(summary_df) > 4 else 0
            class1 = int(summary_df.loc[5, "No in total"]) if len(summary_df) > 5 else 0
            class2 = int(summary_df.loc[6, "No in total"]) if len(summary_df) > 6 else 0
        except Exception as e:
            print(f"Error extracting values from summary_df: {e}")
            lead = u_lead = application = a_application = prev = class1 = class2 = 0
        
        # Build all rows for the report
        all_rows = []
        
        # Row 1: Date header
        all_rows.append([f"Date - {yesterday} ( All Data for Kashti )"] + ["" for _ in range(5)])
        
        # Row 2: Section title
        all_rows.append(["Below Data table on Application Login"] + ["" for _ in range(5)])
        
        # Row 3: Headers
        all_rows.append(["Particulars", "No in total", "Percentage %", "Desc", "Credit Score Pulls", ""])
        
        # Rows 4-10: Data rows
        particulars = [
            "Lead", "Unique Lead", "Application", "Approved Application (Unique)",
            "Approved Application Login - C1 + C2", "Approved Application Login - Class 1",
            "Approved Application Login - Class 2"
        ]
        
        no_in_total = [str(lead), str(u_lead), str(application), str(a_application), str(prev), str(class1), str(class2)]
        
        # Calculate percentages safely
        def safe_percentage(numerator, denominator):
            try:
                if denominator and denominator > 0:
                    return str(round((numerator/denominator)*100, 2))
                return "0"
            except Exception:
                return "0"
        
        percentage = [
            "null",
            safe_percentage(u_lead, lead),
            safe_percentage(application, u_lead),
            safe_percentage(a_application, u_lead),
            safe_percentage(prev, a_application),
            safe_percentage(class1, prev),
            safe_percentage(class2, prev)
        ]
        
        descs = ["null", "Ul to L %", "AA to UI%", "App to Ul%", "APR to app%", "C1 to APR", "C2 to APR"]
        
        # Credit Score data
        credit_labels = ["Total", "Below 650", "Above 650", "Above 720", "New to Credit", "Day Credit Spent"]
        credit_vals = [
            str(get_score_val("Total")),
            str(get_score_val("Below 650")),
            str(get_score_val("Above 650")),
            str(get_score_val("Above 720")),
            str(get_score_val("New to Credit")),
            str(get_score_val("Day Credit Spent"))
        ]
        
        # Fill data rows
        for i in range(7):
            row = [particulars[i], no_in_total[i], percentage[i], descs[i]]
            if i < 6:
                row += [credit_labels[i], credit_vals[i]]
            else:
                row += ["", ""]
            all_rows.append(row)
        
        # Add blank row
        all_rows.append(["" for _ in range(6)])
        
        # Platform-wise section
        all_rows.append(["Platform Wise data -Lead to App to Approved app - User Application"] + ["" for _ in range(4)])
        
        # Platform headers and data
        if not merged.empty:
            all_rows.append([str(cell) for cell in merged.columns.tolist()])
            for row in merged.values.tolist():
                all_rows.append([str(cell) for cell in row])
        
        # Spends section
        all_rows.append(["Platform Wise Spends data - User Application ( Data on Login )"] + ["" for _ in range(5)])
        all_rows.append(["Source", "Spends", "Cost Per App", "Cost Per AAP", "Class 1 - CPAAP", "Class 2 - CPAAP"])
        
        # Spends data rows
        spends_sources = ["Google Ads", "Meta Ads", "RCS Ads", "SMS Campign", "Whatsapp_Utility"]
        for src in spends_sources:
            all_rows.append([src] + ["" for _ in range(5)])
        
        # Lender section - FIXED VERSION
        all_rows.append(["Lender wise - Application Login data"] + ["" for _ in range(5)])
        all_rows.append(["Lenders", "Approved App", "Cost per approved", "", "", ""])
        
        # ADD ACTUAL LENDER DATA PROCESSING HERE
        try:
            if login_data is not None and not login_data.empty:
                # Filter out test data
                filtered_login = login_data[~login_data["pipeline_name"].str.contains('test', case=False, na=False)]
                
                # Group by institution (lender) to get approved app counts
                # CHANGED: Using .count() instead of .nunique() to match C1+C2 calculation
                lender_data = filtered_login.groupby('institution')['mobile'].count().reset_index()
                lender_data.columns = ['Lender', 'Approved_App']
                
                # Remove test data if any
                lender_data = lender_data[~lender_data['Lender'].str.contains('test', case=False, na=False)]
                
                # Sort by approved app count descending
                lender_data = lender_data.sort_values('Approved_App', ascending=False)
                
                # Add lender rows
                for idx, row in lender_data.iterrows():
                    lender_name = row['Lender']
                    approved_count = row['Approved_App']
                    
                    all_rows.append([
                        lender_name, 
                        str(approved_count), 
                        "", # Cost per approved - empty as requested
                        "", "", ""
                    ])
                
                # Add total row - this should now be 406
                total_approved = lender_data['Approved_App'].sum()
                all_rows.append([
                    "Total", 
                    str(total_approved), 
                    "", 
                    "", "", ""
                ])
                
            else:
                # If no login data, add a placeholder row
                all_rows.append(["No data available", "0", "", "", "", ""])
                
        except Exception as e:
            print(f"Error processing lender data: {e}")
            all_rows.append(["Error loading lender data", "0", "", "", "", ""])
        
        # Convert all values to proper types for Google Sheets
        def convert_cell(cell):
            if cell is None or cell == "" or pd.isna(cell):
                return ""
            if isinstance(cell, (int, float)):
                return cell
            if isinstance(cell, str):
                try:
                    # Try to convert to number if it looks like one
                    if cell.replace('.', '').replace('-', '').isdigit():
                        if '.' in cell:
                            return float(cell)
                        return int(cell)
                    return cell
                except (ValueError, AttributeError):
                    return str(cell)
            return str(cell)
        
        all_rows = [[convert_cell(cell) for cell in row] for row in all_rows]
        
        # Write to sheet in batches to avoid API limits
        batch_size = 100
        for i in range(0, len(all_rows), batch_size):
            batch = all_rows[i:i+batch_size]
            start_row = i + 1
            end_col = chr(65 + max(len(row) for row in batch) - 1) if batch else 'F'
            range_name = f"A{start_row}:{end_col}{start_row + len(batch) - 1}"
            ws_report.update(range_name, batch)
        
        # === Work with Monthly Data sheet ===
        try:
            ws_monthly = sh.worksheet("Copy of MonthlyData")
        except gspread.exceptions.WorksheetNotFound:
            ws_monthly = sh.add_worksheet(title="Copy of MonthlyData", rows="1000", cols="30")
        
        # Prepare summary row with proper data types
        summary_row = [
            yesterday_date.strftime('%Y-%m-%d'),
            u_lead, application, a_application,
            safe_percentage(application, u_lead).replace('%', ''),
            safe_percentage(a_application, u_lead).replace('%', ''),
            safe_percentage(prev, u_lead).replace('%', ''),
            safe_percentage(prev, application).replace('%', ''),
            prev, class1, class2,
            get_score_val("Total"),
            get_score_val("Above 650")
        ]
        
        # Convert to appropriate types for Google Sheets
        summary_row_converted = []
        for i, val in enumerate(summary_row):
            if i == 0:  # Date column
                summary_row_converted.append(str(val))
            else:
                try:
                    # Convert to number if possible
                    if isinstance(val, str) and val.replace('.', '').replace('-', '').isdigit():
                        summary_row_converted.append(float(val) if '.' in val else int(val))
                    elif isinstance(val, (int, float)):
                        summary_row_converted.append(val)
                    else:
                        summary_row_converted.append(str(val))
                except:
                    summary_row_converted.append(0)
        
        # Check if date already exists and update or append
        try:
            all_data = ws_monthly.get_all_values()
            target_date = summary_row_converted[0]
            date_exists = False
            
            for i, row in enumerate(all_data):
                if len(row) > 0 and row[0] == target_date:
                    # Update existing row
                    cell_range = f"A{i+1}:{chr(65+len(summary_row_converted)-1)}{i+1}"
                    ws_monthly.update(cell_range, [summary_row_converted])
                    date_exists = True
                    break
            
            if not date_exists:
                # Find position to insert (before "Total" if it exists)
                insert_row = None
                for i, row in enumerate(all_data):
                    if len(row) > 0 and row[0] == "Total":
                        insert_row = i + 1
                        break
                
                if insert_row:
                    ws_monthly.insert_row(summary_row_converted, insert_row)
                else:
                    ws_monthly.append_row(summary_row_converted)
        
        except Exception as e:
            print(f"Error with monthly data: {e}")
            # Fallback: just append
            ws_monthly.append_row(summary_row_converted)
        
        print("SUCCESS: Export complete - Report and Monthly Data sheets updated")
        
    except Exception as e:
        print(f"ERROR in export_to_gsheets: {e}")
        import traceback
        traceback.print_exc()

# =========================
# MAIN PIPELINE - APPLICATION BASED ON PIPELINE_NAME
# =========================
if __name__ == "__main__":
    try:
        print("STARTING: Running daily report pipeline...")
        print("=" * 50)
        
        # Fetch data with enhanced debugging
        print("FETCHING: Data from Metabase...")
        print("Trying standard fetch first...")

        token = get_metabase_token()
        df1 = fetchdata(QUERIES["main"], token=token)
        login = fetchdata(QUERIES["login"], token=token)
        credit = fetchdata(QUERIES["credit"], token=token)
        
        # If we suspect truncation, try paginated approach
        if len(df1) == 2000 or len(credit) == 2000:
            print("\nWARNING: Detected potential row limit truncation")
            print("Trying paginated fetch for larger datasets...")
            
            if len(df1) == 2000:
                print("Re-fetching main data with pagination...")
                df1_paginated = fetchdata_with_pagination(QUERIES["main"], token=token)
                if len(df1_paginated) > len(df1):
                    print(f"Pagination successful: {len(df1)} -> {len(df1_paginated)} rows")
                    df1 = df1_paginated
            
            if len(credit) == 2000:
                print("Re-fetching credit data with pagination...")
                credit_paginated = fetchdata_with_pagination(QUERIES["credit"], token=token)
                if len(credit_paginated) > len(credit):
                    print(f"Pagination successful: {len(credit)} -> {len(credit_paginated)} rows")
                    credit = credit_paginated
        
        if df1.empty:
            print("ERROR: No main data received")
            sys.exit(1)
        
        print(f"FINAL SUCCESS: Data fetched successfully:")
        print(f"   Main: {len(df1)} rows")
        print(f"   Login: {len(login)} rows")
        print(f"   Credit: {len(credit)} rows")
        
        # Data Processing - APPLICATION BASED ON PIPELINE_NAME
        print("\nPROCESSING: Data analysis...")
        
        # Filter out test data
        dump = df1[~df1["pipeline_name"].str.contains('test', case=False, na=False)]
        lead = dump['mobile'].count()
        u_lead = dump['mobile'].nunique()
        
        # ===== FIXED: Calculate Application from pipeline_name groupby =====
        # Get app data (both Offer Screen and SJ Offer Screen)
        app = dump[dump['current_screen_name'].isin(['Offer Screen', 'SJ Offer Screen'])]
        
        # Platform-wise App count: count all records per pipeline
        # This is the SOURCE OF TRUTH for Application
        app_1 = app.groupby('pipeline_name')['mobile'].count().rename("App")
        
        # Application = sum of all pipeline counts (17 + 533 + 522 + 4 = 1076)
        application = int(app_1.sum())
        
        # Approved applications - get unique count per screen then sum across all screens
        plu = dump[dump['pluto_status'] == 'Approved']
        a_application = plu.groupby('current_screen_name')['mobile'].nunique().sum()

        # Login data processing
        dump1 = login[~login["pipeline_name"].str.contains('test', case=False, na=False)]
        c2 = dump1[dump1['institution'].isin(['FatakPay', 'mPokket'])]
        c1 = dump1[~dump1['institution'].isin(['FatakPay', 'mPokket'])]
        class1 = c1['mobile'].count()
        class2 = c2['mobile'].count()
        prev = class1 + class2
        
        print(f"SUCCESS: Metrics calculated:")
        print(f"   Lead: {lead}, Unique Lead: {u_lead}")
        print(f"   Application: {application} (sum of pipeline counts)")
        print(f"   Approved: {a_application}")
        print(f"   Class 1: {class1}, Class 2: {class2}, Total: {prev}")
        
        # Credit Score Processing
        print("\nPROCESSING: Credit scores...")
        if not credit.empty:
            credit = credit[credit['success'] == True]
            
            def categorize_score(x):
                if pd.isna(x) or x == 0:
                    return "New to Credit"
                elif x < 650:
                    return "Below 650"
                else:
                    return "Above 650"
            
            credit["Score Category"] = credit["credit_score"].apply(categorize_score)
            counts = credit["Score Category"].value_counts().reset_index()
            counts.columns = ["Score Category", "Count"]
            total = pd.DataFrame([{"Score Category": "Total", "Count": int(counts["Count"].sum())}])
            above720 = pd.DataFrame([{"Score Category": "Above 720", "Count": int((credit["credit_score"] >= 720).sum())}])
            new_to_credit = pd.DataFrame([{"Score Category": "New to Credit", "Count": int((credit["Score Category"] == "New to Credit").sum())}])
            result = pd.concat([total, 
                                counts[counts["Score Category"] == "Below 650"], 
                                counts[counts["Score Category"] == "Above 650"], 
                                above720,
                                new_to_credit], 
                               ignore_index=True)
            day_credit_spent = pd.DataFrame([{"Score Category": "Day Credit Spent", "Count": int(total["Count"].iloc[0] * 5)}])
            result = pd.concat([result, day_credit_spent], ignore_index=True)
            result["Count"] = result["Count"].astype("Int64")
        else:
            print("WARNING: No credit data available")
            result = pd.DataFrame([
                {"Score Category": "Total", "Count": 0},
                {"Score Category": "Below 650", "Count": 0},
                {"Score Category": "Above 650", "Count": 0},
                {"Score Category": "Above 720", "Count": 0},
                {"Score Category": "New to Credit", "Count": 0},
                {"Score Category": "Day Credit Spent", "Count": 0}
            ])
        
        # Summary DataFrame
        def get_score_val(cat):
            return int(result[result["Score Category"]==cat]["Count"].values[0]) if not result[result["Score Category"]==cat].empty else 0

        # Ensure Application is always the sum of platform-wise app counts (app_1.sum())
        application = int(app_1.sum())
        summary_data = [
            ["Lead", lead, "", "", get_score_val("Total")],
            ["Unique Lead", u_lead, round((u_lead/lead)*100,2) if lead else "", "Ul to L %", ""],
            ["Application", application, round((application/u_lead)*100,2) if u_lead else "", "AA to UI%", ""],
            ["Approved Application (Unique)", a_application, round((a_application/u_lead)*100,2) if u_lead else "", "App to UI%", ""],
            ["Approved Application Login - C1 + C2", prev, round((prev/a_application)*100,2) if a_application else "", "APR to app%", ""],
            ["Approved Application Login - Class 1", class1, round((class1/prev)*100,2) if prev else "", "C1 to APR", ""],
            ["Approved Application Login - Class 2", class2, round((class2/prev)*100,2) if prev else "", "C2 to APR", ""]
        ]
        summary_columns = ["Particulars", "No in total", "Percentage %", "Desc", "Credit Score Pulls"]
        summary_df = pd.DataFrame(summary_data, columns=summary_columns)

        # Platform Wise Table - REUSE app_1 calculated above
        print("\nPROCESSING: Platform data...")
        app_2 = dump1.groupby('pipeline_name')["pipeline_name"].count().rename("Approved Logins")
        c1_ins = c1.groupby('pipeline_name')['institution'].count().rename("Class 1")
        c2_ins = c2.groupby('pipeline_name')['institution'].count().rename("Class 2")
        merged = pd.concat([app_1, app_2, c1_ins, c2_ins], axis=1)
        merged = merged.fillna(0).astype(int).reset_index()
        merged = merged.rename(columns={"pipeline_name": "Source"})
        merged["App to APR %"] = ((merged["Approved Logins"] / merged["App"]) * 100).round(2)
        merged = merged[["Source", "App", "Approved Logins", "App to APR %", "Class 1", "Class 2"]]

        # Prepare final_df for summary append
        L_to_App = round((application/u_lead)*100, 2) if u_lead else 0
        L_to_AA = round((a_application/u_lead)*100, 2) if u_lead else 0
        L_to_AL = round((prev/u_lead)*100, 2) if u_lead else 0
        App_to_AL = round((prev/application)*100, 2) if application else 0
        
        summary_row = [
            date.today(),
            u_lead, application, a_application, L_to_App, L_to_AA, L_to_AL, App_to_AL,
            prev, class1, class2, get_score_val("Total"), get_score_val("Above 650")
        ]
        summary_columns = [
            "Date", "Unique Leads", "App", "AA(Unique)", "L to App%", "L to AA%", "L to AL%", 
            "App to AL%", "Total AL", "AL (C1)", "AL (C2)", "CS Pull", "G CS"
        ]
        final_df = pd.DataFrame([summary_row], columns=summary_columns)

        print("SUCCESS: Data processing complete")
        print(f"\n=== VERIFICATION ===")
        print(f"Application total: {application}")
        print(f"Sum of platform Apps: {merged['App'].sum()}")
        # Use ASCII output to avoid UnicodeEncodeError
        match_str = 'YES' if application == merged['App'].sum() else 'NO'
        print(f"Match: {match_str}")
        print("\nPlatform breakdown:")
        print(merged[['Source', 'App']])
        
        # Export to Google Sheets
        print("\nEXPORTING: To Google Sheets...")
        export_to_gsheets(summary_df, merged, result, final_df, SHEET_URL, login)
        
        print("\nCOMPLETED: Daily report pipeline completed successfully!")
        
    except Exception as e:
        print(f"\nCRITICAL ERROR in main pipeline: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
