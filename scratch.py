import datetime, json, os
import gspread
from google.oauth2.service_account import Credentials

def main():
    creds_path = os.getenv("GOOGLE_APPLICATION_CREDENTIALS")
    if not creds_path:
        print("No creds")
        return
    credentials = Credentials.from_service_account_file(
        creds_path,
        scopes=["https://www.googleapis.com/auth/spreadsheets"]
    )
    client = gspread.authorize(credentials)
    sheet = client.open_by_key("1r20OoEoFyaAEcm3K-1GfzGCLxc6zk3FqGg823S7xYlA")
    worksheet = sheet.worksheet("2026 rok v2")
    
    # Dump row 1509 to 1512 for columns 21 to 30 (zero indexed: 20 to 30)
    for r in range(1508, 1512):
        row_vals = worksheet.row_values(r)
        # Pad row_vals if it's too short
        row_vals += [''] * max(0, 31 - len(row_vals))
        print(f"Row {r}:", row_vals[20:30])

if __name__ == "__main__":
    main()
