# Garmin to Google Sheets Sync (Forma Template)

This script automates the synchronization of health, training, and race prediction data from Garmin Connect to a specific, formatted Google Sheets template ("Forma").

## Features

- **Search-Based Mapping**: Dynamically locates date blocks in the spreadsheet and populates data based on Polish labels.
- **Race Prediction Logic**: Handles "PRZED" (before) and "PO" (after) logic. Pulled from the previous day and current day respectively.
- **Training Status**: Syncs your overall training status (e.g., Peaking, Productive, Strained) for both "PRZED" and "PO" columns.
- **Activity Aggregation**: Sums distance, duration, and training load for multiple activities of the same type on a single day.
- **Health Metrics**: Exports HRV, resting heart rate, and sleep hours (in H:MM format).
- **Activity Links**: Adds a link to the first Garmin activity of the day under "LINK Z ZEGARKA".
- **Safe Updates**: Only writes data to empty cells (or cells containing placeholder zeros like `0:00:00` or `0,0`) to avoid overwriting manual adjustments.

## Prerequisites

### 1. Garmin Connect Account
You need your Garmin Connect email and password.

### 2. Google Service Account
To interact with Google Sheets, you need a Service Account:
1.  Go to the [Google Cloud Console](https://console.cloud.google.com/).
2.  Create a new project.
3.  Enable the **Google Sheets API**.
4.  Navigate to **IAM & Admin > Service Accounts**.
5.  Create a Service Account and download the **JSON Key file**.
6.  **Important**: Open the JSON file and find the `client_email` address. Go to your "Forma" Google Sheet and share it with this email address as an **Editor**.

### 3. Environment Variables
The script uses environment variables for configuration. You can set these in your shell or use a `.env` file (if you add `python-dotenv` support).

| Variable | Description |
| :--- | :--- |
| `EMAIL` | Your Garmin Connect email |
| `PASSWORD` | Your Garmin Connect password |
| `GOOGLE_APPLICATION_CREDENTIALS` | Absolute path to your Service Account JSON key file |
| `SPREADSHEET_ID` | The ID of your Google Sheet (found in the URL) |
| `GARMINTOKENS` | (Optional) Path to store Garmin session tokens (defaults to `~/.garminconnect`) |

## Installation

1.  Clone this repository.
2.  Install the required Python packages:

```bash
pip install gspread google-auth garminconnect cloudscraper garth
```

## Usage

Run the script using Python:

```bash
python3 main.py
```

The script will:
1.  Log in to Garmin Connect.
2.  Fetch sleep, activity, training status, and race prediction data for the last 7 days.
3.  Scan your Google Sheet worksheets for matching dates.
4.  Populate the data into the correct cells, providing detailed logs in the terminal.

## Template Sensitivity

The script relies on specific Polish labels in your spreadsheet to find the correct rows:
- `STATUS`
- `5 KM`, `10 KM`, `21 KM`, `42 KM`
- `BIEG`, `ROWER`, `SPACER`, `PŁYWANIE`, `JOGA`, `INNE`
- `HRV (ZMIENNOŚĆ TĘTNA)`, `TĘTNO SPOCZYNKOWE`, `ILOŚĆ SNU`, `LINK Z ZEGARKA`

If these labels are renamed or moved such that the column search logic fails, the script will skip those metrics.
