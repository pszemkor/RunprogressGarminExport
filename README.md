# Garmin to Google Sheets Sync 

This script automates the synchronization of health, training, and race prediction data from Garmin Connect to a specific, formatted Google Sheets template.

## Features

- **Search-Based Mapping**: Dynamically locates date blocks in the spreadsheet and populates data based on Polish labels.
- **Race Prediction Logic**: Handles "PRZED" (before) and "PO" (after) logic. Pulled from the previous day and current day respectively.
- **Training Status**: Syncs your overall training status (e.g., Peaking, Productive, Strained) for both "PRZED" and "PO" columns.
- **Activity Aggregation**: Sums distance, duration, and training load for multiple activities of the same type on a single day.
- **Health Metrics**: Exports HRV, resting heart rate, and sleep hours (in H:MM format).
- **Weekly HR Zones**: Automatically aggregates your heart rate zones across all activities for every completed week and updates the corresponding "Strefa tętna" block.
- **Activity Links**: Adds a link to the first Garmin activity of the day under "LINK Z ZEGARKA".
- **Subjective Feel**: Extracts how you felt during your primary activity (RPE/Feel) and maps it to your spreadsheet's emojis (e.g. `great 😄`, `poor 😕`).
- **Activity RPE**: Extracts the Rate of Perceived Exertion (RPE) on a scale of 1-10 for each individual activity type and pushes it to the spreadsheet's dropdown.
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
6.  Go to your Google Sheet and share it with the service account email address as an **Editor**.

### 3. Environment Variables
The script uses environment variables for configuration. You can set these in your shell or use a `.env` file (if you add `python-dotenv` support).

| Variable | Description |
| :--- | :--- |
| `GARMIN_PERSONAL_API_KEY` | Official Garmin Connect Personal API key (when using Personal API mode) |
| `USE_PERSONAL_API` | Set to `true` or `1` to default to Personal API instead of `garminconnect` |
| `EMAIL` | Your Garmin Connect email (for `garminconnect` library mode) |
| `PASSWORD` | Your Garmin Connect password (for `garminconnect` library mode) |
| `GOOGLE_APPLICATION_CREDENTIALS` | Absolute path to your Service Account JSON key file (optional on GCP) |
| `SPREADSHEET_ID` | The ID of your Google Sheet (found in the URL) |
| `GARMINTOKENS` | (Optional) Path to store Garmin session tokens (defaults to `~/.garminconnect`) |

## Installation

1. Clone this repository.
2. Install dependencies:

```bash
pip install -r requirements.txt
```

## Usage

### 1. Using Official Garmin Personal API (Recommended for Cloud & Automation)
Obtain your personal API key from [Garmin Connect Personal API](https://connect.garmin.com/app/settings/personalApiToken).

```bash
# Via flag and env var:
export GARMIN_PERSONAL_API_KEY="your-api-key"
python3 main.py --use-personal-api --days 7

# Or pass key directly:
python3 main.py --use-personal-api --api-key "your-api-key" --days 7
```

> **Note on Personal API coverage**: The official Personal API currently provides activities (with training load, duration, distance, subjective feel/RPE, and time-in-zones), sleep, resting heart rate, and overnight HRV. It does not expose Training Status or Race Predictions; those cells will remain empty.

### 2. Using garminconnect Python Library (Default)

```bash
# Sync data for the default 7 past days
python3 main.py

# Sync data for a specific number of past days
python3 main.py --days 14
```

## Google Cloud Platform Deployment (Cloud Functions + Cloud Build)

This repo is configured for automated deployment to **Google Cloud Functions (gen2)** triggered every night at **23:00 Polish time** (Europe/Warsaw) via **Cloud Scheduler**.

### Quick Deployment via Script
```bash
export PROJECT_ID="garminsync-491522"
export GARMIN_PERSONAL_API_KEY="your-api-key"
./deploy_gcp.sh
```

### Cloud Build CI/CD Integration
The included `cloudbuild.yaml` runs tests and deploys the Cloud Function automatically:

```bash
gcloud builds submit --config=cloudbuild.yaml
```

Remember to grant **Editor** access on your Google Sheet to the Cloud Function's runtime service account (`garmin-sync-runner@garminsync-491522.iam.gserviceaccount.com`).

## Template Sensitivity

The script relies on specific Polish labels in your spreadsheet to find the correct rows:
- `STATUS`
- `5 KM`, `10 KM`, `21 KM`, `42 KM`
- `BIEG`, `ROWER`, `SPACER`, `PŁYWANIE`, `JOGA`, `INNE`
- `HRV (ZMIENNOŚĆ TĘTNA)`, `TĘTNO SPOCZYNKOWE`, `ILOŚĆ SNU`, `LINK Z ZEGARKA`

If these labels are renamed or moved such that the column search logic fails, the script will skip those metrics.
