"""Google Cloud Functions entrypoint for GarminSync."""
from __future__ import annotations

import json
import os
import sys
import functions_framework

from main import (
    init_api,
    fetch_sleep_data,
    fetch_training_data,
    fetch_training_status,
    fetch_race_predictions,
    fetch_and_print_hr_zones,
    get_aggregated_hr_zones,
    export_to_google_sheets,
)
import datetime


def run_sync(days: int = 7, use_personal_api: bool = False, api_key: str | None = None) -> dict:
    """Core synchronization routine callable from CLI or Cloud Functions."""
    api = None
    personal_client = None

    if use_personal_api:
        from personal_api import (
            GarminPersonalApiClient,
            fetch_sleep_data_personal,
            fetch_training_data_personal,
        )

        api_key = api_key or os.getenv("GARMIN_PERSONAL_API_KEY")
        if not api_key:
            raise ValueError("GARMIN_PERSONAL_API_KEY is required for Personal API mode.")

        print("🔑 Initializing Garmin Connect Personal API client...")
        personal_client = GarminPersonalApiClient(api_key=api_key)
        personal_client.get_token()
        api = personal_client
    else:
        email = os.getenv("EMAIL")
        password = os.getenv("PASSWORD")
        tokenstore = os.getenv("GARMINTOKENS") or "~/.garminconnect"

        api = init_api(email, password, tokenstore)
        if not api:
            raise RuntimeError("Failed to initialize Garmin API")

    today = datetime.date.today()
    start_date = today - datetime.timedelta(days=days)

    current_week_start = today - datetime.timedelta(days=today.weekday())
    current_week_end = today
    prev_week_start = current_week_start - datetime.timedelta(days=7)
    prev_week_end = current_week_start - datetime.timedelta(days=1)

    fetch_start = min(start_date, prev_week_start)
    fetch_end = today

    if use_personal_api:
        from personal_api import fetch_sleep_data_personal, fetch_training_data_personal

        all_sleep_data = fetch_sleep_data_personal(personal_client, today, days=days)
        all_training_data = fetch_training_data_personal(personal_client, fetch_start, fetch_end)
        status_data = {}
        formatted_preds = {}
    else:
        all_sleep_data = fetch_sleep_data(api, today, days=days)
        all_training_data = fetch_training_data(api, fetch_start, fetch_end)
        status_data = fetch_training_status(api, today, days=days)
        formatted_preds = fetch_race_predictions(api, start_date, today)

    fetch_and_print_hr_zones(api, all_training_data, prev_week_start, prev_week_end)
    fetch_and_print_hr_zones(api, all_training_data, current_week_start, current_week_end)

    weekly_hr_zones_data = {}
    curr = fetch_start
    while curr <= fetch_end:
        if curr.weekday() == 6:
            monday = curr - datetime.timedelta(days=6)
            if monday >= fetch_start:
                zones = get_aggregated_hr_zones(api, all_training_data, monday, curr, quiet=True)
                if zones:
                    weekly_hr_zones_data[curr.isoformat()] = zones
        curr += datetime.timedelta(days=1)

    spreadsheet_id = os.getenv("SPREADSHEET_ID")
    if spreadsheet_id:
        export_to_google_sheets(
            spreadsheet_id,
            all_sleep_data,
            all_training_data,
            status_data,
            formatted_preds,
            weekly_hr_zones_data,
        )
    else:
        print("⚠️ SPREADSHEET_ID env var is missing. Skipping Google Sheets export.")

    return {
        "status": "success",
        "synced_days": days,
        "mode": "personal_api" if use_personal_api else "garminconnect",
        "activities_count": len(all_training_data),
    }


@functions_framework.http
def sync_garmin_http(request):
    """HTTP trigger entrypoint for Cloud Functions gen2 (invoked by Cloud Scheduler or curl)."""
    days = 7
    if request.args and "days" in request.args:
        try:
            days = int(request.args.get("days"))
        except ValueError:
            pass
    elif request.is_json and request.get_json(silent=True):
        body = request.get_json()
        if "days" in body:
            try:
                days = int(body["days"])
            except ValueError:
                pass

    use_personal_api = os.getenv("USE_PERSONAL_API", "true").lower() in ("1", "true", "yes")

    try:
        result = run_sync(days=days, use_personal_api=use_personal_api)
        return (json.dumps(result), 200, {"Content-Type": "application/json"})
    except Exception as e:
        import traceback

        traceback.print_exc()
        print(f"❌ Cloud Function sync failed: {e}")
        return (json.dumps({"status": "error", "message": str(e)}), 500, {"Content-Type": "application/json"})
