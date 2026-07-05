#!/usr/bin/env python3
from __future__ import annotations

import argparse
import datetime
import json
import os
import sys
from getpass import getpass

from garth.exc import GarthException, GarthHTTPError
from garminconnect import (
    Garmin,
    GarminConnectAuthenticationError,
    GarminConnectConnectionError,
    GarminConnectTooManyRequestsError,
)

def get_mfa() -> str:
    """Get MFA token."""
    return input("MFA one-time code: ")

def init_api(email: str | None = None, password: str | None = None, tokenstore: str = "~/.garminconnect") -> Garmin | None:
    """Initialize Garmin API with smart error handling and recovery."""
    tokenstore = os.path.expanduser(tokenstore)

    # First try to login with stored tokens
    try:
        print(f"Attempting to login using stored tokens from: {tokenstore}")
        garmin = Garmin()
        garmin.login(tokenstore)
        print("Successfully logged in using stored tokens!")
        return garmin
    except (FileNotFoundError, GarthHTTPError, GarminConnectAuthenticationError, GarminConnectConnectionError):
        print("No valid tokens found. Requesting fresh login credentials.")

    # Loop for credential entry with retry on auth failure
    try:
        # Get credentials if not provided
        if not email or not password:
            email = input("Email address: ").strip()
            password = getpass("Password: ")

        print("Logging in with credentials...")
        garmin = Garmin(
            email=email, password=password, is_cn=False, return_on_mfa=True
        )
        
        # Override the default requests.Session with Cloudscraper to bypass Garmin's Cloudflare checks
        try:
            import cloudscraper
            scraper = cloudscraper.create_scraper()
            # Copy over user agent added by garth setup
            scraper.headers.update(garmin.garth.sess.headers)
            garmin.garth.sess = scraper
        except ImportError:
            print("💡 Please run 'pip install cloudscraper' if you hit 429 Too Many Requests errors.")

        result1, result2 = garmin.login()

        if result1 == "needs_mfa":
            print("Multi-factor authentication required")
            mfa_code = get_mfa()
            print("🔄 Submitting MFA code...")
            try:
                garmin.resume_login(result2, mfa_code)
                print("✅ MFA authentication successful!")
            except Exception as e:
                print(f"❌ MFA error: {e}")
                sys.exit(1)

        # Save tokens for future use
        token_dir = os.path.dirname(tokenstore)
        if token_dir:
            os.makedirs(token_dir, exist_ok=True)
        garmin.garth.dump(tokenstore)
        print(f"Login successful! Tokens saved to: {tokenstore}")

        return garmin
    except GarminConnectAuthenticationError as err:
        print(f"\n❌ Authentication error: {err}")
        print("💡 Please check your username and password and try again")
        return None
    except Exception as err:
        print(f"❌ Connection error: {err}")
        return None

def export_to_google_sheets(spreadsheet_id: str, sleep_data: dict, training_data: list, status_data: dict, preds_data: dict, weekly_hr_zones: dict = None):
    try:
        import gspread
        from google.oauth2.service_account import Credentials
    except ImportError:
        print("💡 Please run 'pip install gspread google-auth' to enable Google Sheets export.")
        return

    creds_path = os.getenv("GOOGLE_APPLICATION_CREDENTIALS")
    if not creds_path or not os.path.exists(creds_path):
        print("⚠️ GOOGLE_APPLICATION_CREDENTIALS env var is missing or invalid. Skipping Google Sheets export.")
        print("   Set it to the path of your Service Account JSON file.")
        return

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    
    print(f"\n--------------------------------------------------")
    print(f"Exporting data to Forma Google Spreadsheet: {spreadsheet_id}...")
    try:
        creds = Credentials.from_service_account_file(creds_path, scopes=scopes)
        client = gspread.authorize(creds)
        spreadsheet = client.open_by_key(spreadsheet_id)
        
        dates_to_find = list(sleep_data.keys())
        
        for ws in spreadsheet.worksheets():
            try:
                data = ws.get_all_values()
            except Exception as e:
                print(f"Skipping sheet {ws.title} due to error reading: {e}")
                continue
                
            if not data: continue
            batch_updates = []
            
            for date_str in dates_to_find:
                found = False
                r_idx = -1
                c_idx = -1
                
                # Search for the date string in the sheet
                for i, row in enumerate(data):
                    for j, cell_val in enumerate(row):
                        if date_str in str(cell_val):
                            r_idx = i
                            c_idx = j
                            found = True
                            break
                    if found: break
                    
                if not found:
                    continue
                    
                print(f"Found {date_str} in worksheet '{ws.title}' at cell ({r_idx+1}, {c_idx+1})")
                
                # Pre-calculate what we need for this date
                date_obj = datetime.datetime.strptime(date_str, "%Y-%m-%d")
                prev_date_str = (date_obj - datetime.timedelta(days=1)).strftime("%Y-%m-%d")
                
                daily_sleep = sleep_data.get(date_str, {})
                daily_preds = preds_data.get(date_str, {})
                prev_preds = preds_data.get(prev_date_str, {})
                
                # Accumulate activities for the day (metrics + first activity link)
                daily_activities = {}
                daily_activity_links = []  # All activity URLs for the day, in order
                for act in training_data:
                    if act.get("startTime", "").startswith(date_str):
                        t = act.get("activityType", "other")
                        if t not in daily_activities:
                            daily_activities[t] = {"distance": 0, "duration": 0, "load": 0}
                        daily_activities[t]["distance"] += act.get("distanceKm") or 0
                        daily_activities[t]["duration"] += act.get("durationMinutes") or 0
                        daily_activities[t]["load"] += act.get("trainingLoad") or 0
                        act_id = act.get("activityId")
                        if act_id:
                            daily_activity_links.append(f"https://connect.garmin.com/app/activity/{act_id}")

                # Track what we've updated for this day block to ensure we don't over-write if we scan too far
                updated_preds = {"PRZED": set(), "PO": set()}
                updated_acts = set()
                updated_health = set()
                # Gate: only match activity labels after "RODZAJ AKTYWNOŚCI" header is found,
                # preventing false matches on free-form cells like "BIEG DŁUGI"
                in_activity_section = False

                # Scan downwards inside this day's column block
                max_row = min(r_idx + 45, len(data))  # Reduced max_row to 45 to avoid hitting next week
                for i in range(r_idx + 1, max_row):
                    row = data[i]
                    c_val = str(row[c_idx]).upper().strip() if c_idx < len(row) else ""
                    c3_val = str(row[c_idx+3]).upper().strip() if (c_idx+3) < len(row) else ""

                    # Detect the activity section header
                    if "RODZAJ" in c_val and "AKTYWNO" in c_val:
                        in_activity_section = True
                        continue

                    # Helper function to check if cell is effectively empty (including placeholder zeros)
                    def is_empty(col_offset):
                        c_str = str(row[c_idx+col_offset]).strip() if (c_idx+col_offset) < len(row) else ""
                        return c_str == "" or c_str in ["0", "0.0", "0,0", "0:00", "0:00:00", "0.00", "0,00", "-"]

                    def queue_update(label, val, row_idx, col_offset):
                        a1_range = gspread.utils.rowcol_to_a1(row_idx+1, c_idx+col_offset+1)
                        print(f"  📝 [{ws.title} | {date_str}] {label} -> {a1_range} (Value: {val})")
                        batch_updates.append({'range': a1_range, 'values': [[val]]})

                    # 1. Training Status (PRZED = c_idx column, PO = c_idx+3 column)
                    if c_val == "STATUS" and "STATUS_PRZED" not in updated_health:
                        val = normalize_training_status(status_data.get(prev_date_str, {}).get("status", ""))
                        if val and is_empty(2):
                            queue_update("PRZED Status", val, i, 2)
                        updated_health.add("STATUS_PRZED")
                    if c_val == "STATUS" and "STATUS_PO" not in updated_health:
                        val = normalize_training_status(status_data.get(date_str, {}).get("status", ""))
                        if val and is_empty(5):
                            queue_update("PO Status", val, i, 5)
                        updated_health.add("STATUS_PO")

                    # 2. Race Predictions (PRZED = c_idx, PO = c_idx+2)
                    # Use a more flexible string matching by removing spaces
                    for dist_label, key in [("5 KM", "5K"), ("10 KM", "10K"), ("21 KM", "HalfMarathon"), ("42 KM", "Marathon")]:
                        search_label = dist_label.replace(" ", "")
                        if search_label in c_val.replace(" ", "") and key not in updated_preds["PRZED"]:
                            val = prev_preds.get(key)
                            if val and is_empty(2):
                                queue_update(f"PRZED {dist_label}", val, i, 2)
                                updated_preds["PRZED"].add(key)
                        if search_label in c3_val.replace(" ", "") and key not in updated_preds["PO"]:
                            val = daily_preds.get(key)
                            if val and is_empty(5):
                                queue_update(f"PO {dist_label}", val, i, 5)
                                updated_preds["PO"].add(key)

                    # 2. Activities — only match inside the RODZAJ AKTYWNOŚCI section
                    if in_activity_section:
                        act_mapped, d, act_name = False, None, ""
                        if "BIEG" in c_val and "BIEG" not in updated_acts:
                            d, act_mapped, act_name = daily_activities.get("running"), True, "BIEG"
                        elif "ROWER" in c_val and "ROWER" not in updated_acts:
                            d, act_mapped, act_name = daily_activities.get("cycling"), True, "ROWER"
                        elif "SPACER" in c_val and "SPACER" not in updated_acts:
                            d, act_mapped, act_name = daily_activities.get("walking"), True, "SPACER"
                        elif "PŁYWANIE" in c_val and "PŁYWANIE" not in updated_acts:
                            d, act_mapped, act_name = daily_activities.get("swimming"), True, "PŁYWANIE"
                        elif "JOGA" in c_val and "JOGA" not in updated_acts:
                            d, act_mapped, act_name = daily_activities.get("yoga"), True, "JOGA"
                        elif "INNE" in c_val and "INNE" not in updated_acts:
                            other_keys = [k for k in daily_activities.keys() if k not in ["running", "cycling", "walking", "swimming", "yoga"]]
                            d = {"distance": 0, "duration": 0, "load": 0}
                            for k in other_keys:
                                d["distance"] += daily_activities[k]["distance"]
                                d["duration"] += daily_activities[k]["duration"]
                                d["load"] += daily_activities[k]["load"]
                            if d["duration"] == 0: d = None
                            act_mapped, act_name = True, "INNE"

                        if act_mapped and d:
                            updated_acts.add(act_name)
                            if d["distance"] > 0 and is_empty(1):
                                queue_update(f"{act_name} Distance", round(d["distance"], 2), i, 2)
                            if d["duration"] > 0 and is_empty(2):
                                queue_update(f"{act_name} Duration", round(d["duration"], 2), i, 3)
                            if d["load"] > 0 and is_empty(5):
                                queue_update(f"{act_name} Load", round(d["load"], 2), i, 5)
                            else: 
                                print("Load: " + str(row[c_idx+5]).strip())

                    # 3. Health Metrics
                    if "LINK" in c_val and "ZEGARKA" in c_val and "LINK" not in updated_health:
                        if daily_activity_links and is_empty(3):
                            # Write all activity links joined by newlines
                            links_str = "\n".join(daily_activity_links)
                            queue_update("Link z Zegarka", links_str, i, 3)
                            updated_health.add("LINK")
                    elif "HRV" in c_val and "ZMIENNOŚĆ" in c_val and "HRV" not in updated_health:
                        val = daily_sleep.get("avgOvernightHrv")
                        if val and is_empty(2):
                            queue_update("HRV", val, i, 3)
                            updated_health.add("HRV")
                    elif "TĘTNO SPOCZYNKOWE" in c_val and "RHR" not in updated_health:
                        val = daily_sleep.get("restingHeartRate")
                        if val and is_empty(2):
                            queue_update("Resting HR", val, i, 3)
                            updated_health.add("RHR")
                    elif "ILOŚĆ SNU" in c_val and "SLEEP" not in updated_health:
                        val = daily_sleep.get("sleepTimeHours")
                        if val and is_empty(2):
                            queue_update("Sleep Hours", hours_to_time_str(val), i, 3)
                            updated_health.add("SLEEP")
                            
                # 4. Weekly HR Zones (only if it's Sunday)
                if weekly_hr_zones and date_obj.weekday() == 6:
                    zones = weekly_hr_zones.get(date_str)
                    if zones:
                        hr_r_idx = -1
                        hr_c_idx = -1
                        for i in range(r_idx, min(r_idx + 45, len(data))):
                            row = data[i]
                            for j in range(c_idx + 1, len(row)):
                                if "STREFA TĘTNA" in str(row[j]).upper():
                                    hr_r_idx = i
                                    hr_c_idx = j
                                    break
                            if hr_r_idx != -1:
                                break
                                
                        if hr_r_idx != -1:
                            print(f"  Found 'Strefa tętna' table for {date_str} at cell ({hr_r_idx+1}, {hr_c_idx+1})")
                            for z in range(5, 0, -1):
                                search_str = f"STREFA {z}"
                                for i in range(hr_r_idx + 1, min(hr_r_idx + 10, len(data))):
                                    row = data[i]
                                    if hr_c_idx < len(row) and search_str in str(row[hr_c_idx]).upper():
                                        secs = zones.get(z, 0.0)
                                        h, m = divmod(int(secs), 3600)
                                        m, s = divmod(m, 60)
                                        time_str = f"{h:02d}:{m:02d}:{s:02d}"
                                        
                                        target_c = hr_c_idx + 2
                                        queue_update(f"Strefa {z} Time", time_str, i, target_c - c_idx)
                                        break
                                        
            if batch_updates:
                try:
                    ws.batch_update(batch_updates, value_input_option='USER_ENTERED')
                    print(f"✅ Applied {len(batch_updates)} updates to worksheet '{ws.title}'")
                except Exception as e:
                    print(f"❌ Failed to batch update {ws.title}: {e}")

        print("✅ Finished updating Google Sheets!")

    except Exception as e:
        print(f"❌ Failed to export to Google Sheets: {e}")

def hours_to_time_str(decimal_hours: float) -> str:
    """Convert decimal hours (e.g. 5.77) to H:MM string (e.g. '5:46')."""
    if not decimal_hours:
        return ""
    total_minutes = round(decimal_hours * 60)
    h, m = divmod(total_minutes, 60)
    return f"{h}:{m:02d}"

def format_time(seconds):
    if not seconds: return None
    m, s = divmod(seconds, 60)
    h, m = divmod(m, 60)
    return f"{int(h):02d}:{int(m):02d}:{int(s):02d}"

def device_data_phrase_format(phrase):
    return phrase.replace("_", " ").title() if phrase else "Unknown"

def fetch_sleep_data(api, today, days=7):
    print(f"\nFetching sleep data for the past {days} days (ending {today.isoformat()})...")
    all_sleep_data = {}
    for i in range(days):
        target_date = today - datetime.timedelta(days=i)
        date_str = target_date.isoformat()
        print(f"Fetching data for {date_str}...")
        try:
            sleep_data = api.get_sleep_data(date_str)
            if not sleep_data:
                all_sleep_data[date_str] = {"error": "No data returned"}
                continue
                
            daily_dto = sleep_data.get("dailySleepDTO", {})
            
            calendar_date = daily_dto.get("calendarDate", date_str)
            sleep_time_seconds = daily_dto.get("sleepTimeSeconds")
            sleep_time_hours = round(sleep_time_seconds / 3600.0, 2) if sleep_time_seconds is not None else None
            
            resting_hr = sleep_data.get("restingHeartRate")
            avg_hrv = sleep_data.get("avgOvernightHrv")

            all_sleep_data[date_str] = {
                "calendarDate": calendar_date,
                "restingHeartRate": resting_hr,
                "avgOvernightHrv": avg_hrv,
                "sleepTimeHours": sleep_time_hours
            }
        except Exception as e:
            print(f"Failed to fetch sleep data for {date_str}: {e}")
            all_sleep_data[date_str] = {"error": str(e)}

    print("\nSleep data retrieved successfully!")
    print(json.dumps(all_sleep_data, indent=4))
    return all_sleep_data

def fetch_training_data(api, start_date, end_date):
    print("\n--------------------------------------------------")
    print(f"Fetching training/activity data (from {start_date.isoformat()} to {end_date.isoformat()})...")

    all_training_data = []
    try:
        activities = api.get_activities_by_date(start_date.isoformat(), end_date.isoformat())
        for act in activities:
            duration_mins = round((act.get("duration") or 0) / 60.0, 2)
            distance_km = round((act.get("distance") or 0) / 1000.0, 2)
            
            train_load = act.get("activityTrainingLoad")
            
            act_info = {
                "activityId": act.get("activityId"),
                "activityName": act.get("activityName"),
                "activityType": act.get("activityType", {}).get("typeKey"),
                "startTime": act.get("startTimeLocal"),
                "durationMinutes": duration_mins,
                "distanceKm": distance_km,
                "calories": act.get("calories"),
                "averageHeartRate": act.get("averageHR"),
                "trainingLoad": round(train_load, 2) if train_load is not None else None,
                "trainingEffect": act.get("trainingEffectLabel")
            }
            all_training_data.append(act_info)
            
        print("\nTraining data retrieved successfully!")
        print(json.dumps(all_training_data, indent=4))
    except Exception as e:
        print(f"Failed to fetch training data: {e}")
    return all_training_data

# Maps Garmin training status phrases to the exact dropdown values in the sheet
_STATUS_DROPDOWN_VALUES = [
    "Peaking", "Productive", "Maintaining", "Recovery",
    "Unproductive", "Overreaching", "Strained", "Detraining"
]

def normalize_training_status(raw_status: str) -> str:
    """Map a Garmin status phrase (e.g. 'Strained 3') to a sheet dropdown value."""
    if not raw_status:
        return ""
    raw_upper = raw_status.upper()
    for option in _STATUS_DROPDOWN_VALUES:
        if option.upper() in raw_upper:
            return option
    return ""

def fetch_training_status(api, today: datetime.date, days: int = 7) -> dict:
    """Fetch training status for each of the last `days` days, keyed by ISO date string."""
    print("\n--------------------------------------------------")
    print(f"Fetching overall Training Status for the past {days} days (ending {today.isoformat()})...")
    all_status: dict = {}
    for i in range(days):
        target_date = today - datetime.timedelta(days=i)
        date_str = target_date.isoformat()
        try:
            ts = api.get_training_status(date_str)
            if ts and "mostRecentTrainingStatus" in ts:
                latest_ts_data = ts["mostRecentTrainingStatus"].get("latestTrainingStatusData", {})
                if latest_ts_data:
                    device_data = list(latest_ts_data.values())[0]
                    acute_load = device_data.get("acuteTrainingLoadDTO", {}).get("dailyTrainingLoadAcute")
                    chronic_load = device_data.get("acuteTrainingLoadDTO", {}).get("dailyTrainingLoadChronic")
                    phrase = device_data.get("trainingStatusFeedbackPhrase", "Unknown")
                    formatted_phrase = device_data_phrase_format(phrase)
                    print(f"  {date_str}: {formatted_phrase} (Acute: {acute_load}, Chronic: {chronic_load})")
                    all_status[date_str] = {
                        "status": formatted_phrase,
                        "acute_load": acute_load,
                        "chronic_load": chronic_load
                    }
                else:
                    print(f"  {date_str}: No data")
        except Exception as e:
            print(f"  {date_str}: Failed - {e}")
    return all_status

def fetch_race_predictions(api, start_date, end_date):
    print("\n--------------------------------------------------")
    print(f"Fetching race predictions (from {start_date.isoformat()} to {end_date.isoformat()})...")
    formatted_preds = {}
    try:
        predictions = api.get_race_predictions(start_date.isoformat(), end_date.isoformat(), 'daily')
        if predictions:
            for p in predictions:
                date_key = p.get("calendarDate")
                if not date_key: continue
                
                formatted_preds[date_key] = {
                    "5K": format_time(p.get("time5K")),
                    "10K": format_time(p.get("time10K")),
                    "HalfMarathon": format_time(p.get("timeHalfMarathon")),
                    "Marathon": format_time(p.get("timeMarathon"))
                }
            print("\nRace predictions retrieved successfully!")
            print(json.dumps(formatted_preds, indent=4))
        else:
            print("No race prediction data found.")
    except Exception as e:
        print(f"Failed to fetch race predictions: {e}")
    return formatted_preds

_HR_ZONE_CACHE = {}

def get_aggregated_hr_zones(api, training_data, start_date, end_date, quiet=False):
    aggregated_zones = {1: 0.0, 2: 0.0, 3: 0.0, 4: 0.0, 5: 0.0}
    has_hr_data = False

    for act in training_data:
        act_id = act.get("activityId")
        act_name = act.get("activityName", "Unknown Activity")
        act_type = act.get("activityType", "other")
        start_time_str = act.get("startTime")
        
        if not act_id or not start_time_str:
            continue
            
        try:
            act_date = datetime.datetime.strptime(start_time_str[:10], "%Y-%m-%d").date()
        except ValueError:
            continue
            
        if not (start_date <= act_date <= end_date):
            continue
            
        if not quiet:
            print(f"Fetching HR zones for '{act_name}' ({act_type}) on {start_time_str}...")
        try:
            if act_id not in _HR_ZONE_CACHE:
                _HR_ZONE_CACHE[act_id] = api.get_activity_hr_in_timezones(act_id)
                
            hr_zones = _HR_ZONE_CACHE[act_id]
            if hr_zones:
                if not quiet:
                    print(f"  Data retrieved: {len(hr_zones)} zones found.")
                for zone in hr_zones:
                    z_num = zone.get("zoneNumber")
                    secs = zone.get("secsInZone", 0.0)
                    if z_num in aggregated_zones:
                        aggregated_zones[z_num] += secs
                        if secs > 0:
                            has_hr_data = True
            else:
                if not quiet:
                    print("  No HR zone data found for this activity.")
        except Exception as e:
            if not quiet:
                print(f"  Could not retrieve HR zones: {e}")

    return aggregated_zones if has_hr_data else None

def print_hr_zones(aggregated_zones, start_date, end_date):
    print("\n==================================================")
    print(f"AGGREGATED HEART RATE ZONES FROM {start_date.isoformat()} TO {end_date.isoformat()}")
    print("==================================================")
    if aggregated_zones:
        total_active_secs = sum(aggregated_zones.values())
        for z_num in sorted(aggregated_zones.keys(), reverse=True):
            total_secs = aggregated_zones[z_num]
            pct = (total_secs / total_active_secs * 100) if total_active_secs > 0 else 0.0
            
            h, m = divmod(int(total_secs), 3600)
            m, s = divmod(m, 60)
            time_str = f"{h:02d}:{m:02d}:{s:02d}"
            
            zone_desc = {
                1: "Zone 1 (Warm Up)",
                2: "Zone 2 (Easy/Aerobic)",
                3: "Zone 3 (Aerobic/Tempo)",
                4: "Zone 4 (Threshold)",
                5: "Zone 5 (Anaerobic/Max)"
            }.get(z_num, f"Zone {z_num}")
            
            bar_len = 20
            filled_len = int(round(bar_len * pct / 100))
            bar = "█" * filled_len + "░" * (bar_len - filled_len)
            
            print(f"  {zone_desc:<25} : {time_str:>8}  [{bar}] {pct:>5.1f}% ({total_secs:.1f}s)")
    else:
        print("  No heart rate zone data recorded in the exercises during this period.")
    print("==================================================")

def fetch_and_print_hr_zones(api, training_data, start_date, end_date):
    """Fetch HR zone time for each activity and print aggregated times for the [start_date, end_date] interval."""
    print("\n--------------------------------------------------")
    print(f"Fetching heart rate zone data for activities between {start_date.isoformat()} and {end_date.isoformat()} (inclusive)...")
    
    zones = get_aggregated_hr_zones(api, training_data, start_date, end_date)
    print_hr_zones(zones, start_date, end_date)
    return zones

def main():
    parser = argparse.ArgumentParser(description="Sync Garmin data to Google Sheets")
    parser.add_argument("--days", type=int, default=7, help="Number of past days to sync (default: 7)")
    parser.add_argument("--hr-start", type=str, help="Start date for aggregated heart rate zones (YYYY-MM-DD)")
    parser.add_argument("--hr-end", type=str, help="End date for aggregated heart rate zones (YYYY-MM-DD)")
    args = parser.parse_args()

    email = os.getenv("EMAIL")
    password = os.getenv("PASSWORD")
    tokenstore = os.getenv("GARMINTOKENS") or "~/.garminconnect"

    api = init_api(email, password, tokenstore)
    if not api:
        print("Failed to initialize Garmin API")
        sys.exit(1)

    today = datetime.date.today()
    start_date = today - datetime.timedelta(days=args.days)
    
    # Parse HR parameters if passed
    hr_start = None
    hr_end = None
    if args.hr_start or args.hr_end:
        try:
            hr_start = datetime.date.fromisoformat(args.hr_start) if args.hr_start else start_date
            hr_end = datetime.date.fromisoformat(args.hr_end) if args.hr_end else today
        except ValueError as e:
            print(f"❌ Invalid date format for --hr-start or --hr-end. Use YYYY-MM-DD. Error: {e}")
            sys.exit(1)

    # Determine date range for fetching activities to cover standard sync, custom HR zone, and weekly HR zones
    fetch_start = start_date
    fetch_end = today
    if hr_start and hr_start < fetch_start:
        fetch_start = hr_start
    if hr_end and hr_end > fetch_end:
        fetch_end = hr_end
        
    # Calculate dates for current week and previous week
    current_week_start = today - datetime.timedelta(days=today.weekday())
    current_week_end = today
    prev_week_start = current_week_start - datetime.timedelta(days=7)
    prev_week_end = current_week_start - datetime.timedelta(days=1)
    
    if prev_week_start < fetch_start:
        fetch_start = prev_week_start
    
    all_sleep_data = fetch_sleep_data(api, today, days=args.days)
    all_training_data = fetch_training_data(api, fetch_start, fetch_end)
    status_data = fetch_training_status(api, today, days=args.days)
    formatted_preds = fetch_race_predictions(api, start_date, today)

    # Always fetch and print heart rate zones for previous week and current week
    fetch_and_print_hr_zones(api, all_training_data, prev_week_start, prev_week_end)
    fetch_and_print_hr_zones(api, all_training_data, current_week_start, current_week_end)

    # Fetch and print heart rate zones for the exercises if custom parameters are passed
    if hr_start and hr_end:
        fetch_and_print_hr_zones(api, all_training_data, hr_start, hr_end)

    # Calculate weekly HR zones for Google Sheets export
    weekly_hr_zones_data = {}
    curr = fetch_start
    while curr <= fetch_end:
        if curr.weekday() == 6: # Sunday
            monday = curr - datetime.timedelta(days=6)
            if monday >= fetch_start:
                # We have a full week
                zones = get_aggregated_hr_zones(api, all_training_data, monday, curr, quiet=True)
                if zones:
                    weekly_hr_zones_data[curr.isoformat()] = zones
        curr += datetime.timedelta(days=1)

    spreadsheet_id = os.getenv("SPREADSHEET_ID")
    if spreadsheet_id:
        export_to_google_sheets(spreadsheet_id, all_sleep_data, all_training_data, status_data, formatted_preds, weekly_hr_zones_data)
    else:
        print("⚠️ SPREADSHEET_ID env var is missing. Skipping Google Sheets export.")

if __name__ == "__main__":
    main()
