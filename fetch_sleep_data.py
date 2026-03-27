#!/usr/bin/env python3

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

def main():
    email = os.getenv("EMAIL")
    password = os.getenv("PASSWORD")
    tokenstore = os.getenv("GARMINTOKENS") or "~/.garminconnect"

    api = init_api(email, password, tokenstore)
    if not api:
        print("Failed to initialize Garmin API")
        sys.exit(1)

    today = datetime.date.today()
    print(f"\nFetching sleep data for the past 7 days (ending {today.isoformat()})...")

    all_sleep_data = {}
    
    # Fetch data for the past 7 days (including today)
    for i in range(7):
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

    # Fetch training data for the past 7 days
    week_ago = today - datetime.timedelta(days=7)
    print("\n--------------------------------------------------")
    print(f"Fetching training/activity data for the past 7 days (from {week_ago.isoformat()} to {today.isoformat()})...")

    all_training_data = []
    try:
        activities = api.get_activities_by_date(week_ago.isoformat(), today.isoformat())
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

    # Fetch daily training status / load for today
    print("\n--------------------------------------------------")
    print(f"Fetching today's overall Training Status / Load for {today.isoformat()}...")
    try:
        ts = api.get_training_status(today.isoformat())
        if ts and "mostRecentTrainingStatus" in ts:
            latest_ts_data = ts["mostRecentTrainingStatus"].get("latestTrainingStatusData", {})
            if latest_ts_data:
                device_data = list(latest_ts_data.values())[0]
                acute_load = device_data.get("acuteTrainingLoadDTO", {}).get("dailyTrainingLoadAcute")
                chronic_load = device_data.get("acuteTrainingLoadDTO", {}).get("dailyTrainingLoadChronic")
                phrase = device_data.get("trainingStatusFeedbackPhrase", "Unknown")
                
                print(f"Overall Training Status: {DeviceDataPhraseFormat(phrase)}")
                print(f"Acute Training Load: {acute_load}")
                print(f"Chronic Training Load: {chronic_load}")
            else:
                print("No recent training status data found.")
    except Exception as e:
        print(f"Failed to fetch general training status: {e}")

    # Fetch daily race predictions
    print("\n--------------------------------------------------")
    print(f"Fetching race predictions for the past 7 days (from {week_ago.isoformat()} to {today.isoformat()})...")
    try:
        predictions = api.get_race_predictions(week_ago.isoformat(), today.isoformat(), 'daily')
        if predictions:
            formatted_preds = {}
            def format_time(seconds):
                if not seconds: return None
                m, s = divmod(seconds, 60)
                h, m = divmod(m, 60)
                return f"{int(h):02d}:{int(m):02d}:{int(s):02d}"
                
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

if __name__ == "__main__":
    def DeviceDataPhraseFormat(phrase):
        return phrase.replace("_", " ").title() if phrase else "Unknown"
    main()
