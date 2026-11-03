"""Garmin Connect Personal API client and sync adapter.

Official API Documentation:
https://developer.garmin.com/gc-developer-program/personal/personal-api
"""
from __future__ import annotations

import base64
import datetime
import json
import os
import time
from typing import Any, Dict, List, Optional
import requests

DEFAULT_BASE_URL = "https://personal.connectapi.garmin.com"
TOKEN_ENDPOINT = f"{DEFAULT_BASE_URL}/personalapitoken-service/personal-api/personal-api-key/token"


class GarminPersonalApiClient:
    """Client for official Garmin Connect Personal API."""

    def __init__(
        self,
        api_key: Optional[str] = None,
        base_url: str = DEFAULT_BASE_URL,
        bearer_token: Optional[str] = None,
        token_expires_at: Optional[float] = None,
    ):
        self.api_key = api_key or os.getenv("GARMIN_PERSONAL_API_KEY")
        self.base_url = base_url.rstrip("/")
        self._bearer_token = bearer_token
        self._token_expires_at = token_expires_at or 0.0
        self.session = requests.Session()

    def get_token(self, force_refresh: bool = False) -> str:
        """Obtain or return a valid OAuth2 bearer token from the personal API key."""
        now = time.time()
        # Keep a 5-minute safety margin before token expiry (tokens are valid up to 24h)
        if not force_refresh and self._bearer_token and (self._token_expires_at - now > 300):
            return self._bearer_token

        if not self.api_key:
            raise ValueError(
                "Garmin Personal API key is required. Set GARMIN_PERSONAL_API_KEY or pass --api-key."
            )

        headers = {
            "Authorization": f"ApiKey {self.api_key.strip()}",
            "User-Agent": "GarminSync/1.0",
        }
        resp = self.session.post(TOKEN_ENDPOINT, headers=headers, timeout=30)
        resp.raise_for_status()

        # Garmin Personal API token endpoint returns either a raw JWT string (text/plain) or JSON
        raw_text = resp.text.strip().strip('"')
        token = None
        expires_at = now + 86400  # Default 24h fallback

        try:
            data = resp.json()
            if isinstance(data, dict):
                token = data.get("access_token") or data.get("token")
                if "expires_in" in data:
                    expires_at = now + int(data["expires_in"])
            elif isinstance(data, str):
                token = data.strip().strip('"')
        except Exception:
            token = raw_text

        if not token:
            token = raw_text

        # If token is a JWT (header.payload.signature), extract actual exp claim
        parts = token.split(".")
        if len(parts) >= 2:
            try:
                payload_b64 = parts[1] + "=" * (-len(parts[1]) % 4)
                payload = json.loads(base64.urlsafe_b64decode(payload_b64))
                if "exp" in payload:
                    expires_at = float(payload["exp"])
            except Exception:
                pass

        if not token:
            raise ValueError(f"Could not parse access token from response: {resp.text[:200]}")

        self._bearer_token = token
        self._token_expires_at = expires_at
        return token

    def request(self, method: str, path: str, **kwargs) -> Any:
        """Make an authenticated request to the Garmin Personal API."""
        token = self.get_token()
        headers = kwargs.pop("headers", {})
        headers["Authorization"] = f"Bearer {token}"
        headers.setdefault("Accept", "application/json")
        headers.setdefault("User-Agent", "GarminSync/1.0")

        url = f"{self.base_url}{path}" if path.startswith("/") else f"{self.base_url}/{path}"
        resp = self.session.request(method, url, headers=headers, **kwargs)
        if resp.status_code == 401:
            # Token might have expired or been revoked; attempt one refresh
            token = self.get_token(force_refresh=True)
            headers["Authorization"] = f"Bearer {token}"
            resp = self.session.request(method, url, headers=headers, **kwargs)

        if resp.status_code == 204:
            return None

        resp.raise_for_status()
        return resp.json()

    # -------------------------------------------------------------------------
    # API endpoints
    # -------------------------------------------------------------------------

    def get_activities(self, start_date: str, end_date: str, limit: int = 99) -> List[Dict[str, Any]]:
        """Fetch activities in date range [start_date, end_date] (YYYY-MM-DD)."""
        path = "/activitylist-service/activities/search/activities"
        params = {
            "startDate": start_date,
            "endDate": end_date,
            "start": 0,
            "limit": limit,
        }
        data = self.request("GET", path, params=params)
        return data if isinstance(data, list) else []

    def get_activity_details(self, activity_id: int | str) -> Dict[str, Any]:
        """Fetch full details (including summaryDTO with feel/RPE) for an activity."""
        path = f"/activity-service/activity/{activity_id}"
        return self.request("GET", path)

    def get_sleep_data_range(self, start_date: str, end_date: str) -> List[Dict[str, Any]]:
        """Fetch daily sleep summaries for date range [start_date, end_date]."""
        path = "/sleep-service/sleep/dailySleepSummariesByDate"
        params = {"startDate": start_date, "endDate": end_date}
        data = self.request("GET", path, params=params)
        return data if isinstance(data, list) else []

    def get_hrv_range(self, from_date: str, to_date: str) -> Dict[str, Any]:
        """Fetch daily HRV summaries between from_date and to_date."""
        path = f"/hrv-service/hrv/daily/{from_date}/{to_date}"
        return self.request("GET", path) or {}

    def get_daily_heart_rate(self, date_str: str) -> Dict[str, Any]:
        """Fetch daily heart rate (contains restingHeartRate)."""
        path = "/wellness-service/wellness/dailyHeartRate"
        return self.request("GET", path, params={"date": date_str}) or {}

    def get_daily_summary(self, calendar_date: str) -> Dict[str, Any]:
        """Fetch user wellness daily summary (contains restingHeartRate, etc.)."""
        path = "/usersummary-service/usersummary/daily"
        return self.request("GET", path, params={"calendarDate": calendar_date}) or {}


# -------------------------------------------------------------------------
# Adapter Functions Matching main.py Data Contracts
# -------------------------------------------------------------------------

def fetch_sleep_data_personal(client: GarminPersonalApiClient, today: datetime.date, days: int = 7) -> Dict[str, Any]:
    """Fetch sleep, resting HR, and HRV for the past `days` days using Personal API."""
    start_date = (today - datetime.timedelta(days=days - 1)).isoformat()
    end_date = today.isoformat()
    print(f"\n[Personal API] Fetching sleep & wellness data for past {days} days ({start_date} to {end_date})...")

    # 1. Fetch sleep summaries in bulk
    sleep_list = []
    try:
        sleep_list = client.get_sleep_data_range(start_date, end_date)
    except Exception as e:
        print(f"  ⚠️ Failed to fetch sleep range: {e}")

    sleep_by_date: Dict[str, Any] = {}
    for item in sleep_list:
        d = item.get("calendarDate")
        if d:
            sleep_by_date[d] = item

    # 2. Fetch HRV range in bulk
    hrv_map: Dict[str, Any] = {}
    try:
        hrv_resp = client.get_hrv_range(start_date, end_date)
        for hrv in hrv_resp.get("hrvSummaries", []):
            cd = hrv.get("calendarDate")
            if cd:
                # 'lastNightAvg' is the overnight average HRV in the Personal API schema
                hrv_map[cd] = hrv.get("lastNightAvg")
    except Exception as e:
        print(f"  ⚠️ Failed to fetch HRV range: {e}")

    all_sleep_data: Dict[str, Any] = {}
    for i in range(days):
        target_date = today - datetime.timedelta(days=i)
        date_str = target_date.isoformat()

        # Sleep duration
        daily_sleep = sleep_by_date.get(date_str, {})
        sleep_secs = daily_sleep.get("sleepTimeSeconds")
        sleep_hours = round(sleep_secs / 3600.0, 2) if sleep_secs is not None else None

        # Resting heart rate
        resting_hr = None
        try:
            hr_data = client.get_daily_heart_rate(date_str)
            resting_hr = hr_data.get("restingHeartRate")
        except Exception:
            pass

        if resting_hr is None:
            # Fallback to daily summary
            try:
                summary = client.get_daily_summary(date_str)
                resting_hr = summary.get("restingHeartRate")
            except Exception:
                pass

        avg_hrv = hrv_map.get(date_str)

        all_sleep_data[date_str] = {
            "calendarDate": date_str,
            "restingHeartRate": resting_hr,
            "avgOvernightHrv": avg_hrv,
            "sleepTimeHours": sleep_hours,
        }

    return all_sleep_data


def fetch_training_data_personal(
    client: GarminPersonalApiClient,
    start_date: datetime.date,
    end_date: datetime.date,
) -> List[Dict[str, Any]]:
    """Fetch activities within date range using Personal API."""
    print(f"\n[Personal API] Fetching activities from {start_date.isoformat()} to {end_date.isoformat()}...")
    raw_activities = []
    try:
        raw_activities = client.get_activities(start_date.isoformat(), end_date.isoformat())
    except Exception as e:
        print(f"  ❌ Failed to fetch activities: {e}")
        return []

    training_data: List[Dict[str, Any]] = []
    for act in raw_activities:
        act_id = act.get("activityId")
        duration_mins = round((act.get("duration") or 0) / 60.0, 2)
        distance_km = round((act.get("distance") or 0) / 1000.0, 2)
        train_load = act.get("activityTrainingLoad")

        # activityType in SummarizedActivityDTO can be an object with typeKey or direct
        act_type_obj = act.get("activityType")
        if isinstance(act_type_obj, dict):
            act_type = act_type_obj.get("typeKey")
        else:
            act_type = str(act_type_obj) if act_type_obj else "other"

        # Direct feel and RPE might be on the activity or detail summary
        feel = act.get("directWorkoutFeel")
        rpe = act.get("directWorkoutRpe")

        if (feel is None or rpe is None) and act_id:
            try:
                details = client.get_activity_details(act_id)
                summary = details.get("summaryDTO", {})
                if feel is None:
                    feel = summary.get("directWorkoutFeel")
                if rpe is None:
                    rpe = summary.get("directWorkoutRpe")
            except Exception:
                pass

        # Extract time-in-zones directly present in SummarizedActivityDTO
        hr_zones = []
        for z in range(1, 6):
            secs = act.get(f"hrTimeInZone_{z}")
            if secs is not None:
                hr_zones.append({"zoneNumber": z, "secsInZone": float(secs)})

        act_info = {
            "activityId": act_id,
            "activityName": act.get("activityName"),
            "activityType": act_type,
            "startTime": act.get("startTimeLocal"),
            "durationMinutes": duration_mins,
            "distanceKm": distance_km,
            "calories": act.get("calories"),
            "averageHeartRate": act.get("averageHR"),
            "trainingLoad": round(train_load, 2) if train_load is not None else None,
            "trainingEffect": act.get("trainingEffectLabel"),
            "feel": feel,
            "rpe": rpe,
            "hrZones": hr_zones if hr_zones else None,
        }
        training_data.append(act_info)

    return training_data
