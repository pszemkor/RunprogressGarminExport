"""Unit tests for personal_api.py client and sync adapters."""
import datetime
from unittest.mock import Mock, patch
import pytest

from personal_api import (
    GarminPersonalApiClient,
    fetch_sleep_data_personal,
    fetch_training_data_personal,
)


def test_client_get_token():
    client = GarminPersonalApiClient(api_key="test-key-123")
    mock_resp = Mock()
    mock_resp.status_code = 200
    mock_resp.json.return_value = {"access_token": "bearer-abc", "expires_in": 3600}

    with patch.object(client.session, "post", return_value=mock_resp) as mock_post:
        token = client.get_token()
        assert token == "bearer-abc"
        mock_post.assert_called_once()
        headers = mock_post.call_args[1]["headers"]
        assert headers["Authorization"] == "ApiKey test-key-123"

        # Subsequent call before expiry reuses cached token
        token2 = client.get_token()
        assert token2 == "bearer-abc"
        assert mock_post.call_count == 1


def test_client_get_token_raw_jwt():
    import base64
    import json
    import time

    payload = base64.urlsafe_b64encode(json.dumps({"exp": time.time() + 7200}).encode()).decode().rstrip("=")
    fake_jwt = f"header.{payload}.sig"

    client = GarminPersonalApiClient(api_key="test-key-456")
    mock_resp = Mock()
    mock_resp.status_code = 200
    mock_resp.text = fake_jwt
    # When text/plain is returned, resp.json() raises JSONDecodeError
    mock_resp.json.side_effect = ValueError("Invalid JSON")

    with patch.object(client.session, "post", return_value=mock_resp):
        token = client.get_token()
        assert token == fake_jwt



def test_client_request_adds_bearer():
    client = GarminPersonalApiClient(api_key="test-key", bearer_token="bearer-token-xyz", token_expires_at=9999999999)
    mock_resp = Mock()
    mock_resp.status_code = 200
    mock_resp.json.return_value = [{"activityId": 1}]

    with patch.object(client.session, "request", return_value=mock_resp) as mock_req:
        res = client.request("GET", "/test/path")
        assert res == [{"activityId": 1}]
        headers = mock_req.call_args[1]["headers"]
        assert headers["Authorization"] == "Bearer bearer-token-xyz"


def test_fetch_sleep_data_personal():
    client = Mock(spec=GarminPersonalApiClient)
    client.get_sleep_data_range.return_value = [
        {"calendarDate": "2026-11-01", "sleepTimeSeconds": 28800},
        {"calendarDate": "2026-11-02", "sleepTimeSeconds": 25200},
    ]
    client.get_hrv_range.return_value = {
        "hrvSummaries": [
            {"calendarDate": "2026-11-01", "lastNightAvg": 88},
            {"calendarDate": "2026-11-02", "lastNightAvg": 92},
        ]
    }
    client.get_daily_heart_rate.side_effect = lambda d: {"restingHeartRate": 50 if d == "2026-11-01" else 48}

    today = datetime.date(2026, 11, 2)
    data = fetch_sleep_data_personal(client, today, days=2)

    assert "2026-11-01" in data
    assert "2026-11-02" in data
    assert data["2026-11-01"] == {
        "calendarDate": "2026-11-01",
        "restingHeartRate": 50,
        "avgOvernightHrv": 88,
        "sleepTimeHours": 8.0,
    }
    assert data["2026-11-02"]["sleepTimeHours"] == 7.0
    assert data["2026-11-02"]["avgOvernightHrv"] == 92
    assert data["2026-11-02"]["restingHeartRate"] == 48


def test_fetch_training_data_personal():
    client = Mock(spec=GarminPersonalApiClient)
    client.get_activities.return_value = [
        {
            "activityId": 999,
            "activityName": "Cracow Interval Run",
            "activityType": {"typeKey": "running"},
            "startTimeLocal": "2026-11-02 10:00:00",
            "duration": 3600.0,
            "distance": 10000.0,
            "calories": 650.0,
            "averageHR": 155.0,
            "activityTrainingLoad": 140.2,
            "trainingEffectLabel": "VO2 Max",
            "directWorkoutFeel": 100,
            "directWorkoutRpe": 40,
            "hrTimeInZone_1": 300.0,
            "hrTimeInZone_2": 900.0,
            "hrTimeInZone_3": 1200.0,
            "hrTimeInZone_4": 900.0,
            "hrTimeInZone_5": 300.0,
        }
    ]

    start = datetime.date(2026, 11, 1)
    end = datetime.date(2026, 11, 2)
    acts = fetch_training_data_personal(client, start, end)

    assert len(acts) == 1
    act = acts[0]
    assert act["activityId"] == 999
    assert act["activityName"] == "Cracow Interval Run"
    assert act["activityType"] == "running"
    assert act["durationMinutes"] == 60.0
    assert act["distanceKm"] == 10.0
    assert act["feel"] == 100
    assert act["rpe"] == 40
    assert len(act["hrZones"]) == 5
    assert act["hrZones"][3] == {"zoneNumber": 4, "secsInZone": 900.0}


def test_fetch_training_data_personal_fetches_rpe_and_feel_from_details():
    client = Mock(spec=GarminPersonalApiClient)
    client.get_activities.return_value = [
        {
            "activityId": 1001,
            "activityName": "Treadmill Run",
            "activityType": {"typeKey": "treadmill_running"},
            "startTimeLocal": "2026-11-08 10:00:00",
            "duration": 1800.0,
            "distance": 5000.0,
        },
        {
            "activityId": 1002,
            "activityName": "Virtual Ride",
            "activityType": {"typeKey": "virtual_ride"},
            "startTimeLocal": "2026-11-08 16:00:00",
            "duration": 2400.0,
            "distance": 20000.0,
        },
    ]
    client.get_activity_details.side_effect = [
        {"summaryDTO": {"directWorkoutFeel": 75, "directWorkoutRpe": 20}},
        RuntimeError("transient error"),
    ]

    acts = fetch_training_data_personal(
        client, datetime.date(2026, 11, 8), datetime.date(2026, 11, 8)
    )

    assert len(acts) == 2
    assert acts[0]["activityType"] == "treadmill_running"
    assert acts[0]["feel"] == 75
    assert acts[0]["rpe"] == 20
    assert acts[1]["activityType"] == "virtual_ride"
    assert acts[1]["feel"] is None
    assert acts[1]["rpe"] is None
