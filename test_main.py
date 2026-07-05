import datetime
import pytest
from unittest.mock import Mock, patch

import main

from main import (
    hours_to_time_str,
    format_time,
    device_data_phrase_format,
    normalize_training_status,
    fetch_sleep_data,
    fetch_training_data,
    fetch_training_status,
    fetch_race_predictions,
)

def test_hours_to_time_str():
    assert hours_to_time_str(5.77) == "5:46"
    assert hours_to_time_str(0) == ""
    assert hours_to_time_str(None) == ""
    assert hours_to_time_str(1.0) == "1:00"
    assert hours_to_time_str(1.5) == "1:30"

def test_format_time():
    assert format_time(None) is None
    assert format_time(0) is None
    assert format_time(60) == "00:01:00"
    assert format_time(3600) == "01:00:00"
    assert format_time(3661) == "01:01:01"
    assert format_time(45296) == "12:34:56"

def test_device_data_phrase_format():
    assert device_data_phrase_format("productive") == "Productive"
    assert device_data_phrase_format("highly_productive") == "Highly Productive"
    assert device_data_phrase_format(None) == "Unknown"
    assert device_data_phrase_format("") == "Unknown"

def test_normalize_training_status():
    assert normalize_training_status("") == ""
    assert normalize_training_status(None) == ""
    assert normalize_training_status("Strained 3") == "Strained"
    assert normalize_training_status("Peaking") == "Peaking"
    assert normalize_training_status("Recovery") == "Recovery"
    assert normalize_training_status("unknown_status") == ""
    assert normalize_training_status("Highly Productive") == "Productive"

def test_fetch_sleep_data():
    api = Mock()
    api.get_sleep_data.return_value = {
        "dailySleepDTO": {
            "calendarDate": "2026-07-05",
            "sleepTimeSeconds": 28800  # 8 hours
        },
        "restingHeartRate": 55,
        "avgOvernightHrv": 65
    }
    
    today = datetime.date(2026, 7, 5)
    data = fetch_sleep_data(api, today, days=1)
    
    assert "2026-07-05" in data
    assert data["2026-07-05"] == {
        "calendarDate": "2026-07-05",
        "restingHeartRate": 55,
        "avgOvernightHrv": 65,
        "sleepTimeHours": 8.0
    }
    api.get_sleep_data.assert_called_once_with("2026-07-05")

def test_fetch_sleep_data_empty():
    api = Mock()
    api.get_sleep_data.return_value = None
    
    today = datetime.date(2026, 7, 5)
    data = fetch_sleep_data(api, today, days=1)
    
    assert "2026-07-05" in data
    assert data["2026-07-05"]["error"] == "No data returned"

def test_fetch_training_data():
    api = Mock()
    api.get_activities_by_date.return_value = [
        {
            "activityId": 12345,
            "activityName": "Morning Run",
            "activityType": {"typeKey": "running"},
            "startTimeLocal": "2026-07-05 08:00:00",
            "duration": 3600,
            "distance": 10000,
            "calories": 600,
            "averageHR": 140,
            "activityTrainingLoad": 120.5,
            "trainingEffectLabel": "Aerobic"
        }
    ]
    
    start = datetime.date(2026, 7, 1)
    end = datetime.date(2026, 7, 5)
    data = fetch_training_data(api, start, end)
    
    assert len(data) == 1
    assert data[0]["activityId"] == 12345
    assert data[0]["activityName"] == "Morning Run"
    assert data[0]["activityType"] == "running"
    assert data[0]["startTime"] == "2026-07-05 08:00:00"
    assert data[0]["durationMinutes"] == 60.0
    assert data[0]["distanceKm"] == 10.0
    assert data[0]["calories"] == 600
    assert data[0]["averageHeartRate"] == 140
    assert data[0]["trainingLoad"] == 120.5
    assert data[0]["trainingEffect"] == "Aerobic"

def test_fetch_training_status():
    api = Mock()
    api.get_training_status.return_value = {
        "mostRecentTrainingStatus": {
            "latestTrainingStatusData": {
                "some_device_id": {
                    "acuteTrainingLoadDTO": {
                        "dailyTrainingLoadAcute": 500,
                        "dailyTrainingLoadChronic": 450
                    },
                    "trainingStatusFeedbackPhrase": "productive"
                }
            }
        }
    }
    
    today = datetime.date(2026, 7, 5)
    data = fetch_training_status(api, today, days=1)
    
    assert "2026-07-05" in data
    assert data["2026-07-05"] == {
        "status": "Productive",
        "acute_load": 500,
        "chronic_load": 450
    }

def test_fetch_race_predictions():
    api = Mock()
    api.get_race_predictions.return_value = [
        {
            "calendarDate": "2026-07-05",
            "time5K": 1200,          # 20 mins
            "time10K": 2500,         # 41:40
            "timeHalfMarathon": 5400, # 1:30:00
            "timeMarathon": 11400     # 3:10:00
        }
    ]
    
    start = datetime.date(2026, 7, 1)
    end = datetime.date(2026, 7, 5)
    data = fetch_race_predictions(api, start, end)
    
    assert "2026-07-05" in data
    assert data["2026-07-05"] == {
        "5K": "00:20:00",
        "10K": "00:41:40",
        "HalfMarathon": "01:30:00",
        "Marathon": "03:10:00"
    }

@patch("main.fetch_and_print_hr_zones")
@patch("main.fetch_race_predictions")
@patch("main.fetch_training_status")
@patch("main.fetch_training_data")
@patch("main.fetch_sleep_data")
@patch("main.init_api")
@patch("main.os.getenv")
@patch("main.sys.argv", ["main.py"])
def test_main_weekly_hr_zones(
    mock_getenv,
    mock_init_api,
    mock_fetch_sleep,
    mock_fetch_training,
    mock_fetch_status,
    mock_fetch_race,
    mock_fetch_hr,
):
    # Setup mocks
    mock_getenv.return_value = "dummy"
    api_mock = Mock()
    mock_init_api.return_value = api_mock
    
    mock_training_data = []
    mock_fetch_training.return_value = mock_training_data
    
    class MockDate(datetime.date):
        @classmethod
        def today(cls):
            return datetime.date(2026, 7, 8) # Wednesday
            
    with patch("main.datetime.date", MockDate):
        main.main()
        
    # We expect 2 calls to fetch_and_print_hr_zones
    assert mock_fetch_hr.call_count == 2
    
    # Check args for previous week
    call1_args = mock_fetch_hr.call_args_list[0][0]
    assert call1_args[0] == api_mock
    assert call1_args[1] == mock_training_data
    assert call1_args[2] == datetime.date(2026, 6, 29) # Monday prev week
    assert call1_args[3] == datetime.date(2026, 7, 5)  # Sunday prev week
    
    # Check args for current week
    call2_args = mock_fetch_hr.call_args_list[1][0]
    assert call2_args[0] == api_mock
    assert call2_args[1] == mock_training_data
    assert call2_args[2] == datetime.date(2026, 7, 6) # Monday current week
    assert call2_args[3] == datetime.date(2026, 7, 8) # Today

