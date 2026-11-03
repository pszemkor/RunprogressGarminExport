"""Unit tests for Cloud Functions entrypoint."""
from unittest.mock import Mock, patch
import json
import pytest

import function_entry


@patch("function_entry.run_sync")
def test_sync_garmin_http_get(mock_run_sync):
    mock_run_sync.return_value = {"status": "success", "synced_days": 5}
    req = Mock()
    req.args = {"days": "5"}
    req.is_json = False

    resp_body, code, headers = function_entry.sync_garmin_http(req)
    assert code == 200
    assert headers["Content-Type"] == "application/json"
    data = json.loads(resp_body)
    assert data["status"] == "success"
    mock_run_sync.assert_called_once_with(days=5, use_personal_api=True)


@patch("function_entry.run_sync")
def test_sync_garmin_http_error_handling(mock_run_sync):
    mock_run_sync.side_effect = RuntimeError("Something went wrong")
    req = Mock()
    req.args = {}
    req.is_json = False

    resp_body, code, headers = function_entry.sync_garmin_http(req)
    assert code == 500
    data = json.loads(resp_body)
    assert data["status"] == "error"
    assert "Something went wrong" in data["message"]
