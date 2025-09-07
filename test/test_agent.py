# tests/test_agent.py
import json
from unittest.mock import Mock
import azure.functions as func
import function_app as fa

def make_req(payload: dict) -> func.HttpRequest:
    return func.HttpRequest(
        method="POST",
        body=json.dumps(payload).encode("utf-8"),
        url="/api/agent_httptrigger",
        params={},
    )

def test_agent_success(monkeypatch):
    # Prepare mocks for container/blob
    blob_client = Mock()
    blob_client.account_name = "acct"
    blob_client.upload_blob = Mock()

    container_mock = Mock()
    container_mock.get_blob_client.return_value = blob_client

    # Prepare mock table client
    table_mock = Mock()
    table_mock.create_entity = Mock()

    # Monkeypatch accessors in function_app
    monkeypatch.setattr(fa, "_get_container_client", lambda: container_mock)
    monkeypatch.setattr(fa, "_get_table_client", lambda: table_mock)

    payload = {
        "CompanyReseachData1": {"k": 1},
        "CompanyReseachData2": {"k": 2},
        "CompanyReseachData3": {"k": 3},
        "IndustryResearch": {"k": 4},
    }

    req = make_req(payload)
    resp = fa.agent_httptrigger(req)
    assert resp.status_code == 200
    body = json.loads(resp.get_body().decode())
    assert body["status"] == "success"
    table_mock.create_entity.assert_called_once()
    blob_client.upload_blob.assert_called_once()

def test_missing_required_returns_400():
    req = make_req({})
    resp = fa.agent_httptrigger(req)
    assert resp.status_code == 400

def test_invalid_json_returns_400():
    # craft invalid body
    req = func.HttpRequest(method="POST", body=b"not-json", url="/api/agent_httptrigger")
    resp = fa.agent_httptrigger(req)
    assert resp.status_code == 400