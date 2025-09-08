# tests/test_agent.py
import json
import os
import sys
from unittest.mock import Mock

import azure.functions as func

# Ensure project root is on sys.path so tests can import top-level modules
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

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

    # load payload files from test/payloads
    with (
        open("test/payloads/CompanyReseachData1_v2.json") as f1,
        open("test/payloads/CompanyReseachData2_v2.json") as f2,
        open("test/payloads/CompanyReseachData3_v2.json") as f3,
        open("test/payloads/IndustryResearch.json") as fi,
    ):
        pd1 = json.load(f1)
        pd2 = json.load(f2)
        pd3 = json.load(f3)
        pdi = json.load(fi)

    payload = {
        "CompanyReseachData1": pd1,
        "CompanyReseachData2": pd2,
        "CompanyReseachData3": pd3,
        "IndustryResearch": pdi,
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
