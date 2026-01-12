# tests/test_agent.py
import json
import os
import sys
from unittest.mock import Mock

import azure.functions as func

# Ensure project root is on sys.path so tests can import top-level modules
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

import function_app as fa
from helpers.valuemap_transformer import is_valuemap_format, transform_valuemap_to_industry_research


def make_req(payload: dict) -> func.HttpRequest:
    return func.HttpRequest(
        method="POST",
        body=json.dumps(payload).encode("utf-8"),
        url="/api/agent_httptrigger",
        params={},
    )


def _mock_blob_and_container(monkeypatch):
    """Helper to mock blob storage clients."""
    blob_client = Mock()
    blob_client.account_name = "acct"
    blob_client.upload_blob = Mock()

    container_mock = Mock()
    container_mock.get_blob_client.return_value = blob_client

    monkeypatch.setattr(fa, "_get_container_client", lambda: container_mock)
    return blob_client, container_mock


def test_agent_success(monkeypatch):
    """Test successful processing with IndustryResearch format."""
    blob_client, _ = _mock_blob_and_container(monkeypatch)

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
        "CompanyResearchData1": pd1,
        "CompanyResearchData2": pd2,
        "CompanyResearchData3": pd3,
        "IndustryResearch": pdi,
    }

    req = make_req(payload)
    resp = fa.agent_httptrigger(req)
    assert resp.status_code == 200
    body = json.loads(resp.get_body().decode())
    assert body["status"] == "success"
    blob_client.upload_blob.assert_called_once()


def test_agent_success_with_valuemap(monkeypatch):
    """Test successful processing with ValueMap format (auto-transformed)."""
    blob_client, _ = _mock_blob_and_container(monkeypatch)

    # load combined L'Oreal payload that uses ValueMap format
    with open("test/payloads/combined_loreal.json") as f:
        payload = json.load(f)

    req = make_req(payload)
    resp = fa.agent_httptrigger(req)
    assert resp.status_code == 200
    body = json.loads(resp.get_body().decode())
    assert body["status"] == "success"
    blob_client.upload_blob.assert_called_once()


def test_missing_required_returns_400():
    """Test that missing required fields returns 400."""
    req = make_req({})
    resp = fa.agent_httptrigger(req)
    assert resp.status_code == 400


def test_missing_industry_and_valuemap_returns_400():
    """Test that missing both IndustryResearch and ValueMap returns 400."""
    payload = {
        "CompanyResearchData1": {"Company Name": "Test"},
        "CompanyResearchData2": {"Revenue": {"Amount": 1000}},
        "CompanyResearchData3": {"Latest Relevant News": []},
    }
    req = make_req(payload)
    resp = fa.agent_httptrigger(req)
    assert resp.status_code == 400
    body = json.loads(resp.get_body().decode())
    assert "IndustryResearch or ValueMap" in body["error"]


def test_invalid_json_returns_400():
    """Test that invalid JSON returns 400."""
    req = func.HttpRequest(method="POST", body=b"not-json", url="/api/agent_httptrigger")
    resp = fa.agent_httptrigger(req)
    assert resp.status_code == 400


# --- ValueMap Transformer Unit Tests ---


def test_is_valuemap_format_with_benefit_table():
    """Test detection of ValueMap format."""
    valuemap = {"BenefitTable": [{"Challenge": "Test", "KPI": "Test KPI"}]}
    assert is_valuemap_format(valuemap) is True


def test_is_valuemap_format_without_benefit_table():
    """Test detection of non-ValueMap format."""
    industry = {"title": "Test", "headers": ["A"], "rows": []}
    assert is_valuemap_format(industry) is False


def test_is_valuemap_format_with_non_dict():
    """Test detection returns False for invalid types."""
    assert is_valuemap_format("not a dict") is False
    assert is_valuemap_format(None) is False
    assert is_valuemap_format([]) is False  # Empty list


def test_is_valuemap_format_with_direct_array():
    """Test detection returns True for direct array of objects."""
    direct_array = [{"Challenge": "Test", "Description": "Desc"}]
    assert is_valuemap_format(direct_array) is True


def test_transform_valuemap_basic():
    """Test basic ValueMap transformation."""
    valuemap = {
        "BenefitTable": [
            {
                "Challenge": "Challenge 1",
                "Description": "Desc 1",
                "KPI": "KPI 1",
                "Workload": "Workload 1",
                "CalculatedBenefit": 1000,
            },
            {
                "Challenge": "Challenge 2",
                "Description": "Desc 2",
                "KPI": "KPI 2",
                "Workload": "Workload 2",
                "CalculatedBenefit": 2000,
            },
        ]
    }

    result = transform_valuemap_to_industry_research(valuemap, "Test Company")

    assert result["title"] == "Test Company Value Map Benefit Table"
    assert "Challenge" in result["headers"]
    assert "Description" in result["headers"]
    assert "KPI" in result["headers"]
    assert len(result["rows"]) == 2
    assert result["rows"][0]["Challenge"] == "Challenge 1"
    # Values are flattened to strings
    assert result["rows"][1]["CalculatedBenefit"] == "2000"


def test_transform_valuemap_includes_inputs():
    """Test that Inputs column is included and flattened."""
    valuemap = {
        "BenefitTable": [
            {
                "Challenge": "Test",
                "KPI": "Test KPI",
                "Inputs": {"Nested": "Value", "Another": 123},
                "CalculatedBenefit": 500,
            }
        ]
    }

    result = transform_valuemap_to_industry_research(valuemap)

    # Inputs is now included in headers
    assert "Inputs" in result["headers"]
    assert "Challenge" in result["headers"]
    assert "KPI" in result["headers"]
    # Inputs value is flattened to string
    assert result["rows"][0]["Inputs"] == "Nested: Value; Another: 123"


def test_transform_valuemap_empty_company_name():
    """Test title generation without company name."""
    valuemap = {"BenefitTable": [{"Challenge": "Test"}]}

    result = transform_valuemap_to_industry_research(valuemap, "")

    assert result["title"] == "Value Map Benefit Table"


def test_transform_valuemap_empty_benefit_table_returns_minimal():
    """Test that empty BenefitTable returns minimal valid structure."""
    valuemap = {"BenefitTable": []}

    result = transform_valuemap_to_industry_research(valuemap)
    assert result["title"] == "Value Map Benefit Table"
    assert result["headers"] == []
    assert result["rows"] == []


def test_transform_valuemap_missing_benefit_table_returns_minimal():
    """Test that missing BenefitTable returns minimal valid structure."""
    valuemap = {"SomeOtherKey": []}

    result = transform_valuemap_to_industry_research(valuemap)
    assert result["title"] == "Value Map Benefit Table"
    assert result["headers"] == []
    assert result["rows"] == []


def test_transform_valuemap_direct_array():
    """Test that direct array format works without BenefitTable wrapper."""
    valuemap = [
        {"Challenge": "Test Challenge", "Description": "Test Description"},
        {"Challenge": "Another Challenge", "Description": "Another Description"},
    ]

    result = transform_valuemap_to_industry_research(valuemap)
    assert result["title"] == "Value Map Benefit Table"
    assert "Challenge" in result["headers"]
    assert "Description" in result["headers"]
    assert len(result["rows"]) == 2
    assert result["rows"][0]["Challenge"] == "Test Challenge"
