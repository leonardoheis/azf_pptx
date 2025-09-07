import logging
import json
from io import BytesIO
from datetime import datetime, timedelta, timezone
import os
from functools import lru_cache

import azure.functions as func
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE
from pptx.dml.color import RGBColor

from azure.data.tables import TableServiceClient
from azure.storage.blob import BlobServiceClient
from azure.core.exceptions import ResourceExistsError
from azure.ai.projects import AIProjectClient
from azure.identity import DefaultAzureCredential
from azure.data.tables import TableServiceClient
from azure.storage.blob import BlobServiceClient


from config import (
    AZ_STORAGE_CONN_STRING,            # may be None in local env
    AZ_BLOB_TABLE_NAME,
    INPUT_TEMPLATE,
    get_next_output_filename,
    AZ_BLOB_CONTAINER_NAME,
)

from company_research1 import fill_company_research1, fill_company_name_from_json
from company_research2 import fill_company_research2
from company_research3 import fill_company_research3
from industry_research import fill_industry_slides

app = func.FunctionApp(http_auth_level=func.AuthLevel.ANONYMOUS)


@lru_cache(maxsize=1)
def _init_clients():
    """Create and return (table_client, blob_service_client, container_client).

    Cached as a single tuple to avoid module-level mutable globals.
    """
    conn = _get_conn_string()

    # Table client
    svc = TableServiceClient.from_connection_string(conn)
    table_client = svc.get_table_client(table_name=AZ_BLOB_TABLE_NAME)
    try:
        table_client.create_table()
    except ResourceExistsError:
        logging.debug("Table already exists (create_table). Continuing.")
    except Exception:
        logging.exception("Unexpected error creating table")
        raise

    # Blob client
    blob_service_client = BlobServiceClient.from_connection_string(conn)
    container_client = blob_service_client.get_container_client(container=AZ_BLOB_CONTAINER_NAME)
    try:
        container_client.create_container()
    except ResourceExistsError:
        logging.debug("Container already exists (create_container). Continuing.")
    except Exception:
        logging.exception("Unexpected error creating container")
        raise

    return table_client, blob_service_client, container_client

def _get_conn_string() -> str:
    """
    Prefer AZ_STORAGE_CONN_STRING; fallback to AzureWebJobsStorage.
    Validate it early but *don't* fail at import-time.
    """
    conn = AZ_STORAGE_CONN_STRING or os.getenv("AZ_STORAGE_CONN_STRING")
    if not conn:
        conn = os.getenv("AzureWebJobsStorage")  # may be 'UseDevelopmentStorage=true'
    if not conn or len(conn.strip()) == 0:
        raise RuntimeError("Storage connection string missing. "
                           "Set AZ_STORAGE_CONN_STRING or AzureWebJobsStorage.")
    return conn

def _ensure_resources():
    """Compatibility shim retained for call sites; triggers client initialization."""
    _init_clients()

def _get_table_client():
    table_client, _, _ = _init_clients()
    return table_client


def _get_container_client():
    _, _, container_client = _init_clients()
    return container_client


@app.route(route="agent_httptrigger")
def agent_httptrigger(req: func.HttpRequest) -> func.HttpResponse:
    """
    Receives 4 JSON files and produces a PPTX, uploads to Blob, and logs to Tables.
    """
    try:
        req_body = req.get_json()

        required = ["CompanyReseachData1", "CompanyReseachData2", "CompanyReseachData3", "IndustryResearch"]
        missing = [k for k in required if k not in req_body]
        if missing:
            return func.HttpResponse(
                json.dumps({"error": f"Missing required files: {', '.join(missing)}", "status": "error"}),
                status_code=400,
                mimetype="application/json",
            )

        company_data1 = req_body["CompanyReseachData1"]
        company_data2 = req_body["CompanyReseachData2"]
        company_data3 = req_body["CompanyReseachData3"]
        industry_data = req_body["IndustryResearch"]

        for obj in (company_data1, company_data2, company_data3, industry_data):
            if not isinstance(obj, dict):
                return func.HttpResponse(
                    json.dumps({"error": "All files must contain valid JSON objects", "status": "error"}),
                    status_code=400,
                    mimetype="application/json",
                )

        logging.info("Keys1=%s", list(company_data1.keys()))
        logging.info("Keys2=%s", list(company_data2.keys()))
        logging.info("Keys3=%s", list(company_data3.keys()))
        logging.info("KeysIndustry=%s", list(industry_data.keys()))

        processed_data = {
            "company_data1_processed": len(company_data1),
            "company_data2_processed": len(company_data2),
            "company_data3_processed": len(company_data3),
            "industry_data_processed": len(industry_data),
            "total_fields": len(company_data1) + len(company_data2) + len(company_data3) + len(industry_data),
            "timestamp": datetime.now(timezone.utc).isoformat(),
            "output_location": None,
        }

        # Build PPTX + upload
        current_output_file = get_next_output_filename()
        try:
            prs = Presentation(INPUT_TEMPLATE)

            # your fill_* helpers (imported at module scope)

            fill_company_name_from_json(prs, company_data1)
            fill_company_research1(prs, company_data1)
            fill_company_research2(prs, company_data2)
            fill_company_research3(prs, company_data3)
            fill_industry_slides(prs, prs.slides[len(prs.slides) - 1], industry_data)

            buf = BytesIO()
            prs.save(buf)
            buf.seek(0)

            container_client = _get_container_client()
            blob_client = container_client.get_blob_client(blob=current_output_file)
            blob_client.upload_blob(buf.getvalue(), overwrite=True)

            account = blob_client.account_name  # or _blob_service_client.account_name
            processed_data["output_location"] = f"https://{account}.blob.core.windows.net/{AZ_BLOB_CONTAINER_NAME}/{current_output_file}"
            logging.info("Presentation saved: %s", processed_data["output_location"])
        except Exception as e:
            logging.exception("Error building/uploading PPTX: %s", e)

        # Log to Tables
        try:
            table_client = _get_table_client()
            entity = {
                "PartitionKey": "processed_files",
                "RowKey": datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S"),
                "CompanyData1Keys": json.dumps(list(company_data1.keys())),
                "CompanyData2Keys": json.dumps(list(company_data2.keys())),
                "CompanyData3Keys": json.dumps(list(company_data3.keys())),
                "IndustryDataKeys": json.dumps(list(industry_data.keys())),
                "ProcessedData": json.dumps(processed_data),
                "Timestamp": datetime.now(timezone.utc).isoformat(),
            }
            table_client.create_entity(entity=entity)
        except Exception as e:
            logging.exception("Error storing data in table: %s", e)

        return func.HttpResponse(
            json.dumps({
                "status": "success",
                "message": "Processed files and saved PPTX to Blob",
                "data": processed_data,
                "files_received": {
                    "CompanyReseachData1_size": len(company_data1),
                    "CompanyReseachData2_size": len(company_data2),
                    "CompanyReseachData3_size": len(company_data3),
                    "IndustryResearch_size": len(industry_data),
                },
                "output_file": {
                    "filename": current_output_file,
                    "blob_path": current_output_file,
                    "container": AZ_BLOB_CONTAINER_NAME,
                    "full_url": processed_data.get("output_location", "Not available"),
                },
            }),
            status_code=200,
            mimetype="application/json",
        )

    except ValueError as e:
        return func.HttpResponse(json.dumps({"error": f"Invalid JSON: {e}", "status": "error"}),
                                 status_code=400, mimetype="application/json")
    except Exception as e:
        logging.exception("Unexpected error")
        return func.HttpResponse(json.dumps({"error": f"Internal server error: {e}", "status": "error"}),
                                 status_code=500, mimetype="application/json")
