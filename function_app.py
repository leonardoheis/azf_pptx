
import logging
import json
from io import BytesIO

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE
from pptx.dml.color import RGBColor
from datetime import datetime

import azure.functions as func
from azure.ai.projects import AIProjectClient
from azure.identity import DefaultAzureCredential
from azure.data.tables import TableServiceClient
from azure.storage.blob import BlobServiceClient
from datetime import datetime, timedelta, timezone

from config import STORAGE_CONN_STRING, TABLE_NAME, TEMPLATE, get_next_output_filename, CONTAINER_NAME
from company_research1 import fill_company_research1, fill_company_name_from_json
from company_research2 import fill_company_research2
from company_research3 import fill_company_research3
from industry_research import fill_industry_slides

app = func.FunctionApp(http_auth_level=func.AuthLevel.ANONYMOUS)

table_service = TableServiceClient.from_connection_string(conn_str=STORAGE_CONN_STRING)
table_client = table_service.get_table_client(table_name=TABLE_NAME)

# Initialize Blob Service Client
blob_service_client = BlobServiceClient.from_connection_string(conn_str=STORAGE_CONN_STRING)
container_client = blob_service_client.get_container_client(container=CONTAINER_NAME)

# Asegura la tabla exista
try:
    table_client.create_table()
except Exception:
    pass  # Ya existe

# Ensure blob container exists
try:
    container_client.create_container()
except Exception:
    pass  # Already exists

def get_thread_last_active(thread_id):
    try:
        entity = table_client.get_entity(partition_key="thread", row_key=thread_id)
        last_active = datetime.fromisoformat(entity['LastActive'])
        return last_active
    except Exception:
        return None
    

@app.route(route="agent_httptrigger")
def agent_httptrigger(req: func.HttpRequest) -> func.HttpResponse:
    """
    Azure Function that receives 4 JSON files as parameters from Power Automate and processes them.
    
    Expected JSON structure in request body:
    {
        "CompanyReseachData1": { "content": "..." },
        "CompanyReseachData2": { "content": "..." },
        "CompanyReseachData3": { "content": "..." },
        "IndustryResearch": { "content": "..." }
    }
    """
    try:
        # Get the request body
        req_body = req.get_json()
        
        # Validate that all 3 required files are present
        required_files = ["CompanyReseachData1", "CompanyReseachData2", "CompanyReseachData3", "IndustryResearch"]
        missing_files = [file for file in required_files if file not in req_body]
        
        if missing_files:
            return func.HttpResponse(
                json.dumps({
                    "error": f"Missing required files: {', '.join(missing_files)}",
                    "status": "error"
                }),
                status_code=400,
                mimetype="application/json"
            )
        
        # Extract the 4 JSON files
        company_data1 = req_body["CompanyReseachData1"]
        company_data2 = req_body["CompanyReseachData2"]
        company_data3 = req_body["CompanyReseachData3"]
        industry_data = req_body["IndustryResearch"]
        
        # Validate that each file contains valid JSON data
        if not isinstance(company_data1, dict) or not isinstance(company_data2, dict) or not isinstance(company_data3, dict) or not isinstance(industry_data, dict):
            return func.HttpResponse(
                json.dumps({
                    "error": "All files must contain valid JSON objects",
                    "status": "error"
                }),
                status_code=400,
                mimetype="application/json"
            )
        
        # Log the received files for debugging
        logging.info(f"Received CompanyReseachData1 with keys: {list(company_data1.keys())}")
        logging.info(f"Received CompanyReseachData2 with keys: {list(company_data2.keys())}")
        logging.info(f"Received CompanyReseachData3 with keys: {list(company_data3.keys())}")
        logging.info(f"Received IndustryResearch with keys: {list(industry_data.keys())}")
        
        # Process the data using the specialized functions
        # Note: You can call these functions when you have a PowerPoint presentation to process
        # Example usage:
        # prs = Presentation("template.pptx")
        # fill_company_research1(prs, company_data1)
        # fill_company_research2(prs, company_data2)
        # fill_company_research3(prs, company_data3)
        # fill_industry_research(prs, industry_data)
        # prs.save("output.pptx")
        
        # Process the files (you can add your custom logic here)
        # For example, you might want to:
        # - Extract specific data from each file
        # - Combine data from multiple files
        # - Perform calculations or transformations
        # - Store results in Azure Table Storage
        
        # Example processing - modify this according to your needs
        processed_data = {
            "company_data1_processed": len(company_data1),
            "company_data2_processed": len(company_data2),
            "company_data3_processed": len(company_data3),
            "industry_data_processed": len(industry_data),
            "total_fields": len(company_data1) + len(company_data2) + len(company_data3) + len(industry_data),
            "timestamp": datetime.now(timezone.utc).isoformat(),
            "output_location": None  # Will be set after file is saved
        }
        
        try:
            # Generate a fresh output filename for this execution
            current_output_file = get_next_output_filename()
            
            prs = Presentation(TEMPLATE)
            
            fill_company_name_from_json(prs, company_data1)
            
            fill_company_research1(prs, company_data1)
            fill_company_research2(prs, company_data2)
            fill_company_research3(prs, company_data3)
            fill_industry_slides(prs, prs.slides[len(prs.slides) - 1], industry_data)
            
            # Save to blob storage (directly in container root)
            blob_client = container_client.get_blob_client(blob=current_output_file)
            
            # Save presentation to bytes buffer first
            pptx_buffer = BytesIO()
            prs.save(pptx_buffer)
            pptx_buffer.seek(0)
            
            # Upload to blob storage
            blob_client.upload_blob(pptx_buffer.getvalue(), overwrite=True)
            
            print(f"POC generada en Azure Blob Storage: {CONTAINER_NAME}/{current_output_file}")
            logging.info(f"Presentation saved to blob storage: {CONTAINER_NAME}/{current_output_file}")
            
            processed_data["output_location"] = f"https://{blob_service_client.account_name}.blob.core.windows.net/{CONTAINER_NAME}/{current_output_file}"
            
        except Exception as e:
            logging.error(f"Error processing data: {str(e)}")
        # Store the processed data in Azure Table Storage
        try:
            entity = {
                "PartitionKey": "processed_files",
                "RowKey": datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S"),
                "CompanyData1Keys": json.dumps(list(company_data1.keys())),
                "CompanyData2Keys": json.dumps(list(company_data2.keys())),
                "CompanyData3Keys": json.dumps(list(company_data3.keys())),
                "IndustryDataKeys": json.dumps(list(industry_data.keys())),
                "ProcessedData": json.dumps(processed_data),
                "Timestamp": datetime.now(timezone.utc).isoformat()
            }
            table_client.create_entity(entity=entity)
            logging.info("Successfully stored processed data in Azure Table Storage")
        except Exception as e:
            logging.error(f"Error storing data in table: {str(e)}")
            # Continue execution even if storage fails
        
        # Return success response with processed data
        return func.HttpResponse(
            json.dumps({
                "status": "success",
                "message": "Successfully processed 4 JSON files from Power Automate and saved PowerPoint to Azure Blob Storage",
                "data": processed_data,
                "files_received": {
                    "CompanyReseachData1_size": len(company_data1),
                    "CompanyReseachData2_size": len(company_data2),
                    "CompanyReseachData3_size": len(company_data3),
                    "IndustryResearch_size": len(industry_data)
                },
                "output_file": {
                    "filename": current_output_file,
                    "blob_path": current_output_file,
                    "container": CONTAINER_NAME,
                    "full_url": processed_data.get("output_location", "Not available")
                }
            }),
            status_code=200,
            mimetype="application/json"
        )
        
    except ValueError as e:
        # JSON parsing error
        return func.HttpResponse(
            json.dumps({
                "error": f"Invalid JSON in request body: {str(e)}",
                "status": "error"
            }),
            status_code=400,
            mimetype="application/json"
        )
    except Exception as e:
        # General error handling
        logging.error(f"Unexpected error: {str(e)}")
        return func.HttpResponse(
            json.dumps({
                "error": f"Internal server error: {str(e)}",
                "status": "error"
            }),
            status_code=500,
            mimetype="application/json"
        )
