import os
import glob
from datetime import datetime

# Azure Storage Configuration
STORAGE_CONN_STRING = os.environ.get("AzureWebJobsStorage", "")

# Blob Storage Configuration
# Set this environment variable in your Azure Function app settings:
# BLOB_CONTAINER_NAME: Name of the container to store PowerPoint files (default: "pptx-out")
CONTAINER_NAME = os.environ.get("BLOB_CONTAINER_NAME", "pptx-out")

# Table Configuration
TABLE_NAME = "pptxactivity"

# Thread Configuration
THREAD_TIMEOUT_MINUTES = int(os.environ.get("ThreadTimeout", "30"))

# Conversion factor: 1 pt = 12700 EMU
EMU_PER_PT = 12700

TEMPLATE = "templates/pantilla.pptx"

def get_next_output_filename():
    """
    Generates the next available output filename with timestamp for blob storage.
    
    Returns:
        str: Next available filename (e.g., "output_20241201_143022.pptx")
    """
    base_name = "output"
    extension = ".pptx"
    
    # Generate timestamp-based filename to ensure uniqueness
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return f"{base_name}_{timestamp}{extension}"

# Default output file (will be overridden by get_next_output_filename())
OUTPUT_FILE = get_next_output_filename()
