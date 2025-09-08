import os
from datetime import datetime

from pptx.dml.color import RGBColor

# Azure Storage Configuration
AZ_STORAGE_CONN_STRING = os.environ.get("AZ_STORAGE_CONN_STRING") or os.environ.get("AzureWebJobsStorage", "")

# Blob Storage Configuration
# Set this environment variable in your Azure Function app settings:
# BLOB_CONTAINER_NAME: Name of the container to store PowerPoint files (default: "pptx-out")
AZ_BLOB_CONTAINER_NAME = os.environ.get("BLOB_CONTAINER_NAME", "pptx-out")

# Azure Blob Table Storage Name
AZ_BLOB_TABLE_NAME = "pptxactivity"

# Thread Configuration
THREAD_TIMEOUT_MINUTES = int(os.environ.get("ThreadTimeout", "30"))

# Conversion factor: 1 pt = 12700 EMU
EMU_PER_PT = 12700

INPUT_TEMPLATE = "template/plantilla.pptx"

TABLE_HEADER_HEIGHT_PT = 24
TABLE_LINE_HEIGHT_PT = 12

# Table styling colors
TABLE_HEADER_BG_COLOR = RGBColor(0, 70, 122)
TABLE_HEADER_TEXT_COLOR = RGBColor(255, 255, 255)

TABLE_PARAGRAPH_FONT_SIZE_PT = 14

DICT_LIST_CONTENT_FONT_SIZE_PT = 10
SIMPLE_LIST_CONTENT_FONT_SIZE_PT = 10
MULTILINE_CONTENT_FONT_SIZE_PT = 10
SIMPLE_CONTENT_FONT_SIZE_PT = 10


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
