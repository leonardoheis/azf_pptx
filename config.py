import os

# Azure Storage Configuration
STORAGE_CONN_STRING = os.environ.get("AzureWebJobsStorage", "")

# Table Configuration
TABLE_NAME = "PPTX_Activity"

# Thread Configuration
THREAD_TIMEOUT_MINUTES = int(os.environ.get("ThreadTimeout", "30"))

# Conversion factor: 1 pt = 12700 EMU
EMU_PER_PT = 12700
