import sys
from dotenv import load_dotenv
import os

# Modify this if needed
print(sys.path)
sys.path.append('src/classes')

from document_extraction import DocumentExtraction
from document_output import DocumentOutput


## >> Exemple
# Load environment variables from .env file
load_dotenv()

# Access the variables
BASE_PATH = os.getenv("BASE_PATH")
DATA_PATH = os.getenv("DATA_PATH")
RAW_PATH = os.getenv("RAW_PATH")
OUTPUT_PATH = os.getenv("OUTPUT_PATH")
DOCUMENT_NAME = os.getenv("DOCUMENT_NAME")
DOCUMENT_PATH = os.getenv("DOCUMENT_PATH")


## >> Extracting the document content
if __name__ == '__main__':
    document_extractor = DocumentExtraction(DOCUMENT_PATH)
    document_dataframes = document_extractor.extract()
    document_extractor.export(output_format = DocumentOutput.CSV, document_output_path = OUTPUT_PATH)
    document_extractor.export(output_format = DocumentOutput.JSON, document_output_path = OUTPUT_PATH)