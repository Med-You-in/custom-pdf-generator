import pandas as pd
from jinja2 import Environment, FileSystemLoader
import pdfkit
import os
from dotenv import load_dotenv  # Import the load_dotenv function

# Load environment variables from .env file
load_dotenv()

# Initialize the Generator environment and options
options = {
    "enable-local-file-access": True,
    "page-size": "Letter",
    "margin-top": "0in",
    "margin-right": "0in",
    "margin-bottom": "0in",
    "margin-left": "0in",
}

# Load configuration from environment variables
WK_HTML_TO_PDF = os.getenv(
    "WK_HTML_TO_PDF", "C:/Program Files/wkhtmltopdf/bin/wkhtmltopdf.exe"
)  # Default value if not found
DOC_PREFIX = os.getenv("DOC_PREFIX", "documents/")  # Default value if not found
DATA_FOLDER = os.getenv("DATA_FOLDER", "data/")  # Default value if not found
TEMPLATES_FOLDER = os.getenv(
    "TEMPLATES_FOLDER", "templates/"
)  # Default value if not found
IMAGES_ABSOLUTE_PATH = os.getenv(
    "IMAGES_ABSOLUTE_PATH", ""
)  # Default value if not found
SHEET_NAME = os.getenv("SHEET_NAME", "Sheet1")  # Default value if not found
DATA_FILE_NAME = os.getenv(
    "DATA_FILE_NAME", "data-1.xlsx"
)  # Default value if not found

# Load Excel data
df_list = []

# Setup Jinja2 environment
env = Environment(loader=FileSystemLoader(TEMPLATES_FOLDER))
template = env.get_template("pdf_template.html")

# Define output folder
output_folder = os.path.join(DOC_PREFIX, "generated_pdfs")
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

exception_rows = {}
count = 0
config = pdfkit.configuration(wkhtmltopdf=WK_HTML_TO_PDF)

# Manual data insertion
data = {
    "Column1": "Value 1",
    "Column2": "Value 2",
    "Column3": "Value 3",
}

# Render HTML with data
rendered_html = template.render(data)
# Define PDF file path
pdf_filename = os.path.join(output_folder, f"{522}.{data['Column1']}.pdf")

# Convert HTML to PDF and save
try:
    pdfkit.from_string(
        rendered_html, pdf_filename, configuration=config, options=options
    )
except Exception as e:
    # handle any other exceptions
    print(f"An unexpected error occurred: {e}")

print("PDF generated successfully!")
