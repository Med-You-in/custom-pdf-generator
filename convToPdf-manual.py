import pandas as pd
from jinja2 import Environment, FileSystemLoader
import pdfkit
import os


options = {
    "enable-local-file-access": True,
    "page-size": "Letter",
    "margin-top": "0in",
    "margin-right": "0in",
    "margin-bottom": "0in",
    "margin-left": "0in",
}


# Load Excel data
df_list = []

# Setup Jinja2 environment
env = Environment(loader=FileSystemLoader("."))
template = env.get_template("pdf_template.html")

# Define output folder
output_folder = "generated_pdfs"
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

exception_rows = {}
count = 0
config = pdfkit.configuration(
    wkhtmltopdf="C:/Program Files/wkhtmltopdf/bin/wkhtmltopdf.exe"
)

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
