import pandas as pd
from jinja2 import Environment, FileSystemLoader
import pdfkit
import os

#  Initialize the Generator environment and options =======================================================================================
# *************** Check those before running the script *****************
DATA_FILE_NAME = "data-1.xlsx"
SHEET_NAME = "Sheet1"
IMAGES_ABSOLUTE_PATH = (
    "....\documents\images\\"  # set the absolute path for the the images if exists
)
# ***********************************************************************

#  Set the following variables according to your needs
DOC_PREFIX = "documents/"
OUTPUT_FOLDER = "generated_pdfs"
EXCEPTIONS_FILE_NAME = "exceptions.xlsx"
HTML_FILE = "pdf_template.html"
WK_HTML_TO_PDF = "C:/Program Files/wkhtmltopdf/bin/wkhtmltopdf.exe" # try to install wkhtmltopdf from https://wkhtmltopdf.org/downloads.html
IMAGE_SUFFIX = "_image.jpg"

# Initialize the Generator environment and options
options = {
    "enable-local-file-access": True,
    "page-size": "Letter",
    "margin-top": "0in",
    "margin-right": "0in",
    "margin-bottom": "0in",
    "margin-left": "0in",
}

types_list = [
    "Type 1",
    "Type 2",
    "Type 3",
    "Type 4",
]

# Generator functions =======================================================================================


def clean_text(text):
    # Replace specific Unicode characters with a space or an empty string as needed
    text = (
        text.replace("\u200c", " ")
        .replace("\n", " ")
        .replace("*", "", 1)
        .replace("*", ", ")
        .replace("/", " ")
    )
    return text


def clean_col(text):
    # Specific cleaning for 'Coordonnées' column to replace '*' with new lines
    text = (
        text.replace("\u200c", " ")
        .replace("\n", " ")
        .replace("*", "- ", 1)
        .replace("*", "<br> - ")
    )
    return text


def clean_long_text(text):
    reversed_text = text[::-1]
    replaced_text = reversed_text.replace(".", "/$/", 1)
    final_text = replaced_text[::-1]
    # Specific cleaning for 'Coordonnées' column to replace '*' with new lines
    fntext = (
        final_text.replace("\u200c", " ")
        .replace("*", "", 1)
        .replace("-", "")
        .replace(".", "<br> - ")
        .replace(",", "", -1)
        .replace("/$/", ".")
    )
    return fntext


def generate_static_checkboxes(types_list, data):
    checkboxes_html = ""
    for item in types_list:
        checked = "checked" if item in data else ""
        checkboxes_html += f'<label style="font-family: sans-serif;"><input type="checkbox" disabled {checked}> {item}</label><br>'
    return checkboxes_html


# Generation Logic =======================================================================================

# Load Excel data
df_list = []

for sheet_name in [SHEET_NAME]:  # Specify your sheet names
    df = pd.read_excel(
        DOC_PREFIX + DATA_FILE_NAME, sheet_name=sheet_name, engine="openpyxl", header=0
    )
    df_list.append(df)

df_combined = pd.concat(df_list, ignore_index=False)

# Clean the data
for column in df_combined.columns:

    if column == "Column1":
        df_combined[column] = df_combined[column].apply(
            lambda x: clean_col(x) if isinstance(x, str) else x
        )
    # Apply the cleaning function to string columns only
    elif column == "Column2":
        df_combined[column] = df_combined[column].apply(
            lambda x: x.replace("/", " ") if isinstance(x, str) else x
        )
    elif column == "Column3":
        df_combined[column] = df_combined[column].apply(
            lambda x: clean_long_text(x) if isinstance(x, str) else x
        )
    else:
        df_combined[column] = df_combined[column].apply(
            lambda x: clean_text(x) if isinstance(x, str) else x
        )


# This will create a column where each row's value is its index followed by '_image.jpg'
df_combined["image"] = df_combined.index.to_series().apply(
    lambda x: f"{IMAGES_ABSOLUTE_PATH}{x}_image.jpg"
)

# Setup Jinja2 environment
env = Environment(loader=FileSystemLoader("."))
template = env.get_template(DOC_PREFIX + HTML_FILE)

if not os.path.exists(DOC_PREFIX + OUTPUT_FOLDER):
    os.makedirs(DOC_PREFIX + OUTPUT_FOLDER)

exception_rows = {}
count = 0
config = pdfkit.configuration(wkhtmltopdf=WK_HTML_TO_PDF)
# Loop through each row in DataFrame and create a PDF
for index, row in df_combined.iterrows():
    # Prepare data for template
    data = row.to_dict()
    data["Column4"] = generate_static_checkboxes(types_list, data.get("Column3", ""))

    # Render HTML with data
    rendered_html = template.render(data)
    # Define PDF file path
    pdf_filename = os.path.join(
        DOC_PREFIX + OUTPUT_FOLDER, f"{index}_{data['Column1']}.pdf"
    )

    # Convert HTML to PDF and save
    try:
        pdfkit.from_string(
            rendered_html, pdf_filename, configuration=config, options=options
        )
    except Exception as e:
        # handle any other exceptions
        print(f"An unexpected error occurred: {e}")
        exception_rows[count] = {"index": index, "row": row, "error": str(e)}
        count += 1
        print("exception")


# Save the exceptions DataFrame to an Excel file
# Create a DataFrame from the list of dictionaries
exceptions_df = pd.DataFrame(exception_rows)
exceptions_df.to_excel(DOC_PREFIX + EXCEPTIONS_FILE_NAME, index=False)

print("Done")
