import pandas as pd
from jinja2 import Environment, FileSystemLoader
import pdfkit
import os
import logging
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# Configure logging
logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)

# Initialize the Generator environment and options
DATA_FILE_NAME = os.getenv("DATA_FILE_NAME")
SHEET_NAME = os.getenv("SHEET_NAME")
IMAGES_ABSOLUTE_PATH = os.getenv("IMAGES_ABSOLUTE_PATH")
TEMPLATES_FOLDER = os.getenv("TEMPLATES_FOLDER")
DATA_FOLDER = os.getenv("DATA_FOLDER")
DOC_PREFIX = os.getenv("DOC_PREFIX")
WK_HTML_TO_PDF = os.getenv("WK_HTML_TO_PDF")

OUTPUT_FOLDER = "generated_pdfs"
EXCEPTIONS_FILE_NAME = "exceptions.xlsx"
HTML_FILE = "pdf_template.html"
IMAGE_SUFFIX = "_image.jpg"

options = {
    "enable-local-file-access": True,
    "page-size": "A4",
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


def clean_text(text):
    """
    Clean and format text by replacing specific Unicode characters and symbols.

    Args:
        text (str): The text to be cleaned.

    Returns:
        str: The cleaned text.
    """
    text = (
        text.replace("\u200c", " ")
        .replace("\n", " ")
        .replace("*", "", 1)
        .replace("*", ", ")
        .replace("/", " ")
    )
    return text


def clean_col(text):
    """
    Clean and format text for specific columns, replacing '*' with new lines.

    Args:
        text (str): The text to be cleaned.

    Returns:
        str: The cleaned text with new lines.
    """
    text = (
        text.replace("\u200c", " ")
        .replace("\n", " ")
        .replace("*", "- ", 1)
        .replace("*", "<br> - ")
    )
    return text


def clean_long_text(text):
    """
    Clean and format long text by reversing and replacing specific characters.

    Args:
        text (str): The text to be cleaned.

    Returns:
        str: The cleaned text with formatted new lines.
    """
    reversed_text = text[::-1]
    replaced_text = reversed_text.replace(".", "/$/", 1)
    final_text = replaced_text[::-1]
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
    """
    Generate HTML for static checkboxes based on the provided types list.

    Args:
        types_list (list): List of types for checkboxes.
        data (str): Data to check against the types list.

    Returns:
        str: HTML string of checkboxes.
    """
    checkboxes_html = ""
    for item in types_list:
        checked = "checked" if item in data else ""
        checkboxes_html += f'<label style="font-family: sans-serif;"><input type="checkbox" disabled {checked}> {item}</label><br>'
    return checkboxes_html


def load_excel_data(sheet_name):
    """
    Load data from the specified Excel sheet.

    Args:
        sheet_name (str): The name of the sheet to load.

    Returns:
        pd.DataFrame: The loaded DataFrame.
    """
    try:
        df = pd.read_excel(
            DATA_FOLDER + DATA_FILE_NAME,
            sheet_name=sheet_name,
            engine="openpyxl",
            header=0,
        )
        return df
    except Exception as e:
        logging.error(f"Error reading Excel file: {e}")
        return pd.DataFrame()


def clean_dataframe(df):
    """
    Clean the DataFrame columns.

    Args:
        df (pd.DataFrame): The DataFrame to clean.

    Returns:
        pd.DataFrame: The cleaned DataFrame.
    """
    for column in df.columns:
        if column == "Column1":
            df[column] = df[column].apply(
                lambda x: clean_col(x) if isinstance(x, str) else x
            )
        elif column == "Column2":
            df[column] = df[column].apply(
                lambda x: x.replace("/", " ") if isinstance(x, str) else x
            )
        elif column == "Column3":
            df[column] = df[column].apply(
                lambda x: clean_long_text(x) if isinstance(x, str) else x
            )
        else:
            df[column] = df[column].apply(
                lambda x: clean_text(x) if isinstance(x, str) else x
            )
    return df


def generate_image_paths(df):
    """
    Generate image paths based on DataFrame index.

    Args:
        df (pd.DataFrame): The DataFrame to process.

    Returns:
        pd.DataFrame: The DataFrame with image paths added.
    """
    df["image"] = df.index.to_series().apply(
        lambda x: f"{IMAGES_ABSOLUTE_PATH}{x}_image.png"
    )
    logging.info(f"{df["image"].count()} images paths are generated successfully.")
    logging.info(f"{df['image'][0]}")
    return df


def setup_jinja_environment():
    """
    Setup the Jinja2 environment for template rendering.

    Returns:
        Environment: The configured Jinja2 environment.
    """
    return Environment(loader=FileSystemLoader("."))


def create_output_directory():
    """
    Create the output directory if it doesn't exist.
    """
    if not os.path.exists(DOC_PREFIX + OUTPUT_FOLDER):
        os.makedirs(DOC_PREFIX + OUTPUT_FOLDER)


def generate_pdfs(df, template, config):
    """
    Generate PDFs from the DataFrame using the provided template.

    Args:
        df (pd.DataFrame): The DataFrame containing the data.
        template (Template): The Jinja2 template.
        config (pdfkit.configuration): The PDFKit configuration.

    Returns:
        dict: A dictionary of exceptions encountered during PDF generation.
    """
    exception_rows = {}
    count = 0

    for index, row in df.iterrows():
        data = row.to_dict()
        data["Column4"] = generate_static_checkboxes(
            types_list, data.get("Column3", "")
        )

        rendered_html = template.render(data)
        pdf_filename = os.path.join(
            DOC_PREFIX + OUTPUT_FOLDER, f"{index}_{data['Column1']}.pdf"
        )

        try:
            pdfkit.from_string(
                rendered_html, pdf_filename, configuration=config, options=options
            )
            logging.info(f"PDF generated successfully: {pdf_filename}")
        except Exception as e:
            logging.error(f"Error generating PDF for index {index}: {e}")
            exception_rows[count] = {"index": index, "row": row, "error": str(e)}
            count += 1

    return exception_rows


def log_exceptions(exceptions):
    """
    Log exceptions to an Excel file.

    Args:
        exceptions (dict): A dictionary of exceptions.
    """
    if exceptions:
        exceptions_df = pd.DataFrame(exceptions)
        exceptions_df.to_excel(DOC_PREFIX + EXCEPTIONS_FILE_NAME, index=False)
        logging.warning(f"Exceptions logged to {DOC_PREFIX + EXCEPTIONS_FILE_NAME}")
    else:
        logging.info("All PDFs generated successfully.")


def main():
    """
    Main function to generate PDFs from Excel data using a template.
    """
    df_list = [load_excel_data(SHEET_NAME)]
    df_combined = pd.concat(df_list, ignore_index=False)
    df_combined = clean_dataframe(df_combined)
    df_combined = generate_image_paths(df_combined)

    env = setup_jinja_environment()
    template = env.get_template(TEMPLATES_FOLDER + HTML_FILE)

    create_output_directory()

    config = pdfkit.configuration(wkhtmltopdf=WK_HTML_TO_PDF)
    exceptions = generate_pdfs(df_combined, template, config)
    log_exceptions(exceptions)


if __name__ == "__main__":
    main()
    logging.info("PDF generation process completed.")
