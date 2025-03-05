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
options = {
    "enable-local-file-access": True,
    "page-size": "Letter",
    "margin-top": "0in",
    "margin-right": "0in",
    "margin-bottom": "0in",
    "margin-left": "0in",
}

# Load configuration from environment variables with default values
WK_HTML_TO_PDF = os.getenv(
    "WK_HTML_TO_PDF", "C:/Program Files/wkhtmltopdf/bin/wkhtmltopdf.exe"
)
DOC_PREFIX = os.getenv("DOC_PREFIX", "documents/")
DATA_FOLDER = os.getenv("DATA_FOLDER", "data/")
TEMPLATES_FOLDER = os.getenv("TEMPLATES_FOLDER", "templates/")
IMAGES_ABSOLUTE_PATH = os.getenv("IMAGES_ABSOLUTE_PATH", "")
SHEET_NAME = os.getenv("SHEET_NAME", "Sheet1")
DATA_FILE_NAME = os.getenv("DATA_FILE_NAME", "data-1.xlsx")


def setup_jinja_environment(templates_folder):
    """
    Setup the Jinja2 environment for template rendering.

    Args:
        templates_folder (str): The folder containing the templates.

    Returns:
        Environment: The configured Jinja2 environment.
    """
    return Environment(loader=FileSystemLoader(templates_folder))


def create_output_directory(output_folder):
    """
    Create the output directory if it doesn't exist.

    Args:
        output_folder (str): The path to the output directory.
    """
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)


def render_html_to_pdf(data, template, output_folder, index):
    """
    Render HTML with data and convert it to PDF.

    Args:
        data (dict): The data to render.
        template (Template): The Jinja2 template.
        output_folder (str): The path to the output directory.
        index (int): The index for naming the PDF file.

    Returns:
        bool: True if PDF generation is successful, False otherwise.
    """
    rendered_html = template.render(data)
    pdf_filename = os.path.join(output_folder, f"{index}_{data['Column1']}.pdf")

    try:
        pdfkit.from_string(
            rendered_html, pdf_filename, configuration=config, options=options
        )
        logging.info(f"PDF generated successfully: {pdf_filename}")
        return True
    except Exception as e:
        logging.error(f"Error generating PDF for index {index}: {e}")
        return False


def main():
    """
    Main function to generate PDFs from data using a template.
    """
    # Setup Jinja2 environment
    env = setup_jinja_environment(TEMPLATES_FOLDER)
    template = env.get_template("pdf_template.html")

    # Create output directory
    output_folder = os.path.join(DOC_PREFIX, "generated_pdfs")
    create_output_directory(output_folder)

    # Manual data insertion
    data = {
        "Column1": "Value 1",
        "Column2": "Value 2",
        "Column3": "Value 3",
    }

    # Render HTML with data and generate PDF
    success = render_html_to_pdf(data, template, output_folder, 522)

    if success:
        print("PDF generated successfully!")
    else:
        print("Failed to generate PDF.")


if __name__ == "__main__":
    main()