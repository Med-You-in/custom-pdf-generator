import os
import requests
import pandas as pd
import logging

# Configure logging
logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)

IMAGE_SUFFIX = "_image.jpg"
DOC_PREFIX = "documents/"

# Specify the path to your Excel file and the output folder
excel_path = DOC_PREFIX + "data-1.xlsx"
column_name = "liendulogo"
output_folder = DOC_PREFIX + "images_MS"


def read_excel_file(file_path):
    """
    Read the Excel file and return a DataFrame.

    Args:
        file_path (str): The path to the Excel file.

    Returns:
        pd.DataFrame: The DataFrame containing the Excel data.
    """
    try:
        df = pd.read_excel(file_path)
        logging.info(f"Successfully read Excel file: {file_path}")
        return df
    except Exception as e:
        logging.error(f"Error reading Excel file: {e}")
        return pd.DataFrame()


def create_output_directory(directory_path):
    """
    Create the output directory if it doesn't exist.

    Args:
        directory_path (str): The path to the output directory.
    """
    if not os.path.exists(directory_path):
        os.makedirs(directory_path)
        logging.info(f"Created output directory: {directory_path}")


def download_image(image_url, file_path):
    """
    Download an image from a URL and save it to the specified file path.

    Args:
        image_url (str): The URL of the image to download.
        file_path (str): The path where the image will be saved.

    Returns:
        bool: True if the download was successful, False otherwise.
    """
    try:
        response = requests.get(image_url)
        response.raise_for_status()  # Raise an error on bad status

        with open(file_path, "wb") as file:
            file.write(response.content)
        logging.info(f"Downloaded image to {file_path}")
        return True
    except requests.RequestException as e:
        logging.error(f"Failed to download {image_url}: {e}")
        return False


def download_images(excel_path, output_folder, column_name):
    """
    Download images from URLs in the specified column of the Excel file.

    Args:
        excel_path (str): The path to the Excel file.
        output_folder (str): The path to the output folder.
        column_name (str): The name of the column containing image URLs.
    """
    df = read_excel_file(excel_path)
    create_output_directory(output_folder)

    for index, row in df.iterrows():
        image_url = row[column_name]
        if pd.isna(image_url):
            continue

        file_path = os.path.join(output_folder, f"{index}{IMAGE_SUFFIX}")
        download_image(image_url, file_path)


if __name__ == "__main__":
    download_images(excel_path, output_folder, column_name)
    logging.info("Image download process completed.")
