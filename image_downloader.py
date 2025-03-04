import os
import requests
import pandas as pd

IMAGE_SUFFIX = "_image.jpg"
DOC_PREFIX = "documents/"
# Use this if you have a set of images on a column on your data excel file
# Specify the path to your Excel file and the output folder
excel_path = DOC_PREFIX + "data-1.xlsx"
column_name = "liendulogo"
output_folder = DOC_PREFIX + "images_MS"


def download_images(excel_path, output_folder):
    # Read the Excel file
    df = pd.read_excel(excel_path)

    # Ensure the output folder exists
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # Loop through the selected column on column_name
    for index, row in df.iterrows():
        image_url = row[column_name]
        # Check if the URL is a valid URL
        if pd.isna(image_url):
            continue

        try:
            # Download the image
            response = requests.get(image_url)
            response.raise_for_status()  # Raise an error on bad status

            # Save the image to the specified folder
            file_path = os.path.join(output_folder, f"{index}{IMAGE_SUFFIX}")
            with open(file_path, "wb") as file:
                file.write(response.content)
            print(f"Downloaded {file_path}")
        except requests.RequestException as e:
            print(f"Failed to download {image_url}: {e}")


download_images(excel_path, output_folder)
