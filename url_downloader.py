import os
import requests
from urllib.parse import urlparse, unquote
import openpyxl


def download_image(url, folder, sku):
    # Create the folder if it doesn't exist
    if not os.path.exists(folder):
        os.makedirs(folder)

    # Extract the filename from the URL without query parameters
    parsed_url = urlparse(url)
    filename_with_ext = os.path.basename(unquote(parsed_url.path))

    # Split the filename and extension
    _, ext = os.path.splitext(filename_with_ext)

    # Ensure the file has a ".jpg" extension
    if not ext:
        ext = ".jpg"

    # Generate new image name based on SKU
    imgname = sku + ext

    # Download the image
    response = requests.get(url)
    if response.status_code == 200:
        # Save the image to the folder
        with open(os.path.join(folder, imgname), 'wb') as f:
            f.write(response.content)
        print("Image downloaded successfully:", imgname)
    else:
        print("Failed to download image.")


def process_excel_file(filename):
    # Open the Excel file
    try:
        workbook = openpyxl.load_workbook(filename)
    except Exception as e:
        print("Error opening Excel file:", e)
        return

    # Assuming the URLs and SKUs are in the first sheet
    sheet = workbook.active

    # Iterate over rows and download images
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if len(row) >= 2:
            sku = str(row[0]).strip()
            url = str(row[1]).strip()
            if url:
                download_image(url, "downloaded_images", sku)
            else:
                print("No URL provided for SKU:", sku)


if __name__ == "__main__":
    # Get all files in the current directory
    current_dir = os.getcwd()
    files = os.listdir(current_dir)

    # Filter for Excel files
    excel_files = [file for file in files if file.endswith('.xlsx') or file.endswith('.xls')]

    if not excel_files:
        print("No Excel files found in the current directory.")
    else:
        # Process each Excel file found
        for excel_file in excel_files:
            print(f"Processing Excel file: {excel_file}")
            process_excel_file(excel_file)
