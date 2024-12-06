
import os
import cv2
from PIL import Image
import pytesseract
import pandas as pd
import re

def extract_player_names_from_image(image_path):
    """
    Extract player names from a given image file by cropping the relevant section,
    and clean up extra newlines.
    
    Args:
        image_path (str): Path to the image file.
    
    Returns:
        str: Cleaned extracted player names as a string.
    """
    # Load the image
    img = cv2.imread(image_path)

    # Define the region to crop (x_start, y_start, x_end, y_end)
    x_start, y_start, x_end, y_end = 100, 100, 450, 500  # Adjust as necessary

    # Crop the image to only include the player names column
    cropped_img = img[y_start:y_end, x_start:x_end]

    # Convert to grayscale (optional but might improve OCR accuracy)
    gray = cv2.cvtColor(cropped_img, cv2.COLOR_BGR2GRAY)

    # Apply thresholding to binarize the image
    _, thresh_img = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY)

    # Use Tesseract to extract text from the processed image
    img_pil = Image.fromarray(thresh_img)
    player_names_text = pytesseract.image_to_string(img_pil)

    # Clean up the text: remove empty lines and extra spaces
    cleaned_names = "\n".join([line.strip() for line in player_names_text.splitlines() if line.strip()])

    return cleaned_names

def extract_date_from_filename(filename):
    """
    Extract date from filename based on a specific pattern (YYYYMMDD).
    
    Args:
        filename (str): The name of the file.
    
    Returns:
        str: Extracted date in YYYY-MM-DD format.
    """
    # Regex pattern to match the date in the filename
    date_match = re.search(r'_(\d{8})_', filename)
    if date_match:
        date_str = date_match.group(1)
        # Format the date string to YYYY-MM-DD
        formatted_date = f"{date_str[:4]}-{date_str[4:6]}-{date_str[6:]}"
        return formatted_date
    return None

def process_folder(folder_path):
    """
    Process all .webp files in a folder to extract player names and their corresponding dates,
    and return them as a list of dictionaries.
    
    Args:
        folder_path (str): Path to the folder containing .webp files.
    
    Returns:
        list: A list of dictionaries with 'Player Name' and 'Date' extracted from all images.
    """
    player_data = []
    
    # Loop through all files in the folder
    for filename in os.listdir(folder_path):
        if filename.endswith(".webp"):
            # Get the full path of the image file
            image_path = os.path.join(folder_path, filename)
            
            # Extract player names from the image
            player_names = extract_player_names_from_image(image_path)
            
            # Extract the date from the filename
            date = extract_date_from_filename(filename)
            
            # Split player names by newline and create a dictionary for each
            for name in player_names.split("\n"):
                if name:  # Skip empty names
                    player_data.append({"Player Name": name, "Date": date})
    
    return player_data

def export_to_excel(player_data, output_file):
    """
    Export the player data (names and dates) to an Excel file.
    
    Args:
        player_data (list): List of dictionaries with player names and dates.
        output_file (str): The path to the output Excel file.
    """
    # Create a DataFrame from the list of dictionaries
    df = pd.DataFrame(player_data)
    
    # Export the DataFrame to Excel
    df.to_excel(output_file, index=False)

# Define the folder path and output file
folder_path = "C:/ROM"
output_file = "C:/ROM/player_names_with_dates.xlsx"

# Process the folder and get all player names and dates
all_player_data = process_folder(folder_path)

# Export the player names and dates to Excel
export_to_excel(all_player_data, output_file)

print(f"Player names with dates exported to {output_file}")
