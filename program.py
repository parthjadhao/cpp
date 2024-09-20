import os
from PIL import Image
from odf.opendocument import OpenDocumentSpreadsheet
from odf.table import Table, TableRow, TableCell
from odf.text import P

def calculate_average_rgb(image_path):
    # Open the image
    with Image.open(image_path) as img:
        # Convert image to RGB
        img = img.convert("RGB")
        pixels = list(img.getdata())
        
        # Calculate the sum of R, G, B pixels
        total_pixels = len(pixels)
        r_total, g_total, b_total = 0, 0, 0
        
        for r, g, b in pixels:
            r_total += r
            g_total += g
            b_total += b
            
        # Calculate averages
        r_avg = r_total / total_pixels
        g_avg = g_total / total_pixels
        b_avg = b_total / total_pixels
        
    return r_avg, g_avg, b_avg

def process_images(directory_path, output_ods_path):
    # Create a new ODS document
    ods_doc = OpenDocumentSpreadsheet()
    table = Table(name="Image RGB Data")
    
    # Add header row
    header_row = TableRow()
    headers = ["Image Name", "Class", "Red Avg", "Green Avg", "Blue Avg"]
    for header in headers:
        cell = TableCell()
        cell.addElement(P(text=header))
        header_row.addElement(cell)
    table.addElement(header_row)
    
    # Loop through all the files in the directory
    for filename in os.listdir(directory_path):
        if filename.endswith((".png", ".jpg", ".jpeg")):  # Check for image files
            image_path = os.path.join(directory_path, filename)
            
            # Assuming class is part of the filename, e.g., class_image1.jpg
            image_class = filename.split("_")[0]
            
            # Calculate the average RGB values
            r_avg, g_avg, b_avg = calculate_average_rgb(image_path)
            
            # Add a row with the image data
            row = TableRow()
            data = [filename, image_class, r_avg, g_avg, b_avg]
            for item in data:
                cell = TableCell()
                cell.addElement(P(text=str(item)))
                row.addElement(cell)
            table.addElement(row)
    
    # Add the table to the ODS document
    ods_doc.spreadsheet.addElement(table)
    
    # Save the ODS file
    ods_doc.save(output_ods_path)
    print(f"Data saved to {output_ods_path}")

# Usage
directory_path = "./asset"  # Specify the path to your images folder
output_ods_path = "./processedData.ods"         # Specify the path for output ODS file
process_images(directory_path, output_ods_path)