import os
import openpyxl
from PIL import Image, ImageDraw, ImageFont
from photoshop import Session

# Define the path to your spreadsheet and images
spreadsheet_path = '/Users/mac/Documents/GitHub/ai-photoshop-auto/translations.xlsx'
images_folder = '/Users/mac/Documents/aiauto_folder/aiauto_input'
output_folder = '/Users/mac/Documents/aiauto_folder/aiauto_output'

# Load the translations from the spreadsheet
wb = openpyxl.load_workbook(spreadsheet_path)
sheet = wb.active

translations = {}
for row in sheet.iter_rows(min_row=2, values_only=True):
    filename, original_text, translated_text = row
    translations[filename] = translated_text

# Create output folder if it doesn't exist
os.makedirs(output_folder, exist_ok=True)

# Define a function to process each image
def process_image(filename, translated_text):
    image_path = os.path.join(images_folder, filename)
    output_path = os.path.join(output_folder, filename)
    
    with Session() as ps:
        # Open the image in Photoshop
        ps.app.open(image_path)
        doc = ps.app.activeDocument
        
        # Replace the text layer with the translated text
        for layer in doc.artLayers:
            if layer.kind == ps.LayerKind.TextLayer:
                layer.textItem.contents = translated_text
        
        # Save the image
        doc.saveAs(output_path, ps.JPEGSaveOptions())
        doc.close()

# Process each image
for filename, translated_text in translations.items():
    process_image(filename, translated_text)

print("Processing complete")
