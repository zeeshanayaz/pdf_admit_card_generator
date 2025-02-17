import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from PyPDF2 import PdfReader, PdfWriter

# Load the student data from Excel
excel_file = 'students.xlsx'  # Name of your Excel file
df = pd.read_excel(excel_file)  # Read the Excel file into a DataFrame

# PDF settings
output_folder = 'admit_cards'  # Folder to save individual admit card files
page_size = A4  # Set page size to A4
margin = 0.2 * 72  # Convert 0.2 inches to points (1 inch = 72 points)
card_width = (page_size[0] - 2 * margin) / 2  # Two cards per row
card_height = (page_size[1] - 2 * margin) / 4  # Four cards per column

# Function to create a single admit card overlay
def create_admit_card_overlay(c, x, y, student_data):
    c.setFont("Helvetica", 10)
    c.drawString(x + 70, y + 127, f"{student_data['GR #']}")
    c.drawString(x + 70, y + 107, f"{student_data['Candidate Name']}")
    c.drawString(x + 70, y + 88, f"{student_data['Father Name']}")
    c.drawString(x + 70, y + 70, f"{student_data['Grade']}")
    # c.drawString(x + 80, y + 55, f"{student_data['Campus']}")

# Create the PDF overlay
overlay_pdf = 'overlay.pdf'
c = canvas.Canvas(overlay_pdf, pagesize=page_size)

# Loop through the student data and create admit card overlays
for index, row in df.iterrows():
    # Calculate the position for the current card
    row_num = (index // 2) % 4  # 0 to 3 (4 rows per page)
    col_num = index % 2  # 0 or 1 (2 columns per row)
    
    x = margin + col_num * card_width
    y = page_size[1] - margin - (row_num + 1) * card_height
    
    # Create the admit card overlay
    create_admit_card_overlay(c, x, y, row)
    
    # Add a new page after every 8 cards
    if (index + 1) % 8 == 0:
        c.showPage()  # End the current page and start a new one

# Save the overlay PDF
c.save()

# Merge the overlay with the background template
background_pdf = 'admit_card.pdf'  # Your provided background template

# Read the background and overlay PDFs
background = PdfReader(background_pdf)
overlay = PdfReader(overlay_pdf)

# Create a folder to save individual admit card files
import os
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# Save each page as a separate PDF file
for i in range(len(overlay.pages)):
    # Create a fresh copy of the background page
    background_page = background.pages[0]  # Always use the first (blank) page of the background
    overlay_page = overlay.pages[i]
    background_page.merge_page(overlay_page)
    
    # Create a new PDF writer for the current page
    writer = PdfWriter()
    writer.add_page(background_page)
    
    # Save the current page as a separate PDF file
    page_output_pdf = os.path.join(output_folder, f'admit_card_page_{i + 1}.pdf')
    with open(page_output_pdf, 'wb') as out_pdf:
        writer.write(out_pdf)

    print(f"Page {i + 1} saved as '{page_output_pdf}'.")

print("All admit card pages saved successfully!")