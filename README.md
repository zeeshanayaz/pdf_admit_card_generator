# Admit Card Generator

A Python tool to automate the generation of admit cards in bulk using a provided template and student data from an Excel file.

## Features
- Generate admit cards in bulk.
- Use a custom PDF template as the background.
- Save each page as a separate PDF file.
- Easy-to-use and customizable.

## Requirements
- Python 3.x
- Libraries: `pandas`, `reportlab`, `PyPDF2`, `openpyxl`

## Installation
1. Clone the repository:
   ```bash
   git clone https://github.com/zeeshanayaz/pdf_admit_card_generator.git
   ```
2. Install the required libraries:
   ```bash
   pip install pandas reportlab pypdf2 openpyxl
   ```

## Usage
1. Place your Excel file `(students.xlsx)`
2. Place your admit card template `(admit_card.pdf)`
3. Run the script:
   ```bash
   python ./generate_admit_cards.py
   ```
5. The generated admit cards will be saved in `admit_card_1.pdf`.

## Contributing
Contributions are welcome! Please open an issue or submit a pull request.
