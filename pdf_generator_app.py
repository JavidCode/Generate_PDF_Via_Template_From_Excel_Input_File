import pdfrw
import pandas as pd
from tkinter import filedialog, simpledialog
import tkinter as tk
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import os

def get_value(annotation, key):
    if isinstance(annotation[key], pdfrw.objects.PdfString):
        return annotation[key].decode()
    return annotation[key]

def extract_data_from_pdf(pdf_path, consignor_id, field_mapping):
    data = {col: [] for col in field_mapping['Excel']}

    # Read PDF form fields
    template_pdf = pdfrw.PdfReader(pdf_path)

    # Check if annotations are available
    if template_pdf.pages and template_pdf.pages[0].Annots:
        annotations = template_pdf.pages[0].Annots
    else:
        print("No annotations found in the PDF.")
        return data

    for annotation in annotations:
        if annotation['/Subtype'] == '/Widget' and annotation['/FT'] == '/Tx':
            field_name = get_value(annotation, '/T')  # Extract the field name
            if field_name.lower() in field_mapping['PDF']:
                excel_column = field_mapping['PDF'][field_name.lower()]
                field_value = get_value(annotation, '/V')  # Extract the field value
                data[excel_column].append(field_value)

    # Add consignor ID to the extracted data for each row
    data['Consignor'] = [consignor_id] * len(data.get('Size', []))

    return data

def save_output_as_pdf(pdf_data, output_directory, consignor_id, index):
    if not pdf_data['Description1']:
        print("Description list is empty. Skipping PDF generation.")
        return

    output_filename = f"LittleTreasures_Tag_{consignor_id}_{index + 1}.pdf"
    output_path = os.path.join(output_directory, output_filename)

    # Create a PDF document
    c = canvas.Canvas(output_path, pagesize=letter)
    c.drawString(100, 750, f"Size: {pdf_data['Size'][0]}")
    c.drawString(100, 730, f"Consignor: {pdf_data['Consignor'][0]}")

    # Iterate through Description fields
    for i in range(1, 10):
        description_key = f'Description{i}'
        if pdf_data.get(description_key):
            c.drawString(100, 710 - (20 * i), f"{description_key}: {pdf_data[description_key][0]}")

    c.save()

    print(f"PDF generated: {output_filename}")

def main():
    # Prompt the user to enter the consignor ID
    consignor_id = simpledialog.askstring("Input", "Enter Consignor ID:")

    if consignor_id is None:
        print("No Consignor ID entered. Exiting.")
        return

    # Prompt the user to select the PDF template file
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    pdf_template_path = filedialog.askopenfilename(title="Select PDF Template File", filetypes=[("PDF Files", "*.pdf")])

    if not pdf_template_path:
        print("No file selected. Exiting.")
        return

    # Prompt the user to select the Excel input file
    excel_input_path = filedialog.askopenfilename(title="Select Excel Input File", filetypes=[("Excel Files", "*.xlsx")])

    if not excel_input_path:
        print("No file selected. Exiting.")
        return

    # Load the Excel input file into a DataFrame
    try:
        input_df = pd.read_excel(excel_input_path)
    except pd.errors.EmptyDataError:
        print("Empty Excel file. Exiting.")
        return

    # Prompt the user to select the output directory
    output_directory = filedialog.askdirectory(title="Select Output Directory")

    if not output_directory:
        print("No output directory selected. Exiting.")
        return

    # Updated field_mapping
    field_mapping = {
        'PDF': {
            'size': 'Size',
            'description1': 'Description1',
            'size_2': 'Size',
            'description2': 'Description2',
            'size_3': 'Size',
            'description3': 'Description3',
            'size_4': 'Size',
            'description4': 'Description4',
            'size_5': 'Size',
            'description5': 'Description5',
            'size_6': 'Size',
            'description6': 'Description6',
            'size_7': 'Size',
            'description7': 'Description7',
            'size_8': 'Size',
            'description8': 'Description8',
            'size_9': 'Size',
            'description9': 'Description9',
        },
        'Excel': ['Size', 'Description1', 'Size', 'Description2', 'Size', 'Description3',
                  'Size', 'Description4', 'Size', 'Description5', 'Size', 'Description6',
                  'Size', 'Description7', 'Size', 'Description8', 'Size', 'Description9'],
    }

    # Iterate through the rows in the Excel file
    for index, row in input_df.iterrows():
        pdf_data = extract_data_from_pdf(pdf_template_path, consignor_id, field_mapping)
        save_output_as_pdf(pdf_data, output_directory, consignor_id, index)

if __name__ == "__main__":
    main()
