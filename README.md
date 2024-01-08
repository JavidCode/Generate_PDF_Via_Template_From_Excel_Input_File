# Generate_PDF_Via_Template_From_Excel_Input_File
 
Requirements:
Input Data:

Accepts a PDF template and an Excel spreadsheet with two tabs: "Boys" and "Girls."
Each tab contains columns: "Size," "Description," and "Price."
PDF Output:

Generates a PDF with fields "Description" and "Size."
Each PDF may have up to nine tags for multiple products.
Output File Handling:

Saves the generated PDFs to a specified directory.
Appends the cosigner ID and a unique number to each output file name.
Data Mapping:

Maps data from the "Boys" tab to one set of PDFs and from the "Girls" tab to another set.
Multiple Sheets Handling:

Handles the scenario where there are more than one sheet in the Excel file.
Scalability:

Supports the generation of multiple PDFs based on the available data.
Error Handling:

Provides clear error messages if the input data is not in the expected format or if any issues arise during the process.
User Interface (Optional):

Optionally, you may consider a simple command-line interface or configuration file for user input.
Dependencies:

Clearly specifies any external libraries or dependencies required for the program to run.
Documentation:

PDF Output:
Generates PDFs with fields "Description" and "Size."  
Each PDF can contain up to nine tags for different products.
Supports the possibility of a seller having a large number of items (e.g., up to 500), resulting in multiple PDFs.
Ensures proper organization and naming conventions for each PDF, considering the seller's inventory.

UPC Generation and Superimposition:
Generates a unique UPC (Universal Product Code) for each product.
Superimposes the generated UPC onto the existing tag in the PDF, ensuring visibility and adherence to standards.
Provides options or settings for UPC format and placement on the tag.
Supports flexibility in UPC generation, allowing for barcode formats and customization.
