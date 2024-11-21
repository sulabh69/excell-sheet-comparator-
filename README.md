Here’s a `README.md` file that explains how to use the PDF-to-Excel converter and Excel sheet comparator, along with the dependencies you need to install to run the scripts.

---

# PDF to Excel Converter and Excel Sheet Comparator

This repository contains two Python scripts with graphical user interfaces (GUIs) built using `Tkinter`. The tools provided in this repository allow users to:

1. Convert a PDF file containing tables into an Excel file.
2. Compare two Excel files to identify differences between them.

Both tools are easy to use with drag-and-drop functionality, and they allow you to customize the input and output locations.

 Features

 1. PDF to Excel Converter
- Drag and Drop: Allows you to drag and drop a PDF file into the GUI for conversion.
- Excel Export: Converts the tables in the PDF to an Excel sheet.
- Multiple Tables: Handles multiple tables in a PDF, saving each table to a different sheet in the output Excel file.

2. Excel Sheet Comparator
- **Compare Two Excel Files**: Compares two Excel files to find differences in their contents.
- **Error Highlighting**: Cells with differences are highlighted in red, and a comment is added with details of the difference.
- **Generate Report**: The result is saved in a new Excel file, showing the differences and highlighting cells where mismatches occur.

Installation Requirements

Before running the scripts, you need to install a few Python libraries. You can install the required dependencies using `pip`.
Dependencies
- `Tkinter`: A Python library for GUI development.
- `TkinterDnD2`: A library for enabling drag-and-drop functionality in Tkinter.
- `pandas`: A library used for data manipulation and analysis.
- `openpyxl`: A library for reading and writing Excel files.
- `tabula-py`: A Python wrapper for Tabula, a tool for extracting tables from PDFs.

 To install all the required dependencies, run the following command:

```bash
pip install pandas openpyxl tabula-py tk tkinterdnd2
```
 How to Use

PDF to Excel Converter

1. **Launch the application**:
   Run the `PDF_to_Excel_Converter.py` script.
   
   ```bash
   python PDF_to_Excel_Converter.py
   ```

2. Select or Drag & Drop the PDF File:
   - You can either browse for a PDF file using the **Browse** button or drag and drop the PDF into the designated area.
   
3. Set the Output Location:
   - Click on the **Set Output Location** button to specify where the Excel file should be saved.

4. Convert to Excel:
   - After selecting the PDF and setting the output location, click on the **Convert to Excel** button to start the conversion. The program will extract tables from the PDF and save them into an Excel file.

 Excel Sheet Comparator

1. Launch the application:
   Run the `Excel_Sheet_Comparator.py` script.
   
   ```bash
   python Excel_Sheet_Comparator.py
   ```

2. Upload or Drag & Drop Excel Files:
   - You can either browse for the first and second Excel files using the **Upload File** buttons or drag and drop the files into the respective areas.

3. Clear Files:
   - If you need to clear the files, click on the Clear button for each side.

4. Compare the Excel Sheets**:
   - After selecting both Excel files, click on **Compare Excel Sheets**. The program will compare the two files and highlight the differences in a new Excel file.

5. View the Results**:
   - The comparison result will be saved in a new Excel file, where the differences are marked with "Error" and highlighted in red. Comments are added to the cells to show the original and compared values.

 Example of Expected Output

 PDF to Excel Converter:
The output will be an Excel file with multiple sheets, each containing a table extracted from the PDF.

 Excel Sheet Comparator:
The output will be an Excel file with:
- A "Comparison" sheet showing the differences.
- Cells with differences will be highlighted in red, and comments will provide details about the differences.

 Troubleshooting

- If you encounter any issues with `tabula-py`, make sure that Java is installed on your system, as Tabula requires Java to function.
- If any required libraries are missing, you can install them using `pip install <library-name>`.
- Ensure that the paths to the files are correct, especially when setting the output location or loading the input files.



---

 Additional Notes:
1. Java Requirement for Tabula**: `tabula-py` depends on Tabula, which in turn requires Java to be installed on your machine. Make sure Java is installed and properly configured.
2. GUI Responsiveness**: The application uses `TkinterDnD2` for drag-and-drop functionality. Ensure that the library is properly installed to enable this feature.


