# **COMIGHT Assistant 365**

# Introduction

COMIGHT Assistant is a productivity tool offering a wide range of functions to streamline office work and automate repetitive tasks, including batch processing of Excel worksheets, formatting Word documents, and managing files and folders. It aims to help office professionals to boost their productivity and efficiency.

Key features include merging and splitting Excel worksheets, formatting Word documents to official standards, and converting file types. The tool also provides utilities like creating file lists, batch creating folders, and batch creating place cards.

# System Requirements and Usage

- Requires: (1) Windows 10 x64 or later; (2) MS Office 2016 or later, or Microsoft 365; (3) .Net 9.0 SDK or later.

- Extract all files into one folder and double-click COMIGHT.EXE to run. 

- Select a function from the menu at the top of the main window. Follow the on-screen prompts to select a sub-function and files or folders to be processed in the dialog box, and enter the necessary parameters.
 
- Newly created files will be saved in the folder as designated in the Settings.

# Functions

## Start

### **Open Saving Folder** 

- Opens the folder where output files are saved.

### **Settings** 
- Sets parameters for this application.

### **Help** 

- Opens this user manual.

### **Exit** 

- Closes all windows and exits the program.

## Table

### **Batch Process Excel Worksheets**

- **Merge Records:** Vertically merges records from multiple worksheets with the same header to create a summary sheet.

  - Example: Several personnel roster worksheets with the same header are combined into a summary personnel roster worksheet.

- **Accumulate Values:** Sums the values of cells in specified ranges across multiple worksheets to create a summary sheet.

  - Example: Several worksheets are summarized by summing the values in the B2 to C3 range (including cells B2, B3, C2, C3) of each sheet. The resulting summary sheet's B2 cell contains the sum of all B2 cells across the worksheets, B3 contains the sum of all B3 cells, and so on.

- **Extract Cell Data:** Extracts data from specified cell ranges across multiple worksheets to create a summary sheet.

  - Example: Data from the B2 to C3 range (including cells B2, B3, C2, C3) of several worksheets is extracted to a summary sheet. Each row in the summary sheet corresponds to a worksheet, listing the workbook file name, worksheet name, and the values of cells B2, B3, C2, and C3.

- **Convert Textual Numbers into Numberic:** Converts textual numbers in specified cell ranges across multiple worksheets into numberic.

- **Adjust Worksheet Format for Printing:** Automatically adjusts borders, fonts, line breaks, column widths, row heights, and page layout based on the worksheet content for optimal printing.

### **Batch Disassemble Excel Workbook**

- **Split by a Column into Workbooks:** Splits worksheets into separate workbooks based on the values in a specified column. The worksheets must have the same structure.

  - Example: A student roster worksheet contains class information in column B (Class 1, Class 2, Class 3). Splitting by column B creates separate roster worksheets for each class, each stored in a different workbook.

- **Split by a Column into Worksheets:** Same as above, but the split data are stored in different worksheets within a new workbook. 

- **Disassemble by Worksheets:** Disperse the worksheets of a workbook into separate workbooks.
  
### **Batch Unhide Excel Worksheets**

- Make all hidden worksheets visible in all selected workbooks.

### **Batch Extract Tables From Word**

- Extracts tables from Word documents and imports them into Excel workbooks.

## Document

### **Convert Markdown into Word**

- Converts Markdown text into Word documents, keeping formats and styles as the original. If there are tables in the document, the tables will be extracted into Excel worksheets in the meantime.

### **Merge Data Into Document**

- Merges multiple Word documents and Excel spreadsheets into one document file (txt, pdf) for easy uploading to AI chat clients.

### **Batch Format Word Documents**

- Formats Word documents into official styles and adds outline structures.

- Requirements of the source documents: 
 
  - The title shall be at the beginning of the document (If there are multiple articles in the same document, each article must start on a new page, separated from the previous article by a page break or section break). 

  - The body text shall be separated from the title by at least one blank line. 

  - Heading numbers shall be presented in the formats as follows: 

  | Heading Level | Chinese Heading Number Format |
  |---|---|
  | 0 | 第一部分 / 第二部分 / 第三部分 ... <br>（“部分” can be replaced with “篇”“章”“节”）| 
  | 1 | 一、 / 二、 / 三、... |
  | 2 | （一） / （二） / （三） ... |
  | 3 | 1. / 2. / 3. ...|
  | 4 | (1) / (2) / (3) ... |

  - Tables shall be separated from its following text by at least one blank line.

  - The signature shall be separated from the body text by at least one blank line, with the organization's/person's name above (can be multiple, arranged vertically), and the date below. The date format shall be "YYYY年MM月DD日".

  - **Documents with columns, complex tables, or mixed text and images are not applicable!**

### **Batch Repair Word Documents**

- Fixes the problematic Word Documents whose styles and formats cannot be adjusted correctly.

## Tools

### **Batch Convert Office File Types**

- Converts older Excel (.xls), Word (.doc), WPS Spreadsheet (.et), and WPS Document (.wps) files into .xlsx and .docx formats.

### **Create File List**

- Creates a list of all subfolders and files within a specified folder path, including hyperlinks, subpaths, file/folder names, types, and creation times.

### **Batch Create Folders**

- Creates folders based on the folder structure data in an Excel worksheet. Folders are created hierarchically from left to right, with the leftmost folder being the highest level (closest to the root directory) and the rightmost folder being the lowest level (furthest from the root directory). Use the template workbook provided with this application to organize the folder structure data.

### **Batch Create Place Cards**

- Creates place cards (20 x 10 cm) based on the roster data in an Excel worksheet. One item can be split into two lines within a cell. Use the template workbook provided with this application to organize the roster data.

### Remove Markdown Marks in Copied Text

- Removes all Markdown marks in the text copied to the clipboard, and then write the cleaned text back to the clipboard.

### **Create QR Code**

- Creates the QR code from the input text. The text and code are all processed locally to prevent leakage onto the internet.