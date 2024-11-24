# COMIGHT Assistant

## Introduction

COMIGHT Assistant is a productivity tool offering a wide range of functions to streamline office work and automate repetitive tasks, including batch processing of Excel worksheets, formatting Word documents, and managing files and folders. It aims to help office professionals to boost their productivity and efficiency.

Key features include merging and splitting Excel worksheets, comparing data, formatting Word documents to official standards, and converting file types. The tool also provides utilities like creating name cards, making file lists, and batch folder creation. Additionally, a unique built-in browser allows easy copying of text from web pages without manual selection. 

## System Requirements and Usage

- Requires: (1) Windows 10 x64 or later, (2) MS Office 2016 or later, or Microsoft 365, (3) WebView2 Runtime, (4) .Net 8.0 SDK 

- Extract all files into one folder and double-click COMIGHT.EXE to run. 

- Select a function from the menu at the top of the main window. Follow the on-screen prompts to select the files or folders to be processed in the dialog box, and enter the necessary parameters. The generated files and folders are located in the same folder as the source files/folder, or in the "COMIGHT Files" folder on the Desktop.

## Function Introduction

### Start

- **Help:** Opens this user manual.

- **Exit:** Closes all windows and exits the program.

### Table

#### **Batch Unhide Excel Worksheets**

- Make all hidden worksheets visible in all selected workbooks.

#### **Batch Process Excel Worksheets**

- **Merge Records:** Vertically merges records from multiple worksheets with the same header to create a summary sheet.

  - Example: Several personnel roster worksheets with the same header are combined into a summary personnel roster worksheet.

- **Accumulate Values:** Sums the values of cells in specified ranges across multiple worksheets to create a summary sheet.

  - Example: Several worksheets are summarized by summing the values in the B2 to C3 range (including cells B2, B3, C2, C3) of each sheet. The resulting summary sheet's B2 cell contains the sum of all B2 cells across the worksheets, B3 contains the sum of all B3 cells, and so on.

- **Extract Cell Data:** Extracts data from specified cell ranges across multiple worksheets to create a summary sheet.

  - Example: Data from the B2 to C3 range (including cells B2, B3, C2, C3) of several worksheets is extracted to a summary sheet. Each row in the summary sheet corresponds to a worksheet, listing the workbook file name, worksheet name, and the values of cells B2, B3, C2, and C3.

- **Convert Textual Numbers into Numberic:** Converts textual numbers in specified cell ranges across multiple worksheets into numberic.

  - This function generates a summary sheet listing cells that could not be converted, along with their workbook file name, worksheet name, cell addresses, and values.

- **Copy Formulas to Multiple Worksheets:** Copies a formulas from a specified range in a template worksheet to multiple target worksheets. The formulas to be copied must be located in the first worksheet (template worksheet) of the template workbook.

  - Example: A template worksheet has the sum formula "=SUM(A2:D2)" in cell E2, "=SUM(A3:D3)" in E3, and "=SUM(A4:D4)" in E4. This function copies the formulas from the E2:E4 range to multiple target worksheets.

- **Prefix Workbook Filenames with Cell Data:** Extracts data from specified cell ranges across multiple worksheets and uses it as a prefix for the workbook filenames.

  - Example: Several departmental roster worksheets are stored in workbooks named "Roster1.xlsx", "Roster2.xlsx", "Roster3.xlsx". Cells A1 and A2 of each roster worksheet contain the department name and update time, respectively. This function extracts the values from A1 and A2 and adds them as prefixes to the file names, renaming them to "[Department Name] [Update Time]_Roster1.xlsx" (e.g. "R&D Department 2023-09_Roster2.xlsx", "Finance Department 2023-10_Roster3.xlsx").

- **Adjust Worksheet Format for Printing:** Automatically adjusts borders, fonts, line breaks, column widths, row heights, and page layout based on the worksheet content for optimal printing.

#### **Split Excel Worksheet**

- **Split into Workbooks:** Splits a worksheet into separate workbooks based on the values in a specified column. The data to be split must be located in the first worksheet.

  - Example: A student roster worksheet contains class information in column B (Class 1, Class 2, Class 3). Splitting by column B creates separate roster worksheets for each class, each stored in a different workbook.

- **Split into Worksheets:** Same as above, but the split worksheets are stored in different worksheets within a new workbook. The worksheet names are derived from the values in the split column. The number of worksheets cannot exceed 255.

#### **Compare Excel Worksheets**

- Compares data with identical record keys and column fields between a starting data worksheet and an ending data worksheet, listing the differences (and calculating the percentage change for numerical data). The starting and ending data worksheets must be in different workbooks and be the first worksheet in their respective workbooks. Records must be unique (no duplicate record keys).

#### **Screen Stocks**

- Screens stocks based on price-to-book (P/B) and price-to-earnings (P/E) ratios, selecting undervalued stocks with a margin of safety. The stock data Excel worksheet can be exported from stock analysis software. The data worksheet must be the first worksheet in the workbook, with a single header row containing the column field names: Stock Code, Name, Sector, Current Price, P/E Ratio, and P/B Ratio. The filtered results are stored in the second worksheet.

### Document

#### **Batch Format Word Documents**

- Formats Word documents into official styles and adds outline structures.

- **Requirements of the source documents:** 
 
  - The title shall be at the beginning of the document (If there are multiple articles in the same document, each article must start on a new page, separated from the previous article by a page break or section break). 

  - The body text is separated from the title by at least one blank line. 

  - Heading numbers shall be presented in the formats as follows: 

  | Heading Level | Chinese Heading Number Format | English Heading Number Format |
  |---|---|---|
  | 0 | 第一部分 第二部分 第三部分 ... <br>（“部分” can be replaced with “篇”“章”“节”）| N.A. |
  | 1 | 一、 二、 三、... | A. B. C. or 1. 2. 3. ... (with a space behind) |
  | 2 | （一） （二） （三） ... | A.1 B.1 C.1 or 1.1 2.1 3.1 ... (with a space behind) |
  | 3 | 1. 2. 3. ...| A.1.1 B.1.1 C.1.1 or 1.1.1 2.1.1 3.1.1 ... (with a space behind) |
  | 4 | (1) (2) (3) ... | A.1.1.1 B.1.1.1 C.1.1.1 or 1.1.1.1 2.1.1.1 3.1.1.1 ... (with a space behind) |

  - The signature shall be separated from the body text by at least one blank line, with the organization's/person's name above (can be multiple, arranged vertically), and the date below. For Chinese documents, the date format shall be "YYYY年MM月DD日".

  - **Documents with columns, complex tables, or mixed text and images are not applicable!**

#### **Export Document Table into Word Document (Only works for Chinese Documents)**

- Exports the contents of a document table into a Word document, automatically numbers headings at all levels, and formats them according to Chinese government document standards.

- Requirements of the document table:

  - A document table template is provided with this program, with instructions for filling it in. The "Text" column in the "Title" worksheet and the "Heading Level" and "Text" columns in the "Body" worksheet are mandatory and serve as the source of the output document content; the other columns are optional for notes, filtering, etc. 

  - "Heading Level" shall be selected from the dropdown box, within the options of "0级", "1级", "2级", "3级", "4级", "是", "条", respectively. In the exported document, heading numbers will be presented in the formats as follows: 

    | Heading Level | Heading Number Format |
    |---|---| 
    | 0级 | 第一部分 第二部分 第三部分... |
    | 1级 | 一、 二、 三、... |
    | 2级 | （一） （二） （三） ... |
    | 3级 | 1. 2. 3. ...|
    | 4级 | (1) (2) (3) ... |
    | 是 | 一是  二是  三是 ... |
    | 条 | 第一条  第二条  第三条 ... |

  - **None of the 3 worksheets can be deleted, moved, or have their column structure changed. Any rows that are hidden or filtered out will not be exported.**

#### **Import Text into Document Table (Only works for Chinese Documents)**

- Imports the content of the text box in the dialog box into a document table, saved on the Windows desktop.

### Tools

#### **Merge Documents and Tables**

- Converts selected Word documents and Excel spreadsheets to plain text, and then merges them into a single Word document, saved on the Windows desktop.

#### **Batch Convert Office File Types**

- Batch converts older Excel (.xls), Word (.doc), WPS Spreadsheet (.et), and WPS Document (.wps) files into the current .xlsx and .docx formats.

#### **Make File List**

- Creates a list of all subfolders and files within a specified folder path, including hyperlinks, subpaths, file/folder names, types, and creation times.

#### **Make Folders**

- Creates folders based on the folder structure data in an Excel worksheet. Folders are created hierarchically from left to right, with the leftmost folder being the highest level (closest to the root directory) and the rightmost folder being the lowest level (furthest from the root directory). Use the template provided by this program for the folder creation worksheet.

#### **Create Name Cards**

- Creates name cards (20 x 10 cm) based on the data in a roster Excel worksheet. The names for the seat plates must be in column A of the first worksheet, starting from cell A1. One item can be split into two lines within a cell. Each name card can accommodate a maximum of 10 Chinese characters or 25 English characters.

### Browser

- This browser allows you to directly copy entire paragraphs of text from most web pages without having to drag the mouse and hold the Shift key to select and copy. Copied text does not contain any markdown symbols or HTML tags.

- Move the mouse pointer over the text. When a light green border appears around the text block, double-click the mouse. The border will briefly turn pink, indicating that the text has been copied. The Database.xlsx file in the program's folder contains website addresses that are automatically loaded into the browser's favorites dropdown menu upon startup.  You can edit this file as needed (the website list must be in the "Websites" worksheet, and the header cannot be changed).

