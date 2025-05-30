# COMIGHT Assistant

## Introduction

COMIGHT Assistant is a productivity tool offering a wide range of functions to streamline office work and automate repetitive tasks, including batch processing of Excel worksheets, formatting Word documents, and managing files and folders. It aims to help office professionals to boost their productivity and efficiency.

Key features include merging and splitting Excel worksheets, comparing data, formatting Word documents to official standards, and converting file types. The tool also provides utilities like creating name cards, making file lists, and batch folder creation. Additionally, a unique built-in browser allows easy copying of text from web pages without manual selection. 

## System Requirements and Usage

- Requires: (1) Windows 10 x64 or later; (2) MS Office 2016 or later, or Microsoft 365; (3) WebView2 Runtime; (4) .Net 8.0 SDK 

- Extract all files into one folder and double-click COMIGHT.EXE to run. 

- Select a function from the menu at the top of the main window. Follow the on-screen prompts to select a sub-function and files or folders to be processed in the dialog box, and enter the necessary parameters.
 
- Newly created files will be saved in the folder as designated in the Settings.

## Function Introduction

### Start

#### **Open Saving Folder** 

- Opens the folder where output files are saved.

#### **Settings** 
- Sets parameters for this application.

#### **System Info**

- Shows the system infomation of the current computer. The system information can be exported to an Excel workbook.

#### **Help** 

- Opens this user manual.

#### **Exit** 

- Closes all windows and exits the program.

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

- **Copy Formulas to Multiple Worksheets:** Copies a formulas from a specified range in a template worksheet to multiple target worksheets. The formulas to be copied must be located in the first worksheet (template worksheet) of the template workbook.

  - Example: A template worksheet has the sum formula "=SUM(A2:D2)" in cell E2, "=SUM(A3:D3)" in E3, and "=SUM(A4:D4)" in E4. This function copies the formulas from the E2:E4 range to multiple target worksheets.

- **Adjust Worksheet Format for Printing:** Automatically adjusts borders, fonts, line breaks, column widths, row heights, and page layout based on the worksheet content for optimal printing.

#### **Split Excel Worksheet**

- **Split into Workbooks:** Splits worksheets into separate workbooks based on the values in a specified column. The worksheets must have the same structure.

  - Example: A student roster worksheet contains class information in column B (Class 1, Class 2, Class 3). Splitting by column B creates separate roster worksheets for each class, each stored in a different workbook.

- **Split into Worksheets:** Same as above, but the split worksheets are stored in different worksheets within a new workbook. The worksheet names are derived from the values in the split column.

#### **Compare Excel Worksheets**

- Compares data with identical record keys and column fields between starting data worksheets and ending data worksheets, listing the differences (and calculating the percentage change for numerical data). The starting and ending data worksheets must have the same structure, and be arranged in the same sequence in two workbooks. Records must be unique (no duplicate record keys).

### Document

#### **Batch Format Word Documents**

- Formats Word documents into official styles and adds outline structures.

- Requirements of the source documents: 
 
  - The title shall be at the beginning of the document (If there are multiple articles in the same document, each article must start on a new page, separated from the previous article by a page break or section break). 

  - The body text shall be separated from the title by at least one blank line. 

  - Heading numbers shall be presented in the formats as follows: 

  | Heading Level | Chinese Heading Number Format | English Heading Number Format |
  |---|---|---|
  | 0 | 第一部分 / 第二部分 / 第三部分 ... <br>（“部分” can be replaced with “篇”“章”“节”）| Part 1 / Part 2 / Part 3 ... <br> ("Part" can be replaced with "Chapter" "Section" |
  | 1 | 一、 / 二、 / 三、... | 1. / 2. / 3. ... (with a space behind) |
  | 2 | （一） / （二） / （三） ... | 1.1 / 2.1 / 3.1 ... (with a space behind) |
  | 3 | 1. / 2. / 3. ...| 1.1.1 / 2.1.1 / 3.1.1 ... (with a space behind) |
  | 4 | (1) / (2) / (3) ... | 1.1.1.1 / 2.1.1.1 / 3.1.1.1 ... (with a space behind) |

  - For English documents, blank lines shall be placed between different levels of headings, between a heading and its body paragraph, and between paragraphs.

  - Tables shall be separated from its following text by at least one blank line.

  - The signature shall be separated from the body text by at least one blank line, with the organization's/person's name above (can be multiple, arranged vertically), and the date below. For Chinese documents, the date format shall be "YYYY年MM月DD日".

  - **Documents with columns, complex tables, or mixed text and images are not applicable!**

#### **Convert Markdown into Word**

- Converts Markdown text into Word documents, keeping formats and styles as the original. If there are tables in the document, the tables will be extracted into Excel worksheets in the meantime.

  - **The Pandoc application (https://github.com/jgm/pandoc) is necessary for this function. The path of Pandoc executable file shall be set correctly in the Settings.**
 
#### **Export Document Table into Word Document**

- Exports the contents of a document table into a Word document, automatically numbers headings at all levels, and formats them in conformity with official document standards.

- Requirements of the document table:

  - A document table template is provided with this program, with instructions for filling it in. The "Text" column in the "Title" worksheet and the "Heading Level" and "Text" columns in the "Body" worksheet are mandatory and serve as the source of the output document content; the other columns are optional for notes, filtering, etc. 

  - "Heading Level" shall be selected from the dropdown box, within the options of "Lv0", "Lv1", "Lv2", "Lv3", "Lv4", "Enum.", "Itm.", "Immed.", respectively. In the exported document, heading numbers will be presented in the formats as follows: 

    | Heading Level | Chinese Heading Number Format | English Heading Number Format |
    |---|---|---|
    | Lv0 | 第一部分 / 第二部分 / 第三部分 ... | Part 1 / Part 2 / Part 3 ... |
    | Lv1 | 一、 / 二、 / 三、 ... | 1. / 2. / 3. ... |
    | Lv2 | （一） / （二） / （三） ... | 1.1 / 2.1 / 3.1 ... |
    | Lv3 | 1. / 2. / 3. ...| 1.1.1 / 2.1.1 / 3.1.1 ... |
    | Lv4 | (1) / (2) / (3) ... | 1.1.1.1 / 2.1.1.1 / 3.1.1.1 ... |
    | Enum. | 一是 / 二是 / 三是 ... | N.A. |
    | Itm. | 第一条 / 第二条 / 第三条 ... | N.A. |

  - The "Immed." stands for "Immediately following the above paragraph", which means the corresponding text will be deemed as a part of the above paragraph (as a whole, instead of a new paragraph).
  
  - **None of the 2 worksheets can be renamed, deleted, or have their column structure changed. Any rows that are hidden or filtered out will not be exported.**

#### **Import Text into Document Table**

- Imports the content of the text box in the dialog box into a document table.

#### **Merge Data Into Document**

- Merges multiple Word documents and Excel spreadsheets into one document file (txt, pdf) for easy uploading to AI chat clients.

### Tools

#### **Batch Convert Office File Types**

- Batch converts older Excel (.xls), Word (.doc), WPS Spreadsheet (.et), and WPS Document (.wps) files into the current .xlsx and .docx formats.

#### **Batch Repair Word Documents**

- Repairs problematic Word documents that show weird styles and cannot be formatted correctly.

#### **Create File List**

- Creates a list of all subfolders and files within a specified folder path, including hyperlinks, subpaths, file/folder names, types, and creation times.

#### **Create Folders**

- Creates folders based on the folder structure data in an Excel worksheet. Folders are created hierarchically from left to right, with the leftmost folder being the highest level (closest to the root directory) and the rightmost folder being the lowest level (furthest from the root directory). Use the template workbook provided with this application to organize the folder structure data.

#### **Create Place Cards**

- Creates place cards (20 x 10 cm) based on the roster data in an Excel worksheet. One item can be split into two lines within a cell. Use the template workbook provided with this application to organize the roster data.

### Web

#### **Browser**

- This browser allows you to directly copy entire paragraphs of text from most web pages without having to drag the mouse and hold the Shift key to select and copy. Copied text does not contain any markdown symbols or HTML tags, and is free of blank lines or white spaces before or after each paragraph.

- Move the mouse pointer over the text. When a light green border appears around the text block, double-click the mouse. The border will briefly turn pink, indicating that the text has been copied. The "Websites.json" file in the program's folder contains website addresses that are automatically loaded into the browser's dropdown menu upon startup. You can edit this file as needed.
