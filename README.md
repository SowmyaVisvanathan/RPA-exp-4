### Name : Sowmya V
### Reg no : 212222110045
### Date : 
# Experiment 4 - Read and Write Excel Data

## Aim:
To read data from one Excel file and write it into another Excel file using UiPath.

## Software used:
UiPath Studio community edition

## Procedure:

#### 1. Create a New Project in UiPath Studio
- Step 1: Open UiPath Studio Community Edition.
- Step 2: Select "Process" under the New Project section.
- Step 3: Name the project, e.g., "ExcelDataTransferAutomation".
- Step 4: Choose a save location and add a description if necessary.
- Step 5: Click Create to initialize the project.

#### 2. Add an Excel Application Scope to Read Data
- Step 1: In the Activities Panel, search for "Excel Application Scope".
- Step 2: Drag and drop the Excel Application Scope activity onto the design canvas.
- Step 3: Set the WorkbookPath property to the path of the source Excel file, e.g., "C:\Users\sowmy\OneDrive\Desktop\rpa\Book1.xlsx".
- Step 4: Inside the Do section, add a Read Range activity:
- Sheet Name: "Sheet1"

#### 3. Add an Excel Application Scope to Write Data
- Step 1: Add another Excel Application Scope activity below the first one.
- Step 2: Set the WorkbookPath property to the path of the target Excel file, e.g., "C:\Users\sowmy\OneDrive\Desktop\rpa\Book2.xlsx".
- Step 3: Inside the Do section, add a Write Range activity:
- Sheet Name: "Sheet1"
- Starting Cell: "A1"


## Main.xaml:

![image](https://github.com/user-attachments/assets/432b3ab0-b6f4-4842-91de-cd9f77505f1f)


## Output:

![image](https://github.com/user-attachments/assets/ad87f81f-d25b-4e47-ade1-14c0040739c5)


## Result:
Data from one Excel file is successfully read and written into another Excel file using UiPath.


