# splitbook() User Guide

This guide provides detailed instructions for using the `Splitbook` VBA macro, which exports each worksheet in a Microsoft Excel workbook as a separate PDF file. 
It includes setup steps, usage examples, advanced configurations, and troubleshooting tips.

## Table of Contents
- [Overview](#overview)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Basic Usage](#basic-usage)
- [Advanced Usage](#advanced-usage)
  - [Skipping Hidden Worksheets](#skipping-hidden-worksheets)
  - [Customizing Output Directory](#customizing-output-directory)
  - [Handling Special Characters in Worksheet Names](#handling-special-characters-in-worksheet-names)
- [Troubleshooting](#troubleshooting)
- [FAQs](#faqs)
- [Contact](#contact)

## Overview
The `Splitbook` macro automates the process of converting each worksheet in an Excel workbook into an individual PDF file. 
Each PDF is:
1. Formatted in andscape orientation
2. Fitted to one page wide, with multiple pages tall if needed.
3. Saved in the same directory as the workbook, using the worksheet's name as the filename.

This is useful for creating separate reports, archiving worksheets, or sharing individual sheets as PDFs.

## Prerequisites
1. **Microsoft Excel**: Version supporting VBA (e.g., Excel 2010 or later).
2. **Macro-Enabled Workbook**: Save your workbook as .xlsm to enable macro functionality.
3. **Write Permissions**: Ensure you have write access to the directory where the workbook is saved, as PDFs will be created there.
4. **Basic VBA Knowledge**: Familiarity with the VBA Editor is helpful but not required.

## Installation
1. **Download the Macro:**
    - Clone or download the repository:
      ```bash
      git clone https://github.com/adegbalajoshua/Office-Automation-Toolkit/Splitbook-VBA-Macro.git
      ```
    - Alternatively, download `src/Splitbook.vb` from the repository.
2. **Open Excel:**
    - Open the workbook you want to split into PDFs.
    - Ensure it is saved as a macro-enabled workbook (`.xlsm`).
3. **Access the VBA Editor:**
    - Press `Alt + F11` to open the VBA Editor.
    - In the Project Explorer (left panel), right-click your workbook (e.g., `VBAProject (YourWorkbook.xlsm)`).
    - Select `Insert > Module` to create a new module.
4. **Add the Macro:**
    - Open `splitbook.vb` from the repository in a text editor, or copy the code below:
      ```vba
      Sub splitbook()
      Dim xPath As String
      Dim xWs As Worksheet
      xPath = Application.ActiveWorkbook.Path
      Application.ScreenUpdating = False
      Application.DisplayAlerts = False
      For Each xWs In ThisWorkbook.Sheets
          On Error Resume Next
          With xWs.PageSetup
              .Orientation = xlLandscape
              .Zoom = False
              .FitToPagesWide = 1
              .FitToPagesTall = False
          End With
          xWs.ExportAsFixedFormat _
              Type:=xlTypePDF, _
              Filename:=xPath & "\" & xWs.Name & ".pdf", _
              Quality:=xlQualityStandard, _
              IncludeDocProperties:=True, _
              IgnorePrintAreas:=False, _
              OpenAfterPublish:=False
          On Error GoTo 0
      Next
      Application.DisplayAlerts = True
      Application.ScreenUpdating = True
      End Sub
      ```
    - Copy and paste the code into the new module.
    - Save the workbook (`Ctrl + S` or `File > Save`).
5. **Enable Macros:**
    - Go to `File > Options > Trust Center > Trust Center Settings > Macro Settings`.
    - Select "Enable all macros" or "Enable VBA macros" (not recommended for untrusted sources).
    - Save changes and close the dialog.
 
## Basic Usage
1. **Prepare Your Workbook:**
    - Ensure your workbook has multiple worksheets with content you want to export as PDFs.
    - Save the workbook in a directory where you have write permissions.
2. **Run the Macro:**
    - Press ``Alt + F8`` to open the Macro dialog.
    - Select `splitbook` from the list and click `Run`.
    - The macro will:
      - Set each worksheet to landscape orientation.
      - Fit all columns to one page wide.
      - Export each worksheet as a PDF named after the worksheet (e.g., `workbookOne.pdf`, `workbookTwo.pdf`).
      - Save PDFs in the same directory as the workbook.
3. **Check Output:**
    - Navigate to the workbook’s directory to find the generated PDFs.
    - Each PDF corresponds to a worksheet, maintaining the content and formatting.
4. **Test with Sample Workbook:**
    - Use the `examples/sample-workbook.xlsx` file from the repository to test the macro.
    - Add the macro to this workbook (as a `.xlsm` file) and run it to see sample PDFs generated.

## Advanced Usage
### Skipping Hidden Worksheets
By default, `splitbook` exports all worksheets, including hidden ones. 
To skip hidden worksheets, modify the macro by adding a visibility check:
```vba
For Each xWs In ThisWorkbook.Sheets
    If xWs.Visible = xlSheetVisible Then ' Only process visible sheets
        On Error Resume Next
        With xWs.PageSetup
            .Orientation = xlLandscape
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = False
        End With
        xWs.ExportAsFixedFormat _
            Type:=xlTypePDF, _
            Filename:=xPath & "\" & xWs.Name & ".pdf", _
            Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, _
            IgnorePrintAreas:=False, _
            OpenAfterPublish:=False
        On Error GoTo 0
    End If
Next
```
Save the modified macro and run it as usual. Only visible worksheets will be exported.
### Customizing Output Directory
To save PDFs to a specific folder (e.g., `C:\Output\PDFs`), modify the `xPath` variable:
```vba
xPath = "C:\Output\PDFs" ' Replace with your desired path
```
Ensure the folder exists and you have write permissions. Create the folder manually or add code to create it:
```vba
If Dir(xPath, vbDirectory) = "" Then
    MkDir xPath
End If
```
Insert this code after defining `xPath` and before the `For Each` loop.
### Handling Special Characters in Worksheet Names
Worksheet names with special characters (e.g., `/`, `\`, `*`, `?`) may cause errors because they are invalid in filenames. To sanitize names, add a function to replace invalid characters:
```vba
Function SanitizeFileName(fileName As String) As String
    Dim invalidChars As String
    invalidChars = "\/:*?""<>|"
    Dim i As Integer
    For i = 1 To Len(invalidChars)
        fileName = Replace(fileName, Mid(invalidChars, i, 1), "_")
    Next i
    SanitizeFileName = fileName
End Function
```
Then modify the `ExportAsFixedFormat` line:
```vba
Filename:=xPath & "\" & SanitizeFileName(xWs.Name) & ".pdf", _
```
Add the `SanitizeFileName` function to the same module, and it will replace invalid characters with underscores (e.g., `Sheet/1` becomes `Sheet_1`).

## Troubleshooting
1. **No PDFs Generated:**
   - **Cause:** Workbook not saved or no write permissions in the directory.
   - **Solution:** Save the workbook (`File > Save As`) as `.xlsm` in a directory where you have write access.
2. **Macro not running:**
   - **Casue:** Macro disabled in Excel
   - **Solution:** Enable macros in `Trust Center Settings` (see [Installation](#installation)).
3. **Invalid Filename Errors:**
   - **Cause:** Worksheet names contain special characters (e.g., `*`, `/`).
   - **Solution:** Rename worksheets manually or use the sanitize function (see [Handling Special Characters in Worksheet Names](#handling-special-characters-in-worksheet-names)).
4. **PDFs Missing for Some Worksheets:**
   - **Casue:** Errors in worksheet content of formatting.
   - **Solution:** Check the VBA Editor for errors or test with `examples/sample-workbook.xlsx`. Ensure worksheets have valid content.
5. **Hidden Sheets Exported:**
   - **Cause:** Macro processes all worksheets by default.
   - **Solution:** Modify the macro to skip hidden sheets (see [Skipping Hidden Worksheets](#skipping-hidden-worksheets)).
  
## FAQs
**Q: Can I change the PDF page orientation to portrait?**  
A: Yes, modify `.Orientation = xlLandscape` to `.Orientation = xlPortrait` in the macro.  

**Q: How do I open PDFs automatically after creation?**  
A: Change `OpenAfterPublish:=False` to `OpenAfterPublish:=True` in the `ExportAsFixedFormat` line.  

**Q: Can I use this with Excel for Mac?**  
A: Yes, but ensure your Excel version supports VBA. Some path-related issues may occur due to macOS file system differences (e.g., use `":"` instead of `"\"` for paths).  

**Q: How do I test the macro without affecting my workbook?**  
A: Use the `examples/sample-workbook.xlsx` or `tests/test-workbook.xlsm` provided in the repository.

## Contact
For questions, issues, or suggestions:
- Open an issue on the [GitHub repository](https://github.com/adegbalajoshua/Office-Automation-Toolkit/Splitbook-VBA-Macro.git)
- Contact [me](https://www.github.com/adegbalajoshua) on GitHub.











