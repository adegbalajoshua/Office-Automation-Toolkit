# Splitbook VBA Macro

A VBA macro for Microsoft Excel that exports each worksheet in a workbook as a separate PDF file.

## Overview
The `Splitbook` macro automates the conversion of each worksheet in an Excel workbook into an individual PDF file. PDFs are formatted in landscape orientation, fitted to one page wide, and saved in the workbook's directory with the worksheet's name as the filename. This is ideal for generating reports, archiving worksheets, or sharing individual sheets as PDFs.

## Features
1. Exports each worksheet as a separate PDF.
2. Sets landscape orientation and fits columns to one page wide.
3. Saves PDFs in the workbook's directory, named after each worksheet.
4. Fast execution with suppressed screen updates and alerts.

## Prerequisites
1. Microsoft Excel (2010 or later) with VBA support.
2. Workbook saved as macro-enabled (`.xlsm`).
3. Write permissions in the workbook's directory.

## Quick Start
1. **Download the Macro:**
   - Clone the repository:
     
     ```bash
     git clone https://github.com/adegbalajoshua/Office-Automation-Tools/Splitbook-VBA-Macro.git
     ```
   - Or download [src/splitbook.vb](src/splitbook.vb)  
2. **Add to Excel:**
   - Open your workbook and save as `.xlsm`.
   - Press `Alt + F11` to open the VBA Editor.
   - Insert a new module (`Insert > Module`).
   - Copy and paste the code from [src/splitbook.vb](src/splitbook.vb).
   - Save the workbook.
     
4. **Run the Macro:**
   - Enable macros in Excel (`File > Options > Trust Center > Macro Settings`).
   - Press `Alt + F8`, select `splitbook`, and click `Run`.
   - Find PDFs in the workbook’s directory, named after each worksheet.  
For detailed instructions, see [user-guide.md](docs/user-guide.md).

## Repository Structure
1. [src/](src/): Contains the `splitbook.vb` macro code.
2. [docs/](docs/): Detailed documentation, including [user-guide.md](docs/user-guide.md).
3. [examples](examples/): Sample workbook ([sample-workbook.xlsx](examples/sample-workbook.xlsx)) for testing.
4. [tests/](tests/): Test workbook ([test-workbook.xlsm](tests/test-workbook.xlsm)) with the macro embedded.

## Contact
For questions, issues, or suggestions:
- Open an issue on the [GitHub repository](https://github.com/adegbalajoshua/Office-Automation-Toolkit/Splitbook-VBA-Macro.git)
- Contact [me](https://www.github.com/adegbalajoshua) on GitHub.

For advanced usage, troubleshooting, and FAQs, refer to [user-guide.md](docs/user-guide.md).











