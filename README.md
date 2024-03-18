# DX Review 
## Current Release: v4.6.2
Simple VBA Code to (batch) Parse and Summarize Dr Checks XML Files with colorized formatting.

You can read the PDF of the docs folder or click the <a href="https://github.com/benstanfish/DX-Review/blob/main/docs/DX%20Review%20Documentation.pdf">DX Review Documentation.pdf</a> link.
![Initial_View](https://github.com/benstanfish/DX-Review/assets/34006582/b2af4bdb-7973-46e3-8079-f0786fe52868)
**Initial View of Example Output**

![Fully_Expanded_View](https://github.com/benstanfish/DX-Review/assets/34006582/887656cb-2c15-4d5c-956a-6d54bd9392e4)
**Expanded View of Example Output**

*Note:* to expand all the output, click the [2] or [+] at the top of the spreadsheet area to only expand a single region. Click the [1] or [+] to collapse all or selected regions, respectively.

## Download
There are currently ~~three~~ two implementations:

1. Download the <a href="https://github.com/benstanfish/DX-Review/blob/main/DX%20Review%20Plugin%20v2.xlsm">DX Review Plugin v2.xlsm</a> and save to your favorite location. This file comes equipped with a button in the first sheet that you can use to start the program. **Note:** The very first time you open this file, however, you will need to right-click the file and uncheck the *block* option at the bottom of the properties tab.
2. You can download the <a href="https://github.com/benstanfish/DX-Review/blob/main/dxreviewv2.bas">.bas file</a> and import as a module into an existing Excel Workbook. **Note:** This isn't really necessary to do more than once, as the code in the .bas file creates a new Workbook each time it processes XML files.
3. ~~Download the <a href="" onclick="javascript:void(0)">DX Review Plugin v2.xlam</a> and put it in your %APPDATA%/Roaming/Microsoft/Addins folder --- this can also be done (*more easily*) by simply "saving as" a copy of the .xlsm file above. The advantage of this .xlam file is that it you can enable it as an addin so that it is available in every Workbook - you can even add it to the Ribbon with a customized button of your liking.~~ To use the "addin" option, please download the *macro-enabled Workbook* and save-as an *macro-enabled Excel Addin (.xlam")*

You can download the <a href="https://github.com/benstanfish/DrChecksParser/blob/main/bsfdrchecksv2.bas">.bas file</a> and import into a macro-enabled Workbook, or simple download the *<a href="https://github.com/benstanfish/DrChecksParser/blob/main/BSF%20DrChecks%20Plugin%20v2.xlsm">BSF DrChecks Plugin v2.xlsm</a>* Workbook.

Note that you do have to set a reference to the dependencies (discussed in the troubleshooting section of the Documentation), if you import the .bas file.

If you go the import .bas route, you don't need to import the .bas code into each Excel file you work on - in fact, this code operates on files outside Workbook that hosts the module. 

## Quick Start Guide
The only methods visible to the user as a "macro" are the **DXReview_Select_File** and **DXReview_Select_Folder** methods, the latter is recommended in most cases.

1. Export full Dr Checks reports from ProjNet as XML - it's recommended to export the full report. Make sure you only use XML exports from ProjNet.
2. Save one or more of these XML reports in a folder - the names of the files doesn't really matter, just don't overwrite other XML reports.
3. Run the **DXReview_Select_Folder** macro and select the folder containing one or more XML reports.

The code will generate a new Workbook alongside the XML files, with a timestamp. It will include the summary of each XML report as a different tab in the Workbook.

The tabs in the summary Workbook use the <ReviewName> element in the XML file. So it's best to make sure you don't have duplicates of the same XML report - it's best to overwrite older versions of the same XML, or add a timestamp. The file names of the XML don't matter to this parser.

## Dependencies

The code requires that you have the *Microsoft Scripting Runtime (scrrun.dll)* and *Microsoft XML Library v6.0 (MSXML.dll)* --- these are fairly common installs on basically all Windows machines.

You can download these from the **<a href="https://github.com/benstanfish/DX-Review/tree/main/dependencies">dependencies folder</a>**. You'll need to register them in Windows, check the <a href="https://github.com/benstanfish/DX-Review/blob/main/docs/DX%20Review%20Documentation.pdf">Documentation</a> *Troubleshooting section* for step-by-step instructions.
