# DrChecksParser
Simple VBA Code to (batch) Parse and Summarize Dr Checks XML Files with colorized formating.

## Quick Start Guide
You can download the <a href="https://github.com/benstanfish/DrChecksParser/blob/main/bsfdrchecksv2.bas">.bas file</a> and import into a macro-enabled Workbook, or simple download the *<a href="https://github.com/benstanfish/DrChecksParser/blob/main/BSF%20DrChecks%20Plugin%20v2.xlsm">BSF DrChecks Plugin v2.xlsm</a>* Workbook. You don't need to import the .bas code into each Excel file you work on - in fact, this code operates on files outside Workbook that hosts the module. 

The only visible method visible as a "macro" is the MAIN() method, which is merely a wrapper function.

1. Export full Dr Checks reports from ProjNet as XML - it's recommended to export the full report. Make sure you only use XML exports from ProjNet.
2. Save one or more of these XML reports in a folder - the names of the files doesn't really matter, just don't overwrite other XML reports.
3. **Run the MAIN( ) macro** and select the folder containing one or more XML reports.

The code will generate a new Workbook alongside the XML files, with a timestamp. It will include the summary of each XML report as a different tab in the Workbook.

The tabs in the summary Workbook use the <ReviewName> element in the XML file. So it's best to make sure you don't have duplicates of the same XML report - it's best to overwrite older versions of the same XML, or add a timestamp. The file names of the XML don't matter to this parser.

### Documentation

I am currently working on more robust documentation.
