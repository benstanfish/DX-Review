# DrChecksParser
Simple VBA Code to (batch) Parse and Summarize Dr Checks XML Files with colorized formating.

## Quick Start Guide
You can download the .bas file and import into a macro-enabled Workbook, or simple download the *<a href="https://github.com/benstanfish/DrChecksParser/blob/main/BSF%20DrChecks%20Plugin%20v2.xlsm">BSF DrChecks Plugin v2.xlsm</a>* Workbook. The only visible method is the MAIN() method, which is merely a wrapper function.

1. Export full Dr Checks reports from ProjNet as XML - it's recommended to export the full report. Make sure you only use XML exports from ProjNet.
2. Save one or more of these XML reports in a folder - the names of the files doesn't really matter, just don't overwrite other XML reports.
3.  Run the MAIN() function adn select the folder

The code will generate a new Workbook alongside the XML files, with a timestamp. It will include the summary of each XML report as a different tab in the Workbook. You don't necessarily need to import the .bas code into each File - in fact, this code operates on files outside Workbook that hosts the module.

### Documentation

I am currently working on more robust documentation.
