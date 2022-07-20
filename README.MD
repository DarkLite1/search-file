# Description
A PowerShell script designed to search for files. 

This script will search for files where the file name contains a specific string. It wil send out an e-mail with an Excel file in attachment with matching files.

Examples:
- *.pst  : Search for all files with the extension 'pst'
- *kiwi* : Search for all files with the string 'kiwi' in the file name

 All required parameters are read from a [.JSON file](Example.json). The input file can be modified as needed. 

A scheduled task can be created to run the script in an automated way.

# Best practices
- Read the [help section](Search file.ps1) in the script for an explanation of the input file parameters.

# Requirements
All required modules can be found in the [repository](https://github.com/DarkLite1?tab=repositories) and are free to use.