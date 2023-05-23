# CSV Formatter

This is a python project done in spare time.  The goal is to read in data from a CSV file, clean the data, and format in a specific way to allow for budget analysis of property flips.

## Usage
Run using ```python -m main.py```.

There are no command line arguments or arguments.

Selct the CSV using the Select File button.  Currently the expected format of the csv format is hard coded.  Cleaning, fuzzy matching and re-formatting of data is done on initial load of the file.  

Save the file to an excel file by selecting the Save As button.  Data is split out to different sheets as part of the excel writing function 

## Distribution
An executable file for Windows can be found under /dist folder
