# excel-tools
A handful of tools for working with Excel files. Before proceeding, please note that **all** the tools in this repo require Pandas to access the data in the Excel files. If you don't have Pandas installed, you can get it by running `python3 -m pip install pandas`.

## The Tools
* `processor.py` will convert a multi-sheet excel file into a single CSV document. It works by looking for Excel files in a given directory (defualt is the current working directory); when it finds one (or many), it reads the excel file, one sheet at a time, and aggregates the data to a Pandas dataframe. Then it spits out a CSV file containing all the data from all the sheets in the file.
* `xl2csv.py` and `__init__.py` is a server that runs in the background and checks a folder (specified in the script) for Excel files at a given interval. When Excel files are found, the script converts them to CSV (files with multiple sheets will be rendered as multiple CSV documents) that are output to another folder in the server's designated working directory.
