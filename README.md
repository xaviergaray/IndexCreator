# Index Creator
# Status
Currently only the CLI is operational with plans to run the program on a website.
## Requirements
* Python
  * pandas (pip install pandas)
  * pyarrow (pip install pyarrow)
  * docx (pip install python-docx)
  * openpyxl (pip install openpyxl)
  * re
  * argparse
* Excel file (xlsx) representing index
  * Recommended to have Topic, Book, Page, Notes as headers but could be whatever you'd like
## Description
This script will create an index in .docx format and includes some customizable options. Run with -h to find out more.
