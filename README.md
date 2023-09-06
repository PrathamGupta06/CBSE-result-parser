# CBSE Result TXT to Excel Converter and Analyzer
A CLI tool to convert the Class 10th and 12th Result Text File to Excel with proper formatting and analyze the result and display some statistics.
Schools receive the results of students in a .txt file from CBSE. This tool converts the .txt file to .xlsx and then displays some statistics of the student results.
## Features
* ⚡ Simple and easy to use. Single Click
* ️✅ Converts the txt file into a properly formatted Excel file.
* 🔢 Different Spreadsheet page for each Subject
* 📺 Displays the statistics such as 
  * 🎓 Top 5 Male and Female Students
  * 💯 Children with full marks in individual subjects
  * 📔 Number of Distinctions in all 5 subjects
  * 📑 Number of Distinctions
## Demo Files
* 12th : [Click Here]()

## How to Run
### There are two methods to run
1. Google Colab (Easiest)
2. Locally using python

### Method 1: Google Colab (Easiest)
1. Go to the link (https://colab.research.google.com/drive/1ardBfRG_S40qejG5VnCVWIJYEGpSHluI?usp=sharing)
2. Run each cell one by one
3. Rest of the instructions are given in the colab file

### Method 2: Locally using Python
1. Install [Python](https://www.python.org/downloads/) (if not already)
    * While Installing make sure to check the add to system PATH option
2. Download the code by clicking on the Green Code button then Download as ZIP
4. Extract the ZIP file and paste your result text file in the extracted folder
5. Open the terminal in the extracted folder and run `pip install -r requirements.txt`
#### Method 2.1: Using the CLI
6. Run `main.py`
    * Enter the file name
    * Enter the output file name. Make sure to enter .xlsx in after the file name
    * Enter the mode i.e. the class. It can accept only two values i.e. `12th` and `10th`
    * Exported File will automatically Launch.

#### Method 2.2: Using command line arguments (for automation):
6. Run `main.py -i <input_file_name> -o <output_file_name> -c <class>`
    * `<input_file_name>` is the name of the input file
    * `<output_file_name>` is the name of the output file
    * `<class>` is the class of the result. It can accept only two values i.e. `12th` and `10th`
    * Excel file will be saved as `<output_file_name>.xlsx` in the same folder as the input file
## You can run the program again for another class
