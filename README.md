# model-input-creator

Allows the user to load a pool of loans and converts the data into readable input files for AD&Co and Moody's models.

## Description

WeeklyBestEx.py is a Tkinter GUI python script that is used to automate the tape cracking proccess for our cashflow and credit enhancement models. The application receives a loan pool excel file provided by Agency desk and creates two separate model input files for AD&Co LoanKinetics and Moody's MILAN US Model. The files are then ready to input to the corresponding model to return the outputs needed to determine best execution at the loan level.

## Getting Started

### Dependencies

Libraries needed: pyodb, numpy, pandas, tkinter, openpyxl, datetime
Files needed: ADCO Input Template, Moodys Input Template

### Installing

- pip install the libraries stated under dependencies
- download the model input templates and move to an appropriate folder
- within the script change ADCO & Moody's path to where the model input templates are located
- within the script change the output path to desired path for the complete input files
- move script into desired IDE

### Executing program

- Run script within desired IDE
- Click "Open File" Button
- Select the pool of loans you would like to create input files for
- After you select a file wait until you see the path to the file show up in the text box below "Showout File"
- Under the drop down menu that appears, select the worksheet you'd like to create inputs for
- When you have decided on a worksheet, click "Create Inputs"
- You will be prompted with a green "Done!" box when the model input files have been created
- Check your desired output path for two created files, labeled by the corresponding model

## Help

Common Issue: Be cautious to not save over the blank input template files prior to running the python script. This will cause a key error due to a duplication of the "PIW" column. If this error arrises just restore the input template files back to their blank state.

## Authors

Emmanuel Lopez
LinkedIn: (https://www.linkedin.com/in/lopez-emmanuel/)

## Version History

* 1.0
    * Initial Release

