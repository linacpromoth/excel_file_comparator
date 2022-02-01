# excel_file_comparator
Tool to compare two spreadheets

A simple tool written in python to compare two spreadsheets of the formats - **.xlsx/ .xls / .csv**.It helps us to look for addition/deletion/modification that happened between two similar datasets.It is useful in cases when you want to keep track of changes made in a file over a period of time by comparing it with base file.

## Requirements
* Python 3.8.2+
* Python Libraries
    * pandas==1.3.3 ($ pip install pandas)

## Installation
Clone the git and you get the ball rolling!!!
```
$ git clone https://github.com/linacpromoth/excel_file_comparator
$ cd excel_file_comparator
```
Move the **base file** to compare with and the **target file** to comapare against to the **excel_file_comparator** directory and then executes the python file.An output .xlsx file will be generated with the following two sheets
- **output** sheet -> contains the target file values with the changes highlighted.
- **summary** sheet -> contains the summary of the changes made on the target file.  

  
## Usage
```
$ python3 file_comparator.py -b <baseFile.xlsx> -t <targetFile.xlsx>

usage: file_comparator.py [-h] -b BASE_FILE -t TARGET_FILE

Compare target file with base file

optional arguments:
  -h, --help            show this help message and exit
  -b BASE_FILE, --base_file BASE_FILE
                        the base file against which comparison takes place
  -t TARGET_FILE, --target_file TARGET_FILE
                        the target file to be compared with
```
## Example output
![image](https://user-images.githubusercontent.com/98702521/151831926-d5cbdb2f-248a-4099-8a29-169134ce71c8.png)

### Output file 
![image](https://user-images.githubusercontent.com/98702521/151830841-dbaf8b47-0198-4634-bc40-c0b87048f002.png)

### Summary output
![image](https://user-images.githubusercontent.com/98702521/151830963-734b04b4-b211-46f9-baa6-9e7f8bb43cf2.png)


# Limitations
* Duplicates are removed by default
* First common columns between both base and target file is considered as primary identifier column, so its better to ensure that column is unique(case-sensitive).
* Deleted records wont be captured in the output file generated.

