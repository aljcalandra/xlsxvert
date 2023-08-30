# xlsxvert

This is a simplistic script for converting from xlsx to csv files

## Usage

```
usage: xlsx_convert.py [-h] -i INFILE [-o DEST_PATH] [-d DELIMITER]

convert xlsx to csv, outputs to files based on sheet names

optional arguments:
  -h, --help            show this help message and exit
  -i INFILE, --infile INFILE
                        Filename
  -o DEST_PATH, --output_directory DEST_PATH
                        Output Path
  -d DELIMITER, --delimiter DELIMITER

```

Only single character delimiters are allowed with the exception of tab which will be replaced with `\t`.

Files will be output based on sheet names

Example:

```
workbook
\
 sheet1
 sheet2

output_directory (default current)
\
 sheet1.csv
 sheet2.csv
```
