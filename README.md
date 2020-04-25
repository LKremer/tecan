# tecan

## Read Absorbance Files of the Tecan Plate Reader Infinite 200
Automatically determines the table format and the number of wells that were recorded.
Reads all absorbance values accordingly and records all sheet names and file names.
Ignores but warns about empty or malformatted sheets.

## Example
```R
> library(tecan)
> read_tecan("tecan_example.xlsx")
# A tibble: 51 x 5                                                            
   well  blue_absorbance yellow_absorbance sheetname        filename          
   <chr>           <dbl>             <dbl> <chr>            <chr>             
 1 A1                0.6              0.75 plate format 3×9 tecan_example.xlsx
 2 A2                0.7              0.8  plate format 3×9 tecan_example.xlsx
 3 A3                0.8              0.85 plate format 3×9 tecan_example.xlsx
 4 A4                0.9              0.9  plate format 3×9 tecan_example.xlsx
 5 A5                1                0.95 plate format 3×9 tecan_example.xlsx
 6 A6                1.1              1    plate format 3×9 tecan_example.xlsx
 7 A7                1.2              1.05 plate format 3×9 tecan_example.xlsx
 8 A8                1.3              1.1  plate format 3×9 tecan_example.xlsx
 9 A9                1.4              1.15 plate format 3×9 tecan_example.xlsx
10 B1                0.6              0.75 plate format 3×9 tecan_example.xlsx
# … with 41 more rows
Warning messages:
1: In autoread_tecan_sheet(xlsx_path, sheet = s) :
  Ignoring sheet 'unrelated sheet' of 'tecan_example.xlsx' since it has the wrong format.
2: In autoread_tecan_sheet(xlsx_path, sheet = s) :
  Ignoring sheet 'random empty sheet' of 'tecan_example.xlsx' since it has the wrong format.
```
