# Batch TOPAS output to Excel 

## Introduction

This PowerShell script will look for .txt files of QPA results in a folder, then create a .xlsx file of all the results 
and a pivot table. There is also the option to sort the list of phases by Dana classification number using the expanded 
Dana numbers v2 spreadsheet from the [Mineralogical Society of Southern California](https://mineralsocal.org/398-2/). Note, the phase name in the TOPAS output
must be a match in this spreadsheet. Unknown or synthetic minerals will be given a default code of 99.99.99.99.99.99. 

As TOPAS will write results for all refinement convergence situations, or when terminated by the user, the last listed is 
considered the most up to date and the one to export to excel. Assumptions for successful application of the script:
- The TOPAS output is formatted a specific way. Example .INP file given
- All of the TOPAS output .txt files you want converted into excel and shown in the ONE(1) pivot table are all in the same folder
- The .bat and .ps1 files are all in the same folder
- The .txt files containing QPA results arise from only ONE(1) specific XRD data file
- You have manually created a .csv from the Distribution 2025 spreadsheet from the Mineralogical Society of Sourthern California
  - Row 1 contains the following labels
    - A1 = Phase
    - B1 = X
    - C1 = N1
    - D1 = N2
    - E1 = N3
    - F1 = N4
    - G1 = N5
    - H1 = N6
  - Row 2 and onwards are the copied entires from the Distribution 2025 spreadsheet
    - A2 and beyond are the actual mineral names
    - B2 to H2 inclusive and beyond are the values in Columns E to K inclusive (i.e. the expanded Dana numbers v2)
- The above .csv is saved as xdana_cache.csv
- This .csv is in the same folder as the files to convert

## How to run

Double click the Run-Batch-Phases.bat file assuming you have succesfully implemented the above instructions. 

## Suggested formatting of pivot table 

- Click anywhere in pivot table
- Click on Design in the excel ribbon
- Report layout --> Show in tabular form
- Subtotals, do not show subtotals
- Sort by dana number if desired or keep your inp file order
- Toggle Grand totals if needed

## Expandable work 

could be easily expanded to other parameteters if needed, e.g. crystallite size, strain etc. 
