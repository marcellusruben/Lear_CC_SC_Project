# Lear CC SC Project

## About CC-SC

CC-SC is an important procedure in the design phase that the design and project engineer at Lear need to do before deploying a proposal regarding the final design of vehicle seat structure. For a seat structure, it is particularly important to check how many welds can be considered as CC and SC. 

If a weld is considered as CC, it means that this weld carries a heavy amount of load so that a regular maintenance regarding this weld is necessary. Meanwhile, weld that is categorized as SC is also carrying some loads, but not as much as CC welds. Thus, in order to know which welds carry a lot of loads, a result from Finite Element (FE) was used. Then, a report regarding which welds can be considered as CC or SC can be made.

## Problem
- In a design phase, regular modifications and change towards several parts of a seat structure is needed to find the best design possible. This means that for each modification of the design, the CC-SC of the welds needs to be continuously examined again and again.
- However, the examination of CC-SC welds is a very tedious task to do. There are hundreds of welds in each seat structure and each of them needs to be categorized whether they are CC, SC, or regular welds. FE simulation can give us the contour plot of the result, but it can't give the project engineer a detailed overview and report regarding which welds that are considered as CC or SC, how many of the welds that are classified as CC, and what is the percentage of the welds that are categorized as CC or SC. 
- Due to this FE limitation, there should be one engineer that needs to examine the FE result and then manually inputting the classification of the welds (CC or SC) into a spreadsheet. This is very tedious task to do and it takes about two days to finish the categorization of the welds in each seat structure. 

## Objective
The main purpose of this script is to automatize all of the tedious task described in the Problem section above. With this script, no person needed to manually inputting the result from FE to Excel to create some detailed reports. 

This script will automatically manipulate all of the FE raw data and automatically classify all of the welds as CC, SC, or regular welds. As the output, an Excel file will be generated which contains a detailed report regarding which welds that are classified as CC-SC, their location ID, how many welds classified as CC-SC, and lastly, what is the overall percentage of the welds that are classified as CC. This is an important step because one of the final goal of the design phase is to have CC welds as few as possible.

Below is the flow on how this script works:
<p align="center">
  <img width="460" height="300" src="https://github.com/marcellusruben/Lear_CC_SC_Project/blob/master/Flow.png">
</p>

## Files
There are four files in this project: two Python files, one Excel file, and one Text file.
- WS_CC_SC.py - the main file of the Python script to automatically generate Excel report.
- WS_CC_SC_Function.py - the function file to supplement the main Python script.
- InputFile.xls - the input file where the user specify necessary variables to be provided to Python script.
- InputFile.txt - the Text file version of InputFile.xls, which will be used by the Python script to do data wrangling, data manipulation, and generating the report at the end.


