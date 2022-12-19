# Component level test
- Purpose of this program is to convert keysight output files to standard component level testing which includes NEXT, RL and IL.

## Project setup instructions 
To start using this project follow the given description:
1. `Module1.openAllWorkbooks() line 24 testDirectory`
   -  Paste here path As String to folder which contains tests. If files are seperated in different folders/subfolders, then choose folder that is first highest in tree (includes each subfolder).

2. `Module2.conversion() in line 15 folderPathXLSX`
   -  Paste here path As String to folder where you want to save output xlsx files (choose only one destination). 

3. `Module4.ilLimit() line 42 ilFormula`  
`Module4.nextLimit() line 60 nextFormula`   
`Module4.rlLimit() line 77 rlFormula`
   - Formulas are set for C6 testing. If you want to test different category cable, change formulas to ones corresponding to category limit.  
   
## Code requirements/steps
- [x] Convert .csv files with measurements to .xlsx
- [x] Delete redundant measurements
- [x] Adjust units and add limits to worksheets
- [x] Write margins, mark them with green if they are good and red if not, save the worst margin
- [ ] Create a chart with limits and measurements on
- [ ] Create different worksheet containing frequency, limits and margins from previous worksheets with the same measurement type

## License info
Exclusive Licence   

Copyright (c) 2022 Molex Connected Enterprise Solutions   
Author: [Mateusz Pieczykolan](https://github.com/pejczykjr)