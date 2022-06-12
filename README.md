# README.md
## This repo has some fun to play scripts for PowerShell

### The pair of functions `T_2018_WrittenByImportExel.ps1` & `T_2018_WRVE_attch1_ImportExel.ps1`
###  If your goal is to write all *days* starting from Monday
###			and all associated *dates* for the whole year
###			*(pick any year you want)*

fist use the `T_2018_WrittenByImportExel.ps1` function to:
	pick the year you want 
	the function will create two text files
		1. with all days of the year starting from Monday to Sunday
		2. pair each day with the appropriate date

Then the `T_2018_WRVE_attch1_ImportExel.ps1` function:
	will project in to the excel (*at the row of your choice*)
	all days 
	and 
	at the row below all all coresponding dates for the whole year

The `T_2018_WrittenByImportExel.ps1` function is the function you will need to call
	the `T_2018_WrittenByImportExel.ps1` will autocall the `T_2018_WRVE_attch1_ImportExel.ps1` from within 

*make sure you have adjust all variables in both 
`T_2018_WrittenByImportExel.ps1` & `T_2018_WRVE_attch1_ImportExel.ps1` 
to your specification prior to run them*

**have fun playing with them !**
