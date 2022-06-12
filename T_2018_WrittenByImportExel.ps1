
# .\T_2018_WrittenByImportExel.ps1 

#Purpose: 
#	to write the T_2018.xlsx file 
#	throught the use of this code 
#	this file is a combination of functions 

#Package:
import-module -name ImportExcel 

# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #

 # ================================================================#
#1.     .\scrpt_days_in_yr.ps1 
	# to display all dates of a year in a list

# purpose: 
#	in $wk_days_ saves all days of the year  (Mon..Sun)
#	in $dates_   saves all dates of the year (MM/dd/yyyy)


# =================================================#

function get-yr_date ($year, $months)
	{
	
	$wk_days    = @()
	$dates_     = @()     
	for ($mon = 1; $mon -le $months; $mon++)
		{  
		$MaxNumDay = 1  
		$NumDayInMonth = 1..[DateTime]::DaysInMonth($year, $mon) 
		$MaxNumDay = ($NumDayInMonth| measure -Maximum).Count
	
		for ($Day_ = 1; $Day_ -le $MaxNumDay; $Day_++)
			{
			$wk_days += Get-Date -Year $year -Month $mon -Day $Day_ –Format "dddd"	
			$dates_  += Get-Date -Year $year -Month $mon -Day $Day_ –Format "MM/dd/yyyy"	
			}
		}
	
	# to Substring(0,3) the name of the day to 3 char
	$wk_days_ = @()  
	foreach ($day_ in $wk_days) 
		{
    		$short = $day_.subString(0, [System.Math]::Min(3, $day_.Length)) 
		$wk_days_ += $short
		}

	#write-host "beore" -BackGroundColor DarkRed -ForeGroundColor Black
	#write-host $wk_days_ -BackGroundColor Gray -ForeGroundColor DarkBlue
	#write-host $wk_days_.Count -BackGroundColor Gray -ForeGroundColor DarkMagenta
	#write-host $dates_ -BackGroundColor Gray -ForeGroundColor DarkBlue
	#write-host $dates_.Count -BackGroundColor Gray -ForeGroundColor DarkMagenta

	# ------------------------------------- 1st -------------------------------------

	$year_pre = $year-1
	# switch (var_deternine) (var_choice) command = incr./decr. 1ST WK of yr
	switch ($wk_days_[0])
		
		{
    		"Mon" 
			{
			$wk_days_ = $wk_days_
			$dates_ = $dates_
			}
    		"Tue" 
			{
			$Add_1St_day  = @("Mon")
			$Add_1St_date = @("12/31/$year_pre")
			$wk_days_ = $Add_1St_day  + $wk_days_
			$dates_   = $Add_1St_date + $dates_
			}
    		"Wed" 
			{
			$Add_1St_day =@("Mon", "Tue")
			$Add_1St_date = @("12/30/$year_pre", "12/31/$year_pre")
			$wk_days_ = $Add_1St_day  + $wk_days_
			$dates_   = $Add_1St_date + $dates_
			}
    		"Thu" 
			{
			$Add_1St_day =@("Mon", "Tue", "Wed")
			$Add_1St_date = @("12/29/$year_pre", "12/30/$year_pre", "12/31/$year_pre")
			$wk_days_ = $Add_1St_day  + $wk_days_
			$dates_   = $Add_1St_date + $dates_
			}
		"Fri" 
			{
			for ($i = 0; $i -le 2; $i++)
				{
				$wk_days_[$i] = $null
				$dates_[$i]   = $null
				}
			}
		"Sat" 
			{
			for ($i = 0; $i -le 1; $i++)
				{
				$wk_days_[$i] = $null
				$dates_[$i]   = $null
				}
			}
		"Sun" 
			{
			$wk_days_[0] = $null
			$dates_[0]   = $null
			}
		}
	
		# to get rid of $null if present
		$wk_days_ = $wk_days_.Where({ $_ -ne $null })
		$dates_   = $dates_.Where({ $_ -ne $null })

	#write-host "1st" -BackGroundColor DarkRed -ForeGroundColor Black
	#write-host $wk_days_ -BackGroundColor DarkBlue -ForeGroundColor Yellow
	#write-host $wk_days_.Count -BackGroundColor DarkBlue -ForeGroundColor DarkMagenta
	#write-host $dates_ -BackGroundColor DarkBlue -ForeGroundColor Yellow
	#write-host $wk_days.Count -BackGroundColor DarkBlue -ForeGroundColor DarkMagenta

	# ----------------------------- LAST -----------------------------------
	
	$year_post = $year+1
	# switch (var_deternine) (var_choice) command = incr./decr. LAST WK of yr
	switch ($wk_days_[-1])
		
		{
    		"Sun" 
			{
			$wk_days_ = $wk_days_
			$dates_ = $dates_
			}
    		"Mon" 
			{
			$wk_days_[-1] = $null
			$dates_[-1]   = $null
			}
			
    		"Tue" 
			{
			for ($i = 1; $i -le 2; $i++)
				{
				$wk_days_[-$i] = $null
				$dates_[-$i]   = $null
				}
			}
    		"Wed" 
			{
			for ($i = 1; $i -le 3; $i++)
				{
				$wk_days_[-$i] = $null
				$dates_[-$i]   = $null
				}
			}
			
		"Thu" 
			{
			$Add_Ls_day =@("Fri","Sat","Sun")
			$Add_Ls_date = @("1/1/$year_post", "1/2/$year_post", "1/3/$year_post")
			$wk_days_ = $wk_days_ + $Add_Ls_day
			$dates_   = $dates_ + $Add_Ls_date  
			}
			
		"Fri" 
			{
			$Add_Ls_day =@("Sat","Sun")
			$Add_Ls_date = @("1/1/$year_post", "1/2/$year_post")
			$wk_days_ =  $wk_days_ +$Add_Ls_day  
			$dates_   =  $dates_ + $Add_Ls_date 
			}
			
		"Sat" 
			{
			$Add_Ls_day  = @("Sun")
			$Add_LS_date = @("1/1/$year_post")
			$wk_days_ =  $wk_days_ + $Add_Ls_day  
			$dates_   = $dates_+ $Add_Ls_date 
			}
			
		}
		
		# to get rid of $null if present
		$wk_days_ = $wk_days_.Where({ $_ -ne $null })
		$dates_   = $dates_.Where({ $_ -ne $null })
	
	# =================================================	
	# produces a txt file
	$wk_days_ | out-file -LiteralPath "wk_days_.txt"
	$dates_ | out-file -LiteralPath "dates_.txt"
	# ===================================================

	#write-host "LAST" -BackGroundColor DarkRed -ForeGroundColor Black
	#write-host $wk_days_ -BackGroundColor DarkGray -ForeGroundColor DarkBlue
	#write-host $wk_days_.Count -BackGroundColor DarkGray -ForeGroundColor DarkMagenta
	#write-host $dates_ -BackGroundColor DarkGray -ForeGroundColor DarkBlue
	#write-host $dates_.Count -BackGroundColor DarkGray -ForeGroundColor DarkMagenta

	
	# ==================================================================
	write-host ".\scrpt_days_in_yr.ps1" -BackGroundColor Gray -ForeGroundColor DarkGreen
	# ==================================================================
	
	} 

$year   = 2024  # var
$months = 12    # var

get-yr_date $year $months 

# ==================================================================================================================== #
# ==================================================================================================================== #
# ==================================================================================================================== #
.\T_2018_WRVE_attch1_ImportExel.ps1

