# .\T_2018_WRVE_attch1_ImportExel.ps1

# Package: 
import-module -name ImportExcel

# Purpose:
#	it populates ROWS of .xlsx with data (up to 676 col)
#	Can Dispaly values on screen $x (and more) if uncommented


function WriteRow-Excel ($Path, $Sheet, $Val, $RowNum, $ColNum, $col_Strt)
	{ # function
	
	$excel = Open-ExcelPackage -Path $Path	
	$len     = $ColNum.Length         # Length of letteres a - z = 26
	$len_Val = $Val.Length            # Length of $val list
	
	$index = 0       # used in the "else" for 1st column count
	$count = 0       # used in the "if" & "else" -> counts all
			 #	-> keeps track of all values

	for ($i=0; $i -lt $len; $i++)
		{ # for i

		$Jndex = 0                       # used in the "else" for 2nd column count
		for ($j=0; $j -lt $len; $j++)
			{# for j  
			
			if ($count -lt $len_Val)
				{ # for if $len_Val
	
				if ((($col_Strt+$j) -lt $len) -and ($i -eq 0))       # 1st round
					{ # for if

					# indeces of the .xlsx -make into [string] -> have to be
					[string]$1st   = $ColNum[$j+$col_Strt]
					[string]$Value = $Val[$count]

					$excel.$Sheet.Cells[$1st+$RowNum].Value = $Value

					$count++
					}

				# Data of this Bracket are not used (disregard extra combinations)
				elseif ($i -eq 0) 
					{ # for elseif 
					
					}

				else                                                 # all rounds except (1st round)
					{ # for else 

					# indeces of the .xlsx -make into [string] -> have to be
					[string]$1st   = $ColNum[$index-1]
					[string]$2nd   = $ColNum[$Jndex]
					[string]$Value = $Val[$count]

					$excel.$Sheet.Cells[$1st+$2nd+$RowNum].Value = $Value

					
					$count++
					$Jndex++
					}
				} # for if $len_Val
			
			else 
				{ # for -ne if $len_Val
				break; 
				}
			
			} # for j 
		$index++

		} # for i 
	Close-ExcelPackage -ExcelPackage $excel -show
	
	# -----------------------------------------------------------------
	write-host '.\T_2018_WRVE_attch1_ImportExel.ps1' -BackGroundColor DarkRed -ForeGroundColor Black 
	write-host 'fun = WriteRow-Excel' -BackGroundColor DarkRed -ForeGroundColor Black 
	write-host 'outputfile = ' $Path -BackGroundColor DarkRed -ForeGroundColor Black  
	}

# ------------ param ----------------------------------
$Path =   'test.xlsx'
$Sheet =  'Sheet1'              # sample sheet name

# ///////////// variables that change - rest stay the same \\\\\\\\\\\\\\\\\\\\\\\\\
$Val1    = get-content dates_.txt
[string]$RowNum1 = '1'

$Val2      = get-content wk_days_.txt
[string]$RowNum2  = '2'

# //////////////////////////////////////\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
# ------------ pre-fn parameters -----------------------
$a= 97                         # ASCII = a
$z= 122                        # ASCII = a
$col_Strt = 0                  # starting col position 

$int = $a..$z                  # Range from a - z        
           
# ------------ param ----------------------------------
$ColNum =[char[]]$int         # turns Range of int -> Range of letters
# ////////////////////////////////////////////////////////////////////////

WriteRow-Excel $Path $Sheet $Val1 $RowNum1 $ColNum $col_Strt 
WriteRow-Excel $Path $Sheet $Val2 $RowNum2 $ColNum $col_Strt 
