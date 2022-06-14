# scrpt_find_dirfromfile.ps1

# Name = find-dirfromfile
	# purpose = find directory based on file.ext ->
		# write the path of the directory on the screen
		# excecute commands in direcotry
		# reutrn to directory wanted by user

	# more:
		# only checks given locations for finding a document (fast)
		# need to know the name of the file.ext
		# excecute many commands in that directory
		# can return to any direcory chosen

function find-dirfromfile ($item_ , $path_return)       
	{
	# you need to provide the path you want the code to recursely check:
	$locations = @(
		"C:\Users\put down your path",
		"C:\Users\put down your path", 
		"C:\Users\put down your path"
		)

	for ($i=0; $i -lt $locations.length; $i++) 
		{

		$display = get-childitem -path $locations[$i] -recurse  -include $item_ -file 
		
		$full_loc = $display.FullName
		$name_ =$display.name

		if ($name_) 
			{
			# $dir_loc is the dir that the file is located at:
			$dir_loc = $full_loc.replace('\'+$name_, "") 

			set-location $dir_loc
			
			# here you can place as many commands 
			# as you want to be excecuted inside the 
			# files's directory

			# return to disignated directory
			set-location $path_return
			
			write-host $dir_loc  -ForegroundColor DarkGreen -BackgroundColor White
			
			break   
			}

		}
	}

$item_ = "file_name.extension"                                          # this is the file that you are serching for
$path_return = "C:\Users\dir that you want to returb afterwards"        # this is the directory that you want to return when you done 

find-dirfromfile $item_ $path_return 















