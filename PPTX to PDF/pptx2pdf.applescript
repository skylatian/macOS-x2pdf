# credits:
# https://stackoverflow.com/questions/41886380/using-applescript-to-open-ms-powerpoint-2016-file
# https://superuser.com/questions/670893/get-path-of-parent-folder-of-script-location-applescript

on savePowerPointAsPDF(documentPath, pdfPath)
	set inFile to documentPath as alias # this line must be outside of the 'tell application "Microsoft PowerPoint"' block  to avoid issues with the open command
	
	tell application "Microsoft PowerPoint"
		
		launch # open powerpoint
		open inFile # opens file
		
		# create empty file to avoid issues with save command
		set pdfPath to my createEmptyFile(pdfPath) # the handler return a file object (this line must be inside of the 'tell application "Microsoft PowerPoint"' block to avoid issues with the saving command)
		
		# sometimes necesarry to avoid a failure
		delay 1
		
		# saves given ppt(x) as pdf
		save active presentation in pdfPath as save as PDF
		
	end tell
	
end savePowerPointAsPDF

on createEmptyFile(eFile)
	
	# create file (this will do nothing if the file exists)
	do shell script "touch " & quoted form of POSIX path of eFile
	
	# sometimes necesarry to avoid a failure
	delay 1
	
	# returns empty file
	return (POSIX path of eFile) as POSIX file
	
end createEmptyFile

# gets output path from input path (replace extension with .pdf and return as string)
on getOutputPath(docPath)
	
	# casts to string, just in case
	set dpath to docPath as string

	if dpath ends with ".pptx" then
		# replace pptx
		return (text 1 thru -6 of dpath) & "-pptx.pdf"
	else
		# replace ppt
		return (text 1 thru -5 of dpath) & "-ppt.pdf"
	end if
end getOutputPath

# retrive parent directory of given filepath
on getParentDirectory(filePath)

	tell application "Finder"
		
		# retrives parent directory of given filepath
		set parentDirectory to POSIX path of ((filePath as text) & "::")
		
		# debug: show alert with parent directory
		-- set theDialogText to "parent folder: [" & (parentDirectory) & "]" 
		-- display dialog theDialogText
		
	end tell

	return parentDirectory
	
end getParentDirectory

on movePowerPointFilesToSubfolder(sourceFolder, fileToMove, subFolderName)
	tell application "Finder"

		# convert the POSIX path of the source folder to an alias for Finder operations
		set sourceFolderPath to POSIX file sourceFolder as alias

		# concatenate the source folder's path with the subfolder name to create the full path of the subfolder
		set subFolderFullPath to sourceFolderPath & subFolderName as string
		
		# check if the subfolder exists, create it if not
		if not (exists folder subFolderFullPath) then
			make new folder at sourceFolderPath with properties {name:subFolderName}
		end if
		
		# Move file to the subfolder
		move fileToMove to subFolderFullPath
		
	end tell
end movePowerPointFilesToSubfolder

# this "run" loop defines what happens when triggered from quick actions
on run {input, parameters}
	
	repeat with i in input
		set inPath to i as string
		
		# debug: display input path
		-- set theDialogText to "input: [" & inPath & "]"
		-- display dialog theDialogText
		
		# calculate output path for current file
		set outPath to getOutputPath(inPath)
		
		# gets parent directory of current file (kinda redundant but it works)
		set parentDir to getParentDirectory(i)
		
		# save converted PDF in original directory
		savePowerPointAsPDF(inPath, outPath)
		
		# move powerpoint files into subfolder
		movePowerPointFilesToSubfolder(parentDir, inPath, "powerpoint files")
		
	end repeat
	
	# quit powerpoint
	tell application "Microsoft PowerPoint"
		quit
	end tell
	
end run