

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

on docx2pdfCommand(command)

	# set PATH and run command
	set docx2pdfCommand to ("export PATH=$PATH:/Users/skyler/.pyenv/shims/; docx2pdf " & command & "" )
	set outputDiag to do shell script docx2pdfCommand
	
	-- Optional: display the output for debugging
	display dialog outputDiag


end docx2pdfCommand

on run {input, parameters}
	

	
	return input
end run
