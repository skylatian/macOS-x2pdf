on run {input, parameters}

tell application "Finder"
	
	set theItems to input
	
	repeat with itemRef in theItems
		
		set theItemParentPath to (container of itemRef) as text
		set theItemName to (name of itemRef) as string
		set theItemExtension to (name extension of itemRef)
		set theItemExtensionLength to (count theItemExtension) + 1
		set theOutputPath to theItemParentPath & (text 1 thru (-1 - theItemExtensionLength) of theItemName)
		
		tell application "Microsoft PowerPoint"
			
			open itemRef
			
			tell active presentation
				
				save in theOutputPath as save as PDF
				close
				
			end tell
			
		end tell
		
	end repeat
	
end tell

end run