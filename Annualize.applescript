----------------------------------------------------------------------------------------------------------------
-- Script
--          Annualize 
--
-- Description 
--          Move selected messaged into a folder by Year
--          Apply a category based on source folder name
--
-- Author
--          Sri Shivananda
--
-- Date
--          November 04, 2012 
----------------------------------------------------------------------------------------------------------------

-- Get the current year
set currentYear to year of (get current date)

-- This is a script that commands Microsoft Outlook
tell application "Microsoft Outlook"
	
	-- Get the name of the selected foler
	set theFolder to selected folder
	set folderName to name of theFolder
	
	-- Get the "On My Computer" Folder
	set onMyComputer to on my computer
	
	-- Get all the sub folders in the "On My Computer" folder
	set subFolders to name of folders of onMyComputer
	
	-- Get a list of all categories
	set categoryList to name of categories
	
	
	-- Get all the selected messages, timeout if not possible in 5 minutes
	with timeout of 300 seconds
		set allMessages to current messages
	end timeout
	
	-- Iterate through all the messages
	repeat with aMessage in allMessages
		
		-- Get the time sent for the message
		set theDate to time sent of aMessage
		
		-- Get the year in which the message was sent
		set theYear to year of theDate
		
		-- Only run this for messages in the previous years
		if theYear < currentYear then
			
			-- Create a new subfolder with the year
			-- Do this only if the folder does not already exist
			set subFolderName to "Y" & theYear
			if subFolderName is not in subFolders then
				make new mail folder at onMyComputer with properties {name:subFolderName}
				copy subFolderName to end of subFolders
			end if
			
			-- Using the folder name, create a category, if one does not already exist
			if folderName is not in categoryList then
				make new category with properties {name:folderName, show in navigation pane:false}
				copy folderName to end of categoryList
			end if
			
			-- Set the category (foldername) on the message for searching later 	
			set category of aMessage to {category folderName}
			
			-- Move the message to the annualized folder
			move aMessage to folder subFolderName of onMyComputer
		end if
	end repeat
end tell