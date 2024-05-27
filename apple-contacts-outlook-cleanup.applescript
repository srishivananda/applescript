--
-- script: apple-contacts-outlook-cleanup
-- 
-- description: outlook adds some custom fields when contacts are shared between outlook and apple
--                   remove outlook custom fields from apple contacts
-- 
-- actions       : 1. remove outlook url field
--                    2. remove the note "This contact is read-only. To make changes, tap the link above to edit in Outlook."
-- 
-- runtime      : Apple Script Editor on Mac OS
-- 
-- author       : Sri Shivananda
--
tell application "Contacts"
	delete (every url of every person whose value contains "outlook")
	set note of every person whose note starts with "This contact is read-only. To make changes, tap the link above to edit in Outlook." to missing value
	save
end tell
