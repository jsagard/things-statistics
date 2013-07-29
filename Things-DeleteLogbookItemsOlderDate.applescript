#
# Things-DeleteLogbookItemsOlderDate.applescript
#
# Version 1.0
#
# An AppleScript for deleting all items (completed todos and projects) from Logbook list of Things application that older than specific date.
#
# Things Statistics is released under the MIT license:
# Copyright (c) 2013 @ixrevo (Alexander Sokol http://ixrevo.ru)
#

##########
# Config #
##########

#
# Empty Trash before script starts
property emptyTrash : true
#
# Log completed items before script starts
property logCompleted : true
#
# Play sound after script's completion
property playsound : true

############
# Handlers #
############

#
# Convert date to string with 'dd.MM.YYYY' format.
#
# theDate    date    Date to fromat.
#
# return     text    String with formatted date.
#
on dateToString(theDate)
	set cYear to year of theDate
	# month
	set cMonth to month of theDate as integer
	if length of (cMonth as text) is 1 then set cMonth to "0" & cMonth
	# day
	set cDay to day of theDate
	
	return cDay & "." & cMonth & "." & cYear
	
end dateToString

#
# Convert string with 'dd.MM.YYYY' format to date with 1:00:00 time.
#
# theString    text    String to convert.
#
# return       date    Converted date.
#
on stringToDate(theString)
	
	set theString to date "1:00:00 AM" of date theString
	
	set time of theString to 0
	
	return theString
	
end stringToDate

#
# Validate that string have 'dd.MM.YYYY' date format.
#
# theDate    text    Date string to validate.
#
# return     boolean    True if string have correct format or false if not.
#
on isValidDate(theDate)
	try
		if (do shell script "echo " & theDate & " |egrep -c '^(((0?[1-9]|[1-2][0-9]|3[0-1])\\.(0[13578]|(10|12)))|((0[1-9]|[1-2][0-9])\\.02)|((0[1-9]|[1-2][0-9]|30)-(0[469]|11)))\\.[0-9]{4}$'") is "1" then
			return true
		else
			return false
		end if
	on error number 1
		return false
	end try
end isValidDate

###############
# Main Script # 
###############

#
# Empty Trash
#
if emptyTrash then
	tell application "Things"
		empty trash
	end tell
end if
#
# Log Completed Items
#
if logCompleted then
	tell application "Things"
		log completed now
	end tell
end if
#
# Get date after that items (todos and projects) will be deleted.
#
set tmpDate to my dateToString((current date) - 365 * days)
repeat
	set exactDate to display dialog "Trash items from Logbook older than:" default answer tmpDate as text buttons {"Cancel", "Delete"} default button 1 cancel button 1 with title "Delete items from Things Logbook?" with icon caution
	if button returned of exactDate is "Delete" then
		set exactDate to the text returned of exactDate
		set thresholdDate to my stringToDate(exactDate)
		if not (my isValidDate(exactDate)) then
			display alert "Date " & exactDate & " is incorrect. Correct date fromat: dd.mm.yyyy"
		else if thresholdDate is greater than (current date) then
			display alert "Date is greater than current date."
		else if ((date string of thresholdDate) is equal to (date string of (current date))) then
			display alert "Date is equal to current date."
		else
			exit repeat
		end if
	else
		return "Unexpected input. Script stopped."
	end if
	
end repeat
#
#  Delete items (todos and projects)
#
property logbookList : {}
set startTime to current date
tell application "Things"
	set wasDeleted to 0
	
	set my logbookList to to dos of list "Logbook"
	set numberOfLogbookItems to count my logbookList
	
	repeat with i from 1 to numberOfLogbookItems
		# Get current items
		set thisItemProperties to (my logbookList's item i)'s properties
		# Process items
		if thisItemProperties's completion date is less than thresholdDate then
			set wasDeleted to wasDeleted + 1
			move (my logbookList's item i) to list "Trash"
		end if
	end repeat
end tell
#
# Ending
#
if playsound is true then do shell script "/usr/bin/afplay /System/Library/Sounds/Glass.aiff"
# Calculate time of the execution
set endTime to current date
set executionHours to (endTime - startTime) div hours
set executionMinutes to (endTime - startTime) div minutes
set executionSeconds to (endTime - startTime) - (executionHours * hours) - (executionMinutes * minutes)
set executionTime to executionHours & ":" & executionMinutes & ":" & executionSeconds
# Statistics
set finishMsg to "Logbook: " & numberOfLogbookItems & " items." & return
set finishMsg to finishMsg & return
set finishMsg to finishMsg & wasDeleted & " items older than " & my dateToString(thresholdDate) & " was moved to Trash." & return
# Time of ececution
set finishMsg to finishMsg & return
set finishMsg to finishMsg & "Script starting: " & time string of startTime & return
set finishMsg to finishMsg & "Script ending:  " & time string of endTime & return
set finishMsg to finishMsg & "Time of execution: " & executionTime
display dialog (finishMsg)