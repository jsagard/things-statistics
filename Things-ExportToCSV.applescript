#
# Things-ExportToCSV.applescript
#
# If you want export data from things into CSV file than you can use this script.
#
# Things Statistics is released under the MIT license:
# Copyright (c) 2013 @ixrevo (Alexander Sokol http://ixrevo.ru)
#

##########
# Config #
##########

#
# Path to file
property fileFolder : (path to desktop as Unicode text)
#
# Extension of the file
property fileExtension : "csv"
#
# First part of the file name
property fileNamePrefix : "Things"
#
# Export data for last days or weeks
property exportLast : (current date) - (365 * days + 0 * weeks)
#
# Empty Trash before script starts
property emptyTrash : true
#
# Log completed items before script starts
property logCompleted : true
#
# Export completed todos?
property exportTodos : true
#
# Export projects?
property exportProjects : true
#
# Export areas?
property exportAreas : true
#
# Play sound after script's completion
property playsound : true

############
# Handlers #
############

#
# Write text to a file using utf-8 encoding.
#
# theText       text               Text to write to file
# theFile       string or alias    File to write
# checkComma    boolean            Check for the commas and quote text if its contained in the text.
#
# return        nothing
#
on writeToFile(theText, theFile, checkComma)
	
	if class of theText is not text then set theText to theText as text
	
	if checkComma then
		if theText contains "," then set theText to quote & theText & quote
	end if
	
	try
		set theFile to open for access file theFile with write permission
		write theText to theFile as Çclass utf8È starting at eof
		close access theFile
	on error errorMsg number errorNumber
		close access theFile
		error "Can't write to file: " & errorNumber & " " & errorMsg
	end try
end writeToFile

#
# Append ',' to the end of the file.
#
# theFile       string or alias    File to write
#
# return        nothing
#
on nextColumn(theFile)
	try
		set theFile to open for access file theFile with write permission
		write "," to theFile as Çclass utf8È starting at eof
		close access theFile
	on error errorMsg number errorNumber
		close access theFile
		error "Can't write to file: " & errorNumber & " " & errorMsg
	end try
end nextColumn

#
# Append 'end of line' symbol to the end of the file.
#
# theFile       string or alias    File to write
#
# return        nothing
#
on nextRow(theFile)
	try
		set theFile to open for access file theFile with write permission
		write return to theFile as Çclass utf8È starting at eof
		close access theFile
	on error errorMsg number errorNumber
		close access theFile
		error "Can't write to file: " & errorNumber & " " & errorMsg
	end try
end nextRow

###############
# Main Script # 
###############

property listOfTodos : {}
property numberOfTodos : 0
property counterOfExportedTodos : 0
property listOfProjects : {}
property numberOfProjects : 0
property counterOfExportedProjects : 0
property listOfAreas : {}
property numberOfAreas : 0
property counterOfExportedAreas : 0
#
# Start
#
set startTime to current date
#
# Set file names with path to file
#
set currentDate to current date
# year
set cYear to year of currentDate
# month
set cMonth to month of currentDate as integer
if length of (cMonth as text) is 1 then set cMonth to "0" & cMonth
# day
set cDay to day of currentDate
# hour
set cHour to hours of currentDate
if length of (cHour as text) is 1 then set cHour to "0" & cHour
# minute
set cMinute to minutes of currentDate
if length of (cMinute as text) is 1 then set cMinute to "0" & cMinute
# second
set cSecond to seconds of currentDate
if length of (cSecond as text) is 1 then set cSecond to "0" & cSecond
# timestamp
set timeStamp to cDay & "-" & cMonth & "-" & cYear & " " & cHour & "-" & cMinute & "-" & cSecond
# files
set logbookFile to my fileFolder & my fileNamePrefix & " " & "Logbook" & " " & timeStamp & "." & my fileExtension
set projectsFile to my fileFolder & my fileNamePrefix & " " & "Projects" & " " & timeStamp & "." & my fileExtension
set areasFile to my fileFolder & my fileNamePrefix & " " & "Areas" & " " & timeStamp & "." & my fileExtension
#
# Empty Trash
#
if my emptyTrash then
	tell application "Things"
		empty trash
	end tell
end if
#
# Log Completed Items
#
if my logCompleted then
	tell application "Things"
		log completed now
	end tell
end if
#
# Export completed todos
#
if my exportTodos then
	# Set the header
	my writeToFile("Creation Date,Due Date,Completion Date,Name,Area,Project,Tags,Completion Year,Completion Month,Completion Weekday", logbookFile, false)
	my nextRow(logbookFile)
	# Start processing todos
	tell application "Things"
		# Get list of all todos
		set my listOfTodos to to dos
		# Get number of all todos
		set my numberOfTodos to count my listOfTodos
		# Process todos
		repeat with i from 1 to my numberOfTodos
			# Get current todo's properties and completion date
			set thisTodoProperties to (my listOfTodos's item i)'s properties
			set thisTodoCompletionDate to thisTodoProperties's completion date
			if (thisTodoProperties's class is not project) and (thisTodoProperties's status is completed) and (thisTodoCompletionDate is greater than my exportLast) then
				# Creation Date
				my writeToFile(thisTodoProperties's creation date, logbookFile, true)
				my nextColumn(logbookFile)
				# Due Date				
				my writeToFile(thisTodoProperties's due date, logbookFile, true)
				my nextColumn(logbookFile)
				# Completion Date
				my writeToFile(thisTodoCompletionDate, logbookFile, true)
				my nextColumn(logbookFile)
				# Name
				my writeToFile(thisTodoProperties's name, logbookFile, true)
				my nextColumn(logbookFile)
				# Area
				if thisTodoProperties's area is missing value then
					if thisTodoProperties's project is not missing value and area of thisTodoProperties's project is not missing value then
						my writeToFile(name of area of thisTodoProperties's project, logbookFile, true)
					else
						my writeToFile("missing value", logbookFile, true)
					end if
				else
					my writeToFile(name of thisTodoProperties's area, logbookFile, true)
				end if
				my nextColumn(logbookFile)
				# Project
				if thisTodoProperties's project is missing value then
					my writeToFile("missing value", logbookFile, true)
				else
					my writeToFile(name of thisTodoProperties's project, logbookFile, true)
				end if
				my nextColumn(logbookFile)
				# Tags
				my writeToFile(thisTodoProperties's tag names, logbookFile, true)
				my nextColumn(logbookFile)
				# Completion Month
				my writeToFile(year of thisTodoCompletionDate, logbookFile, true)
				my nextColumn(logbookFile)
				# Completion Month
				my writeToFile(month of thisTodoCompletionDate, logbookFile, true)
				my nextColumn(logbookFile)
				# Completion Weekday
				my writeToFile(weekday of thisTodoCompletionDate, logbookFile, true)
				my nextRow(logbookFile)
				set my counterOfExportedTodos to (my counterOfExportedTodos) + 1
			end if
		end repeat
	end tell
end if
#
#  Export Projects and Projects todos
#
if my exportProjects then
	# Set the header to Projects file
	my writeToFile("Creation Date,Due Date,Completion Date,Name,Area,Tags,Completion Year,Completion Month,Completion Weekday", projectsFile, false)
	my nextRow(projectsFile)
	# Start processing todos
	tell application "Things"
		# Get all projects
		set my listOfProjects to projects
		set my numberOfProjects to count my listOfProjects
		# Start processing projects
		repeat with i from 1 to my numberOfProjects
			# Get current projects's properties and completion date
			set thisProjectProperties to (my listOfProjects's item i)'s properties
			set thisProjectCompletionDate to thisProjectProperties's completion date
			if thisProjectProperties's status is completed and thisProjectCompletionDate is greater than my exportLast then
				# Creation Date
				my writeToFile(thisProjectProperties's creation date, projectsFile, true)
				my nextColumn(projectsFile)
				# Due Date
				my writeToFile(thisProjectProperties's due date, projectsFile, true)
				my nextColumn(projectsFile)
				# Completion Date
				my writeToFile(thisProjectCompletionDate, projectsFile, true)
				my nextColumn(projectsFile)
				# Name
				my writeToFile(thisProjectProperties's name, projectsFile, true)
				my nextColumn(projectsFile)
				# Area
				if thisProjectProperties's area is missing value then
					my writeToFile("missing value", projectsFile, true)
				else
					my writeToFile(name of thisProjectProperties's area, projectsFile, true)
				end if
				my nextColumn(projectsFile)
				# Tags
				my writeToFile(thisProjectProperties's tag names, projectsFile, true)
				my nextColumn(projectsFile)
				# Completion Month
				my writeToFile(year of thisProjectCompletionDate, projectsFile, true)
				my nextColumn(projectsFile)
				# Completion Month
				my writeToFile(month of thisProjectCompletionDate, projectsFile, true)
				my nextColumn(projectsFile)
				# Completion Weekday
				my writeToFile(weekday of thisProjectCompletionDate, projectsFile, true)
				my nextRow(projectsFile)
				set my counterOfExportedProjects to (my counterOfExportedProjects) + 1
			end if
		end repeat
	end tell
end if
#
# Export Areas
#
if my exportAreas then
	# Set the header
	my writeToFile("Name,Suspended,Tags", areasFile, false)
	my nextRow(areasFile)
	# Start processing areas
	tell application "Things"
		# Get all areas
		set my listOfAreas to areas
		set my numberOfAreas to count my listOfAreas
		repeat with i from 1 to my numberOfAreas
			# Get current area's properties
			set thisAreaProperties to (my listOfAreas's item i)'s properties
			# Name
			my writeToFile(thisAreaProperties's name, areasFile, true)
			my nextColumn(areasFile)
			# Suspended
			my writeToFile(thisAreaProperties's suspended, areasFile, true)
			my nextColumn(areasFile)
			# Tags
			my writeToFile(thisAreaProperties's tag names, areasFile, true)
			my nextRow(areasFile)
			set my counterOfExportedAreas to (my counterOfExportedAreas) + 1
		end repeat
	end tell
end if
#
# Count Loogbook items
#
tell application "Things"
	set numberOfLogbookItems to count of to dos of list "Logbook"
end tell
#
# Ending
#
if my playsound is true then do shell script "/usr/bin/afplay /System/Library/Sounds/Glass.aiff"
# Calculate time of the execution
set endTime to current date
set executionHours to (endTime - startTime) div hours
set executionMinutes to (endTime - startTime) div minutes
set executionSeconds to (endTime - startTime) - (executionHours * hours) - (executionMinutes * minutes)
set executionTime to executionHours & ":" & executionMinutes & ":" & executionSeconds
# Statistics
set finishMsg to "Todos: " & return & my counterOfExportedTodos & " completed exported from all " & my numberOfTodos & return
set finishMsg to finishMsg & return
set finishMsg to finishMsg & "Projects: " & return & my counterOfExportedProjects & " completed exported from all " & my numberOfProjects & return
set finishMsg to finishMsg & return
set finishMsg to finishMsg & "Areas: " & return & my counterOfExportedAreas & " exported from " & my numberOfAreas & return
set finishMsg to finishMsg & return
set finishMsg to finishMsg & "Logbook: " & numberOfLogbookItems & " items." & return
# Time of ececution
set finishMsg to finishMsg & return
set finishMsg to finishMsg & "Script starting: " & time string of startTime & return
set finishMsg to finishMsg & "Script ending:  " & time string of endTime & return
set finishMsg to finishMsg & "Time of execution: " & executionTime
display dialog (finishMsg)