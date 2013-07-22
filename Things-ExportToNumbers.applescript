#
# Things-ExportToNumbers.applescript
#
# An AppleScript for exporting completed todos, completed projects and areas from Things application into the Numbers document "Things Statistics".
#
# Things Statistics is released under the MIT license:
# Copyright (c) 2013 @ixrevo (Alexander Sokol http://ixrevo.ru)
#

##########
# Config #
##########

#
#
property appendData : false
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
#Export areas?
property exportAreas : true
#
# Export data for last days or weeks
property exportLast : (current date) - (2 * 365 * days + 0 * weeks)
#
# Play sound after script's completion
property playsound : true


############
# Handlers #
############

#
# Convert boolean value to humanfriendly value 'Yes' or 'No'
#
# bolValue    boolean    boolean value to convert
#
# return      text       Yes, No or Undefined
#
on humanBoolean(bolValue)
	if bolValue then
		"Yes"
	else if not bolValue then
		"No"
	else
		return "Undefined"
	end if
end humanBoolean

#
# Get the value of a cell in Numbers' document.
#
# documentName    string or integer    name or index of the target document
# sheetName    string or integer    name or index of the target sheet
# tableName    string or integer    name or index of the target table
# columnName      string or integer    name or index of the target column
# rowNumber    integer              row number of the target cell
#
# returns      Value of the cell
#
on valueOfCell(documentName, sheetName, tableName, columnName, rowNumber)
	tell application "Numbers"
		tell document documentName
			tell sheet sheetName
				tell table tableName
					tell cell rowNumber of column columnName
						set theValue to value
					end tell
				end tell
			end tell
		end tell
	end tell
	try
		set theValue to (theValue as date) - (time to GMT) -- fix for problem with time zone shift
	end try
	return theValue
end valueOfCell

#
# Set the value of a cell in Numbers' document.
#
# cellValue       anything             cell value, using text is strongly recommended
# documentName    string or integer    name or index of the target document
# sheetName       string or integer    name or index of the target sheet
# tableName       string or integer    name or index of the target table
# columnName      string or integer    name or index of the target column
# rowNumber       integer              row number of the target cell
#
# returns     Nothing
#
on setValueOfCell(cellValue, documentName, sheetName, tableName, columnName, rowNumber)
	if class of cellValue is not text then set cellValue to cellValue as text
	tell application "Numbers"
		tell document documentName
			tell sheet sheetName
				tell table tableName
					tell cell rowNumber of column columnName
						set value to cellValue
					end tell
				end tell
			end tell
		end tell
	end tell
end setValueOfCell

#
# Get the column count of a table in Numbers' document.
#
# documentName    string or integer    name or index of the target document
# sheetName       string or integer    name or index of the target sheet
# tableName       string or integer    name or index of the target table
#
# returns     number of columns
#
on columnCountOfTable(documentName, sheetName, tableName)
	tell application "Numbers"
		tell document documentName
			tell sheet sheetName
				tell table tableName
					set rowCount to column count
				end tell
			end tell
		end tell
	end tell
	return rowCount
end columnCountOfTable

#
# Get the row count of a table in Numbers' document.
#
# documentName    string or integer    name or index of the target document
# sheetName       string or integer    name or index of the target sheet
# tableName       string or integer    name or index of the target table
#
# returns         integer              number of rows
#
on rowCountOfTable(documentName, sheetName, tableName)
	tell application "Numbers"
		tell document documentName
			tell sheet sheetName
				tell table tableName
					set rowCount to row count
				end tell
			end tell
		end tell
	end tell
	return rowCount
end rowCountOfTable

#
# Set the row count of a table in Numbers' document.
#
# rowCount        integer              number of rows to set
# documentName    string or integer    name or index of the target document
# sheetName       string or integer    name or index of the target sheet
# tableName       string or integer    name or index of the target table
#
# returns         nothing
#
on setRowCountOfTable(rowCount, documentName, sheetName, tableName)
	tell application "Numbers"
		tell document documentName
			tell sheet sheetName
				tell table tableName
					set row count to rowCount
				end tell
			end tell
		end tell
	end tell
end setRowCountOfTable

#
# Delete the rows between start row and end row in a table in Numbers' document.
# Start and end rows also will be deleted.
#
# startRow        integer              start of the row range to delete
# endRow          integer              end of the row range to delete
# documentName    string or integer    name or index of the target document
# sheetName       string or integer    name or index of the target sheet
# tableName       string or integer    name or index of the target table
#
# returns         nothing
#
on deleteRowsOfTable(startRow, endRow, documentName, sheetName, tableName)
	tell application "Numbers"
		tell document documentName
			tell sheet sheetName
				tell table tableName
					delete (rows startRow thru endRow)
				end tell
			end tell
		end tell
	end tell
end deleteRowsOfTable

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
# Export completed todos
#
if exportTodos then
	tell application "Things"
		# Get list of all todos
		set my listOfTodos to to dos
		set my numberOfTodos to count my listOfTodos
		# Get number of rows of the table (also number of the last row)
		set numberOfTheLastRow to my rowCountOfTable("Things Statistics", "Todos", "Completed Todos")
		
		if not my appendData then
			# Reset table
			my deleteRowsOfTable(3, numberOfTheLastRow, "Things Statistics", "Todos", "Completed Todos")
			set numberOfTheLastRow to 1
		end if
		# Process todos
		repeat with i from 1 to my numberOfTodos
			# Get current todo's properties and completion date
			set thisTodoProperties to (my listOfTodos's item i)'s properties
			set thisTodoCompletionDate to thisTodoProperties's completion date
			if (thisTodoProperties's class is not project) and (thisTodoProperties's status is completed) and (thisTodoCompletionDate is greater than exportLast) then
				# Add new row below
				set numberOfTheLastRow to numberOfTheLastRow + 1
				my setRowCountOfTable(numberOfTheLastRow, "Things Statistics", "Todos", "Completed Todos")
				
				# Creation Date
				my setValueOfCell(thisTodoProperties's creation date, "Things Statistics", "Todos", "Completed Todos", "Creation Date", numberOfTheLastRow)
				# Due Date
				my setValueOfCell(thisTodoProperties's due date, "Things Statistics", "Todos", "Completed Todos", "Due Date", numberOfTheLastRow)
				# Completion Date
				my setValueOfCell(thisTodoCompletionDate, "Things Statistics", "Todos", "Completed Todos", "Completion Date", numberOfTheLastRow)
				# Name
				my setValueOfCell(thisTodoProperties's name, "Things Statistics", "Todos", "Completed Todos", "Name", numberOfTheLastRow)
				# Area
				if thisTodoProperties's area is missing value then
					if thisTodoProperties's project is not missing value and area of thisTodoProperties's project is not missing value then
						my setValueOfCell(name of area of thisTodoProperties's project, "Things Statistics", "Todos", "Completed Todos", "Area", numberOfTheLastRow)
					else
						my setValueOfCell("None", "Things Statistics", "Todos", "Completed Todos", "Area", numberOfTheLastRow)
					end if
				else
					my setValueOfCell(name of thisTodoProperties's area, "Things Statistics", "Todos", "Completed Todos", "Area", numberOfTheLastRow)
				end if
				# Project
				if thisTodoProperties's project is missing value then
					my setValueOfCell("None", "Things Statistics", "Todos", "Completed Todos", "Project", numberOfTheLastRow)
				else
					my setValueOfCell(name of thisTodoProperties's project, "Things Statistics", "Todos", "Completed Todos", "Project", numberOfTheLastRow)
				end if
				# Tags
				my setValueOfCell(thisTodoProperties's tag names, "Things Statistics", "Todos", "Completed Todos", "Tags", numberOfTheLastRow)
				# Completion Year
				my setValueOfCell(year of thisTodoCompletionDate, "Things Statistics", "Todos", "Completed Todos", "Completion Year", numberOfTheLastRow)
				# Completion Month
				my setValueOfCell(month of thisTodoCompletionDate, "Things Statistics", "Todos", "Completed Todos", "Completion Month", numberOfTheLastRow)
				# Completion Weekday
				my setValueOfCell(weekday of thisTodoCompletionDate, "Things Statistics", "Todos", "Completed Todos", "Completion Weekday", numberOfTheLastRow)
				# Completion Week Number
				my setValueOfCell("=WEEKNUM(Completion Date;2)", "Things Statistics", "Todos", "Completed Todos", "Completion Week Number", numberOfTheLastRow)
				
				set my counterOfExportedTodos to (my counterOfExportedTodos) + 1
			end if
		end repeat
	end tell
end if
#
#  Export Projects and Projects todos
#
if my exportProjects then
	tell application "Things"
		# Get all projects
		set my listOfProjects to projects
		set my numberOfProjects to count my listOfProjects
		# Get number of rows of the table (also number of the last row)
		set numberOfTheLastRow to my rowCountOfTable("Things Statistics", "Projects", "Completed Projects")
		
		if not my appendData then
			# Reset table
			my deleteRowsOfTable(3, numberOfTheLastRow, "Things Statistics", "Projects", "Completed Projects")
			set numberOfTheLastRow to 1
		end if
		# Start processing projects
		repeat with i from 1 to my numberOfProjects
			# Get current projects's properties and completion date
			set thisProjectProperties to (my listOfProjects's item i)'s properties
			set thisProjectCompletionDate to thisProjectProperties's completion date
			
			if thisProjectProperties's status is completed and thisProjectCompletionDate is greater than exportLast then
				# Add new row below
				set numberOfTheLastRow to numberOfTheLastRow + 1
				my setRowCountOfTable(numberOfTheLastRow, "Things Statistics", "Projects", "Completed Projects")
				
				# Creation Date
				my setValueOfCell(thisProjectProperties's creation date, "Things Statistics", "Projects", "Completed Projects", "Creation Date", numberOfTheLastRow)
				# Due Date
				my setValueOfCell(thisProjectProperties's due date, "Things Statistics", "Projects", "Completed Projects", "Due Date", numberOfTheLastRow)
				# Completion Date
				my setValueOfCell(thisProjectCompletionDate, "Things Statistics", "Projects", "Completed Projects", "Completion Date", numberOfTheLastRow)
				# Name
				my setValueOfCell(thisProjectProperties's name, "Things Statistics", "Projects", "Completed Projects", "Name", numberOfTheLastRow)
				# Area
				if thisProjectProperties's area is missing value then
					my setValueOfCell("None", "Things Statistics", "Projects", "Completed Projects", "Area", numberOfTheLastRow)
				else
					my setValueOfCell(name of thisProjectProperties's area, "Things Statistics", "Projects", "Completed Projects", "Area", numberOfTheLastRow)
				end if
				# Tags
				my setValueOfCell(thisProjectProperties's tag names, "Things Statistics", "Projects", "Completed Projects", "Tags", numberOfTheLastRow)
				# Completion Year
				my setValueOfCell(year of thisProjectCompletionDate, "Things Statistics", "Projects", "Completed Projects", "Completion Year", numberOfTheLastRow)
				# Completion Month
				my setValueOfCell(month of thisProjectCompletionDate, "Things Statistics", "Projects", "Completed Projects", "Completion Month", numberOfTheLastRow)
				# Completion Weekday
				my setValueOfCell(weekday of thisProjectCompletionDate, "Things Statistics", "Projects", "Completed Projects", "Completion Weekday", numberOfTheLastRow)
				# Completion Week Number
				my setValueOfCell("=WEEKNUM(Completion Date;2)", "Things Statistics", "Projects", "Completed Projects", "Completion Week Number", numberOfTheLastRow)
				
				set my counterOfExportedProjects to (my counterOfExportedProjects) + 1
			end if
		end repeat
	end tell
end if
#
# Export Areas
#
if my exportAreas then
	tell application "Things"
		# Get all areas
		set my listOfAreas to areas
		set my numberOfAreas to count my listOfAreas
		
		# Get number of rows of the table (also number of the last row)
		set numberOfTheLastRow to my rowCountOfTable("Things Statistics", "Areas", "Areas")
		
		if not my appendData then
			# Reset table
			my deleteRowsOfTable(3, numberOfTheLastRow, "Things Statistics", "Areas", "Areas")
			set numberOfTheLastRow to 1
		end if
		
		repeat with i from 1 to my numberOfAreas
			# Get current area's properties
			set thisAreaProperties to (my listOfAreas's item i)'s properties
			
			# Add new row below
			set numberOfTheLastRow to numberOfTheLastRow + 1
			my setRowCountOfTable(numberOfTheLastRow, "Things Statistics", "Areas", "Areas")
			
			# Name
			my setValueOfCell(thisAreaProperties's name, "Things Statistics", "Areas", "Areas", "Name", numberOfTheLastRow)
			# Suspended
			my setValueOfCell(my humanBoolean(thisAreaProperties's suspended), "Things Statistics", "Areas", "Areas", "Suspended", numberOfTheLastRow)
			# Tags
			my setValueOfCell(thisAreaProperties's tag names, "Things Statistics", "Areas", "Areas", "Tags", numberOfTheLastRow)
			
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