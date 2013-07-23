Things Statistics
=================

Contents
--------

### Main Script
*Things-ExportToNumbers*
An AppleScript for exporting completed todos, completed projects and areas from Things application into the Numbers document "Things Statistics".

### File to calc and view your stats
*Things Statistics*
An Apple Numbers documents with predefined formulas and charts for calculate and review statistics of completed todos.

### Supplementary scripts
*Things-DeleteLogbookItemsOlderDate*
An AppleScript for deleting all items (completed todos and projects) from Logbook list of Things application that older than specific date.

*Things-ExportToCSV*
If you want export data from things into CSV file than you can use this script.

Notes
-----

- Tested with Things 2.2.1 and OS X Mountain Lion 10.8.4
- Large amount (over 1000) of todos (include completed) may lead to bad perfomance such as a freezes of the Number or the Things apps and time of execution is about 1-2 hours.
- If you experience perfomance problems you can set earlier "exportLast" date or delete old items by "Things-DeleteLogbookItemsOlderDate" script.

How to use
----------

1. Start Things app and open "Things Statistics" file in Numbers app.
2. Open "Things-ExportToNumbers" script in AppleScript Editor (pre-installed on all Mac OS X systems) and add modifications in "Config" sections if required.
3. Run the sript by "Run" button.
4. Modify "Things Statistics" file to suit your needs (rename Areas and so on).

To Do
------

1. Add ability to add Areas into "Things Statistics" file by script but not manually.
2. To improve perfomance and save old data don't delete all items from "Things Statistics" file if they older than "exportLast" property. May be add new script to move old items to "historical" sheet or table of "Things Statistics" for the purpose of don't check every item while delete process is executing because that can add penalty to perfomance.
3. Make tags tab delimited by heirarchy.

Copyright & License
-------------------

Things Statistics is released under the MIT license:

- http://www.opensource.org/licenses/MIT

The MIT License (MIT)

Copyright (c) 2013 ixrevo (Alexander Sokol http://ixrevo.ru)

Permission is hereby granted, free of charge, to any person obtaining a copy of
this software and associated documentation files (the "Software"), to deal in
the Software without restriction, including without limitation the rights to
use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of
the Software, and to permit persons to whom the Software is furnished to do so,
subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR
COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN
CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.