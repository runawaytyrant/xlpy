# Changelog

v2.0
- Implemented a GUI.
- Switched to a python file with no console.

v3.0
- Changed accepted file types from .csv files to .xls, .xlsx and .xlsm files.
- Data display is now done on the main window rather than on a browser tab via an HTML file.
- Added filtering function on the IDs column, application will now only show students who's IDs start with a certain number.
- Added a label that displays the number of students that met the condition.

v3.1
- Modified the GUI.
- Made the file browsing and filtering into different functions, each tied to their own button.
- Modified all text shown on the app's text widget to eliminate unnecessary spaces.
- Added function for generating a graph out of example data.

v3.2
- Added more possible column combinations and modified the corresponding error message.
- Modified the graph plotting function so it uses data from the excel file and added related error messages.
- Fixed certain file loading related error messages not appearing.
- Added dynamic title.
- Added value labels to the bar charts.

v3.3
- Added percentage calculation of y2/y3 according to y1.
- Changed value labels from numbers to percentages and centered them.
- y1 bar is no longer plotted.
- Reversed the order of the chart legend.

v3.3.1
- Fixed main window size.
- Tested and added new accepted file types: .xlt, .xltx and .xlsb files.
