'Source: https://github.com/InventorCode/iLogicRules
'Title: Open Project Location
'Author: nannerdw
'Description: Opens the folder containing the active .ipj project file

Process.Start(System.IO.Path.GetDirectoryName(ThisApplication.FileLocations.FileLocationsFile))