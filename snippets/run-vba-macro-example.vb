'Source: https://github.com/InventorCode/iLogicRules
'Title: Run VBA Macro - Example
'Author: nannerdw
'Description: Example of how to run a public VBA sub (with or without arguments) from iLogic
'This example uses the "Log" macro from LogFunctions.bas to output a test string to the iLogic logger.

'Example of running a VBA sub that accepts no arguments:
	InventorVb.RunMacro("projectName", "moduleName", "macroName")
	

'Example with arguments:

'In VBA, I have a this line in my ApplicationProject.LogFunctions module:
' Public Sub Log(message As String, Optional level = 3)

'This sub definition contains two arguments, one that is optional with a default value.
'InventorVb.RunMacro does not respect this default value, but instead 
'passes in the type's default value for any arguments that are left blank,
'regardless of whether they are defined as Optional in VBA.

'For example:
	InventorVb.RunMacro("ApplicationProject", "LogFunctions", "Log")
'would be equivalent to typing this in VBA:
' Log message:="", level:=0

'Instead, you should treat all arguments as if they were required when using InventorVb.RunMacro:
	InventorVb.RunMacro("ApplicationProject", "LogFunctions", "Log", "this is a test message from VBA", 3)
