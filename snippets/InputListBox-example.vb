'Source: https://github.com/InventorCode/iLogicRules
'Author: nannerdw
'Last Modified Date: 27 Nov, 2020
'Description: Example of how to use iLogic's built-in InputListBox dialog to prompt the user to select from a list of options
'Its fully-qualified name is Autodesk.iLogic.Runtime.RunDialogs.InputListBox

'Example of inline array declaration:
Dim listItems() As String = New String() {"1/2 in", "1/4 in", "1/8 in", "1/16 in", "1/32 in"}

Dim result As String = InputListBox(
    Prompt := "Pick a value.",
    ListItems := listItems, 
    DefaultValue := listItems(0), 
    Title := "InputListBox Example",
    ListName := "Here are your options:",
	Width := 200 'I changed this from its default so the title wouldn't get cut off
)

MessageBox.Show(If(result = "", "You closed the InputListBox.", "You picked " & result), "InputListBox Result")
