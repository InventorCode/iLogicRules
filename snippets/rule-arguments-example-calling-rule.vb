'Source: https://github.com/InventorCode/iLogicRules
'Author: nannerdw
'Description: This is an example of how to pass RuleArguments to and from an iLogic rule.
'This rule calls rule-arguments-example.vb

Dim x As Integer = CInt(InputBox(Prompt :="Enter an integer", Title :="Input"))

Dim map As Inventor.NameValueMap = ThisApplication.TransientObjects.CreateNameValueMap()
map.Add("x", x)

iLogicVb.RunExternalRule("rule-arguments-example.vb", map)

MessageBox.Show(map.Item("ReturnValue"))