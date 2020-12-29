'Source: https://github.com/InventorCode/iLogicRules
'Author: nannerdw
'Description: This is an example of how to pass RuleArguments to and from an iLogic rule.
'Call this rule using rule-arguments-example-calling-rule.vb

Private Sub Main
	If Not RuleArguments.Exists("x") Then
		MessageBox.Show(text:="The RuleArgument ""x"" must be passed to this rule.", caption:=iLogicVb.RuleName)
		Exit Sub
	End If
	
	Dim x As Integer = RuleArguments("x")
	
	RuleArguments.Arguments.Value("ReturnValue") = x + 1
End Sub
	
'This sub contains info about how RuleArguments work
Private Sub RuleArguments_Info
	'RuleArguments is a wrapper for a NameValueMap, and is used for passing objects between iLogic rules.
	'The objects contained in a NameValueMap are late-bound, which means that RuleArguments are weakly typed.
	
	'When you pass a NameValueMap into RunRule or RunExternalRule,
	'that NameValueMap object will be stored inside the called rule's RuleArguments.
	
	'Calling rule example:
	'------------------
		Dim map As Inventor.NameValueMap = ThisApplication.TransientObjects.CreateNameValueMap()
		map.Add("Arg1", "Arg1Value")
		iLogicVb.RunExternalRule("ruleName", map)
	'------------------
	
	'That NameValueMap object can then be accessed directly inside the called rule like this:
	Dim nvm As NameValueMap = RuleArguments.Arguments
	
	'You can read a value from RuleArguments with:
	arg = RuleArguments("Arg1") ', which is shorthand for:
	arg = RuleArguments.Arguments.Item("Arg1")
	
	'NameValueMap.Item and NameValueMap.Value refer to the same object, however:
	'NameValueMap.Item is read-only
	'NameValueMap.Value is read/write
	
	'Since Value is read/write, it can be used to return objects back to the calling rule:
	RuleArguments.Arguments.Value("ReturnValue1") = "This is a return value"
End Sub
