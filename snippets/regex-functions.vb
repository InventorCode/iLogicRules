'test for:
'
'Sheet number like 100-(500)A: "\-(\d{3})"
'Task number ion drawing sheet, like (100)-500A: "^(\d{3})\-.{4}"
'Units, like 12.145 (in): "([A-Za-z]{1,4})"
'Number with possible decimal point: "(\d+(?:\.\d*)?|\.\d+)"
'A stock length, e.g. Steel Box tube, 2" x 3" - (24)': "\s?-\s(\d{1,3})(?=\')"
'
'
'
'
'


'##########################
'###   Regex Routines   ###
'##########################
'
Sub Main()
	Dim temp1 As String
	Dim temp2 As String

'Return the following: "Steel Channel, C 3x5.3"
	temp1 = ReturnNotRegex("steel Channel, C 3x5.3 - 21', AISC A36", "\s?-\s(\d{1,3})(?=\').*",,True)
	MsgBox(temp1)
'Return the following: "21"
	temp2 = ReturnRegex("Steel Channel, C 3x5.3 - 21', AISC A36", "\s?-\s(\d{1,3})(?=\')",1)
	MsgBox(temp2)
'Return the following: "21"
	temp2 = ReturnRegex("Steel Channel, C 3x5.3 - 21', AISC A36", "(\d)",1,True)
	MsgBox(temp2)

End Sub

'############################
'###   ReturnNotRegex   ###
'############################
' rev 1.1
' MDJ

Function ReturnNotRegex(strIn As String, sPattern As String, Optional boolGlobal As Boolean = False, Optional boolIgnoreCase As Boolean = False) As String
' Return string patterns based on a regex match
' ReturnNotRegex will return the original string minus the matched pattern.  The matched pattern ignores groups.
'
'Syntax:
'	ReturnNotRegex(input string to test, regex match pattern as string)
'
' Returns:
'	A string.
'
' Usage:
' ReturnNotRegex("A string to test for a match.", "\s(to te)st"
'	---> returns: "A string for a match"

	Dim objRegex
    objRegex = CreateObject("vbscript.regexp")
    With objRegex
     .Global = boolGlobal
     .IgnoreCase = boolIgnoreCase
     .Pattern = sPattern
		If .test(strIn) Then
	    	ReturnNotRegex = .Replace(strIn, vbNullString)
		Else
'			MsgBox("match not found")
			ReturnNotRegex = strIn
		End If
    End With

End Function

'#########################
'###   ReturnRegex   ###
'#########################
' rev 1.1
' MDJ

Function ReturnRegex(strIn As String, sPattern As String, i As Integer, Optional boolGlobal As Boolean = False, Optional boolIgnoreCase As Boolean = False) As String
' Return string patterns based on a regex match
' ReturnRegex will return the specified group of matched items from a string.

'Syntax:
'	ReturnRegex(input string to test, regex match pattern as string, group [e.g. $1, $2] to return [1 for first, 2 for second, 0 for all, etc...])
'
' Returns:
'	A string.
'
' Usage:
'ReturnRegex("A string to test for a match.", "ma(tch)", 1)
'	---> returns: "tch"
'ReturnRegex("A string to test for a match.", "ma(tch)", 0)
'	---> returns: "match"

	Dim objRegex As Object
    objRegex = CreateObject("vbscript.regexp")
    Dim objMatches As Object
	Dim objMatch As Object
	Dim strMatch As String
	
    With objRegex
     .Global = boolGlobal
	 .IgnoreCase = boolIgnoreCase
     .Pattern = sPattern
	End With
	
		'Test to make sure we can return a match in the supplied test string
		If objRegex.test(strIn) Then
			
			'if the return group was specified as 1 or greater, provide the group number requested
			If i >= 1 Then
				objMatches = objRegex.Execute(strIn)
				For Each objMatch In objMatches
					strMatch = strMatch & objMatch.submatches(i-1)
				Next
				
			'if the return group was specified as 0 or less, the user just wants the overall match, no groups used.
			ElseIf i = 0 Or i < 0
				objMatches = objRegex.Execute(strIn)
				For Counter = 0 To objMatches.Count - 1
    				msgbox(objMatches.Count)
					strMatch = strMatch & objMatches(Counter).Value
				Next	
			End If
		Else
'			MsgBox("match not found")
			'Do nothing
		End If
	ReturnRegex = strMatch

End Function


'#####################
'###   TestRegex   ###
'#####################
' rev 1.1
' MDJ

Function TestRegex(strIn As String, sPattern As String, Optional boolGlobal As Boolean = False, Optional boolIgnoreCase As Boolean = False) As Boolean
' Return string patterns based on a regex match
' ReturnNotRegex will return a True/False based on a regex match.
'
'Syntax:
'	TestRegex(input string to test, regex match pattern as string)
'
' Returns:
'	A Boolean
'
' Usage:
' TestRegex("A string to test for a match.", "\s(to te)st")
'	---> returns: True

	Dim objRegex
    objRegex = CreateObject("vbscript.regexp")
    With objRegex
     .Global = boolGlobal
     .IgnoreCase = boolIgnoreCase
     .Pattern = sPattern
		If .test(strIn) Then
	    	TestRegex = True
		Else
'			MsgBox("match not found")
			TestRegex = False
		End If
    End With

End Function
