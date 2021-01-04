'Source: https://github.com/InventorCode/iLogicRules

'Author: nannerdw

'Description: This is an example of how to use XML documentation comments in iLogic, which will show up in Intellisense.

'Begin typing an xml documentation comment by typing ''' at the start of a line directly above the class/module/function/etc to be commented.

'iLogic only supports the <summary> and <code> tags.
'<br/> is not supported for newlines; instead, use <code/>
'
'These tags cannot be used inside classes that are declared Private.
'
'iLogic will also recognize <summary> tags that are inside other "Straight VB code" rules,
'and in xml documentation files that are included with dlls that are added to iLogic with AddReference.


''' <summary>
''' This is a summary for Class RuleMain.<code/>
''' This is a multi-line summary.
''' </summary>
Class RuleMain
	''' <summary>This is a summary for the private member variable _doc</summary>
	Private _doc As Document = Nothing
	
	''' <summary>This is a summary for Sub Main</summary>
	Private Sub Main
		_doc = ThisDoc.Document
		
		Dim cls1 As New Class1(Prop1Value:=False)
		var1 = cls1.Var1
		prop1 = cls1.Prop1
	End Sub
	
	''' <summary>
	''' This is a summary for Function Func1
	''' </summary>
	Private Function Func1() As Boolean
		
	End Function
End Class


''' <summary>
''' This is a summary for Class1
''' </summary>
Class Class1
	
	''' <summary>
	''' This is a summary for Property Prop1
	''' </summary>
	Public Property Prop1 As Boolean
		
	''' <summary>
	''' This is a summary for the public member variable Var1
	''' </summary>
	Public Var1 As Boolean = True
	
	''' <summary>
	''' Initializes a new instance of Class1
	''' 
	''' <code>
	''' 'This is a code example. Its text will be formatted differently in Intellisense.
	''' Dim cls1 As New Class1()
	''' </code>
	''' 
	''' Prop1Value = True by default
	''' </summary>
	Sub New(Optional Prop1Value As Boolean = True)
		Me.Prop1 = Prop1Value
	End Sub
	
End Class