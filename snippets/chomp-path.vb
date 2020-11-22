Function ChompPath(ByVal stringTemp As String)

	
	'find position of / character, then delete everything to the right.
		stringTemp = Left(stringTemp,((Len_stringTemp)-InStr(1,StrReverse(stringTemp),"/",1)))
	Return stringTemp
End Function	


'[TODO: make this into a generic chomp function, use split and join?]