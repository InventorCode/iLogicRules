'[TODO: rewrite to be generic for any input object]

Function GetViewPath(ByRef oView As DrawingView)
	Dim filepath As String
	filepath = oView.ReferencedDocumentDescriptor.FullDocumentName
	'usage: GetViewPath(oView)
	If Right(filepath,1) = ">" Then
		'find position of < character, then delete everything to the right.
		filepath = Left(filepath,InStr(1,filepath,"<",1)-1)
		End If
	Return filepath
End Function