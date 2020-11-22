Function GetPartsListPath(ByRef oPartsList As PartsList)
	Dim filepath As String
	filepath = oPartsList.ReferencedDocumentDescriptor.FullDocumentName
	'usage: GetPartsListPath(oPartsList)
	If Right(filepath,1) = ">" Then
		'find position of < character, then delete everything to the right.
		filepath = Left(filepath,InStr(1,filepath,"<",1)-1)
		End If
	Return filepath
End Function