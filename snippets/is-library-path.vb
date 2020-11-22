'Test for library paths:
Function IsLibraryPath()

	oPathname = ThisDoc.Document.FullFileName
	
	Dim oProjectMgr As DesignProjectManager
	oProjectMgr = ThisApplication.DesignProjectManager
	
	' Get the active project
	Dim oProject As DesignProject
	oProject = oProjectMgr.ActiveDesignProject
	
	Dim oLibraryPaths As ProjectPaths
	oLibraryPaths = oProject.LibraryPaths
	
	' look at all library paths
	Dim oLibraryPath As ProjectPath
	For Each oLibraryPath In oLibraryPaths
		oPath = oLibraryPath.Path
		'trim off left most char
		oPath = Right(oPath,Len(oPath)-1)
		If oPathname.Contains(oPath)
			'do something
		Else		
			'do something else
		End If
	Next
	
	'or use 'Contains'
	
	oPathname = ThisDoc.Document.FullFileName
	
	If oPathname.Contains("Libraries") Then
		'do something
		Return True
	Else
		'do something else
		Return False
	End If

End Function



