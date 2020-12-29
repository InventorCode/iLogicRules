'Source: https://github.com/InventorCode/iLogicRules
'Title: Open Selected Derived Reference
'Author: nannerdw
'Description:
'If the selected object (vertex, edge, face, solid body, sketch, sketch block, workfeature, etc) was created by a "Derived Component" feature,
'then the base component of that feature will be opened (equivalent to right clicking the feature and selecting "Open Base Component").

Option Explicit On
Imports Inventor.ObjectTypeEnum

Dim doc As PartDocument = TryCast(ThisDoc.Document,PartDocument)
If doc Is Nothing Then Exit Sub

Dim refFile As ReferencedFileDescriptor

For Each selectedObj As Object In doc.SelectSet
    If selectedObj Is Nothing Then Continue For
	
	With selectedObj
		Select Case .Type

		Case kWorkPointObject, _ '(includes UCS workfeatures)
			 kWorkAxisObject,  _ '..
			 kWorkPlaneObject, _ '..
			 kReferenceFeatureObject, _
			 kPlanarSketchObject, _
			 kSketch3DObject, _
			 kSketchBlockDefinitionObject
			If .ReferenceComponent IsNot Nothing Then
				refFile = .ReferenceComponent.ReferencedFile
			End If

		Case kSketchBlockObject 'instance in sketch (incl. nested blocks)
			If Not .Definition.ReferenceComponent Is Nothing Then
				refFile = .Definition.ReferenceComponent.ReferencedFile
			End If
				
		Case kDerivedPartComponentObject, _
			 kDerivedAssemblyComponentObject
			refFile = .ReferencedFile
			
		Case kSurfaceBodyObject 'Solid Body
			refFile = .CreatedByFeature.ReferenceComponent.ReferencedFile
			
		Case kWorkSurfaceObject 'Surface Body
			If .SurfaceBodies(1).CreatedByFeature.Type = kReferenceFeatureObject Then
				refFile = .SurfaceBodies(1).CreatedByFeature.ReferenceComponent.ReferencedFile
			End If
		
		Case Else
			If .parent IsNot Nothing Then 'For example, PartComponentDefinition has no parent.
				Select Case .parent.Type
				
				Case kSketchBlockDefinitionObject, _
					 kPlanarSketchObject, _
					 kSketch3DObject
					If .parent.ReferenceComponent IsNot Nothing Then
						refFile = .parent.ReferenceComponent.ReferencedFile
					End If 
				
				Case kSurfaceBodyObject 'Solid Body
					If .parent.CreatedByFeature.Type = kReferenceFeatureObject Then
						refFile = .parent.CreatedByFeature.ReferenceComponent.ReferencedFile
					End If
					
				End Select 'Case .parent.Type
			End If '.parent IsNot Nothing
		End Select 'Case .Type
	End With 'selectedObj
	
	'Open the reference file
	Try
    	If refFile IsNot Nothing Then ThisApplication.Documents.Open (refFile.fullFileName)
	Catch ex As Exception
		MessageBox.Show(refFile.fullFileName & vbNewLine & vbNewLine & _
		"Typical cause(s): File may have been renamed or relocated.", "File could not be opened")
	End Try
Next selectedObj

