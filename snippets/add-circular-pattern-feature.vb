'Source: https://github.com/InventorCode/iLogicRules
'Title: Add Circular Pattern Feature
'Author: nannerdw
'Description: Creates a circular pattern feature in a part file.
'
'CircularPatternFeatures.AddByDefinition in Inventor's API has a bug that causes the angle value (in radians) to be the same as the count value.
'AddCircularPatternFeature is provided as a workaround for that bug.

Option Explicit On

Private Sub Main
'Sub Main is provided as an example for calling AddCircularPatternFeature.
'It requires a part file containing two features, named Extrusion1 and Extrusion2
'It will create a circular pattern feature from those two features.
	
	'This rule must be called from a part document
	Dim doc As PartDocument = TryCast(ThisDoc.Document, PartDocument)
	If doc Is Nothing Then Exit Sub
	
	Dim compDef As PartComponentDefinition = doc.ComponentDefinition
	Dim oFeatures As PartFeatures = compDef.Features
	
	'In this example, a list of PartFeatures is passed to the AddCircularPatternFeature function.
	'Just one of these PartFeatures could be passed on its own instead, without the need for a list.
	Dim patternFeatures As New List(Of PartFeature)
	Try
		patternFeatures.Add(oFeatures("Extrusion1"))
		patternFeatures.Add(oFeatures("Extrusion2"))
	Catch ex As ArgumentException
		MessageBox.Show("Part must contain features named Extrusion1 and Extrusion2")
		Exit Sub
	End Try
	
	'The first three WorkAxes in a document are always the X,Y,Z origin axes.
	'You can reference them by name, like I did in patternFeatures above,
	'but remember that the names of the origin work features are editable.
	Dim patternAxis As WorkAxis = compDef.WorkAxes(2)
	
	Dim circPattern As CircularPatternFeature = AddCircularPatternFeature(
		ParentFeatures:=patternFeatures, 
		AxisEntity:=patternAxis, 
		NaturalAxisDirection:=True, 
		Count:=6,
		Angle:="360 deg",
		Name:="My Circular Pattern"
		)
End Sub


''' <summary>
''' Creates a circular pattern feature in a part file.
''' </summary>
''' 
''' <remarks>
''' CircularPatternFeatures.AddByDefinition In Inventor's API has a bug that causes the Angle value (in radians) to be the same as the Count value.
''' <see cref="AddCircularPatternFeature"/> is provided as a workaround for that bug.
''' <br/><br/>
''' Source: <seealso href="https://github.com/InventorCode/iLogicRules"/>
''' </remarks>
''' 
''' <param name="Angle">Can be a Double (in radians), or a parameter expression string.</param>
''' 
''' <param name="Count">Can be an Integer or a parameter expression string.</param>
''' 
''' <param name="ParentFeatures">Can be one of the following: 
''' a single object, 
''' an object that implements IEnumerable(Of Object), 
''' an ObjectCollection
''' </param>
'''
Private Function AddCircularPatternFeature(
	ParentFeatures As Object, 
	AxisEntity As Object, 
	NaturalAxisDirection As Boolean, 
	Count As Object, 
	Angle As Object,
	Optional FitWithinAngle As Boolean = True,
	Optional Name As String = ""
	) As CircularPatternFeature

	Dim objColl As ObjectCollection
	
	If TypeOf ParentFeatures Is ObjectCollection Then
		objColl = ParentFeatures
	Else
		objColl = ThisApplication.TransientObjects.CreateObjectCollection
		
		If TypeOf ParentFeatures Is IEnumerable(Of Object) Then
			For Each obj In ParentFeatures
				objColl.Add(obj)
			Next
		Else 'Assume ParentFeatures is a single object
			objColl.Add(ParentFeatures)
		End If
	End If
	
	Dim compDef As PartComponentDefinition = objColl(1).Parent
	Dim doc As PartDocument = compDef.Document
	Dim app As Inventor.Application = doc.Parent
	Dim oFeatures As PartFeatures = compDef.Features
	Dim circFeatures = oFeatures.CircularPatternFeatures
	
	Dim patternDef As CircularPatternFeatureDefinition = 
		circFeatures.CreateDefinition(objColl, AxisEntity, NaturalAxisDirection, Count, Angle, FitWithinAngle)
	
	Dim oFeature As CircularPatternFeature
	
	Dim oTransaction As Transaction
	
	Try
		oFeature = circFeatures.AddByDefinition(patternDef)
		
		'This will get merged into the "Create Circular Pattern Feature" transaction.
		oTransaction = app.TransactionManager.StartTransaction(doc, "Temp") 
		
		'CircularPatternFeatures.AddByDefinition in Inventor's API has a bug that causes the angle value (in radians) to be the same as the count value.
		'To work around this bug, the angle parameter has to be modified after the feature has been created.
		oFeature.Parameters.Item(1).Expression = Angle
		doc.Update2
		
		If Name <> "" Then oFeature.Name = Name
	
		oTransaction.End
		oTransaction.MergeWithPrevious = True
		
	Catch ex As Exception
		oTransaction.Abort
		Throw
		
	End Try
	
	Return oFeature
End Function
