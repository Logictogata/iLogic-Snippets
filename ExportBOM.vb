' This code exports the "Parts Only" tab in the BOM to an excel file called excelBOM.xlsx

Sub Main()

	ExportBOM

End Sub


Function ExportBOM()
	
	oPath = ThisDoc.WorkspacePath()
	
	Dim oDoc As AssemblyDocument
	oDoc = ThisApplication.ActiveDocument
	
	Dim oBOM As BOM
	oBOM = oDoc.ComponentDefinition.BOM
	oBOM.PartsOnlyViewEnabled = True
	
	' check for occurences, part or assembly
	Dim oComp As ComponentOccurrence
	Dim oComps As ComponentOccurrences
	oComps = ThisDoc.Document.ComponentDefinition.Occurrences
	
	Dim x As Integer = oComps.Count()
	' MessageBox.Show("oComps.count() " & x, "ExportBOM")
	
	If oComps.Count() <= 0 Then
		MessageBox.Show("No components in assembly.", "ExportBOM")
		Exit Function
	Else
		' keep on going partner
	End If
	
	' define the Assembly Component Definition
	Dim oAsmCompDef As ComponentDefinition
	oAsmCompDef = ThisDoc.Document.ComponentDefinition

	' Store active LOD
	Dim tempLOD As LevelOfDetailRepresentation
	tempLOD = oAsmCompDef.RepresentationsManager.ActiveLevelOfDetailRepresentation()
	
	' Activate master LOD
	Dim oLOD As LevelOfDetailRepresentation
    oLOD = oAsmCompDef.RepresentationsManager.LevelOfDetailRepresentations.Item("Master").Activate(True)
	
	' define BOM view
	Dim oPartsOnlyBOMView As BOMView
	oPartsOnlyBOMView = oBOM.BOMViews.Item("Parts Only")
	
	' oPartsOnlyBOMView.Export (oPath & "/" & TITLE, kMicrosoftExcelFormat)
	oPartsOnlyBOMView.Export (oPath & "/" & "excelBOM", kMicrosoftExcelFormat)
	
	' Activate CUSTOM LOD
	oLOD = tempLOD

End Function
