Imports System.Collections.Generic
Imports System.IO

'code taken and modified from:
'http://adndevblog.typepad.com/manufacturing/2015/01/add-external-ilogic-rule-to-event-trigger.html

Sub Main

	msgbox("This rule is obsolete, and may not be used.  ")
	' Dim TemplateFilePath as String = "C:\_vaultWIP\Designs\Templates\Universal\US_FLORIDA\_USF_STANDARD TEMPLATE_SIZE B_11x17.idw"
	' Dim TBName as String = "UNIVERSAL FLORIDA MAIN SHEET"
	' Dim BdrName as String = "UNIVERSAL BORDER - FULL"
	' Dim oTemplateDoc As Document
	' Dim oTemplate As String 
	' Dim oDoc as Document = ThisApplication.ActiveDocument
	
	' If oDoc.DocumentType = kDrawingDocumentObject Then
		' oTemplate = TemplateFilePath 'file to copy events from
	' 'ElseIf oDoc.DocumentType = kPartDocumentObject Then
		' 'oTemplate = "\\server\departments\CAD\Inventor\Templates\Standard.ipt"   'file to copy events from
	' Else
		' 'oTemplate = "\\server\departments\CAD\Inventor\Templates\Standard.iam"   'file to copy events from
		' MsgBox("Must run from drawing")
		' Exit Sub
	' End If
	
	' 'ADDED For SMLYDK to create -FL documents
	' oFileName = DocumentFileName(oDoc)
	' FLPath = InputBox("Please enter the file path you would like to save to: ", "Save Path", _ 
		' Left(oDoc.FullFileName, oDoc.FullFileName.Length - DocumentFileNameExt(oDoc).Length))

	' If FLPath = vbCancel Then
		' MsgBox("You did not make a selection.  Exiting")
		' Exit Sub
	' End If 
	
	' FLName = FLPath & "\" & oFileName & "-FL.dwg"
	' oDoc.SaveAs(FLName, False)
	
	' oDoc = ThisApplication.Documents.ItemByName(FLName)
	
	' ' Dim oModelDoc as Document' = oDoc.AllReferencedDocuments.Item(1)
	
	' ' Try
		' ' oModelDoc = oDoc.AllReferencedDocuments.Item(1)
	' ' Catch 
		' ' MessageBox.Show("There are no models referenced in this drawing.  The Rule cannot continue.  ", "No Model", MessageBoxButtons.OK, MessageBoxIcon.Error)
		' ' Exit Sub
	' ' End Try
	
	' 'If the first reference is not a part or assembly, find the first assembly/part in the list.  
	' ' If oModelDoc.DocumentType <> kAssemblyDocumentObject OrElse oModelDoc.DocumentType <> kPartDocumentObject 
		' ' For docnum = 1 to oDoc.AllReferencedDocuments.Count ' oModelDoc.DocumentType <> kAssemblyDocumentObject OrElse oModelDoc.DocumentType <> kPartDocumentObject
			' ' oModelDoc = oDoc.AllReferencedDocuments.Item(docnum)
			' ' ModelFileName = Right(oModelDoc.FullFileName, oModelDoc.FullFileName.Length - InStrRev(oModelDoc.FullFileName, "\"))
			' ' 'msgbox(ModelFileName)
			' ' If ModelFileName.Contains(oDoc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value) AndAlso oModelDoc.DocumentType <> kPresentationDocumentObject Then  'DocumentType = kAssemblyDocumentObject OrElse oModelDoc.DocumentType = kPartDocumentObject
				' ' 'MsgBox(oModelDoc.FullFileName)
				' ' Exit For
			' ' End If
		' ' Next
	' ' End If
	
	' ' If oModelDoc.DocumentType = kPartDocumentObject Then
			' ' oModelDoc.PropertySets.Item("Design Tracking Properties").Item("Project").Value = ""
	' ' End If
		
	' 'Dim oCtrlDef as ControlDefinitions = ThisApplication.CommandManager.ControlDefinitions
	' 'oCtrlDef.Item("VaultCheckoutTop").Execute
	
	' SharedVariable("LogVar") = "SMLYDK Save as -FL"
	' iLogicVB.RunExternalRule("Write SV to Log.iLogicVB")
	
	' 'check file type 
	
	
	' Dim oOptions as NameValueMap = ThisApplication.TransientObjects.CreateNameValueMap
		' oOptions.Value("DeferUpdates") = True
		' oOptions.Value("FastOpen") = True
		
	' 'open the template
	' oTemplateDoc = ThisApplication.Documents.OpenWithOptions(oTemplate, oOptions, False)
	
	' TitleBlockCopy(oTemplateDoc, oDoc, TBName)
	' 'TitleBlockCopy(oTemplateDoc, oDoc, "UNIVERSAL BOURNE Kit Add")
	' BorderCopy(oTemplateDoc, oDoc, BdrName)
	' ' SymbolCopy(oTemplateDoc, oDoc, "ACAD Leader Right")
	' ' SymbolCopy(oTemplateDoc, oDoc, "ACAD Leader Left")
	' ' SymbolCopy(oTemplateDoc, oDoc, "Univ. Part Notes")
	' ' SymbolCopy(oTemplateDoc, oDoc, "TubeWeldTyp")
	
	' ' Dim IsKit = vbNo
	' ' IsKit = MessageBox.Show("Is this a kit?","Kit?", MessageBoxButtons.YesNo,MessageBoxIcon.Question)
	' ' Try 
		' ' If oModelDoc.PropertySets.Item("Inventor User Defined Properties").Item("DesignLevel") = "Kit" _
		 ' ' OrElse oModelDoc.PropertySets.Item("Inventor User Defined Properties").Item("DesignLevel") = "Element" Then
			' ' IsKit = MessageBox.Show("Is this a kit?","Kit?", MessageBoxButtons.YesNo,MessageBoxIcon.Question)
		' ' End If 
	' ' Catch
	' ' End Try
	
	' Dim oSheets As Sheets = oDoc.Sheets
	' Dim oSheet as Sheet = oDoc.Sheets.Item(1)
	' ' Dim SheetCount = 1
	
	' ' 'if a kit, copy the general notes page, delete the current, and move the newly copied sheet to sheet 2
	' ' If IsKit = vbYes Then
		' ' oTemplateDoc.Sheets.Item(2).CopyTo(oDoc)
		' ' oSheets.Item(2).Delete
		
		' ' Dim oBrowserPane as BrowserPane = oDoc.BrowserPanes.Item("Model")
		' ' Dim ShtGenNotesNew As Sheet = oSheets.Item(oSheets.Count)
		' ' Dim sheetNode As BrowserNode = oBrowserPane.GetBrowserNodeFromObject(ShtGenNotesNew)
		' ' oBrowserPane.Reorder(oBrowserPane.GetBrowserNodeFromObject(oSheet), False, sheetNode)
		
	' ' End If
	
	' oTemplateDoc.Close(True)
	
	' oSheet.Activate
	
	' Dim oBorderDef as BorderDefinition = oDoc.BorderDefinitions.Item(BdrName)
	' Dim oTitleTBDef as TitleBlockDefinition = oDoc.TitleBlockDefinitions.Item(TBName)
	' 'Dim oAddPageDef as TitleBlockDefinition = oDoc.TitleBlockDefinitions.Item("UNIVERSAL BOURNE Kit Add")
	
	' oSheet.Border.Delete
	' oSheet.TitleBlock.Delete
	' oSheet.AddBorder(oBorderDef)
	' oSheet.AddTitleBlock(oTitleTBDef)

	' ' If oDoc.AllReferencedDocuments.Item(1).DocumentType = kPartDocumentObject Then	
		' ' InsertSketchedSymbolOnSheet(oDoc, "Univ. Part Notes")
	' ' End If
	
	' If oSheets.Count > 1 Then
		
		' For Each oSheet In oSheets
			' If SheetCount < oSheets.Count Then
				
				' SheetCount = SheetCount + 1
				' oSheet = oDoc.Sheets.Item(SheetCount)
				' oSheet.Activate
				' oSheet.Border.Delete
				' oSheet.AddBorder(oBorderDef)
				' oSheet.TitleBlock.Delete
				' oSheet.AddTitleBlock(oAddPageDef)
				
			' End If
		' Next
	' End If
	
	' ' PropAdd(oDoc, "NextAssy", "")
	' ' PropAdd(oDoc, "UsedOn", "")
	' ' PropAdd(oModelDoc, "NextAssy", "")
	' ' PropAdd(oModelDoc, "UsedOn", "")
	' ' oDoc.PropertySets.Item("Design Tracking Properties").Item("Mfg Approved By").Value =  "K. Flagg"
	' ' oModelDoc.PropertySets.Item("Design Tracking Properties").Item("Mfg Approved By").Value = "K. Flagg"
	
	' oDoc.Sheets.Item(1).Activate 'goto page 1

	' oDoc.Save
	
End Sub

Sub PropAdd(doc as Document, PropName as String, InitVal as String)
	Dim CustPropSet as PropertySet = doc.PropertySets.Item("Inventor User Defined Properties")
	Try
		prop = CustPropSet.Item(PropName)
	Catch
		CustPropSet.Add("", PropName)
	End Try
End Sub

Sub TitleBlockCopy(SourceDoc as Document, DestDoc as Document, tbName as String)
	' Get the new source title block definition.

	Dim oSourceTitleBlockDef As TitleBlockDefinition = SourceDoc.TitleBlockDefinitions.Item(tbName)

	' Get the new title block definition.
	Dim oNewTitleBlockDef As TitleBlockDefinition = oSourceTitleBlockDef.CopyTo(DestDoc, True)
End Sub

Sub BorderCopy(SourceDoc As Document, DestDoc as Document, BorderName as String)
	Dim oSourceBorderDef as BorderDefinition = SourceDoc.BorderDefinitions.Item(BorderName)
	Dim oNewBorderDef as BorderDefinition = oSourceBorderDef.CopyTo(DestDoc, True)
End Sub

Sub SymbolCopy(SourceDoc As Document, DestDoc as Document, SymbolName as String)
	Dim oSourceSymbolDef as SketchedSymbolDefinition = SourceDoc.SketchedSymbolDefinitions.Item(SymbolName)
	Dim oNewBorderDef as SketchedSymbolDefinition = oSourceSymbolDef.CopyTo(DestDoc, True) 
End Sub


Public Sub InsertSketchedSymbolOnSheet(oDrawDoc as Document, SymbolName as String)

    Dim oSketchedSymbolDef As SketchedSymbolDefinition = oDrawDoc.SketchedSymbolDefinitions.Item(SymbolName)

    Dim oSheet As Sheet = oDrawDoc.Sheets.Item(1)

    Dim oTG As TransientGeometry = ThisApplication.TransientGeometry
	Dim oPoint As Point2d = oTG.CreatePoint2d(0.9652, 26.9748)

    Dim oSketchedSymbol As SketchedSymbol = oSheet.SketchedSymbols.Add(oSketchedSymbolDef, oPoint, 0, 1)
End Sub

Function DocumentFileName(Doc As Document) As String
	Dim FilePath as String
	FilePath = Doc.FullFileName
	DocumentFileName = IO.Path.GetFileNameWithoutExtension(FilePath)
	'MsgBox(DocumentFileName)
End Function

Function DocumentFileNameExt(Doc As Document) As String
	Dim FilePath as String
	FilePath = Doc.FullFileName
	DocumentFileNameExt = IO.Path.GetFileName(FilePath)
	'MsgBox(DocumentFileName)
End Function
