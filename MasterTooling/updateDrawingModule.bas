Attribute VB_Name = "updateDrawingModule"
Option Explicit

Public Sub updateDrawings()

    'Boxes Names:
    'preliminaryBox, titleBox, RevBox, PNBox, designerBox, date1Box, designMechBox, date2Box, designElecBox, date3Box,
    'materialEngBox, date4Box, qualityBox, date5Box, componentBox, date6Box, processBox, date7Box, programBox, date8Box,
    'materialBox, hardnessBox, finishBox, unitBox, nextassemblyBox, noteBox,
    
    Dim i As Integer
    Dim j As Integer
    Dim PartName As String
    Dim PN As String
    Dim AssemblyPartName As String
    Dim AssemblyPN As String
    Dim CurrentMarkRemark As String
    Dim stringLocationinArray As Integer
    Dim Value As String
    Dim ConfName As String
    
    Dim swModel1 As SldWorks.ModelDoc2
    
    Set swApp = Application.SldWorks
    
    j = 0
    
    ReDim DrawingsPaths(0)
    DrawingsPaths(0) = 0
    
'    Debug.Print "FileNames size is: " & UBound(FileNames)
    
    'If assemblyPath = "" Then assemblyPath = MasterToolGeneratorForm.AssembleyPathBox.value
    
    For i = 0 To UBound(FileNames)
        
        If Right(FileNames(i), 6) = "slddrw" Then
            ReDim Preserve DrawingsPaths(j)
            DrawingsPaths(j) = assemblyPath + FileNames(i)
            j = j + 1
        End If
        
    Next i
    
    '***********ReDim Preserve DrawingsPaths(j - 1)**************
    
'    Debug.Print "DrawingsPath size is: " & UBound(DrawingsPaths)
    
'    Debug.Print "Save To Folder is: " & savetofolder
'    For i = 0 To UBound(DrawingsPaths)
'        Debug.Print DrawingsPaths(i)
'    Next i
    
    For i = 0 To UBound(DrawingsPaths)
        
        Set swModel = swApp.OpenDoc6(DrawingsPaths(i), swDocDRAWING, swOpenDocOptions_Silent, "", errors, warnings)
        
        If swModel Is Nothing Then
            swApp.SendMsgToUser2 DrawingsPaths(i) & " Was not found", swMbInformation, swMbOk
            GoTo NextIteration
        End If
        
        Set selMgr = swModel.SelectionManager 'The Accessors for using SelectionMgr Object is SelectionManger which is under IModelDoc2'
        Set swDraw = swModel 'declaring a DrawingDoc could be set like that'

        StringFunctionsModule.partNameandPNfromPath DrawingsPaths(i), PartName, PN 'Function Created by user!
        
        Set swView = swDraw.GetFirstView 'First view you get is the drawing itself!!!!'
        Set swView = swView.GetNextView 'Now you try to get the view....'
        ConfName = swView.ReferencedConfiguration
        
        If PartName = "" Then
            StringFunctionsModule.partNameandPNfromPath ConfName, PartName, PN
        End If
        
        Set swModel1 = swView.ReferencedDocument
        Set swConfig = swModel1.GetConfigurationByName(ConfName)
    
        PartName = UCase(PartName) 'Sets PartName as UpperCase'
        
'        value = PN + (" ") + PartName
'
'        swConfig.AlternateName = value
'        swConfig.UseAlternateNameInBOM = True
        
'        Debug.Print ("AlternateName is: ") & value
        
        changeNote PN, "PNBox@Sheet Format1", 0, 0
        
        If Len(PartName) > 28 Then lineDown PartName
        updateDrawingModule.changeNote PartName, "titleBox@Sheet Format1", 0, 0
        
        StringFunctionsModule.partNameandPNfromPath assemblyFilePath, AssemblyPartName, AssemblyPN
        
        stringLocationinArray = findStringPlace(MasterToolGeneratorForm.UnitListBox.Value, unitList)
        If stringLocationinArray <> -1 Then
            Value = "UNIT" + Chr(13) + unitsNameList(CInt(stringLocationinArray))
            changeNote Value, "unitBox@Sheet Format1", 0, 0
        Else
            changeNote "", "unitBox@Sheet Format1", 0, 0
        End If
        
        If PN = AssemblyPN Then
            changeNote "USED TO MAKE" & Chr(13) & MasterToolGeneratorForm.usedToMakeBox.Value, "nextassemblyBox@Sheet Format1", 0, 0
        Else
            changeNote "NEXT ASSEMBLY" & Chr(13) & AssemblyPN, "nextassemblyBox@Sheet Format1", 0, 0
        End If
        
        CurrentMarkRemark = getNote("noteBox@Sheet Format1", 0, 0)
        
        If CurrentMarkRemark = "" Or CurrentMarkRemark = Empty Or CurrentMarkRemark = Chr(32) Then
        Else
            changeNote "PERMANENTLY MARK PART " & Chr(34) & PN & Chr(34) & " PER MIL-STD-130 APPROX. WHERE SHOWN.", "noteBox@Sheet Format1", 0, 0
        End If
        
        changeNote MasterToolGeneratorForm.DesignerBox.Value, "designerBox@Sheet Format1", 0, 0
        
        changeNote MasterToolGeneratorForm.MechDesignerBox.Value, "designMechBox@Sheet Format1", 0, 0
        
        changeNote MasterToolGeneratorForm.ElectricalEngBox.Value, "designElecBox@Sheet Format1", 0, 0
        
        changeNote MasterToolGeneratorForm.MaterialEngBox.Value, "materialEngBox@Sheet Format1", 0, 0
        
        changeNote MasterToolGeneratorForm.QualityBox.Value, "qualityBox@Sheet Format1", 0, 0
        
        changeNote MasterToolGeneratorForm.CompEngBox.Value, "componentBox@Sheet Format1", 0, 0
        
        changeNote MasterToolGeneratorForm.ProcessEngBox.Value, "processBox@Sheet Format1", 0, 0
    
        changeNote MasterToolGeneratorForm.ProjectMgrBox.Value, "programBox@Sheet Format1", 0, 0
        
        swModel.Save3 swSaveAsOptions_Silent, errors, warnings
        swApp.CloseDoc DrawingsPaths(i)
        
NextIteration:
    Next i
    
    
End Sub


'This function inserts text to a specific note box'
Public Sub changeNote(noteText As String, noteItemName As String, xLocation As Double, yLocation As Double, Optional ByRef status As Boolean)
    
    status = True
    swModel.ClearSelection2 True
    swModel.Extension.SelectByID2 noteItemName, "NOTE", xLocation, yLocation, 0, False, 0, Nothing, 0
    Set swNote = selMgr.GetSelectedObject6(1, -1)
    If swNote Is Nothing Then
        'swApp.SendMsgToUser2 "Note Box not declared!", swMbWarning, swMbOk
        status = False
    Else
        swNote.SetText noteText
    End If

End Sub

'This function gets the note from a note box
Public Function getNote(noteItemName As String, xLocation As Double, yLocation As Double, Optional ByRef status As Boolean) As String

    status = True
    swModel.ClearSelection2 True
    swModel.Extension.SelectByID2 noteItemName, "NOTE", xLocation, yLocation, 0, False, 0, Nothing, 0
    Set swNote = selMgr.GetSelectedObject6(1, -1)
    If swNote Is Nothing Then
         'swApp.SendMsgToUser2 "Note Box not declared!", swMbWarning, swMbOk
         status = False
    Else
         getNote = swNote.GetText()
    End If

End Function

Public Sub saveAsModel(saveOrSaveAs As Boolean, ByVal filelocation As String, saveType As String) 'Using Optional for optional parameters
    
    Dim path As String
    Dim fileName As String
    Dim fileType As String
     
    'saveOrSaveAs Value - to just save use False - to save As use True
    If saveOrSaveAs = False Then
        swModel.Save3 0, errors, warnings
    Else
        path = swModel.GetPathName
        getFileName path, fileName, fileType
        If filelocation = "" Then
            saveType = "pdf"
            filelocation = path + fileName + "." + saveType
        Else
            If saveType = "" Then saveType = "pdf"
            filelocation = filelocation + fileName + "." + saveType ' Chr(92) is \
            Debug.Print filelocation
        End If
    swModel.Extension.SaveAs filelocation, 0, 1 + 2, Nothing, errors, warnings 'Save As Method
    End If

End Sub
