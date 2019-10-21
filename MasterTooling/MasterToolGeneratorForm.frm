VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MasterToolGeneratorForm 
   Caption         =   "Master Tool Generator"
   ClientHeight    =   8025
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15810
   OleObjectBlob   =   "MasterToolGeneratorForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MasterToolGeneratorForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub AssembleyPathBox_Change()

End Sub

Private Sub CheckBox1_Click()

End Sub

Private Sub CommandButton10_Click()
    
    Dim Filter                      As String

    Dim fileName                    As String
    
    Dim fileConfig                  As String
    
    Dim fileDispName                As String
    
    Dim fileOptions                 As Long
    
     
    
    Set swApp = Application.SldWorks
    
    ' This following string has three filters associated with it; note the use
    
    ' of the | character between filters
    
    Filter = "SolidWorks Files (*.sldprt; *.sldasm; *.slddrw)|*.sldprt;*.sldasm;*.slddrw|Filter name (*.fil)|*.fil|All Files (*.*)|*.*"
    
    fileName = swApp.GetOpenFileName("File to Open", "", Filter, fileOptions, fileConfig, fileDispName)
    
    AssembleyPathBox.Value = fileName
    
End Sub


Private Sub CommandButton8_Click()

End Sub

Private Sub GenerateNewToolCmd_Click()

    Dim ToolBoxObject As TextBox
    Dim DimensionsArray(8) As Variant
    Dim errors As Boolean
    Dim i As Integer
    
    For i = 1 To 7
        Set ToolBoxObject = Controls("ToolDimBox" & i)
        DimensionsArray(i - 1) = ToolBoxObject.Value
    Next i
    
    MainRotorBrazingFixtureModule.MainRotorBrazingFixture UnitListBox.Text, DimensionsArray, errors
    
'    If errors = True Then
'
'    End If
   
End Sub

Private Sub GeneratePLCmd_Click()
    
    GeneratePL
    
End Sub

Private Sub GenSignCmd_Click()

    Dim i As Integer
    Dim j As Integer
    Dim stringLocationinArray As Integer
    Dim currentUnitName As String
    Dim fileName As String
    Dim fileType As String
    Dim PartName As String
    Dim PN As String
    Dim fso As Scripting.FileSystemObject
    Set fso = New Scripting.FileSystemObject
    
    Set swApp = Application.SldWorks
    
    stringLocationinArray = findStringPlace(MasterToolGeneratorForm.UnitListBox.Value, unitList)
    currentUnitName = MasterToolGeneratorForm.UnitListBox.Value ' + (" ") + unitsNameList(CInt(stringLocationinArray))
    
    ReDim FileNames(0)
    FileNames(0) = 0
    ReDim DrawingsPaths(0)
    DrawingsPaths(0) = 0
    
    StringFunctionsModule.makeArrayofString FileNamesBox.Text, FileNames
    
    For i = 0 To UBound(FileNames)
        
        If Right(FileNames(i), 6) = "slddrw" Then
            ReDim Preserve DrawingsPaths(j)
            DrawingsPaths(j) = assemblyPath + FileNames(i)
            j = j + 1
        End If
        
    Next i
    
    '************ReDim Preserve DrawingsPaths(j - 1)***************
    
    For i = 0 To UBound(DrawingsPaths)
        
        partNameandPNfromPath DrawingsPaths(i), PartName, PN
        
        If i = 0 Then
            If fso.FileExists(assemblyPath + "PL" + PN + " " + PartName + ".xls") = True Then
                SignatureCard PartName, "PL" + PN
            End If
        End If

        SignatureCard PartName, PN

    Next i
    
End Sub

Private Sub getFileNamesCmd_Click()

'    Dim FileNames() As String
    Dim i As Integer
    Dim Value As String
    Dim CurrentValue As String
    Dim unitIndex As Integer
    Dim flag As Integer
    Dim DocumentType As Integer
    
    flag = 0
    
    Set swApp = Application.SldWorks
    
    assemblyFilePath = AssembleyPathBox.Value
    
    DocumentType = GetDocumentType(Right(assemblyFilePath, 6))
    
    ReDim FileNames(0)
    FileNames(0) = 0
    
    'Clean FilneNamesBox if there was text before
    FileNamesBox.Text = ""
    
    PackandGoModule.getPackandGoFileName assemblyFilePath, DocumentType, GetPackandGoStatus
    
    If GetPackandGoStatus = 1 Then
        assemblyPath = CutPathFrom(assemblyFilePath)
        'SavePathBox.value = savetofolder
        SavePathBox.Enabled = True
    End If

    'FileNamesBox.MultiLine = True
    
    For i = 0 To UBound(FileNames)
        FileNamesBox.Text = FileNamesBox.Text + FileNames(i) + Chr(13)
    Next i
    
    'This Part is to retrieve the current data in the drawings field
    If RetrieveDataFromDrawingCheckBox.Value = True Then
        For i = 0 To UBound(FileNames)
'            FileNamesBox.Text = FileNamesBox.Text + FileNames(i) + Chr(13)
            If Right(FileNames(i), 6) = "slddrw" And flag = 0 Then
                flag = 1
                Set swModel = swApp.OpenDoc6(assemblyPath + FileNames(i), swDocDRAWING, _
                    swOpenDocOptions_Silent, "", errors, warnings)
                    
                If swModel Is Nothing Then
                    swApp.SendMsgToUser2 FileNames(i) & " Was not found", swMbInformation, swMbOk
                    GoTo NextIteration
                End If
                
                Set selMgr = swModel.SelectionManager 'The Accessors for using SelectionMgr Object is SelectionManger which is under IModelDoc2'
                Set swDraw = swModel
                
                'checks for the current value in the Unit box
                CurrentValue = getNote("unitBox@Sheet Format1", 0, 0)
                unitIndex = FindStringPosition(unitsNameList, CurrentValue)
                
                UnitListBox.Text = UnitListBox.List(unitIndex)
                
                'Boxes Names:
                'preliminaryBox, titleBox, RevBox, PNBox, designerBox, date1Box, designMechBox, date2Box, designElecBox, date3Box,
                'materialEngBox, date4Box, qualityBox, date5Box, componentBox, date6Box, processBox, date7Box, programBox, date8Box,
                'materialBox, hardnessBox, finishBox, unitBox, nextassemblyBox, noteBox,
                
                CurrentValue = getNote("designerBox@Sheet Format1", 0, 0)
                DesignerBox.Text = CurrentValue
                
                CurrentValue = getNote("designMechBox@Sheet Format1", 0, 0)
                MechDesignerBox.Text = CurrentValue
                
                CurrentValue = getNote("designElecBox@Sheet Format1", 0, 0)
                ElectricalEngBox.Text = CurrentValue
                
                CurrentValue = getNote("materialEngBox@Sheet Format1", 0, 0)
                MaterialEngBox.Text = CurrentValue
                
                CurrentValue = getNote("qualityBox@Sheet Format1", 0, 0)
                QualityBox.Text = CurrentValue
                
                CurrentValue = getNote("componentBox@Sheet Format1", 0, 0)
                CompEngBox.Text = CurrentValue
                
                CurrentValue = getNote("processBox@Sheet Format1", 0, 0)
                ProcessEngBox.Text = CurrentValue
                
                CurrentValue = getNote("programBox@Sheet Format1", 0, 0)
                ProjectMgrBox.Text = CurrentValue
                
                
                'Checks for the value of the nextAssy and if it is Used to Make or Next Assembly
                CurrentValue = getNote("nextassemblyBox@Sheet Format1", 0, 0)
                If InStr(CurrentValue, "USED TO MAKE") <> 0 Then
                    CurrentValue = Right(CurrentValue, Len(CurrentValue) - 16)
                Else
                    CurrentValue = ""
                End If
                usedToMakeBox.Text = CurrentValue
                
                swApp.CloseDoc assemblyPath + FileNames(i)
                
            End If
            
NextIteration:
        Next i
    End If
    
    setFileNamesCmd.Enabled = True
    
End Sub


Private Sub Label20_Click()

End Sub

Private Sub ReleaseDrawingsCmd_Click()

ReDim FileNames(0)
FileNames(0) = 0
    
StringFunctionsModule.makeArrayofString FileNamesBox.Text, FileNames

ReleaseDrawingsForm.Show
    
End Sub

Private Sub SaveAsPDFCMD_Click()
    
    Dim j As Integer
    Dim i As Integer
    Dim PathForPDF As String
    
    PathForPDF = "C:\Users\FIlie\Documents\Felix Documents IPS\PDF for Quotes\"
    
    Set swApp = Application.SldWorks
    
    j = 0
    
    ReDim DrawingsPaths(0)
    DrawingsPaths(0) = 0
    
    For i = 0 To UBound(FileNames)
        
        If Right(FileNames(i), 6) = "slddrw" Then
            ReDim Preserve DrawingsPaths(j)
            DrawingsPaths(j) = assemblyPath + FileNames(i)
            j = j + 1
        End If
        
    Next i
    
    ReDim Preserve DrawingsPaths(j - 1)
    
    For i = 0 To UBound(DrawingsPaths)
        
        Set swModel = swApp.OpenDoc6(DrawingsPaths(i), swDocDRAWING, swOpenDocOptions_Silent, "", errors, warnings)
        
        If swModel Is Nothing Then
            swApp.SendMsgToUser2 DrawingsPaths(i) & " Was not found", swMbInformation, swMbOk
            GoTo NextIteration
        End If
            
        saveAsModel True, PathForPDF, ""

        swModel.Save3 swSaveAsOptions_Silent, errors, warnings
        
        swApp.CloseDoc DrawingsPaths(i)
        
NextIteration:
    Next i

End Sub

Private Sub setFileNamesCmd_Click()

    Dim i As Long
    Dim j As Integer
    Dim k As Integer
    Dim flag As Integer
    Dim arrayOfMultiConfiParts() As Variant
    Dim ConfName As String
    Dim DocumentType As String
    Dim character As String
    Dim strg As String
    
    Set swApp = Application.SldWorks
    
    ReDim FileNames(0)
    FileNames(0) = 0
    
    StringFunctionsModule.makeArrayofString FileNamesBox.Text, FileNames
    
    strg = FileNamesBox.Text
    
'    For j = 0 To UBound(FileNames)
'
'        Debug.Print "File number " & j + 1 & " FileNames(j)"
'
'        For i = 0 To Len(FileNames(j)) - 1
'
'            character = Left(strg, 1)
'            strg = Right(strg, Len(strg) - 1)
'            Debug.Print character, Asc(character)
'
'        Next i
'
'    Next j
    
    'assemblyFilePath = AssembleyPathBox.value

    For j = 0 To UBound(FileNames)

        Debug.Print FileNames(j)

    Next j

    savetofolder = SavePathBox.Value
    
    If savetofolder = "" Then
        swApp.SendMsgToUser2 "A valid path to save to must be chosen!", swMbStop, swMbOk
        Exit Sub
    End If
    
''    'Checks to see that user doesn't save new assembly at the same folder
''    If savetofolder = assemblyPath Or savetofolder = Left(assemblyPath, Len(assemblyPath) - 1) Then
''        swApp.SendMsgToUser2 "Can't save to same folder as original assembly!", swMbStop, swMbOk
''        Exit Sub
''    End If
    
    If Right(savetofolder, 1) <> Chr(92) Then
        savetofolder = savetofolder + Chr(92)
    End If
     
    DocumentType = GetDocumentType(Right(FileNames(0), 6))
    
    PackandGoModule.savePackandGobyNewFileName assemblyFilePath, savetofolder, DocumentType
    
    'This parts searches for multi-configuration parts
    findFileTypeInArray FileNames, arrayOfMultiConfiParts
    
'    For i = 0 To UBound(arrayOfMultiConfiParts)
'
'        Set swModel = swApp.OpenDoc6(savetofolder + arrayOfMultiConfiParts(i) + (".") + "slddrw", swDocDRAWING, _
'        swOpenDocOptions_Silent, "", errors, warnings)
'
'        If swModel Is Nothing Then
'            swApp.SendMsgToUser2 arrayOfMultiConfiParts(i) & " Was not found", swMbInformation, swMbOk
'            GoTo NextIteration
'        End If
'
'        Set selMgr = swModel.SelectionManager 'The Accessors for using SelectionMgr Object is SelectionManger which is under IModelDoc2'
'        Set swDraw = swModel 'declaring a DrawingDoc could be set like that'
'        Set swView = swDraw.GetFirstView 'First view you get is the drawing itself!!!!'
'        Set swView = swView.GetNextView 'Now you try to get the view....'
'        ConfName = swView.ReferencedConfiguration
'        Set swModel = swView.ReferencedDocument
'        Set swConfig = swModel.GetConfigurationByName(ConfName)
'        swConfig.Name = arrayOfMultiConfiParts(i)
'        swModel.Save3 swSaveAsOptions_Silent, errors, warnings
'        swApp.CloseDoc savetofolder + arrayOfMultiConfiParts(i) + (".") + "slddrw"
'
'NextIteration:
'    Next i
    
    'savetofolder = ""
    SavePathBox.Value = ""
    
    assemblyFilePath = savetofolder + FileNames(0)
    assemblyPath = savetofolder
    AssembleyPathBox.Value = assemblyFilePath

End Sub

Private Sub ToolTypeBox_Change()

Dim labelobject As Label
Dim ToolBoxObject As TextBox
Dim i As Integer
    
Select Case ToolTypeBox.Text
    Case Is = "Main Rotor Coils Brazing"
        ToolDimLabel1.Caption = "RtoCoil"
        ToolDimLabel2.Caption = "NumberCoils"
        ToolDimLabel3.Caption = "CoilWidth"
        ToolDimLabel4.Caption = "CoilLength"
        ToolDimLabel5.Caption = "CoilHeight"
        ToolDimLabel6.Caption = "wireWidth"
        ToolDimLabel7.Caption = "CoilRadius"
        ToolDimLabel8.Visible = False
        ToolDimBox8.Visible = False
    Case Is = "Other File"
        AssembleyPathBox.Enabled = True
        UnitListBox.Value = ""
        UnitListBox.Enabled = False
        
        For i = 1 To 8

            Set labelobject = Controls("ToolDimLabel" & i)
            Set ToolBoxObject = Controls("ToolDimBox" & i)
            labelobject.Caption = ""
            ToolBoxObject.Value = ""
            ToolBoxObject.Enabled = False

        Next
        
End Select
    
If ToolTypeBox.Value <> ToolTypeBox.List(1) And ToolTypeBox.Value <> "Other File" Then
    AssembleyPathBox.Enabled = True ' WAS FALSE
    UnitListBox.Enabled = True
Else
    AssembleyPathBox.Enabled = True
    UnitListBox.Enabled = False
End If

End Sub

Private Sub UnitListBox_Change()

Dim DimensionsArray() As Variant
Dim toolAssembleyPath As String
Dim ToolBoxObject As TextBox
Dim i As Integer
 
'I think that this carp caused the problem!!!!!
'Select Case UnitListBox.value
'    Case "Embraer A4 GN", "Agusta 169 GN", "Cessna Latitude GN", "Boeing CH-47 GN", "SAAB GN"
'        GenerateNewToolCmd.Enabled = True
'        MainRotorBrazingFixtureParameters UnitListBox.value, DimensionsArray, toolAssembleyPath
'        AssembleyPathBox.Text = toolAssembleyPath
'        assemblyFilePath = AssembleyPathBox.Text
'        For i = 1 To 7
'            Set ToolBoxObject = Controls("ToolDimBox" & i)
'            ToolBoxObject.value = DimensionsArray(i - 1)
'            ToolBoxObject.Enabled = True
'        Next i
'    Case Else
'        GenerateNewToolCmd.Enabled = False
'        For i = 1 To 7
'            Set ToolBoxObject = Controls("ToolDimBox" & i)
'            ToolBoxObject.value = ""
'            ToolBoxObject.Enabled = False
'        Next i
'End Select

End Sub


Private Sub UpdateDrawingsCmd_Click()

    ReDim FileNames(0)
    FileNames(0) = 0
    Dim i As Integer
    
    
    StringFunctionsModule.makeArrayofString FileNamesBox.Text, FileNames
    
    updateDrawingModule.updateDrawings
    
End Sub

Private Sub UserForm_Initialize()

ZCopyCodeFileModuel.CopyCodeFile

Dim element As Variant
Dim labelobject As Label
Dim ToolBoxObject As TextBox
Dim i As Integer

ArrayFunctions.intiateArrays

setFileNamesCmd.Enabled = False

SavePathBox.Enabled = False

FileNamesBox.MultiLine = True
FileNamesBox.EnterKeyBehavior = True

'ToolDim1Label.Visible = False
'ToolDim1Box.Visible = False

    With ToolTypeBox
        For Each element In ArrayFunctions.toolNames
        .AddItem element
        Next element
    End With
    
    With UnitListBox
        For Each element In ArrayFunctions.unitList
        .AddItem element
        Next element
    End With
    
    For i = 1 To 8

        Set labelobject = Controls("ToolDimLabel" & i)
        Set ToolBoxObject = Controls("ToolDimBox" & i)
        labelobject.Caption = ""
        ToolBoxObject.Value = ""
        ToolBoxObject.Enabled = False

    Next i
    
'Initial Drawing Parameters
DesignerBox.Value = "F. Ilie"
MechDesignerBox.Value = "----"
ElectricalEngBox.Value = "----"
MaterialEngBox.Value = "A. Subramanian"
QualityBox.Value = "A. Benedictis"
CompEngBox.Value = "----"
ProcessEngBox.Value = "B. Hallowell"
ProjectMgrBox.Value = "R. Zacsh"

Label17.Visible = False
ToolTypeBox.Visible = False
ToolParametersFrame.Visible = False
GenerateNewToolCmd.Visible = False
'UnitListBox.Enabled = False

End Sub


