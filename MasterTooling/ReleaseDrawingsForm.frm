VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ReleaseDrawingsForm 
   Caption         =   "UserForm1"
   ClientHeight    =   6735
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4320
   OleObjectBlob   =   "ReleaseDrawingsForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ReleaseDrawingsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub ReleaseDrawingsFinalCmd_Click()
    
    Dim j As Integer
    Dim i As Integer
    Dim c As Control
    Dim txt As TextBox
    Dim flag As Integer
    Dim userRetrunValue As swMessageBoxResult_e
    Dim fso As Scripting.FileSystemObject
    Dim xlsApp As Excel.Application
    Dim xlsWB As Excel.Workbooks
    Dim ReleaseFolder As String
    Dim status As Boolean
    
    Set fso = New Scripting.FileSystemObject
    
    flag = 0
    
'    For Each c In ReleaseDrawingsForm.Controls 'Goes through all the objects in the form
'        If TypeOf c Is TextBox Then 'Checks only the TextBox ones
'            Set txt = c 'transfers the textbox object out of the controler
'            If IsDate(txt.value) = False And flag = 0 And Not txt.value = Empty Then flag = 1   'IsDate function checks if value is a date but no empty
'        End If
'    Next c

    'DesignerDateBox, MechDesignerDateBox, ElectricalEngDateBox, MaterialEngDateBox,
    'QualityDateBox, CompEngDateBox, ProcessEngDateBox, ProjectMgrDateBox
    
    If IsDate(DesignerDateBox.Value) = False And flag = 0 And Not DesignerDateBox.Value = Empty Then flag = 1
    If IsDate(MechDesignerDateBox.Value) = False And flag = 0 And Not MechDesignerDateBox.Value = Empty Then flag = 1
    If IsDate(ElectricalEngDateBox.Value) = False And flag = 0 And Not ElectricalEngDateBox.Value = Empty Then flag = 1
    If IsDate(MaterialEngDateBox.Value) = False And flag = 0 And Not MaterialEngDateBox.Value = Empty Then flag = 1
    If IsDate(QualityDateBox.Value) = False And flag = 0 And Not QualityDateBox.Value = Empty Then flag = 1
    If IsDate(CompEngDateBox.Value) = False And flag = 0 And Not CompEngDateBox.Value = Empty Then flag = 1
    If IsDate(ProcessEngDateBox.Value) = False And flag = 0 And Not ProcessEngDateBox.Value = Empty Then flag = 1
    If IsDate(ProjectMgrDateBox.Value) = False And flag = 0 And Not ProjectMgrDateBox.Value = Empty Then flag = 1
    
    If flag = 1 Then
        swApp.SendMsgToUser2 "Invalid Dates!", swMbWarning, swMbOk
        Exit Sub
    End If
    
    userRetrunValue = swApp.SendMsgToUser2("Are you Sure you would like to Release the drawing?", swMbQuestion, swMbYesNo)
    If userRetrunValue = swMbHitNo Then Exit Sub
    
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
    
'    If fso.FileExists(assemblyPath + "PL" + Left(FileNames(0), Len(FileNames(0)) - 6) + "xls") = True Then
'        Set xlsApp = New Excel.Application
'        Set xlsWB = xlsApp.Workbooks.Open(assemblyPath + "PL" + Left(FileNames(0), Len(FileNames(0)) - 6) + "xls")
'        'Select both tabs and print, how to check the printing area?
'    End If
    
    
    ReDim Preserve DrawingsPaths(j - 1)
    
    If Right(ReleaseFolderBox.Text, 1) = Chr(92) Then
        ReleaseFolder = ReleaseFolderBox.Text
    Else
        ReleaseFolder = ReleaseFolderBox.Text + Chr(92)
    End If
    
    For i = 0 To UBound(DrawingsPaths)
        
        'If there PL too, the PL is exported to PDF and saved at the relavent folder
        If i = 0 Then
            If fso.FileExists(assemblyPath + "PL" + Left(FileNames(0), Len(FileNames(0)) - 6) + "xls") = True Then
                Dim xlApp As Excel.Application 'all those objects are under the Excel Application
                Dim xlWB As Excel.Workbook
                Dim xlsheets As Variant
                Dim xlsheet As Excel.Worksheet
                Dim LDate As String
                
                Set xlApp = New Excel.Application
                Set xlWB = xlApp.Workbooks.Open(assemblyPath + "PL" + Left(FileNames(0), Len(FileNames(0)) - 6) + "xls")
                
                Set xlsheet = xlWB.Sheets("Parts List")
                
                'This part is to fit the printing size
                With xlsheet.PageSetup
                    .Zoom = False
                    .FitToPagesTall = 1
                    .FitToPagesWide = 1
                End With
                
                Set xlsheet = xlWB.Sheets("Cover Sheet")
                
                LDate = Date
                xlsheet.Range("G5") = "Release Date: " + LDate
                
                xlsheets = Array("Cover Sheet", "Parts List")
                
                xlWB.Sheets(xlsheets).Select
                
                xlApp.ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, fileName:= _
                ReleaseFolder + _
                "PL" + Left(FileNames(0), Len(FileNames(0)) - 6) + "pdf", _
                Quality:=xlQualityStandard, IncludeDocProperties:=True, _
                IgnorePrintAreas:=False, OpenAfterPublish:=True
                
'                'Needs to be tested
'                If printCheckBox.Value = True Then
'                    Set xlsheet = xlWB.Sheets("Cover Sheet")
'                    xlsheets = Array("Cover Sheet", "Parts List")
'                    xlWB.Sheets(xlsheet).PrintOut
'                End If
                
                fso.CopyFile assemblyPath + "PL" + Left(FileNames(0), Len(FileNames(0)) - 6) + "xls", _
                ReleaseFolder + _
                "PL" + Left(FileNames(0), Len(FileNames(0)) - 6) + "xls", True
                
                xlsheet.Select
                
                xlWB.Close SaveChanges:=True
                
                xlApp.Quit
                
            End If
        End If
        
        Set swModel = swApp.OpenDoc6(DrawingsPaths(i), swDocDRAWING, swOpenDocOptions_Silent, "", errors, warnings)
        
        If swModel Is Nothing Then
            swApp.SendMsgToUser2 DrawingsPaths(i) & " Was not found", swMbInformation, swMbOk
            GoTo NextIteration
        End If
        
        Set selMgr = swModel.SelectionManager 'The Accessors for using SelectionMgr Object is SelectionManger which is under IModelDoc2'
        Set swDraw = swModel 'declaring a DrawingDoc could be set like that'
        
        Set swView = swDraw.GetFirstView 'First view you get is the drawing itself!!!!'
        Set swView = swView.GetNextView 'Now you try to get the view....'
        
        changeNote DesignerDateBox.Value, "date1Box@Sheet Format1", 0, 0, status
        If status = False Then GoTo NextIteration
        changeNote MechDesignerDateBox.Value, "date2Box@Sheet Format1", 0, 0
        changeNote ElectricalEngDateBox.Value, "date3Box@Sheet Format1", 0, 0
        changeNote MaterialEngDateBox.Value, "date4Box@Sheet Format1", 0, 0
        changeNote QualityDateBox.Value, "date5Box@Sheet Format1", 0, 0
        changeNote CompEngDateBox.Value, "date6Box@Sheet Format1", 0, 0
        changeNote ProcessEngDateBox.Value, "date7Box@Sheet Format1", 0, 0
        changeNote ProjectMgrDateBox.Value, "date8Box@Sheet Format1", 0, 0
            
        changeNote "--", "RevBox@Sheet Format1", 0, 0
        changeNote " ", "preliminaryBox@Sheet Format1", 0, 0
            
        saveAsModel True, ReleaseFolder, ""

        swModel.Save3 swSaveAsOptions_Silent, errors, warnings
        
        If printCheckBox.Value = True Then
            swModel.Extension.PrintOut4 "", "", Nothing
        End If
        
        swApp.CloseDoc DrawingsPaths(i)
        
NextIteration:
    Next i
    
    Unload Me
    
End Sub
    

Private Sub UserForm_Initialize()

    'Other Boxes values I didn't used yet:
    'For information only 213; Rev 200; Designer 217; Date 194; Design Eng Mech 181; Desgin Eng Mech Date 195; Design Eng Elec. 182; Material Eng 183 Date 196
    'Quality Eng 185, Date 197; Proccess Eng 186, Date 198, Program Mng. 187, Date 199

    caption1.Caption = MasterToolGeneratorForm.DesignerBox.Text
    If MasterToolGeneratorForm.DesignerBox.Text = "" Then DesignerDateBox.Enabled = False
    caption2.Caption = MasterToolGeneratorForm.MechDesignerBox.Text
    If MasterToolGeneratorForm.MechDesignerBox.Text = "" Then DesignerDateBox.Enabled = False
    caption3.Caption = MasterToolGeneratorForm.ElectricalEngBox.Text
    If MasterToolGeneratorForm.ElectricalEngBox.Text = "" Then DesignerDateBox.Enabled = False
    caption4.Caption = MasterToolGeneratorForm.MaterialEngBox.Text
    If MasterToolGeneratorForm.MaterialEngBox.Text = "" Then DesignerDateBox.Enabled = False
    caption5.Caption = MasterToolGeneratorForm.QualityBox.Text
    If MasterToolGeneratorForm.QualityBox.Text = "" Then DesignerDateBox.Enabled = False
    caption6.Caption = MasterToolGeneratorForm.CompEngBox.Text
    If MasterToolGeneratorForm.CompEngBox.Text = "" Then DesignerDateBox.Enabled = False
    caption7.Caption = MasterToolGeneratorForm.ProcessEngBox.Text
    If MasterToolGeneratorForm.ProcessEngBox.Text = "" Then DesignerDateBox.Enabled = False
    caption8.Caption = MasterToolGeneratorForm.ProjectMgrBox.Text
    If MasterToolGeneratorForm.ProjectMgrBox.Text = "" Then DesignerDateBox.Enabled = False
    
    ReleaseFolderBox.Text = ReleaseFolderPath
    
    'THIS IS HOW YOU ADD A CONTROLLER AND LOCATE IT
'    ReleaseForm1.Controls.Add "Forms.TextBox.1", "Name1", True
'    ReleaseForm1!Name1.Top = 216
'    ReleaseForm1!Name1.Left = 126
'    ReleaseForm1!Name1.Text = "COOL!"

End Sub
