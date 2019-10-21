VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm 
   Caption         =   "Drawing Controller"
   ClientHeight    =   3870
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11760
   OleObjectBlob   =   "UserFormImportTitlePN.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Declare Varibles'
Public swApp As SldWorks.SldWorks 'Required always'
Public swModel As SldWorks.ModelDoc2 'Access to top-level (parts, assemblies, drawings)'
Public swDraw As SldWorks.DrawingDoc 'Access to the drawing'
Public selMgr As SldWorks.SelectionMgr
Public swView As SldWorks.View
Public swSheet As SldWorks.Sheet
Public path As String
Public swNote As SldWorks.Note
Public swDrawModel  As SldWorks.ModelDoc2
Public i As Integer
Public ingErrors As Long
Public ingWarnings As Long
Public materialsList As Variant
Public materialsNoteList As Variant
Public unitsList As Variant
Public unitsNameList As Variant


Private Sub ReleaseButton_Click()

If ReleaseCheckBox.Value = False Then
    swApp.SendMsgToUser2 "Check Mark Box Should Be checked Before Releasing!", swMbWarning, swMbOk
Else
    ReleaseForm1.Show 'vbModeless
End If

End Sub

'Other Boxes values I didn't used yet:
'For information only 213; Rev 200; Designer 217; Date 194; Design Eng Mech 181; Desgin Eng Mech Date 195; Design Eng Elec. 182; Material Eng 183 Date 196
'Quality Eng 185, Date 197; Proccess Eng 186, Date 198, Program Mng. 187, Date 199

'New Boxes Names:
'preliminaryBox, titleBox, RevBox, PNBox, designerBox, date1Box, designMechBox, date2Box, designElecBox, date3Box,
'materialEngBox, date4Box, qualityBox, date5Box, componentBox, date6Box, processBox, date7Box, programBox, date8Box,
'materialBox, hardnessBox, finishBox, unitBox, nextassemblyBox, noteBox,

Private Sub UserForm_Initialize()

Dim DocumentType As String
Dim materialCurrentValue As String
Dim unitCurrentValue As String
Dim designEngCurrentValue As String
Dim projectMgrCurrentValue As String
Dim nextAssyCurrentValue As String
Dim hardnessCurrentValue As String
Dim finishCurrentValue As String
Dim paperSize As swDwgPaperSizes_e 'If you want to use the saved prameter names of solidworks
Dim width As Double
Dim height As Double
Dim k As Integer
Dim xlApp As Excel.Application
Dim xlWB As Excel.Workbook
Dim kk As Integer
Dim materialIndex As Integer
Dim unitIndex As Integer

'materialsList and materialsNoteList arrays must be same size
materialsList = Array("ALUMINUM 6061", "4140/4142 ALLOY", "CARBON STEEL", "EPOXY", "WOOD", "Graphit", "GLASS", "GAROLITE", "SEE COMPONENTS")
materialsNoteList = Array("ALUMINUM 6061" & Chr(13) & "PER ASTM B-209, B-211, B-221", "4140/4142 ALLOY" & Chr(13) & "PER ASTM A-29", "CARBON STEEL", "EPOXY", "WOOD", "Graphit", "GLASS", "GAROLITE", "SEE COMPONENTS")

'unitsList and unitsNameList must be same size
unitsList = Array("Cessna Starter Generator", "Agusta 169", "Agusta 609 DC GN", "Agusta 609 AC GN", "Bell 525", "Boeing CH-47", "Cessna Latitude", "Embraer KC390", "Textron", "KC390 Inverter", "Embraer A4 Converter", "G7000")
unitsNameList = Array("SGA1-300-2-A", "SGA11-300-2-A", "GNA15-400D-2-A", "GNA16-25A-2-A", "GNA19-10A-2A", "GNA17-15A-1A", "GNA12-5A-2A", "GNO9-90A-2A", "GNO12-150A-1A", "IN01-10A-1A", "CV05-20A-1A", "TRU3-300-5A")

'Initializing the drop-down list'

Set swApp = Application.SldWorks
Set swModel = swApp.ActiveDoc
'Set swDraw = swModel  ---->>> Could not be called since no one guarantees there is a drawing open

If swModel Is Nothing Then 'Check if object is NULL

    swApp.SendMsgToUser2 "No Document is open!", swMbWarning, swMbOk
    End 'End the App

End If

DocumentType = swModel.GetType

'Checks what kind of file is open, executes only if it is a Drawing document

If DocumentType <> swDocDRAWING Then '= swDocASSEMBLY Or DocumentType = swDocPART - Other way to do that

    swApp.SendMsgToUser2 "Documnet is not a Drawing!", swMbWarning, swMbOk
    End
          
End If

Set selMgr = swModel.SelectionManager 'The Accessors for using SelectionMgr Object is SelectionManger which is under IModelDoc2'
Set swDraw = swModel
Set swSheet = swDraw.GetCurrentSheet
paperSize = swSheet.GetSize(width, height) 'Checks for the paper size

swDraw.EditTemplate
swDraw.EditSketch

'checks for the current value in the Material box
materialCurrentValue = getNote("materialBox@Sheet Format1", 0, 0) ' Using TRUE was incorrect!!!! Functions returns an Integer!!!
materialIndex = FindStringPosition(materialsList, materialCurrentValue)

'checks for the current value in the Unit box
unitCurrentValue = getNote("unitBox@Sheet Format1", 0, 0)
unitIndex = FindStringPosition(unitsNameList, unitCurrentValue)

'Checks for the value of the DesignEngMech
designEngCurrentValue = getNote("designMechBox@Sheet Format1", 0, 0)
DesignEngMechBox.text = designEngCurrentValue

'Checks for the value of the ProjectMgrManagerTextBox
projectMgrCurrentValue = getNote("programBox@Sheet Format1", 0, 0)
ProjectMgrManagerTextBox.text = projectMgrCurrentValue

'Checks for the value of the nextAssy and if it is Used to Make or Next Assembly
nextAssyCurrentValue = getNote("nextassemblyBox@Sheet Format1", 0, 0)
If InStr(nextAssyCurrentValue, "NEXT ASSEMBLY") <> 0 Then
    nextAssyCurrentValue = Right(nextAssyCurrentValue, Len(nextAssyCurrentValue) - 15) 'Remove the NEXT ASSEMBLY text part
    nextAssemblyBox.Value = True
End If
    
If InStr(nextAssyCurrentValue, "USED TO MAKE") <> 0 Then
    nextAssyCurrentValue = Right(nextAssyCurrentValue, Len(nextAssyCurrentValue) - 16) 'Remove the NEXT ASSEMBLY text part
    usedToMakeBox.Value = True
End If
nextAssyBox.text = nextAssyCurrentValue

'Checks for the value of the Hardness
hardnessCurrentValue = getNote("hardnessBox@Sheet Format1", 0, 0)
HardnessBox.text = hardnessCurrentValue

'Checks for the value of the Finish
finishCurrentValue = getNote("finishBox@Sheet Format1", 0, 0)
FinishBox.text = finishCurrentValue

swModel.ClearSelection2 True
swModel.EditSheet
swModel.EditSketch

'Intiated Value for the drop boxes.

materialBox.Clear

For i = 0 To UBound(materialsList)
    materialBox.AddItem materialsList(i)
Next i

materialBox.text = materialBox.List(materialIndex) 'Initializing the first shown Value

unitBox.Clear

For i = 0 To UBound(unitsNameList)
    unitBox.AddItem unitsList(i)
Next i

unitBox.text = unitBox.List(unitIndex) 'Initializing the first shown Value

End Sub

Private Sub fillPartNameandPN_Click()

'Get access to documnet'
       
    'Set Varibles intial values'
        
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    Set selMgr = swModel.SelectionManager 'The Accessors for using SelectionMgr Object is SelectionManger which is under IModelDoc2'
    Set swDraw = swModel 'declaring a DrawingDoc could be set like that'
    Set swView = swDraw.GetFirstView 'First view you get is the drawing itself!!!!'
    Set swView = swView.GetNextView 'Now you try to get the view....'
    
    Dim PN As String
    Dim PartName As String
    Dim sModelName As String
    Dim character As String
    Dim sDrawPath As String
    Dim myPageSetup As SldWorks.PageSetup
    Dim PartNameCurrentValue As String
    Dim EnterCharacterPosition As Integer
    Dim xxx As String
    Dim Value As String
       
    'Checks if there are no view, if so - returns error.
    If swView Is Nothing Then 'Unless there is no view, which is checked here'
        swApp.SendMsgToUser2 "First Set first View!", 1, 0
        End
    End If
    
    Set swDrawModel = swView.ReferencedDocument 'Added for get the reference document to the view
    sModelName = swView.GetReferencedModelName 'Get the Model path
    
    'This function erases the file path file prefix and puts a different prefix
    character = Right(sModelName, 1)
    i = 1
    While character <> Chr(46) ' 46 for .
        sModelName = Left(sModelName, Len(sModelName) - 1)
        character = Right(sModelName, 1)
        i = i + 1
    Wend
    sDrawPath = sModelName + "SLDDRW"
    'function ends
    
    'Debug.Print "sDrawPath  = " & sDrawPath
           
    path = swModel.GetPathName 'Using the Method GetPathName which is under iModelDoc2 interface'
    
    'Checks if the file wasn't Saves As yet
    If path = "" Then
        'swModel.Extension.SaveAs sDrawPath, 0, 1 + 2, Nothing, ingErrors, ingWarnings all of this none required....
        swModel.Save3 0, ingErrors, ingWarnings
        path = swModel.GetPathName 'Using the Method GetPathName which is under iModelDoc2 interface'
    End If
    
    saveAsModel False, "", "" 'Function Created by user!
           
    ' Sets the pageSetup to Scale to Fit
    Set myPageSetup = swModel.PageSetup
    myPageSetup.ScaleToFit = True
        
    swDraw.EditTemplate
    swDraw.EditSketch
 
    Dim MarkNote As String
    
    partNameandPNfromPath path, PartName, PN 'Function Created by user!
    
    PartName = UCase(PartName) 'Sets PartName as UpperCase'
        
    Debug.Print "PN = "; PN 'For Debuging
    
    changeNote PN, "PNBox@Sheet Format1", 0, 0
    
    PartNameCurrentValue = getNote("titleBox@Sheet Format1", 0, 0)
    'Checks if current part names contains Enter Char if so returns its position'
    EnterCharacterPosition = InStr(PartNameCurrentValue, vbCrLf)
    'If it does, it removes the enter char
    If EnterCharacterPosition <> Empty Then
        PartNameCurrentValue = Left(PartNameCurrentValue, EnterCharacterPosition - 1) + Right(PartNameCurrentValue, Len(PartNameCurrentValue) - EnterCharacterPosition - 1)
    End If
    
    'If part Name without the Enter Char is differnet than the files part name, part title should be changed'
    If PartNameCurrentValue <> PartName Then
        If Len(PartName) > 28 Then lineDown PartName 'Checks if title is not too long. If so, lineDown divides the title into two lines
        changeNote PartName, "titleBox@Sheet Format1", 0, 0
    End If
    
    Value = getNote("noteBox@Sheet Format1", 0, 0)
    Debug.Print Value
    
    If Value = "" Or Value = Empty Or Value = Chr(32) Then
        Else: changeNote "PERMANENTLY MARK PART " & Chr(34) & PN & Chr(34) & " PER MIL-STD-130 APPROX. WHERE SHOWN.", "noteBox@Sheet Format1", 0, 0
    End If

    
    'Adds Material Note according to value selected
    i = FindStringPosition(materialsList, materialBox.Value)
    xxx = materialsNoteList(i)
    changeNote xxx, "materialBox@Sheet Format1", 0, 0
      
    'Adds Unit Note according to value selected
    i = FindStringPosition(unitsList, unitBox.Value)
    xxx = "UNIT                    " + unitsNameList(i)
    changeNote xxx, "unitBox@Sheet Format1", 0, 0
       
    'Change Design Eng. Mech Name
    changeNote DesignEngMechBox.Value, "designMechBox@Sheet Format1", 0, 0
    
    'Change Design Eng. Project Manager Name
    changeNote ProjectMgrManagerTextBox.Value, "programBox@Sheet Format1", 0, 0
   
    'Change Next Assy Name
    If usedToMakeBox.Value = True Then changeNote "USED TO MAKE    " & nextAssyBox.Value, "nextassemblyBox@Sheet Format1", 0, 0 'Checks what kind of Extra text should be added
    If nextAssemblyBox.Value = True Then changeNote "NEXT ASSEMBLY  " & nextAssyBox.Value, "nextassemblyBox@Sheet Format1", 0, 0

    'Change Hardness Name
    changeNote HardnessBox.Value, "hardnessBox@Sheet Format1", 0, 0
    
    'Change Finish Name
    changeNote FinishBox.Value, "finishBox@Sheet Format1", 0, 0
            
    'Exit the Edit Sketch Mode'
    
    swModel.ClearSelection2 True
    swModel.EditSheet
    swModel.EditSketch

End Sub


'This function inserts text to a specific note box'
Public Sub changeNote(noteText As String, noteItemName As String, xLocation As Double, yLocation As Double)

   swModel.ClearSelection2 True
   swModel.Extension.SelectByID2 noteItemName, "NOTE", xLocation, yLocation, 0, False, 0, Nothing, 0
   Set swNote = selMgr.GetSelectedObject6(1, -1)
   swNote.SetText noteText

End Sub

'This function gets the note from a note box
Public Function getNote(noteItemName As String, xLocation As Double, yLocation As Double) As String

    swModel.ClearSelection2 True
    swModel.Extension.SelectByID2 noteItemName, "NOTE", xLocation, yLocation, 0, False, 0, Nothing, 0
    Set swNote = selMgr.GetSelectedObject6(1, -1)
    If swNote Is Nothing Then
         swApp.SendMsgToUser2 "Note Box not declared!", swMbWarning, swMbOk
    Else
         getNote = swNote.GetText()
    End If

End Function

'This function either Save As the documnet in a spesific destination and file name and type or it just Save it'
'If no filelocation is provided will be saved in current location
'If no saveType string provided will be saved as PDF'

Public Sub saveAsModel(saveOrSaveAs As Boolean, filelocation As String, saveType As String) 'Using Optional for optional parameters
    
    Dim character As String
    Dim PartName As String
    Dim PN As String
     
    'saveOrSaveAs Value - to just save use False - to save As use True
    If saveOrSaveAs = False Then
        swModel.Save3 0, ingErrors, ingWarnings
    Else
        If filelocation = "" Then
            filelocation = swModel.GetPathName
            saveType = "pdf"
            character = Right(filelocation, 1)
            i = 1
            While character <> Chr(46) ' 46 for .
                filelocation = Left(filelocation, Len(filelocation) - 1)
                character = Right(filelocation, 1)
                i = i + 1
            Wend
            filelocation = filelocation + saveType
        Else
            If saveType = "" Then saveType = "pdf"
            path = swModel.GetPathName
            partNameandPNfromPath path, PartName, PN
            filelocation = filelocation & Chr(92) & PN & " " & PartName & Chr(46) & saveType ' Chr(92) is \
            Debug.Print " filelocation "; filelocation
        End If
    swModel.Extension.SaveAs filelocation, 0, 1 + 2, Nothing, ingErrors, ingWarnings 'Save As Method
    End If

End Sub

'This function devides the file name into PartName and Part Number (P/N)'
'By declaring ByRef values one can get several returns from the functions"

'WHEN USING ByRef variables type must be declared, and anyhow Option Explicit MUST be used!

Public Sub partNameandPNfromPath(path As String, ByRef PartName As String, ByRef PN As String)
    For i = 1 To Len(path) - 2
        If Mid(path, i, 1) = "-" And Mid(path, i + 3, 1) = "-" Then
            PN = Mid(path, i - 4, 13)
            PartName = Mid(path, i + 10, Len(path))
            PartName = Left(PartName, Len(PartName) - 7)
        End If
    Next
End Sub

Function FindStringPosition(arr, v) As Integer
Dim rv As Boolean, lb As Long, ub As Long, i As Long, p As Integer
    lb = LBound(arr)
    ub = UBound(arr)
    For i = lb To ub
        If InStr(v, arr(i)) <> 0 Then
            p = i
            Exit For
        End If
    Next i
    FindStringPosition = p
End Function

'New Boxes Names:
'preliminaryBox, titleBox, RevBox, PNBox, designerBox, date1Box, designMechBox, date2Box, designElecBox, date3Box,
'materialEngBox, date4Box, qualityBox, date5Box, componentBox, date6Box, processBox, date7Box, programBox, date8Box,
'materialBox, hardnessBox, finishBox, unitBox, nextassemblyBox, noteBox,
' If Len(PartName) > 28 Then lineDown PartName 'Checks if title is not too long. If so, lineDown divides the title into two lines

Sub lineDown(ByRef PartName As String)
Dim text1 As String, text2 As String, i As Integer

    For i = 1 To Len(PartName)
        text1 = Left(PartName, i)
        If Right(text1, 1) = Chr(44) Then
            If i > 28 Then
                PartName = text2 & vbCrLf & Right(PartName, Len(PartName) - i)
                Exit For
            End If
        text2 = text1
        End If
    Next i
    
End Sub

