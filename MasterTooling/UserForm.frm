VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm 
   Caption         =   "Drawing Controller"
   ClientHeight    =   3870
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11760
   OleObjectBlob   =   "UserForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'Declare Varibles'


Private Sub ReleaseButton_Click()

If ReleaseCheckBox.Value = False Then
    swApp.SendMsgToUser2 "Check Mark Box Should Be checked Before Releasing!", swMbWarning, swMbOk
Else
    ReleaseDrawingsForm.Show 'vbModeless
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
Dim materialIndex As Integer
Dim unitIndex As Integer
Dim i As Integer

ZCopyCodeFileModuel.CopyCodeFile

intiateArrays

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

'checks for the current value in the Material box
materialCurrentValue = getNote("materialBox@Sheet Format1", 0, 0) ' Using TRUE was incorrect!!!! Functions returns an Integer!!!
materialIndex = findStringPlace(materialCurrentValue, materialsList)
If materialIndex = -1 Then materialBox.Text = materialCurrentValue
'materialIndex = FindStringPosition(materialsList, materialCurrentValue)

'checks for the current value in the Unit box
unitCurrentValue = getNote("unitBox@Sheet Format1", 0, 0)
unitIndex = FindStringPosition(unitsNameList, unitCurrentValue)

'Checks for the value of the DesignEngMech
designEngCurrentValue = getNote("designMechBox@Sheet Format1", 0, 0)
DesignEngMechBox.Text = designEngCurrentValue

'Checks for the value of the ProjectMgrManagerTextBox
projectMgrCurrentValue = getNote("programBox@Sheet Format1", 0, 0)
ProjectMgrManagerTextBox.Text = projectMgrCurrentValue

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
nextAssyBox.Text = nextAssyCurrentValue

'Checks for the value of the Hardness
hardnessCurrentValue = getNote("hardnessBox@Sheet Format1", 0, 0)
HardnessBox.Text = hardnessCurrentValue

'Checks for the value of the Finish
finishCurrentValue = getNote("finishBox@Sheet Format1", 0, 0)
FinishBox.Text = finishCurrentValue

'Intiated Value for the drop boxes.

materialBox.Clear

For i = 0 To UBound(materialsList)
    materialBox.AddItem materialsList(i)
Next i

materialBox.Text = materialCurrentValue 'materialBox.List(materialIndex) 'Initializing the first shown Value

unitBox.Clear

For i = 0 To UBound(unitsNameList)
    unitBox.AddItem unitList(i)
Next i

unitBox.Text = unitBox.List(unitIndex) 'Initializing the first shown Value

End Sub

Private Sub fillPartNameandPN_Click()
  
Dim PN As String
Dim PartName As String
Dim path As String
Dim myPageSetup As SldWorks.PageSetup
Dim Value As String
Dim i As Integer
Dim stringLocationinArray As Integer
Dim CurrentMarkRemark As String

Set swApp = Application.SldWorks
Set swModel = swApp.ActiveDoc
Set selMgr = swModel.SelectionManager 'The Accessors for using SelectionMgr Object is SelectionManger which is under IModelDoc2'
Set swDraw = swModel 'declaring a DrawingDoc could be set like that'
Set swView = swDraw.GetFirstView 'First view you get is the drawing itself!!!!'
Set swView = swView.GetNextView 'Now you try to get the view....'
       
'Checks if there are no view, if so - returns error.
If swView Is Nothing Then 'Unless there is no view, which is checked here'
    swApp.SendMsgToUser2 "First Set first View!", 1, 0
    End
End If

' Sets the pageSetup to Scale to Fit
Set myPageSetup = swModel.PageSetup
myPageSetup.ScaleToFit = True

path = swModel.GetPathName

'Checks if the file wasn't Saved As yet
If path = "" Then
    swModel.Save3 0, errors, warnings
    path = swModel.GetPathName 'Using the Method GetPathName which is under iModelDoc2 interface'
End If

saveAsModel False, "", "" 'Function Created by user!
    
StringFunctionsModule.partNameandPNfromPath path, PartName, PN 'Function Created by user!

PartName = UCase(PartName)
        
changeNote PN, "PNBox@Sheet Format1", 0, 0
        
If Len(PartName) > 28 Then lineDown PartName
updateDrawingModule.changeNote PartName, "titleBox@Sheet Format1", 0, 0

stringLocationinArray = findStringPlace(unitBox.Value, unitList)
If stringLocationinArray <> -1 Then
    Value = "UNIT" + Chr(13) + unitsNameList(CInt(stringLocationinArray))
    changeNote Value, "unitBox@Sheet Format1", 0, 0
Else
    changeNote "", "unitBox@Sheet Format1", 0, 0
End If

'Adds Material Note according to value selected
'changeNote materialBox.Text, "materialBox@Sheet Format1", 0, 0
stringLocationinArray = findStringPlace(materialBox.Value, materialsList)
If stringLocationinArray <> -1 Then
    Value = materialsNoteList(CInt(stringLocationinArray))
    changeNote Value, "materialBox@Sheet Format1", 0, 0
Else
    changeNote "", "materialBox@Sheet Format1", 0, 0
End If

'Change Next Assy Name
If usedToMakeBox.Value = True Then
    changeNote "USED TO MAKE" & Chr(13) & nextAssyBox.Value, "nextassemblyBox@Sheet Format1", 0, 0
    'Checks what kind of Extra text should be added
End If
If nextAssemblyBox.Value = True Then
    changeNote "NEXT ASSEMBLY" & Chr(13) & nextAssyBox.Value, "nextassemblyBox@Sheet Format1", 0, 0
End If

CurrentMarkRemark = getNote("noteBox@Sheet Format1", 0, 0)

If CurrentMarkRemark = "" Or CurrentMarkRemark = Empty Or CurrentMarkRemark = Chr(32) Then
Else
    changeNote "PERMANENTLY MARK PART " & Chr(34) & PN & Chr(34) & " PER MIL-STD-130 APPROX. WHERE SHOWN.", "noteBox@Sheet Format1", 0, 0
End If

changeNote DesignEngMechBox.Value, "designMechBox@Sheet Format1", 0, 0

'        changeNote DesignEngMechBox.Value, "designerBox@Sheet Format1", 0, 0
'
'        changeNote MasterToolGeneratorForm.MechDesignerBox.Value, "designMechBox@Sheet Format1", 0, 0
'
'        changeNote MasterToolGeneratorForm.ElectricalEngBox.Value, "designElecBox@Sheet Format1", 0, 0
'
'        changeNote MasterToolGeneratorForm.MaterialEngBox.Value, "materialEngBox@Sheet Format1", 0, 0
'
'        changeNote MasterToolGeneratorForm.QualityBox.Value, "qualityBox@Sheet Format1", 0, 0
'
'        changeNote MasterToolGeneratorForm.CompEngBox.Value, "componentBox@Sheet Format1", 0, 0
'
'        changeNote MasterToolGeneratorForm.ProcessEngBox.Value, "processBox@Sheet Format1", 0, 0

changeNote ProjectMgrManagerTextBox.Value, "programBox@Sheet Format1", 0, 0

'Change Hardness Name
changeNote HardnessBox.Value, "hardnessBox@Sheet Format1", 0, 0

'Change Finish Name
changeNote FinishBox.Value, "finishBox@Sheet Format1", 0, 0

swModel.Save3 swSaveAsOptions_Silent, errors, warnings

End Sub








