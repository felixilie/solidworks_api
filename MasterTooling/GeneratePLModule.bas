Attribute VB_Name = "GeneratePLModule"
Option Explicit

'#If VBA7 Then
'Declare PtrSafe Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As Long
'#Else
'Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'#End If
'Const WM_QUIT = &H12

Dim vComponents As Variant
Dim visableComponents() As Variant  ''''Declaration needed for dynamic Array later the use ReDim Preserve visableComponents(k) is required! k=0 is a 1 size array'''''
Dim swCustProp As SldWorks.CustomPropertyManager
Dim assemblyPartNames() As Variant
Dim assemblyPNs() As Variant 'Dynamic Variant. This Matrix should have all the info gathered on the assy eventualy
Dim assemblyQTYs() As Variant
Dim ActiveConfiguration As SldWorks.Configuration

Dim FieldName As String
Dim PartName As String
Dim PN As String
Dim path As String
Dim AssemblyName As String
Dim AssemblyPN As String
Dim ConfName As String
Dim success As Boolean
Dim indexer As Long
Dim indexer2 As Long

Dim i As Integer
Dim j As Integer
Dim k As Integer 'Counter for non suppressed \ visiable parts
Dim n As Integer 'Matrix Counter
Dim flag As Integer

'There was aproblem to access the Models from thde components since they were in LightWeight Mode!!!!!
'Changing Between Resolved part to LightWeight could be found here: _
'http://help.solidworks.com/2013/English/api/sldworksapi/Set_All_Assembly_Components_Lightweight_or_Resolved_Example_VB.htm


Sub GeneratePL()

    Dim fileType As String

    Set swApp = Application.SldWorks
    
    '********** For Testing ************
    'assemblyFilePath = "Z:\Documentation in Process\Drawings\Agusta\Agusta Westland 609 Tilt Rotor\GNA15-400D-2-A DC Generator\GNA15-400D-2-A Tooling\Main Rotor Stacking\1034-60-06987 Assembly, Stacking, Rotor, Main.sldasm"
    '***********************************
    
    fileTypeFormPath assemblyFilePath, fileType
    
    If Not (fileType = "SLDASM" Or fileType = "sldasm") Then
        swApp.SendMsgToUser "File is not an assembly!"
        Exit Sub
    End If
    
    Set swModel = swApp.OpenDoc6(assemblyFilePath, GetDocumentType(fileType), swOpenDocOptions_Silent, "", errors, warnings)
    Set assemb = swModel
    
    partNameandPNfromPath1 assemblyFilePath, AssemblyName, AssemblyPN
    
    k = 0
    
    'swApp.SetUserPreferenceIntegerValue swResolveLightweight, 1
    
    vComponents = assemb.GetComponents(True) 'Gets all the components from the assy (+visable, suppresed etc. ) _
    False gets also the childs, True - Top level only
    
    'Goes through all the parts in the assy, take only the visable ones into visableComponents array.
    'Debug.Print VarType(visableComponents(k)) Shows the variant type - good to know
    For i = 0 To UBound(vComponents)
        Set swComponent = vComponents(i)
        If swComponent.GetSuppression <> swComponentSuppressed And swComponent.Visible = swComponentVisible Then
            ReDim Preserve visableComponents(k) 'Re declaration of array size - must use the PRESERVE word while rediming
            Set visableComponents(k) = swComponent
            k = k + 1
        End If
    Next i

    k = k - 1 'So k won't exceed array limits
    n = 2
    
    
'    Debug.Print "visableComponents Size " & UBound(visableComponents)
    
    ''''''''Defining Intial value for assemblyPN'''''''''
    ReDim assemblyPNs(0)
    ReDim assemblyPartNames(0)
    ReDim assemblyQTYs(0)
    assemblyPNs(0) = AssemblyPN 'Intial value for PN
    assemblyPartNames(0) = AssemblyName 'Intial value for Description
    assemblyQTYs(0) = 1 'Intial vaule for Quantity
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Set ActiveConfiguration = swModel.GetActiveConfiguration
    Set swCustProp = swModel.Extension.CustomPropertyManager(ActiveConfiguration.Name)
    swCustProp.Add3 "DESCRIPTION", swCustomInfoText, StrConv(AssemblyName, vbUpperCase), swCustomPropertyReplaceValue
    swCustProp.Add3 "PART NUMBER", swCustomInfoText, StrConv(AssemblyPN, vbUpperCase), swCustomPropertyReplaceValue
    swModel.Save3 swSaveAsOptions_Silent, errors, warnings
    
'    For i = 0 To UBound(vComponents)
    For i = 0 To UBound(visableComponents)
        PartName = ""
        PN = ""
        success = 0
'        Set swComponent = vComponents(i)
        Set swComponent = visableComponents(i)
        swComponent.SetSuppression2 swComponentResolved ' swComponentFullyResolved Fully resolved - recursively resolves the component and any child components see more at: swComponentSuppressionState_e Enumeration
        Set swModel = swComponent.GetModelDoc2
        
        Debug.Print "Part Number " & i + 1 & vbCrLf
    '
    '    Debug.Print "Number Of configurations is: " & swModel.GetConfigurationCount
    '
'        Set ActiveConfiguration = swModel.GetActiveConfiguration
        
        ConfName = swComponent.ReferencedConfiguration ' Worked somehow - SOLVED ALL MY PROBLEMS
        Set ActiveConfiguration = swModel.GetConfigurationByName(ConfName)
    '
        Debug.Print "Configuration Name is: " & ActiveConfiguration.Name
    
        Debug.Print "Configuration Description is: " & ActiveConfiguration.Description
    
        Debug.Print "Configuration Alternate Name: " & ActiveConfiguration.AlternateName
        
        partNameandPNfromPath1 ActiveConfiguration.Name, PartName, PN, success
        
        If success = 0 Then
            
            partNameandPNfromPath1 swModel.GetPathName, PartName, PN, success
            
            If success = 0 Then
            
                PartName = ActiveConfiguration.Description
                PN = ActiveConfiguration.AlternateName
                
            Else

                If swModel.GetConfigurationCount > 1 Then

                PN = PN + ActiveConfiguration.Name

                End If
                
            End If
            
        End If
        
'        PartName = StrConv(PartName, vbProperCase)
'        PN = StrConv(PN, vbProperCase)
        
        Debug.Print PartName, PN, success
        
        'For first Item
        If i = 0 Then
            ReDim Preserve assemblyPNs(1)
            ReDim Preserve assemblyPartNames(1)
            ReDim Preserve assemblyQTYs(1)
            assemblyPNs(1) = PN
            assemblyPartNames(1) = PartName
            assemblyQTYs(1) = 1
            
            Set swCustProp = swModel.Extension.CustomPropertyManager(ConfName)
            swCustProp.Add3 "DESCRIPTION", swCustomInfoText, StrConv(PartName, vbUpperCase), swCustomPropertyReplaceValue
            swCustProp.Add3 "PART NUMBER", swCustomInfoText, StrConv(PN, vbUpperCase), swCustomPropertyReplaceValue
            swModel.Save3 swSaveAsOptions_Silent, errors, warnings
            
        Else
            indexer = IsInArray(PartName, assemblyPartNames)
            indexer2 = IsInArray(PN, assemblyPNs)
    '        Debug.Print indexer
            If indexer <> -1 And indexer2 <> -1 Then
                    assemblyQTYs(indexer) = assemblyQTYs(indexer) + 1 'Another Part with same name was found so one number is added
            Else
                ReDim Preserve assemblyPNs(n) 'Changing the size of the Matrix. Only one dimension could be changed while preserving the data.
                ReDim Preserve assemblyPartNames(n)
                ReDim Preserve assemblyQTYs(n)
                assemblyPartNames(n) = PartName
                assemblyPNs(n) = PN
                assemblyQTYs(n) = 1 '% Should define Integer. Others: String - $, Double #
                
                Set swCustProp = swModel.Extension.CustomPropertyManager(ActiveConfiguration.Name)
                swCustProp.Add3 "DESCRIPTION", swCustomInfoText, StrConv(PartName, vbUpperCase), swCustomPropertyReplaceValue
                swCustProp.Add3 "PART NUMBER", swCustomInfoText, StrConv(PN, vbUpperCase), swCustomPropertyReplaceValue
                swModel.Save3 swSaveAsOptions_Silent, errors, warnings
                
                n = n + 1
            End If
        End If
        
    Next i
    
'    swApp.ActivateDoc assemblyFilePath
'    Set swModel = swApp.ActiveDoc
'
'    swModel.EditRebuild3
'    swModel.Save3 swSaveAsOptions_SaveReferenced, errors, warnings '
'
'    swApp.CloseDoc assemblyFilePath
    
    n = n - 1 'So n won't exceed array limits
    
    'Checking what data assemblyPN contains
'    For j = 1 To n
'        Debug.Print "Part Name: " & assemblyPartNames(j)
'        Debug.Print "Part Number: " & assemblyPNs(j)
'        Debug.Print "QTY: " & assemblyQTYs(j)
'    Next j
    
    Debug.Print "Total number of different components: " & n + 1 'Inculdes the Assy itself
    Debug.Print "Total number of components: " & UBound(visableComponents) + 2 'Inculdes the Assy itself
    
    'How to search in a variant? Is there a function for that? Nice to know
    'Use AddPropertyExtension to set the name, part number and name for each component
    'Export the data into excel sheet
    'How the F%#@*$ to works with attributes? How to add, how to change, to components????  - Not required for now
    
    'Debug.Print "Number of Visable components: " & k
    'SendKeys "^g ^a {DEL}" 'For deleting that Immediate window
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''
    'This part is used to Export the data to an Excel file'''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''Calling objets as range should be done through the hierarchy
    ''''Else might get: "runtime error 462 the remote server machine"
    
    Dim fso As Scripting.FileSystemObject
    Set fso = New Scripting.FileSystemObject
    
    Dim xlApp As Excel.Application 'all those objects are under the Excel Application
    Dim xlWB As Excel.Workbook
    Dim xlWB2 As Excel.Workbook
    Dim xlsheets As Excel.Worksheets
    Dim xlsheet As Excel.Worksheet
    
    Dim assembFolder As String
    Dim LDate As String
    Dim UserName As String
    'UserName = Environ("USERNAME") 'Functions used to get Users name
    UserName = "F. Ilie" 'Since Computer Name is Lenovo.....
    LDate = Date
    
    Dim lastRow As Long
    Dim flag2 As Boolean
    Dim stringLocationinArray As Integer
    
    flag2 = True
    
    assembFolder = CutPathFrom(assemblyFilePath)
    
    Set xlApp = New Excel.Application
    
    On Error Resume Next
    Set xlApp = GetObject(, "Excel.Application")
    
    'We've tried to get Excel but if it's nothing then it isn't open
    If xlApp Is Nothing Then
        Set xlApp = CreateObject("Excel.Application")
        flag2 = False
    End If
    
    'It's good practice to reset error warnings
    On Error GoTo 0
    
    If fso.FileExists(assembFolder + "PL" + AssemblyPN + " " + AssemblyName + ".xls") = True Then
        Set xlWB = xlApp.Workbooks.Open(assembFolder + "PL" + AssemblyPN + " " + AssemblyName + ".xls")
        flag = 1
    Else
        Set xlWB = xlApp.Workbooks.Open("C:\Users\FIlie\Documents\Felix Documents IPS\API\PL Template.xls")
        flag = 0
    End If
    
    xlApp.Visible = True
    
    Set xlsheet = xlWB.Sheets("Cover Sheet")
    xlsheet.Select
    
    xlsheet.Range("F1") = "Parts List: " & Chr(10) & "PL" + assemblyPNs(0)
    xlsheet.Range("F1").Characters(12, Len(xlsheet.Range("F1"))).Font.Bold = False
    xlsheet.Range("A2") = assemblyPartNames(0)
    stringLocationinArray = findStringPlace(MasterToolGeneratorForm.UnitListBox.Value, unitList)
    xlsheet.Range("A3") = "Used On: " + MasterToolGeneratorForm.UnitListBox.Text + " - " + unitsNameList(CInt(stringLocationinArray))
    xlsheet.Range("A3").Characters(8, Len(xlsheet.Range("A3"))).Font.Bold = False
    xlsheet.Range("A5") = "Prepared By: " + MasterToolGeneratorForm.DesignerBox.Text
    xlsheet.Range("G5").Characters(12, Len(xlsheet.Range("G5"))).Font.Bold = False
    'xlsheet.Range("G5") = "Release Date: " ' + LDate
    xlsheet.Range("G5").Characters(13, Len(xlsheet.Range("G5"))).Font.Bold = False
    
    Set xlsheet = xlWB.Sheets("Parts List")
    xlsheet.Select 'SHOULD HAVE BEEN SELECTED AFTER SETTING!
    
    'Cells(1, 1).Select
    
    xlsheet.Range("A1") = assemblyPNs(0) + " " + assemblyPartNames(0)
    xlsheet.Range("B3") = assemblyPNs(0)
    xlsheet.Range("C3") = assemblyPartNames(0)
    xlsheet.Range("H3") = assemblyQTYs(0)
    
    For i = 1 To n
        xlsheet.Range("D3").Offset(i) = assemblyPartNames(i)
        xlsheet.Range("B3").Offset(i) = assemblyPNs(i)
        xlsheet.Range("H3").Offset(i) = assemblyQTYs(i)
    Next i
    
    'This part is used to order by part numbers
    xlsheet.Range("B4", xlsheet.Range("H3").End(xlDown)).Sort Key1:=xlsheet.Range("B4", xlsheet.Range("B4").End(xlDown)) ', Order1:=xlAscending, Header:=xlNo
    'Didn't know how to use the Selection!!!!!
    
    'This is used to print only the Area that contains data.
    lastRow = xlsheet.Range("B4").End(xlDown).Row
    xlsheet.Columns("B:B").AutoFit 'Autofit columns
    xlsheet.PageSetup.PrintArea = xlsheet.Range("A1:M" & lastRow).Address
    
    xlsheet.Range("A1").Select
    Set xlsheet = xlWB.Sheets("Cover Sheet")
    xlsheet.Select
    xlsheet.Range("A1").Select
    
    With xlsheet.PageSetup
     .Zoom = False
     .FitToPagesTall = 1
     .FitToPagesWide = 1
    End With
    
    If flag = 0 Then
        xlWB.SaveAs "C:\Users\FIlie\Documents\Felix Documents IPS\API\PL\" + "PL" + AssemblyPN + " " + AssemblyName 'assembFolder + "PL" + AssemblyPN + " " + AssemblyName ' , xlOpenXMLWorkbookMacroEnabled
        '"C:\Users\FIlie\Documents\Felix Documents IPS\API\PL\" + "PL" + AssemblyPN + " " + AssemblyName
    End If
    
    xlWB.Close SaveChanges:=True
    
    If flag2 = False Then xlApp.Quit
    
'    If flag = 0 Then
'        Set xlWB = xlApp.Workbooks("C:\Users\FIlie\Documents\Felix Documents IPS\API\PL Template.xls")
'        xlWB.Close SaveChanges:=False
'    End If
    
'    xlApp.Visible = True
'    PostMessage xlApp.hwnd, WM_QUIT, 0, 0
'
'    Set xlWB = Nothing
'    Set xlApp = Nothing

End Sub


'This function devides the file name into PartName and Part Number (P/N)'
'By declaring ByRef values one can get several returns from the functions"
'The function now also checks if the file name format is correct!

'WHEN USING ByRef variables type must be declared, and anyhow Option Explicit MUST be used!!!!

Sub partNameandPNfromPath1(ByVal path As String, ByRef PartName As String, ByRef PN As String, _
Optional ByRef success As Boolean)

Dim i As Integer
Dim fileName As String

fileNameFromPath path, fileName
PartName = fileName

For i = 1 To Len(path) - 2
    If Mid(fileName, i, 1) = "-" And Mid(fileName, i + 3, 1) = "-" Then
        PN = Mid(fileName, i - 4, 13)
        PartName = Replace(fileName, PN, "")
        PartName = Replace(PartName, ".", "")
        If PartName <> "" Then PartName = Right(PartName, Len(PartName) - 1)
    Exit For
    End If
Next i

If PartName = "" Or PN = "" Then
    success = False
Else
    success = True
End If

End Sub

Sub fileTypeFormPath(ByVal filePath As String, ByRef fileType As String)

Dim i As Integer
Dim character As String
Dim flag As String

flag = 0
fileType = ""

For i = 1 To Len(filePath)
    character = Right(filePath, 1)
    If character = Chr(46) Then
        flag = 1
        Exit For
    End If
    filePath = Left(filePath, Len(filePath) - 1)
    fileType = character + fileType
Next i

If flag = 0 Then fileType = ""

End Sub

Sub fileNameFromPath(ByVal filePath As String, ByRef fileName As String)

Dim i As Integer
Dim character As String
Dim flag As Integer
Dim flag1 As Integer
Dim fileType As String

flag = 0
flag1 = 0
fileName = ""
fileTypeFormPath filePath, fileType

For i = 1 To Len(filePath)
    character = Right(filePath, 1)
    If character = Chr(92) Then
        flag = 1
        Exit For
    End If
    If character = Chr(46) Then flag1 = 1
    filePath = Left(filePath, Len(filePath) - 1)
    fileName = character + fileName
Next i
    
If flag1 = 1 Then
    fileName = Replace(fileName, fileType, "")
    fileName = Left(fileName, Len(fileName) - 1)
End If

End Sub
 
Function GetAllComponents(ParentComponent As Variant, LevelsDown As Integer) As Variant

Dim swApp As SldWorks.SldWorks
Dim swComponent As SldWorks.Component2
Dim ChildrenComponents As Variant

Set swComponent = ParentComponent
Set ChildrenComponents = swComponent.GetChildren

If LevelsDown = 0 Or ChildrenComponents Is Nothing Then
    End Function
Else
    LevelsDown = LevelesDown - 1
    
End If

End Function

'Function to check if a string exists in an array

Function IsInArray(stringToBeFound As String, arr As Variant) As Long

    Dim i As Long
  ' default return value if value not found in array
    IsInArray = -1

    For i = LBound(arr) To UBound(arr)
      If StrComp(stringToBeFound, arr(i), vbTextCompare) = 0 Then
        IsInArray = i
        Exit For
      End If
    Next i
    
End Function



