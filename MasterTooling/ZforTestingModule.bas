Attribute VB_Name = "ZforTestingModule"
Sub forTestingModule()

Dim fullPath As String
Dim CutPath As String

'fullPath = "C:\Users\FIlie\Documents\Felix Documents IPS\Master Tooling\" + _
'"Main rotor coils Brazing tool\Fixed\91253a197_alloy Steel Flat-head Socket Cap Screw.sldprt"

CutPathFrom fullPath, CutPath

Debug.Print CutPath

End Sub


Sub makeArrayofStringTest()

    'Pre Condition - strings are seprate by Chr(13)
    
    Dim value1 As String
    Dim j As Integer
    Dim position As Integer
    Dim flag As Integer
    Dim FileNamesLocal() As Variant
    
    value1 = "Assem.sldasm" & Chr(13) & _
    "Location Ring, Coil Pressing, Pm, Stator.sldprt" & Chr(13) & "Press Plate, Coil Pressing, Pm, Stator.sldprt" & Chr(13) & _
    "1021-23-02204 Lamination, Stator, Pm.sldprt" & Chr(13) & "1032-23-03217 Insulator, Stator Slot, Pm.sldprt" & Chr(13) & _
    "1032-23-03219 Wedge, Stator Slot, Pm.sldprt" & Chr(13) & "1032-23-03218 Winding, Stator, Pm.sldprt" & Chr(13) & _
    "1032-23-03213 Assembly, Stator, Pm.sldasm" & Chr(13) & "Assem.slddrw" & Chr(13) & _
    "1032-23-03213 Assembly, Stator, Pm.slddrw" & Chr(13) & "1021-23-02203 Assembly, Core, Stator, Pm.slddrw" & Chr(13) & _
    "1032-23-03215 Assembly, Core, Stator, Pm.slddrw" & Chr(13) & "Location Ring, Coil Pressing, Pm, Stator.slddrw" & Chr(13) & _
    "Press Plate, Coil Pressing, Pm, Stator.slddrw" & Chr(13) & "Press Plate, Coil Pressing, Pm, Stator111.slddrw" & Chr(13) & Chr(13)
    
    ReDim FileNamesLocal(0)
    FileNamesLocal(0) = 0
    
    j = 0
    flag = 0
    
    'Cleans All extra Chr(13) at the end
    While Right(value1, 1) = Chr(13)
        value1 = Left(value1, Len(value1) - 1)
    Wend

    Do While flag <> 1
    
        position = InStr(1, value1, Chr(13)) ' InStr finds a string position within another string
        If position = 0 Then
            ReDim Preserve FileNamesLocal(j)
            FileNamesLocal(j) = value1
            flag = 1
            Exit Do
        Else
            ReDim Preserve FileNamesLocal(j)
            FileNamesLocal(j) = Left(value1, position - 1)
            Debug.Print FileNamesLocal(j)
            'FileNamesLocal(j) = Replace(FileNamesLocal(j), Chr(13), "")
            value1 = Right(value1, Len(value1) - position)
            Debug.Print value1
            Debug.Print Asc(Left(value1, 1))
            j = j + 1
        End If
        
        Debug.Print position
        
        If j > 2000 Then Exit Do
        
    Loop

    For j = 0 To UBound(FileNamesLocal)
        Debug.Print FileNamesLocal(j)
    Next j
    
    'Debug.Print Asc(FileNamesLocal(UBound(FileNamesLocal)))
    
    Debug.Print UBound(FileNamesLocal)
    
End Sub

Sub lineDownTest()

'This function is not good

    Dim PartName As String
    
    PartName = "Plate, Fixture, Stacking, Stator, PM"

    Dim text1 As String, text2 As String, i As Integer, firstLine As String, secondLine As String
    
    text1 = ""
    
    text2 = PartName
    
    firstLine = ""
    
    secondLine = ""
    
    
    For i = 1 To Len(PartName)
    
        text1 = text1 + Left(text2, 1) 'Words
        
        text2 = Right(text2, Len(text2) - 1) 'Rest of text
        
        If i < 28 Then
        
            If Left(text2, 1) = Chr(32) Then
                
                firstLine = firstLine + text1
                text1 = ""
                
            End If
        
        End If
        
        If i >= 28 Then
            
            If Left(text2, 1) = Chr(32) Or i = Len(PartName) Then
            
                secondLine = secondLine + text1
                text1 = ""
                
            End If
        
        End If

    Next i
    
    PartName = firstLine + vbCrLf + Right(secondLine, Len(secondLine) - 1)
    
    Debug.Print PartName
    
End Sub

Sub excelOpenTest()
    
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
    
    flag2 = True
    
    assembFolder = "C:\Users\FIlie\Documents\Felix Documents IPS\API\Excel Test\"
    
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
    
    If fso.FileExists(assembFolder + "test1" + ".xls") = True Then
        Set xlWB = xlApp.Workbooks.Open(assembFolder + "test1" + ".xls")
        flag = 1
    Else
        Set xlWB = xlApp.Workbooks.Open("C:\Users\FIlie\Documents\Felix Documents IPS\API\PL Template.xls")
        flag = 0
    End If
    
    xlApp.Visible = True
    
    Set xlsheet = xlWB.Sheets("Cover Sheet")
    xlsheet.Select
    
    xlsheet.Range("F1") = "TEST"
    
    Set xlsheet = xlWB.Sheets("Parts List")
    xlsheet.Select 'SHOULD HAVE BEEN SELECTED AFTER SETTING!
    
    'This part is used to order by part numbers
    xlsheet.Range("B4", xlsheet.Range("H3").End(xlDown)).Sort Key1:=xlsheet.Range("B4", xlsheet.Range("B4").End(xlDown)) ', Order1:=xlAscending, Header:=xlNo
    'Didn't know how to use the Selection!!!!!
    
    'This is used to print only the Area that contains data.
    lastRow = xlsheet.Range("B4").End(xlDown).Row
    'xlsheet.Columns("B:B").AutoFit
    xlsheet.PageSetup.PrintArea = xlsheet.Range("A1:M" & lastRow).Address
    
    'xlsheet.PageSetup.PrintArea = xlsheet.Cells(xlsheet.Range("H3").End(xlDown), 13)
    
    With xlsheet.PageSetup
     .Zoom = False
     .FitToPagesTall = 1
     .FitToPagesWide = 1
    End With
    
    If flag = 0 Then
        xlWB.SaveAs assembFolder + "test1" ' , xlOpenXMLWorkbookMacroEnabled
        'assembFolder + "PL" + AssemblyPN + " " + AssemblyName
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
