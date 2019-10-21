Attribute VB_Name = "StringFunctionsModule"
Option Explicit

Public Sub getFileName(ByRef filePath As String, ByRef fileName As String, ByRef fileType As String)
    
    Dim character As String
    Dim flag As Boolean
    
    flag = False
    
    character = Right(filePath, 1)
    fileType = character
    
    While character <> Chr(46)
        filePath = Left(filePath, Len(filePath) - 1)
        character = Right(filePath, 1)
        If character <> Chr(46) Then fileType = character + fileType
    Wend
    
    character = Right(filePath, 1)
    fileName = character
    
    While character <> Chr(92)
        filePath = Left(filePath, Len(filePath) - 1)
        character = Right(filePath, 1)
        If character <> Chr(92) Then fileName = character + fileName
    Wend
    
    fileName = Left(fileName, Len(fileName) - 1)
    

End Sub

Public Sub makeArrayofString(ByVal value1 As String, ByRef FileNamesLocal() As Variant)

    Dim j As Integer
    Dim position As Integer
    Dim flag As Integer
    
    ReDim FileNamesLocal(0)
    FileNamesLocal(0) = 0
    
    j = 0
    flag = 0
    
    value1 = Replace(value1, Chr(10), "")
    
    'Cleans All extra Chr(13) at the end
    While Right(value1, 1) = Chr(13)
        value1 = Left(value1, Len(value1) - 1)
    Wend

    Do While flag <> 1
    
        position = InStr(1, value1, Chr(13)) ' InStr finds a string position within another string
        If position = 0 Then
            If value1 <> "" Then
                ReDim Preserve FileNamesLocal(j)
                FileNamesLocal(j) = value1
                FileNamesLocal(j) = Replace(FileNamesLocal(j), Chr(13), "")
            End If
            flag = 1
            Exit Do
        Else
            ReDim Preserve FileNamesLocal(j)
            FileNamesLocal(j) = Left(value1, position)
            FileNamesLocal(j) = Replace(FileNamesLocal(j), Chr(13), "")
            value1 = Right(value1, Len(value1) - position)
            j = j + 1
        End If
  
    Loop
    
    

End Sub

Public Sub partNameandPNfromPath(path As String, ByRef PartName As String, ByRef PN As String)
    Dim i As Integer
    For i = 1 To Len(path) - 2
        If Mid(path, i, 1) = "-" And Mid(path, i + 3, 1) = "-" Then
            PN = Mid(path, i - 4, 13)
            PartName = Mid(path, i + 10, Len(path))
            PartName = Left(PartName, Len(PartName) - 7)
        End If
    Next
End Sub

Public Sub lineDown(ByRef PartName As String)


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
    
End Sub

Public Function findStringPlace(stringTofind As String, arrayToSearch As Variant) As Double

'PreCondition - String is present only at one place in the Array

    Dim i As Integer
    Dim flag As Boolean
    flag = False
    
    For i = 0 To UBound(arrayToSearch)
    
        If arrayToSearch(i) = stringTofind Then
            findStringPlace = i
            flag = True
            Exit For
        End If

    Next i

    If flag = False Then findStringPlace = -1

End Function

Public Sub findFileTypeInArray(arrayToSearch As Variant, ByRef arrayToReturn As Variant)

    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim flag As Integer
    
    Dim arrayOfDrawings() As Variant
    Dim arrayOfParts() As Variant
    
    j = 0
    
    For i = 0 To UBound(arrayToSearch)
        If Right(arrayToSearch(i), 6) = "slddrw" Then
            ReDim Preserve arrayOfDrawings(j)
            arrayOfDrawings(j) = Left(arrayToSearch(i), Len(arrayToSearch(i)) - 7)
            'Debug.Print ("File number ") & j & (" is ") & arrayOfDrawings(j)
            j = j + 1
        End If
    Next i
    
    j = 0
    
    For i = 0 To UBound(arrayToSearch)
        If Right(arrayToSearch(i), 6) = "sldasm" Or Right(arrayToSearch(i), 6) = "sldprt" Then
            ReDim Preserve arrayOfParts(j)
            arrayOfParts(j) = Left(arrayToSearch(i), Len(arrayToSearch(i)) - 7)
            'Debug.Print ("File number ") & j & (" is ") & arrayOfParts(j)
            j = j + 1
        End If
    Next i
    
    k = 0
    flag = 0

    For i = 0 To UBound(arrayOfDrawings)
        For j = 0 To UBound(arrayOfParts)
            If arrayOfDrawings(i) = arrayOfParts(j) Then
                flag = 1
            End If
        Next j
        If flag = 0 Then
            ReDim Preserve arrayToReturn(k)
            arrayToReturn(k) = arrayOfDrawings(i)
            'Debug.Print ("Multi Conf Part number ") & k & (" is ") & arrayToReturn(k)
            k = k + 1
        End If
        flag = 0
    Next i

End Sub

Public Function CutPathFrom(ByVal fullPath As String) As String

    Dim i As Integer
    Dim charachter As String
    Dim CutPath As String
    
    charachter = Right(fullPath, 1)
    CutPath = Left(fullPath, Len(fullPath) - 1)
    
    If charachter = Chr(92) Then
        Exit Function
    End If
    
    While charachter <> Chr(92)
    
        charachter = Right(CutPath, 1)
        CutPath = Left(CutPath, Len(CutPath) - 1)
        
    Wend
    
    CutPathFrom = CutPath + Chr(92)
    
End Function

Public Function FindStringPosition(arr, v) As Integer
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

Public Function GetDocumentType(fileType As String) As Integer

    Select Case fileType
        Case Is = "slddrw"
            GetDocumentType = 3
        Case Is = "SLDDRW"
            GetDocumentType = 3
        Case Is = "sldasm"
            GetDocumentType = 2
        Case Is = "SLDASM"
            GetDocumentType = 2
        Case Is = "sldprt"
            GetDocumentType = 1
        Case Is = "SLDPRT"
            GetDocumentType = 1
        Case Else
            GetDocumentType = 0
    End Select
    
End Function
