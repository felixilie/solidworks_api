Attribute VB_Name = "PackandGoModule"
Option Explicit

Public Sub getPackandGoFileName(ByVal openFile As String, _
ByVal DocumentType As Integer, Optional ByRef status As Boolean)

    Dim swApp As SldWorks.SldWorks
    Dim swModelDoc As SldWorks.ModelDoc2
    Dim swModelDocExt As SldWorks.ModelDocExtension
    Dim swPackAndGo As SldWorks.PackandGo
    Dim pgFileNames As Variant
    Dim filePath As String
    Dim fileName As String
    Dim fileType As String
    Dim pgFileStatus As Variant
    Dim pgGetFileNames As Variant
    Dim pgDocumentStatus As Variant
'    Dim status As Boolean
    Dim i As Long
    Dim namesCount As Long
    Dim statuses As Variant
'
'    saveToFolder = "C:\Users\FIlie\Documents\Felix Documents IPS\API\PackAndGoTry"

    Set swApp = Application.SldWorks

    ' Open assembly
    Set swModelDoc = swApp.OpenDoc6(openFile, DocumentType, swOpenDocOptions_Silent, "", errors, warnings)
    Set swModelDocExt = swModelDoc.Extension

    ' Get Pack and Go object
    Set swPackAndGo = swModelDocExt.GetPackAndGo

    ' Get number of documents in assembly
    namesCount = swPackAndGo.GetDocumentNamesCount

    ' Include any drawings, SOLIDWORKS Simulation results, and SOLIDWORKS Toolbox components
    swPackAndGo.IncludeDrawings = True
    'Debug.Print "  Include drawings: " & swPackAndGo.IncludeDrawings
    swPackAndGo.IncludeSimulationResults = False
    'Debug.Print "  Include SOLIDWORKS Simulation results: " & swPackAndGo.IncludeSimulationResults
    swPackAndGo.IncludeToolboxComponents = False
    'Debug.Print "  Include SOLIDWORKS Toolbox components: " & swPackAndGo.IncludeToolboxComponents

    ' Get current paths and filenames of the assembly's documents
    status = swPackAndGo.GetDocumentNames(pgFileNames)
    ReDim FileNames(UBound(pgFileNames))
    'FileNames = pgFileNames
    'Debug.Print ""
    'Debug.Print "  Current path and filenames: "
    If (Not (IsEmpty(pgFileNames))) Then
        For i = 0 To UBound(pgFileNames)
            'Debug.Print "    The path and filename is: " & pgFileNames(i)
            filePath = pgFileNames(i)
            StringFunctionsModule.getFileName filePath, fileName, fileType
            fileName = StrConv(fileName, vbProperCase)
            ReDim Preserve FileNames(i)
            FileNames(i) = fileName + Chr(46) + fileType
            'Debug.Print "    The filename is: " & FileNames(i)
        Next i
    End If
    
    swApp.CloseDoc openFile

End Sub

Public Sub savePackandGobyNewFileName(ByVal openFile As String, ByRef savetofolder As String, _
ByVal DocumentType As Integer)

    Dim swApp As SldWorks.SldWorks
    Dim swModelDoc As SldWorks.ModelDoc2
    Dim swModelDocExt As SldWorks.ModelDocExtension
    Dim swPackAndGo As SldWorks.PackandGo
    Dim pgFileNames As Variant
    Dim filePath As String
    Dim fileName As String
    Dim fileType As String
    Dim pgFileStatus As Variant
    Dim pgGetFileNames As Variant
    Dim pgDocumentStatus As Variant
    Dim status As Boolean
    Dim i As Long
    Dim namesCount As Long
    Dim statuses As Variant

    Set swApp = Application.SldWorks
    
    ' Open assembly
    Set swModelDoc = swApp.OpenDoc6(openFile, DocumentType, swOpenDocOptions_Silent, "", errors, warnings)
    Set swModelDocExt = swModelDoc.Extension

    ' Get Pack and Go object
    '''Debug.Print "Pack and Go"
    Set swPackAndGo = swModelDocExt.GetPackAndGo

    ' Get number of documents in assembly
    namesCount = swPackAndGo.GetDocumentNamesCount
    'Debug.Print "  Number of model documents: " & namesCount

    ' Include any drawings, SOLIDWORKS Simulation results, and SOLIDWORKS Toolbox components
    swPackAndGo.IncludeDrawings = True
    'Debug.Print "  Include drawings: " & swPackAndGo.IncludeDrawings
    swPackAndGo.IncludeSimulationResults = False
    'Debug.Print "  Include SOLIDWORKS Simulation results: " & swPackAndGo.IncludeSimulationResults
    swPackAndGo.IncludeToolboxComponents = False
    'Debug.Print "  Include SOLIDWORKS Toolbox components: " & swPackAndGo.IncludeToolboxComponents
    
    status = swPackAndGo.GetDocumentNames(pgFileNames)
'    Debug.Print ""
'    Debug.Print "  Current path and filenames: "
    If (Not (IsEmpty(pgFileNames))) Then
        For i = 0 To UBound(pgFileNames)
            'Debug.Print "    The path and filename is: " & pgFileNames(i)
            'Debug.Print " pgFileNames Size is " & UBound(pgFileNames)
        Next i
    End If
    
    For i = 0 To UBound(FileNames)
        'Debug.Print "FileNames number " & i & " Is " & FileNames(i)
    Next i
    
    'Debug.Print "FileNames size is " & UBound(FileNames)
    
    For i = 0 To UBound(FileNames)
        pgFileNames(i) = savetofolder + FileNames(i)
        'Debug.Print "Full File Path is " & pgFileNames(i)
    Next i
    
    'Debug.Print " pgFileNames Size is " & UBound(pgFileNames)
    
    status = swPackAndGo.SetDocumentSaveToNames(pgFileNames)
    
    'Debug.Print "PackandGo status : " & status
    
    ' Flatten the Pack and Go folder structure; save all files to the root directory
    swPackAndGo.FlattenToSingleFolder = True

    statuses = swModelDocExt.SavePackAndGo(swPackAndGo)
    
    swApp.CloseDoc openFile

End Sub




