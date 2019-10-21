Attribute VB_Name = "ZCopyCodeFileModuel"
Option Explicit

Sub CopyCodeFile()

    Dim fso As Scripting.FileSystemObject
    Dim swpFilePath As String
    Dim DestinationPath As String
    Dim fileName As String
    Dim fileType As String
    
    swpFilePath = "C:\Users\FIlie\Documents\Felix Documents IPS\API\MasterTooling\MasterTooling.swp"
    DestinationPath = "C:\Users\FIlie\Documents\Felix Documents IPS\API\MasterTooling\OlderVersions"
    
    Set fso = New Scripting.FileSystemObject
    
    getFileNameLocal swpFilePath, fileName, fileType
    
    fso.CopyFile swpFilePath, DestinationPath + "\" + fileName + " " + Format(Now(), "yyyymmddhhnnss") + "." + fileType

End Sub

Private Sub getFileNameLocal(ByVal filePath As String, ByRef fileName As String, ByRef fileType As String)
    
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
