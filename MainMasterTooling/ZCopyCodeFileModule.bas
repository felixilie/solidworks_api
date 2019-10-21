Attribute VB_Name = "ZCopyCodeFileModule"
Option Explicit

Sub CopyCodeFile()

    Dim fso As Scripting.FileSystemObject
    Dim fsoFolder As Scripting.Folder
    Dim DestinationPath As String
    Dim fileName As String
    Dim FileType As String
    Dim swpFilePath As String
    
    Set fso = New Scripting.FileSystemObject
    
    swpFilePath = "C:\Users\FIlie\Documents\Felix Documents IPS\API\MasterTooling\MainMasterTooling.swp"
    
    getFileNameLocal swpFilePath, fileName, FileType
    
    DestinationPath = Left(swpFilePath, Len(swpFilePath) - Len(FileType) - 1 - Len(fileName) - 1) + "\" + fileName + " OlderVersions"
    
    If fso.FolderExists(DestinationPath) = False Then
        fso.CreateFolder DestinationPath
    End If
    
    fso.CopyFile swpFilePath, DestinationPath + "\" + fileName + " " + Format(Now(), "yyyymmddhhnnss") + "." + FileType

End Sub

Private Sub getFileNameLocal(ByVal FilePath As String, ByRef fileName As String, ByRef FileType As String)
    
    Dim character As String
    Dim flag As Boolean
    
    flag = False
    
    character = Right(FilePath, 1)
    FileType = character
    
    While character <> Chr(46)
        FilePath = Left(FilePath, Len(FilePath) - 1)
        character = Right(FilePath, 1)
        If character <> Chr(46) Then FileType = character + FileType
    Wend
    
    character = Right(FilePath, 1)
    fileName = character
    
    While character <> Chr(92)
        FilePath = Left(FilePath, Len(FilePath) - 1)
        character = Right(FilePath, 1)
        If character <> Chr(92) Then fileName = character + fileName
    Wend
    
    fileName = Left(fileName, Len(fileName) - 1)
    

End Sub
