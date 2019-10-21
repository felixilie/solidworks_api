Attribute VB_Name = "Get_all_components1"
Dim swApp As SldWorks.SldWorks
Dim assemb As SldWorks.AssemblyDoc
Dim swModel As SldWorks.ModelDoc2
Dim swModelDocExt As SldWorks.ModelDocExtension
Dim swComponent As SldWorks.Component2
Dim vComponents As Variant
Dim visableComponents() As Variant  ''''Declaration needed for dynamic Array later the use ReDim Preserve visableComponents(k) is required! k=0 is a 1 size array'''''
Dim swCustProp As SldWorks.CustomPropertyManager

Dim assemblyPL() As Variant 'Dynamic Variant. This Matrix should have all the info gathered on the assy eventualy
Dim n As Integer 'Matrix Counter

Dim FieldName As String
Dim PartName As String
Dim PN As String
Dim path As String
Dim ValOut As String
Dim ResolvedValOut As String
Dim WasResolved As Boolean

Dim i As Integer
Dim k As Integer 'Counter for non suppressed \ visiable parts

'There is aproblem to use the Models since they are in LightWeight Mode!!!!!
'Changing Between Resolved part to LightWeight could be found here: _
'_http://help.solidworks.com/2013/English/api/sldworksapi/Set_All_Assembly_Components_Lightweight_or_Resolved_Example_VB.htm


Sub main()

Set swApp = Application.SldWorks
Set swModel = swApp.ActiveDoc
Set assemb = swModel

ReDim assemblyPL(3, 0) 'Only the upper bound of the last dimension in a multidimensional array _
can be changed when you use the Preserve keyword!

k = 0

'swApp.SetUserPreferenceIntegerValue swResolveLightweight, 1

vComponents = assemb.GetComponents(True)

'Debug.Print "Total Number of components: " & UBound(vComponents) + 1 'How many components are _
_Includes Hidden Parts and Suppresed Parts

'Goes through all the parts in the assy, take only the visable ones into visableComponents array.
'Debug.Print VarType(visableComponents(k)) Shows the variant type - good to know
For i = 0 To UBound(vComponents)
    Set swComponent = vComponents(i)
    If swComponent.GetSuppression <> swComponentSuppressed And swComponent.Visible = swComponentVisible Then
        ReDim Preserve visableComponents(k) 'Re declaration of array size - must use the PRESERVE word while rediming!!!!!!!
        Set visableComponents(k) = swComponent
        k = k + 1
        'Debug.Print "Component Name: " & swComponent.Name2
        'Debug.Print "Path for model number " & i + 1; " is " & swComponent.GetPathName
    End If
Next i

k = k - 1 'So k won't exceed array limits
n = 0


Debug.Print "visableComponents Size " & UBound(visableComponents) + 1

For i = 0 To UBound(visableComponents)
    Set swComponent = visableComponents(i)
    swComponent.SetSuppression2 swComponentResolved ' swComponentFullyResolved Fully resolved - recursively resolves the component and any child components see more at: swComponentSuppressionState_e Enumeration
    Set swModel = swComponent.GetModelDoc2
    Set swCustProp = swModel.Extension.CustomPropertyManager("")
    
    partNameandPNfromPath swModel.GetPathName, PartName, PN
    
    Debug.Print "GetPathName: " & swModel.GetPathName
    Debug.Print "PartName: " & PartName
    Debug.Print "PN: " & PN
    
    FieldName = "Description"
    ReDim Preserve assemblyPL(3, n) 'Changing the size of the Matrix. Only one dimension could be changed while preserving the data.
    assemblyPL(0, n) = PartName
    swCustProp.Set2 FieldName, PartName
    
    FieldName = "PartNo"
    assemblyPL(1, n) = PN
    swCustProp.Add3 FieldName, swCustomInfoText, PN, swCustomPropertyReplaceValue
    
    n = n + 1
    
    
    'swCustProp.Get5 FieldName, False, ValOut, ResolvedValOut, WasResolved
    'Debug.Print "VISABLE Component Name: " & swComponent.Name2 & "  Number: " & i
    Debug.Print "Description: " & ValOut
    Debug.Print "i: " & i
Next i

n = n - 1 'So n won't exceed array limits

'Checking what data assemblyPL contains
For j = 1 To n
    Debug.Print "PartName: " & assemblyPL(0, j)
    Debug.Print "Part Number: " & assemblyPL(1, j)
Next j

'How to search in a variant?
' Needs to count how many re-accurances are for each part
'Use AddPropertyExtension to set the name, part number and name for each component
'Export the data into excel sheet
'How the F%#@*$ to works with attributes? How to add, how to change, to components????  - Not required for now

'Debug.Print "Number of Visable components: " & k
'SendKeys "^g ^a {DEL}" 'For deleting that Immediate window

End Sub


'This function devides the file name into PartName and Part Number (P/N)'
'By declaring ByRef values one can get several returns from the functions"
'The function now also checks if the file name format is correct!

'WHEN USING ByRef variables type must be declared, and anyhow Option Explicit MUST be used!!!!

Public Sub partNameandPNfromPath(path As String, ByRef PartName As String, ByRef PN As String)

    path = swModel.GetPathName
    
    'Removes the file name
    character = Right(path, 1)
    While character <> Chr(46) ' 46 for .
        path = Left(path, Len(path) - 1)
        character = Right(path, 1)
    Wend
    
    path = Left(path, Len(path) - 1)
        
    'Takes the file name without the destiantion
        
    FullPartName = ""
    
    character = Right(path, 1)
    While character <> Chr(92) ' 92 for \
        path = Left(path, Len(path) - 1)
        FullPartName = character + FullPartName
        character = Right(path, 1)
    Wend
    
    'Checks if the file name is in the correct format XXXX-XX-XXXXX name, name, etc.
    
    For k = 1 To 13
        If k = 5 Or k = 8 Then
            If Mid(FullPartName, k, 1) <> "-" Then
                PN = "N/A"
                PartName = FullPartName
                'swApp.SendMsgToUser2 "Inncorrect foramt of file Name! No ' - ' between numbers! ", swMbStop, swMbOk
                Exit Sub
            End If
        Else
        End If
        If Not (Mid(FullPartName, k, 1) < "0") And (Mid(FullPartName, k, 1) > "9") And (Mid(FullPartName, k, 1) <> "-") Then
            PN = "N/A"
            PartName = FullPartName
            'swApp.SendMsgToUser2 "Inncorrect foramt of file Name! Only numbers should be used!!!", swMbStop, swMbOk
            Exit Sub
        End If
        If Mid(FullPartName, k, 1) = Chr(32) Then
            PN = "N/A"
            PartName = FullPartName
            'swApp.SendMsgToUser2 "Inncorrect foramt of file Name! Part Number too short!!!", swMbStop, swMbOk
            Exit Sub
        End If
        
    Next k
    
    PN = Mid(FullPartName, 1, 13)
    PartName = Mid(FullPartName, 15, Len(FullPartName))
    
End Sub

