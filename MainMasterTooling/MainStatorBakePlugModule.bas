Attribute VB_Name = "MainStatorBakePlugModule"
Option Explicit

Const PI = 3.14159265358979

Sub MainStatorBakePlug()

    Dim swApp As SldWorks.SldWorks
    Dim swModel As SldWorks.ModelDoc2
    Dim swModelDocExt As SldWorks.ModelDocExtension
    Dim swPart As SldWorks.PartDoc
    Dim swAssy As SldWorks.AssemblyDoc
    Dim swComp As SldWorks.Component2
    Dim swSelectionMgr As SldWorks.SelectionMgr
    Dim swFeature As SldWorks.Feature
    Dim swCircularPatternFeatureData As SldWorks.CircularPatternFeatureData
    Dim swLocCircPatt As SldWorks.LocalCircularPatternFeatureData
    Dim GetTypeName As String
    Dim NbrInstances As Long
    
    Dim lErrors As Long
    Dim lWarnings As Long
    Dim errors As Boolean
    Dim UnitType As String
    Dim value As Boolean
    Dim ToolAssemblyPath As String
    Dim ToolAssemblyFolder As String
    Dim boolstatus As Boolean
    Dim i As Integer
    
'    Dim skippedItemsArray As Variant
    
    Dim inTOmeter As Double
    Dim meterToin As Double
    Dim degTORad As Double
    Dim radToDeg As Double
    
    Dim AssemblyArray() As Variant
    
'    AssemblyArray = Array("", _
'                            "")
    CopyCodeFile
    
    Set swApp = Application.SldWorks
    
    inTOmeter = 0.0254 'Units are in meters
    meterToin = 1 / inTOmeter 'To convert back to inch for checking results
    degTORad = PI / 180 'From degrees to Radians
    radToDeg = 180 / PI
    
    'Part Properties
    Dim CoreID As Double 'Min
    Dim MinUnderConductors As Double
    Dim Length As Double
  
    '***** Tool Dimensions *****
    
    'Plug
    Dim PlugOD As Double 'PlugOD@Sketch1
    Dim PlugStepOD As Double 'PlugStepOD@Sketch1

    '***********************************************************************************************************
    '***********************************************************************************************************
    ToolAssemblyFolder = "C:\Users\FIlie\Documents\Felix Documents IPS\Master Tooling\Main Stator Baking Plug\"
    ToolAssemblyPath = ToolAssemblyFolder + "Plug, Baking, Stator, Main.sldprt"
    '***********************************************************************************************************
    UnitType = "Bell 525" ' "Agusta 609 AC","Agusta 609 DC", "CH47', "SAAB", "Textron", "Scorpion"
    '***********************************************************************************************************
    '***********************************************************************************************************
    
    Select Case UnitType
    
        Case "Bell 525"
            CoreID = 4.228
            MinUnderConductors = 4.228 + 0.05

        Case "Agusta 609 DC"
            CoreID = 5.775
            MinUnderConductors = 5.825
            
        Case "Agusta 609 AC"
            CoreID = 3.78
            MinUnderConductors = 3.984 'Used it and was no Good!

        Case Else
            swApp.SendMsgToUser2 "Data for this unit is not available", swMbStop, swMbOk
            errors = False
            Exit Sub
    End Select

    '***** Calculating Tool Dimensions *****
    
    'Plug
    PlugOD = CoreID - 0.005
    PlugStepOD = CoreID + 0.04
     
    '***** Changing Tool Dimensions *****
    ' DONT FORGET TO CONVERT TO METERS!!!!
    ' DEG in RADIANS!!!!

    Set swModel = swApp.OpenDoc6(ToolAssemblyPath, _
    swDocPART, swOpenDocOptions_Silent, "", lErrors, lWarnings)

    'Plug
'    swApp.ActivateDoc "Plug, Baking, Stator, Main"
'    Set swModel = swApp.ActiveDoc
    swModel.Parameter("PlugOD@Sketch1").SystemValue = PlugOD * inTOmeter
    swModel.Parameter("PlugStepOD@Sketch1").SystemValue = PlugStepOD * inTOmeter
    swModel.EditRebuild3
    swModel.Save3 swSaveAsOptions_Silent, lErrors, lWarnings

'    'Assembly
'    swApp.ActivateDoc "Assembly, Mandrel, Coil Winding, Inductor, AC.sldasm"
'    Set swModel = swApp.ActiveDoc
'    Set swAssy = swModel
'
'    'Unsuppress only the relevant Assemblies
'    'First suppress all the cores:
'    For i = 0 To UBound(AssemblyArray)
'        Set swComp = swAssy.GetComponentByName(AssemblyArray(i))
'        swComp.SetSuppression2 swComponentSuppressed
'    Next i
'    'Next, unsuppress only the relevant Assembly
'    Set swComp = swAssy.GetComponentByName(AssemblyName)
'    swComp.SetSuppression2 swComponentFullyResolved
'
'    swModel.Extension.Rebuild swForceRebuildAll
'    'EditRebuild3
'    swModel.Save3 swSaveAsOptions_Silent, lErrors, lWarnings

End Sub






