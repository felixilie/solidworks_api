Attribute VB_Name = "MainStatorBakingProtectorModule"
Option Explicit

Const PI = 3.14159265358979

Sub MainStatorBakingProtector()

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
    Dim CoreOD As Double
    Dim HeightToHoles As Double
    Dim HeightToOverCT As Double
    Dim NumHoles As Double

    '***** Tool Dimensions *****

    Dim IDOverCore As Double 'IDOverCore@Sketch3
    Dim IDupper As Double 'IDupper@Sketch2
    Dim DistanceToHole As Double 'DistanceToHole@Sketch4
    Dim Height As Double 'Height@Boss-Extrude1

    '***********************************************************************************************************
    '***********************************************************************************************************
    ToolAssemblyFolder = "C:\Users\FIlie\Documents\Felix Documents IPS\Master Tooling\Main Stator Baking Protector\"
    ToolAssemblyPath = ToolAssemblyFolder + "Backing, Protector, Main, Stator 1.SLDPRT"
    '***********************************************************************************************************
    UnitType = "Agusta 609 AC" 'CH47, Agusta 609 AC, Agusta 609 DC
    '***********************************************************************************************************
    '***********************************************************************************************************

    Select Case UnitType

        Case "CH47"

'            AssemblyName = ""

            CoreOD = 2.832 * 2 '
            HeightToHoles = 0.5
            HeightToOverCT = 1.75
            NumHoles = 4
            
        Case "Agusta 609 AC"

'            AssemblyName = ""

            CoreOD = 6.112 '
            HeightToHoles = 0.386
            HeightToOverCT = 1.3
            NumHoles = 4
            
        Case "Agusta 609 DC"

'            AssemblyName = ""

            CoreOD = 6.862 '
            HeightToHoles = 0.383
            HeightToOverCT = 1.75
            NumHoles = 4


        Case Else
            swApp.SendMsgToUser2 "Data for this unit is not available", swMbStop, swMbOk
            errors = False
            Exit Sub
    End Select

    IDOverCore = CoreOD + 0.005
    IDupper = CoreOD - 0.05
    DistanceToHole = HeightToHoles + 0.01
    Height = HeightToOverCT
    

    '***** Changing Tool Dimensions *****
    ' DONT FORGET TO CONVERT TO METERS!!!!
    ' DEG in RADIANS!!!!

    Set swModel = swApp.OpenDoc6(ToolAssemblyPath, _
    swDocPART, swOpenDocOptions_Silent, "", lErrors, lWarnings) 'swDocPART for Part, swDocPART for Assembly

'    swApp.ActivateDoc "Assembly, Mandrel, Coil Winding, Inductor, AC"
'    Set swModel = swApp.ActiveDoc

    swModel.Parameter("IDOverCore@Sketch3").SystemValue = IDOverCore * inTOmeter
    swModel.Parameter("IDupper@Sketch2").SystemValue = IDupper * inTOmeter
    swModel.Parameter("DistanceToHole@Sketch4").SystemValue = DistanceToHole * inTOmeter
    swModel.Parameter("Height@Boss-Extrude1").SystemValue = Height * inTOmeter
    swModel.EditRebuild3
    swModel.Save3 swSaveAsOptions_Silent, lErrors, lWarnings

'    'PressLower
'    swApp.ActivateDoc "Mandrel, Pressing, Lower, Coil, Inductor, AC"
'    Set swModel = swApp.ActiveDoc
'    swModel.Parameter("LowerPressWidth@Sketch1").SystemValue = LowerPressWidth * inTOmeter
'    swModel.Parameter("LowerPressLength@Sketch1").SystemValue = LowerPressLength * inTOmeter
'    swModel.Parameter("LowerPressHeight@Boss-Extrude1").SystemValue = LowerPressHeight * inTOmeter
'    swModel.Parameter("LowerPressRadius@Fillet1").SystemValue = LowerPressRadius * inTOmeter
'    swModel.EditRebuild3
'    swModel.Save3 swSaveAsOptions_Silent, lErrors, lWarnings
'
'    'PressUpper
'    swApp.ActivateDoc "Mandrel, Pressing, Upper, Coil, Inductor, AC"
'    Set swModel = swApp.ActiveDoc
'    swModel.Parameter("UpperPressWidth@Sketch1").SystemValue = UpperPressWidth * inTOmeter
'    swModel.Parameter("UpperPressLength@Sketch1").SystemValue = UpperPressLength * inTOmeter
'    swModel.Parameter("UpperPressHeight@Cut-Extrude1").SystemValue = UpperPressHeight * inTOmeter
'    swModel.Parameter("D1@Boss-Extrude1").SystemValue = (UpperPressHeight + 0.2) * inTOmeter
'    swModel.Parameter("UpperPressRadius@Fillet1").SystemValue = UpperPressRadius * inTOmeter
'    swModel.Parameter("UpperPressCoilWidth@Sketch1").SystemValue = UpperPressCoilWidth * inTOmeter
'    swModel.EditRebuild3
'    swModel.Save3 swSaveAsOptions_Silent, lErrors, lWarnings
'
'    'Adapter
'    swApp.ActivateDoc "Adapter, Mandrel, Coil Winding, Inductor, AC"
'    Set swModel = swApp.ActiveDoc
'    swModel.Parameter("AdapterWidth@Sketch1").SystemValue = AdapterWidth * inTOmeter
'    swModel.Parameter("AdapterLength@Sketch1").SystemValue = AdapterLength * inTOmeter
'    swModel.Parameter("AdapterRadius@Sketch1").SystemValue = AdapterRadius * inTOmeter
'    swModel.EditRebuild3
'    swModel.Save3 swSaveAsOptions_Silent, lErrors, lWarnings
'
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

