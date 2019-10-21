Attribute VB_Name = "ExciterRotorCrushingModule"
Option Explicit

Const PI = 3.14159265358979

Sub ExciterRotorCrushing()

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
    Dim ToolMainSketchPath As String
    Dim boolstatus As Boolean
    Dim longstatus As Long
    Dim i As Integer
    
    Dim skippedItemsArray As Variant
    
    Dim inTOmeter As Double
    Dim meterToin As Double
    Dim degTORad As Double
    Dim radToDeg As Double
    
    Dim CoreNamesArray() As Variant
    
    CoreNamesArray = Array("1034-21-07119 Lamination, Stator, Main-1", _
                            "1034-21-05972 Lamination, Stator, Main-1", _
                            "1015-21-04311 Lamination, Stator, Main-1", _
                            "1015-21-01116 Assembly, Core, Stator, Main-1")
    
    CopyCodeFile
    
    Set swApp = Application.SldWorks
    
    inTOmeter = 0.0254 'Units are in meters
    meterToin = 1 / inTOmeter 'To convert back to inch for checking results
    degTORad = PI / 180 'From degrees to Radians
    radToDeg = 180 / PI
    
    'Part Properties
    Dim CoreName As String
    Dim LamOD As Double 'ID of tooths
    Dim LamODTol As Double
    Dim LamID As Double
    Dim LamIDTol As Double
    Dim BandID As Double
    Dim BandIDTol As Double
    Dim BandThick As Double
    Dim BandThickTol As Double
    Dim CoilID As Double
    Dim CoilIDTol As Double
    Dim DivisionPattern As Integer 'usually 3 or 6
    Dim LocationPinD As Double 'Min Value, LocationPinD@MainSketch
    Dim LocationPinPatternD As Double ' LocationPinPatternD@MainSketch
    Dim DrainHoleD1 As Double ' DrainHoleD1@MainSketch
    Dim DrainHolePattern1 As Double ' DrainHolePattern1@MainSketch
    Dim DrainHoleD2 As Double ' DrainHoleD2@MainSketch
    Dim DrainHolePattern2 As Double ' DrainHolePattern2@MainSketch
    Dim DrainHoleAngle2 As Double ' DrainHoleAngle2@MainSketch
    Dim DrainHoleD3 As Double ' DrainHoleD3@MainSketch
    Dim DrainHolePattern3 As Double ' DrainHolePattern3@MainSketch
    Dim DrainHoleAngle3 As Double ' DrainHoleAngle3@MainSketch
    Dim CoreHeight As Double 'Mid value
    
    '***** Tool Dimensions *****
    
    Dim OD As Double 'OD@MainSketch
    Dim ID As Double 'ID@MainSketch
    Dim ScrewPatternD As Double 'ScrewPatternD@MainSketch
    Dim PinD As Double
    
    
    '***********************************************************************************************************
    '***********************************************************************************************************
    ToolAssemblyFolder = "C:\Users\FIlie\Documents\Felix Documents IPS\Master Tooling\Exciter Rotor Stacking\"
    ToolMainSketchPath = ToolAssemblyFolder + "MainSketch.SLDPRT"
    ToolAssemblyPath = ToolAssemblyFolder + "Assem.SLDASM"
    '***********************************************************************************************************
    UnitType = "Agusta 169" ' "Agusta 609 DC", "Agusta 169', "Latitude", "Agusta 609 AC", "SAAB"
    '***********************************************************************************************************
    '***********************************************************************************************************
    
    Select Case UnitType

        Case "Agusta 169"
        
            'CoreName =
            
            LamOD = 3.965 - 2 * 0.045 'ActualOD - 2 * Tooth width
            LamID = 1.03
            DivisionPattern = 3 'usually 3 or 6
            LocationPinD = 0.31 'Min Value
            LocationPinPatternD = 0.9 * 2 + 0.312
            DrainHoleD1 = 0.312 'Check size with dowel pin - What did I use in the fixture I made?
            DrainHolePattern1 = 2 * (0.9 + 0.312 + 0.125 + 0.312 / 2)
            DrainHoleD2 = 0.1
            DrainHolePattern2 = 1.58 * 2
            DrainHoleAngle2 = 13
            DrainHoleD3 = 3 / 8
            DrainHolePattern3 = 0.9 * 2 + 0.312
            DrainHoleAngle3 = 30
            ScrewPatternD = 1.2 * 2
            CoreHeight = 0.563 'Mid value
            
        Case "Latitude"
        
            'CoreName =
            
            LamOD = 4.15 - 0.04  'ActualOD - 2 * Tooth width
            LamID = 1
            DivisionPattern = 3 'usually 3 or 6
            LocationPinD = 0.313 'Min Value
            LocationPinPatternD = 2.988
            DrainHoleD1 = 0.312
            DrainHolePattern1 = 1.053 * 2
            DrainHoleD2 = 0.16
            DrainHolePattern2 = 1.515 * 2
            DrainHoleAngle2 = 13
            DrainHoleD3 = 3 / 8
            DrainHolePattern3 = 2.5
            DrainHoleAngle3 = 30
            ScrewPatternD = 2.5
            CoreHeight = 0.475 'Mid value

        Case Else
            swApp.SendMsgToUser2 "Data for this unit is not available", swMbStop, swMbOk
            errors = False
            Exit Sub
    End Select
    
    '***** Calculating Tool Dimensions *****
    
    OD = LamOD - 0.2 ' - 0.1
    ID = 3 / 8
    PinD = LocationPinD - 0.001
    
    '***** Changing Tool Dimensions *****
    ' DONT FORGET TO CONVERT TO METERS!!!!
    ' DEG in RADIANS!!!!

    Set swModel = swApp.OpenDoc6(ToolAssemblyPath, _
    swDocASSEMBLY, swOpenDocOptions_Silent, "", lErrors, lWarnings)

'    MainSketch
    swApp.ActivateDoc "MainSketch"
    Set swModel = swApp.ActiveDoc
    swModel.Parameter("LocationPinD@MainSketch").SystemValue = LocationPinD * inTOmeter
    swModel.Parameter("LocationPinPatternD@MainSketch").SystemValue = LocationPinPatternD * inTOmeter
    swModel.Parameter("DrainHoleD1@MainSketch").SystemValue = DrainHoleD1 * inTOmeter
    swModel.Parameter("DrainHolePattern1@MainSketch").SystemValue = DrainHolePattern1 * inTOmeter
    swModel.Parameter("DrainHoleD2@MainSketch").SystemValue = DrainHoleD2 * inTOmeter
    swModel.Parameter("DrainHolePattern2@MainSketch").SystemValue = DrainHolePattern2 * inTOmeter
    swModel.Parameter("DrainHoleAngle2@MainSketch").SystemValue = DrainHoleAngle2 * degTORad
    swModel.Parameter("DrainHoleD3@MainSketch").SystemValue = DrainHoleD3 * inTOmeter
    swModel.Parameter("DrainHolePattern3@MainSketch").SystemValue = DrainHolePattern3 * inTOmeter
    swModel.Parameter("DrainHoleAngle3@MainSketch").SystemValue = DrainHoleAngle3 * degTORad
    swModel.Parameter("OD@MainSketch").SystemValue = OD * inTOmeter
    swModel.Parameter("ID@MainSketch").SystemValue = ID * inTOmeter
    swModel.Parameter("ScrewPatternD@MainSketch").SystemValue = ScrewPatternD * inTOmeter
    swModel.Parameter("PinD@MainSketch").SystemValue = PinD * inTOmeter

    swModel.EditRebuild3
    swModel.Save3 swSaveAsOptions_Silent, lErrors, lWarnings
    
'    Alignment Rod
    swApp.ActivateDoc "Alignment Rod"
    Set swModel = swApp.ActiveDoc
    swModel.Parameter("PinD@Sketch1").SystemValue = PinD * inTOmeter

    swModel.EditRebuild3
    swModel.Save3 swSaveAsOptions_Silent, lErrors, lWarnings

    
'    'This part is used to control if the plate shape is square or hex (like for the augsta DC)
'    Set swFeature = swPart.FeatureByName("Cut-Hex")
'    If NumberOfTabs = 6 Then
'        swFeature.SetSuppression2 swUnSuppressFeature, swAllConfiguration, Nothing
'    Else
'        swFeature.SetSuppression2 swSuppressFeature, swAllConfiguration, Nothing
'    End If
'
'    Set swFeature = swPart.FeatureByName("Boss-Hex")
'    If NumberOfTabs = 6 Then
'        swFeature.SetSuppression2 swUnSuppressFeature, swAllConfiguration, Nothing
'    Else
'        swFeature.SetSuppression2 swSuppressFeature, swAllConfiguration, Nothing
'    End If
    
    
'    'This part is used to control the number of instances in the circular pattern
'    Set swModelDocExt = swModel.Extension
'    boolstatus = swModelDocExt.SelectByID2("CirPattern1", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
'    Set swSelectionMgr = swModel.SelectionManager
'    Set swFeature = swSelectionMgr.GetSelectedObject6(1, -1)
'    Set swCircularPatternFeatureData = swFeature.GetDefinition
'    ' Get or sets the number of instances in the circular-pattern feature
'    swCircularPatternFeatureData.AccessSelections swModel, Nothing 'WASN'T REQUIRED
'    swCircularPatternFeatureData.TotalInstances = NumberOfTabs
'    skippedItemsArray = swCircularPatternFeatureData.SkippedItemArray
'    For i = 0 To UBound(skippedItemsArray)
'        Debug.Print skippedItemsArray(i)
'    Next i
'    'After updating Feature you must use ModifyDefinition, so changes will take place!
'    swFeature.ModifyDefinition swCircularPatternFeatureData, swModel, Nothing
    
'    'Teflon
'    Set swModel = swApp.OpenDoc6(ToolAssemblyFolder + "Teflon.SLDPRT", _
'    swDocPART, swOpenDocOptions_Silent, "", lErrors, lWarnings)
'    swApp.ActivateDoc "Teflon"
'    Set swModel = swApp.ActiveDoc
'
'    swModel.EditRebuild3
'    swModel.Save3 swSaveAsOptions_Silent, lErrors, lWarnings
    
    'Assembly
    swApp.ActivateDoc "Assem"
    Set swModel = swApp.ActiveDoc
    Set swAssy = swModel
''''
'''''    'Unsuppress only the relevant core
'''''    'First suppress all the cores:
'''''    For i = 0 To UBound(CoreNamesArray)
'''''        Set swComp = swAssy.GetComponentByName(CoreNamesArray(i))
'''''        If Not swComp Is Nothing Then
'''''            Debug.Print swComp.Name
'''''            swComp.SetSuppression2 swComponentSuppressed
'''''        End If
'''''    Next i
'''''    'Next, unsuppress only the relevant core
'''''    Set swComp = swAssy.GetComponentByName(CoreName)
'''''    swComp.SetSuppression2 swComponentFullyResolved
''''
'''''    'This part is used to control the number of instances in the circular pattern for ASSEMBLY
'''''    Set swFeature = swAssy.FeatureByName("LocalCirPattern1")
'''''    If swFeature Is Nothing Then Debug.Print "swFeature is nothing"
'''''    Set swLocCircPatt = swFeature.GetDefinition 'Might be beacuse I forgot the SET???? Or didn't work beacuse of SelectByID2???
'''''    swLocCircPatt.TotalInstances = NumberOfTabs
'''''    swFeature.ModifyDefinition swLocCircPatt, swModel, Nothing
''''
    swModel.Extension.Rebuild swForceRebuildAll
    'EditRebuild3
    swModel.Save3 swSaveAsOptions_Silent, lErrors, lWarnings

End Sub








