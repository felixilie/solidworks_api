Attribute VB_Name = "PMStatorStackingModule"
Option Explicit

Const PI = 3.14159265358979

Sub PMStatorStackingNew()

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
    
    Dim skippedItemsArray As Variant
    
    Dim inTOmeter As Double
    Dim meterToin As Double
    Dim degTORad As Double
    Dim radToDeg As Double
    
    Dim CoreNamesArray() As Variant
    
'    CoreNamesArray = Array("1034-21-07119 Lamination, Stator, Main-1", _
'                            "1034-21-05972 Lamination, Stator, Main-1", _
'                            "1015-21-04311 Lamination, Stator, Main-1", _
'                            "1015-21-01116 Assembly, Core, Stator, Main-1")
    
    CopyCodeFile
    
    Set swApp = Application.SldWorks
    
    inTOmeter = 0.0254 'Units are in meters
    meterToin = 1 / inTOmeter 'To convert back to inch for checking results
    degTORad = PI / 180 'From degrees to Radians
    radToDeg = 180 / PI
    
    'Part Properties
    Dim CoreName As String
    Dim NumberOfSlots As Integer
    Dim LamMinOD As Double ' Without Tabs
    Dim LamMinID As Double
    Dim LamThickness As Double
    Dim LamSlotLocationD As Double
    Dim LamSlotMinWidth As Double
    Dim CoreHeight As Double 'Mid value
    
    Dim AlignmentAngle As Double
    Dim InverseSkewDirection As Boolean ' True for Inverse, else False
    
    '***** Tool Dimensions *****
    
    'BottomPlate
    Dim BottomPlateID As Double
    Dim BottomPlateScrewsD As Double
    Dim BottomPlateSize As Double
    Dim BottomPlatePinLocationD As Double
    Dim BottomPlatePinD As Double
    
    'Plate
    Dim PlateSize As Double
    Dim PlateScrewsR As Double
    Dim PlateID As Double
    Dim PlateSlotLocationD As Double
    Dim PlateSlotShiftAngle As Double
    Dim PlateThickness As Double
    Dim PlateSlotD As Double
    Dim PlateSlotAngle As Double
    Dim PlateScrewAngle As Double
    
    'Mandrel
    Dim MandrelOD As Double
    Dim MandrelID As Double
    Dim MandrelHeight As Double
    Dim MandrelScrewsD As Double
    
    'Location Rod
    Dim RodD As Double
    Dim RodL As Double
    
    'Press Cup
    Dim PressCupOD As Double
    Dim PressCupID As Double
    Dim PressCupSocketLocation As Double
    Dim PressSocketAngle As Double
    Dim PressPinLocation As Double
    Dim PressPinD As Double
    
    'Teflon
    Dim TeflonID As Double
    Dim TeflonOD As Double
    Dim TeflonSlotLocationD As Double
    Dim TeflonHoleD As Double
    
    'Grinding Mandrel
    Dim GrindingMandrelCoreID As Double
    Dim GrindingMandrelCoreOD As Double
    Dim GrindingMandrelLength As Double
    Dim GrindingMandrelPinLocationD As Double
    Dim GrindingMandrelPinD As Double
    
    '***********************************************************************************************************
    '***********************************************************************************************************
    ToolAssemblyFolder = "C:\Users\FIlie\Documents\Felix Documents IPS\Master Tooling\PM Stator Stacking NEW\"
    ToolAssemblyPath = ToolAssemblyFolder + "Assem.SLDASM"
    '***********************************************************************************************************
    UnitType = "Agusta 609 AC" ' "Agusta 609 DC", "Agusta 169', "Latitude", "Agusta 609 AC"
    '***********************************************************************************************************
    '***********************************************************************************************************
    
    InverseSkewDirection = False
    
    Select Case UnitType

        Case "Agusta 609 AC"
        
'            CoreName = "1034-21-05972 Lamination, Stator, Main-1"
        
            NumberOfSlots = 36
            LamMinOD = 4.281
            LamMinID = 3.258
            LamThickness = 0.014
            CoreHeight = 0.378 'Average
            LamSlotLocationD = 3.7
            LamSlotMinWidth = 0.231
            
'        Case "Latitude"
'
'            CoreName = "1015-21-04311 Lamination, Stator, Main-1"
'
'            NumberOfSlots = 48
'            LamMinOD = 5.363
'            LamMinID = 3.423
'            LamThickness = 0.014
'            CoreHeight = 1.978
'            LamSlotLocationD = 2.07 * 2
'            LamSlotMinWidth = 0.166
'
'        Case "Agusta 169"
'
'            CoreName = "1015-21-01116 Assembly, Core, Stator, Main-1"
'
'            NumberOfSlots = 48
'            LamMinOD = 5.138
'            LamMinID = 4.16
'            LamThickness = 0.014
'            CoreHeight = 3.048
'            LamSlotLocationD = 4.4
'            LamSlotMinWidth = 0.15
'
'        Case "Agusta 609 DC"
'
'            CoreName = "1034-21-07119 Lamination, Stator, Main-1"
'
'            NumberOfSlots = 72
'            LamMinOD = 6.667
'            LamMinID = 5.78
'            LamThickness = 0.014
'            CoreHeight = 2.067
'            LamSlotLocationD = 6.06
'            LamSlotMinWidth = 0.158
'            InverseSkewDirection = True

        Case Else
            swApp.SendMsgToUser2 "Data for this unit is not available", swMbStop, swMbOk
            errors = False
            Exit Sub
    End Select
    
    '***** Calculating Tool Dimensions *****
    
    'Rod
    RodD = LamSlotMinWidth - 0.003
    RodL = 2 'Round(CoreHeight + 2 * PlateThickness + 0.5, 1) 'RodL@Boss-Extrude1
    
    Debug.Print RodD, RodL
    
    'Bottom Plate
    BottomPlateID = LamMinID + 0.001 'BottomPlateID@Sketch2
    BottomPlateScrewsD = Round(BottomPlateID - 0.5, 2) 'BottomPlateScrewsD@Sketch6
    BottomPlateSize = Round(LamMinOD + 0.7, 1) 'BottomPlateSize@Sketch2
    BottomPlatePinLocationD = LamSlotLocationD 'BottomPlatePinLocationD@Main Sketch
    BottomPlatePinD = RodD - 0.0005 'BottomPlatePinD@Main Sketch
    
    Debug.Print BottomPlateID, BottomPlateScrewsD, BottomPlateSize, BottomPlatePinLocationD, BottomPlatePinD
    
    'Plate
    PlateThickness = 0.375
    PlateSize = Round(LamMinOD + 0.05, 2) 'PlateSize@Sketch2
    PlateID = LamMinID + 0.015 'PlateID@Sketch2
    PlateScrewsR = Round(LamMinOD / 2 + 0.3, 1) 'PlateScrewsR@Sketch1
    PlateSlotLocationD = LamSlotLocationD 'PlateSlotLocationD@Sketch1, PlateSlotLocationD@Sketch15
    PlateSlotD = LamSlotMinWidth + 0.005 'PlateSlotD@Sketch1
    PlateScrewAngle = 45 'PlateScrewAngle@Sketch1

    
    Debug.Print PlateSize, PlateID, PlateScrewsR, PlateSlotLocationD, PlateSlotD, PlateSlotD,
    
    'Mandrel
    MandrelHeight = Round(CoreHeight + PlateThickness * 2 + 0.1, 1) 'MandrelHeight@Boss-Extrude1
    MandrelOD = LamMinID - 0.001 'MandrelOD@Sketch3
    MandrelID = Round(MandrelOD - 1, 1) 'MandrelID@Sketch3
    MandrelScrewsD = BottomPlateScrewsD 'MandrelScrewsD@Sketch4
    
    Debug.Print MandrelHeight, MandrelOD, MandrelID, MandrelScrewsD
    
    'Press Cup
    PressCupID = Round(LamMinID + 0.02, 2)
    PressCupOD = Round(PressCupID + 1, 1)
    PressCupSocketLocation = 2 * PlateScrewsR
    PressSocketAngle = 45 'PressSocketAngle@Sketch4
    PressPinLocation = LamSlotLocationD
    PressPinD = LamSlotMinWidth + 0.01
    
    'Teflon
    TeflonID = LamMinID + 0.015
    TeflonOD = Round(LamMinOD + 0.1, 2)
    TeflonSlotLocationD = LamSlotLocationD
    TeflonHoleD = LamSlotMinWidth + 0.03
    
    'Grinding Mandrel
    GrindingMandrelCoreID = LamMinID - 0.0015 'GrindingMandrelCoreID@Sketch1
    GrindingMandrelCoreOD = LamMinOD - 0.1 'GrindingMandrelCoreOD@Sketch1
    GrindingMandrelLength = CoreHeight - 0.05 'GrindingMandrelLength@Sketch1
    GrindingMandrelPinLocationD = LamSlotLocationD 'GrindingMandrelPinLocationD@Sketch2
    GrindingMandrelPinD = RodD - 0.0005 'GrindingMandrelPinD@Sketch2
    
    '***** Changing Tool Dimensions *****
    ' DONT FORGET TO CONVERT TO METERS!!!!
    ' DEG in RADIANS!!!!
    
    Set swModel = swApp.OpenDoc6(ToolAssemblyPath, _
    swDocASSEMBLY, swOpenDocOptions_Silent, "", lErrors, lWarnings)

    'Bottom Plate
    swApp.ActivateDoc "Bottom Plate"
    Set swModel = swApp.ActiveDoc
    swModel.Parameter("BottomPlateID@Sketch2").SystemValue = BottomPlateID * inTOmeter
    swModel.Parameter("BottomPlateScrewsD@Sketch6").SystemValue = BottomPlateScrewsD * inTOmeter
    swModel.Parameter("BottomPlateSize@Sketch2").SystemValue = BottomPlateSize * inTOmeter
    swModel.Parameter("BottomPlatePinLocationD@Main Sketch").SystemValue = BottomPlatePinLocationD * inTOmeter
    swModel.Parameter("BottomPlatePinD@Main Sketch").SystemValue = BottomPlatePinD * inTOmeter
    swModel.EditRebuild3
    swModel.Save3 swSaveAsOptions_Silent, lErrors, lWarnings
    
    'Plate
    swApp.ActivateDoc "Plate"
    Set swModel = swApp.ActiveDoc
    Set swPart = swModel
    
    swModel.Parameter("PlateSize@Sketch2").SystemValue = PlateSize * inTOmeter
    swModel.Parameter("PlateID@Sketch2").SystemValue = PlateID * inTOmeter
    swModel.Parameter("PlateScrewsR@Sketch1").SystemValue = PlateScrewsR * inTOmeter
    swModel.Parameter("PlateSlotLocationD@Sketch1").SystemValue = PlateSlotLocationD * inTOmeter
    swModel.Parameter("PlateSlotD@Sketch1").SystemValue = PlateSlotD * inTOmeter
    swModel.Parameter("PlateThickness@Boss-Extrude1@Sketch1").SystemValue = PlateThickness * inTOmeter
    swModel.Parameter("PlateScrewAngle@Sketch1").SystemValue = PlateScrewAngle * degTORad 'ANGLE!
    
    Set swModelDocExt = swModel.Extension
    boolstatus = swModelDocExt.SelectByID2("CirPattern3", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
    Set swSelectionMgr = swModel.SelectionManager
    Set swFeature = swSelectionMgr.GetSelectedObject6(1, -1)
    Set swCircularPatternFeatureData = swFeature.GetDefinition
    ' Get or sets the number of instances in the circular-pattern feature
    swCircularPatternFeatureData.AccessSelections swModel, Nothing
    swCircularPatternFeatureData.TotalInstances = NumberOfSlots
    skippedItemsArray = swCircularPatternFeatureData.SkippedItemArray
    swFeature.ModifyDefinition swCircularPatternFeatureData, swModel, Nothing
    
    swModel.EditRebuild3
    swModel.Save3 swSaveAsOptions_Silent, lErrors, lWarnings

    'Mandrel
    swApp.ActivateDoc "Mandrel"
    Set swModel = swApp.ActiveDoc
    swModel.Parameter("MandrelHeight@Boss-Extrude1").SystemValue = MandrelHeight * inTOmeter
    swModel.Parameter("MandrelOD@Sketch3").SystemValue = MandrelOD * inTOmeter
    swModel.Parameter("MandrelID@Sketch3").SystemValue = MandrelID * inTOmeter
    swModel.Parameter("MandrelScrewsD@Sketch4").SystemValue = MandrelScrewsD * inTOmeter
    swModel.EditRebuild3
    swModel.Save3 swSaveAsOptions_Silent, lErrors, lWarnings
    
    'Alignment Rod
    swApp.ActivateDoc "Alignment Rod"
    Set swModel = swApp.ActiveDoc
    swModel.Parameter("RodD@Sketch1").SystemValue = RodD * inTOmeter
    swModel.Parameter("RodL@Boss-Extrude1").SystemValue = RodL * inTOmeter
    swModel.EditRebuild3
    swModel.Save3 swSaveAsOptions_Silent, lErrors, lWarnings
    
    'Press Cup
    swApp.ActivateDoc "Press Cup"
    Set swModel = swApp.ActiveDoc
    Set swPart = swModel
    
    swModel.Parameter("PressCupOD@Sketch1").SystemValue = PressCupOD * inTOmeter
    swModel.Parameter("PressCupSocketLocation@Sketch4").SystemValue = PressCupSocketLocation * inTOmeter
    swModel.Parameter("PressPinLocation@Sketch4").SystemValue = PressPinLocation * inTOmeter
    swModel.Parameter("PressPinD@Sketch4").SystemValue = PressPinD * inTOmeter
    swModel.EditRebuild3
    swModel.Save3 swSaveAsOptions_Silent, lErrors, lWarnings
    
    'Teflon
    Set swModel = swApp.OpenDoc6(ToolAssemblyFolder + "Teflon.SLDPRT", _
    swDocPART, swOpenDocOptions_Silent, "", lErrors, lWarnings)
    swApp.ActivateDoc "Teflon"
    Set swModel = swApp.ActiveDoc
    
    'This part is used to control the number of instances in the circular pattern
    Set swModelDocExt = swModel.Extension
    boolstatus = swModelDocExt.SelectByID2("CirPattern1", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
    Set swSelectionMgr = swModel.SelectionManager
    Set swFeature = swSelectionMgr.GetSelectedObject6(1, -1)
    Set swCircularPatternFeatureData = swFeature.GetDefinition
    ' Get or sets the number of instances in the circular-pattern feature
    swCircularPatternFeatureData.TotalInstances = NumberOfSlots
    'After updating Feature you must use ModifyDefinition, so changes will take place!
    swFeature.ModifyDefinition swCircularPatternFeatureData, swModel, Nothing
    
    swModel.Parameter("TeflonOD@Sketch2").SystemValue = TeflonOD * inTOmeter
    swModel.Parameter("TeflonID@Sketch2").SystemValue = TeflonID * inTOmeter
    swModel.Parameter("TeflonSlotLocationD@Sketch1").SystemValue = TeflonSlotLocationD * inTOmeter
    swModel.Parameter("TeflonHoleD@Sketch3").SystemValue = TeflonHoleD * inTOmeter
    
    swModel.EditRebuild3
    swModel.Save3 swSaveAsOptions_Silent, lErrors, lWarnings
    
    'GrindingMandrel
    Set swModel = swApp.OpenDoc6(ToolAssemblyFolder + "GrindingMandrel.SLDPRT", _
    swDocPART, swOpenDocOptions_Silent, "", lErrors, lWarnings)
    swApp.ActivateDoc "GrindingMandrel"
    Set swModel = swApp.ActiveDoc
    
    swModel.Parameter("GrindingMandrelCoreID@Sketch1").SystemValue = GrindingMandrelCoreID * inTOmeter
    swModel.Parameter("GrindingMandrelCoreOD@Sketch1").SystemValue = GrindingMandrelCoreOD * inTOmeter
    swModel.Parameter("GrindingMandrelLength@Sketch1").SystemValue = GrindingMandrelLength * inTOmeter
    swModel.Parameter("GrindingMandrelPinLocationD@Sketch2").SystemValue = GrindingMandrelPinLocationD * inTOmeter
    swModel.Parameter("GrindingMandrelPinD@Sketch2").SystemValue = GrindingMandrelPinD * inTOmeter
    
    swModel.EditRebuild3
    swModel.Save3 swSaveAsOptions_Silent, lErrors, lWarnings
    
    'Assembly
    swApp.ActivateDoc "Assem"
    Set swModel = swApp.ActiveDoc
    Set swAssy = swModel
    
'    'Unsuppress only the relevant core
'    'First suppress all the cores:
'    For i = 0 To UBound(CoreNamesArray)
'        Set swComp = swAssy.GetComponentByName(CoreNamesArray(i))
'        swComp.SetSuppression2 swComponentSuppressed
'    Next i
'    'Next, unsuppress only the relevant core
'    Set swComp = swAssy.GetComponentByName(CoreName)
'    swComp.SetSuppression2 swComponentFullyResolved
'
'    'This part is used to control the number of instances in the circular pattern for ASSEMBLY
'    Set swFeature = swAssy.FeatureByName("LocalCirPattern1")
'    If swFeature Is Nothing Then Debug.Print "swFeature is nothing"
'    Set swLocCircPatt = swFeature.GetDefinition 'Might be beacuse I forgot the SET???? Or didn't work beacuse of SelectByID2???
'    swLocCircPatt.TotalInstances = NumberOfTabs
'    swFeature.ModifyDefinition swLocCircPatt, swModel, Nothing
    
    swModel.Extension.Rebuild swForceRebuildAll
    'EditRebuild3
    swModel.Save3 swSaveAsOptions_Silent, lErrors, lWarnings

End Sub






