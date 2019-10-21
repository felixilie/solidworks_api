Attribute VB_Name = "ExciterStatorStackingNew"
Option Explicit

Const PI = 3.14159265358979

Sub ExciterStatorStackingNew()

    Dim swApp As SldWorks.SldWorks
    Dim swModel As SldWorks.ModelDoc2
    Dim swModelDocExt As SldWorks.ModelDocExtension
    Dim swAssy As SldWorks.AssemblyDoc
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
    
    Dim inTOmeter As Double
    Dim meterToin As Double
    Dim degTORad As Double
    Dim radToDeg As Double
    
    CopyCodeFile
    
    Set swApp = Application.SldWorks
    
    inTOmeter = 0.0254 'Units are in meters
    meterToin = 1 / inTOmeter 'To convert back to inch for checking results
    degTORad = PI / 180 'From degrees to Radians
    radToDeg = 180 / PI
    
    'Part Properties
    Dim NumberOfSlots As Integer
    Dim NumberOfTabs As Integer
    Dim LamMinOD As Double ' Without Tabs
    Dim LamMinID As Double
    Dim LamThickness As Double
    Dim LamSlotLocationD As Double
    Dim LamPoleMaxWidth As Double
    'Dim LamSlotMinWidth As Double
    Dim CoreHeight As Double 'Mid value
    
    'Dim AlignmentAngle As Double
    
    '***** Tool Dimensions *****
    
    'BottomPlate
    Dim BottomPlateID As Double
    Dim BottomPlateScrewsD As Double
    Dim BottomPlateSize As Double
    Dim BottomPlatePinLocationD As Double
    Dim BottomPlatePinWidth As Double
    
    'Plate
    Dim PlateSize As Double
    Dim PlateScrewsR As Double
    Dim PlateID As Double
    'Dim PlateSlotLocationD As Double
    'Dim PlateSlotShiftAngle As Double
    Dim PlateThickness As Double
    'Dim PlateSlotD As Double
    Dim PlateSlotAngle As Double
    Dim PlatePinLocationD As Double
    Dim PlatePinWidth As Double
    Dim ScrewAngle As Double 'ScrewAngle@Sketch1
    
    'Mandrel
    Dim MandrelOD As Double
    Dim MandrelID As Double
    Dim MandrelHeight As Double
    Dim MandrelScrewsD As Double
    
    'Location Pin
    Dim PinD As Double
    'Dim RodL As Double
    
    'Press Cup
    Dim PressCupOD As Double
    'Dim PressCupID As Double
    Dim PressCupSocketLocation As Double
    Dim PressCupLocatingOD As Double
    
    'Cement Plate
    Dim CementPlateHoleD As Double
    Dim CementPlateOD As Double
    Dim CementPlateHoleLocation As Double
    Dim CementPlateSlotLocationD As Double
    
    'Teflon
    Dim TeflonID As Double
    Dim TeflonOD As Double
    Dim TeflonHoleLocation As Double
    Dim TeflonSlotLocationD As Double
    Dim TeflonHoleD As Double
    
    'Grinding Mandrel
    Dim GrindingMandrelCoreID As Double
    Dim GrindingMandrelCoreOD As Double
    Dim GrindingMandrelLength As Double
    Dim GrindingMandrelPinWidth As Double
    Dim GrindingMandrelPinLocationD As Double
    Dim GrindingMandrelPinD As Double
    
    ToolAssemblyFolder = "C:\Users\FIlie\Documents\Felix Documents IPS\Master Tooling\Exciter Stator Stacking NEW\"
    ToolAssemblyPath = ToolAssemblyFolder + "Assem.SLDASM"
    
    UnitType = "CH" ' "Agusta 609 DC", "Agusta 169', "Latitude", "Agusta 609 AC"
    
    Select Case UnitType

        Case "Agusta 609 AC"
        
            NumberOfSlots = 12
            NumberOfTabs = 4
            LamMinOD = 5.748
            LamMinID = 4.184
            LamThickness = 0.014
            CoreHeight = 0.475
            LamPoleMaxWidth = 0.282
            
'        Case "Latitude"
'
'            NumberOfSlots = 48
'            NumberOfTabs = 4
'            LamMinOD = 5.363
'            LamMinID = 3.423
'            LamThickness = 0.014
'            CoreHeight = 1.978
'            LamSlotLocationD = 2.07 * 2
'            LamSlotMinWidth = 0.166
'
        Case "Agusta 169"

            NumberOfSlots = 10
            NumberOfTabs = 5
            LamMinOD = 5.366
            LamMinID = 3.998
            LamThickness = 0.014
            CoreHeight = 0.591
            LamPoleMaxWidth = 0.309

        Case "Textron"

            NumberOfSlots = 10
            NumberOfTabs = 5
            LamMinOD = 8.05
            LamMinID = 6.215
            LamThickness = 0.014
            CoreHeight = 0.853
            LamPoleMaxWidth = 0.388
            
        Case "CH"

            NumberOfSlots = 8
            NumberOfTabs = 4
            LamMinOD = 5.346
            LamMinID = 4.344
            LamThickness = 0.014
            CoreHeight = 0.375
            LamPoleMaxWidth = 0.452
            ScrewAngle = 22.5

        Case Else
            swApp.SendMsgToUser2 "Data for this unit is not available", swMbStop, swMbOk
            errors = False
            Exit Sub
    End Select
    
    PinD = 0.25
    
    'AlignmentAngle = radToDeg * _
    'Atn((LamMinID + 0.002) / 2 * Sin(degTORad * 360 / NumberOfSlots) / _
    'Sin(degTORad * (180 - 360 / NumberOfSlots) / 2) / CoreHeight)
    
    'AlignmentAngle = Round(AlignmentAngle, 3)
    
    'RodD = Round(LamSlotMinWidth - Tan(degTORad * AlignmentAngle) * LamThickness * Cos(radToDeg * AlignmentAngle) - 0.0025, 3)
    
    'Debug.Print AlignmentAngle, RodD
    
    
    '***** Calculating Tool Dimensions *****
    
    'Bottom Plate
    BottomPlateID = LamMinID + 0.002 'BottomPlateID@Sketch2
    BottomPlateScrewsD = Round(BottomPlateID - 0.5, 2) 'BottomPlateScrewsD@Sketch6
    BottomPlateSize = Round(LamMinOD + 0.7, 2) 'BottomPlateSize@Sketch2
    If UnitType = "Agusta 609 AC" Then BottomPlateSize = BottomPlateSize + 0.1
    BottomPlatePinLocationD = Round(LamMinID + (LamMinOD - LamMinID) / 2, 2)
    BottomPlatePinWidth = LamPoleMaxWidth + 0.002 + PinD
    
    Debug.Print BottomPlateID, BottomPlateScrewsD, BottomPlateSize, BottomPlatePinLocationD, BottomPlatePinWidth
    
    'Plate
    PlateThickness = 0.5
    PlateSize = Round(LamMinOD - 0.15, 2) 'PlateSize@Sketch2
    PlateID = LamMinID + 0.015 'PlateID@Sketch2
    'PlateScrewsR = Round(LamMinID / 2 + (LamMinOD - LamMinID) / 4, 2) 'PlateScrewsR@Sketch1
    'PlateSlotShiftAngle = 360 / NumberOfSlots * PlateThickness / (CoreHeight - LamThickness) 'PlateSlotShiftAngle@Sketch20 , PlateSlotShiftAngle@Sketch19
    'PlateSlotD = LamSlotMinWidth + 0.01 'PlateSlotD@Sketch20, PlateSlotD@Sketch19
    PlateSlotAngle = 360 / NumberOfSlots 'PlateSlotAngle@Sketch1, PlateSlotAngle@Sketch15
    PlatePinLocationD = Round(LamMinID + (LamMinOD - LamMinID) / 2, 2)
    PlateScrewsR = PlatePinLocationD / 2 + 0.1
    PlatePinWidth = LamPoleMaxWidth + 0.002 + PinD
    
    Debug.Print PlateSize, PlateID, PlateScrewsR, PlateSlotAngle, PlatePinLocationD, PlatePinWidth
    
    'Mandrel
    MandrelHeight = Round(CoreHeight + 1, 1)  'MandrelHeight@Boss-Extrude1
    MandrelOD = LamMinID - 0.001 'MandrelOD@Sketch3
    MandrelID = Round(MandrelOD - 1, 1) 'MandrelID@Sketch3
    MandrelScrewsD = BottomPlateScrewsD 'MandrelScrewsD@Sketch4
    
    Debug.Print MandrelHeight, MandrelOD, MandrelID, MandrelScrewsD
    
    'Rod
    'RodD already given above 'RodD@Sketch1
    'RodL = Round(CoreHeight + 2 * PlateThickness + 0.5, 1) 'RodL@Boss-Extrude1
    
    'Debug.Print RodD, RodL
    
    'Press Cup
    'PressCupID = Round(LamMinID + 0.02, 2)
    PressCupOD = Round(LamMinOD + 0.15, 1)
    PressCupSocketLocation = 2 * PlateScrewsR
    PressCupLocatingOD = LamMinID - 0.02 'PressCupLocatingOD@Sketch4
    
    'Cementing Plate
    CementPlateHoleD = 0.375
    CementPlateOD = Round(LamMinOD + 0.1, 2)
    CementPlateHoleLocation = LamMinID - 0.375 - 0.03
    CementPlateSlotLocationD = Round(LamMinID + (LamMinOD - LamMinID) / 2, 2)
    
    'Teflon
    TeflonID = LamMinID - 2 * 0.375 - 0.3
    TeflonOD = Round(LamMinOD + 0.1, 2)
    TeflonHoleLocation = LamMinID - 0.375 - 0.03
    TeflonSlotLocationD = Round(LamMinID + (LamMinOD - LamMinID) / 2, 2)
    TeflonHoleD = 0.375
    
    'Grinding Mandrel
    GrindingMandrelCoreID = LamMinID + 0.03 'GrindingMandrelCoreID@Sketch1, grinding to exact dimension is done later
    GrindingMandrelCoreOD = LamMinOD - 0.1 'GrindingMandrelCoreOD@Sketch1
    GrindingMandrelLength = CoreHeight - 0.05 'GrindingMandrelLength@Sketch1
    GrindingMandrelPinWidth = PlatePinWidth 'GrindingMandrelPinWidth@Sketch2
    GrindingMandrelPinLocationD = Round(LamMinID + (LamMinOD - LamMinID) / 2, 2) 'GrindingMandrelPinLocationD@Sketch2
    GrindingMandrelPinD = PinD - 0.0005 'GrindingMandrelPinD@Sketch2
    
    
    '***** Changing Tool Dimensions *****
    ' DONT FORGET TO CONVERT TO METERS!!!!
    ' DEG in RADIANS!!!!
    
    Set swModel = swApp.OpenDoc6(ToolAssemblyPath, _
    swDocASSEMBLY, swOpenDocOptions_Silent, "", lErrors, lWarnings)

    'Bottom Plate
    swApp.ActivateDoc "Bottom Plate"
    Set swModel = swApp.ActiveDoc
    
    'This part is used to control the number of instances in the circular pattern
    Set swModelDocExt = swModel.Extension
    boolstatus = swModelDocExt.SelectByID2("CirPattern1", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
    Set swSelectionMgr = swModel.SelectionManager
    Set swFeature = swSelectionMgr.GetSelectedObject6(1, -1)
    Set swCircularPatternFeatureData = swFeature.GetDefinition
    ' Get or sets the number of instances in the circular-pattern feature
    swCircularPatternFeatureData.TotalInstances = NumberOfTabs
    'After updating Feature you must use ModifyDefinition, so changes will take place!
    swFeature.ModifyDefinition swCircularPatternFeatureData, swModel, Nothing
    
    swModel.Parameter("BottomPlateID@Sketch2").SystemValue = BottomPlateID * inTOmeter
    swModel.Parameter("BottomPlateScrewsD@Sketch6").SystemValue = BottomPlateScrewsD * inTOmeter
    swModel.Parameter("BottomPlateSize@Sketch2").SystemValue = BottomPlateSize * inTOmeter
    swModel.Parameter("BottomPlatePinLocationD@Sketch9").SystemValue = BottomPlatePinLocationD * inTOmeter
    swModel.Parameter("BottomPlatePinWidth@Sketch9").SystemValue = BottomPlatePinWidth * inTOmeter
    
    swModel.EditRebuild3
    swModel.Save3 swSaveAsOptions_Silent, lErrors, lWarnings

    'Plate
    swApp.ActivateDoc "Plate"
    Set swModel = swApp.ActiveDoc
    
    'This part is used to control the number of instances in the circular pattern
    Set swModelDocExt = swModel.Extension
    boolstatus = swModelDocExt.SelectByID2("CirPattern1", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
    Set swSelectionMgr = swModel.SelectionManager
    Set swFeature = swSelectionMgr.GetSelectedObject6(1, -1)
    Set swCircularPatternFeatureData = swFeature.GetDefinition
    ' Get or sets the number of instances in the circular-pattern feature
    swCircularPatternFeatureData.TotalInstances = NumberOfTabs
    'After updating Feature you must use ModifyDefinition, so changes will take place!
    swFeature.ModifyDefinition swCircularPatternFeatureData, swModel, Nothing

    boolstatus = swModelDocExt.SelectByID2("CirPattern2", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
    Set swSelectionMgr = swModel.SelectionManager
    Set swFeature = swSelectionMgr.GetSelectedObject6(1, -1)
    Set swCircularPatternFeatureData = swFeature.GetDefinition
    ' Get or sets the number of instances in the circular-pattern feature
    swCircularPatternFeatureData.TotalInstances = NumberOfTabs
    'After updating Feature you must use ModifyDefinition, so changes will take place!
    swFeature.ModifyDefinition swCircularPatternFeatureData, swModel, Nothing

    boolstatus = swModelDocExt.SelectByID2("CirPattern5", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
    Set swSelectionMgr = swModel.SelectionManager
    Set swFeature = swSelectionMgr.GetSelectedObject6(1, -1)
    Set swCircularPatternFeatureData = swFeature.GetDefinition
    ' Get or sets the number of instances in the circular-pattern feature
    swCircularPatternFeatureData.TotalInstances = NumberOfTabs
    'After updating Feature you must use ModifyDefinition, so changes will take place!
    swFeature.ModifyDefinition swCircularPatternFeatureData, swModel, Nothing

    boolstatus = swModelDocExt.SelectByID2("CirPattern8", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
    Set swSelectionMgr = swModel.SelectionManager
    Set swFeature = swSelectionMgr.GetSelectedObject6(1, -1)
    Set swCircularPatternFeatureData = swFeature.GetDefinition
    ' Get or sets the number of instances in the circular-pattern feature
    swCircularPatternFeatureData.TotalInstances = NumberOfTabs
    'After updating Feature you must use ModifyDefinition, so changes will take place!
    swFeature.ModifyDefinition swCircularPatternFeatureData, swModel, Nothing
    
    swModel.Parameter("PlateSize@Sketch2").SystemValue = PlateSize * inTOmeter
    swModel.Parameter("PlateID@Sketch2@MainSketch").SystemValue = PlateID * inTOmeter
    swModel.Parameter("PlateScrewsR@Sketch1").SystemValue = PlateScrewsR * inTOmeter
    'swModel.Parameter("PlateSlotLocationD@Sketch1").SystemValue = PlateSlotLocationD * inTOmeter
    'swModel.Parameter("PlateSlotLocationD@Sketch15").SystemValue = PlateSlotLocationD * inTOmeter
    'swModel.Parameter("PlateSlotShiftAngle@Sketch20").SystemValue = PlateSlotShiftAngle * degTORad 'ANGLE!
    'swModel.Parameter("PlateSlotShiftAngle@Sketch19").SystemValue = PlateSlotShiftAngle * degTORad 'ANGLE!
    'swModel.Parameter("PlateSlotD@Sketch20").SystemValue = PlateSlotD * inTOmeter
    'swModel.Parameter("PlateSlotD@Sketch19").SystemValue = PlateSlotD * inTOmeter
    swModel.Parameter("PlateSlotAngle@Sketch1").SystemValue = PlateSlotAngle * degTORad 'ANGLE!
    swModel.Parameter("PlateSlotAngle@Sketch15").SystemValue = PlateSlotAngle * degTORad 'ANGLE!
    swModel.Parameter("PlatePinLocationD@Sketch24").SystemValue = PlatePinLocationD * inTOmeter
    swModel.Parameter("PlatePinWidth@Sketch24").SystemValue = PlatePinWidth * inTOmeter
    swModel.Parameter("ScrewAngle@Sketch1").SystemValue = ScrewAngle * degTORad 'ANGLE!
    
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
    
'    'Alignment Rod
'    swApp.ActivateDoc "Alignment Rod"
'    Set swModel = swApp.ActiveDoc
'    swModel.Parameter("RodD@Sketch1").SystemValue = RodD * inTOmeter
'    swModel.Parameter("RodL@Boss-Extrude1").SystemValue = RodL * inTOmeter
'    swModel.EditRebuild3
'    swModel.Save3 swSaveAsOptions_Silent, lErrors, lWarnings
    
    'Press Cup
    swApp.ActivateDoc "Press Cup"
    Set swModel = swApp.ActiveDoc
    
'    'This part is used to control the number of instances in the circular pattern
'    Set swModelDocExt = swModel.Extension
'    boolstatus = swModelDocExt.SelectByID2("CirPattern1", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
'    Set swSelectionMgr = swModel.SelectionManager
'    Set swFeature = swSelectionMgr.GetSelectedObject6(1, -1)
'    Set swCircularPatternFeatureData = swFeature.GetDefinition
'    ' Get or sets the number of instances in the circular-pattern feature
'    swCircularPatternFeatureData.TotalInstances = NumberOfSlots
'    'After updating Feature you must use ModifyDefinition, so changes will take place!
'    swFeature.ModifyDefinition swCircularPatternFeatureData, swModel, Nothing
    
    swModel.Parameter("PressCupLocatingOD@Sketch4").SystemValue = PressCupLocatingOD * inTOmeter
    swModel.Parameter("PressCupOD@Sketch1").SystemValue = PressCupOD * inTOmeter
    swModel.Parameter("PressCupSocketLocation@Sketch2").SystemValue = PressCupSocketLocation * inTOmeter
    swModel.EditRebuild3
    swModel.Save3 swSaveAsOptions_Silent, lErrors, lWarnings
     
    'Cementing Plate
    Set swModel = swApp.OpenDoc6(ToolAssemblyFolder + "Cementing Plate.SLDPRT", _
    swDocPART, swOpenDocOptions_Silent, "", lErrors, lWarnings)
    swApp.ActivateDoc "Cementing Plate"
    Set swModel = swApp.ActiveDoc
    
    'This part is used to control the number of instances in the circular pattern
    Set swModelDocExt = swModel.Extension
    boolstatus = swModelDocExt.SelectByID2("CirPattern1", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
    Set swSelectionMgr = swModel.SelectionManager
    Set swFeature = swSelectionMgr.GetSelectedObject6(1, -1)
    Set swCircularPatternFeatureData = swFeature.GetDefinition
    ' Get the number of instances in the circular-pattern feature
    swCircularPatternFeatureData.TotalInstances = NumberOfSlots
    'After updating Feature you must use ModifyDefinition, so changes will take place!
    swFeature.ModifyDefinition swCircularPatternFeatureData, swModel, Nothing
    
    swModel.Parameter("CementPlateHoleLocation@Sketch4").SystemValue = CementPlateHoleLocation * inTOmeter
    swModel.Parameter("CementPlateOD@Sketch4").SystemValue = CementPlateOD * inTOmeter
    swModel.Parameter("CementPlateSlotLocationD@Sketch3").SystemValue = CementPlateSlotLocationD * inTOmeter
    swModel.Parameter("CementPlateHoleD@Sketch5").SystemValue = CementPlateHoleD * inTOmeter
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
    
    swModel.Parameter("TeflonHoleLocation@Sketch2").SystemValue = TeflonHoleLocation * inTOmeter
    swModel.Parameter("TeflonOD@Sketch2").SystemValue = TeflonOD * inTOmeter
    swModel.Parameter("TeflonID@Sketch2").SystemValue = TeflonID * inTOmeter
    swModel.Parameter("TeflonSlotLocationD@Sketch3").SystemValue = TeflonSlotLocationD * inTOmeter
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
    swModel.Parameter("GrindingMandrelPinWidth@Sketch2").SystemValue = GrindingMandrelPinWidth * inTOmeter
    swModel.Parameter("GrindingMandrelPinLocationD@Sketch2").SystemValue = GrindingMandrelPinLocationD * inTOmeter
    swModel.Parameter("GrindingMandrelPinD@Sketch2").SystemValue = GrindingMandrelPinD * inTOmeter
    
    swModel.EditRebuild3
    swModel.Save3 swSaveAsOptions_Silent, lErrors, lWarnings
    
    'Assembly
    swApp.ActivateDoc "Assem"
    Set swModel = swApp.ActiveDoc
    Set swAssy = swModel
    
    'This part is used to control the number of instances in the circular pattern for ASSEMBLY
    Set swFeature = swAssy.FeatureByName("LocalCirPattern1")

    If swFeature Is Nothing Then Debug.Print "swFeature is nothing"

    Set swLocCircPatt = swFeature.GetDefinition 'Might be beacuse I forgot the SET???? Or didn't work beacuse of SelectByID2???

    swLocCircPatt.TotalInstances = NumberOfTabs

    swFeature.ModifyDefinition swLocCircPatt, swModel, Nothing
    
    swModel.Extension.Rebuild swForceRebuildAll
    'EditRebuild3
    swModel.Save3 swSaveAsOptions_Silent, lErrors, lWarnings

End Sub






