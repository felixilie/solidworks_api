Attribute VB_Name = "MainStatorStackingNew"
Option Explicit

Const PI = 3.14159265358979

Sub MainStatorStackingNew()

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
    Dim NumberOfSlots As Integer
    Dim NumberOfTabs As Integer
    Dim LamMinOD As Double ' Without Tabs
    Dim LamMinID As Double
    Dim LamThickness As Double
    Dim LamSlotLocationD As Double
    'Dim LamPoleMaxWidth As Double
    Dim LamSlotMinWidth As Double
    Dim CoreHeight As Double 'Mid value
    'Dim LamPolePinLocatorD As Double
    
    Dim AlignmentAngle As Double
    Dim InverseSkewDirection As Boolean ' True for Inverse, else False
    
    '***** Tool Dimensions *****
    
    'BottomPlate
    Dim BottomPlateID As Double
    Dim BottomPlateScrewsD As Double
    Dim BottomPlateSize As Double
    Dim BottomPlateJackScrewLocation As Double 'BottomPlateJackScrewLocation@Main Sketch
    'Dim BottomPlatePinLocationD As Double
    'Dim BottomPlatePinWidth As Double
    
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
    Dim PlateSlotThickness As Double
    Dim PlateSlotRotation As Double 'PlateSlotRotation@Sketch1
    'Dim PlatePinLocationD As Double
    'Dim PlatePinWidth As Double
    
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
    Dim GrindingMandrelPinShiftAngle As Double
    Dim GrindingMandrelPinLocationD As Double
    Dim GrindingMandrelPinD As Double
    
    '***********************************************************************************************************
    '***********************************************************************************************************
    ToolAssemblyFolder = "C:\Users\FIlie\Documents\Felix Documents IPS\Master Tooling\Main Stator Stacking NEW\"
    ToolAssemblyPath = ToolAssemblyFolder + "Assem.SLDASM"
    '***********************************************************************************************************
    UnitType = "SAAB" ' "Agusta 609 DC", "Agusta 169', "Latitude", "Agusta 609 AC", "SAAB", "A4"
    '***********************************************************************************************************
    '***********************************************************************************************************
    
    InverseSkewDirection = False
    PlateSlotRotation = 0
    
    Select Case UnitType

        Case "Agusta 609 AC"
        
            CoreName = "1034-21-05972 Lamination, Stator, Main-1"
        
            NumberOfSlots = 72
            NumberOfTabs = 4
            LamMinOD = 5.823
            LamMinID = 3.778
            LamThickness = 0.014
            CoreHeight = 3.014 'Average
            LamSlotLocationD = 4.162
            LamSlotMinWidth = 0.086
            PlateSlotRotation = 2.5
            
        Case "Latitude"
        
            CoreName = "1015-21-04311 Lamination, Stator, Main-1"

            NumberOfSlots = 48 '41.6 represents - Matt made a mistake,
            'the number used is the number of slot the will generate an alignment angle of 7.5
            NumberOfTabs = 4
            LamMinOD = 5.363
            LamMinID = 3.422 '3.423 Actually
            LamThickness = 0.014
            CoreHeight = 1.978
            LamSlotLocationD = 2.07 * 2
            LamSlotMinWidth = 0.166
            PlateSlotRotation = 3.5

        Case "Agusta 169"
        
            CoreName = "1015-21-01116 Assembly, Core, Stator, Main-1"

            NumberOfSlots = 48
            NumberOfTabs = 4
            LamMinOD = 5.138
            LamMinID = 4.161
            LamThickness = 0.014
            CoreHeight = 3.048
            LamSlotLocationD = 4.4
            LamSlotMinWidth = 0.15
            PlateSlotRotation = 3.5

        Case "Agusta 609 DC"
        
            CoreName = "1034-21-07119 Lamination, Stator, Main-1"

            NumberOfSlots = 72
            NumberOfTabs = 6
            LamMinOD = 6.667
            LamMinID = 5.78
            LamThickness = 0.014
            CoreHeight = 2.067
            LamSlotLocationD = 6.09 '6.06
            LamSlotMinWidth = 0.158
            InverseSkewDirection = True
            PlateSlotRotation = 358
            
        Case "SAAB"
            
            CoreName = "1021-21-02110 Lamination, Stator, Main-1"

            NumberOfSlots = 48
            NumberOfTabs = 4
            LamMinOD = 7.266
            LamMinID = 4.644
            LamThickness = 0.014
            CoreHeight = 2.33
            LamSlotLocationD = 5.374 - 0.08
            LamSlotMinWidth = 0.101
            
        Case "A4"
            
            'CoreName = "1032-21-03207 Assembly, Core, Stator, Main with Skew"

            NumberOfSlots = 72
            NumberOfTabs = 4
            LamMinOD = 8.868
            LamMinID = 7.288
            LamThickness = 0.014
            CoreHeight = 2.828
            LamSlotLocationD = 7.774
            LamSlotMinWidth = 0.172
            InverseSkewDirection = True
            
        Case "Textron"
            
            'CoreName = "1021-21-02110 Lamination, Stator, Main-1"

            NumberOfSlots = 96
            NumberOfTabs = 4
            LamMinOD = 9.765
            LamMinID = 7.634
            LamThickness = 0.014
            CoreHeight = 5.228
            LamSlotLocationD = 7.995
            LamSlotMinWidth = 0.093

        Case Else
            swApp.SendMsgToUser2 "Data for this unit is not available", swMbStop, swMbOk
            errors = False
            Exit Sub
    End Select
    
    'PinD = 0.25
    
    AlignmentAngle = radToDeg * _
    Atn((LamMinID + 0.002) / 2 * Sin(degTORad * 360 / NumberOfSlots) / _
    Sin(degTORad * (180 - 360 / NumberOfSlots) / 2) / CoreHeight)
    
    AlignmentAngle = Round(AlignmentAngle, 3)
    
'    If UnitType = "Latitude" Then AlignmentAngle = 7.5 'Latitude old fixture is screwed!
    
    RodD = Round(LamSlotMinWidth - Tan(degTORad * AlignmentAngle) * LamThickness * Cos(radToDeg * AlignmentAngle) - 0.0025, 3)
    
    Debug.Print AlignmentAngle, RodD
    
    
    '***** Calculating Tool Dimensions *****
    
    'Bottom Plate
    BottomPlateID = LamMinID + 0.002 'BottomPlateID@Sketch2
    BottomPlateScrewsD = Round(BottomPlateID - 0.5, 2) 'BottomPlateScrewsD@Sketch6
    BottomPlateSize = Round(LamMinOD + 0.7, 1) 'BottomPlateSize@Sketch2
    BottomPlateJackScrewLocation = LamSlotLocationD + 0.5
    'BottomPlatePinLocationD = Round(LamMinID + (LamMinOD - LamMinID) / 2, 2)
    'BottomPlatePinWidth = LamPoleMaxWidth + 0.002 + PinD
    
    Debug.Print BottomPlateID, BottomPlateScrewsD, BottomPlateSize ', BottomPlatePinLocationD, BottomPlatePinWidth
    
    'Plate
    PlateThickness = 0.5
    PlateSlotThickness = 0.2
    PlateSize = Round(LamMinOD - 0.08, 2) 'PlateSize@Sketch2 PlateSize = Round(LamMinOD - 0.04, 2)
    PlateID = LamMinID + 0.002 'PlateID@Sketch2
    PlateScrewsR = Round(LamMinOD / 2 + 0.3, 1) 'PlateScrewsR@Sketch1
    PlateSlotLocationD = LamSlotLocationD 'PlateSlotLocationD@Sketch1, PlateSlotLocationD@Sketch15
    PlateSlotShiftAngle = 360 / NumberOfSlots * PlateThickness / (CoreHeight - LamThickness) 'PlateSlotShiftAngle@Sketch20 , PlateSlotShiftAngle@Sketch19
    If InverseSkewDirection = True Then PlateSlotShiftAngle = 360 - PlateSlotShiftAngle
    PlateSlotD = LamSlotMinWidth + 0.005 'PlateSlotD@Sketch20, PlateSlotD@Sketch19
    PlateSlotAngle = 360 / NumberOfSlots 'PlateSlotAngle@Sketch1, PlateSlotAngle@Sketch15
    If InverseSkewDirection = True Then PlateSlotAngle = 360 - PlateSlotAngle
    PlateScrewAngle = (360 / NumberOfTabs) / 2 'PlateScrewAngle@Sketch1
    If NumberOfTabs = 6 Then
        PlateScrewsR = 3.531 'PlateScrewsR = ((PlateSize / 2) / Cos(PlateScrewAngle * degTORad) - LamMinOD) / 2 + LamMinOD
    End If
    'PlatePinLocationD = Round(LamMinID + (LamMinOD - LamMinID) / 2, 2)
    'PlatePinWidth = LamPoleMaxWidth + 0.002 + PinD
    
    Debug.Print PlateSize, PlateID, PlateScrewsR, PlateSlotD, PlateSlotShiftAngle, PlateSlotD, PlateSlotAngle, PlateScrewsR
    
    'Mandrel
    MandrelHeight = Round(CoreHeight + 2 * PlateThickness + 1, 1) 'MandrelHeight@Boss-Extrude1
    MandrelOD = LamMinID - 0.001 'MandrelOD@Sketch3
    MandrelID = Round(MandrelOD - 1, 1) 'MandrelID@Sketch3
    MandrelScrewsD = BottomPlateScrewsD 'MandrelScrewsD@Sketch4
    
    Debug.Print MandrelHeight, MandrelOD, MandrelID, MandrelScrewsD
    
    'Rod
    'RodD already given above 'RodD@Sketch1
    RodL = Round(CoreHeight + 2 * PlateThickness + 0.5, 1) 'RodL@Boss-Extrude1
    
    Debug.Print RodD, RodL
    
    'Press Cup
    PressCupID = Round(LamMinID + 0.02, 2)
    PressCupOD = Round(PressCupID + 1, 1)
    PressCupSocketLocation = 2 * PlateScrewsR
    PressSocketAngle = (360 / NumberOfTabs) / 2 'PressSocketAngle@Sketch4
    
    'Cementing Plate
    CementPlateHoleD = LamSlotMinWidth + 0.03
    CementPlateOD = Round(LamMinOD + 0.1, 2)
    CementPlateHoleLocation = LamMinID - 0.375 - 0.05
    CementPlateSlotLocationD = LamSlotLocationD
    
    'Teflon
    TeflonID = LamMinID - 2 * 0.375 - 0.3
    TeflonOD = Round(LamMinOD + 0.1, 2)
    TeflonHoleLocation = CementPlateHoleLocation
    TeflonSlotLocationD = LamSlotLocationD
    TeflonHoleD = LamSlotMinWidth + 0.03
    
    'Grinding Mandrel
    GrindingMandrelCoreID = LamMinID - 0.0015 'GrindingMandrelCoreID@Sketch1
    GrindingMandrelCoreOD = LamMinOD - 0.1 'GrindingMandrelCoreOD@Sketch1
    GrindingMandrelLength = CoreHeight - 0.05 'GrindingMandrelLength@Sketch1
    GrindingMandrelPinShiftAngle = 360 / NumberOfSlots * 0.15 / (CoreHeight - LamThickness) 'GrindingMandrelPinShiftAngle@Sketch2
    GrindingMandrelPinLocationD = LamSlotLocationD 'GrindingMandrelPinLocationD@Sketch2
    GrindingMandrelPinD = LamSlotMinWidth 'GrindingMandrelPinD@Sketch2
    
    Debug.Print "Grinding Mandrel" & vbNewLine
    Debug.Print GrindingMandrelPinShiftAngle
    
    
    '***** Changing Tool Dimensions *****
    ' DONT FORGET TO CONVERT TO METERS!!!!
    ' DEG in RADIANS!!!!
    
'    If UnitType = "Latitude" Then NumberOfSlots = 48 'Latitude old fixture is screwed!
    
    Set swModel = swApp.OpenDoc6(ToolAssemblyPath, _
    swDocASSEMBLY, swOpenDocOptions_OverrideDefaultLoadLightweight, "", lErrors, lWarnings)

    'Bottom Plate
    swApp.ActivateDoc "Bottom Plate"
    Set swModel = swApp.ActiveDoc
    swModel.Parameter("BottomPlateID@Sketch2").SystemValue = BottomPlateID * inTOmeter
    swModel.Parameter("BottomPlateScrewsD@Sketch6").SystemValue = BottomPlateScrewsD * inTOmeter
    swModel.Parameter("BottomPlateSize@Sketch2").SystemValue = BottomPlateSize * inTOmeter
    swModel.Parameter("BottomPlateJackScrewLocation@Main Sketch").SystemValue = BottomPlateJackScrewLocation * inTOmeter
    'swModel.Parameter("BottomPlatePinLocationD@Sketch9").SystemValue = BottomPlatePinLocationD * inTOmeter
    'swModel.Parameter("BottomPlatePinWidth@Sketch9").SystemValue = BottomPlatePinWidth * inTOmeter
    swModel.EditRebuild3
    swModel.Save3 swSaveAsOptions_Silent, lErrors, lWarnings

    'Plate
    swApp.ActivateDoc "Plate"
    Set swModel = swApp.ActiveDoc
    Set swPart = swModel
    
    'This part is used to control if the plate shape is square or hex (like for the augsta DC)
    Set swFeature = swPart.FeatureByName("Cut-Hex")
    If NumberOfTabs = 6 Then
        swFeature.SetSuppression2 swUnSuppressFeature, swAllConfiguration, Nothing
    Else
        swFeature.SetSuppression2 swSuppressFeature, swAllConfiguration, Nothing
    End If
    
    Set swFeature = swPart.FeatureByName("Boss-Hex")
    If NumberOfTabs = 6 Then
        swFeature.SetSuppression2 swUnSuppressFeature, swAllConfiguration, Nothing
    Else
        swFeature.SetSuppression2 swSuppressFeature, swAllConfiguration, Nothing
    End If
    
    swModel.Parameter("PlateSize@Sketch2").SystemValue = PlateSize * inTOmeter
    swModel.Parameter("PlateID@Sketch2").SystemValue = PlateID * inTOmeter
    swModel.Parameter("PlateScrewsR@Sketch1").SystemValue = PlateScrewsR * inTOmeter
    swModel.Parameter("PlateSlotLocationD@Sketch1").SystemValue = PlateSlotLocationD * inTOmeter
    swModel.Parameter("PlateSlotLocationD@Sketch15").SystemValue = PlateSlotLocationD * inTOmeter
    swModel.Parameter("PlateSlotShiftAngle@Sketch20").SystemValue = PlateSlotShiftAngle * degTORad 'ANGLE!
    swModel.Parameter("PlateSlotShiftAngle@Sketch19").SystemValue = PlateSlotShiftAngle * degTORad 'ANGLE!
    swModel.Parameter("PlateSlotD@Sketch20").SystemValue = PlateSlotD * inTOmeter
    swModel.Parameter("PlateSlotD@Sketch19").SystemValue = PlateSlotD * inTOmeter
    swModel.Parameter("PlateSlotAngle@Sketch1").SystemValue = PlateSlotAngle * degTORad 'ANGLE!
    swModel.Parameter("PlateSlotAngle@Sketch15").SystemValue = PlateSlotAngle * degTORad 'ANGLE!
    swModel.Parameter("PlateScrewAngle@Sketch1").SystemValue = PlateScrewAngle * degTORad 'ANGLE!
    swModel.Parameter("PlateSlotRotation@Sketch1").SystemValue = PlateSlotRotation * degTORad 'ANGLE!
    swModel.Parameter("PlateSlotRotation@Sketch15").SystemValue = PlateSlotRotation * degTORad 'ANGLE!
    'swModel.Parameter("PlatePinLocationD@Sketch24").SystemValue = PlatePinLocationD * inTOmeter
    'swModel.Parameter("PlatePinWidth@Sketch24").SystemValue = PlatePinWidth * inTOmeter
    swModel.Parameter("PlateThickness@Boss-Extrude1").SystemValue = PlateThickness * inTOmeter
    
        'This part is used to control the number of instances in the circular pattern
    Set swModelDocExt = swModel.Extension
    boolstatus = swModelDocExt.SelectByID2("CirPattern1", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
    Set swSelectionMgr = swModel.SelectionManager
    Set swFeature = swSelectionMgr.GetSelectedObject6(1, -1)
    Set swCircularPatternFeatureData = swFeature.GetDefinition
    ' Get or sets the number of instances in the circular-pattern feature
    swCircularPatternFeatureData.AccessSelections swModel, Nothing 'WASN'T REQUIRED
    swCircularPatternFeatureData.TotalInstances = NumberOfTabs
    skippedItemsArray = swCircularPatternFeatureData.SkippedItemArray
    For i = 0 To UBound(skippedItemsArray)
        Debug.Print skippedItemsArray(i)
    Next i
    'After updating Feature you must use ModifyDefinition, so changes will take place!
    swFeature.ModifyDefinition swCircularPatternFeatureData, swModel, Nothing

    boolstatus = swModelDocExt.SelectByID2("CirPattern2", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
    Set swSelectionMgr = swModel.SelectionManager
    Set swFeature = swSelectionMgr.GetSelectedObject6(1, -1)
    Set swCircularPatternFeatureData = swFeature.GetDefinition
    ' Get or sets the number of instances in the circular-pattern feature
    swCircularPatternFeatureData.AccessSelections swModel, Nothing
    swCircularPatternFeatureData.TotalInstances = NumberOfTabs
    skippedItemsArray = swCircularPatternFeatureData.SkippedItemArray
    For i = 0 To UBound(skippedItemsArray)
        Debug.Print skippedItemsArray(i)
    Next i
    'After updating Feature you must use ModifyDefinition, so changes will take place!
    swFeature.ModifyDefinition swCircularPatternFeatureData, swModel, Nothing

    boolstatus = swModelDocExt.SelectByID2("CirPattern5", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
    Set swSelectionMgr = swModel.SelectionManager
    Set swFeature = swSelectionMgr.GetSelectedObject6(1, -1)
    Set swCircularPatternFeatureData = swFeature.GetDefinition
    ' Get or sets the number of instances in the circular-pattern feature
    swCircularPatternFeatureData.AccessSelections swModel, Nothing
    swCircularPatternFeatureData.TotalInstances = NumberOfTabs
    skippedItemsArray = swCircularPatternFeatureData.SkippedItemArray
    For i = 0 To UBound(skippedItemsArray)
        Debug.Print skippedItemsArray(i)
    Next i
    'After updating Feature you must use ModifyDefinition, so changes will take place!
    swFeature.ModifyDefinition swCircularPatternFeatureData, swModel, Nothing

    boolstatus = swModelDocExt.SelectByID2("CirPattern6", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
    Set swSelectionMgr = swModel.SelectionManager
    Set swFeature = swSelectionMgr.GetSelectedObject6(1, -1)
    Set swCircularPatternFeatureData = swFeature.GetDefinition
    ' Get or sets the number of instances in the circular-pattern feature
    swCircularPatternFeatureData.AccessSelections swModel, Nothing
    swCircularPatternFeatureData.TotalInstances = NumberOfTabs
    skippedItemsArray = swCircularPatternFeatureData.SkippedItemArray
    For i = 0 To UBound(skippedItemsArray)
        Debug.Print skippedItemsArray(i)
    Next i
    'After updating Feature you must use ModifyDefinition, so changes will take place!
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
    
    Set swFeature = swPart.FeatureByName("Cut-Extrude2")
    If PressCupSocketLocation > PressCupOD + 0.8 Then
        swFeature.SetSuppression2 swSuppressFeature, swAllConfiguration, Nothing
    Else
        swFeature.SetSuppression2 swUnSuppressFeature, swAllConfiguration, Nothing
    End If
    
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
    
    swModel.Parameter("PressCupID@Sketch1").SystemValue = PressCupID * inTOmeter
    swModel.Parameter("PressCupOD@Sketch1").SystemValue = PressCupOD * inTOmeter
    swModel.Parameter("PressCupSocketLocation@Sketch4").SystemValue = PressCupSocketLocation * inTOmeter
    swModel.Parameter("PressSocketAngle@Sketch4").SystemValue = PressSocketAngle * degTORad
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
    swModel.Parameter("GrindingMandrelPinShiftAngle@Sketch2").SystemValue = GrindingMandrelPinShiftAngle * degTORad
    swModel.Parameter("GrindingMandrelPinLocationD@Sketch2").SystemValue = GrindingMandrelPinLocationD * inTOmeter
    swModel.Parameter("GrindingMandrelPinD@Sketch2").SystemValue = GrindingMandrelPinD * inTOmeter
    
    swModel.EditRebuild3
    swModel.Save3 swSaveAsOptions_Silent, lErrors, lWarnings
    
    'GrindingPlate
    Set swModel = swApp.OpenDoc6(ToolAssemblyFolder + "GrindingPlate.SLDPRT", _
    swDocPART, swOpenDocOptions_Silent, "", lErrors, lWarnings)
    swApp.ActivateDoc "GrindingPlate"
    Set swModel = swApp.ActiveDoc
    
    swModel.Parameter("OD@Sketch1@Sketch1").SystemValue = GrindingMandrelCoreOD * inTOmeter
    swModel.Parameter("ReliefOD@Sketch2").SystemValue = (GrindingMandrelCoreID + 0.1) * inTOmeter

    swModel.EditRebuild3
    swModel.Save3 swSaveAsOptions_Silent, lErrors, lWarnings
    
    'Assembly
    swApp.ActivateDoc "Assem"
    Set swModel = swApp.ActiveDoc
    Set swAssy = swModel
    
    
    'Unsuppress only the relevant core
    'First suppress all the cores:
    For i = 0 To UBound(CoreNamesArray)
        Set swComp = swAssy.GetComponentByName(CoreNamesArray(i))
        If Not swComp Is Nothing Then
            Debug.Print swComp.Name
            swComp.SetSuppression2 swComponentSuppressed
        End If
    Next i
    'Next, unsuppress only the relevant core
    Set swComp = swAssy.GetComponentByName(CoreName)
    swComp.SetSuppression2 swComponentFullyResolved
    
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




