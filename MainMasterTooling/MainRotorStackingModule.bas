Attribute VB_Name = "MainRotorStackingModule"
Option Explicit

Const PI = 3.14159265358979

Sub MainRotorStacking()

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
    
    CoreNamesArray = Array("1034-11-05942 Assembly, Core, Rotor, Main-1", _
                            "1034-11-07040 Assembly, Core, Rotor, Main-1")
    
    CopyCodeFile
    
    Set swApp = Application.SldWorks
    
    inTOmeter = 0.0254 'Units are in meters
    meterToin = 1 / inTOmeter 'To convert back to inch for checking results
    degTORad = PI / 180 'From degrees to Radians
    radToDeg = 180 / PI
    
    'Part Properties
    Dim CoreName As String
    Dim NumberOfPoles As Integer
    Dim LamMinID As Double
    Dim LamCopperRodsLoactionD As Double
    Dim LamCopperRodsD As Double
    Dim LamThickness As Double
    Dim LamPoleMaxWidth As Double
    Dim LamPoleLocationD As Double
    Dim CoreIDAfterGrind As Double 'Min
    Dim CoreHeight As Double 'Mid value
    
    '***** Tool Dimensions *****
    
    'General Tool Dimensions
    Dim ToolOD As Double 'ToolOD@Sketch1
    Dim ToolPoleWidth As Double
    Dim LocationPinD As Double
    Dim LocalCirNumInstances As Double
    Dim ToolScrewAngle As Double 'ToolScrewAngle@Sketch15
    Dim MaxCoreIDnoMandrelID As Double
    
    'Upper Base
    Dim UpperBaseID As Double 'UpperBaseID@Sketch2
    Dim UpperBasePinWidth As Double 'UpperBasePinWidth@Sketch6
    Dim UpperBasePinD As Double 'UpperBasePinD@Sketch6
'    Dim UpperBaseScrewToBaseLoactionD As Double 'UpperBaseScrewToBaseLoactionD@Sketch6
    Dim UpperBaseSmallOD As Double 'UpperBaseSmallOD@Sketch1
    Dim UpperBasePinLoacationD As Double 'UpperBasePinLoacationD@Sketch6
    'ToolOD@Sketch1
    
    'Top
    Dim TopSmallOD As Double 'TopSmallOD@Sketch1
    Dim TopID As Double 'TopID@Sketch2
'    Dim TopScrewToBaseLoactionD As Double 'TopScrewToBaseLoactionD@Sketch15
    Dim TopPinWidth As Double 'TopPinWidth@Sketch15
    Dim TopPinClearanceD As Double 'TopPinClearanceD@Sketch15
    Dim TopPinLocationD As Double 'TopPinLocationD@Sketch15
    'ToolOD@Sketch1
    
    'Mandrel
    Dim MandrelOD As Double 'MandrelID@Sketch1
    Dim MandrelODatBase As Double 'MandrelIDatBase@Sketch1
    Dim MandrelHeight As Double 'ManderlHeight@Sketch1
    Dim MandrelID As Double 'MandrelID@Sketch4
    Dim MandrelScrewLocation As Double 'MandrelScrewLocation@Sketch5
    
    'Base
    Dim BaseScrewLoactionD As Double 'BaseScrewLoactionD@Sketch5
    Dim BaseOD As Double 'BaseOD@Sketch1
    Dim BaseScrewLocation As Double 'BaseScrewLocation@Sketch9
    
    '***********************************************************************************************************
    '***********************************************************************************************************
    ToolAssemblyFolder = "C:\Users\FIlie\Documents\Felix Documents IPS\Master Tooling\Main Rotor Stacking\"
    ToolAssemblyPath = ToolAssemblyFolder + "Assembly, Stacking, Rotor, Main.SLDASM"
    '***********************************************************************************************************
    UnitType = "Agusta 609 DC" ' "Agusta 609 DC"
    '***********************************************************************************************************
    '***********************************************************************************************************
    
'    InverseSkewDirection = False
    
    Select Case UnitType

        Case "Agusta 609 AC"
        
            CoreName = "1034-11-05942 Assembly, Core, Rotor, Main-1"
        
            NumberOfPoles = 4
            LamMinID = 0.938
            LamThickness = 0.014
            LamCopperRodsLoactionD = 3.584
            LamCopperRodsD = 0.062
            LamPoleMaxWidth = 1.077
            LamPoleLocationD = 2.8
            CoreHeight = 3 'Average
            CoreIDAfterGrind = 0.95

        Case "Agusta 609 DC"
        
            CoreName = "1034-11-07040 Assembly, Core, Rotor, Main-1"

            NumberOfPoles = 12
            LamMinID = 3.795
            LamThickness = 0.014
            LamCopperRodsLoactionD = 5.656
            LamCopperRodsD = 0.046
            LamPoleMaxWidth = 0.416
            LamPoleLocationD = 4.85
            CoreHeight = 2.039 'Average
            CoreIDAfterGrind = 3.816

        Case Else
            swApp.SendMsgToUser2 "Data for this unit is not available", swMbStop, swMbOk
            errors = False
            Exit Sub
    End Select
    
    '***** Calculating Tool Dimensions *****
    
    'General Tool
    ToolOD = LamCopperRodsLoactionD - 2 * LamCopperRodsD - 0.01
    ToolPoleWidth = LamPoleMaxWidth + 0.002
    LocationPinD = 0.375
    MaxCoreIDnoMandrelID = 2
    
    If NumberOfPoles = 12 Then
        LocalCirNumInstances = 4
        ToolScrewAngle = 360 / NumberOfPoles + 360 / NumberOfPoles / 2
    End If
    If NumberOfPoles = 4 Then
        LocalCirNumInstances = 2
        ToolScrewAngle = 55
    End If
    
    Debug.Print ToolOD, ToolPoleWidth, LocationPinD
    
    'Upper Base
    UpperBaseID = CoreIDAfterGrind + 0.05
    UpperBasePinD = LocationPinD - 0.0005 'Press Fit
'    UpperBaseScrewToBaseLoactionD = Round(CoreIDAfterGrind + 0.05 + 1, 1)
    UpperBaseSmallOD = Round(ToolOD - 0.1, 2)
    UpperBasePinLoacationD = LamPoleLocationD
    UpperBasePinWidth = ToolPoleWidth
    
    Debug.Print UpperBaseID, UpperBasePinD, UpperBaseSmallOD, _
    UpperBasePinLoacationD, UpperBasePinWidth
    
    'Top
    TopID = CoreIDAfterGrind + 0.05
    TopSmallOD = Round(ToolOD - 0.1, 2)
'    TopScrewToBaseLoactionD = Round(CoreIDAfterGrind + 0.05 + 1, 1)
    TopPinWidth = ToolPoleWidth
    TopPinClearanceD = LocationPinD + 0.011
    TopPinLocationD = LamPoleLocationD
    
    Debug.Print TopID, TopSmallOD, TopPinWidth, TopPinClearanceD, TopPinLocationD
    
    'Mandrel
    MandrelOD = LamMinID - 0.001
    MandrelODatBase = CoreIDAfterGrind + 0.05 - 0.001
    MandrelHeight = 0.825 + 1.6 + CoreHeight - 0.1 'TopHeight + Upper Base Height + CoreHeight - .1
    
    If CoreIDAfterGrind > MaxCoreIDnoMandrelID Then
        MandrelID = Round(MandrelOD - 1.2, 1)
        MandrelScrewLocation = Round((MandrelOD - MandrelID) / 2 + MandrelID, 3)
    End If
    
    Debug.Print MandrelOD, MandrelODatBase, MandrelHeight
    
    'Base
    BaseScrewLoactionD = LamPoleLocationD
    BaseOD = Round(ToolOD - 0.1, 2)
    
    If CoreIDAfterGrind > MaxCoreIDnoMandrelID Then
        BaseScrewLocation = MandrelScrewLocation
    End If
    
    Debug.Print BaseScrewLoactionD, BaseOD
    
    '***** Changing Tool Dimensions *****
    ' DONT FORGET TO CONVERT TO METERS!!!!
    ' DEG in RADIANS!!!!
    
    Set swModel = swApp.OpenDoc6(ToolAssemblyPath, _
    swDocASSEMBLY, swOpenDocOptions_Silent, "", lErrors, lWarnings)

    'Upper Base
    swApp.ActivateDoc "Upper, Base, Fixture, Stacking, Rotor, Main"
    Set swModel = swApp.ActiveDoc
    swModel.Parameter("UpperBaseID@Sketch2").SystemValue = UpperBaseID * inTOmeter
    swModel.Parameter("UpperBasePinWidth@Sketch6").SystemValue = UpperBasePinWidth * inTOmeter
    swModel.Parameter("UpperBasePinD@Sketch6").SystemValue = UpperBasePinD * inTOmeter
'    swModel.Parameter("UpperBaseScrewToBaseLoactionD@Sketch6").SystemValue = UpperBaseScrewToBaseLoactionD * inTOmeter
    swModel.Parameter("UpperBaseSmallOD@Sketch1").SystemValue = UpperBaseSmallOD * inTOmeter
    swModel.Parameter("UpperBasePinLoacationD@Sketch6").SystemValue = UpperBasePinLoacationD * inTOmeter
    swModel.Parameter("ToolScrewAngle@Sketch6").SystemValue = ToolScrewAngle * degTORad
    swModel.Parameter("ToolOD@Sketch1").SystemValue = ToolOD * inTOmeter
    
    'This part is used to control the number of instances in the circular pattern
    Set swModelDocExt = swModel.Extension
    boolstatus = swModelDocExt.SelectByID2("CirPattern1", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
    Set swSelectionMgr = swModel.SelectionManager
    Set swFeature = swSelectionMgr.GetSelectedObject6(1, -1)
    Set swCircularPatternFeatureData = swFeature.GetDefinition
    ' Get the number of instances in the circular-pattern feature
    swCircularPatternFeatureData.TotalInstances = LocalCirNumInstances
    'After updating Feature you must use ModifyDefinition, so changes will take place!
    swFeature.ModifyDefinition swCircularPatternFeatureData, swModel, Nothing
    
    swModel.EditRebuild3
    swModel.Save3 swSaveAsOptions_Silent, lErrors, lWarnings
    
    'Top
    swApp.ActivateDoc "Top, Fixture, Stacking, Rotor, Main"
    Set swModel = swApp.ActiveDoc
    swModel.Parameter("TopSmallOD@Sketch1").SystemValue = TopSmallOD * inTOmeter
    swModel.Parameter("TopID@Sketch2").SystemValue = TopID * inTOmeter
'    swModel.Parameter("TopScrewToBaseLoactionD@Sketch15").SystemValue = TopScrewToBaseLoactionD * inTOmeter
    swModel.Parameter("TopPinWidth@Sketch15").SystemValue = TopPinWidth * inTOmeter
    swModel.Parameter("TopPinClearanceD@Sketch15").SystemValue = TopPinClearanceD * inTOmeter
    swModel.Parameter("TopPinLocationD@Sketch15").SystemValue = TopPinLocationD * inTOmeter
    swModel.Parameter("ToolScrewAngle@Sketch15").SystemValue = ToolScrewAngle * degTORad
    swModel.Parameter("ToolOD@Sketch1").SystemValue = ToolOD * inTOmeter
    
    'This part is used to control the number of instances in the circular pattern
    Set swModelDocExt = swModel.Extension
    boolstatus = swModelDocExt.SelectByID2("CirPattern1", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
    Set swSelectionMgr = swModel.SelectionManager
    Set swFeature = swSelectionMgr.GetSelectedObject6(1, -1)
    Set swCircularPatternFeatureData = swFeature.GetDefinition
    ' Get the number of instances in the circular-pattern feature
    swCircularPatternFeatureData.TotalInstances = LocalCirNumInstances
    'After updating Feature you must use ModifyDefinition, so changes will take place!
    swFeature.ModifyDefinition swCircularPatternFeatureData, swModel, Nothing
    
    swModel.EditRebuild3
    swModel.Save3 swSaveAsOptions_Silent, lErrors, lWarnings
    
    'Mandrel
    swApp.ActivateDoc "Mandrel, Fixture, Stacking, Rotor, Main"
    Set swModel = swApp.ActiveDoc
    swModel.Parameter("MandrelOD@Sketch1").SystemValue = MandrelOD * inTOmeter
    swModel.Parameter("MandrelODatBase@Sketch1").SystemValue = MandrelODatBase * inTOmeter
    swModel.Parameter("MandrelHeight@Sketch1").SystemValue = MandrelHeight * inTOmeter
    
    'This is used to suppress a feature
    Set swPart = swModel
    
    If CoreIDAfterGrind > MaxCoreIDnoMandrelID Then
        swModel.Parameter("MandrelID@Sketch4").SystemValue = MandrelID * inTOmeter
        swModel.Parameter("MandrelScrewLocation@Sketch5").SystemValue = MandrelScrewLocation * inTOmeter
    
        Set swFeature = swPart.FeatureByName("3/8-16 Tapped Hole1")
        swFeature.SetSuppression2 swSuppressFeature, swAllConfiguration, Nothing
        Set swFeature = swPart.FeatureByName("Cut-Extrude1")
        swFeature.SetSuppression2 swUnSuppressFeature, swAllConfiguration, Nothing
        Set swFeature = swPart.FeatureByName("1/4-20 Tapped Hole1")
        swFeature.SetSuppression2 swUnSuppressFeature, swAllConfiguration, Nothing
        
    Else
        Set swFeature = swPart.FeatureByName("3/8-16 Tapped Hole1")
        swFeature.SetSuppression2 swUnSuppressFeature, swAllConfiguration, Nothing
        Set swFeature = swPart.FeatureByName("Cut-Extrude1")
        swFeature.SetSuppression2 swSuppressFeature, swAllConfiguration, Nothing
        Set swFeature = swPart.FeatureByName("1/4-20 Tapped Hole1")
        swFeature.SetSuppression2 swSuppressFeature, swAllConfiguration, Nothing
    End If
    
    swModel.EditRebuild3
    swModel.Save3 swSaveAsOptions_Silent, lErrors, lWarnings
    
    'Base
    swApp.ActivateDoc "Lower, Base, Fixture, Stacking, Rotor, Main"
    Set swModel = swApp.ActiveDoc
    swModel.Parameter("BaseScrewLoactionD@Sketch5").SystemValue = BaseScrewLoactionD * inTOmeter
    swModel.Parameter("BaseOD@Sketch1").SystemValue = BaseOD * inTOmeter
    
    'This is used to suppress a feature
    Set swPart = swModel
    Set swModelDocExt = swModel.Extension
    
    Dim suppressionsuccessful As Boolean
    
    If CoreIDAfterGrind > MaxCoreIDnoMandrelID Then
        swModel.Parameter("BaseScrewLocation@Sketch9").SystemValue = BaseScrewLocation * inTOmeter
        
        Set swFeature = swPart.FeatureByName("CBORE for 1/4 Socket Head Cap Screw2")
        swFeature.SetSuppression2 swUnSuppressFeature, swAllConfiguration, Nothing
        Set swFeature = swPart.FeatureByName("CBORE for 3/8 Socket Head Cap Screw1")
        swFeature.SetSuppression2 swSuppressFeature, swAllConfiguration, Nothing
    Else
        Set swFeature = swPart.FeatureByName("CBORE for 1/4 Socket Head Cap Screw2")
        swFeature.SetSuppression2 swSuppressFeature, swAllConfiguration, Nothing
        Set swFeature = swPart.FeatureByName("CBORE for 3/8 Socket Head Cap Screw1")
        swFeature.SetSuppression2 swUnSuppressFeature, swAllConfiguration, Nothing
    End If
    
    swModel.EditRebuild3
    swModel.Save3 swSaveAsOptions_Silent, lErrors, lWarnings

    'Assembly
    swApp.ActivateDoc "Assembly, Stacking, Rotor, Main"
    Set swModel = swApp.ActiveDoc
    Set swAssy = swModel

    'Unsuppress only the relevant core
    'First suppress all the cores:
    For i = 0 To UBound(CoreNamesArray)
        Set swComp = swAssy.GetComponentByName(CoreNamesArray(i))
        swComp.SetSuppression2 swComponentSuppressed
    Next i
    'Next, unsuppress only the relevant core
    Set swComp = swAssy.GetComponentByName(CoreName)
    swComp.SetSuppression2 swComponentFullyResolved

    'This part is used to control the number of instances in the circular pattern for ASSEMBLY
    Set swFeature = swAssy.FeatureByName("LocalCirPattern1")
    If swFeature Is Nothing Then Debug.Print "swFeature is nothing"
    Set swLocCircPatt = swFeature.GetDefinition 'Might be beacuse I forgot the SET???? Or didn't work beacuse of SelectByID2???
    swLocCircPatt.TotalInstances = LocalCirNumInstances
    swFeature.ModifyDefinition swLocCircPatt, swModel, Nothing
    
    Dim successSuppress As Boolean
    
    If CoreIDAfterGrind > MaxCoreIDnoMandrelID Then
        Set swComp = swAssy.GetComponentByName("Socket Screw-1")
        swComp.SetSuppression2 swComponentSuppressed
        Set swComp = swAssy.GetComponentByName("Socket Screw-17")
        swComp.SetSuppression2 swComponentFullyResolved
        Set swFeature = swAssy.FeatureByName("LocalCirPattern2")
        swFeature.SetSuppression2 swUnSuppressFeature, swAllConfiguration, Nothing
    Else
        Set swComp = swAssy.GetComponentByName("Socket Screw-1")
        swComp.SetSuppression2 swComponentFullyResolved
        Set swComp = swAssy.GetComponentByName("Socket Screw-17")
        swComp.SetSuppression2 swComponentSuppressed
        Set swFeature = swAssy.FeatureByName("LocalCirPattern2")
        swFeature.SetSuppression2 swSuppressFeature, swAllConfiguration, Nothing
    End If

    swModel.Extension.Rebuild swForceRebuildAll
    'EditRebuild3
    swModel.Save3 swSaveAsOptions_Silent, lErrors, lWarnings

End Sub




