Attribute VB_Name = "HairPinModule"
Option Explicit

Const PI = 3.14159265358979

Sub HairPinOld()

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
    
    AssemblyArray = Array("1021-21-02114 Coil, Hair Pin, Stator, Main-1", _
                            "1034-21-07124 Coil, Hair Pin, 72-Pole, Stator, Main-1")

    
    CopyCodeFile
    
    Set swApp = Application.SldWorks
    
    inTOmeter = 0.0254 'Units are in meters
    meterToin = 1 / inTOmeter 'To convert back to inch for checking results
    degTORad = PI / 180 'From degrees to Radians
    radToDeg = 180 / PI
    
    'Part Properties
    Dim AssemblyName As String
    Dim CrossSectionA As Double 'Perpendicular to circle. Maximum with insulation
    Dim CrossSectionB As Double 'Tangent to circle. Maximum with insulation
    Dim InnerLegInnerR As Double 'As from hairpin drawing
    Dim OuterLegInnerR As Double 'As from hairpin drawing
    Dim Angle As Double 'As from hairpin drawing - > could be taken from core, should be on drawing
    Dim SkewAngle As Double 'As from hairpin drawing
    Dim Height As Double
    'Those will be calculated
    Dim HeightDeltaDueToSkew As Double
    Dim HeadHeight As Double 'Should be calculated too - not sure still how
    
    '***** Tool Dimensions *****
    
    'Static Part
    Dim StaticPartIR As Double
    Dim StaticPartSlotA As Double
    Dim StaticPartSlotB As Double
    Dim StaticPartSlotSkewA As Double
    Dim StaticPartSlotSkewB As Double
    Dim StaticPartHeight As Double
    
    'Rotating Part
    Dim RotatingPartOR As Double
    Dim RotatingPartSlotA As Double
    Dim RotatingPartSlotB As Double
    Dim RotatingPartSkewSlotA As Double
    Dim RotatingPartSkewSlotB As Double
    Dim RotatingPartHeight As Double
    
    'Adapter Plate
    Dim AdapterPlateIR As Double
    
    '***********************************************************************************************************
    '***********************************************************************************************************
    ToolAssemblyFolder = "C:\Users\FIlie\Documents\Felix Documents IPS\Master Tooling\Main Stator Hair Pin\"
    ToolAssemblyPath = ToolAssemblyFolder + "Assem Hair Pin Forming Old.sldasm"
    '***********************************************************************************************************
    UnitType = "Agusta 609 DC" ' "Agusta 609 DC", "Agusta 169', "SAAB", "Textron", "Scorpion"
    '***********************************************************************************************************
    '***********************************************************************************************************
    
    Select Case UnitType
    
        Case "SAAB"
            
            AssemblyName = "1021-21-02114 Coil, Hair Pin, Stator, Main-1"

            CrossSectionA = 0.08
            CrossSectionB = 0.165
            InnerLegInnerR = 2.364
            OuterLegInnerR = 2.544
            Angle = 75
            SkewAngle = 7.5
            Height = 2.5 'min 5

        Case "Textron"
            
'            AssemblyName = "1045-21-05374 Coil, Hair Pin, Stator, Main"

            CrossSectionA = 0.068
            CrossSectionB = 0.115
            InnerLegInnerR = 3.873
            OuterLegInnerR = 4.009
            Angle = 45
            SkewAngle = 3.75
            Height = 8 'min

        Case "Agusta 169"

'            AssemblyName = "1015-21-01118 Coil, Hair Pin, Stator, Main"

            CrossSectionA = 0.129
            CrossSectionB = 0.07
            InnerLegInnerR = 2.123
            OuterLegInnerR = 2.207
            Angle = 45
            SkewAngle = 5.15
            Height = 6 'min

        Case "Agusta 609 DC"

            AssemblyName = "1034-21-07124 Coil, Hair Pin, 72-Pole, Stator, Main-1"

            CrossSectionA = 0.137
            CrossSectionB = 0.093
            InnerLegInnerR = 2.933
            OuterLegInnerR = 3.045
            Angle = 25
            SkewAngle = 360 - 6.75 'Inverse skew direction
            Height = 2.2 '4.13 'min
            
        Case "Scorpion"

'            AssemblyName = "0905-21-01368 Coil, Hair Pin, Stator, Main"

            CrossSectionA = 0.137
            CrossSectionB = 0.093
            InnerLegInnerR = 2.536
            OuterLegInnerR = 2.643
            Angle = 37.5
            SkewAngle = 7.5
            Height = 5.25 'min

        Case Else
            swApp.SendMsgToUser2 "Data for this unit is not available", swMbStop, swMbOk
            errors = False
            Exit Sub
    End Select

    HeightDeltaDueToSkew = Round(2 * InnerLegInnerR * Sin(degTORad * Angle / 2) * Sin(degTORad * SkewAngle), 2)
    
    Debug.Print "Height Delta due to skew: " & HeightDeltaDueToSkew
    
    Debug.Print "HairPin Data: " & CrossSectionA, CrossSectionB, InnerLegInnerR, OuterLegInnerR
    '***** Calculating Tool Dimensions ***** According to Chaim old hair pin fixture
    
    'Static Part
    
    StaticPartIR = OuterLegInnerR - 0.002 'StaticPartIR@Sketch2
    'SkewAngle@Sketch6
    StaticPartSlotSkewA = CrossSectionA + 0.003 'StaticPartSlotSkewA@Sketch6
    StaticPartSlotSkewB = CrossSectionB + 0.003 'StaticPartSlotSkewB@Cut-Extrude1 ' ********+.01***********
    StaticPartSlotA = CrossSectionA + 0.003 'StaticPartSlotA@Sketch8
    StaticPartSlotB = CrossSectionB + 0.003 'StaticPartSlotB@Sketch8
    StaticPartHeight = Height 'StaticPartHeight@Sketch2
    
    Debug.Print "Static Part: " & StaticPartIR, StaticPartSlotSkewA, StaticPartSlotSkewB, StaticPartSlotA, StaticPartSlotB
    
    'Rotating Part
    
    RotatingPartOR = StaticPartIR - 0.007 'RotatingPartOR@Sketch1 ' ********- 0.015***********
    RotatingPartSlotA = CrossSectionA + 0.003 'RotatingPartSlotA@Sketch9
    RotatingPartSlotB = CrossSectionB + 0.003 'RotatingPartSlotB@Sketch9
    'SkewAngle@Sketch7
    RotatingPartSkewSlotA = CrossSectionA + 0.003 'RotatingPartSkewSlotA@Sketch7
    RotatingPartSkewSlotB = CrossSectionB + 0.003 'RotatingPartSkewSlotB@Cut-Extrude1
    RotatingPartHeight = Height 'RotatingPartHeight@Boss-Extrude1
    
    Debug.Print "Rotating Part: " & RotatingPartOR, RotatingPartSkewSlotA, RotatingPartSkewSlotB, RotatingPartSlotA, RotatingPartSlotB
    
    'Adapter Plate
    
    AdapterPlateIR = StaticPartIR + 0.22 + 0.005 'AdapterPlateIR@Sketch1
    
    Debug.Print "Adapter Plate: " & AdapterPlateIR
     
    '***** Changing Tool Dimensions *****
    ' DONT FORGET TO CONVERT TO METERS!!!!
    ' DEG in RADIANS!!!!

    Set swModel = swApp.OpenDoc6(ToolAssemblyPath, _
    swDocASSEMBLY, swOpenDocOptions_OverrideDefaultLoadLightweight, "", lErrors, lWarnings)

    'Static Part
    swApp.ActivateDoc "Static Part"
    Set swModel = swApp.ActiveDoc
    swModel.Parameter("StaticPartIR@Sketch2").SystemValue = StaticPartIR * inTOmeter
    swModel.Parameter("SkewAngle@Sketch6").SystemValue = SkewAngle * degTORad
    swModel.Parameter("StaticPartSlotSkewA@Sketch6").SystemValue = StaticPartSlotSkewA * inTOmeter
    swModel.Parameter("StaticPartSlotSkewB@Cut-Extrude1").SystemValue = StaticPartSlotSkewB * inTOmeter
    swModel.Parameter("StaticPartSlotA@Sketch8").SystemValue = StaticPartSlotA * inTOmeter
    swModel.Parameter("StaticPartSlotB@Sketch8").SystemValue = StaticPartSlotB * inTOmeter
    swModel.Parameter("StaticPartHeight@Sketch2").SystemValue = StaticPartHeight * inTOmeter
    swModel.EditRebuild3
    swModel.Save3 swSaveAsOptions_Silent, lErrors, lWarnings

    'Rotating Part
    swApp.ActivateDoc "Rotating Part"
    Set swModel = swApp.ActiveDoc
    swModel.Parameter("RotatingPartOR@Sketch1").SystemValue = RotatingPartOR * inTOmeter
    swModel.Parameter("SkewAngle@Sketch7").SystemValue = SkewAngle * degTORad
    swModel.Parameter("RotatingPartSlotA@Sketch9@Sketch6").SystemValue = RotatingPartSlotA * inTOmeter
    swModel.Parameter("RotatingPartSlotB@Sketch9@Cut-Extrude1").SystemValue = RotatingPartSlotB * inTOmeter
    swModel.Parameter("RotatingPartSkewSlotA@Sketch7").SystemValue = RotatingPartSkewSlotA * inTOmeter
    swModel.Parameter("RotatingPartSkewSlotB@Cut-Extrude1").SystemValue = RotatingPartSkewSlotB * inTOmeter
    swModel.Parameter("RotatingPartHeight@Boss-Extrude1@Sketch2").SystemValue = RotatingPartHeight * inTOmeter
    swModel.EditRebuild3
    swModel.Save3 swSaveAsOptions_Silent, lErrors, lWarnings

    'Adapter Plate
    swApp.ActivateDoc "Adapter Plate"
    Set swModel = swApp.ActiveDoc
    swModel.Parameter("AdapterPlateIR@Sketch1").SystemValue = AdapterPlateIR * inTOmeter
    swModel.EditRebuild3
    swModel.Save3 swSaveAsOptions_Silent, lErrors, lWarnings

    'Assembly
    swApp.ActivateDoc "Assem Hair Pin Forming Old"
    Set swModel = swApp.ActiveDoc
    Set swAssy = swModel

    'Unsuppress only the relevant Assemblies
    'First suppress all the cores:
    For i = 0 To UBound(AssemblyArray)
        Set swComp = swAssy.GetComponentByName(AssemblyArray(i))
        swComp.SetSuppression2 swComponentSuppressed
    Next i
    'Next, unsuppress only the relevant Assembly
    Set swComp = swAssy.GetComponentByName(AssemblyName)
    swComp.SetSuppression2 swComponentFullyResolved

    swModel.Extension.Rebuild swForceRebuildAll
    'EditRebuild3
    swModel.Save3 swSaveAsOptions_Silent, lErrors, lWarnings

End Sub


