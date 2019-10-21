Attribute VB_Name = "HairPinNewestModule"
Option Explicit

Const PI = 3.14159265358979

Sub HairPinNewest()

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
    
    Dim HairPinNameArray() As Variant
    
'    HairPinNameArray = Array("1034-21-07119 Lamination, Stator, Main-1", _
'                            "1034-21-05972 Lamination, Stator, Main-1", _
'                            "1015-21-04311 Lamination, Stator, Main-1", _
'                            "1015-21-01116 Assembly, Core, Stator, Main-1", _
'                            "1021-21-02110 Lamination, Stator, Main-1")
    
    CopyCodeFile
    
    Set swApp = Application.SldWorks
    
    inTOmeter = 0.0254 'Units are in meters
    meterToin = 1 / inTOmeter 'To convert back to inch for checking results
    degTORad = PI / 180 'From degrees to Radians
    radToDeg = 180 / PI
    
    'Part Properties
    'Dim HairPinName As String
    Dim CrossSectionA As Double 'Tangent to circle. Maximum with insulation
    Dim CrossSectionB As Double 'Radial to circle. Maximum with insulation
    Dim InnerLegInnerR As Double 'As from hairpin drawing
    Dim OuterLegInnerR As Double 'As from hairpin drawing
    Dim Angle As Double 'As from hairpin drawing - > could be taken from core, should be on drawing
    Dim SkewAngle As Double 'As from hairpin drawing
    'Those will be calculated
    Dim HeightDeltaDueToSkew As Double
    Dim HeadHeight As Double 'Should be calculated too - not sure still how
    
    '***** Tool Dimensions *****
    
    'General
    Dim SlotA As Double 'Tangent to circle. Maximum with insulation
    Dim SlotB As Double 'Radial to circle. Maximum with insulation
    
    '***********************************************************************************************************
    '***********************************************************************************************************
    ToolAssemblyFolder = "C:\Users\FIlie\Documents\Felix Documents IPS\Master Tooling\Main Stator Hair Pin NEW\1 New\"
    ToolAssemblyPath = ToolAssemblyFolder + "Assembly.SLDASM"
    '***********************************************************************************************************
    UnitType = "Agusta 169" ' "Agusta 609 DC", "Agusta 169', "SAAB", "Textron", "Scorpion"
    '***********************************************************************************************************
    '***********************************************************************************************************
    
    Select Case UnitType
    
        Case "SAAB"
            
'            HairPinName = ""

            CrossSectionA = 0.08
            CrossSectionB = 0.165
            InnerLegInnerR = 2.364
            OuterLegInnerR = 2.544
            Angle = 75
            SkewAngle = 7.5

        Case "Textron"
            
'            HairPinName = ""

            CrossSectionA = 0.068
            CrossSectionB = 0.115
            InnerLegInnerR = 3.873
            OuterLegInnerR = 4.009
            Angle = 45
            SkewAngle = 3.75

        Case "Agusta 169"

'            HairPinName = ""

            CrossSectionA = 0.129
            CrossSectionB = 0.07
            InnerLegInnerR = 2.123
            OuterLegInnerR = 2.207
            Angle = 45
            SkewAngle = 5.15

        Case "Agusta 609 DC"

'            HairPinName = ""

            CrossSectionA = 0.137
            CrossSectionB = 0.093
            InnerLegInnerR = 2.933
            OuterLegInnerR = 3.045
            Angle = 25
            SkewAngle = 6.75
            
        Case "Scorpion"

'            HairPinName = ""

            CrossSectionA = 0.137
            CrossSectionB = 0.093
            InnerLegInnerR = 2.536
            OuterLegInnerR = 2.643
            Angle = 37.5
            SkewAngle = 7.5

        Case Else
            swApp.SendMsgToUser2 "Data for this unit is not available", swMbStop, swMbOk
            errors = False
            Exit Sub
    End Select

    HeightDeltaDueToSkew = Round(2 * InnerLegInnerR * Sin(degTORad * Angle / 2) * Sin(degTORad * SkewAngle), 2)
    
    Debug.Print HeightDeltaDueToSkew
    
    Debug.Print CrossSectionA, CrossSectionB, InnerLegInnerR, OuterLegInnerR
    '***** Calculating Tool Dimensions ***** According to Chaim old hair pin fixture
    
    'Slot
    SlotA = CrossSectionA + 0.002
    SlotB = CrossSectionB + 0.002
     
    '***** Changing Tool Dimensions *****
    ' DONT FORGET TO CONVERT TO METERS!!!!
    ' DEG in RADIANS!!!!

    Set swModel = swApp.OpenDoc6(ToolAssemblyPath, _
    swDocASSEMBLY, swOpenDocOptions_OverrideDefaultLoadLightweight, "", lErrors, lWarnings)
    
    'MainSketch
    swApp.ActivateDoc "MainSketch"
    Set swModel = swApp.ActiveDoc
    swModel.Parameter("OuterLegInnerR@MainSketch").SystemValue = OuterLegInnerR * inTOmeter
    swModel.Parameter("InnerLegInnerR@MainSketch").SystemValue = InnerLegInnerR * inTOmeter
    swModel.Parameter("Angle@MainSketch").SystemValue = Angle * degTORad
    swModel.Parameter("SlotA@MainSketch").SystemValue = SlotA * inTOmeter
    swModel.Parameter("SlotB@MainSketch").SystemValue = SlotB * inTOmeter
    swModel.EditRebuild3
    swModel.Save3 swSaveAsOptions_Silent, lErrors, lWarnings
    
    'Fixed Clamp
    swApp.ActivateDoc "Fixed Clamp"
    Set swModel = swApp.ActiveDoc
    Set swPart = swModel
    swModel.Parameter("HeightDeltaDueToSkew@Cut-Extrude2").SystemValue = HeightDeltaDueToSkew * inTOmeter

    Set swFeature = swPart.FeatureByName("Cut-Extrude3")
    If UnitType = "Textron" Or UnitType = "SAAB" Then
        swFeature.SetSuppression2 swUnSuppressFeature, swAllConfiguration, Nothing
    Else
        swFeature.SetSuppression2 swSuppressFeature, swAllConfiguration, Nothing
    End If
    
    swModel.EditRebuild3
    swModel.Save3 swSaveAsOptions_Silent, lErrors, lWarnings
    
    'Moving Clamp
    swApp.ActivateDoc "Moving Clamp"
    Set swModel = swApp.ActiveDoc
    Set swPart = swModel

    Set swFeature = swPart.FeatureByName("Cut-Extrude1")
    If UnitType = "Textron" Or UnitType = "SAAB" Then
        swFeature.SetSuppression2 swUnSuppressFeature, swAllConfiguration, Nothing
    Else
        swFeature.SetSuppression2 swSuppressFeature, swAllConfiguration, Nothing
    End If
    
    swModel.EditRebuild3
    swModel.Save3 swSaveAsOptions_Silent, lErrors, lWarnings
    
    'Assembly
    swApp.ActivateDoc "Assembly"
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

    swModel.Extension.Rebuild swForceRebuildAll
    'EditRebuild3
    swModel.Save3 swSaveAsOptions_Silent, lErrors, lWarnings

End Sub








