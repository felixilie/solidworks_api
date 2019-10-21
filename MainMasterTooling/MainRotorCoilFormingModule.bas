Attribute VB_Name = "MainRotorCoilFormingModule"
Option Explicit

Const PI = 3.14159265358979

Sub MainRotorCoilForming()

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
    
    AssemblyArray = Array("1034-11-05955THRU05958 Winding, Rotor, Main-1", _
                            "1034-11-07136THRU07138 Winding, Rotor, Main-1")

    
    CopyCodeFile
    
    Set swApp = Application.SldWorks
    
    inTOmeter = 0.0254 'Units are in meters
    meterToin = 1 / inTOmeter 'To convert back to inch for checking results
    degTORad = PI / 180 'From degrees to Radians
    radToDeg = 180 / PI
    
    'Part Properties
    Dim AssemblyName As String
    Dim CrossSectionWidth As Double
    Dim CrossSectionLength As Double
    Dim Radius As Double
    Dim Height As Double
    Dim CoilWidth As Double
  
    '***** Tool Dimensions *****
    
    'Mandrel
    Dim MandrelWidth As Double
    Dim MandrelLength As Double
    Dim MandrelRadius As Double
    
    'Adapter
    Dim AdapterWidth As Double
    Dim AdapterLength As Double
    'Dim LowerPressHeight As Double
    Dim AdapterRadius As Double
    
    'PressLower
    Dim LowerPressWidth As Double
    Dim LowerPressLength As Double
    Dim LowerPressHeight As Double
    Dim LowerPressRadius As Double
    
    'PressUpper
    Dim UpperPressWidth As Double
    Dim UpperPressLength As Double
    Dim UpperPressHeight As Double
    Dim UpperPressRadius As Double
    Dim UpperPressCoilWidth As Double
    
    '***********************************************************************************************************
    '***********************************************************************************************************
    ToolAssemblyFolder = "C:\Users\FIlie\Documents\Felix Documents IPS\Master Tooling\Main Rotor Coil Forming\"
    ToolAssemblyPath = ToolAssemblyFolder + "Assembly, Mandrel, Coil Winding, Inductor, AC.sldasm"
    '***********************************************************************************************************
    UnitType = "Bell 525" ' "Agusta 609 AC","Agusta 609 DC", "CH47', "SAAB", "Textron", "Scorpion"
    '***********************************************************************************************************
    '***********************************************************************************************************
    
    Select Case UnitType
    
        Case "SAAB"
            
'            AssemblyName = ""
'
'            CrossSectionWidth = '
'            CrossSectionLength = '
'            Radius =
'            Height =
'            CoilWidth =

        Case "Bell 525"
            
            AssemblyName = "1045-21-05374 Coil, Hair Pin, Stator, Main"

            AssemblyName = ""

            CrossSectionWidth = 0.53 '
            CrossSectionLength = 3.775 '
            Radius = 0.25
            Height = 0.523
            CoilWidth = 0.255

        Case "CH47"

            AssemblyName = ""

            CrossSectionWidth = 0.93 '
            CrossSectionLength = 3.475 '
            Radius = 0.312
            Height = 0.55
            CoilWidth = 0.308 'MAX

        Case "Agusta 609 DC"

            AssemblyName = "1034-11-07136THRU07138 Winding, Rotor, Main-1"

            CrossSectionWidth = 0.53 ' was changed after CNR12378 from .453
            CrossSectionLength = 2.651 '
            Radius = 0.432 / 2 ' .22
            Height = 0.496
            CoilWidth = 0.255 'MAX
            
        Case "Agusta 609 AC"

            AssemblyName = "1034-11-05955THRU05958 Winding, Rotor, Main-1"

            CrossSectionWidth = 1.115 '
            CrossSectionLength = 3.798 '
            Radius = 0.312
            Height = 0.667
            CoilWidth = 0.305

        Case Else
            swApp.SendMsgToUser2 "Data for this unit is not available", swMbStop, swMbOk
            errors = False
            Exit Sub
    End Select
    
    Debug.Print "Coil - " & CrossSectionWidth, CrossSectionLength, Radius, Height

    '***** Calculating Tool Dimensions *****
    
    'Mandrel
    MandrelWidth = CrossSectionWidth - 0.02 'MandrelWidth@Sketch1  'According to Textron +.002 - According to Latitude - 0.02
    MandrelLength = CrossSectionLength - 0.005 'MandrelLength@Sketch1 'According to Textron +.005 - According to Latitude - .005
    MandrelRadius = Radius 'MandrelRadius@Sketch1
    
    Debug.Print "Mandrel - " & MandrelWidth, MandrelLength, MandrelRadius
    
    'PressLower
    LowerPressWidth = CrossSectionWidth - 0.005 'LowerPressWidth@Sketch1 'According to Textron -.002 - According to Latitude - 0.005
    LowerPressLength = CrossSectionLength - 0.01 'LowerPressLength@Sketch1 'According to Textron -.002 - According to Latitude - 0.01
    LowerPressHeight = Height + 0.5 'LowerPressHeight@Boss-Extrude1
    LowerPressRadius = Radius 'LowerPressRadius@Fillet1
    
    Debug.Print LowerPressWidth, LowerPressLength, LowerPressHeight, LowerPressRadius
    
    'PressUpper
    UpperPressWidth = LowerPressWidth + 0.005 'UpperPressWidth@Sketch1
    UpperPressLength = CrossSectionLength + 0.01 'UpperPressLength@Sketch1
    UpperPressHeight = LowerPressHeight - Height + 0.1 'UpperPressHeight@Cut-Extrude1
    UpperPressRadius = Radius 'UpperPressRadius@Fillet1
    UpperPressCoilWidth = CoilWidth - 0.01 'UpperPressCoilWidth@Sketch1
     
    Debug.Print UpperPressWidth, UpperPressLength, UpperPressHeight, UpperPressRadius
    
    'Adapter
    AdapterWidth = MandrelWidth + 0.005 'AdapterWidth@Sketch1
    AdapterLength = CrossSectionLength + 0.005 'AdapterLength@Sketch1
    AdapterRadius = Radius 'AdapterRadius@Sketch1
     
    '***** Changing Tool Dimensions *****
    ' DONT FORGET TO CONVERT TO METERS!!!!
    ' DEG in RADIANS!!!!

    Set swModel = swApp.OpenDoc6(ToolAssemblyPath, _
    swDocASSEMBLY, swOpenDocOptions_Silent, "", lErrors, lWarnings)

    'Mandrel
    swApp.ActivateDoc "Assembly, Mandrel, Coil Winding, Inductor, AC"
    Set swModel = swApp.ActiveDoc
    swModel.Parameter("MandrelWidth@Sketch1").SystemValue = MandrelWidth * inTOmeter
    swModel.Parameter("MandrelLength@Sketch1").SystemValue = MandrelLength * inTOmeter
    swModel.Parameter("MandrelRadius@Sketch1").SystemValue = MandrelRadius * inTOmeter
    swModel.EditRebuild3
    swModel.Save3 swSaveAsOptions_Silent, lErrors, lWarnings

    'PressLower
    swApp.ActivateDoc "Mandrel, Pressing, Lower, Coil, Inductor, AC"
    Set swModel = swApp.ActiveDoc
    swModel.Parameter("LowerPressWidth@Sketch1").SystemValue = LowerPressWidth * inTOmeter
    swModel.Parameter("LowerPressLength@Sketch1").SystemValue = LowerPressLength * inTOmeter
    swModel.Parameter("LowerPressHeight@Boss-Extrude1").SystemValue = LowerPressHeight * inTOmeter
    swModel.Parameter("LowerPressRadius@Fillet1").SystemValue = LowerPressRadius * inTOmeter
    swModel.EditRebuild3
    swModel.Save3 swSaveAsOptions_Silent, lErrors, lWarnings

    'PressUpper
    swApp.ActivateDoc "Mandrel, Pressing, Upper, Coil, Inductor, AC"
    Set swModel = swApp.ActiveDoc
    swModel.Parameter("UpperPressWidth@Sketch1").SystemValue = UpperPressWidth * inTOmeter
    swModel.Parameter("UpperPressLength@Sketch1").SystemValue = UpperPressLength * inTOmeter
    swModel.Parameter("UpperPressHeight@Cut-Extrude1").SystemValue = UpperPressHeight * inTOmeter
    swModel.Parameter("D1@Boss-Extrude1").SystemValue = (UpperPressHeight + 0.2) * inTOmeter
    swModel.Parameter("UpperPressRadius@Fillet1").SystemValue = UpperPressRadius * inTOmeter
    swModel.Parameter("UpperPressCoilWidth@Sketch1").SystemValue = UpperPressCoilWidth * inTOmeter
    swModel.EditRebuild3
    swModel.Save3 swSaveAsOptions_Silent, lErrors, lWarnings
    
    'Adapter
    swApp.ActivateDoc "Adapter, Mandrel, Coil Winding, Inductor, AC"
    Set swModel = swApp.ActiveDoc
    swModel.Parameter("AdapterWidth@Sketch1").SystemValue = AdapterWidth * inTOmeter
    swModel.Parameter("AdapterLength@Sketch1").SystemValue = AdapterLength * inTOmeter
    swModel.Parameter("AdapterRadius@Sketch1").SystemValue = AdapterRadius * inTOmeter
    swModel.EditRebuild3
    swModel.Save3 swSaveAsOptions_Silent, lErrors, lWarnings

    'Assembly
    swApp.ActivateDoc "Assembly, Mandrel, Coil Winding, Inductor, AC.sldasm"
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




