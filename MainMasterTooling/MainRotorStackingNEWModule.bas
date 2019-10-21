Attribute VB_Name = "MainRotorStackingNEWModule"
Option Explicit

Const PI = 3.14159265358979

Sub MainRotorStackingNEW()

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
    
    'MainSketch
    Dim MandrelOD As Double 'MandrelOD@MainSketch
    Dim ToolOD As Double 'ToolOD@MainSketch
    Dim PinDistance As Double 'PinDistance@MainSketch
    Dim PinDia As Double 'PinDia@MainSketch
    Dim PinPatternIns As Integer 'PinPatternIns@PinPattern
    Dim PINOD As Double 'PINOD@MainSketch
    
    Dim MaxCoreIDnoMandrelID As Double
    Dim ToolScrewAngle As Double 'ToolScrewAngle@Sketch15 WASn't set yet
    
    'Mandrel
    Dim MandrelHeight As Double 'MandrelHeight@Boss-Extrude1
    
    '***********************************************************************************************************
    '***********************************************************************************************************
    ToolAssemblyFolder = "C:\Users\FIlie\Documents\Felix Documents IPS\Master Tooling\Main Rotor Stacking NEW\"
    ToolAssemblyPath = ToolAssemblyFolder + "Assem.SLDASM"
    '***********************************************************************************************************
    UnitType = "A4" ' "Agusta 169", "Agusta 609 DC" ' "Agusta 609 DC"
    '***********************************************************************************************************
    '***********************************************************************************************************
    
'    InverseSkewDirection = False
    
    Select Case UnitType
    
        Case "Latitude"
        
            CoreName = "1015-11-04313 Assembly, Core, Rotor, Main-1"
        
            NumberOfPoles = 4
            LamMinID = 0.977
            LamThickness = 0.014
            LamCopperRodsLoactionD = 3.236
            LamCopperRodsD = 0.078
            LamPoleMaxWidth = 0.892
            LamPoleLocationD = 2.6
            CoreHeight = 1.95 'Average
            CoreIDAfterGrind = 0.95 'Not important here - Fixture isn't used for grind
            
        Case "Agusta 169"
        
            CoreName = "1015-11-01055 Assembly, Core, Rotor, Main-1"
        
            NumberOfPoles = 8
            LamMinID = 1.85
            LamThickness = 0.014
            LamCopperRodsLoactionD = 4.002
            LamCopperRodsD = 0.068
            LamPoleMaxWidth = 0.52
            LamPoleLocationD = 3.175
            CoreHeight = 3.084 'Average
            CoreIDAfterGrind = 0.95 'Not important here - Fixture isn't used for grind

        Case "Agusta 609 AC"
        
            CoreName = "1034-11-05942 Assembly, Core, Rotor, Main-1"
        
            NumberOfPoles = 4
            LamMinID = 0.938
            LamThickness = 0.014
            LamCopperRodsLoactionD = 3.584
            LamCopperRodsD = 0.062
            LamPoleMaxWidth = 1.077
            LamPoleLocationD = 2.4
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
            
        Case "A4"
        
            'CoreName = "1032-11-03227 Assembly, Core, Rotor, Main-1"

            NumberOfPoles = 12
            LamMinID = 1.458
            LamThickness = 0.014
            LamCopperRodsLoactionD = 7.058
            LamCopperRodsD = 0.106
            LamPoleMaxWidth = 0.652
            LamPoleLocationD = 6
            CoreHeight = 2.85 'Average
            CoreIDAfterGrind = 1.477

        Case Else
            swApp.SendMsgToUser2 "Data for this unit is not available", swMbStop, swMbOk
            errors = False
            Exit Sub
    End Select
    
    '***** Calculating Tool Dimensions *****
    
    'General Tool
    ToolOD = LamCopperRodsLoactionD - 2 * LamCopperRodsD - 0.01
    PinDistance = LamPoleMaxWidth + 0.002
    PINOD = 0.25 - 0.0005
    MaxCoreIDnoMandrelID = 2
    PinDia = LamPoleLocationD
    PinPatternIns = 4 'NumberOfPoles
    
    'Mandrel
    MandrelOD = LamMinID - 0.001
    MandrelHeight = 0.825 + 1.6 + CoreHeight - 0.1 'TopHeight + Upper Base Height + CoreHeight - .1
    
    Debug.Print ToolOD, PinDistance, PINOD, MaxCoreIDnoMandrelID, MandrelHeight
    
    '***** Changing Tool Dimensions *****
    ' DONT FORGET TO CONVERT TO METERS!!!!
    ' DEG in RADIANS!!!!
    
    Set swModel = swApp.OpenDoc6(ToolAssemblyPath, _
    swDocASSEMBLY, swOpenDocOptions_Silent, "", lErrors, lWarnings)

    'MainSketch
    swApp.ActivateDoc "MainSketch"
    Set swModel = swApp.ActiveDoc
    swModel.Parameter("MandrelOD@MainSketch").SystemValue = MandrelOD * inTOmeter
    swModel.Parameter("ToolOD@MainSketch").SystemValue = ToolOD * inTOmeter
    swModel.Parameter("PinDistance@MainSketch").SystemValue = PinDistance * inTOmeter
    swModel.Parameter("PinDia@MainSketch").SystemValue = PinDia * inTOmeter
    swModel.Parameter("PINOD@MainSketch").SystemValue = PINOD * inTOmeter
    swModel.Parameter("PinPatternIns@PinPattern").SystemValue = PinPatternIns
    swModel.Parameter("PinClearPatternIns@PinClearPattern").SystemValue = PinPatternIns
     
    swModel.EditRebuild3
    swModel.Save3 swSaveAsOptions_Silent, lErrors, lWarnings
    
    'Mandrel
    swApp.ActivateDoc "Mandrel"
    Set swModel = swApp.ActiveDoc
    swModel.Parameter("MandrelHeight@Boss-Extrude1").SystemValue = MandrelHeight * inTOmeter
    
    swModel.EditRebuild3
    swModel.Save3 swSaveAsOptions_Silent, lErrors, lWarnings

    'Assembly
    swApp.ActivateDoc "Assem"
    Set swModel = swApp.ActiveDoc
    Set swAssy = swModel

    'Unsuppress only the relevant core
    'First suppress all the cores:
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
'    swLocCircPatt.TotalInstances = PinPatternIns
'    swFeature.ModifyDefinition swLocCircPatt, swModel, Nothing

    swModel.Extension.Rebuild swForceRebuildAll
    'EditRebuild3
    swModel.Save3 swSaveAsOptions_Silent, lErrors, lWarnings

End Sub






