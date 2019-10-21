Attribute VB_Name = "ExciterRotorDropFixtureModule"
Option Explicit

Const PI = 3.14159265358979

Sub DropFixture()

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
    Dim LengthToShoulder As Double 'To step where Part is pressed to
    Dim CoreHeight As Double 'Only Core
    Dim CoreOD As Double
    Dim CoreID As Double
    Dim CoreInnerOD As Double
    Dim ShaftSmallOD As Double 'Where ID of Bullet is going to be located to
    Dim CoreToBottomDis As Double 'From bottom of core to bottom of exciter rotor
    'Dim ShaftBigOD As Double 'Where part is going to sit
  
    '***** Tool Dimensions *****
    
    'Locator
    Dim LocatorBigID As Double 'LocatorBigID@Sketch1
    Dim LocatorHeight As Double 'LocatorHeight@Sketch1
    Dim LocatorSmallID As Double 'LocatorSmallID@Sketch1
    Dim LocatorDepth As Double 'LocatorDepth@Sketch1
    Dim LocatorSmallOD As Double 'LocatorSmallOD@Sketch1
    
    'Bullet
    Dim BulletLength As Double 'BulletLength@Sketch1
    Dim BulletID As Double 'BulletID@Sketch1
    Dim BulletOD As Double 'BulletOD@Sketch1
    
    '***********************************************************************************************************
    '***********************************************************************************************************
    ToolAssemblyFolder = "C:\Users\FIlie\Documents\Felix Documents IPS\Master Tooling\Exciter Rotor Drop Fixture\"
    ToolAssemblyPath = ToolAssemblyFolder + "Assembly.sldasm"
    '***********************************************************************************************************
    UnitType = "Rolls Royce" ' "Agusta 609 AC","Agusta 609 DC", "CH47', "SAAB", "Textron", "Scorpion"
    '***********************************************************************************************************
    '***********************************************************************************************************
    
    Select Case UnitType
    
'        Case "SAAB"
'
'
'        Case "Agusta 609 DC"

            
        Case "Agusta 609 AC"

            LengthToShoulder = 2.6
            CoreHeight = 0.475
            CoreOD = 4.15
            CoreID = 1.018 'After Grinding
            CoreInnerOD = 1.27
            ShaftSmallOD = 0.788
            CoreToBottomDis = 0.562
            
        Case "Rolls Royce"

            LengthToShoulder = 2.8
            CoreHeight = 0.475
            CoreOD = 4.15
            CoreID = 1.0925 'After Grinding
            CoreInnerOD = 1.27
            ShaftSmallOD = 0.9846
            CoreToBottomDis = 0.562

        Case Else
            swApp.SendMsgToUser2 "Data for this unit is not available", swMbStop, swMbOk
            errors = False
            Exit Sub
    End Select

    '***** Calculating Tool Dimensions *****
    
    'Bullet
    BulletLength = LengthToShoulder + 0.55
    BulletID = ShaftSmallOD + 0.002
    BulletOD = CoreID - 0.004
     
    Debug.Print BulletLength, BulletID, BulletOD
    
    'Locator
    LocatorBigID = CoreOD + 0.015
    LocatorHeight = CoreToBottomDis + 0.2 + CoreHeight
    LocatorSmallID = BulletOD + 0.05
    LocatorDepth = CoreToBottomDis + 0.1
    LocatorSmallOD = CoreInnerOD + 0.1
    
    Debug.Print LocatorBigID, LocatorHeight, LocatorSmallID, LocatorDepth, LocatorSmallOD

    '***** Changing Tool Dimensions *****
    ' DONT FORGET TO CONVERT TO METERS!!!!
    ' DEG in RADIANS!!!!

    Set swModel = swApp.OpenDoc6(ToolAssemblyPath, _
    swDocASSEMBLY, swOpenDocOptions_Silent, "", lErrors, lWarnings)

    'Bullet
    swApp.ActivateDoc "Bullet"
    Set swModel = swApp.ActiveDoc
    swModel.Parameter("BulletLength@Sketch1").SystemValue = BulletLength * inTOmeter
    swModel.Parameter("BulletID@Sketch1").SystemValue = BulletID * inTOmeter
    swModel.Parameter("BulletOD@Sketch1").SystemValue = BulletOD * inTOmeter
    swModel.EditRebuild3
    swModel.Save3 swSaveAsOptions_Silent, lErrors, lWarnings

    'Locator
    swApp.ActivateDoc "Locator, Core, Rotor, Main"
    Set swModel = swApp.ActiveDoc
    swModel.Parameter("LocatorBigID@Sketch1").SystemValue = LocatorBigID * inTOmeter
    swModel.Parameter("LocatorHeight@Sketch1").SystemValue = LocatorHeight * inTOmeter
    swModel.Parameter("LocatorSmallID@Sketch1").SystemValue = LocatorSmallID * inTOmeter
    swModel.Parameter("LocatorDepth@Sketch1").SystemValue = LocatorDepth * inTOmeter
    swModel.Parameter("LocatorSmallOD@Sketch1").SystemValue = LocatorSmallOD * inTOmeter

    swModel.EditRebuild3
    swModel.Save3 swSaveAsOptions_Silent, lErrors, lWarnings

    'Assembly
    swApp.ActivateDoc "Assembly.sldasm"
    Set swModel = swApp.ActiveDoc
    Set swAssy = swModel

'    'Unsuppress only the relevant Assemblies
'    'First suppress all the cores:
'    For i = 0 To UBound(AssemblyArray)
'        Set swComp = swAssy.GetComponentByName(AssemblyArray(i))
'        swComp.SetSuppression2 swComponentSuppressed
'    Next i
'    'Next, unsuppress only the relevant Assembly
'    Set swComp = swAssy.GetComponentByName(AssemblyName)
'    swComp.SetSuppression2 swComponentFullyResolved

    swModel.Extension.Rebuild swForceRebuildAll
    'EditRebuild3
    swModel.Save3 swSaveAsOptions_Silent, lErrors, lWarnings

End Sub




