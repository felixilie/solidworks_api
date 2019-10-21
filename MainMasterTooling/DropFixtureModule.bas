Attribute VB_Name = "DropFixtureModule"
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
    Dim CoreHeight As Double 'with copper plates
    Dim CoreOD As Double
    Dim CoreID As Double
    Dim ShaftSmallOD As Double 'Where ID of Bullet is going to be located to
    Dim ShaftBigOD As Double 'Where part is going to sit
  
    '***** Tool Dimensions *****
    
    'Locator
    Dim LocatorBigID As Double 'LocatorBigID@Sketch1
    Dim LocatorHeight As Double 'LocatorHeight@Sketch1
    Dim LocatorSmallID As Double 'LocatorSmallID@Sketch1
    Dim LocatorSlot As Double 'LocatorSlot@Sketch1
    
    'Bullet
    Dim BulletLength As Double 'BulletLength@Sketch1
    Dim BulletID As Double 'BulletID@Sketch1
    Dim BulletOD As Double 'BulletOD@Sketch1
    
    '***********************************************************************************************************
    '***********************************************************************************************************
    ToolAssemblyFolder = "C:\Users\FIlie\Documents\Felix Documents IPS\Master Tooling\Drop Fixture\"
    ToolAssemblyPath = ToolAssemblyFolder + "Assembly.sldasm"
    '***********************************************************************************************************
    UnitType = "Agusta 609 DC to core" ' "Agusta 609 AC","Agusta 609 DC", "CH47', "SAAB", "Textron", "Scorpion"
    '***********************************************************************************************************
    '***********************************************************************************************************
    
    Select Case UnitType
    
'        Case "SAAB"
'
'
        Case "Agusta 609 DC shaft to hub"
        
            LengthToShoulder = 0.9
            CoreHeight = 2.15
            CoreOD = 3.82
            CoreID = 0.9975 'After Grinding
            ShaftSmallOD = 0.788
'            ShaftBigOD = 1
            
        Case "Agusta 609 DC to core"
        
            LengthToShoulder = 0.3
            CoreHeight = 2.15
            CoreOD = 5.753
            CoreID = 3.816 'After Grinding
            ShaftSmallOD = 1
'            ShaftBigOD = 3.8195
            
        Case "Agusta 609 AC"

            LengthToShoulder = 0.901
            CoreHeight = 3.05
            CoreOD = 3.744
            CoreID = 0.95 'After Grinding
            ShaftSmallOD = 0.788
'            ShaftBigOD = 3.8

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
    LocatorHeight = CoreHeight / 2
    LocatorSmallID = BulletOD + 0.1
    If UnitType = "Agusta 609 DC to core" Then LocatorSmallID = 1.5
    LocatorSlot = 0.3
    
    Debug.Print LocatorBigID, LocatorHeight, LocatorSmallID, LocatorSlot

    '***** Changing Tool Dimensions *****
    ' DONT FORGET TO CONVERT TO METERS!!!!
    ' DEG in RADIANS!!!!

    Set swModel = swApp.OpenDoc6(ToolAssemblyPath, _
    swDocASSEMBLY, swOpenDocOptions_OverrideDefaultLoadLightweight, "", lErrors, lWarnings)

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
    swModel.Parameter("LocatorSlot@Sketch1").SystemValue = LocatorSlot * inTOmeter
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


