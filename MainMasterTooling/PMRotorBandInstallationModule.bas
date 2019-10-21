Attribute VB_Name = "PMRotorBandInstallationModule"
Option Explicit

Const PI = 3.14159265358979

Sub PMRotorBandInstallation()

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
    
    Dim AssemblyArray() As Variant
    
    AssemblyArray = Array("1034-13-07069 Assembly, Rotor, Pm-1", _
                            "1034-13-05979 Assembly, Rotor, Pm-1", _
                            "1029-13-06090 Assembly, Rotor, Pm-1")
    
    CopyCodeFile
    
    Set swApp = Application.SldWorks
    
    inTOmeter = 0.0254 'Units are in meters
    meterToin = 1 / inTOmeter 'To convert back to inch for checking results
    degTORad = PI / 180 'From degrees to Radians
    radToDeg = 180 / PI
    
    'Thermal expension coefficient
    Dim StainlessSteel303ThermalCoefficient As Double
    Dim TitaniumThermalCoefficient As Double
    
    StainlessSteel303ThermalCoefficient = 17.5 '10^-6*inch/inch/c, for 304 too @ 200c
    TitaniumThermalCoefficient = 9.2 '10^-6*inch/inch/c, @ 200c
    'For more materials data go to website - matweb
    'Part are heated to 210c
    'drop-fixture hole in bottom plate is 1.986" diameter
    
    'Part Properties
    Dim AssemblyName As String
    Dim PMRotorOD As Double
    Dim PMRotorID As Double
    Dim PMRotorThick As Double
    Dim ScrewLocationD As Double ' Without Tabs
    Dim ScrewD As Double
    Dim ScrewProtrudeDepth As Double
    
    '***** Tool Dimensions *****
    
    'LocatorBottomRotorPM
    Dim LocatorBottomRotorPMBandID As Double 'LocatorBottomRotorPMBandID@Sketch1
    Dim LocatorBottomRotorPMBulletID As Double 'LocatorBottomRotorPMBulletID@Sketch1
    Dim LocatorBottomRotorPMHeight As Double 'LocatorBottomRotorPMHeight@Sketch1
    Dim LocatorBottomRotorPMSlotD As Double 'LocatorBottomRotorPMSlotD@Sketch2
    Dim LocatorBottomRotorPMSlotWidth As Double 'LocatorBottomRotorPMSlotWidth@Sketch2
    Dim LocatorBottomRotorPMSlotDepth As Double 'LocatorBottomRotorPMSlotDepth@Cut-Extrude1
    
    'BulletRotorPM
    Dim BulletRotorPMOD As Double 'BulletRotorPMOD@Sketch1
    Dim BulletRotorPMID As Double 'BulletRotorPMID@Sketch1

    'PlateInstallationPM
    Dim PlateInstallationPMID As Double 'PlateInstallationPMID@Sketch1
    Dim PlateInstallationPMOD As Double 'PlateInstallationPMOD@Sketch1
    Dim PlateInstallationPMSlotD As Double 'PlateInstallationPMSlotD@Sketch2
    Dim PlateInstallationPMSlotWidth As Double 'PlateInstallationPMSlotWidth@Sketch2
    Dim PlateInstallationPMSlotDepth As Double 'PlateInstallationPMSlotDepth@Cut-Extrude1
    
    'ShaftRotorPM
    Dim ShaftRotorPM As Double 'ShaftRotorPM@Sketch1
    
    '***********************************************************************************************************
    '***********************************************************************************************************
    ToolAssemblyFolder = "C:\Users\FIlie\Documents\Felix Documents IPS\Master Tooling\PM Rotor Band Installation\"
    ToolAssemblyPath = ToolAssemblyFolder + "Assembly, Band, Rotor, PM.SLDASM"
    '***********************************************************************************************************
    UnitType = "Agusta 609 AC" '"Agusta 609 DC" ' "Agusta 609 DC", "CH47"
    '***********************************************************************************************************
    '***********************************************************************************************************
    
    Select Case UnitType

        Case "Agusta 609 AC"
        
            AssemblyName = "1034-13-05979 Assembly, Rotor, Pm-1"
    
            PMRotorOD = 3.202 + 2 * 0.032 'max 3.202 + 2 * .032
            PMRotorID = 0.781
            PMRotorThick = 0.577
            ScrewLocationD = 2.96
            ScrewD = 0.112 'Dash 4 screw
            ScrewProtrudeDepth = 0.048

        Case "Agusta 609 DC"
        
            AssemblyName = "1034-13-07069 Assembly, Rotor, Pm-1"

            PMRotorOD = 3.127 + 2 * 0.032 'max 3.127 + 2 * .032 DIMESNION GIVEN IS FOR INNER DIAMETER OF BAND!
            PMRotorID = 0.781
            PMRotorThick = 0.674
            ScrewLocationD = 2.879
            ScrewD = 0.112 'Dash 4 screw
            ScrewProtrudeDepth = 0.041
            
        Case "CH47"
            
            AssemblyName = "1029-13-06090 Assembly, Rotor, Pm-1"

            PMRotorOD = 2.686 ' -.000 +.002 DIMESNION GIVEN IS FOR INNER DIAMETER OF BAND!
            PMRotorID = 0.788
            PMRotorThick = 0.507
            ScrewLocationD = 0
            ScrewD = 0
            ScrewProtrudeDepth = 0

        Case Else
            swApp.SendMsgToUser2 "Data for this unit is not available", swMbStop, swMbOk
            errors = False
            Exit Sub
    End Select
       
    '***** Calculating Tool Dimensions *****
    
    'ShaftRotorPM
    ShaftRotorPM = PMRotorID - 0.002
    
    Debug.Print ShaftRotorPM
    
    'BulletRotorPM
    BulletRotorPMID = ShaftRotorPM + 0.003
    BulletRotorPMOD = BulletRotorPMID + 0.26
    
    Debug.Print BulletRotorPMOD, BulletRotorPMOD
    
    'LocatorBottomRotorPM
    LocatorBottomRotorPMBandID = PMRotorOD + 0.002
    LocatorBottomRotorPMBulletID = BulletRotorPMOD + 0.002
    LocatorBottomRotorPMHeight = PMRotorThick - 0.1
    LocatorBottomRotorPMSlotD = ScrewLocationD
    LocatorBottomRotorPMSlotWidth = ScrewD + 0.03 + 0.1 ' + .115 for 609 DC, + .1 for 609 AC
    LocatorBottomRotorPMSlotDepth = ScrewProtrudeDepth + 0.03
    
    Debug.Print LocatorBottomRotorPMBandID, LocatorBottomRotorPMBulletID, LocatorBottomRotorPMHeight
    
    'PlateInstallationPM
    PlateInstallationPMID = ShaftRotorPM + 0.003
    PlateInstallationPMOD = PMRotorOD + 0.1
    PlateInstallationPMSlotD = ScrewLocationD
    PlateInstallationPMSlotWidth = ScrewD + 0.03 + 0.1 ' + .115 for 609 DC, + .1 for 609 AC
    PlateInstallationPMSlotDepth = ScrewProtrudeDepth + 0.03
    
    Debug.Print PlateInstallationPMID, PlateInstallationPMOD


    '***** Changing Tool Dimensions *****
    ' DONT FORGET TO CONVERT TO METERS!!!!
    ' DEG in RADIANS!!!!

    Set swModel = swApp.OpenDoc6(ToolAssemblyPath, _
    swDocASSEMBLY, swOpenDocOptions_Silent, "", lErrors, lWarnings)

    'ShaftRotorPM
    swApp.ActivateDoc "Shaft, Installation, Rotor, PM"
    Set swModel = swApp.ActiveDoc
    swModel.Parameter("ShaftRotorPM@Sketch1").SystemValue = ShaftRotorPM * inTOmeter
    swModel.EditRebuild3
    swModel.Save3 swSaveAsOptions_Silent, lErrors, lWarnings

    'BulletRotorPM
    swApp.ActivateDoc "Bullet, Installation, Rotor, PM"
    Set swModel = swApp.ActiveDoc
    Set swPart = swModel

    swModel.Parameter("BulletRotorPMOD@Sketch1").SystemValue = BulletRotorPMOD * inTOmeter
    swModel.Parameter("BulletRotorPMID@Sketch1").SystemValue = BulletRotorPMID * inTOmeter

    swModel.EditRebuild3
    swModel.Save3 swSaveAsOptions_Silent, lErrors, lWarnings

    'LocatorBottomRotorPM
    swApp.ActivateDoc "Bottom, Installation, Rotor, PM"
    Set swModel = swApp.ActiveDoc
    swModel.Parameter("LocatorBottomRotorPMBandID@Sketch1").SystemValue = LocatorBottomRotorPMBandID * inTOmeter
    swModel.Parameter("LocatorBottomRotorPMBulletID@Sketch1").SystemValue = LocatorBottomRotorPMBulletID * inTOmeter
    swModel.Parameter("LocatorBottomRotorPMHeight@Sketch1").SystemValue = LocatorBottomRotorPMHeight * inTOmeter
    swModel.Parameter("LocatorBottomRotorPMSlotD@Sketch2").SystemValue = LocatorBottomRotorPMSlotD * inTOmeter
    swModel.Parameter("LocatorBottomRotorPMSlotWidth@Sketch2").SystemValue = LocatorBottomRotorPMSlotWidth * inTOmeter
    swModel.Parameter("LocatorBottomRotorPMSlotDepth@Cut-Extrude1").SystemValue = LocatorBottomRotorPMSlotDepth * inTOmeter
    
    swModel.EditRebuild3
    swModel.Save3 swSaveAsOptions_Silent, lErrors, lWarnings

    'PlateInstallationPM
    swApp.ActivateDoc "Plate, Installation, Rotor, PM"
    Set swModel = swApp.ActiveDoc
    swModel.Parameter("PlateInstallationPMID@Sketch1@Sketch1").SystemValue = PlateInstallationPMID * inTOmeter
    swModel.Parameter("PlateInstallationPMOD@Sketch1").SystemValue = PlateInstallationPMOD * inTOmeter
    swModel.Parameter("PlateInstallationPMSlotD@Sketch2").SystemValue = PlateInstallationPMSlotD * inTOmeter
    swModel.Parameter("PlateInstallationPMSlotWidth@Sketch2").SystemValue = PlateInstallationPMSlotWidth * inTOmeter
    swModel.Parameter("PlateInstallationPMSlotDepth@Cut-Extrude1").SystemValue = PlateInstallationPMSlotDepth * inTOmeter
    
    swModel.EditRebuild3
    swModel.Save3 swSaveAsOptions_Silent, lErrors, lWarnings

    'Assembly
    swApp.ActivateDoc "Assembly, Band, Rotor, PM"
    Set swModel = swApp.ActiveDoc
    Set swAssy = swModel

    'Unsuppress only the relevant core
    'First suppress all the cores:
    For i = 0 To UBound(AssemblyArray)
        Set swComp = swAssy.GetComponentByName(AssemblyArray(i))
        If Not swComp Is Nothing Then
            Debug.Print swComp.Name
            swComp.SetSuppression2 swComponentSuppressed
        End If
    Next i
    'Next, unsuppress only the relevant core
    Set swComp = swAssy.GetComponentByName(AssemblyName)
    swComp.SetSuppression2 swComponentFullyResolved

    swModel.Extension.Rebuild swForceRebuildAll
    'EditRebuild3
    swModel.Save3 swSaveAsOptions_Silent, lErrors, lWarnings

End Sub






