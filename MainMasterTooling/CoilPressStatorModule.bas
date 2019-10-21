Attribute VB_Name = "CoilPressStatorModule"
Option Explicit

Const PI = 3.14159265358979

Sub CoilPressStator() '(UnitType As String, ByRef DimensionsArray As Variant, Optional ByRef errors As Boolean = True)
    
    '******TESTING******
    Dim UnitType As String
    Dim DimensionsArray() As Variant
    Dim errors As Boolean
    
    '******TESTING******
    
    
    Dim swApp As SldWorks.SldWorks
    Dim swModel As SldWorks.ModelDoc2
    Dim swModelDocExt As SldWorks.ModelDocExtension
    Dim lErrors As Long
    Dim lWarnings As Long
    Dim value As Boolean
    
    Dim inTOmeter As Double
    Dim meterToin As Double
    Dim degTORad As Double

    'PM stator Dimensions
    Dim CoilID As Double 'Should be as the model!
    Dim coilOD As Double 'As the drawing indicates per MAX
    Dim coilHeight As Double 'As the drawing indicates per MAX
    Dim leadWidth As Double 'As per the model
    Dim CoreID As Double 'nominal according to drawing
    Dim CoreHeight As Double
    Dim InsulationWidth As Double
    Dim InsulationHeight As Double
    
    'Tool Dimensions
    Dim SlotID As Double
    Dim SlotOD As Double
    Dim SlotHieght As Double
    Dim LeadSlot As Double
    Dim LocatorCoreOD As Double
    Dim LocatorCoilOD As Double
    Dim LocatorHeight As Double
    Dim LocatorID As Double
    Dim DtoCore As Double
    Dim InsulationClearWidth As Double 'InsulationClearWidth@Sketch1
    Dim InsulationClearHeight As Double 'InsulationClearHeight@Sketch1
    
    Dim boolstatus As Boolean
    Dim longstatus As Long, longwarnings As Long
    
    Set swApp = Application.SldWorks
    
    inTOmeter = 0.0254 'Units are in meters
    meterToin = 1 / inTOmeter 'To convert back to inch for checking results
    degTORad = PI / 180 'From degrees to Radians
       
    UnitType = "CH47"
    
    Select Case UnitType

        Case "Embraer A4 GN PM"
                CoilID = 3.433 'From Model
                coilOD = 3.812
                coilHeight = 0.44
                leadWidth = 0.12
                CoreID = 3.249
                CoreHeight = 0.907
                InsulationWidth = 0.062
                InsulationHeight = 0.0625
                
        Case "Agusta 609 AC GN PM"
                CoilID = 3.433 'From Model
                coilOD = 4.08
                coilHeight = 0.32
                leadWidth = 0.12
                CoreID = 3.26
                CoreHeight = 0.378
                InsulationWidth = 0.0635 'Used to be .02
                InsulationHeight = 0.065 'Used to be .02
                
        Case "Agusta 609 AC GN Main Stator"
                CoilID = 3.984 'From Model
                coilOD = 4.73
                coilHeight = 0.758 ' minus .020" than size on drawing
                leadWidth = 0.12
                CoreID = 3.778
                CoreHeight = 3.07
                InsulationWidth = 0.062
                InsulationHeight = 0.062
                
        Case "SAAB PM"
                CoilID = 3.433 'From Model
                coilOD = 3.812
                coilHeight = 0.375
                leadWidth = 0.2
                CoreID = 3.249
                CoreHeight = 0.25
                InsulationWidth = 0.062
                InsulationHeight = 0.062
                
        Case "Textron PM"
                CoilID = 3.433 'From Model
                coilOD = 3.812
                coilHeight = 0.44
                leadWidth = 0.2
                CoreID = 3.249
                CoreHeight = 0.237
                InsulationWidth = 0.062
                InsulationHeight = 0.062
                
        Case "CH47"
                CoilID = 2.994 'From Model
                coilOD = 3.406
                coilHeight = 0.375
                leadWidth = 0.2
                CoreID = 2.81
                CoreHeight = 0.3
                InsulationWidth = 0.062
                InsulationHeight = 0.062
                
'        Case "Agusta 169 GN"
'                RtoCoil = 1.214
'                NumberCoils = 8
'
'                coilWidth = 0.544 '+.020/-.000
'                CoilLength = 3.715 '+.020/-.000
'                coilHeight = 0.6
'                wireWidth = 0.205 'Max wire width
'                CoilRadius = 0.25
'        Case "Cessna Latitude GN"
'                RtoCoil = 0.866
'                NumberCoils = 4
'
'                coilWidth = 0.93 '+.020/-.000
'                CoilLength = 2.75 '+.020/-.000
'                coilHeight = 0.5
'                wireWidth = 0.305 'Max wire width
'                CoilRadius = 0.31
'        Case "Boeing CH-47 GN"
'                RtoCoil = 0.822
'                NumberCoils = 4
'
'                coilWidth = 0.93 '+.020/-.000
'                CoilLength = 3.475 '+.020/-.000
'                coilHeight = 0.55
'                wireWidth = 0.308 'Max wire width
'                CoilRadius = 0.31
        Case Else
            'swApp.SendMsgToUser2 "Data for this unit is not available"
            errors = False
            Exit Sub
    End Select
    
    SlotID = Round(CoilID, 2)
    SlotOD = Round(coilOD - 0.02, 2)
    SlotHieght = coilHeight - 0.04
    LeadSlot = Round(leadWidth / 0.4, 1)
    LocatorCoreOD = CoreID - 0.005
    LocatorCoilOD = SlotID + 0.01
    LocatorHeight = SlotHieght
    LocatorID = Round(LocatorCoreOD - 0.5, 2)
    DtoCore = Round((CoreHeight - 0.2) / 2, 2)
    InsulationClearWidth = InsulationWidth + 0.005
    InsulationClearHeight = InsulationHeight + 0.005
    
    
    Debug.Print SlotID, SlotOD, SlotHieght, LeadSlot, LocatorCoreOD, LocatorCoilOD, LocatorHeight, LocatorID, DtoCore
    
    Set swModel = swApp.OpenDoc6("C:\Users\FIlie\Documents\Felix Documents IPS\Master Tooling\PM Stator Coil Pressing\Press Plate, Coil Pressing, PM, Stator.SLDPRT", _
    swDocPART, swOpenDocOptions_Silent, "", lErrors, lWarnings)

'    'Apperantly, ActivateDoc also opens a file if is part of the Assembly
'    swApp.ActivateDoc "Main Plate, Fixture, Brazing, Winding, Rotor, Main"
'    Set swModel = swApp.ActiveDoc

    '*************** DONT FORGET TO CONVERT TO METERS!!!!*********************
    '***************       Angles in Radians!!!!!!!      *********************
    
    ' ************ Press Plate *************

    swModel.Parameter("SlotOD@Sketch1").SystemValue = SlotOD * inTOmeter

    swModel.Parameter("SlotID@Sketch1").SystemValue = SlotID * inTOmeter

    swModel.Parameter("SlotHieght@Sketch1").SystemValue = SlotHieght * inTOmeter

    swModel.Parameter("LeadSlot@Sketch3").SystemValue = LeadSlot * inTOmeter
    
    swModel.Parameter("DtoCore@Sketch1").SystemValue = DtoCore * inTOmeter
    
    swModel.Parameter("PressToCoreOD@Sketch1").SystemValue = LocatorCoreOD * inTOmeter
    
    swModel.Parameter("InsulationClearWidth@Sketch1").SystemValue = InsulationClearWidth * inTOmeter
    
    swModel.Parameter("InsulationClearHeight@Sketch1").SystemValue = InsulationClearHeight * inTOmeter
    

    swModel.EditRebuild3
    swModel.Save3 swSaveAsOptions_Silent, lErrors, lWarnings
    
    ' ************ Location Ring *************

    Set swModel = swApp.OpenDoc6("C:\Users\FIlie\Documents\Felix Documents IPS\Master Tooling\PM Stator Coil Pressing\Location Ring, Coil Pressing, PM, Stator.SLDPRT", _
    swDocPART, swOpenDocOptions_Silent, "", lErrors, lWarnings)

    swModel.Parameter("LocatorCoreOD@Sketch1").SystemValue = LocatorCoreOD * inTOmeter

    swModel.Parameter("LocatorCoilOD@Sketch1").SystemValue = LocatorCoilOD * inTOmeter

    swModel.Parameter("LocatorHeight@Sketch1").SystemValue = LocatorHeight * inTOmeter
    
    swModel.Parameter("LocatorID@Sketch1").SystemValue = LocatorID * inTOmeter
    
    swModel.Parameter("DtoCore@Sketch1").SystemValue = DtoCore * inTOmeter
    
    swModel.Parameter("InsulationClearWidth@Sketch1").SystemValue = InsulationClearWidth * inTOmeter
    
    swModel.Parameter("InsulationClearHeight@Sketch1").SystemValue = InsulationClearHeight * inTOmeter

    swModel.EditRebuild3
    swModel.Save3 swSaveAsOptions_Silent, lErrors, lWarnings

End Sub

