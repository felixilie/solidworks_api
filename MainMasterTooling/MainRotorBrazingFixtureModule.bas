Attribute VB_Name = "MainRotorBrazingFixtureModule"


Option Explicit

Const PI = 3.14159265358979

Sub MainRotorBrazingFixture()

    Dim swApp As SldWorks.SldWorks
    Dim swModel As SldWorks.ModelDoc2
    Dim swModelDocExt As SldWorks.ModelDocExtension
    Dim lErrors As Long
    Dim lWarnings As Long
    Dim errors As Boolean
    Dim value As Boolean
    
    Dim inTOmeter As Double
    Dim meterToin As Double
    Dim degTORad As Double
    
    Dim ToolAssemblyPath As String
    Dim ToolAssemblyFolder As String
    Dim UnitType As String
    
    'Generator Dimension - determine the distance between 2 coils when spread out on the fixture.
    Dim RtoCoil As Double
    Dim NumberCoils As Integer

    'Coil Dimensions
    Dim CoilWidth As Double
    Dim CoilLength As Double
    Dim coilHeight As Double
    Dim CoilRadius As Double
    Dim wireWidth As Double
    
    'Tool Dimensions
    Dim DowelPinDIA As Double
    Dim ToolhalfCoiltoCoil As Double
    Dim ToolhalfCoilWidth As Double
    Dim ToolCoilLength As Double
    Dim ToolCoilHeight As Double
    Dim ToolDistanceSidePinX As Double
    Dim ToolDistanceSidePinY As Double
    Dim ToolWidth As Double
    Dim BottomToolWidth As Double
    Dim ToolLength As Double
    Dim ToolFilet As Double
    
    Dim boolstatus As Boolean
    Dim longstatus As Long, longwarnings As Long
    
    Set swApp = Application.SldWorks
    
    inTOmeter = 0.0254 'Units are in meters
    meterToin = 1 / inTOmeter 'To convert back to inch for checking results
    degTORad = PI / 180 'From degrees to Radians
    
    '***********************************************************************************************************
    '***********************************************************************************************************
    ToolAssemblyFolder = "C:\Users\FIlie\Documents\Felix Documents IPS\Master Tooling\Main rotor coils Brazing tool\"
    ToolAssemblyPath = ToolAssemblyFolder + "Assembly, Tool, Brazing, Coils, Rotor, Main.sldasm"
    UnitType = "Agusta 169 GN" ' "Agusta 609 AC GN", "Agusta 609 DC GN", _
    "Bell 525 GN", "Boeing CH-47 GN", "Cessna Latitude GN", "Embraer A4 GN", _
    "Embraer KC390 GN", "SAAB GN", "Scorpion GN", "TAI GN", "Textron GN", "TRU1")
    
    Select Case UnitType

        Case "Embraer A4 GN"
                RtoCoil = 2.466
                NumberCoils = 12

                CoilWidth = 0.69 '+.020/-.000
                CoilLength = 3.598 '+.020/-.000
                coilHeight = 0.6
                wireWidth = 0.305 'Max wire width
                CoilRadius = 0.31
        Case "Agusta 169 GN"
                RtoCoil = 1.214
                NumberCoils = 8

                CoilWidth = 0.544 '+.020/-.000
                CoilLength = 3.715 '+.020/-.000
                coilHeight = 0.6
                wireWidth = 0.205 'Max wire width
                CoilRadius = 0.25
        Case "Cessna Latitude GN"
                RtoCoil = 0.866
                NumberCoils = 4

                CoilWidth = 0.93 '+.020/-.000
                CoilLength = 2.75 '+.020/-.000
                coilHeight = 0.5
                wireWidth = 0.305 'Max wire width
                CoilRadius = 0.31
        Case "Boeing CH-47 GN"
                RtoCoil = 0.822
                NumberCoils = 4

                CoilWidth = 0.93 '+.020/-.000
                CoilLength = 3.475 '+.020/-.000
                coilHeight = 0.55
                wireWidth = 0.308 'Max wire width
                CoilRadius = 0.31
        Case "Agusta 609 AC GN"
                RtoCoil = 0.882
                NumberCoils = 4

                CoilWidth = 1.115 '+.020/-.000
                CoilLength = 3.798 '+.020/-.000
                coilHeight = 0.667
                wireWidth = 0.305 'Max wire width
                CoilRadius = 0.312

        Case "Agusta 609 DC GN"
                RtoCoil = 2.172
                NumberCoils = 12

                CoilWidth = 0.453  '+.020/-.000
                CoilLength = 2.651 '+.020/-.000
                coilHeight = 0.496
                wireWidth = 0.255  'Max wire width
                CoilRadius = 0.432 / 2 ' .22

        Case Else
            swApp.SendMsgToUser "Data for this unit is not available"
            errors = False
            Exit Sub
    End Select
    
    DowelPinDIA = 0.25
    
    ToolhalfCoiltoCoil = Round(Tan(360 / NumberCoils / 2 * degTORad) * RtoCoil, 3)
    Debug.Print ("ToolhalfCoiltoCoil ") & ToolhalfCoiltoCoil
    ToolhalfCoilWidth = (CoilWidth - 0.005) / 2
    Debug.Print (Chr(10) + "ToolhalfCoilWidth ") & ToolhalfCoilWidth
    ToolCoilLength = CoilLength - 0.01
    Debug.Print (Chr(10) + "ToolCoilLength ") & ToolCoilLength
    ToolDistanceSidePinX = CoilWidth / 2 + 0.005 + wireWidth + DowelPinDIA / 2
    Debug.Print (Chr(10) + "ToolDistanceSidePinX ") & ToolDistanceSidePinX
    ToolDistanceSidePinY = CoilLength / 2 - CoilRadius - 0.01
    Debug.Print (Chr(10) + "ToolDistanceSidePinY ") & ToolDistanceSidePinY
    ToolCoilHeight = coilHeight + 0.6
    Debug.Print (Chr(10) + "ToolCoilHeight ") & ToolCoilHeight
    ToolLength = Round(ToolCoilLength + 2 * ToolhalfCoilWidth + 0.2, 1)
    Debug.Print (Chr(10) + "ToolLength ") & ToolLength
    ToolWidth = Round(2 * (ToolhalfCoilWidth + ToolhalfCoiltoCoil) + wireWidth * 2 + 2 * DowelPinDIA + 0.4, 1)
    Debug.Print (Chr(10) + "ToolWidth ") & ToolWidth
    BottomToolWidth = ToolWidth + 0.5
    Debug.Print (Chr(10) + "BottomToolWidth ") & BottomToolWidth
    ToolFilet = CoilRadius
    Debug.Print (Chr(10) + "ToolFilet ") & ToolFilet
    
    Set swModel = swApp.OpenDoc6(ToolAssemblyPath, _
    swDocASSEMBLY, swOpenDocOptions_OverrideDefaultLoadLightweight, "", lErrors, lWarnings)
    
    'Apperantly, ActivateDoc also opens a file if is part of the Assembly
    swApp.ActivateDoc "Main Plate, Fixture, Brazing, Winding, Rotor, Main"
    Set swModel = swApp.ActiveDoc
    
    ' DONT FORGET TO CONVERT TO METERS!!!!
    
    swModel.Parameter("ToolhalfCoiltoCoil@MainSketch").SystemValue = ToolhalfCoiltoCoil * inTOmeter
    
    swModel.Parameter("ToolhalfCoilWidth@MainSketch").SystemValue = ToolhalfCoilWidth * inTOmeter
    
    swModel.Parameter("ToolCoilLength@MainSketch").SystemValue = ToolCoilLength * inTOmeter
    
    swModel.Parameter("ToolDistanceSidePinX@MainSketch").SystemValue = ToolDistanceSidePinX * inTOmeter
    
    swModel.Parameter("ToolDistanceSidePinY@MainSketch").SystemValue = ToolDistanceSidePinY * inTOmeter
    
    swModel.Parameter("ToolWidth@Sketch3").SystemValue = ToolWidth * inTOmeter
    
    swModel.Parameter("ToolLength@Sketch3").SystemValue = ToolLength * inTOmeter
    
    swModel.EditRebuild3
    swModel.Save3 swSaveAsOptions_Silent, lErrors, lWarnings
    
    
    swApp.ActivateDoc "Center, Tool, Brazing, Coils, Rotor, Main"
    Set swModel = swApp.ActiveDoc
    
    swModel.Parameter("ToolhalfCoiltoCoil@MainSketch").SystemValue = ToolhalfCoiltoCoil * inTOmeter
    
    swModel.Parameter("ToolhalfCoilWidth@MainSketch").SystemValue = ToolhalfCoilWidth * inTOmeter
    
    swModel.Parameter("ToolCoilLength@MainSketch").SystemValue = ToolCoilLength * inTOmeter
    
    swModel.Parameter("ToolDistanceSidePinX@MainSketch").SystemValue = ToolDistanceSidePinX * inTOmeter
    
    swModel.Parameter("ToolDistanceSidePinY@MainSketch").SystemValue = ToolDistanceSidePinY * inTOmeter
    
    swModel.Parameter("ToolCoilHeight@Boss-Extrude1").SystemValue = ToolCoilHeight * inTOmeter
    
    swModel.Parameter("ToolFilet@Fillet1").SystemValue = ToolFilet * inTOmeter
    
    swModel.EditRebuild3
    swModel.Save3 swSaveAsOptions_Silent, lErrors, lWarnings
    
    
    swApp.ActivateDoc "Bottom Plate, Tool, Brazing, Coils, Rotor, Main"
    Set swModel = swApp.ActiveDoc
    Set swModelDocExt = swModel.Extension
    
    swModel.Parameter("ToolhalfCoiltoCoil@MainSketch").SystemValue = ToolhalfCoiltoCoil * inTOmeter
    
    swModel.Parameter("ToolhalfCoilWidth@MainSketch").SystemValue = ToolhalfCoilWidth * inTOmeter
    
    swModel.Parameter("ToolCoilLength@MainSketch").SystemValue = ToolCoilLength * inTOmeter
    
    swModel.Parameter("ToolDistanceSidePinX@MainSketch").SystemValue = ToolDistanceSidePinX * inTOmeter
    
    swModel.Parameter("ToolDistanceSidePinY@MainSketch").SystemValue = ToolDistanceSidePinY * inTOmeter
    
    swModel.Parameter("BottomToolWidth@MainSketch").SystemValue = BottomToolWidth * inTOmeter
    
    swModel.Parameter("ToolLength@MainSketch").SystemValue = ToolLength * inTOmeter
    
    swModel.EditRebuild3
    swModel.Save3 swSaveAsOptions_Silent, lErrors, lWarnings
    
    
    swApp.ActivateDoc "Assembly, Tool, Brazing, Coils, Rotor, Main"
    Set swModel = swApp.ActiveDoc
    
    swModel.EditRebuild3
    swModel.Save3 swSaveAsOptions_Silent, lErrors, lWarnings

End Sub

Sub MainRotorBrazingFixtureParameters(UnitType As String, ByRef DimensionsArray() As Variant, ByRef toolAssembleyPath As String, Optional ByRef errors As Boolean = True)

    Dim swApp As SldWorks.SldWorks
    
    toolAssembleyPath = MainRotorBrazingFixturePath + "Assembly, Tool, Brazing, Coils, Rotor, Main.SLDASM"

    'Generator Dimension - determine the distance between 2 coils when spread out on the fixture.
    Dim RtoCoil As Double
    Dim NumberCoils As Integer
    
    'Coil Dimensions
    Dim CoilWidth As Double
    Dim CoilLength As Double
    Dim coilHeight As Double
    Dim CoilRadius As Double
    Dim wireWidth As Double
      
        Select Case UnitType
    
        Case "Embraer A4 GN"
                RtoCoil = 2.466
                NumberCoils = 12
    
                CoilWidth = 0.69 '+.020/-.000
                CoilLength = 3.598 '+.020/-.000
                coilHeight = 0.6
                wireWidth = 0.305 'Max wire width
                CoilRadius = 0.31
        Case "Agusta 169 GN"
                RtoCoil = 1.214
                NumberCoils = 8
    
                CoilWidth = 0.544 '+.020/-.000
                CoilLength = 3.715 '+.020/-.000
                coilHeight = 0.6
                wireWidth = 0.205 'Max wire width
                CoilRadius = 0.25
        Case "Cessna Latitude GN"
                RtoCoil = 0.866
                NumberCoils = 4
    
                CoilWidth = 0.93 '+.020/-.000
                CoilLength = 2.75 '+.020/-.000
                coilHeight = 0.5
                wireWidth = 0.305 'Max wire width
                CoilRadius = 0.31
        Case "Boeing CH-47 GN"
                RtoCoil = 0.822
                NumberCoils = 4
    
                CoilWidth = 0.93 '+.020/-.000
                CoilLength = 3.475 '+.020/-.000
                coilHeight = 0.55
                wireWidth = 0.308 'Max wire width
                CoilRadius = 0.31
        Case "SAAB GN"
                RtoCoil = 1.124
                NumberCoils = 4
        
                CoilWidth = 1.629 '+.020/-.000
                CoilLength = 2.996 '+.020/-.000
                coilHeight = 0.7
                wireWidth = 0.255 'Max wire width
                CoilRadius = 0.25
        Case Else
'            swApp.SendMsgToUser2 "Data for this unit is not available"
            errors = False
            Exit Sub
    End Select
    
    DimensionsArray = Array(RtoCoil, NumberCoils, CoilWidth, CoilLength, _
    coilHeight, wireWidth, CoilRadius)

End Sub





