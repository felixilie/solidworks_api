Attribute VB_Name = "MainRotorBakingModule"
Option Explicit

Const PI = 3.14159265358979

Sub MainRotorBaking()

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
    Dim longstatus As Long
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
    Dim TotalRotorLength As Double
    Dim DEPartLengthToShoulder As Double 'To step where bearing is usually located
    Dim ADEPartLengthToShoulder As Double 'To where Exciter Rotor is usually located
    Dim ExciterRotorOD As Double
    Dim AfterExciterRotorOD As Double
    Dim BearingOD As Double
    Dim AfterBearingOD As Double
  
    '***** Tool Dimensions *****
    
    'DEPlug
    Dim DELengthToShoulder As Double 'DELengthToShoulder@Sketch1
    Dim DEOD As Double 'DEOD@Sketch1
    Dim DEAfterShoulder As Double 'DEAfterShoulder@Sketch1
    Dim TotalDELenght As Double
    
    'ADEPlug
    Dim ADELengthToShoulder As Double 'ADELengthToShoulder@Sketch1
    Dim ADEOD As Double 'ADEOD@Sketch1
    Dim ADEAfterShoulder As Double 'ADEAfterShoulder@Sketch1
    Dim ADEwall As Double 'ADEwall@Sketch1
    

    '***********************************************************************************************************
    '***********************************************************************************************************
    ToolAssemblyFolder = "C:\Users\FIlie\Documents\Felix Documents IPS\Master Tooling\Main Rotor Baking Plugs\"
    ToolAssemblyPath = ToolAssemblyFolder + "Assem1.sldasm"
    '***********************************************************************************************************
    UnitType = "AFRLSG" '"Agusta169", "Agusta 609 AC","Agusta 609 DC", "CH47', "SAAB", "Textron", "Scorpion"
    '***********************************************************************************************************
    '***********************************************************************************************************
    
    Select Case UnitType
    
        Case "SAAB"

            TotalRotorLength = 9.75
            DEPartLengthToShoulder = 1.271
            ADEPartLengthToShoulder = 3.24
            ExciterRotorOD = 1.2
            AfterExciterRotorOD = 1.3
            BearingOD = 0.985
            AfterBearingOD = 1.2

        Case "Bell 525"

            TotalRotorLength = 0
            DEPartLengthToShoulder = 1.019
            ADEPartLengthToShoulder = 3.115
            ExciterRotorOD = 1.02
            AfterExciterRotorOD = 1.1
            BearingOD = 0.67
            AfterBearingOD = 0.875
            
        Case "CH47"

            TotalRotorLength = 0
            DEPartLengthToShoulder = 2.08
            ADEPartLengthToShoulder = 3.3
            ExciterRotorOD = 1.02
            AfterExciterRotorOD = 1.25
            BearingOD = 0.79
            AfterBearingOD = 1
            
        Case "Agusta 609 AC"

            TotalRotorLength = 10.26
            DEPartLengthToShoulder = 0.901
            ADEPartLengthToShoulder = 3.842
            ExciterRotorOD = 1.021
            AfterExciterRotorOD = 1.105 'Max, plus tolerances
            BearingOD = 0.788
            AfterBearingOD = 0.953
            
        Case "Latitude"

            TotalRotorLength = 9.3
            DEPartLengthToShoulder = 0.96
            ADEPartLengthToShoulder = 3.25
            ExciterRotorOD = 1.02
            AfterExciterRotorOD = 1.14 'Max, plus tolerances
            BearingOD = 0.67
            AfterBearingOD = 0.875

        Case "Agusta169"

            TotalRotorLength = 10.16
            DEPartLengthToShoulder = 1.175
            ADEPartLengthToShoulder = 3.85
            ExciterRotorOD = 1.053
            AfterExciterRotorOD = 1.158 'Max, plus tolerances
            BearingOD = 0.788
            AfterBearingOD = 0.987
            
        Case "A4"

            TotalRotorLength = 11.34
            DEPartLengthToShoulder = 2.56
            ADEPartLengthToShoulder = 3.4
            ExciterRotorOD = 1.38
            AfterExciterRotorOD = 1.485 'Max, plus tolerances
            BearingOD = 1.182
            AfterBearingOD = 1.36
            
        Case "AFRLSG"

            TotalRotorLength = 11
            DEPartLengthToShoulder = 1.165
            ADEPartLengthToShoulder = 4.377
            ExciterRotorOD = 1.095
            AfterExciterRotorOD = 1.2 'Max, plus tolerances
            BearingOD = 0.985
            AfterBearingOD = 1.171

        Case Else
            swApp.SendMsgToUser2 "Data for this unit is not available", swMbStop, swMbOk
            errors = False
            Exit Sub
    End Select

    '***** Calculating Tool Dimensions *****
    
    'DEPlug
    DELengthToShoulder = Round(DEPartLengthToShoulder + 0.05, 2)
    DEOD = BearingOD + 0.002
    DEAfterShoulder = AfterBearingOD + 0.01
    
    
    Debug.Print DELengthToShoulder, DEOD, DEAfterShoulder
    
    'ADEPlug
    ADELengthToShoulder = Round(ADEPartLengthToShoulder + 0.05, 2)
    ADEOD = ExciterRotorOD + 0.002
    ADEAfterShoulder = AfterExciterRotorOD + 0.005
    ADEwall = 14 - 2.3 - TotalRotorLength
     
    Debug.Print ADELengthToShoulder, ADEOD, ADEAfterShoulder

    '***** Changing Tool Dimensions *****
    ' DONT FORGET TO CONVERT TO METERS!!!!
    ' DEG in RADIANS!!!!

    Set swModel = swApp.OpenDoc6(ToolAssemblyPath, _
    swDocASSEMBLY, swOpenDocOptions_OverrideDefaultLoadLightweight, "", lErrors, lWarnings)

    'DEPlug
    swApp.ActivateDoc "Baking, Plug, DE, Rotor, Main"
    Set swModel = swApp.ActiveDoc
    swModel.Parameter("DELengthToShoulder@Sketch1").SystemValue = DELengthToShoulder * inTOmeter
    swModel.Parameter("DEOD@Sketch1").SystemValue = DEOD * inTOmeter
    swModel.Parameter("DEAfterShoulder@Sketch1").SystemValue = DEAfterShoulder * inTOmeter
    
    swModel.EditRebuild3
    swModel.Save3 swSaveAsOptions_Silent, lErrors, lWarnings

    'ADEPlug
    swApp.ActivateDoc "Baking, Plug, ADE, Rotor, Main"
    Set swModel = swApp.ActiveDoc
    swModel.Parameter("ADELengthToShoulder@Sketch1").SystemValue = ADELengthToShoulder * inTOmeter
    swModel.Parameter("ADEOD@Sketch1").SystemValue = ADEOD * inTOmeter
    swModel.Parameter("ADEAfterShoulder@Sketch1").SystemValue = ADEAfterShoulder * inTOmeter
    swModel.Parameter("ADEwall@Sketch1").SystemValue = ADEwall * inTOmeter
    
    swModel.EditRebuild3
    swModel.Save3 swSaveAsOptions_Silent, lErrors, lWarnings

    'Assembly
    swApp.ActivateDoc "Assem1.sldasm"
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








