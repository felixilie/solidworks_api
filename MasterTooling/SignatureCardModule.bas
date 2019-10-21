Attribute VB_Name = "SignatureCardModule"
Sub SignatureCard(PartName As String, PN As String)

    'F0009 Rev A - Document Card Field:
    'signPartNumber, signPartName, signProgram, signDesigner,
    'signDate, signMechEng, MechEngAvaliable, signElecEng, ElecEngAvaliable,
    'MaterialEngAvaliable, ComponentEngAvaliable, QualityAvaliable, ProcessEngAvaliable,
    'signMgr, ProgramAvaliable
    
    'Box Field:
    'UsedToMakeBox, DesignerBox, MechDesignerBox, ElectricalEngBox, MaterialEngBox
    'QualityBox, CompEngBox, ProcessEngBox,ProjectMgrBox

    Dim wdApp As Word.Application 'all those objects are under the Excel Application
    Dim wdWB As Word.Document
    Dim LDate As String
    
    'UserName = Environ("USERNAME") 'Functions used to get Users name
    UserName = "F. Ilie" 'Since Computer Name is Lenovo.....
    LDate = Date
    
    'So the problem you are having with Word asking you to save the global template is because there is already a copy Word open which has rights to the Normal template.
    'When you use CreateObject to set your Word object you are loading up Word a second time which opens Normal template as read only.
    'What you need to do is check if Word is open or not and if it is grab that copy of Word. If it's not then you can open up Word.
    
    
    'We need to continue through errors since if Word isn't
    'open the GetObject line will give an error
    On Error Resume Next
    Set wdApp = GetObject(, "Word.Application")
    
    'We've tried to get Word but if it's nothing then it isn't open
    If wdApp Is Nothing Then
        Set wdApp = CreateObject("Word.Application")
    End If
    
    'It's good practice to reset error warnings
    On Error GoTo 0
    
    'Open your document and ensure its visible and activate after openning
    Set wdWB = wdApp.Documents.Open("C:\Users\FIlie\Documents\Felix Documents IPS\API\F0009 Rev A - Document Card.doc")
    wdApp.Visible = True
    wdApp.Activate
    
    'Set swDraw = swView.ReferencedDocument 'Added for get the reference document to the view
    
    wdWB.SaveAs2 "C:\Users\FIlie\Documents\Felix Documents IPS\API\SignatureCards\" + PN + " " + Format(Now(), "yyyymmddhhnnss") + ".DOC"
    
    'wdFormatXMLDocumentMacroEnabled "C:\Users\FIlie\Documents\Felix Documents IPS\API\SignatureCards " +
    
    With wdApp.Selection.Find
        .ClearFormatting
        .Text = "signPartNumber"
        .Replacement.ClearFormatting
        .Replacement.Text = PN
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    
        .ClearFormatting
        .Text = "signPartName"
        .Replacement.ClearFormatting
        .Replacement.Text = PartName
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    
        .ClearFormatting
        .Text = "signProgram"
        .Replacement.ClearFormatting
        .Replacement.Text = MasterToolGeneratorForm.UnitListBox.Value
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    
        .ClearFormatting
        .Text = "signDesigner"
        .Replacement.ClearFormatting
        .Replacement.Text = MasterToolGeneratorForm.DesignerBox.Value
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    
        .ClearFormatting
        .Text = "signDate"
        .Replacement.ClearFormatting
        .Replacement.Text = Format(Now(), "mm/dd/yyyy")
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
        
        If MasterToolGeneratorForm.MechDesignerBox.Value = "" Or MasterToolGeneratorForm.MechDesignerBox.Value = "----" Then
            .ClearFormatting
            .Text = "signMechEng"
            .Replacement.ClearFormatting
            .Replacement.Text = ""
            .Execute Replace:=wdReplaceAll, Forward:=True, _
            Wrap:=wdFindContinue
            .ClearFormatting
            .Text = "MechEngAvaliable"
            .Replacement.ClearFormatting
            .Replacement.Text = "------------------------N/A------------------------"
            .Execute Replace:=wdReplaceAll, Forward:=True, _
            Wrap:=wdFindContinue
        Else
            .ClearFormatting
            .Text = "signMechEng"
            .Replacement.ClearFormatting
            .Replacement.Text = MasterToolGeneratorForm.MechDesignerBox.Value
            .Execute Replace:=wdReplaceAll, Forward:=True, _
            Wrap:=wdFindContinue
            .ClearFormatting
            .Text = "MechEngAvaliable"
            .Replacement.ClearFormatting
            .Replacement.Text = ""
            .Execute Replace:=wdReplaceAll, Forward:=True, _
            Wrap:=wdFindContinue
        End If
        
        If MasterToolGeneratorForm.ElectricalEngBox.Value = "" Or MasterToolGeneratorForm.ElectricalEngBox.Value = "----" Then
            .ClearFormatting
            .Text = "signElecEng"
            .Replacement.ClearFormatting
            .Replacement.Text = ""
            .Execute Replace:=wdReplaceAll, Forward:=True, _
            Wrap:=wdFindContinue
            .ClearFormatting
            .Text = "ElecEngAvaliable"
            .Replacement.ClearFormatting
            .Replacement.Text = "------------------------N/A------------------------"
            .Execute Replace:=wdReplaceAll, Forward:=True, _
            Wrap:=wdFindContinue
        Else
            .ClearFormatting
            .Text = "signElecEng"
            .Replacement.ClearFormatting
            .Replacement.Text = MasterToolGeneratorForm.ElectricalEngBox.Value
            .Execute Replace:=wdReplaceAll, Forward:=True, _
            Wrap:=wdFindContinue
            .ClearFormatting
            .Text = "ElecEngAvaliable"
            .Replacement.ClearFormatting
            .Replacement.Text = ""
            .Execute Replace:=wdReplaceAll, Forward:=True, _
            Wrap:=wdFindContinue
        End If
        
        If MasterToolGeneratorForm.ProjectMgrBox.Value = "" Or MasterToolGeneratorForm.ProjectMgrBox.Value = "----" Then
            .ClearFormatting
            .Text = "signMgr"
            .Replacement.ClearFormatting
            .Replacement.Text = ""
            .Execute Replace:=wdReplaceAll, Forward:=True, _
            Wrap:=wdFindContinue
            .ClearFormatting
            .Text = "ProgramAvaliable"
            .Replacement.ClearFormatting
            .Replacement.Text = "------------------------N/A------------------------"
            .Execute Replace:=wdReplaceAll, Forward:=True, _
            Wrap:=wdFindContinue
        Else
            .ClearFormatting
            .Text = "signMgr"
            .Replacement.ClearFormatting
            .Replacement.Text = MasterToolGeneratorForm.ProjectMgrBox.Value
            .Execute Replace:=wdReplaceAll, Forward:=True, _
            Wrap:=wdFindContinue
            .ClearFormatting
            .Text = "ProgramAvaliable"
            .Replacement.ClearFormatting
            .Replacement.Text = ""
            .Execute Replace:=wdReplaceAll, Forward:=True, _
            Wrap:=wdFindContinue
        End If
        
        If MasterToolGeneratorForm.MaterialEngBox.Value = "" Or MasterToolGeneratorForm.MaterialEngBox.Value = "----" Then
            .ClearFormatting
            .Text = "MaterialEngAvaliable"
            .Replacement.ClearFormatting
            .Replacement.Text = "------------------------N/A------------------------"
            .Execute Replace:=wdReplaceAll, Forward:=True, _
            Wrap:=wdFindContinue
        Else
            .ClearFormatting
            .Text = "MaterialEngAvaliable"
            .Replacement.ClearFormatting
            .Replacement.Text = ""
            .Execute Replace:=wdReplaceAll, Forward:=True, _
            Wrap:=wdFindContinue
        End If
        
        If MasterToolGeneratorForm.CompEngBox.Value = "" Or MasterToolGeneratorForm.CompEngBox.Value = "----" Then
            .ClearFormatting
            .Text = "ComponentEngAvaliable"
            .Replacement.ClearFormatting
            .Replacement.Text = "------------------------N/A------------------------"
            .Execute Replace:=wdReplaceAll, Forward:=True, _
            Wrap:=wdFindContinue
        Else
            .ClearFormatting
            .Text = "ComponentEngAvaliable"
            .Replacement.ClearFormatting
            .Replacement.Text = ""
            .Execute Replace:=wdReplaceAll, Forward:=True, _
            Wrap:=wdFindContinue
        End If
        
        If MasterToolGeneratorForm.QualityBox.Value = "" Or MasterToolGeneratorForm.QualityBox.Value = "----" Then
            .ClearFormatting
            .Text = "QualityAvaliable"
            .Replacement.ClearFormatting
            .Replacement.Text = "------------------------N/A------------------------"
            .Execute Replace:=wdReplaceAll, Forward:=True, _
            Wrap:=wdFindContinue
        Else
            .ClearFormatting
            .Text = "QualityAvaliable"
            .Replacement.ClearFormatting
            .Replacement.Text = ""
            .Execute Replace:=wdReplaceAll, Forward:=True, _
            Wrap:=wdFindContinue
        End If
        
        If MasterToolGeneratorForm.ProcessEngBox.Value = "" Or MasterToolGeneratorForm.ProcessEngBox.Value = "----" Then
            .ClearFormatting
            .Text = "ProcessEngAvaliable"
            .Replacement.ClearFormatting
            .Replacement.Text = "------------------------N/A------------------------"
            .Execute Replace:=wdReplaceAll, Forward:=True, _
            Wrap:=wdFindContinue
        Else
            .ClearFormatting
            .Text = "ProcessEngAvaliable"
            .Replacement.ClearFormatting
            .Replacement.Text = ""
            .Execute Replace:=wdReplaceAll, Forward:=True, _
            Wrap:=wdFindContinue
        End If
        
        
    End With
    
'    DocumentType = swView.ReferencedDocument.GetType
    
    'Checks what kind of file is open, executes only if it is a Drawing document
    
'    If DocumentType = swDocASSEMBLY Then '= swDocASSEMBLY Or DocumentType = swDocPART - Other way to do that
'
'        wdWB.SaveAs2 "C:\Users\FIlie\Documents\Felix Documents IPS\API\SignatureCards\PL" + PN + " " + Format(Now(), "yyyymmddhhnnss") + ".DOC"
'
'        With wdApp.Selection.Find
'            .ClearFormatting
'            .Text = PN
'            .Replacement.ClearFormatting
'             'Debug.Print ("PL PN") & PN
'            .Replacement.Text = "PL" + PN
'            .Execute Replace:=wdReplaceAll, Forward:=True, _
'            Wrap:=wdFindContinue
'        End With
'
'
'    End If

    wdWB.SaveAs2
    
    Set wdWB = Nothing
    Set wdApp = Nothing

End Sub

