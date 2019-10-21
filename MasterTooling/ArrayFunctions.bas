Attribute VB_Name = "ArrayFunctions"
Option Explicit

Public swApp As SldWorks.SldWorks
Public swModel As SldWorks.ModelDoc2
Public swDraw As SldWorks.DrawingDoc
Public assemb As SldWorks.AssemblyDoc
Public swComponent As SldWorks.Component2
Public swView As SldWorks.View
Public swConfig As SldWorks.Configuration
Public selMgr As SldWorks.SelectionMgr
Public swNote As SldWorks.Note
Public swSheet As SldWorks.Sheet
Public assemblyFilePath As String
Public assemblyPath As String
Public NewAssemblyFilePath As String
Public savetofolder As String
Public ReleaseFolderPath As String

Public toolNames() As Variant
Public unitList() As Variant
Public unitsNameList() As Variant
Public materialsList() As Variant
Public materialsNoteList() As Variant
Public FileNames() As Variant
Public DrawingsPaths() As String
Public MainRotorBrazingFixturePath As String
Public UnitNumbers As Variant
Public HomeFolder As Scripting.Folder
Public GetPackandGoStatus As Boolean
Public SetPackandGoStatus As Boolean

Public warnings As Long
Public errors As Long


Public Sub intiateArrays()
    toolNames = Array("Main Rotor Coils Brazing", "Main Rotor Stacking", _
    "Main Rotor Coil Forming", "Main Rotor Coil Pressing", _
    "Main Rotor Baking", "Main Stator Stacking", _
    "Hair Pin Coil Forming", "Main Stator Baking", _
    "Exciter Rotor Stacking", "Exciter Rotor Baking", _
    "Exciter Stator Stacking", "Exciter Stator Baking", _
    "PM Stator Stacking", "PM Stator Coil Press", _
    "PM Balance", "Fan Balance", "Other File")
    
    unitList = Array("Agusta 169 GN", "Agusta 609 AC GN", "Agusta 609 DC GN", _
    "Bell 525 GN", "Boeing Uclass 30kVA", "Boeing Uclass 600A", _
    "Boeing CH-47 GN", "Cessna Starter Generator", _
    "Cessna Scorpion GN", "Cessna Latitude GN", "Embraer A4 GN", _
    "Embraer KC390 GN", "SAAB GN", "Textron GN", "TRU1")
    
    'NEED TO FIT THE unitList Array!
    unitsNameList = Array("SGA11-300-2-A", "GNA16-25A-2-A", "GNA15-400D-2-A", _
    "GNA19-10A-2A", "GNA14-30A-2A", "SGA18-600-2A", _
    "GNA17-15A-1A ", "SGA1-300-2A", _
    "GNA18-8A-2A", "GNA12-5A-2A", "GNA9-20A-1A", _
    "GNO9-90A-2A", "GNO6-60A-2A", "GNO12-150A-1A", _
    "TRU1-125-1A")
    
    'materialsList and materialsNoteList arrays must be same size
    materialsList = Array("ALUMINUM 6061", "4140/4142 ALLOY", "CARBON STEEL", _
    "EPOXY", "WOOD", "Graphit", "GLASS", "GAROLITE", "SEE COMPONENTS")
    
    
    materialsNoteList = Array("ALUMINUM 6061" & Chr(13) & "PER ASTM B-209, B-211, B-221", _
    "4140/4142 ALLOY" & Chr(13) & "PER ASTM A-29", "CARBON STEEL", "EPOXY", "WOOD", _
    "Graphit", "GLASS", "GAROLITE", "SEE COMPONENTS")
    
    'MainRotorBrazingFixturePath = "C:\Users\FIlie\Documents\Felix Documents IPS\Master Tooling\Main rotor coils Brazing tool\Fixed\"
    
    ReleaseFolderPath = "Z:\Documents To Be Released\Drawings and PL's\Felix\"
    'Z:\Documents To Be Released\Drawings and PL's\Felix\ '"C:\Users\FIlie\Documents\Felix Documents IPS\API\Released Documents\" 'For Testing
    
    GetPackandGoStatus = False
    
End Sub

    



