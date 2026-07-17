Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub Workbook_Open()
Application.ScreenUpdating = False
Sheets("M.M").Visible = 0
Sheets("M.M Semo").Visible = 0
Sheets("ART").Visible = 0




Sheets("Fermentacion").Visible = 1
Sheets("Fermentacion").Select
Range("A9:B39").Value = ""
Sheets("Fermentacion").Visible = 0



Sheets("Coccion").Visible = 1
Sheets("Coccion").Select
Range("A9:B39").Value = ""
Sheets("Coccion").Visible = 0

Sheets("Abatimiento").Visible = 1
Sheets("Abatimiento").Select
Range("A9:B9").Value = ""
Sheets("Abatimiento").Visible = 0

Sheets("Envasado").Visible = 1
Sheets("Envasado").Select
Range("A9:B9").Value = ""
Sheets("Envasado").Visible = 0

Sheets("Amasado").Select
Sheets("A").Visible = 0
Sheets("B").Visible = 0
Sheets("C").Visible = 0
Sheets("D").Visible = 0
Sheets("E").Visible = 0
Sheets("F").Visible = 0
Sheets("G").Visible = 0
Sheets("H").Visible = 0
Sheets("I").Visible = 0
Range("A10:D24").Value = ""
Range("A10").Select






End Sub

