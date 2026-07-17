Attribute VB_Name = "UserForm3"
Private Sub CommandButton1_Click()

End Sub

Private Sub CommandButton2_Click()

PlanillaDeAmasados

End Sub

Private Sub CommandButton4_Click()
Sheets("M.M").Visible = 1
Sheets("M.M Semo").Visible = 1

Sheets("M.M").Select
Range("A1").Select
End Sub

Private Sub CommandButton5_Click()


Sheets("Fermentacion").Visible = 1
Sheets("Coccion").Visible = 1
Sheets("Abatimiento").Visible = 1
Sheets("Envasado").Visible = 1
PlanillasDeProceso
Unload Me

End Sub

Private Sub CommandButton6_Click()
Unload Me

End Sub

