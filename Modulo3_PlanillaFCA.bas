Attribute VB_Name = "Módulo3"
Sub PlanillaFCA()
'
' PlanillaFCA Macro
' Macro grabada el 01/07/2008 por Alejandro

Dim vpuntamas As String
Dim vnomb As String
Dim vfech As Date
Dim vturn As String
Dim vnumeamas As Integer
Dim vcantamas As Integer
Dim vcanthoja As Integer
Dim vcanthojaauxi As Double
Dim vtotaamasxdia As Integer
Dim vamaspagi As Integer
Dim vcantcrudo As Integer
Dim Dia1 As String
Dim Dia2 As String
Dim Dia3 As String
Dim Dia4 As String
Dim Dia5 As String
Dim Dia6 As String
Dim Dia7  As String

Application.ScreenUpdating = False


'contar la cantidad de amasados según el dia ingresado
vvali = Range("A5").Value
If vvali > Date Then
'vfech = InputBox("Ingrese Fecha a Imprimir Planillas FCA")
Range("A10").Select
vfech = ActiveCell
Do While ActiveCell <> ""
    If ActiveCell = vfech Then
    ActiveCell.Offset(0, 4).Select
        If ActiveCell = "SI" Then
        vcantamas = vcantamas + ActiveCell.Offset(0, -2).Value
        ActiveCell.Offset(1, -4).Select
        Else
        vcantcrudos = vcantcrudos + ActiveCell.Offset(0, -2).Value
        ActiveCell.Offset(1, -4).Select
        End If
    Else
    ActiveCell.Offset(1, 0).Select
    End If
Loop
vcanttotal = vcantamas + vcantcrudos


If vcantamas > 0 Then
vcanthojaauxi = vcantamas / 30
vcanthoja = vcanthojaauxi
        If vcanthojaauxi > vcanthoja Then
          vcanthoja = vcanthoja + 1
        End If
    Else
    MsgBox ("No hay planillas FAC para ese día! Corroborar Datos!")
    Range("A10").Select
   ' End
End If


' pego las hojas necesarias segun la cantidad de amasados en FERMENTACION
Sheets("Fermentacion").Select
Range("B7").Value = vfech
vindipega = "A41"
Range("A9:C38").Select
Selection.ClearContents
Rows("41:500").Select

Selection.Delete Shift:=xlUp
For x = 1 To vcanthoja - 1
    Rows("1:40").Select
    Selection.Copy
    Range(vindipega).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveCell.Offset(40, 0).Select
    vindipega = ActiveCell.Address
Next x
'numerar hojas
Range("G6").Select
For x = 1 To vcanthoja
ActiveCell.Value = "Pagina " & x & " de " & vcanthoja
ActiveCell.Offset(40, 0).Select
Next x
' pego la hojas necesarias segun las cantidad de amasados EN ABATIMIENTO
Sheets("Abatimiento").Select
Range("B7").Value = vfech
vindipega = "A41"
Range("A9:C38").Select
Selection.ClearContents
Rows("41:500").Select
Selection.Delete Shift:=xlUp
For x = 1 To vcanthoja - 1
    Rows("1:40").Select
    Selection.Copy
    Range(vindipega).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveCell.Offset(40, 0).Select
    vindipega = ActiveCell.Address
Next x
'numerar hojas
Range("J6").Select
For x = 1 To vcanthoja
ActiveCell.Value = "Pagina " & x & " de " & vcanthoja
ActiveCell.Offset(40, 0).Select
Next x
' pego las hojas necesarias segun la cantidad de amasados en Coccion
Sheets("coccion").Select
Range("B7").Value = vfech
vindipega = "A41"
Range("A9:C38").Select
Selection.ClearContents
Rows("41:500").Select
Selection.Delete Shift:=xlUp
For x = 1 To vcanthoja - 1
    Rows("1:40").Select
    Selection.Copy
    Range(vindipega).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveCell.Offset(40, 0).Select
    vindipega = ActiveCell.Address
Next x
'numerar hojas
Range("J6").Select
For x = 1 To vcanthoja
ActiveCell.Value = "Pagina " & x & " de " & vcanthoja
ActiveCell.Offset(40, 0).Select
Next x
' completo las planillas con los datos del día
Sheets("Amasado").Select
Range("A10").Select
vpuntamas = ActiveCell.Address
vpuntprin = "A9"
Do While ActiveCell <> ""
If ActiveCell = vfech Then
ActiveCell.Offset(0, 4).Select
    If ActiveCell = "SI" Then
    vcantamas = ActiveCell.Offset(0, -2).Value
    vnomb = ActiveCell.Offset(0, -3).Value
    ActiveCell.Offset(1, -4).Select
    vpuntamas = ActiveCell.Address
    Sheets("Fermentacion").Select
    Range(vpuntprin).Select
    For x = 1 To vcantamas
    If vamaspagiF <= 29 Then
    vamaspagiF = vamaspagiF + 1
    ActiveCell.Value = vfech
    ActiveCell.Offset(0, 1).Value = x
    ActiveCell.Offset(0, 2).Value = vnomb
    ActiveCell.Offset(1, 0).Select
    Else
    vamaspagiF = "0"
    ActiveCell.Offset(10, 0).Select
    vamaspagiF = vamaspagiF + 1
    ActiveCell.Value = vfech
    ActiveCell.Offset(0, 1).Value = x
    ActiveCell.Offset(0, 2).Value = vnomb
    ActiveCell.Offset(1, 0).Select
    End If
    Next x
    
    Sheets("Abatimiento").Select
    Range(vpuntprin).Select
    For x = 1 To vcantamas
    If vamaspagiA <= 29 Then
    vamaspagiA = vamaspagiA + 1
    ActiveCell.Value = vfech
    ActiveCell.Offset(0, 1).Value = x
    ActiveCell.Offset(0, 2).Value = vnomb
    ActiveCell.Offset(1, 0).Select
    Else
    ActiveCell.Offset(10, 0).Select
    vamaspagiA = "0"
    vamaspagiA = vamaspagiA + 1
    ActiveCell.Value = vfech
    ActiveCell.Offset(0, 1).Value = x
    ActiveCell.Offset(0, 2).Value = vnomb
    ActiveCell.Offset(1, 0).Select
    End If
    Next x
    
    Sheets("Coccion").Select
    Range(vpuntprin).Select
    For x = 1 To vcantamas
    If vamaspagiC <= 29 Then
    vamaspagiC = vamaspagiC + 1
    ActiveCell.Value = vfech
    ActiveCell.Offset(0, 1).Value = x
    ActiveCell.Offset(0, 2).Value = vnomb
    ActiveCell.Offset(1, 0).Select
    vpuntprin = ActiveCell.Address
    Else
    ActiveCell.Offset(10, 0).Select
      vamaspagiC = "0"
    vamaspagiC = vamaspagiC + 1
    ActiveCell.Value = vfech
    ActiveCell.Offset(0, 1).Value = x
    ActiveCell.Offset(0, 2).Value = vnomb
    ActiveCell.Offset(1, 0).Select
    vpuntprin = ActiveCell.Address
    End If
    Next x
    Sheets("Amasado").Select
    Range(vpuntamas).Select
    Else
    ActiveCell.Offset(1, -4).Select
End If
Else
ActiveCell.Offset(1, 0).Select

End If
Loop

' Comienza el trabajo de los crudos....



If vcantcrudos > 0 Then
vcanthojaauxi = vcantcrudos / 30
vcanthoja = vcanthojaauxi
        If vcanthojaauxi > vcanthoja Then
          vcanthoja = vcanthoja + 1
        End If
    Else
    MsgBox ("No hay CRUDOS para ese día! Corroborar Datos!")
    Range("A10").Select
    End
End If

Sheets("Crudo").Select
Range("B7").Value = vfech
vindipega = "A41"
Range("A9:C38").Select
Selection.ClearContents
Rows("41:500").Select
Selection.Delete Shift:=xlUp
For x = 1 To vcanthoja - 1
    Rows("1:40").Select
    Selection.Copy
    Range(vindipega).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveCell.Offset(40, 0).Select
    vindipega = ActiveCell.Address
Next x
'numerar hojas
Range("I6").Select
For x = 1 To vcanthoja
ActiveCell.Value = "Pagina " & x & " de " & vcanthoja
ActiveCell.Offset(40, 0).Select
Next x


' Completo la planilla "Crudo"
Sheets("Amasado").Select
Range("A10").Select
vpuntamas = ActiveCell.Address
vpuntprin = "A9"
vamaspagiF = "0"
Do While ActiveCell <> ""
If ActiveCell = vfech Then
ActiveCell.Offset(0, 4).Select
    If ActiveCell = "NO" Then
    vcantamas = ActiveCell.Offset(0, -2).Value
    vnomb = ActiveCell.Offset(0, -3).Value
    ActiveCell.Offset(1, -4).Select
    vpuntamas = ActiveCell.Address
    Sheets("Crudo").Select
    Range(vpuntprin).Select
    
    For x = 1 To vcantamas
    If vamaspagiF <= 29 Then
    vamaspagiF = vamaspagiF + 1
    ActiveCell.Value = vfech
    ActiveCell.Offset(0, 1).Value = x
    ActiveCell.Offset(0, 2).Value = vnomb
    ActiveCell.Offset(1, 0).Select
    Else
    vamaspagiF = "0"
    ActiveCell.Offset(10, 0).Select
    vamaspagiF = vamaspagiF + 1
    ActiveCell.Value = vfech
    ActiveCell.Offset(0, 1).Value = x
    ActiveCell.Offset(0, 2).Value = vnomb
    ActiveCell.Offset(1, 0).Select
    End If
    vpuntprin = ActiveCell.Address
    
    
    Next x
    Sheets("Amasado").Select
    Range(vpuntamas).Select
    Else
    ActiveCell.Offset(1, -4).Select
End If
Else
ActiveCell.Offset(1, 0).Select
End If
Loop

' comienza el proceso para la planilla de envasado....
If vcanttotal > 0 Then
vcanthojaauxi = vcanttotal / 30
vcanthoja = vcanthojaauxi
        If vcanthojaauxi > vcanthoja Then
          vcanthoja = vcanthoja + 1
        End If
    Else
    MsgBox ("No hay Amasados ese día! Corroborar Datos!")
    Range("A10").Select
   ' End
End If

' pego las hojas necesarias segun la cantidad de amasados en Envasado
Sheets("Envasado").Select
Range("B7").Value = vfech
vindipega = "A41"
Range("A9:C38").Select
Selection.ClearContents
Rows("41:500").Select
Selection.Delete Shift:=xlUp
For x = 1 To vcanthoja - 1
    Rows("1:40").Select
    Selection.Copy
    Range(vindipega).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveCell.Offset(40, 0).Select
    vindipega = ActiveCell.Address
Next x
'numerar hojas
Range("G6").Select
For x = 1 To vcanthoja
ActiveCell.Value = "Pagina " & x & " de " & vcanthoja
ActiveCell.Offset(40, 0).Select
Next x


Sheets("Amasado").Select
Range("A10").Select
vpuntamas = ActiveCell.Address
vpuntprin = "A9"
vamaspagiF = "0"
Do While ActiveCell <> ""
If ActiveCell = vfech Then
    vcantamas = ActiveCell.Offset(0, 2).Value
    vnomb = ActiveCell.Offset(0, 1).Value
    ActiveCell.Offset(1, 0).Select
    vpuntamas = ActiveCell.Address
    Sheets("Envasado").Select
    Range(vpuntprin).Select
    For x = 1 To vcantamas
    If vamaspagiF <= 29 Then
    vamaspagiF = vamaspagiF + 1
    ActiveCell.Value = vfech
    ActiveCell.Offset(0, 1).Value = x
    ActiveCell.Offset(0, 2).Value = vnomb
    ActiveCell.Offset(1, 0).Select
    Else
    vamaspagiF = "0"
    ActiveCell.Offset(10, 0).Select
    vamaspagiF = vamaspagiF + 1
    ActiveCell.Value = vfech
    ActiveCell.Offset(0, 1).Value = x
    ActiveCell.Offset(0, 2).Value = vnomb
    ActiveCell.Offset(1, 0).Select
    End If
    vpuntprin = ActiveCell.Address
    Next x
    
    Sheets("Amasado").Select
    Range(vpuntamas).Select
Else
ActiveCell.Offset(1, 0).Select

End If
Loop

Else
MsgBox "Esto es una demo, permiso vencido"
End If





'Sheets(Array("Fermentacion", "Abatimiento", "Coccion")).Select
'Sheets("Fermentacion").Activate
'ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True

'ActiveWorkbook.Save
End Sub

