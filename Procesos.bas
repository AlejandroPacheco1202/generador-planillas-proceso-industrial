Attribute VB_Name = "Módulo4"
Sub PlanillasDeProceso()
Attribute PlanillasDeProceso.VB_Description = "Macro grabada el 23/07/2008 por Alejandro"
Attribute PlanillasDeProceso.VB_ProcData.VB_Invoke_Func = " \n14"
'
' PlanillasIndividuales Macro
' Macro grabada el 18/03/2019 por Alejandro

    Dim vfecha As Date
    Dim vnombre As String
    Dim vprevisto As Integer
    Dim vturno As String
    Dim vcodigo As String
    Dim vsolapa As String
    Dim vcantamastota As Integer
    Dim vnumehoja As Integer
    Dim vcanthoja As Double
    Dim vcanthojatotal As Double
    Dim vpuntamas As String
    Dim ving As String
    
    
Application.ScreenUpdating = False

    
If MsgBox("CUIDADO! Desea Imprimir las todas las planillas de amasados?", vbOKCancel, "Cuidado!!!!!!!Producción!!!!!!!!") = vbOK Then

    Range("C10").Select
   
    'voy a contar la cantidad de amasados
    Do While ActiveCell <> ""
        vcantamastota = vcantamastota + ActiveCell
        ActiveCell.Offset(1, 0).Select
    Loop
    If vcantamastota > 30 Then
        vcanthoja = 2
        Else
        vcanthoja = 1
    End If
    

 
    'lleno los amasados
    Sheets("Amasado").Select
    vfecha = Range("A10").Value
    vturno = Range("D10").Value
        
  'lleno encabezados (dejo provisoriamente en 1 la cantidad de hojas)
  x = 1
  Sheets("Fermentacion").Select
  Range("A9:H39").Value = ""
  Range("A7").Value = "Fecha: " & vfecha
  Range("E7").Value = vturno
  Range("F6").Value = "Pag " & x & " de " & vcanthoja
  Sheets("Coccion").Select
  Range("A9:L39").Value = ""
  Range("A7").Value = "Fecha: " & vfecha
  Range("E7").Value = vturno
  Range("I6").Value = "Pag " & x & " de " & vcanthoja
  Sheets("Abatimiento").Select
  Range("A9:L39").Value = ""
  Range("A7").Value = "Fecha: " & vfecha
  Range("E7").Value = vturno
  Range("I6").Value = "Pag " & x & " de " & vcanthoja
  Sheets("Envasado").Select
  Range("A9:I39").Value = ""
  Range("A7").Value = "Fecha: " & vfecha
  Range("E7").Value = vturno
  Range("F6").Value = "Pag " & x & " de " & vcanthoja
     
    
    Sheets("Amasado").Select
    Range("A10").Select
    vpuntamas = ActiveCell.Address
    
    Do While ActiveCell <> ""
        vfecha = ActiveCell
        vnombre = ActiveCell.Offset(0, 1).Value
        vprevisto = ActiveCell.Offset(0, 2).Value
        vturno = ActiveCell.Offset(0, 3).Value
        ActiveCell.Offset(1, 0).Select
        vpuntamas = ActiveCell.Address
        
        'me muevo a la base ART
        Sheets("ART").Visible = 1
        Sheets("ART").Select
        Range("A2").Select
        Do While ActiveCell <> vnombre
            ActiveCell.Offset(1, 0).Select
        Loop
    
        vproceso = ActiveCell.Offset(0, 19).Value
        
        Sheets("Fermentacion").Select
        Range("A9").Select
        Do While ActiveCell <> ""
            ActiveCell.Offset(1, 0).Select
        Loop
        For x = 1 To vprevisto
            ActiveCell = x
            ActiveCell.Offset(0, 1).Value = vnombre
            ActiveCell.Offset(1, 0).Select
        Next x
        
        Sheets("Coccion").Select
        Range("A9").Select
        Do While ActiveCell <> ""
            ActiveCell.Offset(1, 0).Select
        Loop
        For x = 1 To vprevisto
            ActiveCell = x
            ActiveCell.Offset(0, 1).Value = vnombre
            ActiveCell.Offset(1, 0).Select
        Next x
        

        Sheets("Abatimiento").Select
        Range("A9").Select
        Do While ActiveCell <> ""
            ActiveCell.Offset(1, 0).Select
        Loop
        For x = 1 To vprevisto
            ActiveCell = x
            ActiveCell.Offset(0, 1).Value = vnombre
            ActiveCell.Offset(1, 0).Select
        Next x
        
    
        Sheets("Envasado").Select
        Range("A9").Select
        Do While ActiveCell <> ""
            ActiveCell.Offset(1, 0).Select
        Loop
        For x = 1 To vprevisto
            ActiveCell = x
            ActiveCell.Offset(0, 1).Value = vnombre
            ActiveCell.Offset(1, 0).Select
        Next x
        
    
    
    Sheets("Amasado").Select
    Range(vpuntamas).Select
    Loop
    
    
    
Sheets("ART").Visible = 0
ActiveWorkbook.Save

Sheets("Fermentacion").Select
Range("A9").Select

End If




End Sub
