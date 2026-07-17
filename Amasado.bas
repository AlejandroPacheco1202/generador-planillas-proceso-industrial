Attribute VB_Name = "Módulo2"
Sub PlanillaDeAmasados()
Attribute PlanillaDeAmasados.VB_Description = "Macro grabada el 27/06/2008 por Alejandro Pacheco"
Attribute PlanillaDeAmasados.VB_ProcData.VB_Invoke_Func = "a\n14"
'
' PlanillaDeAmasados Macro
' Macro grabada el 27/06/2008 por Alejandro Pacheco
'
' OJO que la macro que imprime las planillas está en el Comand button
'
    
    Dim vfecha As Date
    Dim vnombre As String
    Dim vprevisto As Integer
    Dim vturno As String
    Dim vcodigo As String
    Dim vsolapa As String
    Dim ving1 As String
    Dim ving2 As String
    Dim ving3 As String
    Dim ving4 As String
    Dim ving5 As String
    Dim ving6 As String
    Dim ving7 As String
    Dim ving8 As String
    Dim ving9 As String
    Dim ving10 As String
    Dim ving11 As String
    Dim ving12 As String
    Dim vnumehoja As Integer
    Dim vcanthoja As Integer
    Dim vcanthojatotal As Double
    Dim vpuntamas As String
    Dim ving As String
    
    
Application.ScreenUpdating = False

    
If MsgBox("CUIDADO! Desea Imprimir las todas las planillas de amasados?", vbOKCancel, "Cuidado!!!!!!!Producción!!!!!!!!") = vbOK Then

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
    vcodigo = ActiveCell.Offset(0, 1).Value
    vsolapa = ActiveCell.Offset(0, 2).Value
    ving1 = ActiveCell.Offset(0, 3).Value
    ving2 = ActiveCell.Offset(0, 4).Value
    ving3 = ActiveCell.Offset(0, 5).Value
    ving4 = ActiveCell.Offset(0, 6).Value
    ving5 = ActiveCell.Offset(0, 7).Value
    ving6 = ActiveCell.Offset(0, 8).Value
    ving7 = ActiveCell.Offset(0, 9).Value
    ving8 = ActiveCell.Offset(0, 10).Value
    ving9 = ActiveCell.Offset(0, 11).Value
    ving10 = ActiveCell.Offset(0, 12).Value
    ving11 = ActiveCell.Offset(0, 13).Value
    ving12 = ActiveCell.Offset(0, 14).Value
    ving13 = ActiveCell.Offset(0, 15).Value
    ving14 = ActiveCell.Offset(0, 16).Value
    ving15 = ActiveCell.Offset(0, 17).Value
    vamasxhoja = ActiveCell.Offset(0, 18).Value
    ' voy a la solapa a imprimir, utilizo select case para poder diferenciarlas

    Select Case vsolapa

         Case "A"
         Sheets(vsolapa).Visible = 1
         Sheets(vsolapa).Select
        
            Range("B4").Value = "Fecha:   " & vfecha
            Range("C4").Value = "Producto: " & vnombre
            Range("D4").Value = "Turno:   " & vturno
            Range("B5").Value = ving1
            Range("C5").Value = ving2
            
            'proceso para saber que cantidad de hojas imprimir
            vcanthoja = vprevisto / vamasxhoja
            vcanthojatotal = vprevisto / vamasxhoja
                
            If vcanthojatotal > vcanthoja Then
                vcanthoja = vcanthoja + 1
            End If
         
            y = 1
            x = 1
            Do While y <= vprevisto
                Range("D3").Value = "Pag " & x & " de " & vcanthoja
                x = x + 1
                Range("A6").Value = "Amasado " & y
                y = y + 1
                Range("A12").Value = "Amasado " & y
                y = y + 1
                Range("A18").Value = "Amasado " & y
                y = y + 1
                ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
            Loop
                     
            Sheets(vsolapa).Visible = 0
        
        
        Case "B"
            Sheets(vsolapa).Visible = 1
            Sheets(vsolapa).Select
        
            Range("B4").Value = "Fecha:   " & vfecha
            Range("C4").Value = "Producto: " & vnombre
            Range("E4").Value = "Turno:   " & vturno
            Range("B5").Value = ving1
            Range("C5").Value = ving2
            Range("D5").Value = ving3
            
            'proceso para saber que cantidad de hojas imprimir
            vcanthoja = vprevisto / vamasxhoja
            vcanthojatotal = vprevisto / vamasxhoja
                
            If vcanthojatotal > vcanthoja Then
                vcanthoja = vcanthoja + 1
            End If
         
            y = 1
            x = 1
            Do While y <= vprevisto
                Range("D3").Value = "Pag " & x & " de " & vcanthoja
                x = x + 1
                Range("A6").Value = "Amasado " & y
                y = y + 1
                Range("A12").Value = "Amasado " & y
                y = y + 1
                Range("A18").Value = "Amasado " & y
                y = y + 1
                ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
            Loop
            
            Sheets(vsolapa).Visible = 0
                   
        
        Case "C"
            Sheets(vsolapa).Visible = 1
            Sheets(vsolapa).Select
        
            Range("B4").Value = "Fecha:   " & vfecha
            Range("C4").Value = "Producto: " & vnombre
            Range("E4").Value = "Turno:   " & vturno
            Range("B5").Value = ving1
            Range("C5").Value = ving2
            Range("D5").Value = ving3
            Range("E5").Value = ving4
            
            'proceso para saber que cantidad de hojas imprimir
            vcanthoja = vprevisto / vamasxhoja
            vcanthojatotal = vprevisto / vamasxhoja
                
            If vcanthojatotal > vcanthoja Then
                vcanthoja = vcanthoja + 1
            End If
         
            y = 1
            x = 1
            Do While y <= vprevisto
                Range("D3").Value = "Pag " & x & " de " & vcanthoja
                x = x + 1
                Range("A6").Value = "Amasado " & y
                y = y + 1
                Range("A12").Value = "Amasado " & y
                y = y + 1
                Range("A18").Value = "Amasado " & y
                y = y + 1
                ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
            Loop
            
            Sheets(vsolapa).Visible = 0
                   
         
        Case "D"
            Sheets(vsolapa).Visible = 1
            Sheets(vsolapa).Select
            Range("B4").Value = "Fecha:   " & vfecha
            Range("C4").Value = "Producto: " & vnombre
            Range("E4").Value = "Turno:   " & vturno
            Range("B5").Value = ving1
            Range("C5").Value = ving2
            Range("D5").Value = ving3
            Range("E5").Value = ving4
            Range("F5").Value = ving5
            
            'proceso para saber que cantidad de hojas imprimir
            vcanthoja = vprevisto / vamasxhoja
            vcanthojatotal = vprevisto / vamasxhoja
                
            If vcanthojatotal > vcanthoja Then
                vcanthoja = vcanthoja + 1
            End If
         
            y = 1
            x = 1
            Do While y <= vprevisto
                Range("D3").Value = "Pag " & x & " de " & vcanthoja
                x = x + 1
                Range("A6").Value = "Amasado " & y
                y = y + 1
                Range("A12").Value = "Amasado " & y
                y = y + 1
                Range("A18").Value = "Amasado " & y
                y = y + 1
                ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
            Loop
            
            Sheets(vsolapa).Visible = 0
                
        
        ' a partir de esta solapa, es una hoja por amasado
        Case "E"
            Sheets(vsolapa).Visible = 1
            Sheets(vsolapa).Select
            Range("B4").Value = "Fecha:   " & vfecha
            Range("C4").Value = "Producto: " & vnombre
            Range("E4").Value = "Turno:   " & vturno
            Range("B5").Value = ving1
            Range("C5").Value = ving2
            Range("D5").Value = ving3
            Range("B12").Value = ving4
            Range("C12").Value = ving5
            Range("D12").Value = ving6
            
            y = 1
            Do While y <= vprevisto
                Range("D3").Value = "Pag " & y & " de " & vprevisto
                Range("A6").Value = "Amasado " & y
                y = y + 1
                ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
            Loop
            Sheets(vsolapa).Visible = 0
                
        Case "F"
            Sheets(vsolapa).Visible = 1
            Sheets(vsolapa).Select
            Range("B4").Value = "Fecha:   " & vfecha
            Range("C4").Value = "Producto: " & vnombre
            Range("E4").Value = "Turno:   " & vturno
            Range("B5").Value = ving1
            Range("C5").Value = ving2
            Range("D5").Value = ving3
            Range("E5").Value = ving4
            Range("B12").Value = ving5
            Range("C12").Value = ving6
            Range("D12").Value = ving7
            
            y = 1
            Do While y <= vprevisto
                Range("D3").Value = "Pag " & y & " de " & vprevisto
                Range("A6").Value = "Amasado " & y
                y = y + 1
                ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
            Loop
            Sheets(vsolapa).Visible = 0
                
        
        Case "G"
            Sheets(vsolapa).Visible = 1
            Sheets(vsolapa).Select
            Range("B4").Value = "Fecha:   " & vfecha
            Range("C4").Value = "Producto: " & vnombre
            Range("E4").Value = "Turno:   " & vturno
            Range("B5").Value = ving1
            Range("C5").Value = ving2
            Range("D5").Value = ving3
            Range("E5").Value = ving4
            Range("B12").Value = ving5
            Range("C12").Value = ving6
            Range("D12").Value = ving7
            Range("E12").Value = ving8
            
            y = 1
            Do While y <= vprevisto
                Range("D3").Value = "Pag " & y & " de " & vprevisto
                Range("A6").Value = "Amasado " & y
                y = y + 1
                ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
            Loop
            Sheets(vsolapa).Visible = 0
            
            
            Case "G.O"
            Sheets(vsolapa).Visible = 1
            Sheets(vsolapa).Select
            Range("B4").Value = "Fecha:   " & vfecha
            Range("C4").Value = "Producto: " & vnombre
            Range("E4").Value = "Turno:   " & vturno
            Range("B5").Value = ving1
            Range("C5").Value = ving2
            Range("D5").Value = ving3
            Range("E5").Value = ving4
            Range("B12").Value = ving5
            Range("C12").Value = ving6
            Range("D12").Value = ving7
            Range("E12").Value = ving8
            
            y = 1
            Do While y <= vprevisto
                Range("D3").Value = "Pag " & y & " de " & vprevisto
                Range("A6").Value = "Amasado " & y
                y = y + 1
                ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
            Loop
            Sheets(vsolapa).Visible = 0
        
        Case "H"
            Sheets(vsolapa).Visible = 1
            Sheets(vsolapa).Select
            Range("B4").Value = "Fecha:   " & vfecha
            Range("C4").Value = "Producto: " & vnombre
            Range("E4").Value = "Turno:   " & vturno
            Range("B5").Value = ving1
            Range("C5").Value = ving2
            Range("D5").Value = ving3
            Range("E5").Value = ving4
            Range("F5").Value = ving5
            Range("B12").Value = ving6
            Range("C12").Value = ving7
            Range("D12").Value = ving8
            Range("E12").Value = ving9
            
            y = 1
            Do While y <= vprevisto
                Range("D3").Value = "Pag " & y & " de " & vprevisto
                Range("A6").Value = "Amasado " & y
                y = y + 1
                ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
            Loop
            Sheets(vsolapa).Visible = 0
        
        
        Case "I"
            Sheets(vsolapa).Visible = 1
            Sheets(vsolapa).Select
            Range("B4").Value = "Fecha:   " & vfecha
            Range("C4").Value = "Producto: " & vnombre
            Range("E4").Value = "Turno:   " & vturno
            Range("B5").Value = ving1
            Range("C5").Value = ving2
            Range("D5").Value = ving3
            Range("E5").Value = ving4
            Range("F5").Value = ving5
            Range("B12").Value = ving6
            Range("C12").Value = ving7
            Range("D12").Value = ving8
            Range("E12").Value = ving9
            Range("F12").Value = ving10
            Range("G12").Value = ving11
            Range("G15").Value = ving12
            
            y = 1
            Do While y <= vprevisto
                Range("D3").Value = "Pag " & y & " de " & vprevisto
                Range("A6").Value = "Amasado " & y
                y = y + 1
                ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
            Loop
            Sheets(vsolapa).Visible = 0
        
        Case "J"
            Sheets(vsolapa).Visible = 1
            Sheets(vsolapa).Select
            Range("B4").Value = "Fecha:   " & vfecha
            Range("C4").Value = "Producto: " & vnombre
            Range("E4").Value = "Turno:   " & vturno
            Range("B5").Value = ving1
            Range("C5").Value = ving2
            Range("D5").Value = ving3
            Range("E5").Value = ving4
            Range("F5").Value = ving5
            Range("B12").Value = ving6
            Range("C12").Value = ving7
            Range("D12").Value = ving8
            Range("E12").Value = ving9
            Range("F12").Value = ving10
            Range("B19").Value = ving11
            Range("C19").Value = ving12
            Range("D19").Value = ving13
            Range("E19").Value = ving14
            Range("F19").Value = ving15
            
            y = 1
            Do While y <= vprevisto
                Range("D3").Value = "Pag " & y & " de " & vprevisto
                Range("A6").Value = "Amasado " & y
                y = y + 1
                ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
            Loop
            Sheets(vsolapa).Visible = 0
        
        
        
    End Select

Sheets("Amasado").Select
Range(vpuntamas).Select
If MsgBox("Continua con la impresión", vbOKCancel, "Cuidado!!!!!!!Producción!!!!!!!!") = vbCancel Then
    End
End If


Loop

Sheets("ART").Visible = 0

ActiveWorkbook.Save

End If



End Sub
