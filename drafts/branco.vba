Sub Bloco()
'
' Bloco2 Macro
'
' Atalho do teclado: Opção+Cmd+v

    Range("E13:E1605").Select
    Selection.SpecialCells(xlCellTypeBlanks).EntireRow.Select
    Selection.EntireRow.Delete Shift:=xlUp
    
    Range("B8").Select
    
    Call branco4
    
    Call branco3
    

End Sub



Sub branco4()


Do While ActiveCell.Offset(1, 0).Value <> "TOTAL EM VOLUMES"

If ActiveCell.Value = "" Then


If ActiveCell.Offset(1, 0).Value <> "" Then

ActiveCell.Offset(1, 0).Activate

End If

If ActiveCell.Value = "" Then

ActiveCell.Offset(1, 3).Value = ""

End If

End If
ActiveCell.Offset(1, 0).Activate

Loop

Range("B9").Select

If ActiveCell.Value = "" Then
Selection.EntireRow.Delete Shift:=xlUp

End If



End Sub

Sub branco3()

    Range("E13:E1605").Select
    Selection.SpecialCells(xlCellTypeBlanks).EntireRow.Select
    Selection.EntireRow.Delete Shift:=xlUp


End Sub
Sub Bloco2()
'
' Bloco Macro
'
' Atalho do teclado: Opção+Cmd+z
'

'
    Columns("E:E").Select
    Selection.Replace What:="-", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False
'
    Range("E13:E1605").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.EntireRow.Delete

End Sub

Sub IMPORTACAO()

'deixa a aba importacao visivel e seleciona ela
    Sheets("Importação").Visible = True
    Sheets("Importação").Activate
    

'contador: verifica numero de celulas preenchidas na coluna e _
seleciona da A2 ate a ultima preenchida na coluna K _
limpas os dados da aba importacao
    Final = WorksheetFunction.CountA(Sheets("Importação").Range("g:g"))
    Range("A2:l" & Final + 1).Select
    Selection.EntireRow.Delete
    
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
    Sheets("Bloco de pedidos").Select
'contador verifica quantas linhas na coluna A esta preenchida _
seleciona da celula A12 ate a ultima preenchida e copia para aba importacao
    linha_final = Range("a12").End(xlDown).Row
    
    Range("A13:E" & linha_final).Select
    Selection.Copy

    Sheets("Importação").Select
    
    Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
'formulas para deixar no padrao necessario para importar o pedido
    Range("f2").FormulaR1C1 = "=RC[-4]&RC[-3]&RC[-2]"
    Range("g2").FormulaR1C1 = "=IF(RC[-2]>0,RC[-2],"""")"
    Range("h2").FormulaR1C1 = "=IF(RC[-3]>0,R1C12,"""")"
    Range("i2").FormulaR1C1 = "=IF(RC5>0,0,"""")"
    Range("j2").FormulaR1C1 = "=IF(RC5>0,0,"""")"
    Range("K2").FormulaR1C1 = "=IF(RC5>0,0,"""")"


    
'seleciona o intervalo das formulas e replica para as outras _
linhas ate a ultima linha preenchida
    Range("F2:l2").Select
    linha = WorksheetFunction.CountA(Sheets("Importação").Range("b:b"))
    Selection.AutoFill Destination:=Range("F2:l" & linha), Type:=xlFillDefault
    Application.Calculation = xlAutomatic
    

    Range("F1" & ":" & "l" & linha).Select
    Selection.Copy

'Cria um novo arquivo e passa os dados
    Dim ws As Workbook
    Set ws = Application.Workbooks.Add
    
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Application.CutCopyMode = False
    
   linha = WorksheetFunction.CountA(Range("a:a"))
   Range("D2" & ":" & "D" & linha).Select

    Selection.NumberFormat = "00"

End Sub



