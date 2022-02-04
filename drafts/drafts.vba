Sub center_selected()
'
' center_selected Macro
'

'
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
End Sub

Sub posicao_general_format()
'
' posicao_general_format Macro
'

'
    ActiveSheet.UsedRange.Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Cells.Select
    Selection.Columns.AutoFit
    ActiveWindow.DisplayGridlines = False
End Sub

"
Não clique dentro dessa ou qualquer outra planilha enquanto o programa roda
De outra forma, algumas ações serão realizadas fora da planilha desejada
Mas se acontecer, você ainda pode clicar novamente na pasta em que o programa está trabalhando
Nesse caso, não clique na área com as células, mas na janela
Assim você não seleciona uma célula diferente da qual o programa está trabalhando
Você pode minimizar essa pasta sem problemas. Só clique no botão de (-) da janela
"
