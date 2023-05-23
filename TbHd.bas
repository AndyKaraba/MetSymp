Sub TbHdr()
'
' Table header formatting macro
'

'	

    Dim TbHdr As String
    Dim TbHdrR As Integer
    
    TbHdr = ActiveCell.Value
   ' Debug.Print TbHdr
    TbHdrR = ActiveCell.Row
   ' Debug.Print TbHdrR
    ActiveCell.FormulaR1C1 = ""
    
    Range(Cells(TbHdrR, 2), Cells(TbHdrR, 7)).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range(Cells(TbHdrR, 2), Cells(TbHdrR, 7)).Select
    Selection.Merge
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    ActiveCell.FormulaR1C1 = TbHdr
End Sub