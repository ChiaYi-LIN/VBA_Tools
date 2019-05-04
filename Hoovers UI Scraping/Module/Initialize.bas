Attribute VB_Name = "Initialize"
Sub Initialization()
    With Worksheets("README")
        .Cells.Interior.ColorIndex = xlnofill
        .Columns.ColumnWidth = 25
        .Rows.RowHeight = 15
    End With
    
    With Worksheets("README").Cells.Font
        .Name = "Calibri"
        .Size = 11
        .Bold = False
        .Italic = False
    End With
    
    With Worksheets("DUNS")
        .Cells.Interior.ColorIndex = xlnofill
        .Columns.ColumnWidth = 25
        .Range("B:D").Interior.Color = RGB(217, 225, 242)
        .Columns(5).Interior.Color = RGB(255, 255, 0)
        .Rows.RowHeight = 15
    End With
    
    With Worksheets("DUNS").Cells.Font
        .Name = "Calibri"
        .Size = 11
        .Bold = False
        .Italic = False
    End With
    
    With Worksheets("DUNS").Rows(1)
        .Font.Bold = True
    End With
End Sub

