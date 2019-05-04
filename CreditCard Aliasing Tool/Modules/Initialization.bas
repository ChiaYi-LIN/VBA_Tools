Attribute VB_Name = "Initialization"
Sub WorksheetsInitialization()
    With ActiveWorkbook.Styles("Normal").Font
    .Name = "Calibri"
    .Size = 11
    End With
    Worksheets("Source").Cells(1, 1).Value = "Input Data From Here >>>"
    Worksheets("Fuzzy Lookup").Cells(1, 1).Value = "Select Before Matching >>>"


End Sub
