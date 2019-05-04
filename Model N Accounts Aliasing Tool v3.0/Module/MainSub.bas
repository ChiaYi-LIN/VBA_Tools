Attribute VB_Name = "MainSub"
Sub Main()

Call OptimizeCode_Begin

'清除"Data Cleaner"頁面的所有內容
    
With Worksheets("Data Cleaner").Cells
    .Clear
    .Font.FontStyle = "Calibri"
End With

'進入"Data Cleaner"工作表
ActiveWorkbook.Sheets("Data cleaner").Activate
Dim sourceColNum, sourceRowNum, i As Long
'variable declaration
Dim pctDone As Long

'計算欄數
sourceColNum = Worksheets("Source").Cells(1, 2).End(xlToRight).Column
'Worksheets("Source").Range(Worksheets("Source").Cells(1, 1), Worksheets("Source").Cells(1, 1).End(xlToRight)).Count
'計算列數
sourceRowNum = Worksheets("Source").Range("B1").CurrentRegion.Rows.Count

Worksheets("Data Cleaner").Range(Worksheets("Data Cleaner").Cells(1, 2), Worksheets("Data Cleaner").Cells(1, sourceColNum)).Value = Worksheets("Source").Range(Worksheets("Source").Cells(1, 2), Worksheets("Source").Cells(1, sourceColNum)).Value

For i = 2 To sourceRowNum
    Worksheets("Data Cleaner").Cells(i, 2).Value = CleanCompanyName(Worksheets("Source").Cells(i, 2).Value)
    Worksheets("Data Cleaner").Cells(i, 3).Value = CleanCountryName(Worksheets("Source").Cells(i, 3).Value)
    Worksheets("Data Cleaner").Range(Worksheets("Data Cleaner").Cells(i, 4), Worksheets("Data Cleaner").Cells(i, sourceColNum)).Value = Worksheets("Source").Range(Worksheets("Source").Cells(i, 4), Worksheets("Source").Cells(i, sourceColNum)).Value
           
    'progress bar
    pctDone = i / sourceRowNum * 100
    With Progress
        .FrameProgress.Caption = "Complete: " & pctDone & "%"
        .LabelProgress.Width = pctDone * 2.4
        DoEvents
    End With
Next i

Worksheets("Data Cleaner").Cells.Select
With Selection
    .Font.Name = "Calibri"
End With

ActiveWorkbook.Sheets("Source").Activate
Worksheets("Source").Range("A1").Select
ActiveWorkbook.Sheets("Data Cleaner").Activate
Worksheets("Data Cleaner").Range("A1").Select

Unload Progress

Call OptimizeCode_End

End Sub



