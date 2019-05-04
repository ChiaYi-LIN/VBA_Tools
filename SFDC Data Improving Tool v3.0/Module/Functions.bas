Attribute VB_Name = "Functions"
Sub DUNSOutputIntoCSV()
Dim objWS As Variant
Dim strDesktopPath As String

Application.DisplayAlerts = False

ThisWorkbook.Sheets("DUNS.csv").Visible = True
ThisWorkbook.Sheets("DUNS.csv").Copy

Set objWS = CreateObject("WScript.Shell")
strDesktopPath = objWS.SpecialFolders("Desktop")
FILE_NAME = "DUNS" & "_" & Format(Date, "yyyymmdd") & "_" & Format(Time, "hhmmss")
ActiveWorkbook.SaveAs Filename:=strDesktopPath & "\" & FILE_NAME, FileFormat:=xlCSV, CreateBackup:=False
ActiveWorkbook.Close

ThisWorkbook.Sheets("DUNS.csv").Visible = False
ThisWorkbook.Sheets("DUNS.csv").Cells.Clear
Application.DisplayAlerts = True

End Sub

Sub SFDCOutputIntoCSV()
Dim objWS As Variant
Dim strDesktopPath As String

Application.DisplayAlerts = False

ThisWorkbook.Sheets("Output_csv").Visible = True
ThisWorkbook.Sheets("Output_csv").Copy

Set objWS = CreateObject("WScript.Shell")
strDesktopPath = objWS.SpecialFolders("Desktop")
FILE_NAME = "Poor_Match" & "_" & Format(Date, "yyyymmdd") & "_" & Format(Time, "hhmmss")
ActiveWorkbook.SaveAs Filename:=strDesktopPath & "\" & FILE_NAME, FileFormat:=xlCSV, CreateBackup:=False
ActiveWorkbook.Close

ThisWorkbook.Sheets("Output_csv").Visible = False
ThisWorkbook.Sheets("Output_csv").Cells.Clear
Application.DisplayAlerts = True

End Sub

Sub HooversOutputIntoExcel()
Dim objWS As Variant
Dim strDesktopPath As String

Application.DisplayAlerts = False

ThisWorkbook.Sheets("Output_csv").Visible = True
ThisWorkbook.Sheets("Output_csv").Copy

Set objWS = CreateObject("WScript.Shell")
strDesktopPath = objWS.SpecialFolders("Desktop")
FILE_NAME = "Mass_Update" & "_" & Format(Date, "yyyymmdd") & "_" & Format(Time, "hhmmss")
ActiveWorkbook.SaveAs Filename:=strDesktopPath & "\" & FILE_NAME, FileFormat:=xlWorkbookDefault, CreateBackup:=False
ActiveWorkbook.Close

ThisWorkbook.Sheets("Output_csv").Visible = False
ThisWorkbook.Sheets("Output_csv").Cells.Clear
Application.DisplayAlerts = True

End Sub

Sub ExportResults()
Dim objWS As Variant
Dim strDesktopPath As String

Worksheets("Matching").Activate
Worksheets("Matching").Cells(1, 2).CurrentRegion.Copy Destination:=Worksheets("Output_csv").Cells(1, 1)
Application.DisplayAlerts = False

ThisWorkbook.Sheets("Output_csv").Visible = True
ThisWorkbook.Sheets("Output_csv").Copy

Set objWS = CreateObject("WScript.Shell")
strDesktopPath = objWS.SpecialFolders("Desktop")
FILE_NAME = "Matching_Results" & "_" & Format(Date, "yyyymmdd") & "_" & Format(Time, "hhmmss")
ActiveWorkbook.SaveAs Filename:=strDesktopPath & "\" & FILE_NAME, FileFormat:=xlWorkbookDefault, CreateBackup:=False
ActiveWorkbook.Close

ThisWorkbook.Sheets("Output_csv").Visible = False
ThisWorkbook.Sheets("Output_csv").Cells.Clear
Application.DisplayAlerts = True

MsgBox Chr(34) & FILE_NAME & ".xlsx" & Chr(34) & " has been saved to desktop."
FILE_NAME = ""
Worksheets("Matching").Cells(1, 1).Select
End Sub

Sub CSV_Import(ByVal LoadToSheet As String, ByVal TabelName As String)
Dim ws As Worksheet, strFile As Variant
Application.ScreenUpdating = False
On Error Resume Next
Set ws = ActiveWorkbook.Worksheets("DataLoader_csv")
Worksheets("DataLoader_csv").Cells.Clear
strFile = Application.GetOpenFilename("Text Files (*.csv),*.csv", , "Select file")
'strFile = Application.GetOpenFilename("Text Files (*.csv;*.xlsx;*.xls),*.csv;*.xlsx;*.xls", , "Select file")

On Error GoTo 0
If strFile = False Then Exit Sub
On Error Resume Next
Worksheets(LoadToSheet).ListObjects(TabelName).Range.Select
Selection.Delete
Worksheets(LoadToSheet).Cells.Clear
Worksheets(LoadToSheet).Cells.Clear
Worksheets("DataLoader_csv").Cells.Clear

With ws.QueryTables.Add(Connection:="TEXT;" & strFile, Destination:=ws.Range("A1"))
     .TextFilePlatform = 65001
     .TextFileParseType = xlDelimited
     .TextFileCommaDelimiter = True
     .TextFileColumnDataTypes = Array(2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, _
     2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2)
     .Refresh
End With

If LoadToSheet = "Hoovers" Then
    Call StandardHead
End If
Application.CutCopyMode = False
Worksheets("DataLoader_csv").Cells(1, 1).CurrentRegion.Copy
Worksheets(LoadToSheet).Cells(1, 2).PasteSpecial xlPasteValues
Application.CutCopyMode = False

Worksheets(LoadToSheet).Cells(1, 2).CurrentRegion.Select
Worksheets(LoadToSheet).ListObjects.Add(xlSrcRange, Selection, , xlYes).Name = TabelName
On Error GoTo 0
Call Initialization
Worksheets("DataLoader_csv").Cells.Clear
Worksheets("DataLoader_csv").Activate
Worksheets("DataLoader_csv").Cells(1, 1).Select
Worksheets(LoadToSheet).Activate
Worksheets(LoadToSheet).Cells(1, 1).Select

Application.ScreenUpdating = True
End Sub

Sub StandardHead()
Dim i As Long
Dim nextLoop As Boolean
Dim toDelete, eachToDelete As Variant
toDelete = Array("Primary County", "Phone Number", "Toll-Free Number", "FAX Number", "Mailing Address 1", "Mailing Address 2", _
"Mailing City", "Mailing County", "Mailing State", "Mailing Zip", "Mailing Zip Extension", "Mailing Country", "Latitude", _
"Longitude", "Tradestyle", "Phone", "Fax", "Revenue USD", "Pre Tax Profit USD", "Assets USD", "Liabilities USD", _
"Employees Single Site", "Employees All Sites", "Business Description", "Ownership Type", "Entity Type", "Ticker", _
"Parent Company", "Parent CountryRegion", "Ultimate Parent Company", "Ultimate Parent CountryRegion", "DB Hoovers Industry", _
"US SIC 1987 Code", "US SIC 1987 Description", "NAICS 2012 Code", "NAICS 2012 Description", "UK SIC 2007 Code", _
"UK SIC 2007 Description", "ISIC Rev 4 Code", "ISIC Rev 4 Description", "NACE Rev 2 Code", "NACE Rev 2 Description", _
"ANZSIC 2006 Code", "ANZSIC 2006 Description", "TPS Flag", "Key ID", "Source", "Direct Marketing Status", "Street Line1", _
"Street Line 2", "Street Line 3", "Customer: ID", "")

i = 1
nextLoop = False

'Do While Worksheets("DataLoader_csv").Cells(1, i).Value <> ""
Do While i <= Worksheets("DataLoader_csv").UsedRange.Columns.Count
nextLoop = False

If nextLoop = False Then
For Each eachToDelete In toDelete
    If Worksheets("DataLoader_csv").Cells(1, i).Value = eachToDelete Then
        Worksheets("DataLoader_csv").Cells(1, i).EntireColumn.Delete
        i = i - 1
        nextLoop = True
        Exit For
    End If
Next eachToDelete
End If

If nextLoop = False Then
If Worksheets("DataLoader_csv").Cells(1, i).Value = "Company Name" Then
    Worksheets("DataLoader_csv").Cells(1, i).Value = "Legal Name"
    nextLoop = True
End If
End If

If nextLoop = False Then
If Worksheets("DataLoader_csv").Cells(1, i).Value = "Address Line 1" Then
    Worksheets("DataLoader_csv").Cells(1, i).Value = "Street Line1"
    nextLoop = True
End If
End If

If nextLoop = False Then
If Worksheets("DataLoader_csv").Cells(1, i).Value = "Primary Address 1" Then
    Worksheets("DataLoader_csv").Cells(1, i).Value = "Street Line1"
    nextLoop = True
End If
End If

If nextLoop = False Then
If Worksheets("DataLoader_csv").Cells(1, i).Value = "Address Line 2" Then
    Worksheets("DataLoader_csv").Cells(1, i).Value = "Street Line 2"
    nextLoop = True
End If
End If

If nextLoop = False Then
If Worksheets("DataLoader_csv").Cells(1, i).Value = "Primary Address 2" Then
    Worksheets("DataLoader_csv").Cells(1, i).Value = "Street Line 2"
    nextLoop = True
End If
End If

If nextLoop = False Then
If Worksheets("DataLoader_csv").Cells(1, i).Value = "Address Line 3" Then
    Worksheets("DataLoader_csv").Cells(1, i).Value = "Street Line 3"
    nextLoop = True
End If
End If

If nextLoop = False Then
If Worksheets("DataLoader_csv").Cells(1, i).Value = "Primary City" Then
    Worksheets("DataLoader_csv").Cells(1, i).Value = "City"
    nextLoop = True
End If
End If

If nextLoop = False Then
If Worksheets("DataLoader_csv").Cells(1, i).Value = "State Or Province" Then
    Worksheets("DataLoader_csv").Cells(1, i).Value = "State/Province"
    nextLoop = True
End If
End If

If nextLoop = False Then
If Worksheets("DataLoader_csv").Cells(1, i).Value = "Postal Code" Then
    Worksheets("DataLoader_csv").Cells(1, i).Value = "ZIP"
    nextLoop = True
End If
End If

If nextLoop = False Then
If Worksheets("DataLoader_csv").Cells(1, i).Value = "Zip Code" Then
    Worksheets("DataLoader_csv").Cells(1, i).Value = "ZIP"
    nextLoop = True
End If
End If

If nextLoop = False Then
If Worksheets("DataLoader_csv").Cells(1, i).Value = "Primary Zip Extension" Then
    Worksheets("DataLoader_csv").Cells(1, i).Value = "ZIP Extension"
    nextLoop = True
End If
End If

If nextLoop = False Then
If Worksheets("DataLoader_csv").Cells(1, i).Value = "CountryRegion" Then
    Worksheets("DataLoader_csv").Cells(1, i).Value = "Country"
    nextLoop = True
End If
End If

If nextLoop = False Then
If Worksheets("DataLoader_csv").Cells(1, i).Value = "Primary Country" Then
    Worksheets("DataLoader_csv").Cells(1, i).Value = "Country"
    nextLoop = True
End If
End If

If nextLoop = False Then
If Worksheets("DataLoader_csv").Cells(1, i).Value = "Web Address" Then
    Worksheets("DataLoader_csv").Cells(1, i).Value = "Website"
    nextLoop = True
End If
End If

If nextLoop = False Then
If Worksheets("DataLoader_csv").Cells(1, i).Value = "URL" Then
    Worksheets("DataLoader_csv").Cells(1, i).Value = "Website"
    nextLoop = True
End If
End If

If nextLoop = False Then
If Worksheets("DataLoader_csv").Cells(1, i).Value = "D-U-N-S Number" Then
    Worksheets("DataLoader_csv").Cells(1, i).Value = "DUNS"
    nextLoop = True
End If
End If

i = i + 1
Loop

Worksheets("DataLoader_csv").Cells(1, i).Value = "DUNS verified"
End Sub

Sub DataTransmitOld()
    Dim i, j, k, SfdcRow, SFDCColumn, HooversRow, HooversColumn, pctDone, countRowNum As Long
    
    Worksheets("Matching").Cells.Clear
    Worksheets("Matching").Cells.Clear
    
    Application.CutCopyMode = False
    Worksheets("Salesforce Customers").Activate
    Set SFDC_ALL_DATA = Worksheets("Salesforce Customers").Cells(1, 2).CurrentRegion
    SFDC_ALL_DATA.Copy
    Worksheets("Matching").Activate
    Worksheets("Matching").Cells(1, 2).PasteSpecial xlPasteAll
    
    'SFDC_ALL_DATA.Copy Destination:=Worksheets("Matching").Cells(1, 2)
    
    Application.CutCopyMode = False
    Worksheets("Matching").Columns(SFDC_DUNS.Column).Cut
    Worksheets("Matching").Columns(SFDC_ALL_DATA.Columns.Count + 2).Insert Shift:=xlToRight
    Application.CutCopyMode = False
    
    countRowNum = Worksheets("Matching").Cells(1, 2).CurrentRegion.Rows.Count
    
    With Range(Worksheets("Matching").Cells(1, SFDC_ALL_DATA.Columns.Count + 1), Worksheets("Matching").Cells(countRowNum, SFDC_ALL_DATA.Columns.Count + 1))
        .Value = .Value
    End With
    
    
    
    Worksheets("Hoovers").Activate
    Set HOOVERS_ALL_DATA = Worksheets("Hoovers").Cells(1, 2).CurrentRegion
    
    Worksheets("Hoovers").Columns(HOOVERS_DUNS.Column).Cut
    On Error Resume Next
    Worksheets("Hoovers").Columns(2).Insert Shift:=xlToRight
    On Error GoTo 0
    Application.CutCopyMode = False
    
    Set HOOVERS_ALL_DATA = Worksheets("Hoovers").Cells(1, 2).CurrentRegion
    
    SfdcRow = SFDC_ALL_DATA.Rows.Count
    SFDCColumn = SFDC_ALL_DATA.Columns.Count
    HooversRow = HOOVERS_ALL_DATA.Rows.Count
    HooversColumn = HOOVERS_ALL_DATA.Columns.Count
    
    GoTo Endall
    i = 2
    k = 2
        For j = (1 + SFDCColumn + 1) To (SFDCColumn + HooversColumn)
            Worksheets("Matching").Activate
            Worksheets("Matching").Cells(1, j).Value = Worksheets("Hoovers").Cells(1, j - SFDCColumn + 1).Value
            Worksheets("Matching").Cells(i, j).Activate
            
            If IF_HAVE_DUNS = True Then
                ActiveCell.FormulaR1C1 = _
                "=IFERROR(VLOOKUP([@" & SFDC_DUNS.Value & "],Hoovers_data[#All]," & k & ",FALSE), " & Chr(34) & Chr(34) & ")"
                k = k + 1
            Else
                ActiveCell.FormulaR1C1 = _
                "=IFERROR(VLOOKUP([@[" & SFDC_DUNS.Value & "]],Hoovers_data[#All]," & k & ",FALSE), " & Chr(34) & Chr(34) & ")"
                k = k + 1
            End If
            
            pctDone = (j - (1 + SFDCColumn + 1)) * 100 / ((SFDCColumn + HooversColumn) - (1 + SFDCColumn + 1))
            
            With Progress
                .theFrameProgress.Caption = "Merging SFDC Data Table & Hoovers Data Table. Complete: " & pctDone & "%"
                .theLabelProgress.Width = pctDone * 2.4
                DoEvents
            End With
        Next j
        
    Worksheets("Matching").Range(Cells(1, 2), Cells(SfdcRow, SFDCColumn)).Interior.Color = RGB(221, 235, 247)
    Worksheets("Matching").Range(Cells(1, SFDCColumn + 1), Cells(SfdcRow, SFDCColumn + 1)).Interior.Color = RGB(255, 255, 0)
    Worksheets("Matching").Range(Cells(1, SFDCColumn + 2), Cells(SfdcRow, SFDCColumn + HooversColumn)).Interior.Color = _
    RGB(226, 239, 218)
    
Endall:
    Unload Progress
    'Call ZeroToBlank
End Sub

Sub ShowScreen()
    Application.ScreenUpdating = True
End Sub

Sub CellsToText()
Dim eachWorksheet As Worksheet
For Each eachWorksheet In ThisWorkbook.Worksheets
    eachWorksheet.Cells.NumberFormat = "@"
Next eachWorksheet
End Sub

Sub ZeroToBlank()
    Dim countRowNum, allColumn As Long
    Dim eachColumn As Range
    'Dim replaceText As String
    
    Worksheets("Matching").Activate
    countRowNum = Worksheets("Matching").Cells(1, 2).CurrentRegion.Rows.Count
    allColumn = SFDC_ALL_DATA.Columns.Count + HOOVERS_ALL_DATA.Columns.Count
    
    'Worksheets("Matching").UsedRange.Value = Worksheets("Matching").UsedRange.Value
    'Range(Worksheets("Matching").Cells(1, 2), Worksheets("Matching").Cells(countRowNum, allColumn)).Replace _
    What:="0", Replacement:="", LookAt:=xlWhole
    
    For Each eachColumn In Worksheets("Matching").UsedRange.Columns
        If InStr(1, Worksheets("Matching").Cells(1, eachColumn.Column).Value, "ZIP") > 0 Then
            With Range(Worksheets("Matching").Cells(1, eachColumn.Column), Worksheets("Matching").Cells(countRowNum, eachColumn.Column))
                .NumberFormat = "@"
                .Value = .Value
            End With
        End If
    Next eachColumn
    
    With Worksheets("Matching").UsedRange
        .Value = .Value
    End With
    
    Range(Worksheets("Matching").Cells(1, 2), Worksheets("Matching").Cells(countRowNum, allColumn)).Replace _
    What:="0", Replacement:="", LookAt:=xlWhole

End Sub

Sub SelectFirstCell()
    Worksheets("Salesforce Customers").Activate
    Worksheets("Salesforce Customers").Cells(1, 1).Select
    Worksheets("Hoovers").Activate
    Worksheets("Hoovers").Cells(1, 1).Select
    Worksheets("Matching").Activate
    Worksheets("Matching").Cells(1, 1).Select
End Sub

Sub DataTransmit()
    Dim i, j, k, SfdcRow, SFDCColumn, HooversRow, HooversColumn, pctDone, countRowNum As Long
    Dim checkNumber As Boolean
    
    Worksheets("Matching").Cells.Clear
    Worksheets("Matching").Cells.Clear
    
    Application.CutCopyMode = False
    Worksheets("Salesforce Customers").Activate
    Set SFDC_ALL_DATA = Worksheets("Salesforce Customers").Cells(1, 2).CurrentRegion
    Worksheets("Salesforce Customers").Columns(SFDC_DUNS.Column).Cut
    Worksheets("Salesforce Customers").Columns(SFDC_ALL_DATA.Columns.Count + 2).Insert Shift:=xlToRight
    
    Application.CutCopyMode = False
    Set SFDC_ALL_DATA = Worksheets("Salesforce Customers").Cells(1, 2).CurrentRegion
    SFDC_ALL_DATA.Copy Destination:=Worksheets("Matching").Cells(1, 2)
    
    Application.CutCopyMode = False
    'Worksheets("Matching").Columns(SFDC_DUNS.Column).Cut
    'Worksheets("Matching").Columns(SFDC_ALL_DATA.Columns.Count + 2).Insert Shift:=xlToRight
    'Application.CutCopyMode = False
    
    Worksheets("Matching").Activate
    countRowNum = Worksheets("Matching").Cells(1, 2).CurrentRegion.Rows.Count
    With Range(Worksheets("Matching").Cells(1, SFDC_ALL_DATA.Columns.Count + 1), Worksheets("Matching").Cells(countRowNum, SFDC_ALL_DATA.Columns.Count + 1))
        .Value = .Value
    End With
    
    Worksheets("Hoovers").Activate
    Worksheets("Hoovers").Columns(HOOVERS_DUNS.Column).Cut
    On Error Resume Next
    Worksheets("Hoovers").Columns(2).Insert Shift:=xlToRight
    On Error GoTo 0
    Application.CutCopyMode = False
    countRowNum = Worksheets("Hoovers").Cells(1, 2).CurrentRegion.Rows.Count
    With Range(Worksheets("Hoovers").Cells(1, 2), Worksheets("Hoovers").Cells(countRowNum, 2))
        .Value = .Value
    End With
    
    Set HOOVERS_ALL_DATA = Worksheets("Hoovers").Cells(1, 2).CurrentRegion
    SfdcRow = SFDC_ALL_DATA.Rows.Count
    SFDCColumn = SFDC_ALL_DATA.Columns.Count
    HooversRow = HOOVERS_ALL_DATA.Rows.Count
    HooversColumn = HOOVERS_ALL_DATA.Columns.Count
    
    On Error Resume Next
        checkNumber = False
        checkNumber = Application.WorksheetFunction.IsNumber(Worksheets("Matching").Cells(2, SFDC_DUNS.Column).Value * 1)
    On Error GoTo 0
    
    i = 2
    k = 2
        For j = (1 + SFDCColumn + 1) To (SFDCColumn + HooversColumn)
            Worksheets("Matching").Activate
            Worksheets("Matching").Cells(1, j).Value = Worksheets("Hoovers").Cells(1, j - SFDCColumn + 1).Value
            Worksheets("Matching").Cells(i, j).Activate
            
            If IF_HAVE_DUNS = True Then
                'ActiveCell.FormulaR1C1 = _
                "=IFERROR(VLOOKUP([@" & SFDC_DUNS.Value & "],Hoovers_data[#All]," & k & ",FALSE), " & Chr(34) & Chr(34) & ")"
                ActiveCell.FormulaR1C1 = _
                "=IFERROR(VLOOKUP([@" & SFDC_DUNS.Value & "]*1,Hoovers_data[#All]," & k & ",FALSE), " & Chr(34) & Chr(34) & ")"
                k = k + 1
            Else
                If checkNumber Then
                    ActiveCell.FormulaR1C1 = _
                    "=IFERROR(VLOOKUP([@[" & SFDC_DUNS.Value & "]]*1,Hoovers_data[#All]," & k & ",FALSE), " & Chr(34) & Chr(34) & ")"
                    k = k + 1
                Else
                    ActiveCell.FormulaR1C1 = _
                    "=IFERROR(VLOOKUP([@[" & SFDC_DUNS.Value & "]],Hoovers_data[#All]," & k & ",FALSE), " & Chr(34) & Chr(34) & ")"
                    k = k + 1
                End If
            End If
            
            pctDone = ((j - (1 + SFDCColumn + 1)) / ((SFDCColumn + HooversColumn) - (1 + SFDCColumn + 1))) * 100
            If pctDone > 99 Then
                pctDone = 99
            End If
            With Progress
                .theFrameProgress.Caption = "Merging SFDC Data Table & Hoovers Data Table. Complete: " & Int(pctDone) & "%"
                .theLabelProgress.Width = pctDone * 2.4
                DoEvents
            End With
        Next j
        
    Worksheets("Matching").Range(Cells(1, 2), Cells(SfdcRow, SFDCColumn)).Interior.Color = RGB(221, 235, 247)
    Worksheets("Matching").Range(Cells(1, SFDCColumn + 1), Cells(SfdcRow, SFDCColumn + 1)).Interior.Color = RGB(255, 255, 0)
    Worksheets("Matching").Range(Cells(1, SFDCColumn + 2), Cells(SfdcRow, SFDCColumn + HooversColumn)).Interior.Color = _
    RGB(226, 239, 218)

End Sub

Sub ExportPoorMatch()
    Dim eachColumn As Range
    Dim countRowNum, currentColumn As Long
    
    currentColumn = 1
    Worksheets("Matching").Activate
    countRowNum = Worksheets("Matching").Cells(1, 2).CurrentRegion.Rows.Count
    For Each eachColumn In Worksheets("Matching").UsedRange.Columns
        If Worksheets("Matching").Cells(1, eachColumn.Column).Interior.Color = RGB(221, 235, 247) Then
            Range(Worksheets("Matching").Cells(1, eachColumn.Column), Worksheets("Matching").Cells(countRowNum, _
            eachColumn.Column)).SpecialCells(xlCellTypeVisible).Copy Destination:=Worksheets("Output_csv").Cells(1, _
            currentColumn)
            currentColumn = currentColumn + 1
        End If
        
        If Worksheets("Matching").Cells(1, eachColumn.Column).Interior.Color = RGB(255, 255, 0) Then
            Worksheets("Matching").Cells(1, eachColumn.Column).Value = "Original " & Worksheets("Matching").Cells(1, eachColumn.Column).Value
            Range(Worksheets("Matching").Cells(1, eachColumn.Column), Worksheets("Matching").Cells(countRowNum, _
            eachColumn.Column)).SpecialCells(xlCellTypeVisible).Copy Destination:=Worksheets("Output_csv").Cells(1, _
            currentColumn)
            currentColumn = currentColumn + 1
        End If
    Next eachColumn
    
    Call SFDCOutputIntoCSV
    MsgBox Chr(34) & FILE_NAME & ".csv" & Chr(34) & " has been saved to desktop."
    FILE_NAME = ""
    Worksheets("Matching").Cells(1, 1).Select
End Sub


Sub ExportToUpdate()
    Dim eachColumn As Range
    Dim countRowNum, currentColumn As Long
    Dim HaveToReplace As Boolean
    
    currentColumn = 1
    Worksheets("Matching").Activate
    countRowNum = Worksheets("Matching").Cells(1, 2).CurrentRegion.Rows.Count
    For Each eachColumn In Worksheets("Matching").UsedRange.Columns
        If Worksheets("Matching").Cells(1, eachColumn.Column).Value = "Customer: ID" Then
            Range(Worksheets("Matching").Cells(1, eachColumn.Column), Worksheets("Matching").Cells(countRowNum, _
            eachColumn.Column)).SpecialCells(xlCellTypeVisible).Copy Destination:=Worksheets("Output_csv").Cells(1, _
            currentColumn)
            currentColumn = currentColumn + 1
        End If
        If Worksheets("Matching").Cells(1, eachColumn.Column).Interior.Color = RGB(226, 239, 218) Then
            Range(Worksheets("Matching").Cells(1, eachColumn.Column), Worksheets("Matching").Cells(countRowNum, _
            eachColumn.Column)).SpecialCells(xlCellTypeVisible).Copy Destination:=Worksheets("Output_csv").Cells(1, _
            currentColumn)
            currentColumn = currentColumn + 1
        End If
    Next eachColumn
    
    HaveToReplace = True
    For Each eachColumn In Worksheets("Output_csv").UsedRange.Columns
        If Worksheets("Output_csv").Cells(1, eachColumn.Column).Value = "Customer: ID2" And HaveToReplace Then
            Worksheets("Output_csv").Cells(1, eachColumn.Column).Value = "Customer: ID"
            HaveToReplace = False
        End If
        
        If Worksheets("Output_csv").Cells(1, eachColumn.Column).Value = "Legal Name2" And HaveToReplace Then
            Worksheets("Output_csv").Cells(1, eachColumn.Column).Value = "Legal Name"
            HaveToReplace = False
        End If
        
        If Worksheets("Output_csv").Cells(1, eachColumn.Column).Value = "DUNS22" And HaveToReplace Then
            Worksheets("Output_csv").Cells(1, eachColumn.Column).Value = "DUNS"
            HaveToReplace = False
        End If
        
        If Worksheets("Output_csv").Cells(1, eachColumn.Column).Value = "DUNS2" And HaveToReplace Then
            Worksheets("Output_csv").Cells(1, eachColumn.Column).Value = "DUNS"
            HaveToReplace = False
        End If
        
        If Worksheets("Output_csv").Cells(1, eachColumn.Column).Value = "Street Line12" And HaveToReplace Then
            Worksheets("Output_csv").Cells(1, eachColumn.Column).Value = "Street Line1"
            HaveToReplace = False
        End If
        
        If Worksheets("Output_csv").Cells(1, eachColumn.Column).Value = "Street Line 22" And HaveToReplace Then
            Worksheets("Output_csv").Cells(1, eachColumn.Column).Value = "Street Line 2"
            HaveToReplace = False
        End If
        
        If Worksheets("Output_csv").Cells(1, eachColumn.Column).Value = "Street Line 32" And HaveToReplace Then
            Worksheets("Output_csv").Cells(1, eachColumn.Column).Value = "Street Line 3"
            HaveToReplace = False
        End If
        
        If Worksheets("Output_csv").Cells(1, eachColumn.Column).Value = "City2" And HaveToReplace Then
            Worksheets("Output_csv").Cells(1, eachColumn.Column).Value = "City"
            HaveToReplace = False
        End If
        
        If Worksheets("Output_csv").Cells(1, eachColumn.Column).Value = "State/Province2" And HaveToReplace Then
            Worksheets("Output_csv").Cells(1, eachColumn.Column).Value = "State/Province"
            HaveToReplace = False
        End If
        
        If Worksheets("Output_csv").Cells(1, eachColumn.Column).Value = "ZIP2" And HaveToReplace Then
            Worksheets("Output_csv").Cells(1, eachColumn.Column).Value = "ZIP"
            HaveToReplace = False
        End If
        
        If Worksheets("Output_csv").Cells(1, eachColumn.Column).Value = "Country2" And HaveToReplace Then
            Worksheets("Output_csv").Cells(1, eachColumn.Column).Value = "Country"
            HaveToReplace = False
        End If
        
        If Worksheets("Output_csv").Cells(1, eachColumn.Column).Value = "Website2" And HaveToReplace Then
            Worksheets("Output_csv").Cells(1, eachColumn.Column).Value = "Website"
            HaveToReplace = False
        End If
        
        If Worksheets("Output_csv").Cells(1, eachColumn.Column).Value = "DUNS verified2" And HaveToReplace Then
            Worksheets("Output_csv").Cells(1, eachColumn.Column).Value = "DUNS verified"
            HaveToReplace = False
        End If
        
        HaveToReplace = True
    Next eachColumn
    
    Call HooversOutputIntoExcel
    MsgBox Chr(34) & FILE_NAME & ".xlsx" & Chr(34) & " has been saved to desktop."
    FILE_NAME = ""
    Worksheets("Matching").Cells(1, 1).Select
End Sub

Function DunsNumberDigitToNine(ByVal theString As String)
Dim stringDigit, needMoreDigit As Long
stringDigit = Len(theString)
needMoreDigit = 9 - stringDigit
theString = String(needMoreDigit, "0") & theString
DunsNumberDigitToNine = theString
End Function

Sub DUNSFormat()
Dim countRowNum, countColNum As Long
Dim eachHeader, eachCell, clearSimilarity As Range

    Worksheets("Matching").Activate
    countRowNum = Worksheets("Matching").Cells(1, 1).CurrentRegion.Rows.Count
    countColNum = Worksheets("Matching").Cells(1, 1).CurrentRegion.Columns.Count
    For Each eachHeader In Range(Worksheets("Matching").Cells(1, 1), Worksheets("Matching").Cells(1, countColNum)).Cells
        If InStr(1, eachHeader.Value, "DUNS", vbTextCompare) > 0 Then
            If InStr(1, eachHeader.Value, "DUNS verified", vbTextCompare) = 0 Then
                'eachHeader.EntireColumn.NumberFormat = "General"
                eachHeader.EntireColumn.NumberFormat = "@"
                For Each eachCell In Range(Worksheets("Matching").Cells(2, eachHeader.Column), Worksheets("Matching").Cells( _
                countRowNum, eachHeader.Column)).Cells
                If eachCell.Value <> "" And eachCell.Value <> "0" Then
                    eachCell.Value = DunsNumberDigitToNine(eachCell.Value)
                    'eachCell.FormulaR1C1 = _
                "=" & Chr(34) & DunsNumberDigitToNine(eachCell.Value) & Chr(34)
                Else
                    If eachCell.Interior.Color = RGB(226, 239, 218) Then
                        For Each clearSimilarity In Range(Worksheets("Matching").Cells(eachCell.Row, 2), _
                        Worksheets("Matching").Cells(eachCell.Row, countColNum)).Cells
                            If clearSimilarity.Interior.Color = RGB(255, 255, 255) Then
                                clearSimilarity.Value = ""
                            End If
                        Next clearSimilarity
                    End If
                End If
                Next eachCell
            End If
        End If
    Next eachHeader
    
End Sub


Sub ie()
MsgBox Sheets("Matching").Cells(12, 17).Value
MsgBox Sheets("Matching").Cells(12, 17).Text
Sheets("Matching").Cells(12, 17).NumberFormat = "@"
Sheets("Matching").Cells(12, 17).Value = "06560"
End Sub
