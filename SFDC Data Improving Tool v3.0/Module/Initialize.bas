Attribute VB_Name = "Initialize"
Public DUNS_RANGE As Range
Public DUNS_ROWS_COUNT, ALL_COLUMN As Long
Public SFDC_ALL_DATA, HOOVERS_ALL_DATA As Range
Public SFDC_LEGAL_NAME, SFDC_COUNTRY, SFDC_CITY, SFDC_ADDRESS, SFDC_DUNS As Range
Public HOOVERS_LEGAL_NAME, HOOVERS_COUNTRY, HOOVERS_CITY, HOOVERS_ADDRESS, HOOVERS_DUNS As Range
Public CUSTOM_SET_WEIGHT As Boolean
Public WEIGHT_LEGAL_NAME, WEIGHT_COUNTRY, WEIGHT_CITY, WEIGHT_ADDRESS As Double
Public SIM_LEGAL_NAME, SIM_COUNTRY, SIM_CITY, SIM_ADDRESS, SIM_INTEGRATED As Boolean
Public IF_HAVE_DUNS As Boolean
Public FILE_NAME As String
Public GOOGLE_HAS_QUERY, USE_GOOGLE_API As Boolean
Public QUERY_USED As Long

Sub Initialization()
Dim eachWorksheet As Worksheet
For Each eachWorksheet In ThisWorkbook.Worksheets
    If eachWorksheet.Name <> "Matching" Then
        With eachWorksheet.Cells.Font
            .Name = "Calibri"
            .Size = 11
            .Bold = False
            .Italic = False
        End With
    Else
        With eachWorksheet.Cells.Font
            .Name = "Calibri"
            .Size = 11
            .Bold = False
            .Italic = False
            .Color = vbBlack
        End With
    End If
    
    If eachWorksheet.Name <> "Matching" Then
        If eachWorksheet.Name = "Readme" Then
            With eachWorksheet
                .Columns.ColumnWidth = 25
                .Columns(2).ColumnWidth = 7.5
                .Columns(3).ColumnWidth = 22
                .Columns(4).ColumnWidth = 34.5
                .Columns(5).ColumnWidth = 85.5
                .Rows.RowHeight = 45
            End With
        Else
            With eachWorksheet
                .Cells.Interior.ColorIndex = xlnofill
                .Columns.ColumnWidth = 25
                .Rows.RowHeight = 15
                .Rows(1).Font.Bold = True
            End With
        End If
    Else
        With eachWorksheet
            .Columns.ColumnWidth = 25
            .Rows.RowHeight = 15
            .Rows(1).Font.Bold = True
        End With
    End If
    
    eachWorksheet.Activate
    eachWorksheet.Cells(1, 1).Select
Next eachWorksheet
End Sub

Sub ClearAllCells()
Dim eachWorksheet As Worksheet
For Each eachWorksheet In ThisWorkbook.Worksheets
    If eachWorksheet.Name <> "Readme" Then
        eachWorksheet.Cells.Clear
        eachWorksheet.Cells.Clear
    End If
Next eachWorksheet

End Sub

Sub ResetButtons()
    Dim btn As Shape
    Dim ws As Worksheet
    
    For Each ws In Worksheets
        If ws.Visible = xlSheetVisible Then
            ws.Select
            ActiveWindow.Zoom = 100
            ActiveWindow.ScrollColumn = 1
            ActiveWindow.ScrollRow = 1
        End If
    Next ws
    
    Set btn = Worksheets("Readme").Shapes("ResetAllBtn")
    With btn
        .Height = 51
        .Width = 105
        .Left = 16.5
        .Top = 19.5
    End With
    
    Set btn = Worksheets("Readme").Shapes("ClearAllBtn")
    With btn
        .Height = 51
        .Width = 105
        .Left = 16.5
        .Top = 109.5
    End With
    
    Set btn = Worksheets("Salesforce Customers").Shapes("ImportSFDCcsv")
    With btn
        .Height = 36
        .Width = 105
        .Left = 16.5
        .Top = 19.5
    End With
    
    Set btn = Worksheets("Salesforce Customers").Shapes("OutputDUNS")
    With btn
        .Height = 36
        .Width = 105
        .Left = 16.5
        .Top = 79.5
    End With
    
    Set btn = Worksheets("Hoovers").Shapes("ImportHooversCSV")
    With btn
        .Height = 36
        .Width = 105
        .Left = 16.5
        .Top = 19.5
    End With
    
    Set btn = Worksheets("Matching").Shapes("MatchSetConfig")
    With btn
        .Height = 36
        .Width = 105
        .Left = 16.5
        .Top = 19.5
    End With
    
    Set btn = Worksheets("Matching").Shapes("SFDCOutput")
    With btn
        .Height = 36
        .Width = 105
        .Left = 16.5
        .Top = 79.5
    End With
    
    Set btn = Worksheets("Matching").Shapes("MassUpdate")
    With btn
        .Height = 36
        .Width = 105
        .Left = 16.5
        .Top = 139.5
    End With
    
    Set btn = Worksheets("Matching").Shapes("MatchingScoringResults")
    With btn
        .Height = 36
        .Width = 105
        .Left = 16.5
        .Top = 199.5
    End With
End Sub

Sub ResetAll()
    Application.ScreenUpdating = False
    Call Initialization
    Call ResetButtons
    Worksheets("Readme").Activate
    Application.ScreenUpdating = True
End Sub

Sub ClearAll()
    Application.ScreenUpdating = False
    Call ClearAllCells
    Call ResetAll
    Worksheets("Readme").Activate
    Application.ScreenUpdating = True
End Sub

