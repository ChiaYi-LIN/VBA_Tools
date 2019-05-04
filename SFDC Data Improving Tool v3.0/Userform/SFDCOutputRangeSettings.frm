VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SFDCOutputRangeSettings 
   Caption         =   "Settings"
   ClientHeight    =   930
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   7332
   OleObjectBlob   =   "SFDCOutputRangeSettings.frx":0000
   StartUpPosition =   1  '©ÒÄÝµøµ¡¤¤¥¡
End
Attribute VB_Name = "SFDCOutputRangeSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''No use currently
Private Sub ReadyToOutput_Click()
    
    Dim SFDCRange As Range
    Dim SFDCColumn As Range
    Dim CopyColumn, countRowNum As Long
    
    Worksheets("Output_csv").Cells.Clear
    Worksheets("Matching").Activate
    countRowNum = Worksheets("Matching").Cells(1, 2).CurrentRegion.Rows.Count
    
    CopyColumn = 1
    Set SFDCRange = Nothing
    On Error Resume Next
    Set SFDCRange = Range(SFDCDataRange.Value)
    On Error GoTo 0
            If SFDCDataRange.Value = "" Then
                MsgBox "Please select all SFDC header cells."
            ElseIf SFDCRange Is Nothing Then
                MsgBox "Invalid range. Please check again."
            Else
                For Each SFDCColumn In SFDCRange.Columns
                    'columns(SFDCColumn.Column).SpecialCells(xlCellTypeVisible).Copy Destination:=Worksheets("Output_csv").columns(CopyColumn)
                    Range(Cells(1, SFDCColumn.Column), Cells(countRowNum, SFDCColumn.Column)).SpecialCells(xlCellTypeVisible).Copy _
                    Destination:=Worksheets("Output_csv").Cells(1, CopyColumn)
                    
                    CopyColumn = CopyColumn + 1
               Next SFDCColumn
                
                Call SFDCOutputIntoCSV
                MsgBox Chr(34) & "Poor_Match_SFDC_Customers.csv" & Chr(34) & " has been saved to desktop."
            End If
    
    Cells(1, 1).Select
    Unload Me
End Sub



