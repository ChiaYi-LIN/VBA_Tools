VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} OutputSetting 
   Caption         =   "Settings"
   ClientHeight    =   930
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   7332
   OleObjectBlob   =   "OutputSetting.frx":0000
   StartUpPosition =   1  '©ÒÄÝµøµ¡¤¤¥¡
End
Attribute VB_Name = "OutputSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub OKBtn_Click()
    Worksheets("DUNS.csv").Cells.Clear
    Worksheets("Salesforce Customers").Activate
    DUNS_ROWS_COUNT = 0
    Set DUNS_RANGE = Nothing
    On Error Resume Next
    Set DUNS_RANGE = Range(DUNSCell.Value)
    On Error GoTo 0
            If DUNSCell.Value = "" Then
                MsgBox "Please select DUNS Number header cell."
            ElseIf DUNS_RANGE Is Nothing Then
                MsgBox "Invalid range. Please check again."
            Else
                DUNS_ROWS_COUNT = Cells(1, DUNS_RANGE.Column).CurrentRegion.Rows.Count
                Columns(DUNS_RANGE.Column).Copy Destination:=Worksheets("DUNS.csv").Columns(1)
                
                Dim Counter, i As Long
                Counter = 1
                For i = 1 To DUNS_ROWS_COUNT
                If Worksheets("DUNS.csv").Cells(i, 1).Value <> "" Then
                    Worksheets("DUNS.csv").Cells(Counter, 2).Value = Worksheets("DUNS.csv").Cells(i, 1).Value
                    Counter = Counter + 1
                End If
                Next i
                Worksheets("DUNS.csv").Cells(1, 1).EntireColumn.Delete
                
                Call DUNSOutputIntoCSV
                MsgBox Chr(34) & FILE_NAME & ".csv" & Chr(34) & " has been saved to desktop."
            End If
    
    Set DUNS_RANGE = Nothing
    FILE_NAME = ""
    Cells(1, 1).Select
    Unload Me
End Sub

