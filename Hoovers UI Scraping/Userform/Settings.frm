VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Settings 
   Caption         =   "Settings"
   ClientHeight    =   6345
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   5208
   OleObjectBlob   =   "Settings.frx":0000
   StartUpPosition =   1  '©ÒÄÝµøµ¡¤¤¥¡
End
Attribute VB_Name = "Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CancelBtn_Click()
    Unload Me
End Sub

Private Sub CellsRange_BeforeDragOver(Cancel As Boolean, ByVal Data As MSForms.DataObject, ByVal x As stdole.OLE_XPOS_CONTAINER, ByVal y As stdole.OLE_YPOS_CONTAINER, ByVal DragState As MSForms.fmDragState, Effect As MSForms.fmDropEffect, ByVal Shift As Integer)

End Sub

Private Sub ShowPassword_Click()
    If ShowPassword.Value = False Then
        Password.PasswordChar = "*"
    Else
        Password.PasswordChar = ""
    End If
End Sub

Private Sub StartBtn_Click()
    If (SearchDuns.Value = False And SearchInformation.Value = False) Or (AllCells.Value = False And FromCells.Value = False) Or Account.Value = "" Or Password.Value = "" Then
        MsgBox "Please input each value."
    Else
        If SearchDuns.Value = True Then
            HAVE_DUNS = False
        Else
            HAVE_DUNS = True
        End If
        
        If FromCells.Value = True Then
            Set SELECT_RANGE = Nothing
            On Error Resume Next
                Set SELECT_RANGE = Range(CellsRange.Value)
            On Error GoTo 0
            If CellsRange.Value = "" Then
                MsgBox "Please select the row which you would like to start with."
            ElseIf SELECT_RANGE Is Nothing Then
                MsgBox "Invalid range. Please check again."
            Else
                
                START_ROW = SELECT_RANGE.Row
                LOGIN_ACCOUNT = Account.Value
                LOGIN_PASSWORD = Password.Value
                Call RememberMe
                Call loginHoovers
                'With Progress
                '    .theFrameProgress.Caption = "Complete: 0%"
                '    .theLabelProgress.Width = 0
                '    .Show
                'End With
                
            End If
        Else
            
            START_ROW = 2
            LOGIN_ACCOUNT = Account.Value
            LOGIN_PASSWORD = Password.Value
            Call RememberMe
            Call loginHoovers
            'With Progress
            '    .theFrameProgress.Caption = "Complete: 0%"
            '    .theLabelProgress.Width = 0
            '    .Show
            'End With
            
        End If
        
    End If
End Sub

Sub RememberMe()
    If RememberAll.Value = True Then
        Worksheets("Account").Cells(1, 1).Value = Account.Value
        Worksheets("Account").Cells(1, 2).Value = Password.Value
        
    Else
        Worksheets("Account").Cells(1, 1).Value = ""
        Worksheets("Account").Cells(1, 2).Value = ""
        
    End If
End Sub

Private Sub CheckRemember()
    If Worksheets("Account").Cells(1, 1).Value <> "" Then
        Account.Value = Worksheets("Account").Cells(1, 1).Value
    End If
    If Worksheets("Account").Cells(1, 2).Value <> "" Then
        Password.Value = Worksheets("Account").Cells(1, 2).Value
    End If
End Sub

Private Sub UserForm_Activate()
    Call CheckRemember
    RememberAll.Value = True
End Sub






