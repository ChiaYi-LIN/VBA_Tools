VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Deduplicate 
   Caption         =   "Processing..."
   ClientHeight    =   1185
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   5676
   OleObjectBlob   =   "Deduplicate.frx":0000
   StartUpPosition =   1  '©ÒÄÝµøµ¡¤¤¥¡
End
Attribute VB_Name = "Deduplicate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Activate()
    If PROGRESS_BAR_MODE = 1 Then
        Call dedup
    ElseIf PROGRESS_BAR_MODE = 2 Then
        Call DataProcessForMasterAliased
    ElseIf PROGRESS_BAR_MODE = 3 Then
        Call DeleteDupDataByID
    Else
        MsgBox "Please check the processing mode."
    End If
End Sub


Private Sub UserForm_Click()
    Unload Me
End Sub
