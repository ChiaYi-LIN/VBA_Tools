VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} theProgress 
   Caption         =   "Processing..."
   ClientHeight    =   1185
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   5676
   OleObjectBlob   =   "theProgress.frx":0000
   StartUpPosition =   1  '���ݵ�������
End
Attribute VB_Name = "theProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Activate()
    Call Main
End Sub

Private Sub UserForm_Click()
    Unload Me
End Sub
