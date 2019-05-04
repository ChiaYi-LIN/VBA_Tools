VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GIDMatchProgress 
   Caption         =   "Matching GID..."
   ClientHeight    =   1185
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   5676
   OleObjectBlob   =   "GIDMatchProgress.frx":0000
   StartUpPosition =   1  '©ÒÄÝµøµ¡¤¤¥¡
End
Attribute VB_Name = "GIDMatchProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Activate()
    Call OIDandGIDMatching
End Sub
