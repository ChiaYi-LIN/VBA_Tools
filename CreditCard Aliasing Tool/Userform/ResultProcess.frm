VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ResultProcess 
   Caption         =   "Exporting Results..."
   ClientHeight    =   1215
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   5628
   OleObjectBlob   =   "ResultProcess.frx":0000
   StartUpPosition =   1  '©ÒÄÝµøµ¡¤¤¥¡
End
Attribute VB_Name = "ResultProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Activate()

    'Call PassValueBetweenForms(Range("A1"), Range("A1"), "", Range("A1"), Range("A1"), Range("A1"), Range("A1"), Range("A1"), 1)
    Call CreditCard(PUBCompanyName, PUBOIDRef)
End Sub

Private Sub UserForm_Click()
    Unload Me
End Sub
