VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GIDMatch 
   Caption         =   "Settings"
   ClientHeight    =   9435.001
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   6300
   OleObjectBlob   =   "GIDMatch.frx":0000
   StartUpPosition =   1  '©ÒÄÝµøµ¡¤¤¥¡
End
Attribute VB_Name = "GIDMatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CancelBtn_Click()
    Unload Me
End Sub

Private Sub ClearBtn_Click()
    ModelNCompany.Text = ""
    ModelNCity.Text = ""
    ModelNCountry.Text = ""
    ModelNOID.Text = ""
    ModelNGID.Text = ""
    ModelNState.Text = ""
    'ModelNStatus.Text = ""
    SFDCCompany.Text = ""
    SFDCCity.Text = ""
    SFDCGID.Text = ""
    SFDCStatus.Text = ""
    SFDCCountry.Text = ""
    SFDCState.Text = ""
End Sub

Private Sub Label15_Click()

End Sub

Private Sub Label8_Click()

End Sub

Private Sub StartBtn_Click()
    If ModelNCompany.Text = "" Or ModelNCity.Text = "" Or ModelNCountry.Text = "" Or ModelNOID.Text = "" Or SFDCCompany.Text = "" _
    Or SFDCGID.Text = "" Or SFDCCountry.Text = "" Then
        MsgBox "Please Input All Configs"
    Else
        Set gModelNCompany = Nothing
        Set gModelNCity = Nothing
        Set gModelNCountry = Nothing
        Set gModelNOID = Nothing
        Set gModelNGID = Nothing
        Set gModelNState = Nothing
        'Set gModelNStatus = Nothing
        Set gSFDCCompany = Nothing
        Set gSFDCCity = Nothing
        Set gSFDCGID = Nothing
        Set gSFDCStatus = Nothing
        Set gSFDCCountry = Nothing
        Set gSFDCState = Nothing
        On Error Resume Next
        Set gModelNCompany = Range(ModelNCompany.Text)
        Set gModelNCity = Range(ModelNCity.Text)
        Set gModelNCountry = Range(ModelNCountry.Text)
        Set gModelNOID = Range(ModelNOID.Text)
        Set gModelNGID = Range(ModelNGID.Text)
        Set gModelNState = Range(ModelNState.Text)
        'Set gModelNStatus = Range(ModelNStatus.Text)
        Set gSFDCCompany = Range(SFDCCompany.Text)
        Set gSFDCCity = Range(SFDCCity.Text)
        Set gSFDCGID = Range(SFDCGID.Text)
        Set gSFDCStatus = Range(SFDCStatus.Text)
        Set gSFDCCountry = Range(SFDCCountry.Text)
        Set gSFDCState = Range(SFDCState.Text)
        On Error GoTo 0
        If gModelNCompany Is Nothing Or gModelNCity Is Nothing Or gModelNCountry Is Nothing Or gModelNOID Is Nothing _
        Or gSFDCCompany Is Nothing Or gSFDCGID Is Nothing Or gSFDCCountry Is Nothing Then
            MsgBox "There Is Invalid Input Value"
        Else
            With GIDMatchProgress
                .LabelProgressGID.Width = 0
            End With
    
            GIDMatchProgress.Show
        End If
    End If
End Sub

Private Sub UserForm_Click()

End Sub
