VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ConfigSetting 
   Caption         =   "Config Settings"
   ClientHeight    =   7050
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   5088
   OleObjectBlob   =   "ConfigSetting.frx":0000
   StartUpPosition =   1  '©ÒÄÝµøµ¡¤¤¥¡
End
Attribute VB_Name = "ConfigSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CancelButton_Click()
    Unload ConfigSetting
End Sub

Private Sub ClearSettingButton_Click()
    CompanyNameRef.Text = ""
    LocationNameRef.Text = ""
    LevelHeader.Text = ""
    GIDRef = ""
    InformationRef.Text = ""
    LocationLevel.Text = ""
End Sub

Private Sub GIDRef_BeforeDragOver(Cancel As Boolean, ByVal Data As MSForms.DataObject, ByVal x As stdole.OLE_XPOS_CONTAINER, ByVal y As stdole.OLE_YPOS_CONTAINER, ByVal DragState As MSForms.fmDragState, Effect As MSForms.fmDropEffect, ByVal Shift As Integer)

End Sub

Private Sub LocationLevel_Change()
    ConfigSetting.LevelLabel.Caption = "Select " & ConfigSetting.LocationLevel.Value & " Name Header"
End Sub

Private Sub OIDRef_BeforeDragOver(Cancel As Boolean, ByVal Data As MSForms.DataObject, ByVal x As stdole.OLE_XPOS_CONTAINER, ByVal y As stdole.OLE_YPOS_CONTAINER, ByVal DragState As MSForms.fmDragState, Effect As MSForms.fmDropEffect, ByVal Shift As Integer)

End Sub

Private Sub OkButton_Click()
    Dim checkFlagOne, checkFlagTwo, checkFlagThree, checkFlagFour, checkFlagFive, checkFlagSix, checkFlagSeven, checkFlagEight As Range
        
    If CompanyNameRef.Text = "" Or LocationNameRef.Text = "" Or GIDRef = "" Or InformationRef.Text = "" Or OIDRef.Text = "" _
    Or LevelHeader = "" Or ParentHeader = "" Or StatusHeader.Text = "" Then
    MsgBox "Please Input All Configs"
    Else
        On Error Resume Next
        Set checkFlagOne = Range(CompanyNameRef.Value)
        Set checkFlagTwo = Range(LocationNameRef.Value)
        Set checkFlagThree = Range(GIDRef.Value)
        Set checkFlagFour = Range(InformationRef.Value)
        Set checkFlagFive = Range(OIDRef.Value)
        Set checkFlagSix = Range(LevelHeader.Value)
        Set checkFlagSeven = Range(ParentHeader.Value)
        Set checkFlagEight = Range(StatusHeader.Value)
        If checkFlagOne Is Nothing Or checkFlagTwo Is Nothing Or checkFlagThree Is Nothing Or checkFlagFour Is Nothing Or checkFlagFive _
        Is Nothing Or checkFlagSix Is Nothing Or checkFlagSeven Is Nothing Or checkFlagEight Is Nothing Then
        MsgBox "There Is Invalid Range Value"
        Else
            Set PUBCompanyName = Range(CompanyNameRef.Value)
            Set PUBLocationName = Range(LocationNameRef.Value)
            PUBLocationLevel = LocationLevel.Value
            Set PUBGIDRef = Range(GIDRef.Value)
            Set PUBInformation = Range(InformationRef.Value)
            Set PUBOIDRef = Range(OIDRef.Value)
            Set PUBLevelHeader = Range(LevelHeader.Value)
            Set PUBParentHeader = Range(ParentHeader.Value)
            Set PUBStatusHeader = Range(StatusHeader.Value)
            'Call IndicateMasterAndAliased(Range(CompanyNameRef.Value), Range(LocationNameRef.Value), LocationLevel.Value, _
            Range(GIDRef.Value), Range(InformationRef.Value), Range(OIDRef.Value), Range(LevelHeader.Value), _
            Range(ParentHeader.Value))
            ResultProcess.LabelProgressResults.Width = 0
            ResultProcess.Show
            
        End If
    End If
End Sub

Private Sub UserForm_Click()

End Sub
