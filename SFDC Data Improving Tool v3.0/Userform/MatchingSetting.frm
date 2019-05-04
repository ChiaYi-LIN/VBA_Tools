VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MatchingSetting 
   Caption         =   "Settings"
   ClientHeight    =   6975
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   10236
   OleObjectBlob   =   "MatchingSetting.frx":0000
   StartUpPosition =   1  '©ÒÄÝµøµ¡¤¤¥¡
End
Attribute VB_Name = "MatchingSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub StartMatching_Click()
    Dim emptyWeight As Long
    emptyWeight = 0
    WEIGHT_LEGAL_NAME = 0
    WEIGHT_COUNTRY = 0
    WEIGHT_CITY = 0
    WEIGHT_ADDRESS = 0
    
    On Error Resume Next
    Set SFDC_LEGAL_NAME = Nothing
    Set SFDC_COUNTRY = Nothing
    Set SFDC_CITY = Nothing
    Set SFDC_ADDRESS = Nothing
    Set SFDC_DUNS = Nothing
    Set HOOVERS_LEGAL_NAME = Nothing
    Set HOOVERS_COUNTRY = Nothing
    Set HOOVERS_CITY = Nothing
    Set HOOVERS_ADDRESS = Nothing
    Set HOOVERS_DUNS = Nothing
    
    Set SFDC_LEGAL_NAME = Range(uSFDCLegalName.Value)
    Set SFDC_COUNTRY = Range(uSFDCCountry.Value)
    Set SFDC_CITY = Range(uSFDCCity.Value)
    Set SFDC_ADDRESS = Range(uSFDCAddress.Value)
    Set SFDC_DUNS = Range(uSFDCDUNS.Value)
    Set HOOVERS_LEGAL_NAME = Range(uHooversLegalName.Value)
    Set HOOVERS_COUNTRY = Range(uHooversCountry.Value)
    Set HOOVERS_CITY = Range(uHooversCity.Value)
    Set HOOVERS_ADDRESS = Range(uHooversAddress.Value)
    Set HOOVERS_DUNS = Range(uHooversDUNS.Value)
    On Error GoTo 0
    
    If uSFDCLegalName.Value = "" Or uSFDCCountry.Value = "" Or uSFDCCity.Value = "" Or _
    uSFDCAddress.Value = "" Or uSFDCDUNS.Value = "" Or uHooversLegalName.Value = "" Or _
    uHooversCountry.Value = "" Or uHooversCity.Value = "" Or uHooversAddress.Value = "" Or uHooversDUNS.Value = "" Then
        MsgBox "Please input all ranges."
    ElseIf SFDC_LEGAL_NAME Is Nothing Or SFDC_COUNTRY Is Nothing Or SFDC_CITY Is Nothing Or _
    SFDC_ADDRESS Is Nothing Or SFDC_DUNS Is Nothing Or HOOVERS_LEGAL_NAME Is Nothing Or HOOVERS_COUNTRY Is Nothing Or _
    HOOVERS_CITY Is Nothing Or HOOVERS_ADDRESS Is Nothing Or HOOVERS_DUNS Is Nothing Then
        MsgBox "There is invalid range. Please check again."
    Else
        If CheckLegalName.Value = False And CheckCountry.Value = False And CheckCity.Value = False And _
        CheckAddress.Value = False And CheckIntegrated.Value = False Then
        MsgBox "Please select at least one output criterion."
        Else
            If SetWeight.Value = True Then
                If WeightLegalName.Value = "" Then
                    emptyWeight = emptyWeight + 1
                End If
                If WeightCountry.Value = "" Then
                    emptyWeight = emptyWeight + 1
                End If
                If WeightCity.Value = "" Then
                    emptyWeight = emptyWeight + 1
                End If
                If WeightAddress.Value = "" Then
                    emptyWeight = emptyWeight + 1
                End If
                
                On Error Resume Next
                WEIGHT_LEGAL_NAME = Val(WeightLegalName.Value)
                WEIGHT_COUNTRY = Val(WeightCountry.Value)
                WEIGHT_CITY = Val(WeightCity.Value)
                WEIGHT_ADDRESS = Val(WeightAddress.Value)
                On Error GoTo 0
            End If
        
            If SetWeight.Value = True And emptyWeight > 0 Then
                MsgBox "Please set the weight for each criterion."
            Else
                If SetWeight.Value = True And (WEIGHT_LEGAL_NAME + WEIGHT_COUNTRY + WEIGHT_CITY + WEIGHT_ADDRESS) <> 100 Then
                    MsgBox "Total weight should be 100."
                Else
                    If CheckLegalName.Value = True Then
                        SIM_LEGAL_NAME = True
                    End If
                    If CheckCountry.Value = True Then
                        SIM_COUNTRY = True
                    End If
                    If CheckCity.Value = True Then
                        SIM_CITY = True
                    End If
                    If CheckAddress.Value = True Then
                        SIM_ADDRESS = True
                    End If
                    If CheckIntegrated.Value = True Then
                        SIM_INTEGRATED = True
                    End If
                    
                    If HaveDUNS.Value = True Then
                        IF_HAVE_DUNS = True
                    ElseIf NoDUNS.Value = True Then
                        IF_HAVE_DUNS = False
                    End If
                    
                    USE_GOOGLE_API = GeoApi.Value
                    
                    Progress.theLabelProgress.Width = 0
                    Progress.theFrameProgress.Caption = "Merging SFDC Data Table & Hoovers Data Table. Complete: 0%"
                    Progress.Show
                    
                    On Error Resume Next
                    Set SFDC_ALL_DATA = Nothing
                    Set HOOVERS_ALL_DATA = Nothing
                    Set SFDC_LEGAL_NAME = Nothing
                    Set SFDC_COUNTRY = Nothing
                    Set SFDC_CITY = Nothing
                    Set SFDC_ADDRESS = Nothing
                    Set SFDC_DUNS = Nothing
                    Set HOOVERS_LEGAL_NAME = Nothing
                    Set HOOVERS_COUNTRY = Nothing
                    Set HOOVERS_CITY = Nothing
                    Set HOOVERS_ADDRESS = Nothing
                    Set HOOVERS_DUNS = Nothing
                    On Error GoTo 0
                    
                    Unload Me
                    Application.ScreenUpdating = True
                End If
            End If
        End If
    End If
End Sub

Private Sub CancelMatching_Click()
    Unload Me
End Sub

Private Sub CheckIntegrated_Change()
    If CheckIntegrated.Value = True And CUSTOM_SET_WEIGHT = False Then
    IntegrateEach.Value = True
    SetWeight.Value = False
    ElseIf CheckIntegrated.Value = True And CUSTOM_SET_WEIGHT = True Then
    IntegrateEach.Value = False
    SetWeight.Value = True
    Else
    IntegrateEach.Value = False
    SetWeight.Value = False
    End If
End Sub

Private Sub ClearMatching_Click()
    Call InitializeMatchingSetting
End Sub

Private Sub HaveDUNS_Click()
    LabelSFDCDUNS.Caption = "DUNS :"
    LabelHooversDUNS.Caption = "DUNS :"
End Sub

Private Sub HaveDUNS_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Worksheets("Salesforce Customers").Activate
End Sub

Private Sub NoDUNS_Click()
    LabelSFDCDUNS.Caption = "GID :"
    LabelHooversDUNS.Caption = "GID :"
End Sub

Private Sub NoDUNS_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Worksheets("Salesforce Customers").Activate
End Sub

Private Sub UserForm_Activate()
    Call InitializeMatchingSetting
End Sub

Private Sub IntegrateEach_Click()
    CUSTOM_SET_WEIGHT = False
    IntegrateEach.Value = True
    CheckIntegrated.Value = True
End Sub

Private Sub SetWeight_Click()
    CUSTOM_SET_WEIGHT = True
    SetWeight.Value = True
    CheckIntegrated.Value = True
End Sub

Private Sub InitializeMatchingSetting()
    uSFDCLegalName.Value = ""
    uSFDCCountry.Value = ""
    uSFDCCity.Value = ""
    uSFDCAddress.Value = ""
    uSFDCDUNS.Value = ""
    uHooversLegalName.Value = ""
    uHooversCountry.Value = ""
    uHooversCity.Value = ""
    uHooversAddress.Value = ""
    uHooversDUNS.Value = ""
    WeightLegalName.Value = ""
    WeightCountry.Value = ""
    WeightCity.Value = ""
    WeightAddress.Value = ""
    Call IntegrateEach_Click
    CheckLegalName.Value = True
    CheckCountry.Value = True
    CheckCity.Value = True
    CheckAddress.Value = True
    HaveDUNS.Value = True
    GeoApi.Value = False
End Sub

Private Sub WeightAddress_Change()
    Dim RemoveChar As Integer
    Dim DefaultValue As Long
    RemoveChar = 1
    If Not IsNumeric(WeightAddress.Value) And WeightAddress.Value <> "" Then
        If Len(WeightAddress.Value) < 1 Then RemoveChar = 0
        WeightAddress.Value = Left(WeightAddress.Value, Len(WeightAddress.Value) - RemoveChar)
    End If
    WeightAddress.Value = Replace(WeightAddress.Value, " ", "", , , vbTextCompare)
End Sub

Private Sub WeightCity_Change()
    Dim RemoveChar As Integer
    Dim DefaultValue As Long
    RemoveChar = 1
    If Not IsNumeric(WeightCity.Value) And WeightCity.Value <> "" Then
        If Len(WeightCity.Value) < 1 Then RemoveChar = 0
        WeightCity.Value = Left(WeightCity.Value, Len(WeightCity.Value) - RemoveChar)
    End If
    WeightCity.Value = Replace(WeightCity.Value, " ", "", , , vbTextCompare)
End Sub

Private Sub WeightCountry_Change()
    Dim RemoveChar As Integer
    Dim DefaultValue As Long
    RemoveChar = 1
    If Not IsNumeric(WeightCountry.Value) And WeightCountry.Value <> "" Then
        If Len(WeightCountry.Value) < 1 Then RemoveChar = 0
        WeightCountry.Value = Left(WeightCountry.Value, Len(WeightCountry.Value) - RemoveChar)
    End If
    WeightCountry.Value = Replace(WeightCountry.Value, " ", "", , , vbTextCompare)
End Sub

Private Sub WeightLegalName_Change()
    Dim RemoveChar As Integer
    Dim DefaultValue As Long
    RemoveChar = 1
    If Not IsNumeric(WeightLegalName.Value) And WeightLegalName.Value <> "" Then
        If Len(WeightLegalName.Value) < 1 Then RemoveChar = 0
        WeightLegalName.Value = Left(WeightLegalName.Value, Len(WeightLegalName.Value) - RemoveChar)
    End If
    WeightLegalName.Value = Replace(WeightLegalName.Value, " ", "", , , vbTextCompare)
End Sub

Private Sub uSFDCDUNS_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As stdole.OLE_XPOS_CONTAINER, ByVal Y As stdole.OLE_YPOS_CONTAINER)
    Worksheets("Salesforce Customers").Activate
End Sub
Private Sub uSFDCLegalName_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As stdole.OLE_XPOS_CONTAINER, ByVal Y As stdole.OLE_YPOS_CONTAINER)
    Worksheets("Salesforce Customers").Activate
End Sub
Private Sub uSFDCCountry_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As stdole.OLE_XPOS_CONTAINER, ByVal Y As stdole.OLE_YPOS_CONTAINER)
    Worksheets("Salesforce Customers").Activate
End Sub
Private Sub uSFDCCity_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As stdole.OLE_XPOS_CONTAINER, ByVal Y As stdole.OLE_YPOS_CONTAINER)
    Worksheets("Salesforce Customers").Activate
End Sub
Private Sub uSFDCAddress_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As stdole.OLE_XPOS_CONTAINER, ByVal Y As stdole.OLE_YPOS_CONTAINER)
    Worksheets("Salesforce Customers").Activate
End Sub
Private Sub uHooversDUNS_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As stdole.OLE_XPOS_CONTAINER, ByVal Y As stdole.OLE_YPOS_CONTAINER)
    Worksheets("Hoovers").Activate
End Sub
Private Sub uHooversLegalName_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As stdole.OLE_XPOS_CONTAINER, ByVal Y As stdole.OLE_YPOS_CONTAINER)
    Worksheets("Hoovers").Activate
End Sub
Private Sub uHooversCountry_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As stdole.OLE_XPOS_CONTAINER, ByVal Y As stdole.OLE_YPOS_CONTAINER)
    Worksheets("Hoovers").Activate
End Sub
Private Sub uHooversCity_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As stdole.OLE_XPOS_CONTAINER, ByVal Y As stdole.OLE_YPOS_CONTAINER)
    Worksheets("Hoovers").Activate
End Sub
Private Sub uHooversAddress_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As stdole.OLE_XPOS_CONTAINER, ByVal Y As stdole.OLE_YPOS_CONTAINER)
    Worksheets("Hoovers").Activate
End Sub
