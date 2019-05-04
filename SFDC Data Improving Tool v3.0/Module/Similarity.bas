Attribute VB_Name = "Similarity"
#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr) 'For 64 Bit Systems
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'For 32 Bit Systems
#End If

Sub GetSimilarity()
Dim i, countRowNum, countColNum, currentColumn, j As Long
Dim legalnameSimilarity, countrySimilarity, citySimilarity, addressSimilarity, integratedSimilarity As Double
Dim String1, String2 As String
Dim wholeString1, wholeString2 As String
Dim GoogleStreet1, GoogleStreet2 As String
Dim pctDone As Long
Dim fullAddress1, fullAddress2 As String
Dim seperateAddress1, seperateAddress2 As Range
Dim ToSeperate As Boolean
Dim keepString1, keepString2, newAddressSimilarity As String

With Progress
    .theFrameProgress.Caption = "Calculating Data Similarity. Complete: 0%"
    .theLabelProgress.Width = pctDone * 2.4
End With

GOOGLE_HAS_QUERY = True
QUERY_USED = 0

Call DictionaryInitialization
Worksheets("Matching").Activate
countRowNum = Cells(1, 1).CurrentRegion.Rows.Count
countColNum = Cells(1, 1).CurrentRegion.Columns.Count

GoTo SkipToNext:

For i = 2 To countRowNum
GoogleStreet1 = standardizeAddress(Cells(i, STREET_SALESFORCE).Value)
GoogleStreet2 = standardizeAddress(Cells(i, STREET_DATADOTCOM).Value)

Cells(i, STREET_SALESFORCE_GOOGLE).Value = GoogleStreet1
Cells(i, STREET_DATADOTCOM_GOOGLE).Value = GoogleStreet2

If googlestreet = "ERROR (OVER_QUERY_LIMIT)" Or GoogleStreet2 = "ERROR (OVER_QUERY_LIMIT)" Then
    Call OptimizeCode_End
    Unload Progress
    MsgBox "Check the internet explorer. (Stopped at row " & i & " )"
    Exit Sub
End If

'progress bar
    pctDone = i / countRowNum * 100
    With Progress
        .theFrameProgress.Caption = "Complete(1/2): " & pctDone & "%"
        .theLabelProgress.Width = pctDone * 2.4
        DoEvents
    End With
    
GoogleStreet1 = ""
GoogleStreet2 = ""
Next i

SkipToNext:

currentColumn = countColNum + 1
If SIM_INTEGRATED = True Then
    Cells(1, currentColumn).Value = "Integrated Similarity"
    currentColumn = currentColumn + 1
End If
If SIM_LEGAL_NAME = True Then
    Cells(1, currentColumn).Value = "Legal Name Similarity"
    currentColumn = currentColumn + 1
End If
If SIM_COUNTRY = True Then
    Cells(1, currentColumn).Value = "Country Similarity"
    currentColumn = currentColumn + 1
End If
If SIM_CITY = True Then
    Cells(1, currentColumn).Value = "City Similarity"
    currentColumn = currentColumn + 1
End If
If SIM_ADDRESS = True Then
    Cells(1, currentColumn).Value = "Address Similarity"
    currentColumn = currentColumn + 1
End If
If SIM_INTEGRATED = True Or (SIM_LEGAL_NAME = True And SIM_COUNTRY = True And SIM_CITY And SIM_ADDRESS) Then
    Cells(1, currentColumn).Value = "Matching Level"
    currentColumn = currentColumn + 1
End If
Worksheets("Matching").Range(Cells(1, countColNum + 1), Cells(1, currentColumn - 1)).Interior.Color = RGB(255, 255, 102)

For i = 2 To countRowNum
'For i = 2 To 8
'MsgBox Trim(Cells(i, 1 + SFDC_ALL_DATA.Columns.Count + HOOVERS_LEGAL_NAME.Column - 2).Value)

If Trim(Cells(i, 1 + SFDC_ALL_DATA.Columns.Count + HOOVERS_LEGAL_NAME.Column - 2).Value) <> "" And _
    Trim(Cells(i, 1 + SFDC_ALL_DATA.Columns.Count + HOOVERS_LEGAL_NAME.Column - 2).Value) <> "0" Then
    legalnameSimilarity = 0
    countrySimilarity = 0
    citySimilarity = 0
    addressSimilarity = 0
    Set seperateAddress1 = Nothing
    Set seperateAddress2 = Nothing
    
    currentColumn = countColNum + 1
    If SIM_INTEGRATED = True Then
        currentColumn = currentColumn + 1
    End If
    
    If SIM_LEGAL_NAME = True Or (SIM_INTEGRATED = True And CUSTOM_SET_WEIGHT = False) Or (SIM_INTEGRATED = True And CUSTOM_SET_WEIGHT = True And WEIGHT_LEGAL_NAME > 0) Then
    String1 = CleanCompanyName(Cells(i, SFDC_LEGAL_NAME.Column).Value)
    String2 = CleanCompanyName(Cells(i, 1 + SFDC_ALL_DATA.Columns.Count + HOOVERS_LEGAL_NAME.Column - 2).Value)
    wholeString1 = wholeString1 & String1 & " "
    wholeString2 = wholeString2 & String2 & " "
    If IsTheSame(String1, String2) Then
        legalnameSimilarity = 1
    Else
        legalnameSimilarity = Similarity(String1, String2)
    End If
        If SIM_LEGAL_NAME = True Then
            Cells(i, currentColumn).Value = legalnameSimilarity
            currentColumn = currentColumn + 1
        End If
    End If
    
    If SIM_COUNTRY = True Or (SIM_INTEGRATED = True And CUSTOM_SET_WEIGHT = False) Or (SIM_INTEGRATED = True And CUSTOM_SET_WEIGHT = True And WEIGHT_COUNTRY > 0) Then
    String1 = CleanCountryName(Cells(i, SFDC_COUNTRY.Column).Value)
    String2 = CleanCountryName(Cells(i, 1 + SFDC_ALL_DATA.Columns.Count + HOOVERS_COUNTRY.Column - 2).Value)
    wholeString1 = wholeString1 & String1 & " "
    wholeString2 = wholeString2 & String2 & " "
    If USE_GOOGLE_API Then
        keepString1 = keepString1 & String1 & " "
        keepString2 = keepString2 & String2 & " "
    End If
    If IsTheSame(String1, String2) Then
        countrySimilarity = 1
    Else
        countrySimilarity = Similarity(String1, String2)
    End If
        If SIM_COUNTRY = True Then
            Cells(i, currentColumn).Value = countrySimilarity
            currentColumn = currentColumn + 1
        End If
    End If
    
    If SIM_CITY = True Or (SIM_INTEGRATED = True And CUSTOM_SET_WEIGHT = False) Or (SIM_INTEGRATED = True And CUSTOM_SET_WEIGHT = True And WEIGHT_CITY > 0) Then
    String1 = CleanCityName(Cells(i, SFDC_CITY.Column).Value)
    String2 = CleanCityName(Cells(i, 1 + SFDC_ALL_DATA.Columns.Count + HOOVERS_CITY.Column - 2).Value)
    wholeString1 = wholeString1 & String1 & " "
    wholeString2 = wholeString2 & String2 & " "
    If USE_GOOGLE_API Then
        keepString1 = keepString1 & String1 & " "
        keepString2 = keepString2 & String2 & " "
    End If
    If IsTheSame(String1, String2) Then
        citySimilarity = 1
    Else
        citySimilarity = Similarity(String1, String2)
    End If
        If SIM_CITY = True Then
            Cells(i, currentColumn).Value = citySimilarity
            currentColumn = currentColumn + 1
        End If
    End If
    
    If SIM_ADDRESS = True Or (SIM_INTEGRATED = True And CUSTOM_SET_WEIGHT = False) Or (SIM_INTEGRATED = True And CUSTOM_SET_WEIGHT = True And WEIGHT_ADDRESS > 0) Then
    String1 = ""
    On Error Resume Next
    For Each seperateAddress1 In SFDC_ADDRESS.Columns
        String1 = String1 & " " & Cells(i, seperateAddress1.Column).Value
    Next seperateAddress1
    On Error GoTo 0
    String1 = Replace(Trim(String1), ",", " ", , , vbTextCompare)
    
    String2 = ""
    j = 0
    ToSeperate = False
    For Each seperateAddress2 In HOOVERS_ADDRESS.Columns
        String2 = String2 & " " & Cells(i, 1 + SFDC_ALL_DATA.Columns.Count + seperateAddress2.Column - 2).Value
        If Len(Cells(i, 1 + SFDC_ALL_DATA.Columns.Count + seperateAddress2.Column - 2).Value) > 35 Then
            ToSeperate = True
        End If
    Next seperateAddress2
    String2 = Replace(Trim(String2), ",", " ", , , vbTextCompare)
    If USE_GOOGLE_API Then
        keepString1 = keepString1 & String1 & " "
        keepString2 = keepString2 & String2 & " "
    End If
    If ToSeperate Then
        For Each seperateAddress2 In HOOVERS_ADDRESS.Columns
            Cells(i, 1 + SFDC_ALL_DATA.Columns.Count + seperateAddress2.Column - 2).Value = Mid(String2, 35 * j + 1, 35)
            j = j + 1
        Next seperateAddress2
    End If
    
    wholeString1 = wholeString1 & String1 & " "
    wholeString2 = wholeString2 & String2 & " "
    If IsTheSame(String1, String2) Then
        addressSimilarity = 1
    Else
        addressSimilarity = Similarity(String1, String2)
    End If
        If SIM_ADDRESS = True Then
            Cells(i, currentColumn).Value = addressSimilarity
            currentColumn = currentColumn + 1
        End If
    End If
    
    If SIM_INTEGRATED = True Then
        If CUSTOM_SET_WEIGHT = False Then
            integratedSimilarity = Similarity(wholeString1, wholeString2)
            Cells(i, countColNum + 1).Value = integratedSimilarity
        Else
            integratedSimilarity = (legalnameSimilarity * WEIGHT_LEGAL_NAME + _
            countrySimilarity * WEIGHT_COUNTRY + citySimilarity * WEIGHT_CITY + addressSimilarity * WEIGHT_ADDRESS) / 100
            Cells(i, countColNum + 1).Value = integratedSimilarity
        End If
    End If
    
    If SIM_INTEGRATED = True Then
        If legalnameSimilarity = 1 Or countrySimilarity = 1 Or citySimilarity = 1 Or addressSimilarity = 1 Then
            Cells(i, currentColumn).Value = "Ideal Match"
        End If
        If legalnameSimilarity <= 0.65 Or countrySimilarity <= 0.75 Or citySimilarity <= 0.75 Or addressSimilarity <= 0.66 Then
            Cells(i, currentColumn).Value = "Poor Match"
            If integratedSimilarity > 0.79 Then
                If legalnameSimilarity > 0.5 Then
                    If addressSimilarity > 0.5 Then
                        Cells(i, currentColumn).Value = "Partial Match"
                        If USE_GOOGLE_API Then
                            If GOOGLE_HAS_QUERY Then
                                GoogleStreet1 = standardizeAddress(keepString1)
                                GoogleStreet2 = standardizeAddress(keepString2)
                                addressSimilarity = Similarity(GoogleStreet1, GoogleStreet2)
                                newAddressSimilarity = Str(Round(addressSimilarity, 2))
                                If legalnameSimilarity = 1 Or countrySimilarity = 1 Or citySimilarity = 1 Or addressSimilarity = 1 Then
                                    Cells(i, currentColumn).Value = "Ideal Match (" & newAddressSimilarity & ")"
                                End If
                                If legalnameSimilarity <= 0.65 Or countrySimilarity <= 0.75 Or citySimilarity <= 0.75 Or addressSimilarity <= 0.66 Then
                                    Cells(i, currentColumn).Value = "Poor Match (" & newAddressSimilarity & ")"
                                    If integratedSimilarity > 0.79 Then
                                        If legalnameSimilarity > 0.5 Then
                                            If addressSimilarity > 0.5 Then
                                                Cells(i, currentColumn).Value = "Partial Match (" & newAddressSimilarity & ")"
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            If legalnameSimilarity > 0.5 Then
                If countrySimilarity = 1 Then
                    If citySimilarity = 1 Then
                        If addressSimilarity = 1 Then
                            Cells(i, currentColumn).Value = "Partial Match (Check Legal Name)"
                        End If
                    End If
                End If
            End If
            If legalnameSimilarity = 1 Then
                If countrySimilarity = 1 Then
                    If citySimilarity = 1 Then
                        If addressSimilarity > 0.5 Then
                            Cells(i, currentColumn).Value = "Partial Match (Check Address)"
                        End If
                    End If
                End If
            End If
            If legalnameSimilarity = 1 Then
                If addressSimilarity > 0.85 Then
                    If countrySimilarity = 1 Then
                        If citySimilarity <> 1 Then
                            Cells(i, currentColumn).Value = "Partial Match (Check City)"
                        End If
                    End If
                    If countrySimilarity <> 1 Then
                        If citySimilarity = 1 Then
                            Cells(i, currentColumn).Value = "Partial Match (Check Country)"
                        End If
                    End If
                End If
            End If
        End If
        currentColumn = currentColumn + 1
    End If
    
    String1 = ""
    String2 = ""
    wholeString1 = ""
    wholeString2 = ""
    keepString1 = ""
    keepString2 = ""
    GoogleStreet1 = ""
    GoogleStreet2 = ""
    
    pctDone = i / countRowNum * 100
    If pctDone > 99 Then
        pctDone = 99
    End If
    With Progress
        .theFrameProgress.Caption = "Calculating Data Similarity. Complete: " & pctDone & "%"
        .theLabelProgress.Width = pctDone * 2.4
        DoEvents
    End With
End If
Next i

Call DeleteDictionary
Worksheets("Matching").UsedRange.AutoFilter
End Sub

Function Similarity(ByVal String1 As String, _
    ByVal String2 As String, _
    Optional ByRef RetMatch As String, _
    Optional min_match = 1) As Single
Dim b1() As Byte, b2() As Byte
Dim lngLen1 As Long, lngLen2 As Long
Dim lngResult As Long


If UCase(String1) = UCase(String2) Then
    Similarity = 1
Else:
    lngLen1 = Len(String1)
    lngLen2 = Len(String2)
    If (lngLen1 = 0) Or (lngLen2 = 0) Then
        Similarity = 0
    Else:
        b1() = StrConv(UCase(String1), vbFromUnicode)
        b2() = StrConv(UCase(String2), vbFromUnicode)
        lngResult = Similarity_sub(0, lngLen1 - 1, _
        0, lngLen2 - 1, _
        b1, b2, _
        String1, _
        RetMatch, _
        min_match)
        Erase b1
        Erase b2
        If lngLen1 >= lngLen2 Then
            Similarity = lngResult / lngLen1
        Else
            Similarity = lngResult / lngLen2
        End If
    End If
End If

End Function

Private Function Similarity_sub(ByVal start1 As Long, ByVal end1 As Long, _
                                ByVal start2 As Long, ByVal end2 As Long, _
                                ByRef b1() As Byte, ByRef b2() As Byte, _
                                ByVal FirstString As String, _
                                ByRef RetMatch As String, _
                                ByVal min_match As Long, _
                                Optional recur_level As Integer = 0) As Long
'* CALLED BY: Similarity *(RECURSIVE)

Dim lngCurr1 As Long, lngCurr2 As Long
Dim lngMatchAt1 As Long, lngMatchAt2 As Long
Dim i As Long
Dim lngLongestMatch As Long, lngLocalLongestMatch As Long
Dim strRetMatch1 As String, strRetMatch2 As String

If (start1 > end1) Or (start1 < 0) Or (end1 - start1 + 1 < min_match) _
Or (start2 > end2) Or (start2 < 0) Or (end2 - start2 + 1 < min_match) Then
    Exit Function '(exit if start/end is out of string, or length is too short)
End If

For lngCurr1 = start1 To end1
    For lngCurr2 = start2 To end2
        i = 0
        Do Until b1(lngCurr1 + i) <> b2(lngCurr2 + i)
            i = i + 1
            If i > lngLongestMatch Then
                lngMatchAt1 = lngCurr1
                lngMatchAt2 = lngCurr2
                lngLongestMatch = i
            End If
            If (lngCurr1 + i) > end1 Or (lngCurr2 + i) > end2 Then Exit Do
        Loop
    Next lngCurr2
Next lngCurr1

If lngLongestMatch < min_match Then Exit Function

lngLocalLongestMatch = lngLongestMatch
RetMatch = ""

lngLongestMatch = lngLongestMatch _
+ Similarity_sub(start1, lngMatchAt1 - 1, _
start2, lngMatchAt2 - 1, _
b1, b2, _
FirstString, _
strRetMatch1, _
min_match, _
recur_level + 1)
If strRetMatch1 <> "" Then
    RetMatch = RetMatch & strRetMatch1 & "*"
Else
    RetMatch = RetMatch & IIf(recur_level = 0 _
    And lngLocalLongestMatch > 0 _
    And (lngMatchAt1 > 1 Or lngMatchAt2 > 1) _
    , "*", "")
End If


RetMatch = RetMatch & Mid$(FirstString, lngMatchAt1 + 1, lngLocalLongestMatch)


lngLongestMatch = lngLongestMatch _
+ Similarity_sub(lngMatchAt1 + lngLocalLongestMatch, end1, _
lngMatchAt2 + lngLocalLongestMatch, end2, _
b1, b2, _
FirstString, _
strRetMatch2, _
min_match, _
recur_level + 1)

If strRetMatch2 <> "" Then
    RetMatch = RetMatch & "*" & strRetMatch2
Else
    RetMatch = RetMatch & IIf(recur_level = 0 _
    And lngLocalLongestMatch > 0 _
    And ((lngMatchAt1 + lngLocalLongestMatch < end1) _
    Or (lngMatchAt2 + lngLocalLongestMatch < end2)) _
    , "*", "")
End If

Similarity_sub = lngLongestMatch

End Function

Function CleanCompanyName(strSource As String) As String
    Dim i, j As Long
    Dim strResult, originalResult As String
    Dim characterToStandardize As Long
    
    For i = 1 To Len(strSource)
        If ((Asc(Mid(strSource, i, 1)) >= 48 And Asc(Mid(strSource, i, 1)) <= 57) Or (Asc(Mid(strSource, i, 1)) >= 65 And Asc(Mid(strSource, i, 1)) <= 90) Or (Asc(Mid(strSource, i, 1)) >= 97 And Asc(Mid(strSource, i, 1)) <= 122) Or Asc(Mid(strSource, i, 1)) = 38) = False Then
            strResult = strResult & " "
        Else
            strResult = strResult & Mid(strSource, i, 1)
        End If
    Next i
        
    'Concatenate leading consecutive single letters
    strResult = Trim(strResult)
    strResult = strResult & " "
    originalResult = strResult
    For i = 1 To Len(originalResult)
    If Mid(originalResult, i, 1) <> " " Then
        If Mid(originalResult, i + 1, 1) <> " " Then
            Exit For
        End If
        j = i
        While Mid(originalResult, j + 1, 1) = " " And Mid(originalResult, j + 2, 1) <> " " And Mid(originalResult, j + 3, 1) = " "
        strResult = Replace(strResult, " ", "", i, 1, 1)
        j = j + 2
        Wend
        If Len(originalResult) <> Len(strResult) Then
            Exit For
        End If
    End If
    Next i
        
    'Remove trailing corporate designations
    strResult = Trim(strResult)
    strResult = " " & strResult & " "
    
    '" & " and " AND "
    While InStr(1, strResult, " & ", 1) <> 0
    strResult = Replace(strResult, " & ", " ", , , 1)
    Wend
    strResult = Replace(strResult, "& ", "&", , , 1)
    
    While InStr(1, strResult, " AND ", 1) <> 0
    strResult = Replace(strResult, " AND ", " ", , , 1)
    Wend
     
    While InStr(1, strResult, " THE ", 1) <> 0
    strResult = Replace(strResult, " THE ", " ", , , 1)
    Wend
    
    strResult = Replace(strResult, " CO ", " ", , , 1)
    strResult = Replace(strResult, " LTD ", " ", , , 1)
    strResult = Replace(strResult, " LIMITED ", " ", , , 1)
    strResult = Replace(strResult, " INC ", " ", , , 1)
    strResult = Replace(strResult, " LLC ", " ", , , 1)
    strResult = Replace(strResult, " PTY ", " ", , , 1)
    strResult = Replace(strResult, " KG ", " ", , , 1)
    strResult = Replace(strResult, " GMBH ", " ", , , 1)
    strResult = Replace(strResult, " PTE ", " ", , , 1)
    strResult = Replace(strResult, " AS ", " ", , , 1)
    strResult = Replace(strResult, " CORPORATION ", " ", , , 1)
    strResult = Replace(strResult, " COMPANY ", " ", , , 1)
    strResult = Replace(strResult, " CORP ", " ", , , 1)
    strResult = Replace(strResult, " AB ", " ")
    strResult = Replace(strResult, " DE C V ", " ", , , 1)
    strResult = Replace(strResult, " SA ", " ", , , 1)
    strResult = Replace(strResult, " SRL ", " ", , , 1)
    strResult = Replace(strResult, " S P A ", " ", , , 1)
    strResult = Replace(strResult, " SPA ", " ", , , 1)
    strResult = Replace(strResult, " AG ", " ", , , 1)
    strResult = Replace(strResult, " GROUP ", " ", , , 1)
    strResult = Replace(strResult, " SP ", " ", , , 1)
    strResult = Replace(strResult, " Z O O ", " ", , , 1)
    strResult = Replace(strResult, " S L ", " ", , , 1)
    strResult = Replace(strResult, " ZOO ", " ", , , 1)
    strResult = Replace(strResult, " S R L ", " ", , , 1)
    strResult = Replace(strResult, " S A ", " ", , , 1)
    strResult = Replace(strResult, " A S ", " ", , , 1)
    strResult = Replace(strResult, " LT ", " ", , , 1)
    strResult = Replace(strResult, " EMS ", " ", , , 1)
    strResult = Replace(strResult, " SE ", " ", , , 1)
    
    'N/A Unknown
    If InStr(1, strResult, " NA ", 1) <> 0 Or InStr(1, strResult, " N A ", 1) <> 0 Or InStr(1, strResult, " UNKNOWN ", 1) <> 0 Then
        strResult = ""
    End If
    
    'Standardize certain common words - ELECTR
    For i = 0 To UBound(theElectronic)
    characterToStandardize = 0
    characterToStandardize = InStr(1, strResult, theElectronic(i), 1)
    While characterToStandardize <> 0
        strResult = Replace(strResult, theElectronic(i), " ", , , 1)
        characterToStandardize = 0
        characterToStandardize = InStr(1, strResult, theElectronic(i), 1)
    Wend
    Next i
    
    'Standardize certain common words - Tech
    For i = 0 To UBound(theTechnology)
    characterToStandardize = 0
    characterToStandardize = InStr(1, strResult, theTechnology(i), 1)
    While characterToStandardize <> 0
        strResult = Replace(strResult, theTechnology(i), " ", , , 1)
        characterToStandardize = 0
        characterToStandardize = InStr(1, strResult, theTechnology(i), 1)
    Wend
    Next i
    
    'Standardize certain common words - SYS
    For i = 0 To UBound(theSystem)
    characterToStandardize = 0
    characterToStandardize = InStr(1, strResult, theSystem(i), 1)
    While characterToStandardize <> 0
        strResult = Replace(strResult, theSystem(i), " ", , , 1)
        characterToStandardize = 0
        characterToStandardize = InStr(1, strResult, theSystem(i), 1)
    Wend
    Next i

    'Standardize certain common words - SCI
    For i = 0 To UBound(theScience)
    characterToStandardize = 0
    characterToStandardize = InStr(1, strResult, theScience(i), 1)
    While characterToStandardize <> 0
        strResult = Replace(strResult, theScience(i), " ", , , 1)
        characterToStandardize = 0
        characterToStandardize = InStr(1, strResult, theScience(i), 1)
    Wend
    Next i
    
    'Standardize certain common words - ENG
    For i = 0 To UBound(theEngineer)
    characterToStandardize = 0
    characterToStandardize = InStr(1, strResult, theEngineer(i), 1)
    While characterToStandardize <> 0
        strResult = Replace(strResult, theEngineer(i), " ", , , 1)
        characterToStandardize = 0
        characterToStandardize = InStr(1, strResult, theEngineer(i), 1)
    Wend
    Next i
    
    'Standardize certain common words - AUTO
    For i = 0 To UBound(theAutomation)
    characterToStandardize = 0
    characterToStandardize = InStr(1, strResult, theAutomation(i), 1)
    While characterToStandardize <> 0
        strResult = Replace(strResult, theAutomation(i), " ", , , 1)
        characterToStandardize = 0
        characterToStandardize = InStr(1, strResult, theAutomation(i), 1)
    Wend
    Next i
    
    'Standardize certain common words - ENTERPRISE
    For i = 0 To UBound(theEnterprise)
    characterToStandardize = 0
    characterToStandardize = InStr(1, strResult, theEnterprise(i), 1)
    While characterToStandardize <> 0
        strResult = Replace(strResult, theEnterprise(i), " ", , , 1)
        characterToStandardize = 0
        characterToStandardize = InStr(1, strResult, theEnterprise(i), 1)
    Wend
    Next i
     
    ''''''''''''''''''''''''''''''
    While InStr(1, strResult, "  ", 1) <> 0
    strResult = Replace(strResult, "  ", " ", , , 1)
    Wend
    'While InStr(1, strResult, " ", 1) <> 0
    'strResult = Replace(strResult, " ", "", , , 1)
    'Wend
         
    CleanCompanyName = Trim(strResult)
End Function

Function CleanCountryName(strSource As String) As String
    Dim strResult As String
    
    strResult = strSource
    strResult = Replace(strResult, "Korea, Republic of", "South Korea", , , 1)
    strResult = Replace(strResult, "RUSSIAN FEDERATION", "RUSSIA", , , 1)
    strResult = Replace(strResult, "Vietnam", "Viet Nam", , , 1)
    strResult = Replace(strResult, "Virgin Islands, British", "Virgin Islands (British)", , , 1)
    strResult = Replace(strResult, "Korea, Republic of", "Republic of Korea", , , 1)
    strResult = Replace(strResult, "USA", "United States", , , 1)
    strResult = Replace(strResult, "Macedonia, The Former Yugoslav Republic of", "Macedonia", , , 1)
    strResult = Replace(strResult, "United Kingdom", "England", , , 1)
    strResult = Replace(strResult, "The  Netherlands", "Netherlands", , , 1)
    CleanCountryName = strResult
End Function

Function CleanCityName(strSource As String) As String
    Dim strResult As String
    strResult = strSource
    strResult = Replace(strResult, "96317 Kronach", "Kronach", , , 1)
    strResult = Replace(strResult, "A.BANGPAKONG, CHACHOENGSAO", "BANG PAKONG", , , 1)
    strResult = Replace(strResult, "AC Amsterdam", "Amsterdam", , , 1)
    strResult = Replace(strResult, "AIR PORT CITY", "Airport City", , , 1)
    strResult = Replace(strResult, "Aki-gun, Hiroshima", "AKI-GUN", , , 1)
    strResult = Replace(strResult, "Alcobendas Madrid", "ALCOBENDAS", , , 1)
    strResult = Replace(strResult, "Taoyuan City", "Taoyuan", , , 1)
    strResult = Replace(strResult, "Taipei City", "Taipei", , , 1)
    strResult = Replace(strResult, "Taichung City", "Taichung", , , 1)
    strResult = Replace(strResult, "Sha Tin New Territories", "SHA TIN", , , 1)
    strResult = Replace(strResult, "Road Town Tortola", "ROAD TOWN", , , 1)
    strResult = Replace(strResult, "Kwun Tong Kowloon", "KWUN TONG", , , 1)
    strResult = Replace(strResult, "Kowloon Bay Kowloon", "KOWLOON BAY", , , 1)
    strResult = Replace(strResult, "Hsinchu City", "Hsinchu", , , 1)
    strResult = Replace(strResult, "Bangalore", "Bengaluru", , , 1)
    strResult = Replace(strResult, "Hsinchu City", "Hsinchu", , , 1)
    strResult = Replace(strResult, "Rome", "Roma", , , 1)
    strResult = Replace(strResult, "Saint", "St", , , 1)
    strResult = Replace(strResult, "City", "", , , 1)
    strResult = Replace(strResult, "-", " ", , , 1)
    CleanCityName = strResult
End Function

Sub OptimizeCode_Begin()

Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual
ActiveSheet.DisplayPageBreaks = False

End Sub

Sub OptimizeCode_End()

Application.ScreenUpdating = True
Application.EnableEvents = True
Application.Calculation = xlCalculationAutomatic
ActiveSheet.DisplayPageBreaks = True

End Sub

Function standardizeAddress(ByVal address As String)
Dim Web, Status As String
Dim ie As Object
Dim result As String

If GOOGLE_HAS_QUERY = False Then
    standardizeAddress = address
    Exit Function
End If

Web = "https://maps.googleapis.com/maps/api/geocode/xml?address=" & address & "&sensor=false&language=en"
Set ie = CreateObject("InternetExplorer.Application")
On Error Resume Next

With ie
.Visible = False
.Navigate (Web)
Do While ie.Busy Or ie.readyState <> 4: DoEvents: Loop
Sleep 100
Status = Tuning(ie.document.getElementsByTagName("status")(0).innertext, "status")
If Status = "OVER_QUERY_LIMIT" Or Status = "REQUEST_DENIED" Then
    Web = "https://maps.googleapis.com/maps/api/geocode/xml?address=" & address & "&sensor=false&language=en&key=AIzaSyCfsIO5xJOKtl1_QF3NXrcJSAAyf6FgIiE"
    ie.Navigate (Web)
    Do While ie.Busy Or ie.readyState <> 4: DoEvents: Loop
    QUERY_USED = QUERY_USED + 1
    Sleep 100
    Status = Tuning(ie.document.getElementsByTagName("status")(0).innertext, "status")
    If Status = "OVER_QUERY_LIMIT" Or Status = "REQUEST_DENIED" Then
        GOOGLE_HAS_QUERY = False
        MsgBox "Query limit has all been used, please purchase more."
        standardizeAddress = address
        Exit Function
    ElseIf Status = "ZERO_RESULTS" Then
        standardizeAddress = address
        Exit Function
    End If
ElseIf Status = "ZERO_RESULTS" Then
    standardizeAddress = address
    Exit Function
End If

Dim Eachformatted_address As Object
Set Eachformatted_address = ie.document.getElementsByTagName("formatted_address")
result = Eachformatted_address(0).innertext
result = Tuning(result, "formatted_address")
Set Eachformatted_address = Nothing
End With

On Error GoTo 0
ie.Quit
Set ie = Nothing

Sleep 100
standardizeAddress = result
End Function

Function Tuning(ByVal b As String, ByVal a As String) As String
Dim c, d As String
c = "</" & a & ">"
b = Replace(b, c, "")

d = "<" & a & "/>"
b = Replace(b, d, "")


a = "<" & a & ">"
b = Replace(b, a, "")
Tuning = b
End Function

Function IsTheSame(ByVal String1 As String, ByVal String2 As String) As Boolean
Dim theSame As Boolean
theSame = False
If theSame = False And InStr(1, String1, String2, vbTextCompare) > 0 Then
    theSame = True
End If
If theSame = False And InStr(1, String2, String1, vbTextCompare) > 0 Then
    theSame = True
End If

If theSame Then
    IsTheSame = True
Else
    IsTheSame = False
End If
End Function


Function GetStandardizedAddress(ByVal address As String)
Dim Web, Status As String
Dim ie As Object
Dim result As String

Web = "https://maps.googleapis.com/maps/api/geocode/xml?address=" & address & "&sensor=false&language=en"
Set ie = CreateObject("InternetExplorer.Application")
On Error Resume Next

With ie
.Visible = False
.Navigate (Web)
Do While ie.Busy Or ie.readyState <> 4: DoEvents: Loop
Sleep 100
Status = Tuning(ie.document.getElementsByTagName("status")(0).innertext, "status")
If Status = "OVER_QUERY_LIMIT" Or Status = "REQUEST_DENIED" Then
    Web = "https://maps.googleapis.com/maps/api/geocode/xml?address=" & address & "&sensor=false&language=en&key=AIzaSyCfsIO5xJOKtl1_QF3NXrcJSAAyf6FgIiE"
    ie.Navigate (Web)
    Do While ie.Busy Or ie.readyState <> 4: DoEvents: Loop
    Sleep 100
    Status = Tuning(ie.document.getElementsByTagName("status")(0).innertext, "status")
    If Status = "OVER_QUERY_LIMIT" Or Status = "REQUEST_DENIED" Then
        GetStandardizedAddress = "OVER_QUERY_LIMIT"
        Exit Function
    ElseIf Status = "ZERO_RESULTS" Then
        GetStandardizedAddress = "ZERO_RESULTS"
        Exit Function
    End If
ElseIf Status = "ZERO_RESULTS" Then
    GetStandardizedAddress = "ZERO_RESULTS"
    Exit Function
End If

Dim Eachformatted_address As Object
Set Eachformatted_address = ie.document.getElementsByTagName("formatted_address")
result = Eachformatted_address(0).innertext
result = Tuning(result, "formatted_address")
Set Eachformatted_address = Nothing
End With

On Error GoTo 0
ie.Quit
Set ie = Nothing

GetStandardizedAddress = result
End Function

