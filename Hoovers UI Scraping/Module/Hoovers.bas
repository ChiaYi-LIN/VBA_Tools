Attribute VB_Name = "Hoovers"
#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr) 'For 64 Bit Systems
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'For 32 Bit Systems
#End If

Public Const SEARCH_NAME = 2
Public Const SEARCH_COUNTRY = 3
Public Const SEARCH_CITY = 4
Public Const SEARCH_DUNS = 5
Public Const RESULT_NAME = 6
Public Const RESULT_DUNS = 7
Public Const RESULT_COUNTRY = 8
Public Const RESULT_STATE = 9
Public Const RESULT_CITY = 10
Public Const RESULT_STREET = 11
Public Const RESULT_EXTEND_STREET = 12
Public Const RESULT_ZIP = 13
Public Const RESULT_FULL_ADDRESS = 14
Public Const RESULT_LOCATION_TYPE = 15
Public Const RESULT_UPNAME = 16
Public Const RESULT_UPDUNS = 17
Public Const RESULT_WEBSITE = 18
Public Const RESULT_COMMENT = 19
Public HAVE_DUNS As Boolean
Public START_ROW As Long
Public SELECT_RANGE As Range
Public LOGIN_ACCOUNT, LOGIN_PASSWORD As String
Public REMEMBER_ACCOUNT, REMEMBER_PASSWORD As String


Sub HeadersAndFormat()
With Worksheets("DUNS")
    .Cells(1, SEARCH_NAME).Value = "Legal Name"
    .Cells(1, SEARCH_COUNTRY).Value = "Country"
    .Cells(1, SEARCH_CITY).Value = "City"
    .Cells(1, SEARCH_DUNS).Value = "DUNS"
    .Cells(1, RESULT_NAME).Value = "Legal Name"
    .Cells(1, RESULT_DUNS).Value = "DUNS"
    .Cells(1, RESULT_COUNTRY).Value = "Country"
    .Cells(1, RESULT_STATE).Value = "State"
    .Cells(1, RESULT_CITY).Value = "City"
    .Cells(1, RESULT_STREET).Value = "Street 1"
    .Cells(1, RESULT_EXTEND_STREET).Value = "Street 2"
    .Cells(1, RESULT_ZIP).Value = "Zip"
    .Cells(1, RESULT_FULL_ADDRESS) = "Full Address"
    .Cells(1, RESULT_LOCATION_TYPE).Value = "Location Type"
    .Cells(1, RESULT_UPNAME).Value = "Ultimate Parent Name"
    .Cells(1, RESULT_UPDUNS).Value = "Ultimate Parent DUNS"
    .Cells(1, RESULT_WEBSITE).Value = "Website"
    .Cells(1, RESULT_COMMENT).Value = "Comment"
    
    .Columns(SEARCH_DUNS).NumberFormat = "@"
    .Columns(RESULT_DUNS).NumberFormat = "@"
End With
End Sub
'''Not Being Used
Sub loginDNBHoovers()
  Dim IE As Object
  Dim HTMLDoc As Object
  Dim objCollection As Object
  Dim DunsNumber, checkDuns, checkCompanyName As String
  Dim i, j, countRowNum, numberOfResults, currentColumn As Long
  Dim allCompanyResults, eachCompanyResult, companyName, companyDuns, overView As Object
  Dim viewCompanyName, viewCompanyCoun, viewCompanyCity, viewCompanyStreet As Object
  Dim success As Boolean
  Dim getFullAddress, eachLevelAddress, getLocationType, getKeyInfo, getUltimateParent, getUPName, getUPDuns, getWebsite As Object
  Dim checkUPName, checkUPDuns, checkExtendStreet As Object
  Dim PctDone As Long
  
  On Error Resume Next
  Set IE = Nothing
  On Error GoTo 0
  
  Worksheets("DUNS").Activate
  countRowNum = Cells(1, 1).CurrentRegion.Rows.Count
  
  Set IE = CreateObject("InternetExplorer.Application")
  IE.Visible = False
  IE.Navigate "https://app.avention.com/login?F1330836414546BWPUHR=_"
  
  Do While IE.Busy Or IE.ReadyState <> 4: Loop
  Sleep 100
  
  Set HTMLDoc = IE.Document
  Set objCollection = IE.Document.getElementById("remember")
  objCollection.Click
  HTMLDoc.getElementById("username").innertext = ""
  Set objCollection = IE.Document.getElementById("login")
  objCollection.getElementsByTagname("button")(0).Click
  
  Do While IE.Busy Or IE.ReadyState <> 4: Loop
  Sleep 1000
  HTMLDoc.getElementById("password").innertext = ""
  
  Set objCollection = IE.Document.getElementById("login")
  objCollection.getElementsByTagname("button")(0).Click

  Do While IE.Busy Or IE.ReadyState <> 4: Loop
  Sleep 5000
  
  If HAVE_DUNS Then
    HTMLDoc.getElementById("search").innertext = Cells(i, SEARCH_DUNS).Value
  Else
    MsgBox InStr(InStr(1, HTMLDoc.getElementsByTagname("form")(0).innerhtml, Chr(34) & "search" & Chr(34), vbTextCompare), HTMLDoc.getElementsByTagname("form")(0).innerhtml, "value=" & Chr(34), vbTextCompare)
    HTMLDoc.getElementsByTagname("form")(0).innerhtml = Mid(HTMLDoc.getElementsByTagname("form")(0).innerhtml, 1, InStr(InStr(1, HTMLDoc.getElementsByTagname("form")(0).innerhtml, Chr(34) & "search" & Chr(34), vbTextCompare), HTMLDoc.getElementsByTagname("form")(0).innerhtml, "value=" & Chr(34), vbTextCompare) + 6) & "NXP.com" & Mid(Source, i + 1, Len(Source) - i)
    HTMLDoc.getElementById("search").innertext = Cells(i, SEARCH_NAME).Value & " " & CountryNameTuning(Cells(i, SEARCH_COUNTRY).Value) & " " & Cells(i, SEARCH_CITY).Value
  End If
  
  
  If HAVE_DUNS Then
  For i = 2 To countRowNum
    Cells(i, SEARCH_DUNS).Value = DunsNumberDigitToNine(Cells(i, SEARCH_DUNS).Value)
  Next i
  End If
  
  On Error Resume Next
  Set allCompanyResults = Nothing
  Set eachCompanyResult = Nothing
  Set companyName = Nothing
  Set companyDuns = Nothing
  On Error GoTo 0
  
  IE.Quit
  Set IE = Nothing
End Sub
'''Usable
Sub loginHoovers()
  
  'TEST
  'HAVE_DUNS = True
  'START_ROW = 2
  'LOGIN_ACCOUNT = ""
  'LOGIN_PASSWORD = ""
  '
  
  Dim IE As Object
  Dim HTMLDoc As Object
  Dim objCollection As Object
  Dim DunsNumber, checkDuns, checkCompanyName As String
  Dim i, j, countRowNum, numberOfResults, currentColumn As Long
  Dim allCompanyResults, eachCompanyResult, companyName, companyDuns, overView As Object
  Dim viewCompanyName, viewCompanyCoun, viewCompanyCity, viewCompanyStreet As Object
  Dim success As Boolean
  Dim getFullAddress, eachLevelAddress, getLocationType, getKeyInfo, getUltimateParent, getUPName, getUPDuns, getWebsite As Object
  Dim checkUPName, checkUPDuns, checkExtendStreet, checkLoginError As Object
  Dim PctDone As Long
  
  On Error Resume Next
  Set IE = Nothing
  On Error GoTo 0
  
  Worksheets("DUNS").Activate
  countRowNum = Cells(START_ROW, SEARCH_NAME).CurrentRegion.Rows.Count
  
  Set IE = CreateObject("InternetExplorer.Application")
  IE.Visible = True
  IE.Navigate "https://subscriber.hoovers.com/H/login/login.html"
  
  Do While IE.Busy Or IE.ReadyState <> 4: Loop
  Sleep 100
  
  Set HTMLDoc = IE.Document

  With HTMLDoc
  HTMLDoc.getElementById("j_username").Value = LOGIN_ACCOUNT
  HTMLDoc.getElementById("j_password").Value = LOGIN_PASSWORD

  End With

  Set objCollection = IE.Document.getElementById("j_submit")
  objCollection.Click
  
  Do While IE.Busy Or IE.ReadyState <> 4: Loop
  Sleep 5000
  
  On Error Resume Next
  Set checkLoginError = HTMLDoc.getElementById("loginError").getElementsByTagname("p")
  On Error GoTo 0
  
  If Not checkLoginError Is Nothing Then
  If Trim(checkLoginError(0).innertext) = "The username or password you entered is incorrect." Then
    MsgBox "Login failed."
    Exit Sub
  ElseIf InStr(1, checkLoginError(0).innertext, "Your account has been locked.", vbTextCompare) > 0 Then
    MsgBox "Your account has been locked."
    Exit Sub
  End If
  End If
  
  Set objCollection = IE.Document.getElementsByClassname("ui-button ui-widget ui-state-default ui-corner-all ui-button-text-only")
  objCollection(1).Click
  
  If HAVE_DUNS Then
  For i = 2 To countRowNum
    Cells(i, SEARCH_DUNS).Value = DunsNumberDigitToNine(Cells(i, SEARCH_DUNS).Value)
  Next i
  End If
  
  For i = START_ROW To countRowNum
  On Error Resume Next
  Set allCompanyResults = Nothing
  Set eachCompanyResult = Nothing
  Set companyName = Nothing
  Set companyDuns = Nothing
  On Error GoTo 0
  
  Do While IE.Busy Or IE.ReadyState <> 4: Loop
  Sleep 100
  
  If HAVE_DUNS Then
    HTMLDoc.getElementById("searchField").Value = Cells(i, SEARCH_DUNS).Value
  Else
    HTMLDoc.getElementById("searchField").Value = Cells(i, SEARCH_NAME).Value & " " & CountryNameTuning(Cells(i, SEARCH_COUNTRY).Value) & " " & Cells(i, SEARCH_CITY).Value
  End If
  
  Set objCollection = IE.Document.getElementById("btnSearch")
  objCollection.Click
  
  Do While IE.Busy Or IE.ReadyState <> 4: Loop
  Sleep 1000
  Set allCompanyResults = IE.Document.getElementsByClassname("component categories clearBoth")
  
  On Error Resume Next
  Set eachCompanyResult = allCompanyResults(0).getElementsByTagname("p")
  On Error GoTo 0
  
  If Not eachCompanyResult(0) Is Nothing Then
    If Trim(eachCompanyResult(0).innertext) = "No company results found." Then
        Cells(i, RESULT_NAME).Value = "No company results found."
    End If
  End If
   
  If eachCompanyResult(0) Is Nothing Then
  Set eachCompanyResult = allCompanyResults(0).getElementsByTagname("tbody")(0).getElementsByTagname("tr")
  numberOfResults = eachCompanyResult.Length - 1
  
  For j = 0 To numberOfResults
    success = False
    
    Set companyDuns = eachCompanyResult(j).getElementsByTagname("td")
    checkDuns = DunsNumberTuning(companyDuns(0).innertext)
    
    If checkDuns = Cells(i, SEARCH_DUNS) Or Not HAVE_DUNS Then
    
    Set companyName = eachCompanyResult(j).getElementsByTagname("a")
    checkCompanyName = companyName(0).innertext
    If checkCompanyName = "Nonmarketable" Then
        Set overView = companyName(1)
        checkCompanyName = overView.innertext
        Cells(i, RESULT_NAME).Value = checkCompanyName
        Cells(i, RESULT_COMMENT).Value = "Nonmarketable"
    ElseIf checkCompanyName = "Out of Business" Then
        Set overView = companyName(1)
        checkCompanyName = overView.innertext
        Cells(i, RESULT_NAME).Value = checkCompanyName
        Cells(i, RESULT_COMMENT).Value = "Out of Business"
    ElseIf Trim(checkCompanyName) = "" Then
        Set overView = companyName(1)
        checkCompanyName = overView.innertext
        Cells(i, RESULT_NAME).Value = checkCompanyName
    Else
        Set overView = companyName(0)
        Cells(i, RESULT_NAME).Value = checkCompanyName
    End If
    
    Cells(i, RESULT_DUNS).Value = checkDuns
      
    success = True
    
    Else
        Cells(i, RESULT_NAME).Value = "No company results found."
    End If
    
    checkCompanyName = ""
    checkDuns = ""
    Set companyName = Nothing
    Set companyDuns = Nothing
    
    '
    'Query Company Information
    '
    '
    If success Then
        overView.Click
        Do While IE.Busy Or IE.ReadyState <> 4: Loop
        Sleep 1000

        If Trim(IE.Document.getElementById("adr")) <> "" Then
        Set getFullAddress = IE.Document.getElementById("adr")
        Set getLocationType = IE.Document.getElementById("companyLocationType")
        Set checkExtendStreet = getFullAddress.getElementsByClassname("extended-address")
        Set eachLevelAddress = getFullAddress.getElementsByTagname("span")
        Set getKeyInfo = IE.Document.getElementById("kInfo")
        Set getUltimateParent = getKeyInfo.getElementsByTagname("tr")
        Set checkUPName = getUltimateParent(5).getElementsByTagname("th")
        Set getUPName = getUltimateParent(5).getElementsByTagname("a")
        Set checkUPDuns = getUltimateParent(6).getElementsByTagname("th")
        Set getUPDuns = getUltimateParent(6).getElementsByTagname("td")
        Set getWebsite = IE.Document.getElementsByClassname("url ext")
        
        If Not getFullAddress.getElementsByClassname("street-address")(0) Is Nothing Then
        If Trim(getFullAddress.getElementsByClassname("street-address")(0).innertext) <> "" Then
        Cells(i, RESULT_STREET).Value = getFullAddress.getElementsByClassname("street-address")(0).innertext
        End If
        End If
        
        If Not getFullAddress.getElementsByClassname("extended-address")(0) Is Nothing Then
        If Trim(getFullAddress.getElementsByClassname("extended-address")(0).innertext) <> "" Then
        Cells(i, RESULT_EXTEND_STREET).Value = getFullAddress.getElementsByClassname("extended-address")(0).innertext
        End If
        End If
        
        If Not getFullAddress.getElementsByClassname("locality")(0) Is Nothing Then
        If Trim(getFullAddress.getElementsByClassname("locality")(0).innertext) <> "" Then
        Cells(i, RESULT_CITY).Value = getFullAddress.getElementsByClassname("locality")(0).innertext
        End If
        End If
        
        If Not getFullAddress.getElementsByClassname("region")(0) Is Nothing Then
        If Trim(getFullAddress.getElementsByClassname("region")(0).innertext) <> "" Then
        Cells(i, RESULT_STATE).Value = getFullAddress.getElementsByClassname("region")(0).innertext
        End If
        End If
        
        If Not getFullAddress.getElementsByClassname("zip")(0) Is Nothing Then
        If Trim(getFullAddress.getElementsByClassname("zip")(0).innertext) <> "" Then
        Cells(i, RESULT_ZIP).Value = getFullAddress.getElementsByClassname("zip")(0).innertext
        End If
        End If
        
        If Not getFullAddress.getElementsByClassname("country-name")(0) Is Nothing Then
        If Trim(getFullAddress.getElementsByClassname("country-name")(0).innertext) <> "" Then
        Cells(i, RESULT_COUNTRY).Value = getFullAddress.getElementsByClassname("country-name")(0).innertext
        End If
        End If
        
        If Not getFullAddress Is Nothing Then
        If Trim(getFullAddress.innertext) <> "" Then
        Cells(i, RESULT_FULL_ADDRESS).Value = getFullAddress.innertext
        Cells(i, RESULT_FULL_ADDRESS).Value = Trim(Replace(Cells(i, RESULT_FULL_ADDRESS).Value, Chr(10), " ", , , vbTextCompare))
        Cells(i, RESULT_FULL_ADDRESS).Value = Trim(Replace(Cells(i, RESULT_FULL_ADDRESS).Value, "  ", " ", , , vbTextCompare))
        End If
        End If
        
        If Not getLocationType Is Nothing Then
        If Trim(getLocationType.innertext) <> "" Then
        Cells(i, RESULT_LOCATION_TYPE).Value = getLocationType.innertext
        End If
        End If
        
        If checkUPName(0).innertext = "Ultimate Parent" Then
        If Not getUPName Is Nothing Then
        Cells(i, RESULT_UPNAME).Value = getUPName(0).innertext
        End If
        End If
        
        If checkUPDuns(0).innertext = "Ultimate Parent D-U-N-S" Then
        If Not getUPDuns Is Nothing Then
        Cells(i, RESULT_UPDUNS).Value = getUPDuns(0).innertext
        End If
        End If
       
        If Not getWebsite(0) Is Nothing Then
        If Trim(getWebsite(0).innertext) <> "" Then
            Cells(i, RESULT_WEBSITE).Value = getWebsite(0).innertext
        End If
        End If
        
        Set overView = Nothing
        Set getFullAddress = Nothing
        Set eachLevelAddress = Nothing
        Set getLocationType = Nothing
        Set getKeyInfo = Nothing
        Set getUltimateParent = Nothing
        Set getUPName = Nothing
        Set getUPDuns = Nothing
        Set getWebsite = Nothing
        Set checkUPName = Nothing
        Set checkExtendStreet = Nothing
        
        End If
        Exit For
        
    End If
  Next j
  End If
  
  'If START_ROW = countRowNum Then
  '  PctDone = 100
  'Else
  '  PctDone = (i - START_ROW) * 100 / (countRowNum - START_ROW)
  'End If
  'With Progress
  '  .theLabelProgress.Width = PctDone * 2.4
  '  .theFrameProgress.Caption = "Complete: " & PctDone & "%"
  '  DoEvents
  'End With
  
  Next i
  
  Set objCollection = IE.Document.getElementById("logout")
  objCollection.Click
  Do While IE.Busy Or IE.ReadyState <> 4: Loop
  Sleep 1000
  IE.Quit
  Set IE = Nothing
  
  'Unload Progress
  Unload Settings
End Sub

Function DunsNumberTuning(ByVal theString As String)
    Dim result As String
    Dim startNumOfDuns As Integer
    
    startNumOfDuns = InStr(1, theString, "#", 1) + 2
    result = Mid(theString, startNumOfDuns, 9)
    
    DunsNumberTuning = result
End Function

Function DunsNumberDigitToNine(ByVal theString As String)
Dim stringDigit, needMoreDigit As Long
stringDigit = Len(theString)
needMoreDigit = 9 - stringDigit
theString = String(needMoreDigit, "0") & theString
DunsNumberDigitToNine = theString
End Function

Function CountryNameTuning(ByVal inputString As String)
    Dim needTune As Long
    
    needTune = 0
    
    needTune = InStr(1, inputString, "USA", vbTextCompare)
    If needTune > 1 Then
        CountryNameTuning = "United States"
        Exit Function
    Else
        CountryNameTuning = inputString
    End If
    
End Function
