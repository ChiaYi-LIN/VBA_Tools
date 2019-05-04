Attribute VB_Name = "AllFunctions"
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

'去除所有符號(除了"&")，並把所有空格變為單一空格
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
        strResult = Replace(strResult, theElectronic(i), " ELECTR ", , , 1)
        characterToStandardize = 0
        characterToStandardize = InStr(1, strResult, theElectronic(i), 1)
    Wend
    Next i
    
    'Standardize certain common words - Tech
    For i = 0 To UBound(theTechnology)
    characterToStandardize = 0
    characterToStandardize = InStr(1, strResult, theTechnology(i), 1)
    While characterToStandardize <> 0
        strResult = Replace(strResult, theTechnology(i), " TECH ", , , 1)
        characterToStandardize = 0
        characterToStandardize = InStr(1, strResult, theTechnology(i), 1)
    Wend
    Next i
    
    'Standardize certain common words - SYS
    For i = 0 To UBound(theSystem)
    characterToStandardize = 0
    characterToStandardize = InStr(1, strResult, theSystem(i), 1)
    While characterToStandardize <> 0
        strResult = Replace(strResult, theSystem(i), " SYS ", , , 1)
        characterToStandardize = 0
        characterToStandardize = InStr(1, strResult, theSystem(i), 1)
    Wend
    Next i

    'Standardize certain common words - SCI
    For i = 0 To UBound(theScience)
    characterToStandardize = 0
    characterToStandardize = InStr(1, strResult, theScience(i), 1)
    While characterToStandardize <> 0
        strResult = Replace(strResult, theScience(i), " SCI ", , , 1)
        characterToStandardize = 0
        characterToStandardize = InStr(1, strResult, theScience(i), 1)
    Wend
    Next i
    
    'Standardize certain common words - ENG
    For i = 0 To UBound(theEngineer)
    characterToStandardize = 0
    characterToStandardize = InStr(1, strResult, theEngineer(i), 1)
    While characterToStandardize <> 0
        strResult = Replace(strResult, theEngineer(i), " ENG ", , , 1)
        characterToStandardize = 0
        characterToStandardize = InStr(1, strResult, theEngineer(i), 1)
    Wend
    Next i
    
    'Standardize certain common words - AUTO
    For i = 0 To UBound(theAutomation)
    characterToStandardize = 0
    characterToStandardize = InStr(1, strResult, theAutomation(i), 1)
    While characterToStandardize <> 0
        strResult = Replace(strResult, theAutomation(i), " AUTO ", , , 1)
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
    strResult = Replace(strResult, "RUSSIAN FEDERATION", "RUSSIA", , , 1)
    CleanCountryName = strResult
End Function
'Export datas into Table
Sub ExportCleanedData(ByVal exportLocation As String)
Dim targetSheet As String
Dim tableNumber As String

targetSheet = "Model N Data"

Worksheets(targetSheet).Activate
On Error Resume Next
ActiveSheet.ListObjects("Model N Data Table").Range.Select
Selection.Delete

Worksheets("Data Cleaner").Activate
Worksheets("Data Cleaner").Range("B1").CurrentRegion.Select
Selection.Copy

Worksheets(targetSheet).Activate
ActiveSheet.Paste Destination:=Worksheets(targetSheet).Cells(1, 2)
ActiveSheet.ListObjects.Add(xlSrcRange, Range("B1").CurrentRegion, , xlYes).Name = "Model N Data Table"

Cells(1, 1).Select
Application.CutCopyMode = False
Worksheets("Data Cleaner").Activate
Cells(1, 1).Select
End Sub
'HIde the same results in Fuzzy Lookup
Sub dedup()
Dim countRowNum, countColNum, i, j, IDOneColumn, IDTwoColumn As Long
Dim customerIDOne, customerIDTwo, countryOne, countryTwo As Range
Dim pctDone As Long

On Error Resume Next
Set customerIDOne = Application.InputBox("Select the Header Cell", "First Customer OID Range Select", Type:=8)
If customerIDOne Is Nothing Then
    Unload Deduplicate
    Exit Sub
End If

Set countryOne = Application.InputBox("Select the header cell of location level (Country in general) which should be matched exactly. Those which are not matched exactly would be hidden. (If the location that should be matched exactly is City or State, then choose the corresponding cell)", "First Country Range Select", Type:=8)
If countryOne Is Nothing Then
    Unload Deduplicate
    Exit Sub
End If

Set customerIDTwo = Application.InputBox("Select the Header Cell", "Second Customer OID Range Select", Type:=8)
If customerIDTwo Is Nothing Then
    Unload Deduplicate
    Exit Sub
End If

Set countryTwo = Application.InputBox("Select the header cell of location level (Country in general) which should be matched exactly. Those which are not matched exactly would be hidden. (If the location that should be matched exactly is City or State, then choose the corresponding cell)", "Second Country Range Select", Type:=8)
If countryTwo Is Nothing Then
    Unload Deduplicate
    Exit Sub
End If

On Error GoTo 0
IDOneColumn = customerIDOne.Column
IDTwoColumn = customerIDTwo.Column

Call OptimizeCode_Begin
Worksheets("Fuzzy Lookup").Activate

countColNum = Worksheets("Fuzzy Lookup").Cells(1, 2).CurrentRegion.Columns.Count
countRowNum = Worksheets("Fuzzy Lookup").Cells(1, 2).CurrentRegion.Rows.Count

If Cells(1, 2).Value <> "" And Cells(1, countColNum).Value <> "Exact check" Then
Cells(1, countColNum + 1).Value = "Exact check"
For i = 2 To countRowNum
If Cells(i, IDOneColumn).Value = Cells(i, IDTwoColumn).Value Then
    Cells(i, countColNum + 1).Value = "True"
Else
    Cells(i, countColNum + 1).Value = "False"
End If
Next i


For i = 2 To countRowNum
'5296274 is green
'colorindex 37 is blue
If Cells(i, 1).EntireRow.Interior.Color <> 5296274 And Cells(i, countColNum + 1).Value = "False" Then
    If Cells(i, countryOne.Column).Value <> Cells(i, countryTwo.Column).Value Then
        Cells(i, 1).EntireRow.Interior.ColorIndex = 37
    Else
        For j = i + 1 To countRowNum
            On Error Resume Next
            If Cells(i, IDOneColumn).Value = Cells(j, IDTwoColumn).Value Then
                If Cells(i, IDTwoColumn).Value = Cells(j, IDOneColumn).Value Then
                    Cells(j, 1).EntireRow.Interior.Color = 5296274
                    Exit For
                End If
            End If
        Next j
    End If
End If

'progress bar
    pctDone = i / countRowNum * 100
    With Deduplicate
        .FrameProgressDup.Caption = "Complete: " & pctDone & "%"
        .LabelProgressDup.Width = pctDone * 2.4
        DoEvents
    End With
Next i

Cells(1, 1).CurrentRegion.Select
Selection.AutoFilter
With Cells(1, 1).CurrentRegion
    .AutoFilter Field:=countColNum + 1, Criteria1:="FALSE"
    .AutoFilter Field:=1, Operator:=xlFilterNoFill
End With
Cells(1, 1).Select

ElseIf Cells(1, countColNum).Value = "Exact check" Then
Cells(1, 1).CurrentRegion.Select
Selection.AutoFilter
With Cells(1, 1).CurrentRegion
    .AutoFilter Field:=countColNum, Criteria1:="FALSE"
    .AutoFilter Field:=1, Operator:=xlFilterNoFill
End With
Cells(1, 1).Select

End If

PROGRESS_BAR_MODE = 0
Unload Deduplicate
Call ResetButtonPosition
Call OptimizeCode_End

End Sub
'Export Fuzzy Lookup results to Master & Aliased
Sub DataProcessForMasterAliased()

Dim countRowNum, countColNum, i, j, CompanyOneColumn, CompanyTwoColumn As Long
Dim CompanyNameOne, CompanyNameTwo As Range
Dim pctDone As Long
Dim currentCompany As String


On Error Resume Next
Set CompanyNameOne = Application.InputBox("Select the Header Cell", "First Company Name Range Select", Type:=8)
If CompanyNameOne Is Nothing Then
    Unload Deduplicate
    Exit Sub
End If
Set CompanyNameTwo = Application.InputBox("Select the Header Cell", "Second Company Name Range Select", Type:=8)
If CompanyNameTwo Is Nothing Then
    Unload Deduplicate
    Exit Sub
End If
On Error GoTo 0
CompanyOneColumn = CompanyNameOne.Column
CompanyTwoColumn = CompanyNameTwo.Column

Call OptimizeCode_Begin

Worksheets("Fuzzy Lookup").Activate
countColNum = Worksheets("Fuzzy Lookup").Cells(1, 2).CurrentRegion.Columns.Count
countRowNum = Worksheets("Fuzzy Lookup").Cells(1, 2).CurrentRegion.Rows.Count

Range(Worksheets("Master & Aliased").Cells(1, CompanyOneColumn), Worksheets("Master & Aliased").Cells(1, CompanyTwoColumn - 1)).Value = _
Range(Worksheets("Fuzzy Lookup").Cells(1, CompanyOneColumn), Worksheets("Fuzzy Lookup").Cells(1, CompanyTwoColumn - 1)).Value
Worksheets("Master & Aliased").Cells(1, CompanyTwoColumn).Value = Worksheets("Fuzzy Lookup").Cells(1, countColNum - 1).Value

currentRow = 2
For i = 2 To countRowNum
    If Cells(i, countColNum) = "False" And Cells(i, 1).Interior.Color <> 5296274 And Cells(i, 1).EntireRow.Interior.ColorIndex <> 37 _
    And Cells(i, CompanyOneColumn).Value <> "" Then
    
    Range(Worksheets("Master & Aliased").Cells(currentRow, CompanyOneColumn), Worksheets("Master & Aliased").Cells(currentRow, CompanyTwoColumn - 1)).Value = _
    Range(Worksheets("Fuzzy Lookup").Cells(i, CompanyOneColumn), Worksheets("Fuzzy Lookup").Cells(i, CompanyTwoColumn - 1)).Value
    Worksheets("Master & Aliased").Cells(currentRow, CompanyTwoColumn).Value = Worksheets("Fuzzy Lookup").Cells(i, countColNum - 1).Value
    
    currentRow = currentRow + 1
    Worksheets("Master & Aliased").Cells(currentRow, CompanyOneColumn).Value = Worksheets("Fuzzy Lookup").Cells(i, CompanyOneColumn).Value
    Range(Worksheets("Master & Aliased").Cells(currentRow, CompanyOneColumn + 1), Worksheets("Master & Aliased").Cells(currentRow, CompanyTwoColumn)).Value = _
    Range(Worksheets("Fuzzy Lookup").Cells(i, CompanyTwoColumn + 1), Worksheets("Fuzzy Lookup").Cells(i, countColNum - 1)).Value
    'Show original company name
    'Worksheets("Master & Aliased").Cells(currentRow, countColNum - CompanyTwoColumn + 1).Value = Worksheets("Fuzzy Lookup").Cells(i, CompanyTwoColumn).Value
    currentRow = currentRow + 1
    End If
    
    'progress bar
    pctDone = i / countRowNum * 100
    With Deduplicate
        .FrameProgressDup.Caption = "Complete: " & pctDone & "%"
        .LabelProgressDup.Width = pctDone * 2.4
        DoEvents
    End With
Next i

PROGRESS_BAR_MODE = 0
Unload Deduplicate
Call ResetButtonPosition
Call OptimizeCode_End
End Sub
'Delete dupicated ID and invalid company names
Sub DeleteDupDataByID()
Dim countRowNum, countColNum, i, j, CompanyColumn, IDColumn, countIfDup As Long
Dim companyName, CustomerID As Range
Dim pctDone As Long
Dim currentCompany As String
Dim startRow, endRow As Long

On Error Resume Next
Set companyName = Application.InputBox("Select the Header Cell", "Company Name Range Select", Type:=8)
If companyName Is Nothing Then
    Unload Deduplicate
    Exit Sub
End If
Set CustomerID = Application.InputBox("Select the Header Cell", "Customer OID Range Select", Type:=8)
If CustomerID Is Nothing Then
    Unload Deduplicate
    Exit Sub
End If
On Error GoTo 0
CompanyColumn = companyName.Column
IDColumn = CustomerID.Column

Call OptimizeCode_Begin
Worksheets("Master & Aliased").Activate
countColNum = Worksheets("Master & Aliased").Cells(1, 2).CurrentRegion.Columns.Count
countRowNum = Worksheets("Master & Aliased").Cells(1, 2).CurrentRegion.Rows.Count

''''''''''''''''''''''''''''''''''''''''''''''''
For i = countRowNum To 2 Step -1
    countIfDup = Application.WorksheetFunction.CountIf(Range(Cells(1, IDColumn), Cells(countRowNum, IDColumn)), Cells(i, IDColumn).Value)
    If countIfDup > 1 Then
        Cells(i, 1).EntireRow.Delete
    End If
    
    'progress bar
    pctDone = (countRowNum - i) / (countRowNum - 2) * 100
    With Deduplicate
        .FrameProgressDup.Caption = "Complete(1/2): " & pctDone & "%"
        .LabelProgressDup.Width = pctDone * 2.4
        DoEvents
    End With
Next i

countRowNum = Worksheets("Master & Aliased").Cells(1, 2).CurrentRegion.Rows.Count
For i = countRowNum To 2 Step -1
    countIfDup = Application.WorksheetFunction.CountIf(Range(Cells(1, CompanyColumn), Cells(countRowNum, CompanyColumn)), _
    Cells(i, CompanyColumn).Value)
    If countIfDup = 1 Then
        'Delete
        Cells(i, 1).EntireRow.Delete
        'Highlighten
        'Cells(i, 1).EntireRow.Interior.Color = 5296274
    Else
        Cells(i, 1).EntireRow.Interior.ColorIndex = xlNone
    End If
    
    'progress bar
    pctDone = (countRowNum - i) / (countRowNum - 2) * 100
    With Deduplicate
        .FrameProgressDup.Caption = "Complete(2/2): " & pctDone & "%"
        .LabelProgressDup.Width = pctDone * 2.4
        DoEvents
    End With
Next i
''''''''''''''''''''''''''''''''''''''''''''''''
'TESTING
'
'
'currentCompany = Cells(countRowNum, CompanyColumn).Value
'startRow = countRowNum
'endRow = countRowNum
'For i = countRowNum - 1 To 1 Step -1
'    If Cells(i, CompanyColumn).Value = currentCompany Then
'        endRow = i
'    Else
'        For j = startRow To endRow Step -1
'            countIfDup = Application.WorksheetFunction.CountIf(Range(Cells(endRow, IDColumn), Cells(startRow, IDColumn)), Cells(j, IDColumn).Value)
'            If countIfDup > 1 Then
'                Cells(j, 1).EntireRow.Delete
'            End If
'        Next j
'        startRow = i
'        currentCompany = Cells(i, CompanyColumn).Value
'    End If
'Next i
''''''''''''''''''''''''''''''''''''''''''''
PROGRESS_BAR_MODE = 0
Unload Deduplicate
Call ResetButtonPosition
Call OptimizeCode_End

End Sub
'Master & Alased process
Sub IndicateMasterAndAliased(ByVal companyName As Range, ByVal locationName As Range, ByVal LocationLevel As String, _
    ByVal GIDName As Range, ByVal theInformation As Range, ByVal OIDName As Range, ByVal LevelHeader As Range, ByVal ParentHeader As _
    Range, ByVal StatusHeader As Range)

Dim col As Range
Dim companyNameCol, locationNameCol, GIDNameCol, theInformationCol, OIDNameCol, countRowNum, countColNum, i, j, k, addColumn As Long
Dim startRow, endRow, minOID As Long
Dim currentCompanyName, nextCompanyName As String
Dim informationCheck, countryCheck As Boolean
Dim pctDone As Long
pctDone = 0

Call OptimizeCode_Begin
ResultProcess.FrameProgressResults.Caption = "Complete: " & pctDone & "%"
Worksheets("Master & Aliased").Activate

companyNameCol = companyName.Column
locationNameCol = locationName.Column
GIDNameCol = GIDName.Column
theInformationCol = theInformation.Column
OIDNameCol = OIDName.Column

countColNum = Worksheets("Master & Aliased").Cells(1, 2).CurrentRegion.Columns.Count
countRowNum = Worksheets("Master & Aliased").Cells(1, 2).CurrentRegion.Rows.Count

If Cells(1, 2).Value = "Comment" Then
    Unload ResultProcess
    Unload ConfigSetting
    MsgBox "Results already shown"
    Exit Sub
End If

ActiveWorkbook.Worksheets("Master & Aliased").Sort.SortFields.Clear
ActiveWorkbook.Worksheets("Master & Aliased").Sort.SortFields.Add Key:=Range(Cells(1, LevelHeader.Column), Cells(countRowNum, LevelHeader.Column)) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
ActiveWorkbook.Worksheets("Master & Aliased").Sort.SortFields.Add Key:=Range(Cells(1, companyNameCol), Cells(countRowNum, companyNameCol)) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Master & Aliased").Sort
        .SetRange Cells(1, 2).CurrentRegion
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
'
'columns to insert
addColumn = 8
Columns("B:I").Insert Shift:=xlToRight
'
Cells(1, 2).Value = "Comment"
Cells(1, 3).Value = "OID to be changed to"
Cells(1, 4).Value = "Master or Aliased"
Cells(1, 5).Value = "No location level results"
Cells(1, 6).Value = "Location level matching"
Cells(1, 7).Value = "GID"
Cells(1, 8).Value = "Customer Information Difference"
Cells(1, 9).Value = "Real Display Name"
companyNameCol = companyNameCol + addColumn
locationNameCol = locationNameCol + addColumn
GIDNameCol = GIDNameCol + addColumn
theInformationCol = theInformationCol + addColumn
OIDNameCol = OIDNameCol + addColumn
countColNum = countColNum + addColumn

minOID = 0
informationCheck = True
countryCheck = False
startRow = 2
currentCompanyName = Cells(2, companyNameCol).Value


'
'get real company name from source
For i = 2 To countRowNum
Dim realComName As Range
Dim OIDColInSource As Long
OIDColInSource = OIDNameCol - addColumn
Set realComName = Worksheets("Source").Columns(OIDColInSource).Find(Worksheets("Master & Aliased").Cells(i, OIDNameCol).Value, _
LookIn:=xlValues, lookat:=xlWhole)
Worksheets("Master & Aliased").Cells(i, 9).Value = Worksheets("Source").Cells(realComName.Row, 2).Value
'Progress bar
    pctDone = i / countRowNum * 100
    With ResultProcess
        .FrameProgressResults.Caption = "Complete(1/2): " & pctDone & "%"
        .LabelProgressResults.Width = pctDone * 2.4
        DoEvents
    End With
Next i
'
'

For i = 3 To countRowNum + 1

If currentCompanyName = Cells(i, companyNameCol).Value Then
    endRow = i
Else

' Check if the information is all the same
For j = startRow To endRow
    For k = startRow + 1 To endRow
        For Each col In theInformation.Columns
            If Cells(j, col.Column).Value <> Cells(k, col.Column).Value Then
                informationCheck = False
            End If
        Next col
    Next k
Next j

'Check the information, GID and country
For j = startRow To endRow
    If informationCheck = False Then
        Cells(j, 8).Value = "At Least One Difference"
    Else
        Cells(j, 8).Value = "All The Same"
    End If
    
    If Cells(j, GIDNameCol) <> "" Then
        Cells(j, 7).Value = "Has GID"
    Else
        Cells(j, 7).Value = "No GID"
    End If
    
''''''''''''''''''''''''''''''''''''
'Older Version
'    If StrComp(Cells(j, locationNameCol), CountryName, 1) = 0 Then
'        Cells(j, 6).Value = "Country"
'        countryCheck = True
'    Else
'        Cells(j, 6).Value = "City"
'    End If
'''''''''''''''''''''''''''''''''''''
    If StrComp(Cells(j, locationNameCol), Cells(j, LevelHeader.Column)) = 0 Then
        Cells(j, 6).Value = "Is " & LocationLevel
        countryCheck = True
    Else
        Cells(j, 6).Value = "Not " & LocationLevel
    End If
Next j

'Check the situation of no country
For j = startRow To endRow
    If countryCheck = True Then
        Cells(j, 5).Value = "Have " & LocationLevel
    Else
        Cells(j, 5).Value = "No " & LocationLevel
    End If
Next j

'Get minimum OID
For j = startRow To endRow
'    If Cells(j, 6).Value = "Is " & LocationLevel And Cells(j, 7).Value = "Has GID" Then
        If minOID = 0 Then
            minOID = Cells(j, OIDNameCol).Value
        ElseIf Cells(j, OIDNameCol).Value < minOID Then
            minOID = Cells(j, OIDNameCol).Value
        End If
'    End If
Next j
'For j = startRow To endRow
'    If Cells(j, 6).Value = "Is " & LocationLevel And minOID = 0 Then
'        minOID = Cells(j, OIDNameCol).Value
'        For k = startRow To endRow
'            If Cells(k, 6).Value = "Is " & LocationLevel And Cells(k, OIDNameCol).Value < minOID Then
'                minOID = Cells(k, OIDNameCol).Value
'            End If
'        Next k
'    End If
'Next j
'For j = startRow To endRow
'    If Cells(j, 5).Value = "No " & LocationLevel Then
'        If minOID = 0 Then
'            minOID = Cells(startRow, OIDNameCol).Value
'        Else
'            If Cells(j, OIDNameCol).Value < minOID Then
'                minOID = Cells(j, OIDNameCol).Value
'            End If
'        End If
'    End If
'Next j

'Master & Aliased
For j = startRow To endRow
    If Cells(j, OIDNameCol).Value = minOID Then
        Cells(j, 4).Value = "Master"
        Cells(j, 3).Value = ""
    Else
        Cells(j, 4).Value = "Aliased"
        Cells(j, 3).Value = minOID
    End If
Next j

''''''''''''''''''''''''''''''''''''''''
'New way to define master and aliased
'
Dim findNewMaster, selectionRange As Range
For j = startRow To endRow
    Set selectionRange = Range(Cells(startRow, 9), Cells(endRow, 9))
    If Cells(j, 4).Value = "Master" Then
        If Application.WorksheetFunction.CountIf(selectionRange, Cells(j, 9)) > 1 And _
        Cells(j, 6).Value = "Not " & LocationLevel Then
            Set findNewMaster = Worksheets("Master & Aliased").Range(Cells(startRow, 9), Cells(endRow, 9)).Find(Cells(j, 9).Value, _
            LookIn:=xlValues, lookat:=xlWhole)
            ', after:=Cells(startRow, 9)
            If Not findNewMaster Is Nothing Then
                Do
                If Cells(findNewMaster.Row, 6).Value = "Is " & LocationLevel Then
                    Cells(findNewMaster.Row, 4).Value = "Master"
                    Cells(findNewMaster.Row, 4).Interior.ColorIndex = 22
                    Cells(j, 4).Value = "Aliased"
                    For k = startRow To endRow
                        Cells(k, 3).Value = Cells(findNewMaster.Row, OIDNameCol).Value
                    Next k
                    Cells(findNewMaster.Row, 3).Value = ""
                    Exit Do
                End If
                Set findNewMaster = Worksheets("Master & Aliased").Range(Cells(startRow, 9), Cells(endRow, 9)).FindNext(findNewMaster)
                Loop While findNewMaster.Row <> j
                Set findNewMaster = Nothing
            End If
        End If
    End If
Next j
'''''''''''''''''''''''''''''''''''''''
'
'
'Master no GID, Aliased has GID
'Master GID and Aliased GID are different
'Master no Parent, Aliased has Parent
'Master Parent and Aliased Parent are different
'Master Status and Aliased Status are different
For j = startRow To endRow
    If Cells(j, 4).Value = "Master" And Cells(j, 7).Value = "No GID" Then
    For k = startRow To endRow
        If Cells(k, 4).Value = "Aliased" And Cells(k, 7).Value = "Has GID" Then
            If Not InStr(1, Cells(k, 2).Value, "GID", vbTextCompare) > 0 Then
                Cells(k, 2).Value = Cells(k, 2).Value & "Check GID. "
            End If
        End If
    Next k
    End If
    
    If Cells(j, 4).Value = "Master" And Cells(j, 7).Value = "Has GID" Then
    For k = startRow To endRow
        If Cells(k, 4).Value = "Aliased" And Cells(k, 7).Value = "Has GID" Then
            If Cells(j, GIDName.Column).Value <> Cells(k, GIDName.Column).Value Then
                If Not InStr(1, Cells(k, 2).Value, "GID", vbTextCompare) > 0 Then
                    Cells(k, 2).Value = Cells(k, 2).Value & "Check GID. "
                End If
            End If
        End If
    Next k
    End If
    
    If Cells(j, 4).Value = "Master" And Cells(j, ParentHeader.Column).Value = "" Then
    For k = startRow To endRow
        If Cells(k, 4).Value = "Aliased" And Cells(k, ParentHeader.Column).Value <> "" Then
            If Not InStr(1, Cells(k, 2).Value, "Parent", vbTextCompare) > 0 Then
                Cells(k, 2).Value = Cells(k, 2).Value & "Check Parent. "
            End If
        End If
    Next k
    End If
    
    If Cells(j, 4).Value = "Master" And Cells(j, ParentHeader.Column).Value <> "" Then
    For k = startRow To endRow
        If Cells(k, 4).Value = "Aliased" And Cells(k, ParentHeader.Column).Value <> "" Then
            If Cells(j, ParentHeader.Column).Value <> Cells(k, ParentHeader.Column).Value Then
                If Not InStr(1, Cells(k, 2).Value, "Parent", vbTextCompare) > 0 Then
                    Cells(k, 2).Value = Cells(k, 2).Value & "Check Parent. "
                End If
            End If
        End If
    Next k
    End If
    
    If Cells(j, 4).Value = "Master" Then
    For k = startRow To endRow
        If Cells(k, 4).Value = "Aliased" Then
            If Cells(j, StatusHeader.Column).Value <> Cells(k, StatusHeader.Column).Value And Cells(j, StatusHeader.Column).Value <> "Active" Then
                If Not InStr(1, Cells(k, 2).Value, "Status", vbTextCompare) > 0 Then
                    Cells(k, 2).Value = Cells(k, 2).Value & "Check Status. "
                End If
            End If
        End If
    Next k
    End If
    
    If Cells(j, 8).Value = "At Least One Difference" And Cells(j, 4).Value = "Aliased" Then
        If Cells(j, 2).Value <> "" Then
            Cells(j, 2).Value = Cells(j, 2).Value & "Check Information. "
        Else
            Cells(j, 2).Value = "Check Information. "
        End If
    End If
Next j

minOID = 0
countryCheck = False
informationCheck = True
currentCompanyName = Cells(i, companyNameCol).Value
startRow = i
End If
Next i

''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Export Results
''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim reComment, reMOID, reMName, reMLoc, reMaOrAl, reAOID, reAName, reALoc, reACoun, rePar, reGID, reType, reCate, _
reCACC, reStat, reIfDup, reSim, reMCoun, reMPar, reMGID, reMtype, reMStatus, reAStatus As Long
reComment = 2
reMOID = 3
reMName = 4
reMLoc = 5
reMCoun = 6
reMPar = 7
reMGID = 8
reMStatus = 9
reMtype = 10
'reMaOrAl = 6
reAOID = 10 + theInformation.Columns.Count
reAName = 11 + theInformation.Columns.Count
reALoc = 12 + theInformation.Columns.Count
reACoun = 13 + theInformation.Columns.Count
rePar = 14 + theInformation.Columns.Count
reGID = 15 + theInformation.Columns.Count
reAStatus = 16 + theInformation.Columns.Count
reType = 17 + theInformation.Columns.Count
'reCate = 14
'reCACC = 15
'reStat = 16
Worksheets("Results").Cells.Clear
Worksheets("Results").Cells.ColumnWidth = 25
Worksheets("Results").Cells(1, reComment).Value = "Comment"
Worksheets("Results").Cells(1, reMOID).Value = "Master OID"
Worksheets("Results").Cells(1, reMName).Value = "Master Name"
Worksheets("Results").Cells(1, reMLoc).Value = "Master Location"
Worksheets("Results").Cells(1, reMCoun).Value = "Master Country"
Worksheets("Results").Cells(1, reMPar).Value = "Master Parent"
Worksheets("Results").Cells(1, reMGID).Value = "Master GID"
Worksheets("Results").Cells(1, reMStatus).Value = "Master Status"
'Worksheets("Results").Cells(1, reMaOrAl).Value = "Master or Aliased"
Worksheets("Results").Cells(1, reAOID).Value = "Aliased OID"
Worksheets("Results").Cells(1, reAName).Value = "Aliased Name"
Worksheets("Results").Cells(1, reALoc).Value = "Aliased Location"
Worksheets("Results").Cells(1, reACoun).Value = "Aliased Country"
Worksheets("Results").Cells(1, rePar).Value = "Aliased Parent"
Worksheets("Results").Cells(1, reGID).Value = "Aliased GID"
Worksheets("Results").Cells(1, reAStatus).Value = "Aliased Status"
'Worksheets("Results").Cells(1, reType).Value = "Type"
'Worksheets("Results").Cells(1, reCate).Value = "Category"
'Worksheets("Results").Cells(1, reCACC).Value = "CACC"
'Worksheets("Results").Cells(1, reStat).Value = "Status"
k = reMtype
For Each col In theInformation.Columns
    Worksheets("Results").Cells(1, k).Value = "Master " & Worksheets("Master & Aliased").Cells(1, col.Column).Value
    k = k + 1
Next col
k = reType
For Each col In theInformation.Columns
    Worksheets("Results").Cells(1, k).Value = "Aliased " & Worksheets("Master & Aliased").Cells(1, col.Column).Value
    k = k + 1
Next col
reSim = k
Worksheets("Results").Cells(1, reSim).Value = "Similarity"
reIfDup = k + 1
Worksheets("Results").Cells(1, reIfDup).Value = "Duplicated account name and location check"

j = 2

Dim MasterRow As Long
For i = 2 To countRowNum
If Worksheets("Master & Aliased").Cells(i, 4) = "Aliased" Then
    Worksheets("Results").Cells(j, reComment).Value = Worksheets("Master & Aliased").Cells(i, 2).Value
    'If Worksheets("Master & Aliased").Cells(i, 3).Value <> "" Then
        Worksheets("Results").Cells(j, reMOID).Value = Worksheets("Master & Aliased").Cells(i, 3).Value
    'Else
     '   Worksheets("Results").Cells(j, reMOID).Value = Worksheets("Master & Aliased").Cells(i, OIDNameCol).Value
    'End If
    'Master
    MasterRow = Range(Worksheets("Master & Aliased").Cells(1, OIDNameCol), Worksheets("Master & Aliased").Cells(countRowNum, _
    OIDNameCol)).Find(Worksheets("Results").Cells(j, reMOID).Value, LookIn:=xlValues, lookat:=xlWhole).Row
    Worksheets("Results").Cells(j, reMLoc).Value = Worksheets("Master & Aliased").Cells(MasterRow, locationNameCol).Value
    Worksheets("Results").Cells(j, reMCoun).Value = Worksheets("Master & Aliased").Cells(MasterRow, LevelHeader.Column).Value
    Worksheets("Results").Cells(j, reMPar).Value = Worksheets("Master & Aliased").Cells(MasterRow, ParentHeader.Column).Value
    Worksheets("Results").Cells(j, reMGID).Value = Worksheets("Master & Aliased").Cells(MasterRow, GIDName.Column).Value
    Worksheets("Results").Cells(j, reMStatus).Value = Worksheets("Master & Aliased").Cells(MasterRow, StatusHeader.Column).Value
    k = reMtype
    For Each col In theInformation.Columns
        Worksheets("Results").Cells(j, k).Value = Worksheets("Master & Aliased").Cells(MasterRow, col.Column).Value
        k = k + 1
    Next col
    'Aliased
    'Worksheets("Results").Cells(j, reMaOrAl).Value = Worksheets("Master & Aliased").Cells(i, 4).Value
    Worksheets("Results").Cells(j, reAOID).Value = Worksheets("Master & Aliased").Cells(i, OIDNameCol).Value
    Worksheets("Results").Cells(j, reALoc).Value = Worksheets("Master & Aliased").Cells(i, locationNameCol).Value
    Worksheets("Results").Cells(j, reACoun).Value = Worksheets("Master & Aliased").Cells(i, LevelHeader.Column).Value
    Worksheets("Results").Cells(j, rePar).Value = Worksheets("Master & Aliased").Cells(i, ParentHeader.Column).Value
    Worksheets("Results").Cells(j, reGID).Value = Worksheets("Master & Aliased").Cells(i, GIDName.Column).Value
    Worksheets("Results").Cells(j, reAStatus).Value = Worksheets("Master & Aliased").Cells(i, StatusHeader.Column).Value
    Worksheets("Results").Cells(j, reSim).Value = Worksheets("Master & Aliased").Cells(i, countColNum + 1).Value
    k = reType
    For Each col In theInformation.Columns
        Worksheets("Results").Cells(j, k).Value = Worksheets("Master & Aliased").Cells(i, col.Column).Value
        k = k + 1
    Next col
j = j + 1
End If
Next i

Dim sourceOIDRow As Long
Dim sourceCompanyName, sourceCompanyLocation As String
Dim sourceCompanyNameRange As Range
Dim countRowNumTwo, OIDNameColTwo, locationNameColTwo As Long
Dim findMOID As Range
OIDNameColTwo = OIDNameCol - addColumn
locationNameColTwo = locationNameCol - addColumn
countRowNum = Worksheets("Results").Cells(1, 3).End(xlDown).Row
countRowNumTwo = Worksheets("Source").Cells(1, 2).CurrentRegion.Rows.Count
For i = 2 To countRowNum
    'sourceOIDRow = Worksheets("Source").Cells(1, 2).CurrentRegion.Find(Worksheets("Results").Cells(i, 3).Value, LookIn:=xlValues).Row
    sourceOIDRow = Worksheets("Source").Columns(OIDNameColTwo).Find(Worksheets("Results").Cells(i, reMOID).Value, _
    LookIn:=xlValues, lookat:=xlWhole).Row
    sourceCompanyName = Worksheets("Source").Cells(sourceOIDRow, 2).Value
    'sourceCompanyLocation = Worksheets("Source").Cells(sourceOIDRow, 3).Value
    Worksheets("Results").Cells(i, reMName).Value = sourceCompanyName
    'Worksheets("Results").Cells(i, reMLoc).Value = sourceCompanyLocation
    
    'sourceOIDRow = Worksheets("Source").Cells(1, 2).CurrentRegion.Find(Worksheets("Results").Cells(i, 6).Value, LookIn:=xlValues).Row
    sourceOIDRow = Worksheets("Source").Columns(OIDNameColTwo).Find(Worksheets("Results").Cells(i, reAOID).Value, _
    LookIn:=xlValues, lookat:=xlWhole).Row
    sourceCompanyName = Worksheets("Source").Cells(sourceOIDRow, 2).Value
    Worksheets("Results").Cells(i, reAName).Value = sourceCompanyName
    
    'Find matster oid in aliased oid
    Set findMOID = Worksheets("Results").Columns(reAOID).Find(Worksheets("Results").Cells(i, reMOID).Value, _
    LookIn:=xlValues, lookat:=xlWhole)
    If Not findMOID Is Nothing Then
        Worksheets("Results").Cells(i, reComment).Value = "Invalid Aliased Match. " & Worksheets("Results").Cells(i, reComment).Value
    End If
    Set findMOID = Nothing
        
    'check dup
    If Worksheets("Results").Cells(i, reMName).Value = Worksheets("Results").Cells(i, reAName).Value And _
    Worksheets("Results").Cells(i, reALoc).Value = Worksheets("Results").Cells(i, reACoun).Value And _
    Worksheets("Results").Cells(i, reMLoc).Value <> Worksheets("Results").Cells(i, reACoun).Value Then
    Worksheets("Results").Cells(i, reIfDup).Value = "Failed to aliased due to duplicated accout name"
    End If
    
    'Progress bar
    pctDone = i / countRowNum * 100
    With ResultProcess
        .FrameProgressResults.Caption = "Complete(2/2): " & pctDone & "%"
        .LabelProgressResults.Width = pctDone * 2.4
        DoEvents
    End With
Next i

''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Dim countDup, countDupTwo As Long
'Dim findObject, findObjectTwo As Range
'Dim findValue As String
'pctDone = 0
'With ResultProcess
'    .FrameProgressResults.Caption = "Optimizing results: " & pctDone & "%"
'    .LabelProgressResults.Width = pctDone * 2.4
'End With
'countRowNum = Worksheets("Results").Cells(1, 2).CurrentRegion.Rows.Count
'countDup = 0
'countDupTwo = 0
'For i = countRowNum To 2 Step -1
'    Set findObject = Nothing
'    Set findObjectTwo = Nothing
'    'col 3: master OID, col 6: aliased OID
'    Set findObject = Worksheets("Results").Columns(reAOID).Find(Worksheets("Results").Cells(i, reMOID).Value, _
'    LookIn:=xlValues, lookat:=xlWhole)
'    Set findObjectTwo = Worksheets("Results").Columns(reAOID).Find(Worksheets("Results").Cells(i, reAOID).Value, _
'    LookIn:=xlValues, lookat:=xlWhole)
'    If Not findObjectTwo Is Nothing Then
'        findValue = findObjectTwo.Address
'    End If
'    If findObjectTwo.Row = i Then
'        Set findObjectTwo = Worksheets("Results").Columns(reAOID).FindNext(findObjectTwo)
'        If findObjectTwo.Row = i Then
'            Set findObjectTwo = Nothing
'        End If
'    End If
'
'    If Not findObjectTwo Is Nothing Then
'        If Worksheets("Results").Cells(i, reMOID).Value <> Worksheets("Results").Cells(findObjectTwo.Row, reMOID).Value Then
'            With Worksheets("Results")
'                .Cells(i, reMOID).Interior.ColorIndex = 42
'                .Cells(findObjectTwo.Row, reMOID).Interior.ColorIndex = 42
'            End With
'        Else
'            Worksheets("Results").Cells(i, 1).EntireRow.Delete
'        End If
'    End If
'
'    If Not findObject Is Nothing And Not findObjectTwo Is Nothing Then
'    Do
'        If Worksheets("Results").Cells(findObject.Row, reMOID).Value = _
'        Worksheets("Results").Cells(findObjectTwo.Row, reMOID).Value Then
'            If Worksheets("Results").Cells(i, 1).Interior.ColorIndex = xlNone Then
'                Worksheets("Results").Cells(i, 1).EntireRow.Interior.ColorIndex = 22
'            End If
'            'If Worksheets("Results").Cells(findObject.Row, 1).Interior.ColorIndex = xlNone Then
'                Worksheets("Results").Cells(findObject.Row, 1).EntireRow.Interior.ColorIndex = xlNone
'            'End If
'            'If Worksheets("Results").Cells(findObjectTwo.Row, 1).Interior.ColorIndex = xlNone Then
'                Worksheets("Results").Cells(findObjectTwo.Row, 1).EntireRow.Interior.ColorIndex = xlNone
'            'End If
'        Else
'        '    If Worksheets("Results").Cells(i, 1).Interior.ColorIndex = xlNone Then
'        '        Worksheets("Results").Cells(i, 1).EntireRow.Interior.ColorIndex = 36
'        '    End If
'        '    If Worksheets("Results").Cells(findObject.Row, 1).Interior.ColorIndex = xlNone Then
'        '        Worksheets("Results").Cells(findObject.Row, 1).EntireRow.Interior.ColorIndex = 44
'        '    End If
'        '    If Worksheets("Results").Cells(findObjectTwo.Row, 1).Interior.ColorIndex = xlNone Then
'        '    Worksheets("Results").Cells(findObjectTwo.Row, 1).EntireRow.Interior.ColorIndex = 44
'        '    End If
'        End If
'        'If Worksheets("Results").Cells(findObject.Row, 3).Value <> Worksheets("Results").Cells(findObjectTwo.Row, 3).Value Then
'            'Worksheets("Results").Cells(findObject.Row, 1).EntireRow.Interior.ColorIndex = 44
'            'Worksheets("Results").Cells(findObjectTwo.Row, 1).EntireRow.Interior.ColorIndex = 44
'        'End If
'        'countDupTwo = findObjectTwo.Row
'    Set findObjectTwo = Worksheets("Results").Columns(reAOID).FindNext(findObjectTwo)
'    Loop While findValue <> findObjectTwo.Address
'    End If
'
'
'    'If countDup <> 0 Then
'    'ElseIf countDupTwo <> 0 And Worksheets("Results").Cells(i, 3).Value = Worksheets("Results").Cells(countDupTwo, 3).Value Then
'    'End If
'    'countDup = 0
'    'countDupTwo = 0
'
'    'Progress bar
'    pctDone = (countRowNum + 2 - i) / countRowNum * 100
'    With ResultProcess
'        .FrameProgressResults.Caption = "Optimizing results: " & pctDone & "%"
'        .LabelProgressResults.Width = pctDone * 2.4
'        DoEvents
'    End With
'Next i
''''''''''''''''''''''''''''''''''''''''''''''''''''''
Call OptimizeCode_End
Unload ResultProcess
Unload ConfigSetting
End Sub

'
'Still Testing
'
Sub PassValueBetweenForms(ByVal companyName As Range, ByVal locationName As Range, ByVal CountryName As String, _
    ByVal GIDName As Range, ByVal theInformation As Range, ByVal OIDName As Range, ByVal CountryHeader As Range, ByVal ParentHeader As _
    Range, ByVal theOption As Long)
    Static companyNameVar, locationNameVar, GIDNameVar, theInformationVar, OIDNameVar, CountryHeaderVar, ParentHeaderVar As Range
    Static CountryNameVar As String
    If theOption = 0 Then
        companyNameVar = companyName
        locationNameVar = locationName
        GIDNameVar = GIDName
        theInformationVar = theInformation
        OIDNameVar = OIDName
        CountryHeaderVar = CountryHeader
        ParentHeaderVar = ParentHeader
        CountryNameVar = CountryName
    ElseIf theOption = 1 Then
        'Call IndicateMasterAndAliased(companyNameVar, locationNameVar, CountryNameVar, GIDNameVar, theInformationVar, OIDNameVar, _
        CountryHeaderVar, ParentHeaderVar)
    End If
End Sub
