Attribute VB_Name = "AllFunctions"
'Function AlphaNumericOnly(strSource As String) As String

Public CalcState As Long
Public EventState As Boolean
Public PageBreakState As Boolean

Sub OptimizeCode_Begin()

Application.ScreenUpdating = False

EventState = Application.EnableEvents
'Application.EnableEvents = False

CalcState = Application.Calculation
Application.Calculation = xlCalculationManual

PageBreakState = ActiveSheet.DisplayPageBreaks
ActiveSheet.DisplayPageBreaks = False

End Sub

Sub OptimizeCode_End()

ActiveSheet.DisplayPageBreaks = PageBreakState
Application.Calculation = CalcState
'Application.EnableEvents = EventState
Application.ScreenUpdating = True

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
    strResult = Replace(strResult, "VIET NAM", "VIETNAM", , , 1)
    If InStr(1, strResult, "Korea", 1) <> 0 Then
        strResult = "SOUTH KOREA"
    End If
    
    CleanCountryName = strResult
End Function
'Export datas into Table
Sub ExportCleanedData(ByVal exportLocation As String)
Dim targetSheet As String
Dim tableNumber As String

If exportLocation = "One" Then
    targetSheet = "(1) Model N"
    tableNumber = "1"
ElseIf exportLocation = "Two" Then
    targetSheet = "(2) SFDC"
    tableNumber = "2"
Else
    MsgBox "Error in function ExportCleanedData()"
End If

Worksheets(targetSheet).Activate
On Error Resume Next
'ActiveSheet.ListObjects("Table Of Cleaned Data Sheet " & tableNumber).Range.Select
ActiveSheet.ListObjects(targetSheet).Range.Select
Selection.Delete

Worksheets("Data Cleaner").Activate
Worksheets("Data Cleaner").Range("B1").CurrentRegion.Select
Selection.Copy

Worksheets(targetSheet).Activate
ActiveSheet.Paste Destination:=Worksheets(targetSheet).Cells(1, 2)
'ActiveSheet.ListObjects.Add(xlSrcRange, Range("B1").CurrentRegion, , xlYes).Name = "Table Of Cleaned Data Sheet " & tableNumber
ActiveSheet.ListObjects.Add(xlSrcRange, Range("B1").CurrentRegion, , xlYes).Name = targetSheet


Cells(1, 1).Select
Application.CutCopyMode = False
Worksheets("Data Cleaner").Activate
Cells(1, 1).Select
End Sub

'
'OID GID Matching
'
Sub OIDandGIDMatching()
    Dim countRowNum, countColNum, i, countryNameMatch, j, outputCurRow, addColumn, startRow, endRow, k As Long
    Dim currentRange As Range
    Dim currentCompany As Long
    Dim theFlag As Boolean
    Dim pctDone As Long
    Dim pct As Long
    Call OptimizeCode_Begin
    
    Worksheets("Fuzzy Lookup").Activate
    Set currentRange = Cells(1, 2).CurrentRegion
    countRowNum = currentRange.Rows.Count
    countColNum = currentRange.Columns.Count
    
    currentRange.Sort Key1:=Columns(gModelNCountry.Column), Order1:=xlAscending, key2:=Columns(gModelNCompany.Column), _
    order2:=xlAscending, Header:=xlYes
    
    If Cells(1, countColNum).Value <> "Country Match" Then
    countryNameMatch = countColNum + 1
    Cells(1, countryNameMatch).Value = "Country Match"
    For i = 2 To countRowNum
        If UCase(Cells(i, gModelNCountry.Column).Value) = UCase(Cells(i, gSFDCCountry.Column).Value) Then
            Cells(i, countryNameMatch).Value = "TRUE"
            If Not gModelNState Is Nothing Then
                If UCase(Cells(i, gModelNState.Column).Value) <> UCase(Cells(i, gSFDCState.Column).Value) Then
                    Cells(i, countryNameMatch).Value = "FALSE"
                End If
            End If
        Else
            Cells(i, countryNameMatch).Value = "FALSE"
        End If
    '
    pctDone = i / countRowNum * 100 * 0.6
    With GIDMatchProgress
        .FrameProgressGID.Caption = "Complete: " & pctDone & "%"
        .LabelProgressGID.Width = pctDone * 2.4
        DoEvents
    End With
    '
    Next i
    End If
    
    Set currentRange = Cells(1, 2).CurrentRegion
    countRowNum = currentRange.Rows.Count
    countColNum = currentRange.Columns.Count
    
    currentRange.Select
    Selection.AutoFilter
    With currentRange
        .AutoFilter Field:=countColNum, Criteria1:="TRUE"
        .AutoFilter Field:=countColNum - 1, Criteria1:=">0"
    End With
    Cells(1, 1).Select
    
    Cells(1, 1).CurrentRegion.Copy Worksheets("Results").Cells(1, 1)
    Worksheets("Results").Cells(1, 1).Value = ""
   
    Worksheets("Results").Activate
    
    Set currentRange = Cells(1, 2).CurrentRegion
    
    currentRange.Sort Key1:=Columns(gModelNOID.Column), Order1:=xlAscending, DataOption1:=xlSortNormal, _
    Header:=xlYes
        
    If Not gModelNCity Is Nothing And Not gSFDCCity Is Nothing Then
    currentCompany = 1
    startRow = 1
    endRow = 1
    theFlag = False
    For i = 2 To countRowNum
        If currentCompany = 1 Then
            currentCompany = Cells(i, gModelNOID.Column).Value
            startRow = i
        ElseIf Cells(i, gModelNOID.Column).Value = currentCompany Then
            endRow = i
        Else
            For j = startRow To endRow
                If UCase(Cells(j, gModelNCity.Column).Value) = UCase(Cells(j, gSFDCCity.Column).Value) And _
                (UCase(Cells(j, gSFDCStatus.Column).Value) = "ACTIVE" Or Cells(j, gSFDCStatus.Column).Value = "0") And theFlag = False Then
                    Cells(j, 2).Interior.ColorIndex = 32
                    theFlag = True
                Else
                    Cells(j, 2).Interior.ColorIndex = 22
                End If
            Next j
            
            If theFlag = False Then
                For k = startRow To endRow
                If UCase(Cells(j, gModelNCity.Column).Value) = UCase(Cells(j, gSFDCCity.Column).Value) Then
                    Cells(k, 2).Interior.ColorIndex = 32
                    theFlag = True
                    Exit For
                End If
                Next k
            End If
            
            If theFlag = False Then
                For k = startRow To endRow
                If UCase(Cells(k, gSFDCStatus.Column).Value) = "ACTIVE" Or Cells(k, gSFDCStatus.Column).Value = "0" Then
                    Cells(k, 2).Interior.ColorIndex = 32
                    theFlag = True
                    Exit For
                End If
                Next k
            End If
            
            If theFlag = False Then
                For k = startRow To endRow
                    Cells(k, countColNum + 1).Value = "Mutiple Results. "
                Next k
            End If
            
            theFlag = False
            currentCompany = Cells(i, gModelNOID.Column).Value
            startRow = i
        End If
    '
    pct = pctDone
    pctDone = 60 + i / countRowNum * 100 * 0.05
    With GIDMatchProgress
        .FrameProgressGID.Caption = "Complete: " & pctDone & "%"
        .LabelProgressGID.Width = pctDone * 2.4
        If pctDone >= pct + 1 Then
            DoEvents
            pct = pctDone
        End If
    End With
    '
    Next i
    End If
    For i = 2 To countRowNum
        If Cells(i, 2).Interior.ColorIndex = 32 Then
            Cells(i, 2).Interior.ColorIndex = xlNone
        End If
    '
    pct = pctDone
    pctDone = 65 + i / countRowNum * 100 * 0.05
    With GIDMatchProgress
        .FrameProgressGID.Caption = "Complete: " & pctDone & "%"
        .LabelProgressGID.Width = pctDone * 2.4
        If pctDone >= pct + 1 Then
            DoEvents
            pct = pctDone
        End If
    End With
    '
    Next i
       
    Set currentRange = Cells(1, 2).CurrentRegion
    
    Cells(1, countColNum + 1).Value = "Comment"
   
    If Not gSFDCStatus Is Nothing Then
        For i = 2 To countRowNum
            If UCase(Cells(i, gSFDCStatus.Column).Value) <> "ACTIVE" And UCase(Cells(i, gSFDCStatus.Column).Value) <> "0" Then
                Cells(i, countColNum + 1).Value = Cells(i, countColNum + 1).Value & "SFDC is Inactive. "
            End If
    '
    pctDone = 70 + i / countRowNum * 100 * 0.2
    With GIDMatchProgress
        .FrameProgressGID.Caption = "Complete: " & pctDone & "%"
        .LabelProgressGID.Width = pctDone * 2.4
        DoEvents
    End With
    '
        Next i
    End If
    
    If Not gModelNGID Is Nothing And Not gSFDCGID Is Nothing Then
        For i = 2 To countRowNum
            If Cells(i, gModelNGID.Column).Value <> Cells(i, gSFDCGID.Column).Value And Cells(i, gModelNGID.Column).Value <> "" Then
                Cells(i, countColNum + 1).Value = Cells(i, countColNum + 1).Value & "Compare GID. "
            ElseIf Cells(i, gModelNGID.Column).Value = Cells(i, gSFDCGID.Column).Value Then
                Cells(i, gModelNCompany.Column).Interior.ColorIndex = 45
            End If
    '
    pct = pctDone
    pctDone = 90 + i / countRowNum * 100 * 0.1
    With GIDMatchProgress
        .FrameProgressGID.Caption = "Complete: " & pctDone & "%"
        .LabelProgressGID.Width = pctDone * 2.4
        If pctDone >= pct + 1 Then
            DoEvents
            pct = pctDone
        End If
    End With
    '
        Next i
    End If
    
    
    Set currentRange = Cells(1, 2).CurrentRegion
    
    currentRange.Sort Key1:=Columns(gModelNCountry.Column), Order1:=xlAscending, key2:=Columns(gModelNCompany.Column), _
    order2:=xlAscending, Header:=xlYes
    
    currentRange.Select
    Selection.AutoFilter
    With currentRange
        .AutoFilter Field:=1, Operator:=xlFilterNoFill
    End With
    Cells(1, 1).Select
    
    
    Call ResetButtonPosition
    Call OptimizeCode_End
    Unload GIDMatchProgress
    Unload GIDMatch
End Sub

