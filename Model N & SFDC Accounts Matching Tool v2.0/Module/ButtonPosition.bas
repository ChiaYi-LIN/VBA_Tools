Attribute VB_Name = "ButtonPosition"
Sub ResetButtonPosition()
    Dim btn As Shape
    Set btn = Worksheets("README First").Shapes("Reset")
    With btn
        .Height = 58.85
        .Width = 65.25
        .Left = 3.75
        .Top = 15.75
    End With
    
    Set btn = Worksheets("Source").Shapes("ClearSource")
    With btn
        .Height = 27.75
        .Width = 80.25
        .Left = 27
        .Top = 33
    End With

    Set btn = Worksheets("(1) Model N").Shapes("ClearTableOne")
    With btn
        .Height = 27.75
        .Width = 80.25
        .Left = 27
        .Top = 33
    End With
    
    Set btn = Worksheets("(2) SFDC").Shapes("ClearTableTwo")
    With btn
        .Height = 27.75
        .Width = 80.25
        .Left = 27
        .Top = 33
    End With
    
    Set btn = Worksheets("Data Cleaner").Shapes("StartButton")
    With btn
        .Height = 27.75
        .Width = 88.5
        .Left = 24.75
        .Top = 33
    End With
    
    Set btn = Worksheets("Data Cleaner").Shapes("ClearData")
    With btn
        .Height = 27.75
        .Width = 88.5
        .Left = 24.75
        .Top = 65
    End With
    
    Set btn = Worksheets("Data Cleaner").Shapes("ExporterOne")
    With btn
        .Height = 57
        .Width = 88.5
        .Left = 24.75
        .Top = 106.5
    End With
    
    Set btn = Worksheets("Data Cleaner").Shapes("ExporterTwo")
    With btn
        .Height = 57
        .Width = 88.5
        .Left = 24.75
        .Top = 181.5
    End With
    
    Set btn = Worksheets("Fuzzy Lookup").Shapes("OIDGIDMatch")
    With btn
        .Height = 45.75
        .Width = 120
        .Left = 8.25
        .Top = 30.75
    End With
    
    Set btn = Worksheets("Fuzzy Lookup").Shapes("ClearMatchingData")
    With btn
        .Height = 45.75
        .Width = 120
        .Left = 8.25
        .Top = 90
    End With
    
    Set btn = Worksheets("Results").Shapes("ClearResults")
    With btn
        .Height = 45.75
        .Width = 120
        .Left = 8.25
        .Top = 30.75
    End With
End Sub
