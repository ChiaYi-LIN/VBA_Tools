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

    Set btn = Worksheets("Model N Data").Shapes("ClearTableOne")
    With btn
        .Height = 27.75
        .Width = 80.25
        .Left = 27
        .Top = 33
    End With
    
    
    Set btn = Worksheets("Data Cleaner").Shapes("StartButton")
    With btn
        .Height = 27.75
        .Width = 80.25
        .Left = 27
        .Top = 33
    End With
    
    Set btn = Worksheets("Data Cleaner").Shapes("ClearData")
    With btn
        .Height = 27.75
        .Width = 80.25
        .Left = 27
        .Top = 65
    End With
    
    Set btn = Worksheets("Data Cleaner").Shapes("ExporterOne")
    With btn
        .Height = 57
        .Width = 80.25
        .Left = 27
        .Top = 106.5
    End With
    
    Set btn = Worksheets("Fuzzy Lookup").Shapes("HighlightSameResults")
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
    
    Set btn = Worksheets("Fuzzy Lookup").Shapes("ExportMatchedData")
    With btn
        .Height = 45.75
        .Width = 120
        .Left = 8.25
        .Top = 150.75
    End With
    
    Set btn = Worksheets("Master & Aliased").Shapes("DeleteDupID")
    With btn
        .Height = 45.75
        .Width = 120
        .Left = 8.25
        .Top = 30.75
    End With
    
    Set btn = Worksheets("Master & Aliased").Shapes("MasterAliasedIndicator")
    With btn
        .Height = 45.75
        .Width = 120
        .Left = 8.25
        .Top = 90
    End With
    
    Set btn = Worksheets("Master & Aliased").Shapes("ClearMasterAliased")
    With btn
        .Height = 45.75
        .Width = 120
        .Left = 8.25
        .Top = 150.75
    End With
    
    Set btn = Worksheets("Results").Shapes("ClearResults")
    With btn
        .Height = 45.75
        .Width = 120
        .Left = 8.25
        .Top = 30.75
    End With
End Sub
