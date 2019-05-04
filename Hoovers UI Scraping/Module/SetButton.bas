Attribute VB_Name = "SetButton"
Sub ButtonPosition()
    Dim btn As Shape
    
    Set btn = Worksheets("README").Shapes("ResetButton")
    btn.Top = 46
    btn.Left = 19
    btn.Height = 28.5
    btn.Width = 99

    Set btn = Worksheets("DUNS").Shapes("StartButton")
    btn.Top = 46
    btn.Left = 19
    btn.Height = 28.5
    btn.Width = 99
    
    Set btn = Worksheets("DUNS").Shapes("ClearButton")
    btn.Top = 91
    btn.Left = 19
    btn.Height = 28.5
    btn.Width = 99


End Sub
