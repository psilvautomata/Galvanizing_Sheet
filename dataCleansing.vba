Sub dataClean()

Dim ws As Variant

Application.ScreenUpdating = False 'Disable screen updating
Application.EnableEvents = False 'Disable excel events

ws = ActiveSheet.Name 'Worksheet name

If ws = "Soufer" Then 'Clear data to Soufer

    Range("B6:F75").ClearContents
    Range("B6").Select
    
ElseIf ws = "N 5580" Then 'Clear data to 5580

    Range("B6:I75").ClearContents
    Range("B6").Select
    
ElseIf ws = "CV 300-345 STi" Then 'Clear data to STi

    Range("B11:E23").ClearContents
    Range("B33:E45").ClearContents
    Range("B57:E69").ClearContents
    Range("H11:K23").ClearContents
    Range("H33:K45").ClearContents
    Range("H57:K69").ClearContents
    Range("B11").Select
    
End If

Application.ScreenUpdating = True 'Enable screen updating
Application.EnableEvents = True 'Enable excel events

End Sub
