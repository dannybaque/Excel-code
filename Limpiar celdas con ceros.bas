Attribute VB_Name = "M�dulo1"
Sub RemoveZero()

    Dim rango As Range
    Set rango = Selection
    
    For Each celda In rango
        If celda.Value = 0 Then
            celda.ClearContents
        End If
    Next
    
End Sub
