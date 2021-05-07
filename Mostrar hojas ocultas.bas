Attribute VB_Name = "Módulo2"
Sub ShowHidenSheets()
    Dim sheet As Worksheet
    For Each sheet In ActiveWorkbook.Worksheets
        sheet.Visible = xlSheetVisible
    Next
End Sub
