Attribute VB_Name = "M�dulo2"
Sub ShowHidenSheets()
    Dim sheet As Worksheet
    For Each sheet In ActiveWorkbook.Worksheets
        sheet.Visible = xlSheetVisible
    Next
End Sub
