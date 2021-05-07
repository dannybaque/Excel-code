VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
 Dim nombreFichero As String
 'Nombre fichero = 20170827134320.xlsx
 nombreFichero = "d:/temp/excel/" & Year(Date) & Month(Date) & Day(Date) & Hour(Time) & Minute(Time) & Second(Time) & ".xslx"
 ThisWorkbook.SaveAs Filename:=nombreFichero
 
End Sub
