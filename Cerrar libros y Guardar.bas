Attribute VB_Name = "Módulo3"
Sub CloseAndSaveAllChanges()
    Dim libriActual As Workbook
        For Each libroActual In Workbooks
        If libroActual.Name <> "Cerrar Libros y Salvar.xlsm" Then
            libroActual.clos savechanges:=True
        End If
    Next libroActual
    
End Sub
