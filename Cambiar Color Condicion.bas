Attribute VB_Name = "Módulo4"
Sub ChangeCellColorCondition()
    Dim miRango As Range
    Set miRango = Range("A1:A10")
    For Each celdaActual In miRango
       If celdaActual.Value = "Valor1" Then celdaActual.Interior.Color = RGB(255, 0, 0)
       If celdaActual.Value = "Valor2" Then celdaActual.Interior.Color = RGB(0, 255, 0)
    Next
End Sub
