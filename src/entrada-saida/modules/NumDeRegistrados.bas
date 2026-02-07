Attribute VB_Name = "NumDeRegistrados"
Sub Num()
TelaPrincipal.txtNRegistrado.Caption = ThisWorkbook.Worksheets("Config").Range("I3")
TelaPrincipal.txtNSaidas.Caption = ThisWorkbook.Worksheets("Config").Range("I6")
End Sub


