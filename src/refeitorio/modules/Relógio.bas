Attribute VB_Name = "Relógio"
Public RunWhen As Double
Sub StartCiclo()
Verificador.LabelHoraAtual.Caption = Format(Now, "hh:mm:ss")
RunWhen = Now + TimeSerial(0, 0, 1)
Application.OnTime RunWhen, "StartCiclo", , True 'Aqui está o loop'
End Sub
Sub StopCiclo()
    Application.OnTime RunWhen, "StartCiclo", , False
End Sub

Sub StartCiclo2()
Verificador2.LabelHoraAtual.Caption = Format(Now, "hh:mm:ss")
RunWhen = Now + TimeSerial(0, 0, 1)
Application.OnTime RunWhen, "StartCiclo2", , True 'Aqui está o loop'
End Sub
Sub StopCiclo2()
    Application.OnTime RunWhen, "StartCiclo2", , False
End Sub
Sub StartCiclo3()
Verificador3.LabelHoraAtual.Caption = Format(Now, "hh:mm:ss")
RunWhen = Now + TimeSerial(0, 0, 1)
Application.OnTime RunWhen, "StartCiclo3", , True 'Aqui está o loop'
End Sub
Sub StopCiclo3()
    Application.OnTime RunWhen, "StartCiclo3", , False
End Sub

