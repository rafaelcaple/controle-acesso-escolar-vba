Attribute VB_Name = "Relogio"
Public RunWhen As Double

Sub StartCiclo()
    Verificador.LabelHoraAtual.Caption = Format(Now, "hh:mm:ss")
    RunWhen = Now + TimeSerial(0, 0, 1)
    Application.OnTime RunWhen, "StartCiclo", , True 'Aqui esta o loop
End Sub

Sub StopCiclo()
    Application.OnTime RunWhen, "StartCiclo", , False
End Sub

Sub StartCiclo2()
    VerificadorSaida.LabelHoraAtual.Caption = Format(Now, "hh:mm:ss")
    RunWhen = Now + TimeSerial(0, 0, 1)
    Application.OnTime RunWhen, "StartCiclo2", , True 'Aqui esta o loop
End Sub

Sub StopCiclo2()
    Application.OnTime RunWhen, "StartCiclo2", , False
End Sub
