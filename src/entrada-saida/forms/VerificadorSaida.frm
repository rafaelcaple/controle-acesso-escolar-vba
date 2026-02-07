VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} VerificadorSaida 
   Caption         =   "| CIDA | REGISTRO DE SAÍDA"
   ClientHeight    =   8910.001
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   12300
   OleObjectBlob   =   "VerificadorSaida.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "VerificadorSaida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'SAIDA
Private Sub TMatricula_Change()
    Do While Len(TMatricula.Text) = 15
    
    If Not IsNumeric(TMatricula.Value) Then
        TMatricula.Value = ""
        MsgBox "Apenas número são permitidos", , "| CIDA |"
        Exit Sub
    End If
    
    Call Verificar
    Loop
End Sub
Sub Verificar()
 On Error GoTo Erro
'VARIAVEIS
Dim Pasta As String
Dim ext As String
Dim underline As String
Dim senha As String

Pasta = ThisWorkbook.Worksheets("Config").Range("B4")
ext = ".jpg"
underline = "_"
senha = ThisWorkbook.Worksheets("Config").Range("B18")

Set Rng = Range("A2")
valor = TMatricula

    Do While Rng.Value <> ""
        Rng.Select
        
            
            If Str(Rng.Value) = Str(valor) Then
            
                'SE JÁ FOI REGISTRADO
                If Rng.Offset(0, 5).Value = "Sim" Then
                    
                    'Mensagem
                    Mensagem.Caption = ThisWorkbook.Worksheets("Config").Range("B13")
                    Mensagem.ForeColor = vbRed
                    NomeAluno.Caption = Rng.Offset(0, 1).Value
                    Turma.Caption = Rng.Offset(0, 2).Value
                    HoraRegistrada.Caption = Format(Rng.Offset(0, 6).Value, "HH:MM:SS")
                    
                    
                    'Atualiza a Foto
                    Foto.Visible = True
                    Foto2.Visible = False
                    VerificadorSaida.Foto.Picture = LoadPicture(Pasta & NomeAluno.Caption & underline & TMatricula & underline & Turma.Caption & ext)
                     
                    'Limpa o campo de digitação da matricula e seleciona
                    TMatricula = ""
                    TMatricula.SetFocus
                    
                'SE A ENTRADA NÃO FOI REGISTRADA ANTES DA SAIDA
                ElseIf Rng.Offset(0, 3).Value = "" Then
                
                    'Mensagem
                    Mensagem.Caption = ThisWorkbook.Worksheets("Config").Range("D11")
                    Mensagem.ForeColor = RGB(255, 124, 35)
                    NomeAluno.Caption = Rng.Offset(0, 1).Value
                    Turma.Caption = Rng.Offset(0, 2).Value
                    
                    'Limpa o campo de digitação da matricula e seleciona
                    TMatricula = ""
                    TMatricula.SetFocus
                       
                'SE AINDA NÃO FOI REGISTRADO
                Else
                
                    'Tira proteção da planilha para registrar
                    Worksheets("Controle").Unprotect Password:=senha
                    
                    'Faz o Registro na planilha
                    Rng.Offset(0, 5).Value = "Sim"
                    Rng.Offset(0, 6).Value = Time
                    Rng.Offset(0, 7).Value = Date
                    
                    'Protege
                    Worksheets("Controle").Protect Password:=senha, AllowFiltering:=True, AllowSorting:=True
                    
                    'Mensagem
                    Mensagem.Caption = ThisWorkbook.Worksheets("Config").Range("B11")
                    Mensagem.ForeColor = vbGreen

                    NomeAluno.Caption = Rng.Offset(0, 1).Value
                    Turma.Caption = Rng.Offset(0, 2).Value
                    HoraRegistrada.Caption = Format(Rng.Offset(0, 6).Value, "HH:MM:SS")
                    
                    'Atualiza a Foto
                    Foto.Visible = True
                    Foto2.Visible = False
                    VerificadorSaida.Foto.Picture = LoadPicture(Pasta & NomeAluno.Caption & underline & TMatricula & underline & Turma.Caption & ext)
                    
                    'Limpa o campo de digitação da matricula e seleciona
                    TMatricula = ""
                    TMatricula.SetFocus
                    
                    'Atualiza o N de pessoas registradas
                    Call Num
                    
                    Exit Sub
                End If
                    Exit Sub
            End If
         Set Rng = Rng.Offset(1, 0)
    Loop
        'Quando a matrícula não existir
      If Str(Rng.Value) <> Str(valor) Then
      
            'Mensagem
            Mensagem.Caption = ThisWorkbook.Worksheets("Config").Range("B15")
            Mensagem.ForeColor = RGB(255, 182, 0)
            
            NomeAluno.Caption = ""
            Turma.Caption = ""
            HoraRegistrada.Caption = ""

            'Atualiza a Foto
            Foto.Visible = False
            Foto2.Visible = True
                     
            'Limpa o campo de digitação da matricula e seleciona
            TMatricula = ""
            TMatricula.SetFocus
           Exit Sub
           End If
Erro:
            MsgBox "Erro: " & Err.Description
            Resume Next
End Sub
Private Sub UserForm_Initialize()
Application.ScreenUpdating = False
Call StartCiclo2
End Sub
Private Sub UserForm_Terminate()
    Application.ScreenUpdating = True
    Call StopCiclo2
    ActiveWindow.ScrollRow = 1
    
End Sub

Private Sub btnPesquisa_Click()
opcaoRegistro = 2
Pesquisa.Show
End Sub

