VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TelaPrincipal 
   Caption         =   "| CIDA |"
   ClientHeight    =   5955
   ClientLeft      =   105
   ClientTop       =   420
   ClientWidth     =   7290
   OleObjectBlob   =   "TelaPrincipal.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TelaPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub btnLimpar_Click()
Call Limpar
End Sub

Private Sub btnOpenVerificador_Click()
    Verificador.Show
End Sub

Private Sub btnSaida_Click()
    VerificadorSaida.Show
End Sub

Private Sub btnSalvar_Click()
    Call Salvar
End Sub

Private Sub btnSave_Click()
ActiveWorkbook.Save
End Sub

Private Sub UserForm_Activate()
txtNRegistrado.Caption = ThisWorkbook.Worksheets("Config").Range("I3")
txtNSaidas.Caption = ThisWorkbook.Worksheets("Config").Range("I6")
End Sub

Private Sub UserForm_Terminate()
    ActiveWindow.ScrollRow = 1
End Sub

