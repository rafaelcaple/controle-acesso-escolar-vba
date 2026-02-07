VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Pesquisa 
   Caption         =   "| CIDA | PESQUISA"
   ClientHeight    =   3825
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8985.001
   OleObjectBlob   =   "Pesquisa.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Pesquisa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private turmaSelecionada As String
Private Sub UserForm_Activate()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Refeitorio")
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    Dim rngNames As Range
    Set rngNames = ws.Range("B2:B" & lastRow)
    
    ' Obter a lista de nomes em uma matriz bidimensional
    Dim namesArray As Variant
    namesArray = rngNames.Value
    
    ' Converter a matriz bidimensional para uma matriz unidimensional
    Dim namesList() As String
    ReDim namesList(1 To UBound(namesArray))
    Dim i As Long
    For i = 1 To UBound(namesArray)
        namesList(i) = namesArray(i, 1)
    Next i
    
    ' Ordenar a matriz unidimensional em ordem alfabética usando o QuickSort
    QuickSort namesList, LBound(namesList), UBound(namesList)
    
    ' Atribuir a matriz unidimensional ao ComboBox cbNome
    cbNome.List = namesList
    PreencherComboBoxTurma
End Sub
Private Sub cbNome_Change()
Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Refeitorio")
    
Dim selectedName As String
    selectedName = cbNome.Value
    
    Dim colunaPesquisa As Range
    Set colunaPesquisa = ws.Range("B2:B" & ws.Cells(ws.Rows.Count, "B").End(xlUp).Row)
    
    Dim celulaEncontrada As Range
    Set celulaEncontrada = colunaPesquisa.Find(selectedName, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not celulaEncontrada Is Nothing Then
        txtMat.Caption = celulaEncontrada.Offset(0, -1).Value
    Else
        txtMat.Caption = "Não encontrada"
    End If
    
End Sub

Sub PreencherComboBoxTurma()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Refeitorio")
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    
    Dim rngTurmas As Range
    Set rngTurmas = ws.Range("C2:C" & lastRow)
    
    ' Criar um objeto Dictionary para armazenar turmas únicas
    Dim dictTurmas As Object
    Set dictTurmas = CreateObject("Scripting.Dictionary")
    
    Dim celula As Range
    For Each celula In rngTurmas
        ' Adicionar a turma ao Dictionary (turmas repetidas não serão adicionadas novamente)
        dictTurmas(celula.Value) = 1
    Next celula
    
    ' Limpar o ComboBox de turmas
    cbTurma.Clear
    
    ' Adicionar turmas únicas ao ComboBox
    Dim Turma As Variant
    For Each Turma In dictTurmas.Keys
        cbTurma.AddItem Turma
    Next Turma
End Sub

Private Sub cbTurma_Change()
     If cbTurma.Value <> "" Then
        turmaSelecionada = cbTurma.Value
        FiltrarNomesPorTurma
    Else
        ' Se a turma estiver vazia, limpa o ComboBox de nomes e mostra todos os nomes
        turmaSelecionada = ""
        MostrarTodosOsNomes
    End If
End Sub

Sub FiltrarNomesPorTurma()
     Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Refeitorio")

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    Dim rngNames As Range
    Set rngNames = ws.Range("B2:B" & lastRow)

    Dim rngTurmas As Range
    Set rngTurmas = ws.Range("C2:C" & lastRow)

    ' Obter a lista de nomes e turmas em matrizes
    Dim namesArray As Variant
    Dim turmasArray As Variant
    namesArray = rngNames.Value
    turmasArray = rngTurmas.Value

    ' Limpar o ComboBox de nomes
    cbNome.Clear

    ' Adicionar nomes ao ComboBox baseado na turma selecionada
    Dim i As Long
    For i = 1 To UBound(namesArray)
        If turmasArray(i, 1) = turmaSelecionada Then
            cbNome.AddItem namesArray(i, 1)
        End If
    Next i
End Sub
Sub MostrarTodosOsNomes()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Refeitorio")

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    Dim rngNames As Range
    Set rngNames = ws.Range("B2:B" & lastRow)

    Dim namesArray As Variant
    namesArray = rngNames.Value

    ' Cria uma matriz para armazenar os nomes
    Dim nomes() As String
    ReDim nomes(1 To UBound(namesArray))

    Dim i As Long
    Dim contadorNomes As Long
    contadorNomes = 0

    ' Percorre a lista de nomes e armazena-os na matriz
    For i = 1 To UBound(namesArray)
        If Not IsEmpty(namesArray(i, 1)) Then
            contadorNomes = contadorNomes + 1
            nomes(contadorNomes) = namesArray(i, 1)
        End If
    Next i

    ' Redimensiona a matriz para remover os elementos vazios (caso existam)
    ReDim Preserve nomes(1 To contadorNomes)

    ' Ordena a matriz de nomes em ordem alfabética
    Call QuickSort(nomes, 1, UBound(nomes))

    cbNome.Clear

    ' Adiciona os nomes ordenados ao ComboBox de nomes
    For i = 1 To UBound(nomes)
        cbNome.AddItem nomes(i)
    Next i
End Sub

' Algoritmo de ordenação QuickSort para ordenar a matriz de nomes em ordem alfabética
Sub QuickSort(ByRef arr() As String, ByVal left As Long, ByVal right As Long)
    Dim i As Long
    Dim j As Long
    Dim pivot As String
    Dim temp As String

    i = left
    j = right
    pivot = arr((left + right) \ 2)

    While i <= j
        While StrComp(arr(i), pivot, vbTextCompare) < 0
            i = i + 1
        Wend

        While StrComp(arr(j), pivot, vbTextCompare) > 0
            j = j - 1
        Wend

        If i <= j Then
            temp = arr(i)
            arr(i) = arr(j)
            arr(j) = temp
            i = i + 1
            j = j - 1
        End If
    Wend

    If left < j Then QuickSort arr, left, j
    If i < right Then QuickSort arr, i, right
End Sub
Private Sub btnRegistrarPq_Click()

Select Case opcaoRegistro
    Case 1
        Verificador.TMatricula.Text = txtMat.Caption
    Case 2
        Verificador2.TMatricula.Text = txtMat.Caption
    Case 3
        Verificador3.TMatricula.Text = txtMat.Caption
End Select

End Sub

Private Sub cbTurma_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
HookListBoxScroll Me, Me.cbTurma
End Sub
Private Sub cbNome_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
HookListBoxScroll Me, Me.cbNome
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
UnhookListBoxScroll
End Sub


