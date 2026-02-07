Attribute VB_Name = "SalvarLimpar"

Sub Salvar()
 
Dim saveDate As Date
Dim saveTime As Variant
Dim formatTime As String
Dim formatDate As String
Dim backupFolder As String
Dim FileExt As String
Dim ThisFileName As String
Dim FileName As String

saveDate = Now
FileExt = ".xlsm"

formatDate = Format(saveDate, "YYYY-MM-DD hh-mm")

Application.DisplayAlerts = False
backupFolder = ThisWorkbook.Worksheets("Config").Range("B3")
FileName = ThisWorkbook.Worksheets("Config").Range("B6")

ThisFileName = FileName & " " & formatDate & FileExt

ActiveWorkbook.Save

ActiveWorkbook.SaveCopyAs FileName:=backupFolder & ThisFileName
Application.DisplayAlerts = True
MsgBox "Salvo com sucesso! No diretório " & backupFolder, , "| CIDA |"

ActiveWorkbook.Save

End Sub

Sub Limpar()
Dim senha As String
senha = ThisWorkbook.Worksheets("Config").Range("B18")

Worksheets("Refeitorio").Unprotect Password:=senha

Worksheets("Refeitorio").Range("D2:D1000").ClearContents
Worksheets("Refeitorio").Range("E2:E1000").ClearContents
Worksheets("Refeitorio").Range("F2:F1000").ClearContents
Worksheets("Refeitorio").Range("G2:G1000").ClearContents
Worksheets("Refeitorio").Range("H2:H1000").ClearContents
Worksheets("Refeitorio").Range("I2:I1000").ClearContents
Worksheets("Refeitorio").Range("J2:J1000").ClearContents

Worksheets("Refeitorio").Protect Password:=senha, AllowFiltering:=True, AllowSorting:=True
Call Num
End Sub


