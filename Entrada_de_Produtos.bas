Attribute VB_Name = "Entrada_de_Produtos"
'Entrada de Produtos da Multivarejo para a Via Varejo
Sub EntradaMultivarejo()

Dim lin As Long
Dim linha As Long
Dim mes As String
Dim contDados As Long

lin = WorksheetFunction.CountA(Columns("I")) + 2
linha = 2
contDados = 0

Do Until Sheets("BANCO DE DADOS").Cells(lin, 1) = ""

If Sheets("BANCO DE DADOS").Cells(lin, 3) = Sheets("POSTOS").Cells(linha, 1) Then
Sheets("BANCO DE DADOS").Cells(lin, 2) = Sheets("POSTOS").Cells(linha, 4) 'Fornecedor
Sheets("BANCO DE DADOS").Cells(lin, 9) = Sheets("POSTOS").Cells(linha, 3) 'Posto
Sheets("BANCO DE DADOS").Cells(lin, 10) = Sheets("POSTOS").Cells(linha, 5) 'Analista
lin = lin + 1
linha = 2
contDados = contDados + 1
Else
If Sheets("POSTOS").Cells(linha, 1).Value <> "" Then
linha = linha + 1
Else
If Sheets("POSTOS").Cells(linha, 1).Value <> Sheets("BANCO DE DADOS").Cells(lin, 3) Then
MsgBox "Código de Fornecedor Inválido", vbCritical, "DADOS INVÁLIDOS!"
Exit Sub
End If
End If
End If
Loop
MsgBox contDados & "  DADOS CARREGADOS", vbInformation, "DADOS CAREEGADOS!"
contDados = 0
Exit Sub
End Sub
'MENU INICIAL
Sub selectOperacao()
Application.ScreenUpdating = False
    If Inicial.OptionRomaneioDeEntrada = True Then
        Range("A3").Select
        Sheets("ROMANEIO").Visible = False
        Sheets("PROTOCOLO").Visible = False
    End If
    
    If Inicial.OptionRomaneioDeSaida = True Then
        Sheets("ROMANEIO").Visible = True
        Sheets("ROMANEIO").Select
        Sheets("PROTOCOLO").Visible = False
        Range("B13").Select
    End If
    'seleciona entrada posto
    If Inicial.OptionProtEntradaPA = True Then
        Sheets("PROTOCOLO").Visible = True
        Sheets("PROTOCOLO").Select
        Range("B12").Select
        Sheets("ROMANEIO").Visible = False
    End If
        Unload Inicial
Application.ScreenUpdating = True
End Sub



