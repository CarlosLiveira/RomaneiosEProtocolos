Attribute VB_Name = "Geral"
'Vai para a planilha do DADOS
Sub abaBancoDados()
Sheets("BANCO DE DADOS").Visible = True
Sheets("BANCO DE DADOS").Select
Sheets("PROTOCOLO").Visible = False
Sheets("ROMANEIO").Visible = False
Range("A1").Select
End Sub
'Abre o formulário den opções
Sub abreForm()
Inicial.Show
End Sub
'Calcula e atualiza os dias na posição
Sub dataDeOntem()
Dim dtHoje As Date
Dim dtOntem As Date
dtHoje = Date
dtOntem = Sheets("POSTOS").Range("A17")

If dtOntem = dtHoje Then
MsgBox "DIAS NA POSIÇÃO JÁ ATUALIZADOS HOJE!", vbInformation, "ATUALIZAÇÃO!"
Else
Call atualizaDiasNaPosicao
End If
Exit Sub
End Sub
'Calcula e atualiza Range de dias na posição
Sub atualizaDiasNaPosicao()
    Application.ScreenUpdating = False
    
    Dim dias As Long
    Dim lin1 As Long
    lin1 = 3
    'ABRE BANCO DE DADOS
    Workbooks.Open fileName:=ThisWorkbook.Path & "/BASE DE DADOS.xlsx"
    'Workbooks("BASE DE DADOS.xlsx").Activate
    Do Until Sheets("DADOS").Cells(lin1, 1) = ""
          
        If Sheets("DADOS").Cells(lin1, 16) <> "FECHADO" And Sheets("DADOS").Cells(lin1, 13) = "ENVIADO AO POSTO" Then
                dias = Date - Sheets("DADOS").Cells(lin1, 17)
            Else
            If Sheets("DADOS").Cells(lin1, 16) <> "FECHADO" And Sheets("DADOS").Cells(lin1, 13) = "TRIAGEM CQ" Then
                dias = Date - Sheets("DADOS").Cells(lin1, 11)
            End If
        End If
 
        Sheets("DADOS").Cells(lin1, 14) = dias
        If Sheets("DADOS").Cells(lin1, 16) = "FECHADO" Then
            Sheets("DADOS").Cells(lin1, 14) = "" 'Dias na Posição
            Sheets("DADOS").Cells(lin1, 15) = "" 'Range de Dias Na posição
        Else
        Select Case dias
            Case Is < 21
                Sheets("DADOS").Cells(lin1, 15) = "Até 20 dias"
            Case Is < 31
                Sheets("DADOS").Cells(lin1, 15) = "De 21 a 30 dias"
            Case Is < 61
                Sheets("DADOS").Cells(lin1, 15) = "De 31 a 60 dias"
            Case Is > 60
                Sheets("DADOS").Cells(lin1, 15) = "Acima de 60 dias"
        End Select
        End If
        lin1 = lin1 + 1
    Loop
    Sheets("DADOS").Range("F1") = Date
    Sheets("DADOS").Range("G1") = Time
    
    'SALVA BANCO DE DADOS
    Workbooks("BASE DE DADOS.xlsx").Save
    'FECHA BANCO DE DADOS
    Workbooks("BASE DE DADOS.xlsx").Close
    Workbooks("ENTRADA NF  CONTROLE MULTIVAREJO.xlsm").Activate
    MsgBox "DIAS NA POSIÇÃO ATUALIZADOS COM SUCESSO!", vbInformation, "ATUALIZAÇÃO"
    Application.ScreenUpdating = True
End Sub

'INICIALIZA APLICATIVO
Sub inicio()
Dim Entrada As String
Dim data As String
Sheets("PROTOCOLO").Range("K1") = Date
Sheets("ROMANEIO").Range("K1") = Date
Call dataDeOntem
Inicial.Show
End Sub
'Mostrar botões
Sub Mostrar_Botoes()
    Application.ScreenUpdating = False
        Sheets("PROTOCOLO").Select
        Call mostraBotaoProtocolo
        Sheets("ROMANEIO").Select
        Call mostraBotaoRomaneio
        Sheets("BANCO DE DADOS").Select
        MsgBox "BOTÃO EXIBIDO COM SUCESSO", vbInformation, "EXIBIÇÃO DE BOTÃO ATIVA"
    Application.ScreenUpdating = True
End Sub


Sub Retornar_a_tela_inicial()
    Sheets("BANCO DE DADOS").Visible = True
    Sheets("BANCO DE DADOS").Select
    Sheets("POSTOS").Visible = False
    'Sheets("ROMANEIO").Visible = False
    'Sheets("PROTOCOLO").Visible = False
End Sub
