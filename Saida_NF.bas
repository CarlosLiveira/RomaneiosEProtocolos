Attribute VB_Name = "Saida_NF"
'Carrega Romaneio de Envio para Multivarejo
Sub Saida_Do_Posto()
    Application.ScreenUpdating = False
    Dim id As String
    id = Left(ThisWorkbook.Path, 1)
    Workbooks.Open (id & ":\01 Monitoria %2f Inspetoria %2f Administrativo\001 - OPERAÇÃO MULTIVAREJO\005 - APLICATIVO\BASE DE DADOS.xlsx")
    
    Dim CustoSimulado As Currency
    Dim CodProduto As String
    Dim mes As String
    Dim retorno As Range, retValor As Range
    Dim i As Long
    Dim tipoSaldo As String
    Dim wsRoma As Worksheet, wsBd As Worksheet, wsBsVal As Worksheet
            Set wsRoma = Workbooks("ENTRADA NF  CONTROLE MULTIVAREJO.xlsm").Sheets("ROMANEIO")
            Set wsBd = Workbooks("BASE DE DADOS.xlsx").Sheets("DADOS")
            Set wsBsVal = Workbooks("ENTRADA NF  CONTROLE MULTIVAREJO.xlsm").Sheets("BASE_VALORES")
            
    Dim produto As Entrada
            Set produto = New Entrada
    F = WorksheetFunction.Count(wsRoma.Columns("B")) + 12
    
    For i = 13 To F
        produto.rg = wsRoma.Cells(i, 2)
        Workbooks("BASE DE DADOS.xlsx").Activate
        Set retorno = wsBd.Range("A:A").Find(What:=produto.rg, LookAt:=xlWhole)
            
        If Not (retorno Is Nothing) Then
            '****** Romaneio Recebe dados do Banco
            'Código Fornecedor
            retorno.Offset(0, 2).Select
            wsRoma.Cells(i, 5) = Selection
            'Código do Produto
            retorno.Offset(0, 3).Select
            wsRoma.Cells(i, 6) = Selection
            CodProduto = Selection
            
            'Busca Valor de custo simulado
            Workbooks("ENTRADA NF  CONTROLE MULTIVAREJO.xlsm").Activate
            Sheets("BASE_VALORES").Visible = True
            Sheets("BASE_VALORES").Select
            Set retValor = wsBsVal.Range("A:A").Find(What:=CodProduto, LookAt:=xlWhole)
            If Not (retValor Is Nothing) Then
            retValor.Offset(0, 3).Select
            CustoSimulado = Selection
            End If
            Workbooks("BASE DE DADOS.xlsx").Activate
            
            'Descrição do produto
            retorno.Offset(0, 4).Select
            wsRoma.Cells(i, 7) = Selection
            'Nota Fiscal
            retorno.Offset(0, 22).Select
            Selection = wsRoma.Cells(i, 4)
            wsRoma.Cells(i, 8) = wsRoma.Cells(i, 4)
            
            'Preenche o tipo de saldo No Romaneio(STATUS)
            tipoSaldo = UCase(wsRoma.Cells(i, 3))
            wsRoma.Range("H" & i) = wsRoma.Range("D" & i) 'NF
            Select Case tipoSaldo
            Case Is = "A":
            wsRoma.Cells(i, 9) = "SALDO A"
            Case Is = "B":
            wsRoma.Cells(i, 9) = "SALDO B"
            Case Is = "C":
            wsRoma.Cells(i, 9) = "SALDO C"
            Case Is = "D":
            wsRoma.Cells(i, 9) = "DEVOLUÇAO"
            Case Is = "E":
            wsRoma.Cells(i, 9) = "ESTOQUE"
            Case Is = "R":
            wsRoma.Cells(i, 9) = "REPROVADO"
            Case Is = "RET":
            wsRoma.Cells(i, 9) = "RETORNO"
            End Select
            
            '******Atualiza Dados do Banco
            'Posições
            retorno.Offset(0, 12) = "PROCESSO DE CONSERTO ENCERRADO"
            'Data Expedida
            retorno.Offset(0, 13) = ""
            'Data Expedida
            retorno.Offset(0, 14) = ""
            'Status (Aberto/Fechado)
            retorno.Offset(0, 15) = "FECHADO"
            'Destino tipo saldo
            retorno.Offset(0, 17) = wsRoma.Cells(i, 9)
            'Data Expedida
            retorno.Offset(0, 16) = Date
            'Mês saída
            mes = MonthName(Month(Date))
            retorno.Offset(0, 20) = UCase(mes)
            'Custo Simulado
            retorno.Offset(0, 21) = CustoSimulado
            'Número do Romaneio
            retorno.Offset(0, 24).Select
            Selection = wsRoma.Range("K2")
            
        Else
        'Pega dados na planilha de preenchimento
        MsgBox "RG - " & produto.rg & " - NÃO CADASTRADO!", vbCritical, "AVISO"
        
        End If
    Next
    'Ativa o arquivo que contém o protocolo
    Workbooks("ENTRADA NF  CONTROLE MULTIVAREJO.xlsm").Sheets("ROMANEIO").Activate
    'Chama a fução que oculta as linhas vazias do protocolo
    Call ocultaRoma
    Call SalvaRomaneio
    'IMPRIME ROMANEIO
ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
    'SALVA BANCO DE DADOS
    Workbooks("BASE DE DADOS.xlsx").Save
    'FECHA BANCO DE DADOS
    Workbooks("BASE DE DADOS.xlsx").Close
    MsgBox "ROMANEIO SALVO COM SUCESSO", vbInformation, "AVISO"
    Sheets("BASE_VALORES").Visible = False
    Application.ScreenUpdating = True
End Sub

