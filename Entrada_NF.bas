Attribute VB_Name = "Entrada_NF"
'Analisa existencia do rg na base de dados
Sub Entrada_NF()
    Application.ScreenUpdating = False
    'ABRE O BANCO DE DADOS
    Workbooks.Open fileName:=ThisWorkbook.Path & "/BASE DE DADOS.xlsx"
    
    Dim TotalCarregado As Integer
    Dim retorno As Range, retPostos As Range
    Dim i As Long, qtdGeral As Long
    Dim ws As Worksheet, wsBd As Worksheet, wsPostos As Worksheet
            Set ws = Workbooks("ENTRADA NF  CONTROLE MULTIVAREJO.xlsm").Sheets("BANCO DE DADOS")
            Set wsPostos = Workbooks("ENTRADA NF  CONTROLE MULTIVAREJO.xlsm").Sheets("POSTOS")
            Set wsBd = Workbooks("BASE DE DADOS.xlsx").Sheets("DADOS")
    Dim produto As Entrada
            Set produto = New Entrada
    F = WorksheetFunction.Count(ws.Columns("A")) + 2
    x = WorksheetFunction.Count(wsBd.Columns("A")) + 3
    i = 3
    
    For i = 3 To F
        produto.rg = ws.Cells(i, 1)
        produto.NF = ws.Cells(i, 7)
        Workbooks("BASE DE DADOS.xlsx").Activate
        Set retorno = wsBd.Range("A:A").Find(What:=produto.rg, LookAt:=xlWhole)
             
        If Not (retorno Is Nothing) Then
            retorno.Offset(0, 6).Select
            Selection = produto.NF
            
        Else
        'Pega dados na planilha de preenchimento
        produto.rg = ws.Cells(i, 1)
        produto.CodFornecedor = ws.Cells(i, 3)
        produto.CodProduto = ws.Cells(i, 4)
        produto.DescricaoProd = ws.Cells(i, 5)
        produto.CustoUnitario = ws.Cells(i, 6)
        produto.NF = ws.Cells(i, 7)
        produto.Serie = ws.Cells(i, 8)
        
        'Pega dados da planilha POSTOS
        Workbooks("ENTRADA NF  CONTROLE MULTIVAREJO.xlsm").Activate
        Sheets("POSTOS").Visible = True
        Sheets("POSTOS").Select
        Set retPostos = wsPostos.Range("A:A").Find(What:=produto.CodFornecedor, LookAt:=xlWhole)
             
        If Not (retPostos Is Nothing) Then
            retPostos.Offset(0, 2).Select
            produto.Posto = Selection
            retPostos.Offset(0, 3).Select
            produto.Fornecedor = Selection
            retPostos.Offset(0, 4).Select
            produto.Analista = Selection
          
        'Insere dados no banco de dados
        Workbooks("BASE DE DADOS.xlsx").Activate
        wsBd.Cells(x, 1) = produto.rg
        wsBd.Cells(x, 2) = produto.Fornecedor
        wsBd.Cells(x, 3) = produto.CodFornecedor
        wsBd.Cells(x, 4) = produto.CodProduto
        wsBd.Cells(x, 5) = produto.DescricaoProd
        wsBd.Cells(x, 6) = produto.CustoUnitario
        wsBd.Cells(x, 7) = produto.NF
        wsBd.Cells(x, 8) = produto.Serie
        wsBd.Cells(x, 9) = produto.Posto
        wsBd.Cells(x, 10) = produto.Analista
        wsBd.Cells(x, 11) = Date 'produto.DataEntrada
        produto.MesEntrada = MonthName(Month(Date))
        wsBd.Cells(x, 12) = UCase(produto.MesEntrada)
        wsBd.Cells(x, 13) = "TRIAGEM CQ"
        wsBd.Cells(x, 14) = 0
        wsBd.Cells(x, 15) = "Até 20 dias"
        wsBd.Cells(x, 16) = "ABERTO"
        wsBd.Cells(x, 19) = "CQ"
         
        x = x + 1
        End If
        End If
        TotalCarregado = i - 2
    Next
    
    qtdGeral = WorksheetFunction.Count(Workbooks("BASE DE DADOS.xlsx").Worksheets("DADOS").Range("A:A"))
    
    'SALVA O BANCO DE DADOS
    Workbooks("BASE DE DADOS.xlsx").Save
    'FECHA O BANCO DE DADOS
    Workbooks("BASE DE DADOS.xlsx").Close
    'SELECIONA A PLANILHA PRINCIPAL
    Workbooks("ENTRADA NF  CONTROLE MULTIVAREJO.xlsm").Activate
    Sheets("POSTOS").Visible = False
    Sheets("BANCO DE DADOS").Select
    Range("J1") = qtdGeral
    MsgBox i - 3 & " - DADOS CARREGADOS COM SUCESSO", vbInformation, "AVISO"
    Range("H1") = TotalCarregado + Range("H1")
    Call limpaLista
    Range("A3").Select
    Application.ScreenUpdating = True
End Sub

'Limpa a área de colagem do aplicativo
Sub limpaLista()
Sheets("BANCO DE DADOS").Range("A3:AA1000") = ""
End Sub
