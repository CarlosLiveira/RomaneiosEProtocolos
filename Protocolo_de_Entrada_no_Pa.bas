Attribute VB_Name = "Protocolo_de_Entrada_no_Pa"
'Carrega Protocolo de Entrada no PA
Sub Entrada_No_Posto()
    Application.ScreenUpdating = False
    'ABRE O BANCO DE DADOS
    Workbooks.Open fileName:=ThisWorkbook.Path & "/BASE DE DADOS.xlsx"
    
    Dim retorno As Range
    Dim i As Long
    Dim wsProt As Worksheet, wsBd As Worksheet
            Set wsProt = Workbooks("ENTRADA NF  CONTROLE MULTIVAREJO.xlsm").Sheets("PROTOCOLO")
            Set wsBd = Workbooks("BASE DE DADOS.xlsx").Sheets("DADOS")
    Dim produto As Entrada
            Set produto = New Entrada
    F = WorksheetFunction.Count(wsProt.Columns("B")) + 11
    
    For i = 12 To F
        produto.rg = wsProt.Cells(i, 2)
        Workbooks("BASE DE DADOS.xlsx").Activate
        Set retorno = wsBd.Range("A:A").Find(What:=produto.rg, LookAt:=xlWhole)
             
        If Not (retorno Is Nothing) Then
            'Código Fornecedor
            retorno.Offset(0, 2).Select
            wsProt.Cells(i, 4) = Selection
            'Código do Produto
            retorno.Offset(0, 3).Select
            wsProt.Cells(i, 5) = Selection
            'Descrição do produto
            retorno.Offset(0, 4).Select
            wsProt.Cells(i, 6) = Selection
            'Série
            retorno.Offset(0, 7).Select
            wsProt.Cells(i, 7) = Selection
            'Nota Fiscal
            retorno.Offset(0, 6).Select
            wsProt.Cells(i, 8) = Selection
            'Posições
            retorno.Offset(0, 12) = "ENVIADO AO POSTO"
            'Data Expedida
            retorno.Offset(0, 16) = Date
            'Posto ou CQ
            retorno.Offset(0, 18) = "POSTO"
            
            'Número do Protocolo
            retorno.Offset(0, 23).Select
            Selection = wsProt.Range("J2")
            'Série
             wsProt.Cells(i, 7) = wsProt.Cells(i, 3)
            
        Else
        'Pega dados na planilha de preenchimento
        MsgBox "RG - " & produto.rg & " - NÃO CADASTRADO!", vbCritical, "AVISO"
        
        End If
    Next
    'Ativa o arquivo que contém o protocolo
    Workbooks("ENTRADA NF  CONTROLE MULTIVAREJO.xlsm").Sheets("PROTOCOLO").Activate
    'Chama a fução que oculta as linhas vazias do protocolo
    Call ocultaProt
    Call SalvaProtocolo
    'IMPRIME PROTOCOLO
ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
     'SALVA O BANCO DE DADOS
    Workbooks("BASE DE DADOS.xlsx").Save
    'FECHA O BANCO DE DADOS
    Workbooks("BASE DE DADOS.xlsx").Close
    MsgBox "PROTOCOLO SALVO COM SUCESSO", vbInformation, "AVISO"
    Application.ScreenUpdating = True
End Sub

'Limpa protocolo
Sub LimpProtocolo()
Attribute LimpProtocolo.VB_ProcData.VB_Invoke_Func = " \n14"
    Range("B12:K111").Select
    Selection.ClearContents
    Range("B12").Select
    
    Range("J2") = Range("J2") + 1
    Range("B12").Select
    Call mostraProt
End Sub
'Oculta Linhas Vazias no Protocolo
Sub ocultaProt()
  Dim i As Integer
  Application.ScreenUpdating = False
    With Sheets("PROTOCOLO")
      .Cells.EntireRow.Hidden = False
          For i = 12 To 111
            Select Case .Range("B" & i).Value
              Case 0
              .Rows(i & ":" & i).EntireRow.Hidden = True
            End Select
         Next i
    End With
    Application.ScreenUpdating = True
End Sub
'Exibe Linhas Vazias No Protocolo
 Sub mostraProt()
  Dim i As Integer
  Application.ScreenUpdating = False
    With Sheets("PROTOCOLO")
      .Cells.EntireRow.Hidden = False
          For i = 12 To 111
            Select Case .Range("B" & i).Value
              Case 0
              .Rows(i & ":" & i).EntireRow.Hidden = False
            End Select
         Next i
    End With
    Application.ScreenUpdating = True
End Sub

'Oculta botões do Protocolo
Sub ocultaBotaoProtocolo()
ActiveSheet.Shapes("LimpaProt").Visible = False
ActiveSheet.Shapes("Carrega_dados_Prot").Visible = False
ActiveSheet.Shapes("Edita_Txt_Prot").Visible = False
ActiveSheet.Shapes("Volta_Bd_Prot").Visible = False
End Sub
'Mostra botões do Protocolo
Sub mostraBotaoProtocolo()
ActiveSheet.Shapes("LimpaProt").Visible = True
ActiveSheet.Shapes("Carrega_dados_Prot").Visible = True
ActiveSheet.Shapes("Edita_Txt_Prot").Visible = True
ActiveSheet.Shapes("Volta_Bd_Prot").Visible = True
Range("B12").Select
End Sub

'EDITA TXT PROTOCOLO
Sub editaTXTProt()
    Selection.TextToColumns Destination:=Range("B12"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
        :=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
End Sub

