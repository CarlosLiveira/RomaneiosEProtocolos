Attribute VB_Name = "Romaneio_de_Envio_para_Multi"
' EditaTXT
Sub EditaTXT()
Application.DisplayAlerts = False
    Range("B13:B112").Select
    Selection.TextToColumns Destination:=Range("B13"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 1), Array(10, 1), Array(12, 1)), TrailingMinusNumbers _
        :=True
    Range("B13").Select
Application.DisplayAlerts = True
End Sub

'Limpa Romaneio
Sub LimpRoma()
Dim num As String
    Range("B13:K112").Select
    Selection.ClearContents
    num = Range("K2")
    Range("K2") = Left(num, 4) + 1 & "L"
    Range("B13").Select
    Call mostraRoma
End Sub
'Mostra Linhas do Romaneio
 Sub ocultaRoma()
  Dim i As Integer
  Application.ScreenUpdating = False
    With Sheets("ROMANEIO")
      .Cells.EntireRow.Hidden = False
          For i = 13 To 112
            Select Case .Range("B" & i).Value
              Case 0
              .Rows(i & ":" & i).EntireRow.Hidden = True
            End Select
         Next i
    End With
    Application.ScreenUpdating = True
End Sub
'Mostra Linhas do Romaneio
 Sub mostraRoma()
  Dim i As Integer
  Application.ScreenUpdating = False
    With Sheets("ROMANEIO")
      .Cells.EntireRow.Hidden = False
          For i = 13 To 112
            Select Case .Range("B" & i).Value
              Case 0
              .Rows(i & ":" & i).EntireRow.Hidden = False
            End Select
         Next i
    End With
    Application.ScreenUpdating = True
End Sub

'Oculta botões do Romaneio
Sub ocultaBotaoRomaneio()
    ActiveSheet.Shapes("limpaRomaneio").Visible = False
    ActiveSheet.Shapes("CarregaRomaneio").Visible = False
    ActiveSheet.Shapes("Edita_Txt_Roma").Visible = False
    ActiveSheet.Shapes("Volta_Bd_Roma").Visible = False
End Sub
'Mostra botões do Romaneio
Sub mostraBotaoRomaneio()
    ActiveSheet.Shapes("limpaRomaneio").Visible = True
    ActiveSheet.Shapes("CarregaRomaneio").Visible = True
    ActiveSheet.Shapes("Edita_Txt_Roma").Visible = True
    ActiveSheet.Shapes("Volta_Bd_Roma").Visible = True
    Range("B13").Select
End Sub

