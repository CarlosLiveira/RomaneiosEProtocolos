Attribute VB_Name = "SalvRomaneio"
Sub SalvaRomaneio()
Attribute SalvaRomaneio.VB_ProcData.VB_Invoke_Func = " \n14"
Application.ScreenUpdating = False

Dim titulo As String
Dim id As String
titulo = Range("K2")
id = Left(ThisWorkbook.Path, 1)

  Call ocultaBotaoRomaneio
  
    Sheets("ROMANEIO").Copy
    ChDir _
        id & ":\01 Monitoria %2f Inspetoria %2f Administrativo\001 - OPERAÇÃO MULTIVAREJO\003 - ROMANEIOS DE ENVIO\PENDENTE DE ENVIO PARA WESLEY"
    ActiveWorkbook.SaveAs fileName:= _
        id & ":\01 Monitoria %2f Inspetoria %2f Administrativo\001 - OPERAÇÃO MULTIVAREJO\003 - ROMANEIOS DE ENVIO\PENDENTE DE ENVIO PARA WESLEY\ROMANEIO_" & titulo & ".xlsx" _
        , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    ActiveWindow.Close
     
   Sheets("ROMANEIO").Select
   Call mostraBotaoRomaneio
   Application.ScreenUpdating = True
End Sub
