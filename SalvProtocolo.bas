Attribute VB_Name = "SalvProtocolo"
Sub SalvaProtocolo()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Dim titulo As String
    Dim codForne As String
    Dim id As String
    
    titulo = Range("J2")
    codForne = Range("D12")
    id = Left(ThisWorkbook.Path, 1)
      
        Call ocultaBotaoProtocolo
        
        Sheets("PROTOCOLO").Copy
        
        Select Case codForne
            Case 48910:
            ChDir _
                id & ":\01 Monitoria %2f Inspetoria %2f Administrativo\001 - OPERAÇÃO MULTIVAREJO\002 - PROTOCOLOS DE ENTRADA NO P.A\VAGNER ELETRO"
            ActiveWorkbook.SaveAs fileName:= _
                id & ":\01 Monitoria %2f Inspetoria %2f Administrativo\001 - OPERAÇÃO MULTIVAREJO\002 - PROTOCOLOS DE ENTRADA NO P.A\VAGNER ELETRO\Protocolo Entrada e Saída Postos_N°" & titulo & ".xlsx" _
                , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
            ActiveWindow.Close
            
            Case 2114:
            ChDir _
                id & ":\01 Monitoria %2f Inspetoria %2f Administrativo\001 - OPERAÇÃO MULTIVAREJO\002 - PROTOCOLOS DE ENTRADA NO P.A\GIMENEZ"
            ActiveWorkbook.SaveAs fileName:= _
                id & ":\01 Monitoria %2f Inspetoria %2f Administrativo\001 - OPERAÇÃO MULTIVAREJO\002 - PROTOCOLOS DE ENTRADA NO P.A\GIMENEZ\Protocolo Entrada e Saída Postos_N°" & titulo & ".xlsx" _
                , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
            ActiveWindow.Close
            
             Case 23279:
            ChDir _
                id & ":\01 Monitoria %2f Inspetoria %2f Administrativo\001 - OPERAÇÃO MULTIVAREJO\002 - PROTOCOLOS DE ENTRADA NO P.A\GIMENEZ"
            ActiveWorkbook.SaveAs fileName:= _
                id & ":\01 Monitoria %2f Inspetoria %2f Administrativo\001 - OPERAÇÃO MULTIVAREJO\002 - PROTOCOLOS DE ENTRADA NO P.A\GIMENEZ\Protocolo Entrada e Saída Postos_N°" & titulo & ".xlsx" _
                , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
            ActiveWindow.Close
            
            
             Case 25100:
            ChDir _
                id & ":\01 Monitoria %2f Inspetoria %2f Administrativo\001 - OPERAÇÃO MULTIVAREJO\002 - PROTOCOLOS DE ENTRADA NO P.A\GIMENEZ"
            ActiveWorkbook.SaveAs fileName:= _
                id & ":\01 Monitoria %2f Inspetoria %2f Administrativo\001 - OPERAÇÃO MULTIVAREJO\002 - PROTOCOLOS DE ENTRADA NO P.A\GIMENEZ\Protocolo Entrada e Saída Postos_N°" & titulo & ".xlsx" _
                , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
            ActiveWindow.Close
            
             Case 7642:
            ChDir _
                id & ":\01 Monitoria %2f Inspetoria %2f Administrativo\001 - OPERAÇÃO MULTIVAREJO\002 - PROTOCOLOS DE ENTRADA NO P.A\GIMENEZ"
            ActiveWorkbook.SaveAs fileName:= _
                id & ":\01 Monitoria %2f Inspetoria %2f Administrativo\001 - OPERAÇÃO MULTIVAREJO\002 - PROTOCOLOS DE ENTRADA NO P.A\GIMENEZ\Protocolo Entrada e Saída Postos_N°" & titulo & ".xlsx" _
                , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
            ActiveWindow.Close
            
             Case 3901:
            ChDir _
                id & ":\01 Monitoria %2f Inspetoria %2f Administrativo\001 - OPERAÇÃO MULTIVAREJO\002 - PROTOCOLOS DE ENTRADA NO P.A\GIMENEZ"
            ActiveWorkbook.SaveAs fileName:= _
                id & ":\01 Monitoria %2f Inspetoria %2f Administrativo\001 - OPERAÇÃO MULTIVAREJO\002 - PROTOCOLOS DE ENTRADA NO P.A\GIMENEZ\Protocolo Entrada e Saída Postos_N°" & titulo & ".xlsx" _
                , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
            ActiveWindow.Close
            
             Case 24333:
            ChDir _
                id & ":\01 Monitoria %2f Inspetoria %2f Administrativo\001 - OPERAÇÃO MULTIVAREJO\002 - PROTOCOLOS DE ENTRADA NO P.A\GIMENEZ"
            ActiveWorkbook.SaveAs fileName:= _
                id & ":\01 Monitoria %2f Inspetoria %2f Administrativo\001 - OPERAÇÃO MULTIVAREJO\002 - PROTOCOLOS DE ENTRADA NO P.A\GIMENEZ\Protocolo Entrada e Saída Postos_N°" & titulo & ".xlsx" _
                , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
            ActiveWindow.Close
             
            Case 5048:
            ChDir _
                id & ":\01 Monitoria %2f Inspetoria %2f Administrativo\001 - OPERAÇÃO MULTIVAREJO\002 - PROTOCOLOS DE ENTRADA NO P.A\MADSON"
            ActiveWorkbook.SaveAs fileName:= _
                id & ":\01 Monitoria %2f Inspetoria %2f Administrativo\001 - OPERAÇÃO MULTIVAREJO\002 - PROTOCOLOS DE ENTRADA NO P.A\MADSON\Protocolo Entrada e Saída Postos_N°" & titulo & ".xlsx" _
                , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
            ActiveWindow.Close
            
            Case 5016:
             ChDir _
                id & ":\01 Monitoria %2f Inspetoria %2f Administrativo\001 - OPERAÇÃO MULTIVAREJO\002 - PROTOCOLOS DE ENTRADA NO P.A\WP"
            ActiveWorkbook.SaveAs fileName:= _
                id & ":\01 Monitoria %2f Inspetoria %2f Administrativo\001 - OPERAÇÃO MULTIVAREJO\002 - PROTOCOLOS DE ENTRADA NO P.A\WP\Protocolo Entrada e Saída Postos_N°" & titulo & ".xlsx" _
                , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
            ActiveWindow.Close
            
            Case 3870:
             ChDir _
                id & ":\01 Monitoria %2f Inspetoria %2f Administrativo\001 - OPERAÇÃO MULTIVAREJO\002 - PROTOCOLOS DE ENTRADA NO P.A\WP"
            ActiveWorkbook.SaveAs fileName:= _
                id & ":\01 Monitoria %2f Inspetoria %2f Administrativo\001 - OPERAÇÃO MULTIVAREJO\002 - PROTOCOLOS DE ENTRADA NO P.A\WP\Protocolo Entrada e Saída Postos_N°" & titulo & ".xlsx" _
                , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
            ActiveWindow.Close
            
            Case 48166:
             ChDir _
                id & ":\01 Monitoria %2f Inspetoria %2f Administrativo\001 - OPERAÇÃO MULTIVAREJO\002 - PROTOCOLOS DE ENTRADA NO P.A\WP"
            ActiveWorkbook.SaveAs fileName:= _
                id & ":\01 Monitoria %2f Inspetoria %2f Administrativo\001 - OPERAÇÃO MULTIVAREJO\002 - PROTOCOLOS DE ENTRADA NO P.A\WP\Protocolo Entrada e Saída Postos_N°" & titulo & ".xlsx" _
                , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
            ActiveWindow.Close
            
            Case 3816:
            ChDir _
                id & ":\01 Monitoria %2f Inspetoria %2f Administrativo\001 - OPERAÇÃO MULTIVAREJO\002 - PROTOCOLOS DE ENTRADA NO P.A\CUSTOMIZA"
            ActiveWorkbook.SaveAs fileName:= _
                id & ":\01 Monitoria %2f Inspetoria %2f Administrativo\001 - OPERAÇÃO MULTIVAREJO\002 - PROTOCOLOS DE ENTRADA NO P.A\CUSTOMIZA\Protocolo Entrada e Saída Postos_N°" & titulo & ".xlsx" _
                , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
            ActiveWindow.Close
            
             Case 14048:
            ChDir _
                id & ":\01 Monitoria %2f Inspetoria %2f Administrativo\001 - OPERAÇÃO MULTIVAREJO\002 - PROTOCOLOS DE ENTRADA NO P.A\CUSTOMIZA"
            ActiveWorkbook.SaveAs fileName:= _
                id & ":\01 Monitoria %2f Inspetoria %2f Administrativo\001 - OPERAÇÃO MULTIVAREJO\002 - PROTOCOLOS DE ENTRADA NO P.A\CUSTOMIZA\Protocolo Entrada e Saída Postos_N°" & titulo & ".xlsx" _
                , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
            ActiveWindow.Close
            
            Case 66679:
             ChDir _
                id & ":\01 Monitoria %2f Inspetoria %2f Administrativo\001 - OPERAÇÃO MULTIVAREJO\002 - PROTOCOLOS DE ENTRADA NO P.A\VAGNER ELETRO"
            ActiveWorkbook.SaveAs fileName:= _
                id & ":\01 Monitoria %2f Inspetoria %2f Administrativo\001 - OPERAÇÃO MULTIVAREJO\002 - PROTOCOLOS DE ENTRADA NO P.A\VAGNER ELETRO\Protocolo Entrada e Saída Postos_N°" & titulo & ".xlsx" _
                , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
            ActiveWindow.Close
        End Select
    
        Sheets("PROTOCOLO").Select
       Call mostraBotaoProtocolo
       Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub


