Attribute VB_Name = "Espelha_Dados"
'ESPELHA DADOS NA CONSULTA
Sub espelharDados()
    Application.ScreenUpdating = False
    Dim id As String
    id = Left(ThisWorkbook.Path, 1)

Workbooks.Open fileName:= _
        id & ":\01 Monitoria %2f Inspetoria %2f Administrativo\001 - OPERAÇÃO MULTIVAREJO\001  -  CONSULTA\CONSULTA BASE.xlsm"

Workbooks.Open fileName:= _
        id & ":\01 Monitoria %2f Inspetoria %2f Administrativo\001 - OPERAÇÃO MULTIVAREJO\005 - APLICATIVO\BASE DE DADOS.xlsx"

Workbooks("BASE DE DADOS").Worksheets("DADOS").Range("A3:Y60000").Copy
Workbooks("CONSULTA BASE").Worksheets("DATABASE").Range("A3").PasteSpecial xlPasteValues
Workbooks("CONSULTA BASE").Worksheets("DATABASE").Range("F1") = Date
Workbooks("CONSULTA BASE").Worksheets("DATABASE").Range("G1") = Format(Time, "hh:mm")

'Salva e Fecha Base de Consulta
Workbooks("CONSULTA BASE").Close savechanges:=True
Application.CutCopyMode = False
'Salva e Fecha Base de Dados
Workbooks("BASE DE DADOS").Close savechanges:=True
Application.CutCopyMode = False

Application.ScreenUpdating = True
MsgBox "DADOS ESPELHADOS COM SUCESSO!", vbInformation, "BACKUP DE DADOS"
Exit Sub
End Sub





