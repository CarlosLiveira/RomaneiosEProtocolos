Private Sub Workbook_AfterSave(ByVal Success As Boolean)
    Call ExportarCodigo
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
Range("H1") = ""
Sheets("POSTOS").Range("A17") = Date
ThisWorkbook.Save
End Sub

Private Sub Workbook_Open()
Call inicio
End Sub

