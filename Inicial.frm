VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Inicial 
   Caption         =   "OPERAÇÃO MULTVAREJO"
   ClientHeight    =   3690
   ClientLeft      =   50
   ClientTop       =   380
   ClientWidth     =   6240
   OleObjectBlob   =   "Inicial.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Inicial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Chama a função para atualizar dias na posição
Private Sub Bt_Atualiza_Dias_Click()
Call atualizaDiasNaPosicao
End Sub
'Exibe dados relativos aos postos para alteração
Private Sub Btn_Atualizar_Dados_Click()
    Sheets("POSTOS").Visible = True
    Sheets("POSTOS").Select
    Sheets("BANCO DE DADOS").Visible = xlSheetVeryHidden
    Sheets("ROMANEIO").Visible = False
    Sheets("PROTOCOLO").Visible = False
End Sub

Private Sub CommandButton1_Click()
Call selectOperacao
End Sub

