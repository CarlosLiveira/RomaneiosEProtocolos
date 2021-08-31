VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Entrada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Esta Classe possiu sete propriedades: RG, Fornecedor,Codigo Fornecedor,Codigo Produto,Descricao Produto,Custo Unitario e NF
'representados pelas variáveis pRG, pFornecedor, pCodFornecedor, pCodProduto, pDescricaoProd, pCustoUnitario e pNF respectivamente.

Private pRG As String
Private pFornecedor As String
Private pCodFornecedor As String
Private pCodProduto As String
Private pDescricaoProd As String
Private pCustoUnitario As Currency
Private pNF As String
Private pSerie As String
Private pPosto As String
Private pAnalista As String
Private pDataEntrada As Date
Private pMesEntrada As String

'Abaixo seguem as declarações de propriedade de como são lidas e gravadas.

'Propriedade RG:
Public Property Get rg() As String
rg = pRG
End Property
Public Property Let rg(s As String)
pRG = s
End Property

'Propriedade Fornecedor:
Public Property Get Fornecedor() As String
Fornecedor = pFornecedor
End Property
Public Property Let Fornecedor(s As String)
pFornecedor = s
End Property

'Propriedade CodFornecedor:
Public Property Get CodFornecedor() As String
CodFornecedor = pCodFornecedor
End Property
Public Property Let CodFornecedor(s As String)
pCodFornecedor = s
End Property

'Propriedade CodProduto:
Public Property Get CodProduto() As String
CodProduto = pCodProduto
End Property
Public Property Let CodProduto(s As String)
pCodProduto = s
End Property

'Propriedade DescricaoProd:
Public Property Get DescricaoProd() As String
DescricaoProd = pDescricaoProd
End Property
Public Property Let DescricaoProd(s As String)
pDescricaoProd = s
End Property

'Propriedade CustoUnitario:
Public Property Get CustoUnitario() As Currency
CustoUnitario = pCustoUnitario
End Property
Public Property Let CustoUnitario(c As Currency)
pCustoUnitario = c
End Property

'Propriedade CustoUnitario:
Public Property Get NF() As String
NF = pNF
End Property
Public Property Let NF(s As String)
pNF = s
End Property

'Propriedade Serie:
Public Property Get Serie() As String
Serie = pSerie
End Property
Public Property Let Serie(s As String)
pSerie = s
End Property

'Propriedade Posto:
Public Property Get Posto() As String
Posto = pPosto
End Property
Public Property Let Posto(s As String)
pPosto = s
End Property

'Propriedade Analista:
Public Property Get Analista() As String
Analista = pAnalista
End Property
Public Property Let Analista(s As String)
pAnalista = s
End Property

'Propriedade Data Entrada:
Public Property Get DataEntrada() As Date
DataEntrada = pDataEntrada
End Property
Public Property Let DataEntrada(d As Date)
pDataEntrada = d
End Property

'Propriedade Mes Entrada:
Public Property Get MesEntrada() As String
MesEntrada = pMesEntrada
End Property
Public Property Let MesEntrada(d As String)
pMesEntrada = d
End Property


'Conforme o código para cada uma das propriedades existe um procedimento Get e outro Let.
' O Get é chamado quando se deseja ler o valor de uma propriedade. Ou seja ler o valor que já existe na classe
'O Let é chamado quando se deseja atribuir um valor à uma propriedade, em outras palavras quando você deseja inserir um valor na classe









