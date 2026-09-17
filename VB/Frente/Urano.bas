Attribute VB_Name = "Urano"
Option Explicit
' ************** Declarações da impressora Urano 1 estação **************

' ************** Declarações da impressora Urano 1 estação  TERMICA  **************



Declare Function DLLG2_Versao Lib "DLLG2.dll" (ByVal Versao As String, ByVal TamVersao As Long) As String
Declare Function DLLG2_IniciaDriver Lib "DLLG2.dll" (ByVal Canal As String) As Long
Declare Function DLLG2_ConfiguraDriver Lib "DLLG2.dll" (ByVal Handle As Long, ByVal Speed As Long) As Long
Declare Function DLLG2_ObtemCodErro Lib "DLLG2.dll" (ByVal Handle As Long) As Long
Declare Function DLLG2_ObtemNomeErro Lib "DLLG2.dll" (ByVal Handle As Long, ByVal NomeErro As String, ByVal TamNomeErro As Long) As String
Declare Function DLLG2_ObtemCircunstancia Lib "DLLG2.dll" (ByVal Handle As Long, ByVal Circunstancia As String, ByVal TamNomeCircunstancia As Long) As String
Declare Function DLLG2_ObtemRetornos Lib "DLLG2.dll" (ByVal Handle As Long, ByVal Retornos As String, ByVal TamRetorno As Long) As String
Declare Sub DLLG2_DefineTimeout Lib "DLLG2.dll" (ByVal Handle As Long, ByVal TempoMaximo As Long)
Declare Function DLLG2_LeTimeout Lib "DLLG2.dll" (ByVal Handle As Long) As Long
Declare Function DLLG2_ExecutaComando Lib "DLLG2.dll" (ByVal Handle As Long, ByVal Comando As String) As Long
Declare Function DLLG2_EncerraDriver Lib "DLLG2.dll" (ByVal Handle As Long) As Long
Declare Sub DLLG2_SetaArquivoLog Lib "DLLG2.dll" (ByVal NomeArquivo As String)
Declare Function DLLG2_ObtemNomeLog Lib "DLLG2.dll" (ByVal NomeArquivo As String, ByVal TamNomeArquivo As Long) As String
Declare Sub DLLG2_LimpaParams Lib "DLLG2.dll" (ByVal Handle As Long)
Declare Function DLLG2_ListaParams Lib "DLLG2.dll" (ByVal Handle As Long, ByVal LstParams As String, ByVal TamLstParams As Long) As String
Declare Sub DLLG2_AdicionaParam Lib "DLLG2.dll" (ByVal Handle As Long, ByVal NomeParam As String, ByVal ValorParam As String, ByVal TipoParam As Long)
Declare Function DLLG2_TotalRetornos Lib "DLLG2.dll" (ByVal Handle As Long) As Long
Declare Function DLLG2_Retorno Lib "DLLG2.dll" (ByVal Handle As Long, ByVal Indice As Long, ByVal NomeRetorno As String, ByVal TamNomeRetorno As Long, ByVal ValorRetorno As String, ByVal TamValorRetorno As Long) As Long
Declare Function DLLG2_LeRegistrador Lib "DLLG2.dll" (ByVal Handle As Long, ByVal NomeRegistrador As String, ByVal NomeComando As String, ByVal TamNomeComando As Long) As Long

Private Declare Function DLLG2_ViaTransporteInicia Lib "DLLG2.dll" (ByVal Handle As Long, ByVal COO As String, ByVal ValorTotalCupom As String, ByVal NomeComprador As String, ByVal EnderecoComprador As String, ByVal TextoLivre1 As String, ByVal TextoLivre2 As String) As Long
Private Declare Function DLLG2_ViaTransporteLinhas Lib "DLLG2.dll" (ByVal Handle As Long, ByVal Linhas As String) As Long
Private Declare Function DLLG2_ViaTransporteEncerra Lib "DLLG2.dll" (ByVal Handle As Long, ByVal Operador As String) As Long


