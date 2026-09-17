VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.4#0"; "ReportX.ocx"
Begin VB.Form frmRelatorio 
   ClientHeight    =   4275
   ClientLeft      =   1875
   ClientTop       =   2280
   ClientWidth     =   8850
   LinkTopic       =   "Form1"
   ScaleHeight     =   7.541
   ScaleMode       =   7  'Centimeter
   ScaleWidth      =   15.61
   Begin ReportX.ReportSection Secao 
      Align           =   1  'Align Top
      Height          =   465
      Index           =   7
      Left            =   0
      Top             =   3030
      Width           =   8850
      _ExtentX        =   15610
      _ExtentY        =   820
      Tipo            =   8
   End
   Begin ReportX.ReportSection Secao 
      Align           =   1  'Align Top
      Height          =   465
      Index           =   4
      Left            =   0
      Top             =   1635
      Width           =   8850
      _ExtentX        =   15610
      _ExtentY        =   820
      Tipo            =   5
      Ordem           =   1
   End
   Begin ReportX.ReportSection Secao 
      Align           =   1  'Align Top
      Height          =   465
      Index           =   2
      Left            =   0
      Top             =   930
      Width           =   8850
      _ExtentX        =   15610
      _ExtentY        =   820
      Tipo            =   3
      Ordem           =   1
   End
   Begin ReportX.ReportSection Secao 
      Align           =   1  'Align Top
      Height          =   465
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   8850
      _ExtentX        =   15610
      _ExtentY        =   820
      Tipo            =   1
   End
   Begin ReportX.ReportSection Secao 
      Align           =   1  'Align Top
      Height          =   465
      Index           =   1
      Left            =   0
      Top             =   465
      Width           =   8850
      _ExtentX        =   15610
      _ExtentY        =   820
      Tipo            =   2
   End
   Begin ReportX.ReportSection Secao 
      Align           =   1  'Align Top
      Height          =   240
      Index           =   3
      Left            =   0
      Top             =   1395
      Width           =   8850
      _ExtentX        =   15610
      _ExtentY        =   423
   End
   Begin ReportX.ReportSection Secao 
      Align           =   1  'Align Top
      Height          =   465
      Index           =   6
      Left            =   0
      Top             =   2565
      Width           =   8850
      _ExtentX        =   15610
      _ExtentY        =   820
      Tipo            =   7
   End
   Begin ReportX.ReportSection Secao 
      Align           =   1  'Align Top
      Height          =   465
      Index           =   5
      Left            =   0
      Top             =   2100
      Width           =   8850
      _ExtentX        =   15610
      _ExtentY        =   820
      Tipo            =   6
   End
   Begin ReportX.ReportMain Relatorio 
      Height          =   480
      Left            =   30
      TabIndex        =   0
      Top             =   3690
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Pagina          =   9
      Divisao         =   1
      Regua           =   -1  'True
      Escala          =   7
      Titulo          =   ""
      HTMLUnico       =   -1  'True
   End
End
Attribute VB_Name = "frmRelatorio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Os eventos abaixo estão comentados, mas não são nessários.
' Após a compreensão do processo de impressão esses eventos
' podem ser alterados no Template do formulário.

' Não coloque nenhum tipo de código nesses evento que possam fazer
' o sistema ficar em "loop contínuo", pois é necessário que o controle volte
' do evento para a conclusão do processo de impressão. Qualquer tipo de
' erro que possa ocorrer deve ser evitado com um código de manipulação
' de erro através dos comandos "On Error"

' Obs: Se você estiver utilizando um objeto recordset não esqueça de
' fechá-lo no evento Form_Unload. O ActiveX não tem um objeto recordset, apenas
' uma referência ao objeto que você passa para ele, portanto qualquer referência
' a propriedade recordset do ReportMain estará sendo feita ao objeto ajustado.

' Esse formulário de exemplo contém os controles de seção (ReportSection) e
' o controle principal de impressão (ReportMain). Os controles de seção são
' "Containers" e os elementos de impressão de cada seção deve ser inseridos
' "dentro" deles. Para a impressão podem ser utlizados os seguintes objetos:
' Label, ReportField, PictureBox, Line e Shape (Exceto os tipos Rounded)

' Todas as propriedades são baseadas na escala de "Twips", pois o
' ActiveX não permite a mudança de escala do container.
' Uma referência visual de outras escalas pode ser obtida ajustando
' a propriedade "Escala" do ReportMain.

' Pode-se ajustar as posições em run-time utilizando a propriedade MargemEsquerda
' e a propriedade Left do Controle que estiver dentro da seção, mas para isso a
' propriedade Escala do ReportMain deve estar ajustada para Twips.

Public Sub Config()
    
    ' É recomendado criar um procedimento público nesse formulário
    ' que será utilizado para a impressão do relatório.
    
    ' Ajuste o Recordset
    
    ' Dim Rds as DAO.Recordset
    ' Set Rds = db.OpenRecordset("Consulta", dbOpenSnapshot)
    ' Set Relatorio.Recordset = Rds
    
    ' ou
    
    ' Set Relatorio.Recordset = Data1.Recordset
    
    ' ou
    
    ' Dim Rds as ADODB.Recordset
    ' Set Rds = New ADODB.Recordset
    ' Rds.Open "Consulta", Conexao, adOpenStatic
    
    ' Informe o título que deve aparecer na barra de título do relatório
    Relatorio.Titulo = "Título do Relatório"
    
    ' Ative o relatório
    Relatorio.Ativar
    
    ' A partir daqui, todo o processo é feito pelo ActiveX
    ' O controle será devolvido ao programa logo após o término
    ' da impressão.
    
    ' É conveniente colocar aqui o fechamento dos objetos
    
    ' Rds.Close
    ' Set Rds = Nothing
    
    Unload Me
    
End Sub

Private Sub Relatorio_Continua(Cancelar As Boolean)

    ' No início de cada página esse evento é disparado e
    ' o relatório pode ser cancelado colocando Cancelar = True

    ' OBS: Depois de algumas atualizações do componente o comportamento
    ' desse evento ficou confuso portanto não utilize-o.
    ' Ele foi mantido apenas para efeito de compatibilidade das versões.
    
End Sub

Private Sub Relatorio_Erro(ByVal Numero As Long)

    ' Indica a ocorrência de um erro durante a impressão:
    
    ' If Numero < 0 then
        ' Os código de erros previstos estão no arquivo de Ajuda.
    '   MsgBox "Erro Previsto:" & Numero - vbObjectError
    ' Else
    '   MsgBox "Erro não Previsto:" & Numero
    ' End If
    
End Sub

Private Sub Relatorio_FormatarControle(Controle As Object, ByVal Valor As String)

    ' Esse evento pode ser utilizado para formatar um controle de uma seção
    ' de acordo com o registro corrente.
    ' Por exemplo, para um Label que contenha um numero inteiro
    ' e esse numero possa ser negativo ou positivo conforme o registro:
    ' If Controle.Name = "Label1" Then
    '   If Val(Valor) < 0 Then
    '     Controle.ForeColor = RGB(255,0,0)
    '   Else
    '     Controle.ForeColor = RGB(0,0,255)
    '   End If
    ' End If
    
End Sub

Private Sub Relatorio_FormulaCampo(ByVal Campo As String, Valor As Variant)

    ' Esse evento é disparado para cada campo (ReportField) que tiver a propriedade
    ' Formula em True. O valor da propriedade Campo será passado no parâmetro "Campo"
    ' e o parâmetro Valor deve ser ajustado com o resultado.
    
    ' Normalmente esse resultado é baseado nos campos do Recordset.
    ' Para retornar o valor da página corrente, por exemplo, coloque um
    ' ReportField na seção de Cabeçalho, ajuste a propriedade Campo para
    ' "NumeroPagina" e Formula para True, em seguida coloque o seguinte código:
    
    'If Campo = "NumeroPagina" Then
    '   Valor = Relatorio.PaginaAtual
    'End If
    
    ' Nesse momento o recordset está apontando para o registro que está sendo impresso
    ' na seção de "Detalhe".
    
End Sub

Private Sub Relatorio_FormulaGrupo(ByVal Ordem As Byte, Valor As Variant)

    ' Esse evento é disparado pelo controle para saber a fórmula que comanda a quebra
    ' do grupo. O parâmetro "Ordem" contém o número do grupo e o parâmetro valor
    ' deve ser ajustado para o valor que controla a quebra do grupo.
    
    ' Por exemplo, se o grupo 1 é quebrado por cliente através do campo "Cliente" do recordset,
    ' que contém o código do cliente, então o valor deve ser ajustado da seguinte maneira:
    ' If Ordem = 1 Then
    '     Valor = Relatorio.Recordset("Cliente")
    ' End If
    
    ' Obs: Se existir apenas 1 grupo não é necessário fazer a verificação do parâmetro Ordem
    ' Para que as quebras de grupo fiquem coerentes, é necessário que o recordset utilizado
    ' esteja ordendo pelo campo que controla o grupo, ou seja, no exemplo anterior o recordset
    ' deveria ser do tipo: "SELECT * FROM Clientes ORDER BY Cliente"
        
    ' Obs: No caso de estar utilizando um recordset tipo "Table", o índice ajustado deve corresponder
    ' ao campo utilizado para a quebra do grupo.
    
End Sub

Private Sub Relatorio_ImprimiuRegistro(Cancelar As Boolean)

    ' Esse evento é disparado após a impressão de cada registro.
    ' Pode verificar aqui qualquer tipo de condição e no sistema e
    ' cancelar o relatório, ajustando o parâmetro Cancelar para True
    
End Sub

Private Sub Relatorio_IniciarGrupo(ByVal Ordem As Byte)

    ' Esse evento deve ser utilizado somente nos relatório que contém
    ' as seções do tipo "Cabeçalho de Grupo" e "Rodapé de Grupo"
    
    ' Esse evento é disparado sempre que um Grupo é inciado
    ' O parâmetro Ordem contém o número do grupo. Os grupos
    ' estão organizados de "fora para dentro" do detalhe, em cascata,
    ' assim o grupo 1 é mais externo que o grupo 2, por exemplo.
    
    ' Para que o relatório com grupos funcione corretamente, é necessário
    ' que existam ambas a seções do grupo, "Cabeçalho de Grupo" e
    ' "Rodapé de Grupo" para que o relatório possa ser impresso corretamente.
    ' Se não quizer que um dos dois, ou os dois apareçam ajuste a propriedade
    ' "Mostrar" da seção para False.
    
    ' Esse evento é utlizado para o caso de haver algum acumulador de totais
    ' para o grupo. Esse acumulador deve ser zerado nesse evento para que
    ' mostre o resultado correto na impressão do rodapé do grupo.
    ' Obs: Os acumuladores não podem ser variáveis locais do evento
    ' Podem ser privativos do formulário ou globais do projeto.
    
End Sub

Private Sub Relatorio_IniciarRelatorio(ByVal Impressora As Boolean, Cancelar As Boolean)
    
    ' O evento é ativado no início do relatório ou quando o usuário troca as configurações
    ' de impressora ou de página.
    
    ' Se o relatorio estiver sendo direcionado para a Visualização, o parâmetro
    ' Impressora é passado como False. Se ele estiver sendo enviado para a impressora
    ' ou sendo gravado em arquivo o parâmetro é passado como True.
    
    ' Quando o parâmetro Cancelar é ajustado para True
    ' e o parâmetro Impressora é igual a False, a impressão será automaticamente
    ' direcionada para a impressora e o evento será novamente disparado.
    ' Para cancelar efetivamente o relatório o parâmetro Cancelar deve ser
    ' ajustado para True quando o parâmetro Impressora for True.
    
    ' Pode-se colocar aqui, por exemplo, para impedir que a impressão
    ' seja direcionada para a impressora:
    ' If Impressora Then
    '     MsgBox "Essa versão não permite a impressão em papel!"
    '     Cancelar = True
    ' End If
    
    ' Ou, para evitar a visualização, passando direto para a impressora:
    ' If Not Impressora Then
    '     Cancelar = True
    ' End If
    ' Nas versões mais recente existem as propriedades Visualizar e Imprimir
    ' que são mais eficientes que utilizar esse evento.
            
End Sub

Private Sub Relatorio_IniciarSecao(ByVal Secao As ReportX.TSecao, ByVal Ordem As Byte)

    ' Esse evento é disparado antes da impressão das seções
    ' Para identificar a seção, use o parâmetro Secao.
    ' Se for uma seção do tipo Cabecalho de Grupo ou Rodapé de Grupo
    ' use o parâmetro Ordem para identificá-la.
        
End Sub

Private Sub Relatorio_ValidarRegistro(Invalido As Boolean, Cancelar As Boolean)
    
    ' Esse evento é executado antes de cada impressão de cada registro
    ' Se o flag "Invalido" for colocado para True, o registro corrente é ignorado.
    ' Se "Cancelar" é colocado em True o relatório é cancelado
    
    ' Utilizar os filtros do próprio Recordset é mais rápido e eficiente.
    ' Pode-se utilizar esse recurso quando o relatório é baseado em um
    ' Recordset do tipo "Table" e os filtro não estão disponíveis.
        
End Sub

