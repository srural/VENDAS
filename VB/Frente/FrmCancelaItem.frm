VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmCancelaItem 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edita Item"
   ClientHeight    =   1860
   ClientLeft      =   3930
   ClientTop       =   2925
   ClientWidth     =   4020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   4020
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cancela [ESC]"
      Height          =   390
      Left            =   585
      TabIndex        =   3
      Top             =   1050
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok [ENTER]"
      Default         =   -1  'True
      Height          =   390
      Left            =   2070
      TabIndex        =   1
      Top             =   1035
      Width           =   1215
   End
   Begin MSMask.MaskEdBox MebItem 
      DataField       =   "Custo"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """R$""#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   2
      EndProperty
      Height          =   435
      Left            =   240
      TabIndex        =   2
      Tag             =   "S"
      Top             =   360
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   767
      _Version        =   393216
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "0"
      PromptChar      =   "_"
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Digite o Número do item"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "FrmCancelaItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim quantidade As Currency
Dim desc As String
Dim Unit As String
Dim Valor As Currency
Dim Status As String
Dim Criterio As String
Dim CodigoBarra As String
Dim Tamanho As String
'Dim Criterio As String
Dim Resp As Byte
Dim Flag As Boolean
Dim soma As Currency
Resp = 7
Flag = True

If MebItem.Text <> "" Then
     VarItem = MebItem  ' ListVda.SelectedItem.Index

    Unload Me
      
    If FrmFrente.ListVda.ListItems.Count > 0 Then
        If Flag = True And Resp = 7 Then
            If FrmFrente.ListVda.ListItems.Count > 1 Then
                Criterio = "Item = " & VarItem
                FrmFrente.ListVda.FindItem Criterio
                If MsgBox("Confirma exclusão do " & FrmFrente.ListVda.ListItems(CInt(VarItem)).ListSubItems(4).Text, vbYesNo, App.Title) = vbYes Then
                    Criterio = "CodItp = " & FrmFrente.ListVda.SelectedItem.SubItems(7)
                    FrmFrente.RecVendas.FindFirst Criterio
                    If Not FrmFrente.RecVendas.NoMatch Then
                        soma = CCur(FrmFrente.soma) - Format((FrmFrente.ListVda.ListItems(CInt(VarItem)).ListSubItems(6).Text) / 100, "#,##0.000")
                        FrmFrente.MebTotal = Format(soma, "##,##0.00")
                        FrmFrente.ListVda.ListItems(CInt(VarItem)).ListSubItems(3).Text = "(Cancelado)"
                        FrmFrente.RecVendas.Edit
                        FrmFrente.RecVendas.Delete
'                        FrmFrente.ContList = FrmFrente.ContList - 1
'                        RecVendas.Update
                    End If
                End If
            Else
                MsgBox "Exclui o pedido inteiro", vbInformation, App.Title
            End If
        End If
            
    End If
End If
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Activate()
Label1.Caption = "Digite o Número do item"
MebItem.SetFocus
End Sub

Private Function CodTrib(Codigo As Integer) As String
If Codigo = 0 Or FrmFrente.Registros.Fields("Aliquota") = 0 Then
     CodTrib = "T03"
Else
    Select Case Codigo
        Case 1
            CodTrib = "T" + psCompString(1, Trim(str(FrmFrente.Registros.Fields("Aliquota") - 1)), 2)
        Case 2
            CodTrib = "F00"
        Case 3
            CodTrib = "I00"
        Case 4
            CodTrib = "N00"
    End Select
End If
Select Case Impressora
    Case 1
        CodTrib = CodTrib
    Case 2
        CodTrib = Left(CodTrib, 1)
End Select
End Function

Private Function VlrAliquota(Codigo As Integer) As String
If Codigo = 0 Or FrmFrente.Registros.Fields("Aliquota") = 0 Then
     VlrAliquota = "03"
Else
     VlrAliquota = psCompString(1, Trim(str(FrmFrente.Registros.Fields("Aliquota") - 1)), 2)
End If
End Function



