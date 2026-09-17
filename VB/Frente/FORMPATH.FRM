VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Object = "{D6EEA3C0-6216-11CF-BE62-0080C72EDD2D}#1.0#0"; "MARQUEE.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9510
   LinkTopic       =   "Form1"
   ScaleHeight     =   6540
   ScaleWidth      =   9510
   StartUpPosition =   3  'Windows Default
   Begin MarqueeObjectsCtl.Marquee Marquee1 
      Height          =   735
      Left            =   3600
      TabIndex        =   10
      Top             =   3960
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   1296
   End
   Begin VB.TextBox Text3 
      Height          =   675
      Left            =   300
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   5640
      Width           =   2835
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Sistema"
      Height          =   675
      Left            =   1080
      TabIndex        =   8
      Top             =   4740
      Width           =   1215
   End
   Begin SysInfoLib.SysInfo SysInfo1 
      Left            =   180
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox Rtb 
      Height          =   1575
      Left            =   3480
      TabIndex        =   7
      Top             =   1620
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   2778
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"formpath.frx":0000
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Abrir"
      Height          =   615
      Left            =   5280
      TabIndex        =   6
      Top             =   240
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   960
      Width           =   5655
   End
   Begin MSComDlg.CommonDialog Comom1 
      Left            =   3540
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   3900
      Width           =   2835
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   555
      Left            =   600
      TabIndex        =   3
      Top             =   3120
      Width           =   2055
   End
   Begin VB.FileListBox File1 
      Height          =   1065
      Left            =   240
      TabIndex        =   2
      Top             =   1800
      Width           =   2775
   End
   Begin VB.DirListBox Dir1 
      Height          =   990
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   2775
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
   Begin VB.Line Line2 
      X1              =   60
      X2              =   3120
      Y1              =   4620
      Y2              =   4620
   End
   Begin VB.Line Line1 
      X1              =   3360
      X2              =   3360
      Y1              =   120
      Y2              =   6360
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim caminho As String

caminho = Dir1.Path & "\" & File1.FileName

Text1.Text = caminho
End Sub

Private Sub Command2_Click()
Dim caminho As String

Comom1.ShowOpen

caminho = Comom1.FileName
Text2 = caminho
Rtb.LoadFile (caminho)

End Sub

Private Sub Command3_Click()
Text3.Text = SysInfo1.OSVersion
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1
End Sub

Private Sub Form_Load()


End Sub
