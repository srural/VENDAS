VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form FormWeb 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta - ACIUR/Urupês"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   7245
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   1500
      Left            =   0
      TabIndex        =   0
      Top             =   2610
      Width           =   7215
      ExtentX         =   12726
      ExtentY         =   2646
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Image Image1 
      Height          =   2655
      Left            =   0
      Picture         =   "FormWeb.frx":0000
      Stretch         =   -1  'True
      Top             =   -45
      Width           =   7215
   End
End
Attribute VB_Name = "FormWeb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
WebBrowser1.Navigate "http://www.aciur.org.br/part/65/ConsultaSCPC.html"
'WebBrowser1.Navigate "http://www.sidcomp.com.br/suporte"
End Sub

