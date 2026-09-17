VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRecordset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Essa é uma classe para traduzir as propriedades
' de um recordset usadas no ReportX para o modelo
Private pRecordset As Recordset
Private pRSSec As Collection
Private pRSAtual As Integer

Public Sub AddSec(RS As Recordset)
    pRSSec.Add RS
End Sub

Public Property Get RSSec(Index As Variant)
    Set RSSec = pRSSec.Item(Index)
End Property

Public Property Set RSPri(Value As Recordset)
    Set pRecordset = Value
End Property
Public Property Get RSPri() As Recordset
    Set RSPri = pRecordset
End Property

Public Property Get RecordCount() As Long
    RecordCount = pRecordset.RecordCount
End Property

Public Property Get AbsolutePosition() As Long
    AbsolutePosition = pRecordset.AbsolutePosition
End Property

Public Property Get Eof() As Boolean
    Dim RS As Recordset
    ' O Eof nesse caso será TRUE apenas
    ' quando todos os EOFs dos RS secundarios
    ' forem TRUE e do principal tambem
    If pRecordset.Eof Then
        Eof = True
        For Each RS In pRSSec
            If Not RS.Eof Then
                Eof = False
                Exit For
            End If
        Next
    Else
        Eof = False
    End If
End Property

Public Property Get Bof() As Boolean
    Bof = pRecordset.Bof
End Property

Public Sub MoveNext()
    
    ' Só faz o move next no recordset
    ' principal, se todos os recordsets
    ' secundarios estirem em EOF
    
    pRSAtual = BuscaGrupo(1)
    If pRSAtual = -1 Then
        pRecordset.MoveNext
    Else
        pRSSec.Item(pRSAtual).MoveNext
        pRSAtual = BuscaGrupo(pRSAtual)
        If pRSAtual = -1 Then
            pRecordset.MoveNext
        Else
            pRSAtual = pRSAtual - 1
        End If
    End If
    
End Sub

Private Function BuscaGrupo(Indice As Integer) As Integer

    Dim RS As Recordset
        
    If Indice > pRSSec.Count Then
        BuscaGrupo = -1
    Else
        Set RS = pRSSec.Item(Indice)
        If RS.Eof Then
            BuscaGrupo = BuscaGrupo(Indice + 1)
        Else
            BuscaGrupo = Indice
        End If
    End If

End Function

Public Sub MovePrevious()
    pRecordset.MovePrevious
End Sub

Public Sub MoveFirst()
    Dim RS As Recordset
    For Each RS In pRSSec
        'If RS.State = adStateOpen Then
            If Not RS.Bof Then RS.MoveFirst
        'End If
    Next
    pRecordset.MoveFirst
End Sub

' Propriedade Defaul para atribuir os Fields.
Public Property Get Fields(Value As Variant) As Field
    Set Fields = pRecordset.Fields(Value)
End Property

Public Sub OpenSec(Grupo As Integer, SQL As String)

    Dim RS As Recordset
    
    ' Na colecao os indices sempre começam em 1
    ' Nos objetos das secoes começam em 0
    Set RS = pRSSec(Grupo + 1)
    SqlDiv(Grupo + 1) = SQL
    ' fecha os RS secundarios antes de abrir
    'If RS.State = adStateOpen Then
    '    RS.Close
    'End If
    Set RS = Banco.OpenRecordset(SQL)

    
End Sub

Public Sub CloseSec()
On Error Resume Next

    Dim RS As Recordset
    For Each RS In pRSSec
        'If RS.State = adStateOpen Then
            RS.Close
        'End If
    Next
    
End Sub

Public Sub PosicionaRSAtual()
    
    Dim RS As Recordset
    Dim i As Integer
    
    pRSAtual = -1
    
    For i = 1 To pRSSec.Count
        Set RS = pRSSec.Item(i)
        If Not RS.Eof Then
            pRSAtual = i - 1
            Exit For
        End If
    Next
    
End Sub

Public Property Get RSAtual() As Integer
    RSAtual = pRSAtual
End Property

Private Sub Class_Initialize()
    Set pRSSec = New Collection
End Sub

Private Sub Class_Terminate()
    CloseSec
    pRecordset.Close
    Set pRSSec = Nothing
End Sub

