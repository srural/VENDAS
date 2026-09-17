Attribute VB_Name = "ModuloProtecao"
Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long

Public Function GeraComplemento(NroCpf As String, Complemento As String) As String

Dim CpfLimpo As String
Dim Compl As Currency
Dim X As Currency

CpfLimpo = Replace(NroCpf, ",", "")
CpfLimpo = Replace(CpfLimpo, ".", "")
CpfLimpo = Replace(CpfLimpo, "-", "")
CpfLimpo = Replace(CpfLimpo, "_", "")

If Len(CpfLimpo) > 0 Then

    For X = 1 To Len(CpfLimpo)
        Compl = Compl + (Mid(CpfLimpo, X, 1) * X)
    Next
    
    GeraComplemento = Format(Complemento + Int((Compl / 10)), "0000")

Else

    GeraComplemento = 0
    
End If

End Function

Public Function Cripta(Numero As String, tipo As Integer) As String
Dim X As Currency
Dim MontaString As String
Dim NumAt As String

Numero = Replace(Numero, ".", "")
Numero = Replace(Numero, "-", "")

For X = 1 To Len(Numero)
    NumAt = Mid(Numero, X, 1)

    If X = 2 Or X = 3 Or X = 6 Or X = 7 Or X = 9 Or X = 11 Then
        'If NumAt Mod 2 = 0 Then
            NumAt = TabelaConv(NumAt, tipo)
        'End If
    Else
        If NumAt Mod 2 = 0 Then
            NumAt = NumAt + 1
        Else
            NumAt = NumAt - 1
        End If
    End If
    
    Mid(Numero, X, 1) = NumAt
    
Next

    'Cripta = Mid(Numero, 1, 4) & "." & Mid(Numero, 5, 4) & "." & Mid(Numero, 9, 4) & "." & Mid(Numero, 13, 3)
    Cripta = Numero

End Function

Public Function TabelaConv(Numero As String, tipo As Integer) As String
Dim X  As Currency

    If tipo = 1 Then
        Select Case Numero
            Case "0"
                TabelaConv = "W"
            Case "1"
                TabelaConv = "Z"
            Case "2"
                TabelaConv = "N"
            Case "3"
                TabelaConv = "H"
            Case "4"
                TabelaConv = "L"
            Case "5"
                TabelaConv = "E"
            Case "6"
                TabelaConv = "T"
            Case "7"
                TabelaConv = "K"
            Case "8"
                TabelaConv = "Q"
            Case "9"
                TabelaConv = "V"
        End Select
    ElseIf tipo = 2 Then
        Select Case Numero
            Case "W"
                TabelaConv = "0"
            Case "Z"
                TabelaConv = "1"
            Case "N"
                TabelaConv = "2"
            Case "H"
                TabelaConv = "3"
            Case "L"
                TabelaConv = "4"
            Case "E"
                TabelaConv = "5"
            Case "T"
                TabelaConv = "6"
            Case "K"
                TabelaConv = "7"
            Case "Q"
                TabelaConv = "8"
            Case "V"
                TabelaConv = "9"
        End Select
    End If
End Function

Public Function DriveSerial(ByVal Drive As String) As Long
  'Valor que retornará o serial do HD
  Dim RetVal As Long

  Dim HDNameBuffer As String * 256 'Nome do HD
  Dim FSBuffer As String * 256     'FS do HD
  Dim a As Long                    'auxiliar
  Dim b As Long                    'auxiliar

  Call GetVolumeInformation(Drive, HDNameBuffer, 256, RetVal, a, b, FSBuffer, 256)

  DriveSerial = RetVal
  
End Function

Public Function GeraDigito(Numero As Currency) As Currency

Dim X As Currency
Dim Compl As Currency
Dim NumTemp As Currency

    For X = 1 To Len(Numero)
        Compl = Compl + (Mid(Numero, X, 1) * X)
    Next
    
    NumTemp = Compl Mod 100
    GeraDigito = NumTemp

End Function

'Function LimpaString(Campo As String) As String
'Dim X As Integer
'Dim CampoLimpo As String

'For X = 1 To Len(Campo)
'    If Mid$(Campo, X, 1) >= Chr(32) And Mid$(Campo, X, 1) <= Chr(122) Then
'            CampoLimpo = CampoLimpo + Mid$(Campo, X, 1)
'    End If
'Next
'LimpaString = CampoLimpo
'End Function

'Public Function NameOfPC(MachineName As String) As Long
'Dim NameSize As Long
'Dim X As Long
'MachineName = Space$(16)
'NameSize = Len(MachineName)
'X = GetComputerName(MachineName, NameSize)
'End Function

Public Function EnCripta(strText As String) As String
Dim I As Long, c As Long
Dim strBuff As String
Dim strPwd As String '

strPwd = "Sr2090"

If Len(strPwd) Then
  For I = 1 To Len(strText)
  c = Asc(Mid$(strText, I, 1))
  c = c + Asc(Mid$(strPwd, (I Mod Len(strPwd)) + 1, 1))
  strBuff = strBuff & Chr$(c And &HFF)
  Next I
Else
  strBuff = strText
End If

EnCripta = strBuff

End Function

Public Function DeCripta(strText As String) As String
Dim I As Long, c As Long
Dim strBuff As String
Dim strPwd As String

strPwd = "Sr2090"

If Len(strPwd) Then
For I = 1 To Len(strText)
c = Asc(Mid$(strText, I, 1))
c = c - Asc(Mid$(strPwd, (I Mod Len(strPwd)) + 1, 1))
strBuff = strBuff & Chr$(c And &HFF)
Next I
Else
strBuff = strText
End If
DeCripta = strBuff
End Function

