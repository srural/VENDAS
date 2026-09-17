Attribute VB_Name = "ChecaModulo"

Public Function CalculaDacCPF(pNumDoc As String) As Boolean

Dim DacInfo As String, dac1 As String, dac2 As String, StrAux As String
Dim Cpf As String

CalculaDacCPF = False

DacInfo = Right$(pNumDoc, 2)

Cpf = Mid(pNumDoc, 1, 3)
Cpf = Cpf & Mid(pNumDoc, 5, 3)
Cpf = Cpf & Mid(pNumDoc, 9, 3)
'Cpf = Cpf & Mid(pNumDoc, 13, 2)

StrAux = Cpf

'StrAux = Mid$(pNumDoc, 1, 9)

dac1 = DAC11B(StrAux)
StrAux = Trim$(StrAux) & Trim$(dac1)

dac2 = DAC11B(StrAux)
StrAux = Trim$(StrAux) & Trim$(dac2)

If StrComp(DacInfo, dac1 & dac2) = 0 Then
    CalculaDacCPF = True
End If

End Function

Public Function DAC11B(pcVarNum) As String

Dim nSoma As Integer, nInicio As Integer, nResto As Integer, nDac As Integer
Dim nContador As Integer, nContaFator As Integer

nInicio = Len(Trim(pcVarNum))
nSoma = 0
nContaFator = 2

For nContador = nInicio To 1 Step -1
    nSoma = nSoma + (Val(Mid(pcVarNum, nContador, 1)) * nContaFator)
    
    nContaFator = nContaFator + 1
    nContaFator = IIf(nContaFator > 99, 2, nContaFator)
Next nContador

nResto = nSoma Mod 11
nDac = IIf(nResto = 0 Or nResto = 1, 0, 11 - nResto)

DAC11B = Trim(str(nDac))

End Function

Public Function ContaNum(Number1 As String) As Boolean

Dim X As Integer
Dim soma As Integer
'Dim Number1 As Long
ReDim Conta(10) As Integer

'Number1 = Val(Number)

For X = 1 To 10
    Conta(X) = Mid(Number1, X, 1)
Next X

For X = 1 To 9
    If X Mod 2 <> 0 Then
    Conta(X) = Conta(X) * 2
    End If
Next X

For X = 1 To 9
    Select Case Conta(X)
        Case Is = 10
            soma = soma + 1
        
        Case Is = 11
            soma = soma + 2
        
        Case Is = 12
            soma = soma + 3
            
        Case Is = 13
            soma = soma + 4
        
        Case Is = 14
            soma = soma + 5
        
        Case Is = 15
            soma = soma + 6
        
        Case Is = 16
            soma = soma + 7
        
        Case Is = 17
            soma = soma + 8
            
        Case Is = 18
            soma = soma + 9
        
        Case Else
            soma = soma + Conta(X)
    End Select
Next X

soma = 10 - (soma Mod 10)

If soma <> Conta(10) Then
    ContaNum = False
Else
    ContaNum = True
End If

End Function

Public Function VerificaCGC(Numero As String) As Boolean

    Dim CGC As String

    VerificaCGC = False
    
    CGC = Mid(Numero, 1, 2)
    CGC = CGC & Mid(Numero, 4, 3)
    CGC = CGC & Mid(Numero, 8, 3)
    CGC = CGC & Mid(Numero, 12, 4)
    CGC = CGC & Mid(Numero, 17, 2)

    '----- Valida argumento -----'
    If Len(CGC) <> 14 Then
        VerificaCGC = False
        Exit Function
    End If

    '----- Valida primeiro dígito -----'
    If Modulo11(Left(CGC, 12)) <> Mid(CGC, 13, 1) Then
        VerificaCGC = False
        Exit Function
    End If
    
    '----- Valida o segundo dígito -----'
    If Modulo11(Left(CGC, 13)) <> Mid(CGC, 14, 1) Then
        VerificaCGC = False
        Exit Function
    End If

    VerificaCGC = True

End Function

Public Function Modulo11(Numero As String) As String

    Dim I As Integer
    Dim Produto As Integer
    Dim Multiplicador As Integer
    Dim Digito As Integer

    '----- Valida argumento -----'
    If Not IsNumeric(Numero) Then
        Modulo11 = ""
        Exit Function
    End If

    '----- Calcula o dígito no módulo 11 -----'
    Multiplicador = 2
    For I = Len(Numero) To 1 Step -1
        Produto = Produto + Val(Mid(Numero, I, 1)) * Multiplicador
        Multiplicador = IIf(Multiplicador = 9, 2, Multiplicador + 1)
    Next I

    '----- Exceção -----'
    Digito = 11 - Int(Produto Mod 11)
    Digito = IIf(Digito = 10 Or Digito = 11, 0, Digito)
    
    '----- Retorna -----'
    Modulo11 = Trim(str(Digito))

End Function

