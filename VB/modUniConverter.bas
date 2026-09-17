Attribute VB_Name = "modUniConverter"
Option Explicit

Public Type RECT
   Left                 As Long
   Top                  As Long
   Right                As Long
   Bottom               As Long
End Type

Public Type POINTAPI
   x                    As Long
   y                    As Long
End Type

Public Declare Function DrawTextW Lib "user32" (ByVal hdc As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByRef lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Public Declare Function TextOutW Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As Long, ByVal nCount As Long) As Long
Public Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByRef lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As String, ByVal lpUsedDefaultChar As Long) As Long

Private Const CP_UTF8 = 65001

Public Enum CodeEnum
   cgNone = 0
   cgANSI = 1
   cgUTF8 = 2
   cgUTF16 = 3
   cgHTML = 4
End Enum
#If False Then  'PreserveEnumCase
   Private cgNone, cgANSI, cgUTF8, cgUTF16, cgHTML
#End If

Public Function Any_Uni(ByVal sText As String) As String

   If IsUtf8(sText) Then
      If IsHTML(sText) Then
         Any_Uni = HTML_Uni(sText)
      Else
         Any_Uni = UTF8_Decode(sText)
      End If
   Else 'IsUtf16
      Any_Uni = sText
   End If

End Function

'Purpose: Input(ANSI, UTF-8, UTF-16), Output(UTF-16)
Public Function DecodeText(sText As String) As String
   Dim sUni             As String
   sUni = UTF8_Decode(sText)
   If sText = UTF8_Encode(sUni) Then
      DecodeText = sUni
   Else
      DecodeText = sText
   End If
End Function

Public Function fbIsNumeric(ByRef vRow As Variant) As Boolean
   Select Case VarType(vRow)
      Case 2 To 6, 14, 17  ' Integer, Long, Single, Double, Currency, Decimal, Byte
         fbIsNumeric = True
   End Select
End Function

'Purpose:Returns True if string is HTML
Public Function IsHTML(ByVal sText As String) As Boolean
   Dim Pos              As Long
   Dim lASC             As Long

   Pos = InStr(sText, "&#")

   Select Case Pos '   &#2
      Case 0
      Case Else
         Pos = Pos + 2
         If Pos <= Len(sText) Then
            lASC = Asc(Mid$(sText, Pos + 2, 1))
            If (lASC >= 48) And (lASC <= 57) Then
               IsHTML = True
               Exit Function
            End If
         End If
   End Select

End Function

'Purpose:Returns True if string is Utf16
Public Function IsUtf16(ByVal sText As String) As Boolean
   Dim i                As Long
   Dim lLen             As Long
   Dim lAscW            As Long

   lLen = Len(sText)
   For i = 1 To lLen
      lAscW = AscW(Mid$(sText, i))
      If lAscW < 0 Then
         lAscW = lAscW + 65536
      End If
      If (lAscW > 255) Then
         IsUtf16 = True
         Exit Function
      End If
   Next

End Function

'Purpose:Returns True if string is Utf8
Public Function IsUtf8(ByVal sText As String) As Boolean
   Dim sUni             As String
   sUni = UTF8_Decode(sText)
   IsUtf8 = sText = UTF8_Encode(sUni)
End Function

Public Function HTML_Uni(ByVal sText As String) As String
   Dim vVar             As Variant
   Dim i                As Long
   Dim sTemp            As String

   On Error Resume Next

   vVar = Split(sText, "&#")

   For i = 0 To UBound(vVar)
      sTemp = vVar(i)
      If Len(sTemp) > 1 And Right$(sTemp, 1) = ";" Then
         sTemp = Left$(sTemp, Len(sTemp) - 1)
      End If
      If IsNumeric(sTemp) Then
         sTemp = ChrW$(Val(sTemp))
      End If
      HTML_Uni = HTML_Uni & sTemp
   Next

End Function

Public Function Uni_HTML(ByVal sText As String) As String
   Dim i                As Long
   Dim lLen             As Long
   Dim lAscW            As Long
   Dim sTemp            As String

   lLen = Len(sText)
   For i = 1 To lLen
      sTemp = Mid$(sText, i, 1)
      lAscW = AscW(sTemp)
      If lAscW < 0 Then
         lAscW = lAscW + 65536
      End If
      If (lAscW > 255) Then
         Uni_HTML = Uni_HTML & "&#" & CStr(lAscW)
         If i <> lLen Then
            Uni_HTML = Uni_HTML & ";"
         End If
      Else
         Uni_HTML = Uni_HTML & sTemp
      End If
   Next
End Function

'Purpose: Convert Unicode to Vb code.
Public Function Uni_VB(ByVal sText As String) As String
   Dim lLen             As Long
   Dim i                As Long
   Dim sChar            As String
   Dim lChar            As Integer
   Dim CodeEnum         As CodeEnum

   lLen = Len(sText)

   If lLen Then
      For i = 1 To lLen
         sChar = Mid$(sText, i, 1)
         lChar = AscW(sChar)
         If (lChar >= &H0) And (lChar <= &HFF) Then
            Select Case CodeEnum
               Case cgNone
                  Uni_VB = Uni_VB & Chr$(34)
               Case cgANSI

               Case cgUTF16
                  Uni_VB = Uni_VB & " & " & Chr$(34)
            End Select
            CodeEnum = cgANSI
            Uni_VB = Uni_VB & sChar
         Else
            If CodeEnum = cgANSI Then
               Uni_VB = Uni_VB & Chr$(34)
            End If
            CodeEnum = cgUTF16
            Uni_VB = Uni_VB & " & ChrW$(&H" & Hex$(lChar) & ")"
         End If
      Next

      If CodeEnum = cgANSI Then
         Uni_VB = Uni_VB & Chr$(34)
      End If

   End If
End Function

'Purpose:Convert Utf8 to Unicode
Public Function UTF8_Decode(ByVal sUTF8 As String) As String

   Dim lngUtf8Size      As Long
   Dim strBuffer        As String
   Dim lngBufferSize    As Long
   Dim lngResult        As Long
   Dim bytUtf8()        As Byte
   Dim n                As Long

   If LenB(sUTF8) Then
      On Error GoTo EndFunction
      bytUtf8 = StrConv(sUTF8, vbFromUnicode)
      lngUtf8Size = UBound(bytUtf8) + 1
      On Error GoTo 0
      'Set buffer for longest possible string i.e. each byte is
      'ANSI<=&HFF, thus 1 unicode(2 bytes)for every utf-8 character.
      lngBufferSize = lngUtf8Size * 2
      strBuffer = String$(lngBufferSize, vbNullChar)
      'Translate using code page 65001(UTF-8)
      lngResult = MultiByteToWideChar(CP_UTF8, 0, bytUtf8(0), _
         lngUtf8Size, StrPtr(strBuffer), lngBufferSize)
      'Trim result to actual length
      If lngResult Then
         UTF8_Decode = Left$(strBuffer, lngResult)
         'Debug.Print UTF8_Decode
      End If
   End If

EndFunction:

End Function

'Purpose:Convert Unicode string to UTF-8.
Public Function UTF8_Encode(ByVal strUnicode As String) As String
   Dim i                As Long
   Dim TLen             As Long
   Dim lPtr             As Long
   Dim UTF16            As Long
   Dim UTF8_EncodeLong  As String

   TLen = Len(strUnicode)
   If TLen = 0 Then Exit Function

   Dim lngBufferSize    As Long
   Dim lngResult        As Long
   Dim bytUtf8()        As Byte
   'Set buffer for longest possible string.
   lngBufferSize = TLen * 3 + 1
   ReDim bytUtf8(lngBufferSize - 1)
   'Translate using code page 65001(UTF-8).
   lngResult = WideCharToMultiByte(CP_UTF8, 0, StrPtr(strUnicode), _
      TLen, bytUtf8(0), lngBufferSize, vbNullString, 0)
   'Trim result to actual length.
   If lngResult Then
      lngResult = lngResult - 1
      ReDim Preserve bytUtf8(lngResult)
      UTF8_Encode = StrConv(bytUtf8, vbUnicode)
   End If

End Function

