Attribute VB_Name = "Retirar_Acentuação"
Option Explicit


Public Function RetirarAcento(ByVal Texto As String) As String

Dim i As Integer

For i = 0 To 255

    If i = 39 Or i = 94 Or i = 96 Or i = 126 Or i = 145 _
        Or i = 146 Or i = 168 Or i = 180 Then
        Texto = Replace(Texto, Chr(i), "")
    End If
        
    If i >= 192 And i <= 197 Then
        Texto = Replace(Texto, Chr(i), "A")
    End If
    
    If i = 199 Then
        Texto = Replace(Texto, Chr(i), "C")
    End If
    
    If i >= 200 And i <= 203 Then
        Texto = Replace(Texto, Chr(i), "E")
    End If
    
    If i >= 204 And i <= 207 Then
        Texto = Replace(Texto, Chr(i), "I")
    End If
    
    If i = 209 Then
        Texto = Replace(Texto, Chr(i), "N")
    End If
    
    If i >= 210 And i <= 214 Then
        Texto = Replace(Texto, Chr(i), "O")
    End If
    
    If i >= 217 And i <= 220 Then
        Texto = Replace(Texto, Chr(i), "U")
    End If
    
    If i = 221 Then
        Texto = Replace(Texto, Chr(i), "Y")
    End If
    
    If i >= 224 And i <= 230 Then
        Texto = Replace(Texto, Chr(i), "a")
    End If
    
    If i = 231 Then
        Texto = Replace(Texto, Chr(i), "c")
    End If
    
    If i >= 232 And i <= 235 Then
        Texto = Replace(Texto, Chr(i), "e")
    End If
    
    If i >= 236 And i <= 239 Then
        Texto = Replace(Texto, Chr(i), "i")
    End If
    
    If i = 240 Then
        Texto = Replace(Texto, Chr(i), "o")
    End If
    
    If i = 241 Then
        Texto = Replace(Texto, Chr(i), "n")
    End If
    
    If i >= 242 And i <= 246 Then
        Texto = Replace(Texto, Chr(i), "o")
    End If
    
    If i >= 249 And i <= 252 Then
        Texto = Replace(Texto, Chr(i), "u")
    End If
    
    If i = 253 Or i = 255 Then
        Texto = Replace(Texto, Chr(i), "y")
    End If
Next i

RetirarAcento = Texto

End Function
