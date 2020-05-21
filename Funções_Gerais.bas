Attribute VB_Name = "modFunções_Gerais"

Option Explicit

'principais cores ****************************************************************
Public Enum Cor
    Branco = &HFFFFFF

    Amarelo = &HFFFF&
    AmareloClaro = &H84FFFF
    AmareloEscuro = &HC7C6&
    
    Pele = &HC6E7FF
    Laranja = &H80FF&
    Castanho = &H41C6&
    Marrom = &H84&
    
    Vermelho = &HFF&
    VermelhoClaro = &H8486FF
    VermelhoEscuro = &HC6&
    
    Rosa = &HFF86FF
    RosaClaro = &H8486FF
    RosaPink = &HFF00FF
    
    Roxo = &HC600C6
    RoxoEscuro = &H840084
    
    Azul = &HFF0000
    AzulCiano = &HC6C700
    AzulClaro = &HFFFF00
    AzulClaríssimo = &HFFFFC6
    AzulEscuro = &HC60000
    
    Violeta = &HFF8684
    VioletaClaro = &HFFC7C6
    
    Verde = &HC700&
    VerdeClaro = &HFF00&
    VerdeSuave = &H84FF84
    VerdeClaríssimo = &HC6FFC6
    VerdeEscuro = &H8600&
    
    Cinza = &H848684
    CinzaClaríssimo = &HE7E7E7
    CinzaClaro = &HC6C7C6
    CinzaEscuro = &H424142
    
    Preto = &H0&
End Enum
'*********************************************************************************

'Gerenciar arquivos INI **********************************************************
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias _
    "WritePrivateProfileStringA" (ByVal Seção As String, ByVal Item As String, _
    ByVal Valor As String, ByVal NomeDoArquivo As String) As Long

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias _
    "GetPrivateProfileStringA" (ByVal Seção As String, ByVal Item As String, _
    ByVal ValorPadrão As String, ByVal ValorRetornado As String, _
    ByVal Tamanho As Long, ByVal NomeDoArquivo As String) As Long
'*********************************************************************************

'Captura a posição do mouse em qualquer parte da tela ****************************
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Type POINTAPI
    X As Long
    Y As Long
End Type
Public Pos As POINTAPI
'*********************************************************************************

'Mantem o formulário acima dos outros ********************************************
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Const hWnd_TOP = 0
Public Const hWnd_TOPMOST = -1
Public Const hWnd_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
'*********************************************************************************

'Inicializar junto com o sistema *************************************************
Public Enum Usuário
    Usuário_Atual = 0
    Todos_Usuários = 1
    Ambos = 2
End Enum

Public Enum Incialização
    Inicializar = 0
    Retirar_Inicialização = 1
End Enum
'*********************************************************************************

'Verfica rapidamente se um caminho existe ****************************************
Public Declare Function PathFileExistsA Lib "shlwapi.dll" _
    (ByVal pszPath As String) As Long
'*********************************************************************************

Public Function Ap(ByVal Texto As String) As String
'Coloca apóstrofos em um texto para ser utilizado em consultas SQLs

If InStr(Texto, Chr(39)) Then
    Ap = Chr(39) & Replace(Texto, "'", "''") & Chr(39)
Else
    Ap = Chr(39) & Texto & Chr(39)
End If

End Function

Public Sub GravarINI(Item As String, Valor As String, Optional Seção As String, Optional NomeDoArquivo As String)
'grava um arquivo ini

If Seção = "" Then Seção = App.EXEName
If NomeDoArquivo = "" Then NomeDoArquivo = App.Path & "\Config.ini"

WritePrivateProfileString Seção, Item, Valor, NomeDoArquivo

End Sub
Public Function LerINI(Item As String, Optional ValorPadrão As String, Optional Seção As String, Optional NomeDoArquivo As String) As String
'lê um arquivo ini

If Seção = "" Then Seção = App.EXEName
If NomeDoArquivo = "" Then NomeDoArquivo = App.Path & "\Config.ini"

Dim i As Long
Dim Valor As String * 1024
i = GetPrivateProfileString(Seção, Item, ValorPadrão, Valor, Len(Valor), NomeDoArquivo)

If i > 0 Then
    LerINI = Left(Valor, i)
Else:
    LerINI = ""
End If

End Function

Public Function CaminhoExiste(Caminho As String) As Boolean
'Verifica a existencia de um arquivo ou pasta

If PathFileExistsA(Caminho) = 1 Then
    CaminhoExiste = True
Else
    CaminhoExiste = False
End If

End Function

Public Function NomeDeArquivo(Nome As String) As String
'altera um texto, retirando caracteres especiais,
'para ser usado como nome de um arquivo

Nome = Replace(Nome, "\", "_")
Nome = Replace(Nome, "|", "_")
Nome = Replace(Nome, "/", "_")
Nome = Replace(Nome, "?", "_")
Nome = Replace(Nome, "*", "_")
Nome = Replace(Nome, ":", "_")
Nome = Replace(Nome, ">", "_")
Nome = Replace(Nome, "<", "_")
Nome = Replace(Nome, Chr(34), "_")

Dim Nome191 As String * 191
Nome191 = Nome

NomeDeArquivo = Nome191

End Function

Public Function NomeDePasta(Nome As String) As String
'altera um texto, retirando caracteres especiais,
'para ser usado como nome de uma pasta

Nome = Replace(Nome, "\", "_")
Nome = Replace(Nome, "|", "_")
Nome = Replace(Nome, "/", "_")
Nome = Replace(Nome, "?", "_")
Nome = Replace(Nome, "*", "_")
Nome = Replace(Nome, ":", "_")
Nome = Replace(Nome, ">", "_")
Nome = Replace(Nome, "<", "_")
Nome = Replace(Nome, Chr(34), "_")

Dim Nome180 As String * 180
Nome180 = Nome

NomeDePasta = Nome180

End Function

Public Sub PosiçãoDoMouse()
'Captura a posição atual do mouse e coloca na variável Pos
GetCursorPos Pos

End Sub

Public Sub Form_Acima(FormHwnd As Long, Acima As Boolean)

'liga ou desliga a opção de manter um formulário acima dos outros
If Acima = True Then
    SetWindowPos FormHwnd, hWnd_TOPMOST, 0, 0, 0, 0, FLAGS
Else:
    SetWindowPos FormHwnd, hWnd_NOTOPMOST, 0, 0, 0, 0, FLAGS
End If

End Sub

Public Sub IniComSistema(Inicialização As Incialização, Usuário As Usuário)

On Error Resume Next
Dim RKRunUser As String
RKRunUser = "HKEY_CURRENT_USER\Software\Microsoft\Windows\" _
    & "CurrentVersion\Run\" & App.EXEName
    
Dim RKRunMachine As String
RKRunMachine = "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\" _
    & "CurrentVersion\Run\" & App.EXEName

Dim DirProgram As String
DirProgram = App.Path & "\" & App.EXEName

Dim RegEdit As Object
Set RegEdit = CreateObject("wscript.shell")

If Inicialização = Inicializar Then
    If Usuário = Usuário_Atual Or Usuário = Ambos Then
        RegEdit.regwrite RKRunUser, DirProgram, "REG_SZ"
    End If
    
    If Usuário = Todos_Usuários Or Usuário = Ambos Then
        RegEdit.regwrite RKRunMachine, DirProgram, "REG_SZ"
    End If
End If

If Inicialização = Retirar_Inicialização Then
    If Usuário = Usuário_Atual Or Usuário = Ambos Then
        RegEdit.regdelete RKRunUser
    End If
    
    If Usuário = Todos_Usuários Or Usuário = Ambos Then
        RegEdit.regdelete RKRunMachine
    End If
End If
    
End Sub

