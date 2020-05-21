VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form f1 
   BackColor       =   &H00404040&
   Caption         =   "Consulta Hemeroteca"
   ClientHeight    =   6450
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9345
   Icon            =   "f1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6450
   ScaleWidth      =   9345
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   810
      TabIndex        =   12
      Top             =   2340
      Visible         =   0   'False
      Width           =   7575
   End
   Begin VB.CommandButton botNormalizar 
      Caption         =   "Normalizar tabela"
      Height          =   305
      Left            =   1830
      TabIndex        =   11
      Top             =   750
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Frame frmInserir 
      BackColor       =   &H00808080&
      Caption         =   "Inserir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   150
      TabIndex        =   6
      Top             =   90
      Width           =   8385
      Begin VB.CommandButton botInserir 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Artigo de revista"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   3
         Left            =   6300
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   210
         Width           =   2000
      End
      Begin VB.CommandButton botInserir 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Artigo de Guarulhos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   2
         Left            =   4230
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   210
         Width           =   2000
      End
      Begin VB.CommandButton botInserir 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Biografia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   210
         Width           =   2000
      End
      Begin VB.CommandButton botInserir 
         BackColor       =   &H008080FF&
         Caption         =   "Assunto do Fichário"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   210
         Width           =   2000
      End
   End
   Begin VB.CommandButton botOk 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   370
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1050
      Width           =   645
   End
   Begin VB.CommandButton botPrioridade 
      BackColor       =   &H008080FF&
      Caption         =   "^^ Fichário ^^"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   370
      Left            =   4530
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1050
      Width           =   1725
   End
   Begin VB.TextBox txtLocalizar 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   150
      MaxLength       =   100
      TabIndex        =   1
      Top             =   1050
      Width           =   3675
   End
   Begin MSFlexGridLib.MSFlexGrid lstBusca 
      Height          =   4665
      Left            =   150
      TabIndex        =   0
      Top             =   1440
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   8229
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColorBkg    =   12632256
      WordWrap        =   -1  'True
      ScrollBars      =   2
      SelectionMode   =   1
      FormatString    =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label legP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Prioridade"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4530
      TabIndex        =   4
      Top             =   840
      Width           =   870
   End
   Begin VB.Label legLocalizar 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Localizar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   150
      TabIndex        =   2
      Top             =   840
      Width           =   780
   End
End
Attribute VB_Name = "f1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Option Compare Text

Dim ÞIndLista As Long

'variáveis usadas para ajustar à resolução do monitor ****************************
Dim MyForm As FRMSIZE
Dim DesignX As Integer
Dim DesignY As Integer

Private Sub botInserir_Click(Index As Integer)

If Index = 0 Then
    With f2
        .Tag = 0
        .Left = Me.Left + (Me.Width - f2.Width) / 2
        .Top = Me.Top + (Me.Height - f2.Height) / 2
        .legT.Caption = "Cadastro de novo assunto do fichário"
        .botExcluir.Visible = False
        .Show 1
    End With
End If

If Index = 1 Then
    With f3
        .Tag = 0
        .Left = Me.Left + (Me.Width - f2.Width) / 2
        .Top = Me.Top + (Me.Height - f2.Height) / 2
        .legT.Caption = "Cadastro de nova biografia"
        .botExcluir.Visible = False
        .Show 1
    End With
End If

If Index = 2 Then
    With f4
        .Tag = 0
        .Left = Me.Left + (Me.Width - f2.Width) / 2
        .Top = Me.Top + (Me.Height - f2.Height) / 2
        .legT.Caption = "Cadastro de novo assunto de Guarulhos"
        .botExcluir.Visible = False
        .Show 1
    End With
End If

If Index = 3 Then
    With f5
        .Tag = 0
        .Left = Me.Left + (Me.Width - f2.Width) / 2
        .Top = Me.Top + (Me.Height - f2.Height) / 2
        .legT.Caption = "Cadastro de novo assunto de revista"
        .botExcluir.Visible = False
        .Show 1
    End With
End If



End Sub

Private Sub botNormalizar_Click()

On Error GoTo Erro

DoEvents
Dim Tabela As String
Tabela = Trim(txtLocalizar)
 
If øCon.State = 1 Then
    Dim Busca As New ADODB.Recordset
    Busca.ActiveConnection = øCon
    Busca.CursorType = adOpenStatic
    
    Busca.Open "Select * From " & Tabela & " Where Assunto <> NULL Order By Cod"
    
    
    If Busca.RecordCount > 0 Then
        Busca.MoveLast: Busca.MoveFirst
        
        Dim i As Long, j As Long
        j = Busca.RecordCount
        For i = i To Busca.RecordCount - 1
            Dim AssuntoSA As String
            AssuntoSA = RetirarAcento(Busca("Assunto"))
            øCon.Execute "Update " & Tabela _
                & " Set AssuntoSA = " & Ap(AssuntoSA) _
                & " Where Cod = " & Busca("Cod")
            Busca.MoveNext
            j = j - 1
            botNormalizar.Caption = "Normalizando (" & j & ")"
        Next i
    End If
    Busca.Close
End If

botNormalizar.Caption = "Normalizar tabela"

Exit Sub

Erro:
If Err.Number = -2147217865 Or Err.Number = -2147217900 Then
    txtLocalizar = Tabela & " <-Tabela não encontrada"
Else
    If Err.Number = -2147217904 Then
        txtLocalizar = Tabela & " <-Campo 'AssuntoSA' não encontrado"
    Else
        MsgBox "Erro: " & Err.Number & " - " & Err.Description
        Debug.Print Err.Number
    End If
End If

End Sub

Private Sub botOk_Click()

DoEvents
With lstBusca

Dim Termo As String
Termo = Trim(txtLocalizar)
Termo = RetirarAcento(Termo)

Dim Tabela As String, NomeLocal As String
Select Case botPrioridade.Caption
    Case "^^ Fichário ^^"
    Tabela = "tabFichário"
    NomeLocal = "Fichário"
    
    Case "^^ Biografias ^^"
    Tabela = "tabBiografias"
    NomeLocal = "Biografias"
    
    Case "^^ Guarulhos ^^"
    Tabela = "tabGuarulhos"
    NomeLocal = "Guarulhos"

    Case "^^ Revistas ^^"
    Tabela = "tabRevista"
    NomeLocal = "Revista"
End Select

'inicializa a lista
.Visible = False
.Clear
.Rows = 1
.ColAlignment(1) = 1
.ForeColor = Cor.Preto
.TextMatrix(0, 1) = "Assunto"
.TextMatrix(0, 2) = "Local / suporte"

If øCon.State = 1 Then
    Dim Busca As New ADODB.Recordset
    Busca.ActiveConnection = øCon
    Busca.CursorType = adOpenStatic
    
    Dim i As Integer
    i = 1
    
    Busca.Open "Select * From " & Tabela & " Where AssuntoSA Like '%" _
        & Replace(Termo, "'", "''") & "%' Order By Assunto"
        
    .Rows = Busca.RecordCount + 1
    If Busca.RecordCount > 0 Then
        Busca.MoveLast: Busca.MoveFirst
        
        
        For i = i To .Rows - 1
            .TextMatrix(i, 0) = Busca("Cod") & ""
            .TextMatrix(i, 1) = Busca("Assunto") & ""
            
            If NomeLocal = "Revista" Then
                .TextMatrix(i, 2) = NomeLocal & ": " & Busca("Revista") _
                    & vbNewLine & "Número: " & Busca("Número") _
                    & " | Data: " & Busca("Data")
                
            ElseIf NomeLocal = "Fichário" Then
                .TextMatrix(i, 2) = NomeLocal & ">" & Busca("Pasta")
                If Trim(Busca("Subpasta")) <> "" Then
                    .TextMatrix(i, 2) = .TextMatrix(i, 2) & ">" & Busca("Subpasta")
                End If
            Else
                .TextMatrix(i, 2) = NomeLocal & ">" & Busca("Pasta")
            End If
            Busca.MoveNext
        Next i
    End If
    Busca.Close
    
    
    If NomeLocal <> "Fichário" Then
        Busca.Open "Select * From tabFichário Where AssuntoSA Like '%" _
            & Replace(Termo, "'", "''") & "%' Order By Assunto"
            
    
        .Rows = .Rows + Busca.RecordCount
        If Busca.RecordCount > 0 Then
            Busca.MoveLast: Busca.MoveFirst
            
            For i = i To .Rows - 1
                .TextMatrix(i, 0) = Busca("Cod") & ""
                .TextMatrix(i, 1) = Busca("Assunto") & ""
                .TextMatrix(i, 2) = "Fichário" & ">" & Busca("Pasta")
                If Trim(Busca("Subpasta")) <> "" Then
                    .TextMatrix(i, 2) = .TextMatrix(i, 2) & ">" & Busca("Subpasta")
                End If
                Busca.MoveNext
            Next i
        End If
        Busca.Close
    End If
    
    If NomeLocal <> "Biografias" Then
        Busca.Open "Select * From tabBiografias Where AssuntoSA Like '%" _
            & Replace(Termo, "'", "''") & "%' Order By Assunto"
            
    
        .Rows = .Rows + Busca.RecordCount
        If Busca.RecordCount > 0 Then
            Busca.MoveLast: Busca.MoveFirst
            
            For i = i To .Rows - 1
                .TextMatrix(i, 0) = Busca("Cod") & ""
                .TextMatrix(i, 1) = Busca("Assunto") & ""
                .TextMatrix(i, 2) = "Biografias" & ">" & Busca("Pasta")
                Busca.MoveNext
            Next i
        End If
        Busca.Close
    End If
    
    If NomeLocal <> "Guarulhos" Then
        Busca.Open "Select * From tabGuarulhos Where AssuntoSA Like '%" _
            & Replace(Termo, "'", "''") & "%' Order By Assunto"
            
    
        .Rows = .Rows + Busca.RecordCount
        If Busca.RecordCount > 0 Then
            Busca.MoveLast: Busca.MoveFirst
            
            For i = i To .Rows - 1
                .TextMatrix(i, 0) = Busca("Cod") & ""
                .TextMatrix(i, 1) = Busca("Assunto") & ""
                .TextMatrix(i, 2) = "Guarulhos" & ">" & Busca("Pasta")
                Busca.MoveNext
            Next i
        End If
        Busca.Close
    End If
    
    If NomeLocal <> "Revista" Then
        Busca.Open "Select * From tabRevista Where AssuntoSA Like '%" _
            & Replace(Termo, "'", "''") & "%' Order By Assunto"
            
    
        .Rows = .Rows + Busca.RecordCount
        If Busca.RecordCount > 0 Then
            Busca.MoveLast: Busca.MoveFirst
            
            For i = i To .Rows - 1
                .TextMatrix(i, 0) = Busca("Cod") & ""
                .TextMatrix(i, 1) = Busca("Assunto") & ""
                .TextMatrix(i, 2) = "Revista" & ": " & Busca("Revista") _
                    & vbNewLine & "Número: " & Busca("Número") _
                    & " | Data: " & Busca("Data")
                Busca.MoveNext
            Next i
        End If
        Busca.Close
    End If
End If

For i = 1 To .Rows - 1
    If Left(.TextMatrix(i, 2), 8) = "Fichário" Then
        .Col = 2: .Row = i: .CellBackColor = Cor.VermelhoClaro
    End If
    
    If Left(.TextMatrix(i, 2), 10) = "Biografias" Then
        .Col = 2: .Row = i: .CellBackColor = Cor.VerdeClaríssimo
    End If
    
    If Left(.TextMatrix(i, 2), 9) = "Guarulhos" Then
        .Col = 2: .Row = i: .CellBackColor = Cor.AzulClaríssimo
    End If
    
    If Left(.TextMatrix(i, 2), 7) = "Revista" Then
        
        Dim j As Integer
        j = InStr(1, .TextMatrix(i, 2), vbNewLine)
        If j > 0 Then
            txt1 = Left(.TextMatrix(i, 2), j)
        Else
            txt1 = ""
        End If
        
        If TextWidth(txt1) > .ColWidth(2) Then
            .RowHeight(i) = 750
        Else
            .RowHeight(i) = 500
        End If
        
        .Col = 2: .Row = i: .CellBackColor = Cor.AmareloClaro
    End If
Next i

If .Rows = 1 Then
    .Rows = 2
    .TextMatrix(1, 1) = "<<NENHUMA OCORRÊNCIA ENCONTRDA>>"
    .ColAlignment(1) = 3
    .ForeColor = Cor.Vermelho
End If

.Visible = True
End With

End Sub

Private Sub botPrioridade_Click()

Select Case botPrioridade.Caption
    Case "^^ Fichário ^^"
    botPrioridade.Caption = "^^ Biografias ^^"
    botPrioridade.BackColor = Cor.VerdeClaríssimo
    
    Case "^^ Biografias ^^"
    botPrioridade.Caption = "^^ Guarulhos ^^"
    botPrioridade.BackColor = Cor.AzulClaríssimo
    
    Case "^^ Guarulhos ^^"
    botPrioridade.Caption = "^^ Revistas ^^"
    botPrioridade.BackColor = Cor.AmareloClaro

    Case "^^ Revistas ^^"
    botPrioridade.Caption = "^^ Fichário ^^"
    botPrioridade.BackColor = Cor.VermelhoClaro
End Select

botOk = True

End Sub



Private Sub Form_Load()

'ajusta à resolução do monitor ****************************************************
Dim ScaleFactorX As Single, ScaleFactorY As Single  ' Scaling factors
' Size of Form in Pixels at design resolution
DesignX = 800
DesignY = 600
RePosForm = True   ' Flag for positioning Form
DoResize = False   ' Flag for Resize Event
' Set up the screen values
Xtwips = Screen.TwipsPerPixelX
Ytwips = Screen.TwipsPerPixelY
Ypixels = Screen.Height / Ytwips ' Y Pixel Resolution
Xpixels = Screen.Width / Xtwips  ' X Pixel Resolution

' Determine scaling factors
ScaleFactorX = (Xpixels / DesignX)
ScaleFactorY = (Ypixels / DesignY)
ScaleMode = 1  ' twips
'Exit Sub  ' uncomment to see how Form1 looks without resizing
Resize_For_Resolution ScaleFactorX, ScaleFactorY, Me

MyForm.Height = Me.Height ' Remember the current size
MyForm.Width = Me.Width
'**********************************************************************************

lstBusca.Rows = 1
lstBusca.TextMatrix(0, 1) = "Assunto"
lstBusca.TextMatrix(0, 2) = "Local / suporte"
lstBusca.ColAlignment(1) = 0

AjustarLista


End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
botNormalizar.Visible = False
End Sub

Private Sub Form_Resize()

AjustarLista

On Error Resume Next
txtLocalizar.SetFocus

End Sub

Public Sub AjustarLista()

If Me.Height < 3000 Or Me.Width < 4000 Then Exit Sub

lstBusca.Width = Me.Width - lstBusca.Left * 3
lstBusca.Height = Me.Height - lstBusca.Top - 600

lstBusca.ColWidth(0) = 0
lstBusca.ColWidth(1) = (lstBusca.Width / 100) * 65
lstBusca.ColWidth(2) = (lstBusca.Width / 100) * 35


txtLocalizar.Width = lstBusca.ColWidth(1) - botOk.Width
botNormalizar.Left = txtLocalizar.Left + txtLocalizar.Width - botNormalizar.Width + 10
botOk.Left = txtLocalizar.Left + txtLocalizar.Width + 40
botPrioridade.Left = txtLocalizar.Left + txtLocalizar.Width + botOk.Width + 70
botPrioridade.Width = lstBusca.ColWidth(2) - 60
legP.Left = botPrioridade.Left

frmInserir.Width = lstBusca.Width
Dim i As Integer
For i = 0 To botInserir.UBound
    botInserir(i).Width = (frmInserir.Width - (botInserir.Count + 1) * 90) / 4
    If i = 0 Then
        botInserir(i).Left = 90
    Else
        botInserir(i).Left = botInserir(i - 1).Left + botInserir(i - 1).Width + 90
    End If
    
Next i

End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub lstBusca_DblClick()

If lstBusca.Row < 1 Then Exit Sub

If Left(lstBusca.TextMatrix(lstBusca.Row, 2), 8) = "Fichário" Then
With f2
    .Left = Me.Left + (Me.Width - .Width) / 2
    .Top = Me.Top + (Me.Height - .Height) / 2
    .Tag = lstBusca.TextMatrix(lstBusca.Row, 0)
    .legT.Caption = "Assunto do fichário (cod: " & .Tag & ")"
    
    If øCon.State = 1 Then
        Dim Busca As New ADODB.Recordset
        Busca.ActiveConnection = øCon
        Busca.CursorType = adOpenStatic

        Busca.Open "Select * From tabFichário Where Cod = " & .Tag
        
        If Busca.RecordCount > 0 Then
            .txtAssunto = Busca("Assunto") & ""
            .txtPasta = Busca("Pasta") & ""
            .txtSubpasta = Busca("Subpasta") & ""
        End If
        Busca.Close
    End If
    
    .botExcluir.Visible = True
    .Show 1
End With
End If



If Left(lstBusca.TextMatrix(lstBusca.Row, 2), 9) = "Biografia" Then
With f3
    .Left = Me.Left + (Me.Width - .Width) / 2
    .Top = Me.Top + (Me.Height - .Height) / 2
    .Tag = lstBusca.TextMatrix(lstBusca.Row, 0)
    .legT.Caption = "Biografia (cod: " & .Tag & ")"
    
    If øCon.State = 1 Then

        Busca.ActiveConnection = øCon
        Busca.CursorType = adOpenStatic

        Busca.Open "Select * From tabBiografias Where Cod = " & .Tag
        
        If Busca.RecordCount > 0 Then
            .txtAssunto = Busca("Assunto") & ""
            .txtPasta = Busca("Pasta") & ""
            .txtSubpasta = Busca("Comentário") & ""
        End If
        Busca.Close
    End If
    
    .botExcluir.Visible = True
    .Show 1
End With
End If

If Left(lstBusca.TextMatrix(lstBusca.Row, 2), 9) = "Guarulhos" Then
    With f4
    .Left = Me.Left + (Me.Width - .Width) / 2
    .Top = Me.Top + (Me.Height - .Height) / 2
    .Tag = lstBusca.TextMatrix(lstBusca.Row, 0)
    .legT.Caption = "Assunto de Guarulhos (cod: " & .Tag & ")"
    
    If øCon.State = 1 Then

        Busca.ActiveConnection = øCon
        Busca.CursorType = adOpenStatic

        Busca.Open "Select * From tabGuarulhos Where Cod = " & .Tag
        
        If Busca.RecordCount > 0 Then
            .txtAssunto = Busca("Assunto") & ""
            .txtPasta = Busca("Pasta") & ""
            .txtSubpasta = Busca("Comentário") & ""
        End If
        Busca.Close
    End If
    
    .botExcluir.Visible = True
    .Show 1
End With
End If


If Left(lstBusca.TextMatrix(lstBusca.Row, 2), 7) = "Revista" Then
With f5
    .Left = Me.Left + (Me.Width - .Width) / 2
    .Top = Me.Top + (Me.Height - .Height) / 2
    .Tag = lstBusca.TextMatrix(lstBusca.Row, 0)
    .legT.Caption = "Assunto de revista (cod: " & .Tag & ")"
    
    If øCon.State = 1 Then

        Busca.ActiveConnection = øCon
        Busca.CursorType = adOpenStatic

        Busca.Open "Select * From tabRevista Where Cod = " & .Tag
        
        If Busca.RecordCount > 0 Then
            .txtAssunto = Busca("Assunto") & ""
            .txtRevista = Busca("Revista") & ""
            .txtNúmero = Busca("Número") & ""
            .txtData = Busca("Data") & ""
        End If
        Busca.Close
    End If
    
    .botExcluir.Visible = True
    .Show 1
End With
End If
  
End Sub

Private Sub lstBusca_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If lstBusca.Rows > 1 Then
    ÞIndLista = lstBusca.Row
End If

End Sub


Private Sub lstBusca_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If lstBusca.Rows > 1 And ÞIndLista > 0 Then
    lstBusca.Col = 0
    lstBusca.ColSel = 1
    lstBusca.RowSel = ÞIndLista
End If

End Sub


Private Sub txtLocalizar_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then
    botOk = True
End If

End Sub


Private Sub txtLocalizar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Shift = 7 Then
    botNormalizar.Visible = True
Else
    
End If

End Sub


