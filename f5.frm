VERSION 5.00
Begin VB.Form f5 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hemeroteca: Revista"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7680
   Icon            =   "f5.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   7680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtData 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5700
      MaxLength       =   150
      TabIndex        =   10
      Top             =   1160
      Width           =   1755
   End
   Begin VB.CommandButton botFechar 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Fechar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6300
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1740
      Width           =   1185
   End
   Begin VB.CommandButton botExcluir 
      BackColor       =   &H008080FF&
      Caption         =   "Excluir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1740
      Width           =   1185
   End
   Begin VB.CommandButton botSalvar 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Salvar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1740
      Width           =   1185
   End
   Begin VB.TextBox txtNúmero 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3800
      MaxLength       =   150
      TabIndex        =   5
      Top             =   1160
      Width           =   1755
   End
   Begin VB.TextBox txtRevista 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   180
      MaxLength       =   150
      TabIndex        =   3
      Top             =   1160
      Width           =   3465
   End
   Begin VB.TextBox txtAssunto 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   180
      MaxLength       =   150
      TabIndex        =   0
      Top             =   510
      Width           =   7305
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data"
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
      Left            =   5700
      TabIndex        =   11
      Top             =   950
      Width           =   420
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Número"
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
      Left            =   3800
      TabIndex        =   6
      Top             =   945
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Revista"
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
      Left            =   210
      TabIndex        =   4
      Top             =   960
      Width           =   660
   End
   Begin VB.Label legT 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Cadastro de novo assunto de revista"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   30
      TabIndex        =   2
      Top             =   30
      Width           =   7605
   End
   Begin VB.Label leg1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Assunto"
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
      Left            =   210
      TabIndex        =   1
      Top             =   300
      Width           =   690
   End
End
Attribute VB_Name = "f5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'variáveis usadas para ajustar à resolução do monitor ****************************
Dim MyForm As FRMSIZE
Dim DesignX As Integer
Dim DesignY As Integer

Option Compare Text
Option Explicit

Private Sub botExcluir_Click()

If MsgBox("Deseja mesmo excluir este assunto de revista?", vbYesNo, "Hemeroteca") = vbYes Then
    If øCon.State = 1 Then
        øCon.Execute "Delete * From tabRevista Where Cod = " & CDbl(Me.Tag)
    End If
    
    f1.lstBusca.TextMatrix(f1.lstBusca.Row, 1) = "<<EXCLUÍDO>>"
    f1.lstBusca.TextMatrix(f1.lstBusca.Row, 2) = "<<EXCLUÍDO>>"
    Unload Me
End If

End Sub

Private Sub botFechar_Click()

Unload Me

End Sub

Private Sub botSalvar_Click()

Dim Assunto As String, AssuntoSA As String, Revista As String, _
    Número As String, Data As String
Assunto = Trim(txtAssunto)
AssuntoSA = RetirarAcento(Assunto)
Revista = Trim(txtRevista)
Número = Trim(txtNúmero)
Data = Trim(txtData)

If Assunto = "" Then
    txtAssunto.SetFocus
    Exit Sub
End If

If Revista = "" Then
    txtRevista.SetFocus
    Exit Sub
End If

If øCon.State = 1 Then
    If botExcluir.Visible = True Then
        øCon.Execute "Update tabRevista Set " _
        & "Assunto = " & Ap(Assunto) _
        & ", AssuntoSA = " & Ap(AssuntoSA) _
        & ", Revista = " & Ap(Revista) _
        & ", Número = " & Ap(Número) _
        & ", Data = " & Ap(Data) _
        & " Where Cod = " & CDbl(Me.Tag)
    Else
        øCon.Execute "Insert Into tabRevista " _
        & "(Assunto, AssuntoSA, Revista, Número, Data) Values (" _
        & Ap(Assunto) & ", " & Ap(AssuntoSA) & ", " _
        & Ap(Revista) & ", " & Ap(Número) & ", " & Ap(Data) & ")"
    End If
End If

'If Trim(f1.txtLocalizar) <> "" Then f1.botOk = True
Unload Me


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

End Sub

Private Sub Text3_Change()

End Sub




Private Sub txtDocumentos_KeyPress(KeyAscii As Integer)

If KeyAscii <> vbKeyBack Then
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End If

End Sub


