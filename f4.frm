VERSION 5.00
Begin VB.Form f4 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hemeroteca: Guarulhos"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7680
   Icon            =   "f4.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   7680
   StartUpPosition =   3  'Windows Default
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
   Begin VB.TextBox txtSubpasta 
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
      Left            =   4020
      MaxLength       =   150
      TabIndex        =   5
      Top             =   1160
      Width           =   3465
   End
   Begin VB.TextBox txtPasta 
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Comentário"
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
      Left            =   4050
      TabIndex        =   6
      Top             =   945
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pasta"
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
      Width           =   495
   End
   Begin VB.Label legT 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "Cadastro de novo assunto de Guarulhos"
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
Attribute VB_Name = "f4"
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

If MsgBox("Deseja mesmo excluir este assunto de Guarulhos?", vbYesNo, "Hemeroteca") = vbYes Then
    If øCon.State = 1 Then
        øCon.Execute "Delete * From tabGuarulhos Where Cod = " & CDbl(Me.Tag)
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

Dim Assunto As String, AssuntoSA As String, Pasta As String, Comentário As String
Assunto = Trim(txtAssunto)
AssuntoSA = RetirarAcento(Assunto)
Pasta = Trim(txtPasta)
Comentário = Trim(txtSubpasta)

If Assunto = "" Then
    txtAssunto.SetFocus
    Exit Sub
End If

If øCon.State = 1 Then
    If botExcluir.Visible = True Then
        øCon.Execute "Update tabGuarulhos Set " _
        & "Assunto = " & Ap(Assunto) _
        & ", AssuntoSA = " & Ap(AssuntoSA) _
        & ", Pasta = " & Ap(Pasta) _
        & ", Comentário = " & Ap(Comentário) _
        & " Where Cod = " & CDbl(Me.Tag)
    Else
        øCon.Execute "Insert Into tabGuarulhos " _
        & "(Assunto, AssuntoSA, Pasta, Comentário) Values (" _
        & Ap(Assunto) & ", " & Ap(AssuntoSA) & ", " _
        & Ap(Pasta) & ", " & Ap(Comentário) & ")"
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


