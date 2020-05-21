Attribute VB_Name = "Inicial"
Option Explicit

Public øServidor As String
Public øCon As New ADODB.Connection

Public Sub Main()

If Existe(App.Path & "\Backup") = False Then
    MkDir App.Path & "\Backup"
End If
    
    'If CaminhoExiste(App.Path & "\Backup\" _
        & Format(Date, "dd-mm-yyyy") & ".mdb") = False Then
        'f1.MsgInfo "->Fazendo cópia do banco de dados", Cor.CinzaClaro
        'FileCopy App.Path & "\BdSenha.mdb", App.Path & "\Backup\" _
            & Format(Date, "dd-mm-yyyy") & ".mdb"
    'End If
    
    Dim Banco As String
    Banco = App.Path & "\HEMEROTECA.mdb"
    If Existe(Banco) = 1 Then
        øCon.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" _
        & Banco & ";Mode=ReadWrite;Persist Security Info=False"
        øCon.Open
        f1.Show
    Else

    End If


f1.Show

End Sub

