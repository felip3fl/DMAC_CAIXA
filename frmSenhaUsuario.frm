VERSION 5.00
Begin VB.Form frmSenhaUsuario 
   Caption         =   "Liberação do Controle de Estoque"
   ClientHeight    =   1620
   ClientLeft      =   4065
   ClientTop       =   3870
   ClientWidth     =   4155
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1620
   ScaleWidth      =   4155
   Begin VB.CommandButton cmdRetornar 
      Caption         =   "&Retornar"
      Height          =   360
      Left            =   2955
      TabIndex        =   2
      Top             =   1155
      Width           =   1005
   End
   Begin VB.Frame fraUsuarioSenha 
      Height          =   1005
      Left            =   150
      TabIndex        =   3
      Top             =   30
      Width           =   3825
      Begin VB.TextBox txtSenha 
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   2535
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   360
         Width           =   1020
      End
      Begin VB.TextBox txtusuario 
         Height          =   315
         Left            =   855
         TabIndex        =   0
         Top             =   375
         Width           =   990
      End
      Begin VB.Label lblSenha 
         Caption         =   "Senha"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   1980
         TabIndex        =   5
         Top             =   495
         Width           =   900
      End
      Begin VB.Label lblusuario 
         AutoSize        =   -1  'True
         Caption         =   "Usuario"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   225
         TabIndex        =   4
         Top             =   495
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmSenhaUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdRetornar_Click()

    Unload Me

End Sub

Private Sub Form_Load()
    
    
    txtsenha.Text = ""
    txtusuario.Text = ""
    
End Sub

Private Sub txtSenha_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        cmdRetornar.SetFocus
    End If
    
End Sub

Private Sub txtSenha_LostFocus()
    
    
    If txtsenha.Text <> "" And txtusuario.Text <> "" Then
        If LiberaSenha(txtusuario.Text, txtsenha.Text) = True Then
            Glb_UsuarioEstoque = txtusuario.Text
            Glb_SenhaEstoque = txtsenha.Text
            Unload Me
            If glb_LiberaSenha = 1 Then
                frmAjusteEstoque.Show
            ElseIf glb_LiberaSenha = 2 Then
                frmTransferenciaEntrada.Show
            ElseIf glb_LiberaSenha = 3 Then
                frmPromocao.Show
            End If
        Else
            MsgBox "Senha Incorreta, Verifique", vbCritical, "Atenção"
            txtsenha.Text = ""
            txtusuario.SelStart = 0
            txtusuario.SelLength = Len(txtusuario.Text)
            txtusuario.SetFocus
        End If
    Else
        txtusuario.Text = ""
        txtsenha.Text = ""
    End If
    
End Sub




Private Sub txtusuario_GotFocus()
    
    txtusuario.SelStart = 0
    txtusuario.SelLength = Len(txtusuario.Text)
    txtusuario.SetFocus
    
End Sub

Private Sub txtusuario_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        txtsenha.SelStart = 0
        txtsenha.SelLength = Len(txtsenha.Text)
        txtsenha.SetFocus
    End If
    
End Sub
