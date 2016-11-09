VERSION 5.00
Begin VB.Form frmSenhaLiberacao 
   Caption         =   "Senha"
   ClientHeight    =   1305
   ClientLeft      =   3525
   ClientTop       =   3720
   ClientWidth     =   4080
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1305
   ScaleWidth      =   4080
   Begin VB.CommandButton cmdRetornar 
      Caption         =   "&Retornar"
      Height          =   330
      Left            =   3150
      TabIndex        =   5
      Top             =   915
      Width           =   795
   End
   Begin VB.Frame frasenha 
      Height          =   795
      Left            =   90
      TabIndex        =   0
      Top             =   15
      Width           =   3855
      Begin VB.TextBox txtsenha 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2565
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   270
         Width           =   1080
      End
      Begin VB.TextBox txtusuario 
         Height          =   315
         Left            =   750
         TabIndex        =   2
         Top             =   270
         Width           =   1080
      End
      Begin VB.Label lblSenha 
         AutoSize        =   -1  'True
         Caption         =   "Senha"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   1950
         TabIndex        =   3
         Top             =   375
         Width           =   465
      End
      Begin VB.Label lblusuario 
         AutoSize        =   -1  'True
         Caption         =   "Usuario"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   135
         TabIndex        =   1
         Top             =   375
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmSenhaLiberacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdRetornar_Click()
    
    Unload Me
    
End Sub

Private Sub txtSenha_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        If txtSenha.Text <> "" Then
            cmdRetornar.SetFocus
        End If
    End If
    
End Sub

Private Sub txtSenha_LostFocus()
    
    If txtUsuario.Text = "" Then
        txtUsuario.SetFocus
    ElseIf txtSenha.Text <> "" Then
        If LiberaSenha(txtUsuario.Text, txtSenha.Text) = True Then
            Screen.MousePointer = 11
            LiberaBc2000 2
            Screen.MousePointer = 0
            Unload Me
        Else
            MsgBox "Senha incorreta", vbCritical, "Atenção"
            txtUsuario.SelStart = 0
            txtUsuario.SelLength = Len(txtUsuario.Text)
            txtUsuario.SetFocus
            txtSenha.Text = ""
        End If
    End If
    
End Sub

Private Sub txtusuario_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        If txtUsuario.Text <> "" Then
            txtSenha.SelStart = 0
            txtSenha.SelLength = Len(txtSenha.Text)
            txtSenha.SetFocus
        End If
    End If
    
End Sub
