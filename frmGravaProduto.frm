VERSION 5.00
Begin VB.Form frmGravaProduto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Grava Produto"
   ClientHeight    =   1860
   ClientLeft      =   3840
   ClientTop       =   2565
   ClientWidth     =   3645
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   3645
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdRetonar 
      Caption         =   "&Retornar"
      Height          =   345
      Left            =   2685
      TabIndex        =   3
      Top             =   1395
      Width           =   840
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "&Gravar"
      Default         =   -1  'True
      Height          =   360
      Left            =   1830
      TabIndex        =   2
      Top             =   1395
      Width           =   840
   End
   Begin VB.Frame fraReferencia 
      Height          =   1260
      Left            =   165
      TabIndex        =   1
      Top             =   45
      Width           =   3315
      Begin VB.OptionButton optProdu 
         Caption         =   "Fornecedor"
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   2
         Left            =   150
         TabIndex        =   6
         Top             =   480
         Width           =   1110
      End
      Begin VB.OptionButton optProdu 
         Caption         =   "Código Barras"
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   1
         Left            =   1470
         TabIndex        =   5
         Top             =   180
         Width           =   1290
      End
      Begin VB.OptionButton optProdu 
         Caption         =   "Referência"
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   0
         Left            =   150
         TabIndex        =   4
         Top             =   180
         Value           =   -1  'True
         Width           =   1080
      End
      Begin VB.TextBox txtReferencia 
         Height          =   330
         Left            =   675
         MaxLength       =   20
         TabIndex        =   0
         Top             =   780
         Width           =   1830
      End
   End
End
Attribute VB_Name = "frmGravaProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGravar_Click()
    Screen.MousePointer = 11
    If optProdu(0).Value = True Then
        GravaProduto (txtReferencia.Text)
    ElseIf optProdu(1).Value = True Then
        GravaCodigoBarras (txtReferencia.Text)
    ElseIf optProdu(2).Value = True Then
        GravaFornecedor (txtReferencia.Text)
    End If
    txtReferencia.Text = ""
    txtReferencia.SetFocus
    Screen.MousePointer = 0
End Sub

Private Sub cmdRetonar_Click()
    
    'rdoCNRetaguarda.Close
    Unload Me
    
End Sub

Private Sub Form_Load()
    
    Left = (Screen.Width - Width) / 2
    Top = (Screen.Height - Height) / 4
    'optProdu(0).Value = True
    Screen.MousePointer = 11
    On Error Resume Next
    MsgBox "Atenção para executar este processo sera preciso estar conectado a InterNet", vbInformation, "Atenção"
    'Conexao.Close
   ' If ConectaODBCRetaguarda(Conexao, Cliptografia(GLB_Usuario), Cliptografia(GLB_Senha)) = False Then
   '     MsgBox "Não foi possivel conectar-se no servidor", vbCritical, "ERRO"
   '     Unload Me
   '     Exit Sub
   ' End If
    
     ConectaODBCRetaguarda
             
               
    If wConectouRetaguarda = False Then
       MsgBox "Não Conectou no banco da Matriz. Favor Verificar se internet está no Ar", vbCritical, "Atenção"
     '  rdoCNRetaguarda.Close
       Unload Me
       Exit Sub
    Exit Sub
    End If
    
   ' rdoCNRetaguarda.Close
   ' optProdu(0).Value = True
    Screen.MousePointer = 0
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
rdoCNRetaguarda.Close
End Sub

Private Sub optProdu_Click(Index As Integer)

    If optProdu(0).Value = True Then
        txtReferencia.MaxLength = 7
    ElseIf optProdu(1).Value = True Then
        txtReferencia.MaxLength = 20
    ElseIf optProdu(2).Value = True Then
        txtReferencia.MaxLength = 3
    End If

End Sub
